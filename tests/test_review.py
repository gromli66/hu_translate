# -*- coding: utf-8 -*-
"""Вычитка в браузере без портала: цвета строк, боковая панель, правка и её разнос, «Применить», память переводов."""
import json
import time

import pytest

from server import config, db, jobs, security
from test_server import env, client, login, wait, SECRET      # noqa: F401 — фикстуры и помощники сервиса

import pipeline as P

HU_SAME = "A szelep zárva."


def seg(i, text):
    return {"id": i, "text": text, "page": 0}


@pytest.fixture
def job(env):
    """Готовый перевод Алисы из двух документов; фраза HU_SAME есть в обоих."""
    (config.PROJECTS / "demo" / "g.md").write_text(
        "## Термины\n\n| Венгерский | Русский | Комментарий |\n|---|---|---|\n| szelep | клапан (запорный) | |\n",
        encoding="utf-8")
    alice = db.one("select id from users where login='alice'")["id"]
    db.x("update users set token_enc=? where login='alice'", (security.enc(SECRET),))
    uid = "job-review"
    work = jobs.job_dir(uid) / "work" / "json"
    work.mkdir(parents=True)
    (jobs.job_dir(uid) / "job.json").write_text(
        json.dumps({"project": str(config.PROJECTS / "demo" / "project.json"), "review": True}), encoding="utf-8")
    doc1 = {"file": "x", "segs": [seg(1, HU_SAME), seg(2, "Hiba van."), seg(3, "Nyomás 5 MPa."), seg(4, "30TL02"),
                                  seg(5, "Szivattyú indul.")],
            "res": {"1": {"ru": "Клапан открыт.", "issues": [], "review": {"serious": True, "reason": "zárva = закрыт"}},
                    "2": {"ru": "Ошибка есть.", "issues": ["нет кодов/чисел: 1"]},
                    "3": {"ru": "Давление 5 МПа.", "issues": [],
                          "autofix": {"accepted": True, "before": "Давление 50 МПа.", "reason": "число"}},
                    "4": {"ru": "30TL02", "issues": [], "copied": True},
                    "5": {"ru": "Насос запускается.", "issues": [], "gloss_miss": ["szivattyú"]}}}
    doc2 = {"file": "y", "segs": [seg(1, HU_SAME)], "res": {"1": {"ru": "Клапан открыт.", "issues": []}}}
    (work / "doc1.json").write_text(json.dumps(doc1, ensure_ascii=False), encoding="utf-8")
    (work / "doc2.json").write_text(json.dumps(doc2, ensure_ascii=False), encoding="utf-8")
    db.x("insert into jobs(uid, owner, project, files, status, created) values (?, ?, 'demo', '[\"doc1.docx\"]', 'done', ?)",
         (uid, alice, time.time()))
    return uid


def test_rows_kinds(client, job):
    login(client, "alice")
    rows = {r["id"]: r["kind"] for r in client.get(f"/jobs/{job}/review/rows?doc=doc1").json()}
    assert rows == {"1": "suspect", "2": "error", "3": "auto", "5": "term"}        # скопированный код не показывается


def test_side_panel(client, job):
    login(client, "alice")
    t = client.get(f"/jobs/{job}/review/side?doc=doc1&sid=1").text
    assert "zárva = закрыт" in t and "клапан" in t and "заказчик" in t
    t = client.get(f"/jobs/{job}/review/side?doc=doc1&sid=3").text
    assert "Давление 50 МПа." in t and "Вернуть «было»" in t


def test_edit_spreads_to_same_phrase_and_records(client, job):
    login(client, "alice")
    r = client.post(f"/jobs/{job}/review/edit", data={"doc": "doc1", "sid": "1", "ru": "Клапан закрыт."}).json()
    assert r == {"changed": [{"id": "1", "ru": "Клапан закрыт.", "kind": "edited"}], "pending": 1}
    work = jobs.job_dir(job) / "work"
    for doc in ("doc1", "doc2"):                                                   # та же фраза в другом документе
        res = json.loads((work / "json" / f"{doc}.json").read_text(encoding="utf-8"))["res"]["1"]
        assert res["ru"] == "Клапан закрыт." and res["edited"] and "review" not in res
    e = db.one("select * from edits")
    assert (e["hu"], e["ru_before"], e["ru_after"]) == (HU_SAME, "Клапан открыт.", "Клапан закрыт.")
    tm = db.one("select * from tm")
    assert (tm["hu_key"], tm["ru"], tm["status"]) == (P.tmkey(HU_SAME), "Клапан закрыт.", "unconfirmed")
    client.post(f"/jobs/{job}/review/edit", data={"doc": "doc1", "sid": "1", "ru": "Клапан закрыт!"})
    assert len(json.loads((jobs.job_dir(job) / "edits_pending.json").read_text(encoding="utf-8"))) == 1
    assert db.one("select count(*) n, max(ru) ru from tm")["n"] == 1                  # память — одна запись на автора


def test_edit_blocked_while_busy_and_for_strangers(client, job):
    login(client, "bob")
    assert client.post(f"/jobs/{job}/review/edit", data={"doc": "doc1", "sid": "1", "ru": "x"}).status_code == 404
    db.x("update jobs set status='running' where uid=?", (job,))
    login(client, "alice")
    assert client.post(f"/jobs/{job}/review/edit", data={"doc": "doc1", "sid": "1", "ru": "x"}).status_code == 409
    assert client.get(f"/jobs/{job}/review", follow_redirects=False).headers["location"] == "/?m=busy"


def test_apply_runs_edits_mode(client, job):
    login(client, "alice")
    client.post(f"/jobs/{job}/review/edit", data={"doc": "doc1", "sid": "1", "ru": "Клапан закрыт."})
    r = client.post(f"/jobs/{job}/review/apply", follow_redirects=False)
    assert r.headers["location"] == "/?m=applying"
    assert db.one("select mode from jobs where uid=?", (job,))["mode"] == "edits"
    wait(job, "done")
    j = db.one("select * from jobs where uid=?", (job,))
    assert j["mode"] == "translate" and not (jobs.job_dir(job) / "edits_pending.json").exists()


def test_tm_for_next_job_respects_owner_and_approval(env, monkeypatch):
    monkeypatch.setenv("FAKE_SLEEP", "3")
    ids = {r["login"]: r["id"] for r in db.q("select id, login from users")}
    for who in ("alice", "bob"):
        db.x("update users set token_enc=? where login=?", (security.enc(SECRET), who))
    now = time.time()
    db.x("insert into tm(project, hu_key, ru, status, author, created) values ('demo', 'a', 'своё Алисы', 'unconfirmed', ?, ?)",
         (ids["alice"], now))
    db.x("insert into tm(project, hu_key, ru, status, author, created) values ('demo', 'b', 'утверждено', 'approved', ?, ?)",
         (ids["root"], now))
    for who in ("alice", "bob"):
        (jobs.job_dir(who) / "in").mkdir(parents=True)
        db.x("insert into jobs(uid, owner, project, files, status, created) values (?, ?, 'demo', '[]', 'queued', ?)",
             (who, ids[who], now))
    s = jobs.Scheduler()
    try:
        s.tick()
        tm = {who: json.loads((jobs.job_dir(who) / "tm.json").read_text(encoding="utf-8")) for who in ("alice", "bob")}
        assert tm["alice"] == {"a": "своё Алисы", "b": "утверждено"}
        assert tm["bob"] == {"b": "утверждено"}                    # чужая неподтверждённая память не действует
    finally:
        for p, _, lf in s.procs.values():
            p.kill(); p.wait(); lf.close()


def test_propagate_number_variants_without_portal(tmp_path, monkeypatch):
    g = tmp_path / "g.md"
    g.write_text("## Т\n\n| Венгерский | Русский | Комментарий |\n|---|---|---|\n| szelep | клапан | |\n", encoding="utf-8")
    work = tmp_path / "work"
    (work / "json").mkdir(parents=True)
    segs = [seg(i, f"Oldalszám: {i}") for i in (1, 2, 3)]
    res = {"1": {"ru": "Лист: 1", "issues": [], "edited": True},
           "2": {"ru": "Номер страницы: 2", "issues": []}, "3": {"ru": "Номер страницы: 3", "issues": []}}
    (work / "json" / "d.json").write_text(json.dumps({"file": "d", "segs": segs, "res": res}, ensure_ascii=False),
                                          encoding="utf-8")
    import translate as T
    monkeypatch.setattr(T.Translator, "_call", lambda *a, **k: pytest.fail("портал не должен вызываться"))
    changed = P.propagate(work, {"glossary": str(g)}, [("Oldalszám: 1", "Лист: 1")])
    out = changed["d"]["res"]
    assert out["2"]["ru"] == "Лист: 2" and out["3"]["ru"] == "Лист: 3" and out["2"]["rag"]["before"] == "Номер страницы: 2"
    assert out["2"]["rag"]["numbers"]                       # подстановка чисел — правка, а не «проверить»
    from server.review import kind
    assert kind(out["2"]) == "edited"
