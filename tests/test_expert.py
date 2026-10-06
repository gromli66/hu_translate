# -*- coding: utf-8 -*-
"""Глоссарии в БД и экраны эксперта: снимок для движка, импорт глоссария заказчика, кандидаты, память, права."""
import io
import json
import time
from pathlib import Path

import pytest

from server import config, db, jobs, security, glossary_db as GD
from test_server import env, client, login, SECRET      # noqa: F401 — фикстуры и помощники сервиса

import pipeline as P

REPO_PROJECTS = config.ROOT / "projects"
CUSTOMER_DOCX = Path(r"C:/project/pid_pipeline_deploy/_scratch/hu_translate_2026-10/input/_Большие_и_глоссарии/"
                     r"Глоссарии/_Глоссарий HU-RU.docx")
MD = ("# Г\n\n## Правила применения\n\n- коды не переводятся\n\n## Системы\n\n"
      "| Венгерский | Русский | Комментарий |\n|---|---|---|\n| szelep | клапан | |\n| szivattyú | насос | |\n")


@pytest.fixture
def demo(env):
    (config.PROJECTS / "demo" / "g.md").write_text(MD, encoding="utf-8")
    ids = {r["login"]: r["id"] for r in db.q("select id, login from users")}
    db.x("update users set role='expert' where login='bob'")
    return ids


def entries_of(tr):
    return [(e.hu, e.ru, e.note, e.section) for e in tr.entries]


@pytest.mark.skipif(not (REPO_PROJECTS / "paks" / "data" / "glossary_FAT.md").exists(), reason="нет данных проекта paks")
def test_snapshot_gives_engine_same_glossary(tmp_path, monkeypatch):
    """Снимок из БД должен давать движку ровно то же, что файлы проекта: иначе переводы (и кэш портала) поплывут."""
    monkeypatch.setattr(config, "DATA", tmp_path / "data")
    monkeypatch.setattr(config, "PROJECTS", REPO_PROJECTS)
    db.init()
    ver, pj = GD.snapshot("paks")
    orig = P.make_translator(P.load_project(REPO_PROJECTS / "paks" / "project.json"))
    snap = P.make_translator(P.load_project(pj))
    assert ver == 1 and snap.rules == orig.rules
    strict = lambda tr: sum("строгий" in (e.note or "") for e in tr.extra_entries)
    assert strict(snap) == strict(orig) > 0
    assert [(e.hu, e.ru) for e in snap.entries] == [(e.hu, e.ru) for e in orig.entries]
    assert entries_of(snap)[:len(orig.entries)] == entries_of(orig)


@pytest.mark.skipif(not CUSTOMER_DOCX.exists(), reason="нет docx-глоссария заказчика")
def test_customer_docx_import_is_subset_of_fat_md():
    """Приёмка из плана: импорт _Глоссарий HU-RU.docx даёт подмножество glossary_FAT.md."""
    rules, entries = GD.parse_docx(CUSTOMER_DOCX)
    _, fat = GD.parse_md((REPO_PROJECTS / "paks" / "data" / "glossary_FAT.md").read_text(encoding="utf-8"))
    pairs = {(hu, ru) for hu, ru, _, _ in fat}
    assert len(entries) > 700 and "Глоссарий обязателен" in rules
    missing = [(hu, ru) for hu, ru, _, _ in entries if (hu, ru) not in pairs]
    assert not missing, missing[:5]


def test_candidates_merge_and_decide(demo):
    GD.ensure_project("demo")
    items = [{"hu": "tápszivattyú", "ru": "питательный насос", "count": 3, "example": "A tápszivattyú indul."},
             {"hu": "szelep", "ru": "клапан"},                       # уже термин — не кандидат
             {"hu": "szivattyú", "ru": "помпа"}]                     # другой перевод известного термина — конфликт
    assert GD.merge_candidates("demo", items, "terms", demo["alice"]) == 2
    GD.merge_candidates("demo", items[:1], "edits", demo["bob"])
    c = db.one("select * from glossary where layer='candidate' and hu='tápszivattyú'")
    assert c["count"] == 6 and json.loads(c["authors"]) == sorted([demo["alice"], demo["bob"]])
    assert db.one("select note from glossary where hu='szivattyú' and layer='candidate'")["note"] == "в глоссарии: насос"
    v0 = GD.version("demo")
    GD.decide_candidate(c["id"], demo["bob"], accept=True, ru="насос питательной воды", strict=True)
    assert GD.version("demo") == v0 + 1
    _, pj = GD.snapshot("demo")
    tr = P.make_translator(P.load_project(pj))
    e = next(e for e in tr.extra_entries if e.hu == "tápszivattyú")
    assert e.ru == "насос питательной воды" and "строгий" in e.note       # принятый строгим — обязателен в движке
    rej = db.one("select id from glossary where hu='szivattyú' and layer='candidate'")["id"]
    GD.decide_candidate(rej, demo["bob"], accept=False)
    assert GD.merge_candidates("demo", [{"hu": "szivattyú", "ru": "помпа"}], "edits", demo["alice"]) == 0   # не всплывает


def test_expert_pages_and_roles(client, demo):
    GD.ensure_project("demo")
    GD.merge_candidates("demo", [{"hu": "tápszivattyú", "ru": "питательный насос", "example": "x"}], "terms", demo["alice"])
    login(client, "alice")
    assert client.get("/expert/candidates?project=demo").status_code == 403
    login(client, "bob")
    page = client.get("/expert/candidates?project=demo").text
    assert "tápszivattyú" in page and "глоссарий: версия 1" in page
    cid = db.one("select id from glossary where layer='candidate'")["id"]
    r = client.post(f"/expert/candidates/{cid}", data={"action": "accept", "ru": "насос ПВ"})
    assert "Принято: tápszivattyú → насос ПВ" in r.text
    assert "версия 2" in client.get("/expert/glossary?project=demo&q=táp").text
    g = db.one("select id from glossary where layer='expert' and hu='tápszivattyú'")["id"]
    assert "сохранено" in client.post(f"/expert/glossary/{g}", data={"action": "save", "ru": "насос ПВ", "strict": "true"}).text
    assert db.one("select strict from glossary where id=?", (g,))["strict"] == 1
    cust = db.one("select id from glossary where layer='customer'")["id"]
    assert "глоссарий заказчика" in client.post(f"/expert/glossary/{cust}", data={"action": "delete"}).text


def test_memory_approval(client, demo):
    now = time.time()
    for who, ru in (("alice", "Клапан закрыт."), ("root", "Клапан закрыт!")):
        db.x("insert into tm(project, hu_key, ru, author, created) values ('demo', 'A szelep zárva.', ?, ?, ?)",
             (ru, demo[who], now))
    login(client, "bob")
    page = client.get("/expert/memory?project=demo").text
    assert "Клапан закрыт." in page and "Клапан закрыт!" in page
    tid = db.one("select id from tm where ru='Клапан закрыт.'")["id"]
    assert "Утверждено" in client.post(f"/expert/memory/{tid}", data={"action": "approve"}).text
    assert {r["ru"]: r["status"] for r in db.q("select ru, status from tm")} == {"Клапан закрыт.": "approved",
                                                                                "Клапан закрыт!": "rejected"}


def test_import_customer_glossary_flow(client, demo):
    GD.ensure_project("demo")
    login(client, "bob")
    new_md = MD.replace("| szivattyú | насос | |", "| szivattyú | насосный агрегат | |\n| tartály | бак | |")
    r = client.post("/expert/import", data={"project": "demo"},
                    files={"file": ("g.md", io.BytesIO(new_md.encode()), "text/markdown")})
    assert "Новых: <strong>1</strong>" in r.text and "изменится перевод: <strong>1</strong>" in r.text
    token = r.text.split('name="token" value="')[1].split('"')[0]
    client.post("/expert/import/confirm", data={"token": token})
    assert GD.version("demo") == 2
    assert {r["hu"]: r["ru"] for r in db.q("select hu, ru from glossary where layer='customer'")} == \
        {"szelep": "клапан", "szivattyú": "насосный агрегат", "tartály": "бак"}


def test_candidates_collected_after_job(demo):
    GD.ensure_project("demo")
    uid = "job-c"
    jobs.job_dir(uid).mkdir(parents=True)
    (jobs.job_dir(uid) / "candidates.json").write_text(json.dumps(
        [{"hu": "tápszivattyú", "ru": "питательный насос", "count": 2, "example": "x"}], ensure_ascii=False), encoding="utf-8")
    db.x("insert into jobs(uid, owner, project, files, status, created) values (?, ?, 'demo', '[]', 'running', ?)",
         (uid, demo["alice"], time.time()))
    jobs.Scheduler._end(uid, "translate", True, "")
    c = db.one("select * from glossary where layer='candidate'")
    assert (c["hu"], c["source"], c["count"]) == ("tápszivattyú", "terms", 2)
    assert not (jobs.job_dir(uid) / "candidates.json").exists()


def test_terms_from_edits_parses_portal_answer(monkeypatch):
    import termx
    answers = iter(['Термин: [{"hu": "relétér", "ru": "релейное помещение"}]', "[]"])
    monkeypatch.setattr(termx, "chat", lambda *a, **k: {"text": next(answers)})
    out = termx.from_edits([("A relétérben.", "в релейном шкафу", "в релейном помещении"),
                            ("Kész.", "Готово", "Сделано")])
    assert out == [{"hu": "relétér", "ru": "релейное помещение", "example": "A relétérben."}]
