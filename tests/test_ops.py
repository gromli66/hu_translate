# -*- coding: utf-8 -*-
"""Шаг 4: копия базы, удаление по сроку, замер качества, лимит загрузки, ограничение попыток входа."""
import io
import os
import json
import time

from server import config, db, jobs, maintenance, expert
from test_server import env, client, login, SECRET, upload      # noqa: F401 — фикстуры и помощники сервиса


def add_job(uid, owner, status="done", finished=None, summary=None):
    db.x("insert into jobs(uid, owner, project, files, status, created, finished, summary_initial, glossary_version) "
         "values (?, ?, 'demo', '[\"a.docx\"]', ?, ?, ?, ?, 1)",
         (uid, owner, status, time.time(), finished or time.time(), json.dumps(summary) if summary else None))
    (jobs.job_dir(uid) / "in").mkdir(parents=True, exist_ok=True)


def test_backup_keeps_last(env):
    d = config.DATA / "backups"
    d.mkdir(parents=True)
    for day in ("2026-01-01", "2026-01-02", "2026-01-03"):
        (d / f"hut_{day}.db").write_bytes(b"old")
    path = maintenance.backup(keep=2)
    assert path.exists() and path.stat().st_size > 0
    assert sorted(p.name for p in d.glob("hut_*.db")) == ["hut_2026-01-03.db", path.name]


def test_expire_and_prune_only_when_enabled(env):
    alice = db.one("select id from users where login='alice'")["id"]
    add_job("old", alice, finished=time.time() - 100 * 86400)
    add_job("new", alice)
    add_job("running", alice, status="running", finished=time.time() - 100 * 86400)
    assert maintenance.expire_jobs(0) == 0 and jobs.job_dir("old").exists()          # по умолчанию не удаляем
    assert maintenance.expire_jobs(90) == 1
    assert not jobs.job_dir("old").exists() and jobs.job_dir("new").exists() and jobs.job_dir("running").exists()
    assert db.one("select status from jobs where uid='old'")["status"] == "expired"
    cache = config.DATA / "cache"
    cache.mkdir()
    (cache / "a.json").write_text("{}")
    (cache / "b.json").write_text("{}")
    os.utime(cache / "a.json", (time.time() - 40 * 86400,) * 2)
    assert maintenance.prune_cache(0) == 0 and maintenance.prune_cache(30) == 1
    assert [p.name for p in cache.iterdir()] == ["b.json"]


def test_summary_initial_survives_edits(env):
    alice = db.one("select id from users where login='alice'")["id"]
    add_job("j", alice, status="running")
    (jobs.job_dir("j") / "progress.json").write_text(json.dumps({"summary": {"segments": 10, "issues": 2}}))
    jobs.Scheduler._end("j", "translate", True, "")
    (jobs.job_dir("j") / "progress.json").write_text(json.dumps({"summary": {"segments": 10, "issues": 0}}))
    jobs.Scheduler._end("j", "edits", True, "")
    j = db.one("select summary, summary_initial from jobs where uid='j'")
    assert json.loads(j["summary_initial"])["issues"] == 2 and json.loads(j["summary"])["issues"] == 0


def test_quality_rows_and_page(client):
    ids = {r["login"]: r["id"] for r in db.q("select id, login from users")}
    add_job("q1", ids["alice"], summary={"segments": 100, "issues": 4, "suspicious": 10})
    for seg in ("1", "2", "2"):                                    # одна строка правлена дважды — считается один раз
        db.x("insert into edits(job, doc, seg, hu, ru_before, ru_after, author, created) values "
             "('q1', 'd', ?, 'x', 'a', 'b', ?, ?)", (seg, ids["alice"], time.time()))
    q = expert.quality_rows("demo")
    assert q["version"][0]["p_edited"] == 2.0 and q["version"][0]["p_susp"] == 10.0
    db.x("update users set role='expert' where login='bob'")
    login(client, "bob")
    assert "Правят редакторы" in client.get("/expert/quality?project=demo").text


def test_upload_limit(client, monkeypatch):
    from server import app as A
    monkeypatch.setattr(A, "MAX_UPLOAD", 10)
    login(client, "alice")
    client.post("/profile/token", data={"token": SECRET})
    r = client.post("/jobs", files=[("files", ("big.docx", io.BytesIO(b"x" * 100), "application/octet-stream"))],
                    data={"project": "demo"}, follow_redirects=False)
    assert r.headers["location"] == "/?m=too_big"
    assert not db.q("select 1 from jobs") and not list((config.DATA / "jobs").glob("*"))


def test_login_throttle(client):
    from server import app as A
    A.FAILS.clear()
    for _ in range(5):
        assert "Неверный логин" in client.post("/login", data={"login": "alice", "password": "wrong-pass"}).text
    r = client.post("/login", data={"login": "alice", "password": "password1"}, follow_redirects=False)
    assert "Слишком много неудачных попыток" in r.text
    A.FAILS.clear()
