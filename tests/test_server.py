# -*- coding: utf-8 -*-
"""Веб-сервис без портала: вход, токен, задача от загрузки до скачивания, права, очередь.
Задачу выполняет tests/fake_runner.py с тем же контрактом, что server/runner.py."""
import io
import sys
import json
import time
import zipfile
from pathlib import Path

import pytest
from cryptography.fernet import Fernet
from fastapi.testclient import TestClient

from server import config, db, jobs, security

SECRET = "tok-SECRET-123"


@pytest.fixture
def env(tmp_path, monkeypatch):
    monkeypatch.setattr(config, "DATA", tmp_path / "data")
    monkeypatch.setattr(config, "PROJECTS", tmp_path / "projects")
    monkeypatch.setenv("HUT_SECRET_KEY", Fernet.generate_key().decode())
    monkeypatch.setattr(jobs, "RUNNER", [sys.executable, str(Path(__file__).parent / "fake_runner.py")])
    (tmp_path / "projects" / "demo").mkdir(parents=True)
    (tmp_path / "projects" / "demo" / "project.json").write_text('{"glossary": "g.md"}', encoding="utf-8")
    db.init()
    for login, role in (("alice", "user"), ("bob", "user"), ("root", "admin")):
        db.x("insert into users(login, name, role, pw_hash, created) values (?, ?, ?, ?, ?)",
             (login, login.title(), role, security.hash_pw("password1"), time.time()))
    return tmp_path


@pytest.fixture
def client(env):
    from server.app import app
    with TestClient(app) as c:
        yield c


def login(c, who):
    c.cookies.clear()
    r = c.post("/login", data={"login": who, "password": "password1"}, follow_redirects=False)
    assert r.status_code == 303


def upload(c, *names, review=True):
    files = [("files", (n, io.BytesIO(b"doc"), "application/octet-stream")) for n in names]
    return c.post("/jobs", files=files, data={"project": "demo", **({"review": "true"} if review else {})},
                  follow_redirects=False)


def wait(uid, *statuses, timeout=20):
    end = time.time() + timeout
    while time.time() < end:
        st = db.one("select status from jobs where uid=?", (uid,))["status"]
        if st in statuses:
            return st
        time.sleep(0.2)
    raise AssertionError(f"задача {uid} не дошла до {statuses}, сейчас {st}")


def last_uid():
    return db.q("select uid from jobs order by created desc")[0]["uid"]


def test_login_and_redirects(client):
    r = client.get("/", follow_redirects=False)
    assert r.status_code == 303 and r.headers["location"] == "/login"
    r = client.get("/jobs/list", headers={"HX-Request": "true"})
    assert r.headers.get("HX-Redirect") == "/login"
    r = client.post("/login", data={"login": "alice", "password": "wrong-pass"})
    assert "Неверный логин или пароль" in r.text
    login(client, "alice")
    assert "Ваши переводы" in client.get("/").text


def test_job_needs_token(client):
    login(client, "alice")
    r = upload(client, "a.docx")
    assert r.headers["location"] == "/?m=no_token"
    assert not db.q("select 1 from jobs")


def test_job_runs_end_to_end_and_token_stays_secret(client, env):
    login(client, "alice")
    client.post("/profile/token", data={"token": SECRET})
    enc = db.one("select token_enc from users where login='alice'")["token_enc"]
    assert SECRET not in enc and security.dec(enc) == SECRET
    r = upload(client, "../../evil.docx", "notes.txt", "b.PDF")
    assert r.headers["location"] == "/?m=queued"
    uid = last_uid()
    jd = jobs.job_dir(uid)
    assert sorted(p.name for p in (jd / "in").iterdir()) == ["b.PDF", "evil.docx"]   # .txt отброшен, путь срезан
    assert wait(uid, "done", "failed") == "done"
    assert "Готово · строк 10 · просят внимания 3" in client.get("/jobs/list").text
    z = zipfile.ZipFile(io.BytesIO(client.get(f"/jobs/{uid}/download").content))
    assert sorted(z.namelist()) == ["b_RU.PDF", "evil_RU.docx"]
    assert json.loads((jd / "token_seen.json").read_text())["len"] == len(SECRET)     # токен дошёл до исполнителя
    for f in [jd / "log.txt", jd / "job.json", config.DATA / "hut.db"]:
        assert SECRET.encode() not in f.read_bytes(), f                               # и нигде не записан открыто


def test_failed_job_shows_reason(client):
    login(client, "alice")
    client.post("/profile/token", data={"token": SECRET})
    upload(client, "fail.docx")
    assert wait(last_uid(), "done", "failed") == "failed"
    assert "Не получилось: работаю | RuntimeError: тестовый сбой" in client.get("/jobs/list").text


def test_other_user_cannot_touch_job(client):
    login(client, "alice")
    client.post("/profile/token", data={"token": SECRET})
    upload(client, "a.docx")
    uid = last_uid()
    wait(uid, "done")
    login(client, "bob")
    assert client.get(f"/jobs/{uid}/download").status_code == 404
    assert client.post(f"/jobs/{uid}/cancel").status_code == 404
    assert "a.docx" not in client.get("/jobs/list").text
    login(client, "root")                                                             # админ видит все задачи
    assert "a.docx" in client.get("/jobs/list").text


def test_cancel_queued(client, monkeypatch):
    monkeypatch.setattr(config, "MAX_JOBS", 0)                                        # очередь стоит
    login(client, "alice")
    client.post("/profile/token", data={"token": SECRET})
    upload(client, "a.docx")
    uid = last_uid()
    assert "В очереди: 1-й" in client.get("/jobs/list").text
    client.post(f"/jobs/{uid}/cancel")
    assert db.one("select status from jobs where uid=?", (uid,))["status"] == "cancelled"


def test_one_job_per_user(env, monkeypatch):
    monkeypatch.setenv("FAKE_SLEEP", "5")
    for login_ in ("alice", "bob"):
        db.x("update users set token_enc=? where login=?", (security.enc(SECRET), login_))
    ids = {r["login"]: r["id"] for r in db.q("select id, login from users")}
    for k, who in enumerate(("alice", "alice", "bob")):
        uid = f"job{k}"
        (jobs.job_dir(uid) / "in").mkdir(parents=True)
        (jobs.job_dir(uid) / "in" / "a.docx").write_bytes(b"doc")
        db.x("insert into jobs(uid, owner, project, files, status, created) values (?, ?, 'demo', '[\"a.docx\"]', 'queued', ?)",
             (uid, ids[who], time.time() + k))
    s = jobs.Scheduler()
    try:
        s.tick()
        running = {r["uid"] for r in db.q("select uid from jobs where status='running'")}
        assert running == {"job0", "job2"}            # второй задаче Алисы ждать: один токен — один прогон
    finally:
        for p, _, lf in s.procs.values():
            p.kill()
            p.wait()
            lf.close()


def test_progress_weights(tmp_path):
    sys.path.insert(0, str(config.ROOT / "src"))
    from server.runner import Progress, W_REVIEW
    pr = Progress(tmp_path / "p.json", W_REVIEW)
    pr("перевод", 5, 10)
    p = json.loads((tmp_path / "p.json").read_text(encoding="utf-8"))
    assert p["pct"] == 27.5 and p["stage"] == "перевод"
