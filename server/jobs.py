# -*- coding: utf-8 -*-
"""Очередь задач: поток внутри веб-приложения раз в секунду
  - забирает завершившиеся процессы задач и пишет итог в БД;
  - останавливает задачи, которые пользователь отменил;
  - запускает задачи из очереди: не больше одной на пользователя (два прогона на одном токене портал не держит,
    отвечает «200 без choices») и не больше HUT_MAX_JOBS всего.
Задача — отдельный процесс server/runner.py. После перезапуска сервера прерванные задачи снова встают в очередь:
шаги конвейера продолжают с места, ответы портала в кэше."""
import os
import sys
import json
import time
import logging
import threading
import subprocess
from pathlib import Path

from server import config, db, security

RUNNER = [sys.executable, "-m", "server.runner"]
log = logging.getLogger("hut.jobs")


def job_dir(uid):
    return config.DATA / "jobs" / uid


def read_progress(uid):
    try:
        return json.loads((job_dir(uid) / "progress.json").read_text(encoding="utf-8"))
    except (OSError, ValueError):
        return {}


def _tail(path, n=3):
    try:
        lines = [ln.strip() for ln in Path(path).read_text(encoding="utf-8", errors="replace").splitlines() if ln.strip()]
    except OSError:
        return ""
    return " | ".join(lines[-n:])[-400:]


class Scheduler:
    def __init__(self):
        self.procs = {}                 # uid -> (Popen, owner, файл лога)
        self.stop_ev = threading.Event()
        self.thread = None

    def start(self):
        db.x("update jobs set status='queued', started=null where status='running'")
        db.x("update jobs set status='cancelled', finished=? where status='cancelling'", (time.time(),))
        self.stop_ev.clear()
        self.thread = threading.Thread(target=self._loop, name="hut-scheduler", daemon=True)
        self.thread.start()

    def stop(self):
        self.stop_ev.set()
        if self.thread:
            self.thread.join(timeout=5)
        for p, _, lf in self.procs.values():   # сервер останавливается — задачи продолжатся после запуска
            p.terminate()
            lf.close()

    def _loop(self):
        while not self.stop_ev.is_set():
            try:
                self.tick()
            except Exception:                   # поток очереди не должен умирать от одной сбойной задачи
                log.exception("ошибка очереди задач")
            self.stop_ev.wait(1)

    def tick(self):
        for uid, (p, owner, lf) in list(self.procs.items()):
            rc = p.poll()
            if rc is None:
                if db.one("select status from jobs where uid=?", (uid,))["status"] == "cancelling":
                    p.terminate()
                continue
            lf.close()
            del self.procs[uid]
            self._finish(uid, rc)
        busy = {owner for _, owner, _ in self.procs.values()}
        for job in db.q("select j.uid, j.owner, u.token_enc from jobs j join users u on u.id = j.owner "
                        "where j.status='queued' order by j.created"):
            if len(self.procs) >= config.MAX_JOBS:
                break
            if job["owner"] in busy:
                continue
            token = security.dec(job["token_enc"])
            if not token:
                db.x("update jobs set status='failed', finished=?, error=? where uid=?",
                     (time.time(), "Нет токена портала. Впишите его в профиле и запустите перевод снова.", job["uid"]))
                continue
            self._spawn(job["uid"], job["owner"], token)
            busy.add(job["owner"])

    def _spawn(self, uid, owner, token):
        jd = job_dir(uid)
        env = dict(os.environ, ROSATOM_AI_TOKEN=token, HU_CACHE=str(config.DATA / "cache"),
                   PYTHONIOENCODING="utf-8", PYTHONUNBUFFERED="1")
        lf = open(jd / "log.txt", "a", encoding="utf-8")
        p = subprocess.Popen(RUNNER + [str(jd)], cwd=config.ROOT, env=env, stdout=lf, stderr=subprocess.STDOUT)
        self.procs[uid] = (p, owner, lf)
        db.x("update jobs set status='running', started=? where uid=?", (time.time(), uid))
        log.info("задача %s запущена", uid)

    def _finish(self, uid, rc):
        st = db.one("select status from jobs where uid=?", (uid,))["status"]
        now = time.time()
        if st == "cancelling":
            db.x("update jobs set status='cancelled', finished=? where uid=?", (now, uid))
        elif rc == 0:
            summ = read_progress(uid).get("summary") or {}
            db.x("update jobs set status='done', finished=?, summary=? where uid=?",
                 (now, json.dumps(summ, ensure_ascii=False), uid))
        else:
            db.x("update jobs set status='failed', finished=?, error=? where uid=?",
                 (now, _tail(job_dir(uid) / "log.txt") or f"процесс завершился с кодом {rc}", uid))
        log.info("задача %s: %s (код %s)", uid, st, rc)
