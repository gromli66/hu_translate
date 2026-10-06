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
        db.x("update jobs set status='cancelled', finished=? where status='cancelling' and mode='translate'",
             (time.time(),))
        db.x("update jobs set status='done', mode='translate' where status='cancelling' and mode='edits'")
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
        for job in db.q("select j.uid, j.owner, j.project, j.mode, u.token_enc from jobs j "
                        "join users u on u.id = j.owner where j.status='queued' order by j.created"):
            if len(self.procs) >= config.MAX_JOBS:
                break
            if job["owner"] in busy:
                continue
            token = security.dec(job["token_enc"])
            if not token:
                self._end(job["uid"], job["mode"], False, "Нет токена портала. Впишите его в профиле и запустите снова.")
                continue
            self._spawn(job, token)
            busy.add(job["owner"])

    def _spawn(self, job, token):
        uid, owner = job["uid"], job["owner"]
        jd = job_dir(uid)
        # память переводов для задачи: утверждённое по проекту + неподтверждённое самого владельца;
        # утверждённое идёт последним и при совпадении фразы побеждает
        tm = {r["hu_key"]: r["ru"] for r in db.q(
            "select hu_key, ru from tm where project=? and (status='approved' or (status='unconfirmed' and author=?)) "
            "order by status='approved', created", (job["project"], owner))}
        (jd / "tm.json").write_text(json.dumps(tm, ensure_ascii=False), encoding="utf-8")
        env = dict(os.environ, ROSATOM_AI_TOKEN=token, HU_CACHE=str(config.DATA / "cache"),
                   PYTHONIOENCODING="utf-8", PYTHONUNBUFFERED="1")
        lf = open(jd / "log.txt", "a", encoding="utf-8")
        p = subprocess.Popen(RUNNER + [str(jd), job["mode"]], cwd=config.ROOT, env=env, stdout=lf,
                             stderr=subprocess.STDOUT)
        self.procs[uid] = (p, owner, lf)
        db.x("update jobs set status='running', started=? where uid=?", (time.time(), uid))
        log.info("задача %s запущена", uid)

    def _finish(self, uid, rc):
        j = db.one("select status, mode from jobs where uid=?", (uid,))
        if j["status"] == "cancelling":
            self._end(uid, j["mode"], None, "")
        elif rc == 0:
            self._end(uid, j["mode"], True, "")
        else:
            self._end(uid, j["mode"], False, _tail(job_dir(uid) / "log.txt") or f"процесс завершился с кодом {rc}")
        log.info("задача %s (%s): %s, код %s", uid, j["mode"], j["status"], rc)

    @staticmethod
    def _end(uid, mode, ok, error):
        """ok: True — успех, False — сбой, None — отменено. Сбой или отмена применения правок не портят
        готовый перевод: задача остаётся «готово», правки ждут следующего «Применить»."""
        now = time.time()
        if ok:
            summ = json.dumps(read_progress(uid).get("summary") or {}, ensure_ascii=False)
            db.x("update jobs set status='done', mode='translate', finished=?, summary=?, error=null where uid=?",
                 (now, summ, uid))
        elif mode == "edits":
            db.x("update jobs set status='done', mode='translate', finished=?, error=? where uid=?",
                 (now, "Правки не применились: " + error if error else None, uid))
        elif ok is None:
            db.x("update jobs set status='cancelled', finished=? where uid=?", (now, uid))
        else:
            db.x("update jobs set status='failed', finished=?, error=? where uid=?", (now, error, uid))
