# -*- coding: utf-8 -*-
"""Обслуживание раз в сутки (вызывает очередь задач): копия базы, удаление старых задач и кэша по сроку.
Удаление по сроку выключено по умолчанию (0): включает админ в docker/.env, когда решит, сколько хранить.
Журнал правок, память переводов и глоссарии лежат в базе и удалением задач не затрагиваются."""
import os
import time
import shutil
import sqlite3
import logging

from server import config, db

log = logging.getLogger("hut.maintenance")


def _int_env(name, default):
    try:
        return int(os.environ.get(name, "").strip() or default)
    except ValueError:
        return default


def backup(keep=None):
    """Копия базы средствами SQLite (согласованная, без остановки сервиса) → data/backups/hut_<дата>.db."""
    keep = keep if keep is not None else _int_env("HUT_BACKUP_KEEP", 14)
    d = config.DATA / "backups"
    d.mkdir(parents=True, exist_ok=True)
    dest = d / f"hut_{time.strftime('%Y-%m-%d')}.db"
    src = sqlite3.connect(config.DATA / "hut.db")
    try:
        out = sqlite3.connect(dest)
        with out:
            src.backup(out)
        out.close()
    finally:
        src.close()
    for old in sorted(d.glob("hut_*.db"))[:-keep] if keep > 0 else []:
        old.unlink()
    return dest


def expire_jobs(days=None):
    """Файлы задач старше срока удаляются, строка задачи остаётся (статус expired) — для истории и замера качества."""
    days = days if days is not None else _int_env("HUT_KEEP_DAYS", 0)
    if days <= 0:
        return 0
    edge = time.time() - days * 86400
    n = 0
    for j in db.q("select uid from jobs where status in ('done', 'failed', 'cancelled') and finished < ?", (edge,)):
        shutil.rmtree(config.DATA / "jobs" / j["uid"], ignore_errors=True)
        db.x("update jobs set status='expired' where uid=?", (j["uid"],))
        n += 1
    return n


def prune_cache(days=None):
    """Ответы портала старше срока удаляются: повтор такого запроса снова пойдёт на портал."""
    days = days if days is not None else _int_env("HUT_CACHE_DAYS", 0)
    if days <= 0:
        return 0
    edge = time.time() - days * 86400
    n = 0
    for f in (config.DATA / "cache").glob("*.json"):
        if f.stat().st_mtime < edge:
            f.unlink(missing_ok=True)
            n += 1
    return n


class Daily:
    """Раз в сутки (при первом тике нового дня) — копия базы и чистка по сроку."""

    def __init__(self):
        self.day = None

    def tick(self):
        today = time.strftime("%Y-%m-%d")
        if today == self.day:
            return
        self.day = today
        try:
            path = backup()
            log.info("копия базы: %s; задач удалено по сроку: %s; ответов кэша удалено: %s",
                     path.name, expire_jobs(), prune_cache())
        except (OSError, sqlite3.Error):
            log.exception("обслуживание не выполнено")
