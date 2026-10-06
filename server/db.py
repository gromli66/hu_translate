# -*- coding: utf-8 -*-
"""SQLite: пользователи, сессии, задачи, правки редакторов, память переводов.
Соединение на каждый вызов: страницы и очередь работают из разных потоков."""
import sqlite3
from contextlib import contextmanager

from server import config

SCHEMA = """
create table if not exists users(
  id integer primary key, login text unique not null, name text not null default '',
  role text not null default 'user' check (role in ('user', 'expert', 'admin')),
  pw_hash text not null, token_enc text, active integer not null default 1, created real not null);
create table if not exists sessions(
  id text primary key, user_id integer not null references users(id), expires real not null);
create table if not exists jobs(
  uid text primary key, owner integer not null references users(id), project text not null, files text not null,
  review integer not null default 1, status text not null, error text, summary text,
  created real not null, started real, finished real);
create index if not exists jobs_status on jobs(status, created);
create table if not exists edits(
  id integer primary key, job text not null references jobs(uid), doc text not null, seg text not null,
  hu text not null, ru_before text not null, ru_after text not null, author integer not null references users(id),
  created real not null);
create table if not exists tm(
  id integer primary key, project text not null, hu_key text not null, ru text not null,
  status text not null default 'unconfirmed' check (status in ('unconfirmed', 'approved', 'rejected')),
  author integer not null references users(id), job text, created real not null,
  decided_by integer references users(id), decided_at real,
  unique (project, hu_key, author));
"""
# колонки, добавленные после первого выката: (таблица, колонка, объявление)
MIGRATIONS = [("jobs", "mode", "text not null default 'translate'")]


@contextmanager
def conn():
    config.DATA.mkdir(parents=True, exist_ok=True)
    c = sqlite3.connect(config.DATA / "hut.db", timeout=30)
    c.row_factory = sqlite3.Row
    try:
        yield c
        c.commit()
    finally:
        c.close()


def init():
    with conn() as c:
        c.execute("pragma journal_mode=wal")
        c.executescript(SCHEMA)
        for table, col, decl in MIGRATIONS:
            if col not in {r[1] for r in c.execute(f"pragma table_info({table})")}:
                c.execute(f"alter table {table} add column {col} {decl}")


def q(sql, args=()):
    with conn() as c:
        return c.execute(sql, args).fetchall()


def one(sql, args=()):
    with conn() as c:
        return c.execute(sql, args).fetchone()


def x(sql, args=()):
    with conn() as c:
        return c.execute(sql, args).rowcount
