# -*- coding: utf-8 -*-
"""SQLite: пользователи, сессии, задачи. Соединение на каждый вызов: страницы и очередь работают из разных потоков."""
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
"""


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


def q(sql, args=()):
    with conn() as c:
        return c.execute(sql, args).fetchall()


def one(sql, args=()):
    with conn() as c:
        return c.execute(sql, args).fetchone()


def x(sql, args=()):
    with conn() as c:
        return c.execute(sql, args).rowcount
