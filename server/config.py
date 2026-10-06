# -*- coding: utf-8 -*-
"""Настройки веб-сервиса из окружения (в контейнере — строки .env рядом с docker-compose.yml). Модули читают их как config.X во время вызова,
поэтому тесты подменяют значения monkeypatch'ем без перезагрузки модулей."""
import os
from pathlib import Path

ROOT = Path(__file__).resolve().parent.parent


def _env(name, default):
    return os.environ.get(name, "").strip() or default


DATA = Path(_env("HUT_DATA", str(ROOT / "data")))               # база, задачи, кэш портала, ключ шифрования
PROJECTS = Path(_env("HUT_PROJECTS", str(ROOT / "projects")))   # проекты: папки с project.json
MAX_JOBS = int(_env("HUT_MAX_JOBS", "3"))                       # одновременно задач на весь сервер
COOKIE_SECURE = _env("HUT_COOKIE_SECURE", "0") == "1"           # 1 — сервис за HTTPS
PORTAL_BASE = _env("ROSATOM_AI_BASE", "https://go.ai-rosatom.ru").rstrip("/")
SESSION_DAYS = 14
