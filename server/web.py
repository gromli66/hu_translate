# -*- coding: utf-8 -*-
"""Общее для страниц: шаблоны, текущий пользователь, сообщения после действий."""
import time
from pathlib import Path

from fastapi import Request
from fastapi.templating import Jinja2Templates

from server import db

HERE = Path(__file__).resolve().parent
templates = Jinja2Templates(directory=HERE / "templates")
COOKIE = "hut_session"
ROLE_RU = {"user": "пользователь", "expert": "эксперт", "admin": "админ"}

# сообщения после действий: код в адресе → текст (сам текст в адрес не попадает)
MESSAGES = {
    "queued": ("Перевод поставлен в очередь. Прогресс — в списке ниже.", "ok"),
    "no_token": ("Сначала впишите токен портала в профиле — переводы идут под ним.", "warn"),
    "no_files": ("Нужны файлы DOCX или PDF. Другие форматы переводчик не читает.", "warn"),
    "bad_project": ("Такого проекта нет. Выберите проект из списка.", "warn"),
    "token_saved": ("Токен сохранён. Проверьте его кнопкой «Проверить».", "ok"),
    "token_removed": ("Токен удалён.", "ok"),
    "pw_changed": ("Пароль изменён.", "ok"),
    "pw_bad": ("Текущий пароль не подошёл.", "warn"),
    "pw_short": ("Новый пароль — не короче 8 знаков.", "warn"),
    "applying": ("Правки применяются к похожим строкам, файлы пересобираются. Прогресс — в списке переводов.", "ok"),
    "busy": ("Перевод сейчас в работе — вычитка откроется, когда он закончится.", "warn"),
}


class NeedLogin(Exception):
    pass


def current_user(request: Request):
    sid = request.cookies.get(COOKIE)
    if not sid:
        return None
    return db.one("select u.* from sessions s join users u on u.id = s.user_id "
                  "where s.id=? and s.expires>? and u.active=1", (sid, time.time()))


def need_user(request: Request):
    u = current_user(request)
    if not u:
        raise NeedLogin()
    return u


def page(request, name, user=None, **ctx):
    m = MESSAGES.get(request.query_params.get("m", ""))
    return templates.TemplateResponse(request, name, dict(ctx, user=user, role_ru=ROLE_RU, msg=m))


def own_job(u, uid):
    """Задача, если она есть и принадлежит пользователю (админу — любая)."""
    j = db.one("select * from jobs where uid=?", (uid,))
    if not j or (j["owner"] != u["id"] and u["role"] != "admin"):
        return None
    return j
