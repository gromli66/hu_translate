# -*- coding: utf-8 -*-
"""Веб-интерфейс переводчика: FastAPI + шаблоны Jinja2 + HTMX (живые обновления без своего JavaScript).
Запуск: uvicorn server.app:app --port 8010. Пользователей заводит админ: python manage.py user add …"""
import io
import json
import time
import shutil
import secrets
import zipfile
import logging
import urllib.error
import urllib.parse
import urllib.request
from pathlib import Path
from typing import List
from contextlib import asynccontextmanager

from fastapi import FastAPI, Request, Form, UploadFile, File
from fastapi.responses import HTMLResponse, RedirectResponse, Response
from fastapi.staticfiles import StaticFiles

from server import config, db, security, jobs, review
from server import expert
from server.web import HERE, COOKIE, NeedLogin, NotAllowed, need_user, page, own_job, projects

logging.basicConfig(level=logging.INFO, format="%(asctime)s %(name)s %(message)s")
SCHED = jobs.Scheduler()
ALLOWED = (".docx", ".pdf")


@asynccontextmanager
async def lifespan(_app):
    db.init()
    SCHED.start()
    yield
    SCHED.stop()


app = FastAPI(lifespan=lifespan, docs_url=None, redoc_url=None)
app.mount("/static", StaticFiles(directory=HERE / "static"), name="static")
app.include_router(review.router)
app.include_router(expert.router)


# ------------------------------------------------------------------ вход
@app.exception_handler(NeedLogin)
async def _need_login(request: Request, _exc):
    if request.headers.get("HX-Request"):          # фрагмент HTMX: перейти на вход всей страницей
        return Response(status_code=200, headers={"HX-Redirect": "/login"})
    return RedirectResponse("/login", status_code=303)


@app.exception_handler(NotAllowed)
async def _not_allowed(request: Request, _exc):
    return HTMLResponse("<p>Этот раздел — для экспертов. Роль выдаёт администратор.</p><p><a href='/'>К переводам</a></p>",
                        status_code=403)


@app.get("/health")
def health():
    return {"ok": True}


@app.get("/login", response_class=HTMLResponse)
def login_page(request: Request):
    return page(request, "login.html", error=None)


@app.post("/login")
def login(request: Request, login: str = Form(...), password: str = Form(...)):
    u = db.one("select * from users where login=? and active=1", (login.strip(),))
    if not u or not security.check_pw(password, u["pw_hash"]):
        return page(request, "login.html", error="Неверный логин или пароль.", login=login)
    sid = secrets.token_urlsafe(32)
    db.x("insert into sessions(id, user_id, expires) values (?, ?, ?)",
         (sid, u["id"], time.time() + config.SESSION_DAYS * 86400))
    db.x("delete from sessions where expires<?", (time.time(),))
    r = RedirectResponse("/", status_code=303)
    r.set_cookie(COOKIE, sid, max_age=config.SESSION_DAYS * 86400, httponly=True, samesite="lax",
                 secure=config.COOKIE_SECURE)
    return r


@app.post("/logout")
def logout(request: Request):
    db.x("delete from sessions where id=?", (request.cookies.get(COOKIE, ""),))
    r = RedirectResponse("/login", status_code=303)
    r.delete_cookie(COOKIE)
    return r


# ------------------------------------------------------------------ переводы
def _plural(n, one, few, many):
    n10, n100 = n % 10, n % 100
    return one if n10 == 1 and n100 != 11 else few if 2 <= n10 <= 4 and not 12 <= n100 <= 14 else many


def _eta(sec):
    if sec is None:
        return ""
    if sec < 60:
        return "меньше минуты"
    m = round(sec / 60)
    return f"ещё ~{m} мин" if m < 90 else f"ещё ~{m // 60} ч {m % 60} мин"


def job_views(user):
    admin = user["role"] == "admin"
    rows = db.q("select j.*, u.name, u.login from jobs j join users u on u.id = j.owner "
                + ("" if admin else "where j.owner=? ") + "order by j.created desc limit 50",
                () if admin else (user["id"],))
    queued = [r["uid"] for r in db.q("select uid from jobs where status='queued' order by created")]
    out = []
    for r in rows:
        files = json.loads(r["files"])
        v = {"uid": r["uid"], "status": r["status"], "project": r["project"], "mode": r["mode"],
             "title": files[0] + (f" + {len(files) - 1} {_plural(len(files) - 1, 'файл', 'файла', 'файлов')}"
                                  if len(files) > 1 else ""),
             "created": time.strftime("%d.%m %H:%M", time.localtime(r["created"])),
             "owner": (r["name"] or r["login"]) if admin else "", "error": r["error"] or ""}
        if r["status"] == "running":
            p = jobs.read_progress(r["uid"])
            v.update(pct=p.get("pct", 0), stage=p.get("stage", "запуск"), eta=_eta(p.get("eta_s")))
        elif r["status"] == "queued":
            v["pos"] = queued.index(r["uid"]) + 1 if r["uid"] in queued else 1
        elif r["status"] == "done":
            s = json.loads(r["summary"] or "{}")
            v.update(segments=s.get("segments", 0), attention=s.get("issues", 0) + s.get("suspicious", 0))
        out.append(v)
    return out


@app.get("/", response_class=HTMLResponse)
def index(request: Request):
    u = need_user(request)
    return page(request, "jobs.html", u, projects=projects(), jobs=job_views(u), has_token=bool(u["token_enc"]))


@app.get("/jobs/list", response_class=HTMLResponse)
def jobs_list(request: Request):
    u = need_user(request)
    return page(request, "_jobs.html", u, jobs=job_views(u))


@app.post("/jobs")
def create_job(request: Request, files: List[UploadFile] = File(...), project: str = Form(...),
               review: bool = Form(False)):
    u = need_user(request)
    if not u["token_enc"]:
        return RedirectResponse("/?m=no_token", status_code=303)
    if project not in projects():
        return RedirectResponse("/?m=bad_project", status_code=303)
    good = [f for f in files if f.filename and Path(f.filename).suffix.lower() in ALLOWED]
    if not good:
        return RedirectResponse("/?m=no_files", status_code=303)
    uid = time.strftime("%Y%m%d-%H%M%S-") + secrets.token_hex(3)
    jd = jobs.job_dir(uid)
    (jd / "in").mkdir(parents=True)
    names = []
    for f in good:
        name = Path(f.filename.replace("\\", "/")).name          # только имя: «../../x.docx» → «x.docx»
        stem, suf, k = Path(name).stem, Path(name).suffix, 1
        while name in names:
            k += 1
            name = f"{stem} ({k}){suf}"
        with open(jd / "in" / name, "wb") as out:
            shutil.copyfileobj(f.file, out)
        names.append(name)
    spec = {"project": str(config.PROJECTS / project / "project.json"), "review": review}
    (jd / "job.json").write_text(json.dumps(spec, ensure_ascii=False), encoding="utf-8")
    db.x("insert into jobs(uid, owner, project, files, review, status, created) values (?, ?, ?, ?, ?, 'queued', ?)",
         (uid, u["id"], project, json.dumps(names, ensure_ascii=False), int(review), time.time()))
    return RedirectResponse("/?m=queued", status_code=303)


@app.post("/jobs/{uid}/cancel", response_class=HTMLResponse)
def cancel_job(request: Request, uid: str):
    u = need_user(request)
    j = own_job(u, uid)
    if not j:
        return Response(status_code=404)
    if j["status"] == "queued" and j["mode"] == "edits":       # перевод готов, отменяется только применение правок
        db.x("update jobs set status='done', mode='translate' where uid=?", (uid,))
    elif j["status"] == "queued":
        db.x("update jobs set status='cancelled', finished=? where uid=?", (time.time(), uid))
    elif j["status"] == "running":
        db.x("update jobs set status='cancelling' where uid=?", (uid,))
    return page(request, "_jobs.html", u, jobs=job_views(u))


@app.get("/jobs/{uid}/download")
def download(request: Request, uid: str):
    u = need_user(request)
    j = own_job(u, uid)
    if not j or j["status"] != "done":
        return Response("Перевод не найден или ещё не готов.", status_code=404)
    out = jobs.job_dir(uid) / "work" / "out"
    buf = io.BytesIO()
    with zipfile.ZipFile(buf, "w", zipfile.ZIP_DEFLATED) as z:
        for f in sorted(out.glob("*")):
            if f.is_file():
                z.write(f, f.name)
    name = Path(json.loads(j["files"])[0]).stem
    return Response(buf.getvalue(), media_type="application/zip",
                    headers={"Content-Disposition": "attachment; filename=\"translation.zip\"; "
                             f"filename*=UTF-8''{urllib.parse.quote('перевод_' + name + '.zip')}"})


# ------------------------------------------------------------------ профиль
@app.get("/profile", response_class=HTMLResponse)
def profile(request: Request):
    u = need_user(request)
    return page(request, "profile.html", u, has_token=bool(u["token_enc"]))


@app.post("/profile/token")
def save_token(request: Request, token: str = Form("")):
    u = need_user(request)
    token = token.strip()
    db.x("update users set token_enc=? where id=?", (security.enc(token) if token else None, u["id"]))
    return RedirectResponse("/profile?m=" + ("token_saved" if token else "token_removed"), status_code=303)


def check_token(token):
    """Дешёвый запрос к порталу (список моделей) — токен действует или нет. Токен никуда не пишется."""
    req = urllib.request.Request(config.PORTAL_BASE + "/api/models", headers={"Authorization": "Bearer " + token})
    try:
        with urllib.request.urlopen(req, timeout=15) as r:
            return r.status == 200, "Токен действует."
    except urllib.error.HTTPError as e:
        return False, ("Портал не принял токен. Возьмите новый в настройках портала." if e.code in (401, 403)
                       else f"Портал ответил ошибкой {e.code}. Попробуйте позже.")
    except (urllib.error.URLError, TimeoutError, OSError):
        return False, "Портал недоступен с сервера. Сообщите администратору."


@app.post("/profile/token/check", response_class=HTMLResponse)
def check_token_view(request: Request, token: str = Form("")):
    u = need_user(request)
    t = token.strip() or security.dec(u["token_enc"])
    if not t:
        ok, text = False, "Токена нет: впишите его и сохраните."
    else:
        ok, text = check_token(t)
    return HTMLResponse(f'<span class="{"ok" if ok else "warn"}">{text}</span>')


@app.post("/profile/password")
def change_password(request: Request, old: str = Form(...), new: str = Form(...)):
    u = need_user(request)
    if not security.check_pw(old, u["pw_hash"]):
        return RedirectResponse("/profile?m=pw_bad", status_code=303)
    if len(new) < 8:
        return RedirectResponse("/profile?m=pw_short", status_code=303)
    db.x("update users set pw_hash=? where id=?", (security.hash_pw(new), u["id"]))
    return RedirectResponse("/profile?m=pw_changed", status_code=303)
