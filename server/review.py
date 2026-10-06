# -*- coding: utf-8 -*-
"""Вычитка в браузере: таблица Tabulator (строки — JSON), боковая панель (фрагмент HTMX), правка прямо в ячейке.

Правка сразу пишется:
  - в JSON задачи — во все одинаковые фразы всех документов задачи;
  - в журнал правок (edits) — сырьё для кандидатов в термины и замера качества;
  - в память переводов (tm) как «не подтверждено»: действует только в задачах автора, пока эксперт не утвердит.
Разнос по похожим строкам (числа, перевод по образцу) и пересборка файлов — кнопкой «Применить»:
задача уходит в очередь в режиме edits и идёт под токеном владельца. Пока она идёт, правка закрыта."""
import re
import json
import time
import threading
from functools import lru_cache

from fastapi import APIRouter, Request, Form
from fastapi.responses import HTMLResponse, JSONResponse, RedirectResponse, Response

from server import config, db, jobs
from server.web import need_user, page, own_job

import pipeline as P            # чтение JSON задачи, ключ фразы; к порталу веб-процесс не обращается
import glossary as G

router = APIRouter()
_guard, _locks = threading.Lock(), {}
BUSY = ("queued", "running", "cancelling")
KINDS = [("error", "Ошибки"), ("suspect", "Подозрения"), ("auto", "Изменено автоматически"),
         ("term", "Термины"), ("edited", "Правки")]


def job_lock(uid):
    with _guard:
        return _locks.setdefault(uid, threading.Lock())


def kind(r):
    """Цвет строки — как в таблице вычитки xlsx; правка человека важнее прежних пометок."""
    if r.get("edited") or (r.get("rag") or {}).get("numbers"):   # подстановка чисел в правку — тоже правка
        return "edited"
    if r.get("issues"):
        return "error"
    if (r.get("review") or {}).get("serious"):
        return "suspect"
    if (r.get("autofix") or {}).get("accepted") or r.get("rag"):
        return "auto"
    if r.get("warn") or r.get("gloss_miss"):
        return "term"
    return "ok"


def work_of(uid):
    return jobs.job_dir(uid) / "work"


def pending_file(uid):
    return jobs.job_dir(uid) / "edits_pending.json"


def pending(uid):
    try:
        return json.loads(pending_file(uid).read_text(encoding="utf-8"))
    except (OSError, ValueError):
        return []


def load_doc(uid, doc):
    f = work_of(uid) / "json" / f"{doc}.json"
    return json.loads(f.read_text(encoding="utf-8")) if f.exists() else None


@lru_cache(maxsize=4)
def _translator(project_json, mtime):
    """Глоссарии проекта для боковой панели (термины строки). Перечитываются, если project.json поменялся."""
    return P.make_translator(P.load_project(project_json))


def translator_for(job):
    pj = config.PROJECTS / job["project"] / "project.json"
    return _translator(str(pj), pj.stat().st_mtime)


def usable_job(request, uid):
    u = need_user(request)
    j = own_job(u, uid)
    return u, j


# ------------------------------------------------------------------ страница
@router.get("/jobs/{uid}/review", response_class=HTMLResponse)
def review_page(request: Request, uid: str, doc: str = ""):
    u, j = usable_job(request, uid)
    if not j:
        return Response("Перевод не найден.", status_code=404)
    if j["status"] in BUSY:
        return RedirectResponse("/?m=busy", status_code=303)
    docs = [f.stem for f in P.doc_files(work_of(uid))]
    if not docs:
        return Response("У этого перевода нет результатов для вычитки.", status_code=404)
    doc = doc if doc in docs else docs[0]
    d = load_doc(uid, doc)
    counts = {k: 0 for k, _ in KINDS}
    total = 0
    for s in d["segs"]:
        r = d["res"].get(str(s["id"])) or {}
        if r.get("copied") or not r:
            continue
        total += 1
        k = kind(r)
        if k in counts:
            counts[k] += 1
    files = json.loads(j["files"])
    return page(request, "review.html", u, job=j, uid=uid, docs=docs, doc=doc, kinds=KINDS, counts=counts,
                total=total, pending=len(pending(uid)), title=files[0], error=j["error"] or "")


@router.get("/jobs/{uid}/review/rows")
def rows(request: Request, uid: str, doc: str):
    _, j = usable_job(request, uid)
    d = load_doc(uid, doc) if j else None
    if not d:
        return JSONResponse([], status_code=404)
    out = []
    for s in d["segs"]:
        r = d["res"].get(str(s["id"])) or {}
        if not r or r.get("copied"):
            continue
        out.append({"id": str(s["id"]), "page": (s["page"] + 1) if s.get("page") is not None else "",
                    "hu": P.cl(s["text"]), "ru": P.cl(r.get("ru")), "kind": kind(r)})
    return out


@router.get("/jobs/{uid}/review/side", response_class=HTMLResponse)
def side(request: Request, uid: str, doc: str, sid: str):
    u, j = usable_job(request, uid)
    d = load_doc(uid, doc) if j else None
    seg = next((s for s in d["segs"] if str(s["id"]) == sid), None) if d else None
    if not seg:
        return HTMLResponse('<p class="muted">Строка не найдена.</p>', status_code=404)
    r = d["res"].get(sid) or {}
    tr = translator_for(j)
    terms, seen = [], set()
    for e in G.find(P.cl(seg["text"]), tr.entries, limit=20):
        hu = e.hu.split(" / ")[0].strip()                     # плотный глоссарий перечисляет формы через «/»
        ru = re.sub(r"\s*\([^)]*\)", "", e.ru).strip() or e.ru
        if (hu.lower(), ru.lower()) not in seen:
            seen.add((hu.lower(), ru.lower()))
            terms.append({"hu": hu, "ru": ru, "src": "заказчик" if id(e) in tr.main_ids else "комплект"})
    terms = terms[:12]
    before = (r.get("autofix") or {}).get("before") if (r.get("autofix") or {}).get("accepted") else \
        (r.get("rag") or {}).get("before")
    return page(request, "_review_side.html", u, uid=uid, doc=doc, sid=sid, r=r, kind=kind(r), terms=terms,
                before=P.cl(before) if before else "", busy=j["status"] in BUSY)


# ------------------------------------------------------------------ правка
@router.post("/jobs/{uid}/review/edit")
def edit(request: Request, uid: str, doc: str = Form(...), sid: str = Form(...), ru: str = Form(...)):
    u, j = usable_job(request, uid)
    if not j:
        return JSONResponse({"error": "Перевод не найден."}, status_code=404)
    if j["status"] in BUSY:
        return JSONResponse({"error": "Сейчас применяются правки — подождите, пока закончится."}, status_code=409)
    new = ru.strip()
    if not new:
        return JSONResponse({"error": "Пустой перевод не сохраняется. Esc — отменить правку."}, status_code=400)
    with job_lock(uid):
        docs = {f.stem: json.loads(f.read_text(encoding="utf-8")) for f in P.doc_files(work_of(uid))}
        d = docs.get(doc)
        seg = next((s for s in d["segs"] if str(s["id"]) == sid), None) if d else None
        if not seg:
            return JSONResponse({"error": "Строка не найдена."}, status_code=404)
        old = d["res"][sid]["ru"]
        if P.cl(old) == new:
            return {"changed": [], "pending": len(pending(uid))}
        key = P.tmkey(seg["text"])
        changed, touched = [], set()
        for stem, dd in docs.items():                 # одинаковая фраза — одинаковая правка во всех документах
            for s in dd["segs"]:
                r = dd["res"].get(str(s["id"]))
                if r and not r.get("copied") and P.tmkey(s["text"]) == key:
                    for k in ("review", "autofix", "rag"):          # пометки были о прежнем переводе
                        r.pop(k, None)
                    r.update(ru=new, issues=[], edited=True)
                    touched.add(stem)
                    if stem == doc:
                        changed.append({"id": str(s["id"]), "ru": new, "kind": "edited"})
        for stem in touched:
            P.save_doc(work_of(uid), stem, docs[stem])
        # [венгерский, стало, было (самое первое — до всех правок этой фразы), автор]
        prev = next((p for p in pending(uid) if P.tmkey(p[0]) == key), None)
        first_before = prev[2] if prev and len(prev) > 2 else P.cl(old)
        pend = [p for p in pending(uid) if P.tmkey(p[0]) != key] + [[P.cl(seg["text"]), new, first_before, u["id"]]]
        pending_file(uid).write_text(json.dumps(pend, ensure_ascii=False), encoding="utf-8")
    now = time.time()
    db.x("insert into edits(job, doc, seg, hu, ru_before, ru_after, author, created) values (?, ?, ?, ?, ?, ?, ?, ?)",
         (uid, doc, sid, P.cl(seg["text"]), P.cl(old), new, u["id"], now))
    db.x("insert into tm(project, hu_key, ru, status, author, job, created) values (?, ?, ?, 'unconfirmed', ?, ?, ?) "
         "on conflict(project, hu_key, author) do update set ru=excluded.ru, status='unconfirmed', job=excluded.job, "
         "created=excluded.created, decided_by=null, decided_at=null",
         (j["project"], key, new, u["id"], uid, now))
    return {"changed": changed, "pending": len(pend)}


@router.post("/jobs/{uid}/review/apply")
def apply(request: Request, uid: str):
    _, j = usable_job(request, uid)
    if not j:
        return Response("Перевод не найден.", status_code=404)
    if j["status"] in BUSY or not pending(uid):
        return RedirectResponse(f"/jobs/{uid}/review", status_code=303)
    db.x("update jobs set status='queued', mode='edits', error=null where uid=?", (uid,))
    return RedirectResponse("/?m=applying", status_code=303)
