# -*- coding: utf-8 -*-
"""Экраны эксперта (роль expert или admin): кандидаты в термины, память переводов, глоссарий, импорт глоссария
заказчика. Каждое решение по глоссарию повышает версию проекта — следующие переводы идут с ним.
Память переводов версии не меняет: очередь читает её из БД при запуске каждого перевода."""
import json
import time
import secrets
import tempfile
from pathlib import Path

from fastapi import APIRouter, Request, Form, UploadFile, File
from fastapi.responses import HTMLResponse, RedirectResponse, Response

from server import config, db, glossary_db as GD
from server.web import need_user, page, projects, NotAllowed

router = APIRouter(prefix="/expert")
PAGE = 100


def need_expert(request: Request):
    u = need_user(request)
    if u["role"] not in ("expert", "admin"):
        raise NotAllowed()
    return u


def project_of(request, project):
    names = projects()
    p = project if project in names else (names[0] if names else "")
    if p:
        GD.ensure_project(p)
    return p


def ctx(request, u, project, tab, **kw):
    return dict(kw, project=project, projects=projects(), tab=tab,
                version=GD.version(project) if project else 0)


@router.get("", response_class=HTMLResponse)
def expert_home(request: Request):
    need_expert(request)
    return RedirectResponse("/expert/candidates", status_code=303)


# ------------------------------------------------------------------ кандидаты
def candidate_rows(project, show, offset):
    rows = db.q("select * from glossary where project=? and layer='candidate' and status='active'", (project,))
    by_hu = {}
    for r in rows:
        by_hu.setdefault(r["hu"].lower(), []).append(r)
    out = []
    for r in rows:
        conflict = len(by_hu[r["hu"].lower()]) > 1 or r["note"].startswith("в глоссарии:")
        if show == "conflicts" and not conflict or show in ("edits", "terms") and r["source"] != show:
            continue
        out.append({"id": r["id"], "hu": r["hu"], "ru": r["ru"], "note": r["note"], "source": r["source"],
                    "count": r["count"], "authors": len(json.loads(r["authors"])), "conflict": conflict,
                    "examples": json.loads(r["examples"])[:2]})
    # сначала то, что подтверждают разные люди и чаще встречается; варианты одного термина — рядом
    out.sort(key=lambda c: (-c["authors"], -c["count"], c["hu"].lower()))
    return out[offset:offset + PAGE], len(out)


@router.get("/candidates", response_class=HTMLResponse)
def candidates(request: Request, project: str = "", show: str = "all", offset: int = 0):
    u = need_expert(request)
    project = project_of(request, project)
    rows, total = candidate_rows(project, show, offset) if project else ([], 0)
    return page(request, "expert_candidates.html", u,
                **ctx(request, u, project, "candidates", rows=rows, total=total, show=show, offset=offset, step=PAGE))


@router.post("/candidates/{cid}", response_class=HTMLResponse)
def decide(request: Request, cid: int, action: str = Form(...), ru: str = Form("")):
    u = need_expert(request)
    c = GD.decide_candidate(cid, u["id"], accept=action in ("accept", "strict"), ru=ru, strict=action == "strict")
    if not c:
        return HTMLResponse('<tr><td colspan="5" class="muted">Уже решено.</td></tr>')
    text = {"accept": "Принято", "strict": "Принято строгим", "reject": "Отклонено"}[action]
    shown = ru.strip() or c["ru"]
    return page(request, "_decided_row.html", u, text=text, hu=c["hu"], ru=shown if action != "reject" else c["ru"],
                ok=action != "reject")


# ------------------------------------------------------------------ память переводов
@router.get("/memory", response_class=HTMLResponse)
def memory(request: Request, project: str = ""):
    u = need_expert(request)
    project = project_of(request, project)
    rows = db.q("select t.*, coalesce(nullif(u.name, ''), u.login) who from tm t join users u on u.id = t.author "
                "where t.project=? and t.status in ('unconfirmed', 'approved') order by t.created desc", (project,))
    groups = {}
    for r in rows:
        g = groups.setdefault(r["hu_key"], {"hu": r["hu_key"], "variants": [], "approved": None})
        if r["status"] == "approved":
            g["approved"] = r["ru"]
        else:
            g["variants"].append({"id": r["id"], "ru": r["ru"], "who": r["who"], "job": r["job"],
                                  "created": time.strftime("%d.%m %H:%M", time.localtime(r["created"]))})
    pending_groups = [g for g in groups.values() if g["variants"]]
    pending_groups.sort(key=lambda g: (-len({v["ru"] for v in g["variants"]}), g["hu"]))   # конфликты — первыми
    return page(request, "expert_memory.html", u, **ctx(request, u, project, "memory", groups=pending_groups[:PAGE],
                                                        total=len(pending_groups)))


@router.post("/memory/{tid}", response_class=HTMLResponse)
def decide_memory(request: Request, tid: int, action: str = Form(...)):
    u = need_expert(request)
    t = db.one("select * from tm where id=?", (tid,))
    if not t or t["status"] != "unconfirmed":
        return HTMLResponse('<p class="muted">Уже решено.</p>')
    now = time.time()
    if action == "approve":
        # одна фраза — один утверждённый перевод: остальные варианты этой фразы отклоняются
        db.x("update tm set status='rejected', decided_by=?, decided_at=? where project=? and hu_key=? and id<>?",
             (u["id"], now, t["project"], t["hu_key"], tid))
        db.x("update tm set status='approved', decided_by=?, decided_at=? where id=?", (u["id"], now, tid))
        return HTMLResponse(f'<p class="ok">Утверждено: {_esc(t["ru"])}</p>')
    db.x("update tm set status='rejected', decided_by=?, decided_at=? where id=?", (u["id"], now, tid))
    return HTMLResponse('<p class="muted">Вариант отклонён.</p>')


def _esc(s):
    return (s or "").replace("&", "&amp;").replace("<", "&lt;").replace(">", "&gt;")


# ------------------------------------------------------------------ глоссарий
def search_terms(project, q, layer):
    sql = "select * from glossary where project=? and layer in ('customer', 'expert') and status='active'"
    args = [project]
    if layer in ("customer", "expert"):
        sql += " and layer=?"
        args.append(layer)
    if q.strip():
        sql += " and (hu like ? or ru like ?)"
        args += [f"%{q.strip()}%"] * 2
    return db.q(sql + " order by layer, hu collate nocase limit ?", args + [PAGE])


@router.get("/glossary", response_class=HTMLResponse)
def glossary_page(request: Request, project: str = "", q: str = "", layer: str = ""):
    u = need_expert(request)
    project = project_of(request, project)
    counts = {r["layer"]: r["n"] for r in db.q("select layer, count(*) n from glossary where project=? and "
                                               "status='active' group by layer", (project,))}
    return page(request, "expert_glossary.html", u, **ctx(request, u, project, "glossary", q=q, layer=layer,
                                                          rows=search_terms(project, q, layer), counts=counts))


@router.get("/glossary/rows", response_class=HTMLResponse)
def glossary_rows(request: Request, project: str, q: str = "", layer: str = ""):
    u = need_expert(request)
    return page(request, "_gloss_rows.html", u, rows=search_terms(project, q, layer), project=project)


@router.post("/glossary/add")
def glossary_add(request: Request, project: str = Form(...), hu: str = Form(...), ru: str = Form(...),
                 strict: bool = Form(False)):
    u = need_expert(request)
    if project in projects() and hu.strip() and ru.strip():
        now = time.time()
        db.x("update glossary set status='rejected', decided_by=?, decided_at=? where project=? and layer='expert' "
             "and status='active' and lower(hu)=lower(?)", (u["id"], now, project, hu.strip()))
        db.x("insert into glossary(project, layer, hu, ru, section, strict, source, created, created_by, decided_by, "
             "decided_at) values (?, 'expert', ?, ?, 'Принято экспертом', ?, 'manual', ?, ?, ?, ?)",
             (project, hu.strip(), ru.strip(), int(strict), now, u["id"], u["id"], now))
        GD.bump(project)
    return RedirectResponse(f"/expert/glossary?project={project}&q={hu.strip()}", status_code=303)


@router.post("/glossary/{gid}", response_class=HTMLResponse)
def glossary_edit(request: Request, gid: int, action: str = Form(...), ru: str = Form(""), strict: bool = Form(False)):
    u = need_expert(request)
    g = db.one("select * from glossary where id=? and layer='expert' and status='active'", (gid,))
    if not g:
        return HTMLResponse('<tr><td colspan="5" class="muted">Запись не найдена или это глоссарий заказчика.</td></tr>')
    now = time.time()
    if action == "delete":
        db.x("update glossary set status='rejected', decided_by=?, decided_at=? where id=?", (u["id"], now, gid))
        GD.bump(g["project"])
        return HTMLResponse(f'<tr><td colspan="5" class="muted">Удалено: {_esc(g["hu"])}</td></tr>')
    if ru.strip() and (ru.strip() != g["ru"] or int(strict) != g["strict"]):
        db.x("update glossary set ru=?, strict=?, decided_by=?, decided_at=? where id=?",
             (ru.strip(), int(strict), u["id"], now, gid))
        GD.bump(g["project"])
    return page(request, "_gloss_rows.html", u, rows=[db.one("select * from glossary where id=?", (gid,))],
                project=g["project"], saved=True)


# ------------------------------------------------------------------ качество
def quality_rows(project):
    """По месяцам и по версиям глоссария: переводов, строк, доля строк с правкой редактора, с подозрением проверки
    и с формальным замечанием. Сводка — сразу после перевода (до правок); правки — строки, исправленные человеком."""
    edited = {r["job"]: r["n"] for r in db.q("select job, count(distinct doc || '/' || seg) n from edits group by job")}
    rows = db.q("select uid, finished, glossary_version, summary_initial from jobs "
                "where project=? and summary_initial is not null order by finished", (project,))
    by = {"month": {}, "version": {}}
    for r in rows:
        s = json.loads(r["summary_initial"])
        for kind, key in (("month", time.strftime("%Y-%m", time.localtime(r["finished"]))),
                          ("version", r["glossary_version"] or 1)):
            g = by[kind].setdefault(key, {"key": key, "jobs": 0, "segments": 0, "edited": 0, "suspicious": 0, "issues": 0})
            g["jobs"] += 1
            g["segments"] += s.get("segments", 0)
            g["suspicious"] += s.get("suspicious", 0)
            g["issues"] += s.get("issues", 0)
            g["edited"] += edited.get(r["uid"], 0)
    out = {}
    for kind, groups in by.items():
        lst = []
        for g in groups.values():
            n = max(g["segments"], 1)
            lst.append(dict(g, p_edited=100 * g["edited"] / n, p_susp=100 * g["suspicious"] / n,
                            p_issues=100 * g["issues"] / n))
        out[kind] = sorted(lst, key=lambda g: g["key"])
    return out


@router.get("/quality", response_class=HTMLResponse)
def quality(request: Request, project: str = ""):
    u = need_expert(request)
    project = project_of(request, project)
    return page(request, "expert_quality.html", u, **ctx(request, u, project, "quality", q=quality_rows(project)))


# ------------------------------------------------------------------ импорт глоссария заказчика
@router.get("/import", response_class=HTMLResponse)
def import_page(request: Request, project: str = ""):
    u = need_expert(request)
    project = project_of(request, project)
    return page(request, "expert_import.html", u, **ctx(request, u, project, "import", preview=None))


@router.post("/import", response_class=HTMLResponse)
def import_preview(request: Request, project: str = Form(...), file: UploadFile = File(...)):
    u = need_expert(request)
    project = project_of(request, project)
    suf = Path(file.filename or "").suffix.lower()
    with tempfile.NamedTemporaryFile(suffix=suf, delete=False) as tmp:
        tmp.write(file.file.read())
    try:
        rules, entries = GD.parse_upload(tmp.name)
    except ValueError as e:
        return page(request, "expert_import.html", u, **ctx(request, u, project, "import", preview=None, error=str(e)))
    finally:
        Path(tmp.name).unlink(missing_ok=True)
    if not entries:
        return page(request, "expert_import.html", u, **ctx(request, u, project, "import", preview=None,
                                                            error="В файле не нашлось ни одной пары «венгерский — русский»."))
    token = secrets.token_hex(8)
    imp = config.DATA / "imports"
    imp.mkdir(parents=True, exist_ok=True)
    (imp / f"{token}.json").write_text(json.dumps({"project": project, "rules": rules, "entries": entries,
                                                   "file": file.filename}, ensure_ascii=False), encoding="utf-8")
    diff = GD.diff_customer(project, entries)
    return page(request, "expert_import.html", u, **ctx(request, u, project, "import", token=token, file=file.filename,
                                                        preview=diff, n=len(entries), rules=bool(rules)))


@router.post("/import/confirm")
def import_confirm(request: Request, token: str = Form(...)):
    u = need_expert(request)
    f = config.DATA / "imports" / f"{Path(token).name}.json"
    if not f.exists():
        return RedirectResponse("/expert/import", status_code=303)
    d = json.loads(f.read_text(encoding="utf-8"))
    GD.replace_customer(d["project"], d["rules"], [tuple(e) for e in d["entries"]], u["id"])
    f.unlink()
    return RedirectResponse(f"/expert/glossary?project={d['project']}&layer=customer", status_code=303)
