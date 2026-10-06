# -*- coding: utf-8 -*-
"""Глоссарии проекта в БД — слоями и с версией.

Слои:
  customer  — глоссарий заказчика: основа и высший приоритет; меняется только импортом целиком;
  expert    — термины, проверенные экспертом (при первом запуске — глоссарий комплекта из файла проекта);
              «строгий» — обязателен: перевод без него уходит на повтор;
  candidate — кандидаты: из новых комплектов (предпроход терминов) и из правок редакторов; в перевод не идут,
              пока эксперт не примет.
Любое решение эксперта по глоссарию повышает версию проекта. Перевод получает снимок своей версии —
md-файлы в формате движка (data/glossary/<проект>/v<N>/), поэтому его можно воспроизвести."""
import re
import json
import time
from pathlib import Path

from server import config, db

import glossary as G                         # путь к src добавляет server.review / runner

ROWS_HEAD = "| Венгерский | Русский | Комментарий |\n|---|---|---|"


# ------------------------------------------------------------------ проект
def ensure_project(name):
    """Проект в БД; при первом обращении слои заполняются из файлов проекта (глоссарий заказчика и комплекта)."""
    if db.one("select 1 from projects where name=?", (name,)):
        return
    pj = config.PROJECTS / name / "project.json"
    prj = json.loads(pj.read_text(encoding="utf-8"))
    rules = ""
    now = time.time()
    rows = []
    for layer, key in (("customer", "glossary"), ("expert", "terms")):
        if not prj.get(key):
            continue
        f = (pj.parent / prj[key]).resolve()
        if not f.exists():
            continue
        r, entries = parse_md(f.read_text(encoding="utf-8"))
        if layer == "customer":
            rules = r
        rows += [(name, layer, hu, ru, note, sec, int("строгий" in note), "file", now) for hu, ru, note, sec in entries]
    with db.conn() as c:
        c.execute("insert or ignore into projects(name, rules, version, updated) values (?, ?, 1, ?)", (name, rules, now))
        if c.total_changes:
            c.executemany("insert into glossary(project, layer, hu, ru, note, section, strict, source, created) "
                          "values (?, ?, ?, ?, ?, ?, ?, ?, ?)", rows)


def version(name):
    ensure_project(name)
    return db.one("select version from projects where name=?", (name,))["version"]


def bump(name):
    db.x("update projects set version=version+1, updated=? where name=?", (time.time(), name))


# ------------------------------------------------------------------ разбор файлов глоссария
def parse_md(text):
    """md в формате движка → (правила, [(hu, ru, комментарий, раздел)]). Тот же разбор, что glossary.load."""
    rules = text.split("## Правила применения", 1)[1].split("\n## ", 1)[0].strip() if "## Правила применения" in text else ""
    out, section = [], ""
    for ln in text.splitlines():
        if ln.startswith("## "):
            section = ln[3:].strip()
            continue
        if not ln.startswith("|") or ln.startswith("|---") or ln.startswith("| Венгерский"):
            continue
        cells = [c.strip() for c in ln.strip().strip("|").split("|")]
        if len(cells) >= 2 and cells[0] and cells[1] and G._alts(cells[0], cells[1]):
            out.append((cells[0], cells[1], cells[2] if len(cells) > 2 else "", section))
    return rules, out


HEADER_WORDS = {"венгерский", "magyar", "hu", "hungarian", "термин"}


def parse_docx(path):
    """docx заказчика: разделы — заголовки, правила — список под «Правила применения», термины — таблицы
    «венгерский | русский | комментарий» (первая строка-шапка пропускается)."""
    from docx import Document
    from docx.table import Table
    from docx.text.paragraph import Paragraph
    d = Document(path)
    rules, out, section = [], [], ""
    for el in d.element.body.iterchildren():
        tag = el.tag.rsplit("}", 1)[-1]
        if tag == "p":
            p = Paragraph(el, d)
            txt = p.text.strip()
            if not txt:
                continue
            if p.style.name.lower().startswith(("heading", "заголовок")):
                section = txt
            elif section.lower().startswith("правила применения"):
                rules.append("- " + txt)
        elif tag == "tbl":
            for row in Table(el, d).rows:
                cells = [c.text.strip() for c in row.cells]
                if len(cells) < 2 or not cells[0] or not cells[1] or cells[0].lower() in HEADER_WORDS:
                    continue
                if G._alts(cells[0], cells[1]):
                    out.append((cells[0], cells[1], cells[2] if len(cells) > 2 else "", section))
    return "\n".join(rules), out


def parse_xlsx(path):
    """xlsx: первый лист, колонки A — венгерский, B — русский, C — комментарий; строка-шапка пропускается."""
    from openpyxl import load_workbook
    out = []
    for row in load_workbook(path, read_only=True).active.iter_rows(values_only=True):
        cells = [str(c).strip() if c is not None else "" for c in row[:3]] + ["", "", ""]
        if cells[0] and cells[1] and cells[0].lower() not in HEADER_WORDS and G._alts(cells[0], cells[1]):
            out.append((cells[0], cells[1], cells[2], ""))
    return "", out


def parse_upload(path):
    suf = Path(path).suffix.lower()
    if suf == ".md":
        return parse_md(Path(path).read_text(encoding="utf-8"))
    if suf == ".docx":
        return parse_docx(path)
    if suf == ".xlsx":
        return parse_xlsx(path)
    raise ValueError("Глоссарий принимается в md, docx или xlsx.")


# ------------------------------------------------------------------ снимок для движка
def _cell(s):
    return (s or "").replace("|", "/").replace("\n", " ").strip()


def _md(title, rules, rows):
    parts, section = [f"# {title}", ""], None
    if rules:
        parts += ["## Правила применения", "", rules, ""]
    for r in rows:
        if r["section"] != section:
            section = r["section"]
            parts += ["", f"## {section or 'Термины'}", "", ROWS_HEAD]
        note = r["note"]
        if r["strict"] and "строгий" not in note:
            note = ("строгий; " + note).strip("; ")
        if not r["strict"]:
            note = re.sub(r"(?<![\w])строгий(?![\w])(;\s*)?", "", note).strip("; ")
        parts.append(f"| {_cell(r['hu'])} | {_cell(r['ru'])} | {_cell(note)} |")
    return "\n".join(parts) + "\n"


def entries_md(title, rules, entries):
    """md в формате движка из [(hu, ru, комментарий, раздел)] — глоссарий нового проекта (manage.py project add)."""
    rows = [{"hu": hu, "ru": ru, "note": note, "section": sec, "strict": int("строгий" in note)}
            for hu, ru, note, sec in entries]
    return _md(title, rules, rows)


def snapshot(name):
    """(версия, путь к project.json снимка). Снимок версии создаётся один раз и потом переиспользуется."""
    ver = version(name)
    d = config.DATA / "glossary" / name / f"v{ver}"
    pj = d / "project.json"
    if pj.exists():
        return ver, pj
    d.mkdir(parents=True, exist_ok=True)
    rules = db.one("select rules from projects where name=?", (name,))["rules"]
    for layer, fname in (("customer", "customer.md"), ("expert", "terms.md")):
        rows = db.q("select hu, ru, note, section, strict from glossary where project=? and layer=? and status='active' "
                    "order by id", (name, layer))
        (d / fname).write_text(_md(f"Глоссарий {name}: {layer}, версия {ver}", rules if layer == "customer" else "", rows),
                               encoding="utf-8")
    src = json.loads((config.PROJECTS / name / "project.json").read_text(encoding="utf-8"))
    prj = {"glossary": str(d / "customer.md"), "terms": str(d / "terms.md"), "tm": str(d / "tm_unused.json")}
    if src.get("conventions"):
        prj["conventions"] = str((config.PROJECTS / name / src["conventions"]).resolve())
    if src.get("domain"):
        prj["domain"] = src["domain"]
    pj.write_text(json.dumps(prj, ensure_ascii=False, indent=1), encoding="utf-8")
    return ver, pj


# ------------------------------------------------------------------ кандидаты
def known_terms(name):
    """{венгерский в нижнем регистре: русский} действующих терминов заказчика и эксперта."""
    return {r["hu"].lower(): r["ru"] for r in db.q(
        "select hu, ru from glossary where project=? and layer in ('customer', 'expert') and status='active'", (name,))}


def merge_candidates(name, items, source, author):
    """items: [{hu, ru, count?, example?}]. Одинаковая пара (hu, ru) копит частоту, авторов и примеры.
    Отклонённая экспертом пара не всплывает снова. Пара, совпадающая с действующим термином, не нужна;
    если венгерский уже в глоссарии с другим русским — это кандидат-конфликт («правят глоссарий»)."""
    ensure_project(name)
    known = known_terms(name)
    added = 0
    for it in items:
        hu, ru = (it.get("hu") or "").strip(), (it.get("ru") or "").strip()
        if not hu or not ru or not G._alts(hu, ru) or known.get(hu.lower(), "").lower() == ru.lower():
            continue
        row = db.one("select * from glossary where project=? and layer='candidate' and lower(hu)=lower(?) "
                     "and lower(ru)=lower(?)", (name, hu, ru))
        ex = [it["example"]] if it.get("example") else []
        if row is None:
            note = f"в глоссарии: {known[hu.lower()]}" if hu.lower() in known else ""
            db.x("insert into glossary(project, layer, hu, ru, note, source, count, authors, examples, created, created_by) "
                 "values (?, 'candidate', ?, ?, ?, ?, ?, ?, ?, ?, ?)",
                 (name, hu, ru, note, source, int(it.get("count") or 1), json.dumps([author]),
                  json.dumps(ex, ensure_ascii=False), time.time(), author))
            added += 1
        elif row["status"] == "active":
            authors = sorted(set(json.loads(row["authors"])) | {author})
            examples = (json.loads(row["examples"]) + ex)[:3]
            db.x("update glossary set count=count+?, authors=?, examples=? where id=?",
                 (int(it.get("count") or 1), json.dumps(authors), json.dumps(examples, ensure_ascii=False), row["id"]))
    return added


def decide_candidate(cid, user, accept, ru=None, strict=False):
    """Принять: кандидат становится термином эксперта (тот же венгерский у эксперта заменяется), версия растёт.
    Отклонить: пара больше не всплывает."""
    c = db.one("select * from glossary where id=? and layer='candidate'", (cid,))
    if not c or c["status"] != "active":
        return None
    now = time.time()
    if not accept:
        db.x("update glossary set status='rejected', decided_by=?, decided_at=? where id=?", (user, now, cid))
        return c
    ru = (ru or c["ru"]).strip()
    db.x("update glossary set status='rejected', decided_by=?, decided_at=? "
         "where project=? and layer='expert' and status='active' and lower(hu)=lower(?)", (user, now, c["project"], c["hu"]))
    db.x("insert into glossary(project, layer, hu, ru, note, section, strict, source, created, created_by, decided_by, decided_at) "
         "values (?, 'expert', ?, ?, ?, 'Принято экспертом', ?, ?, ?, ?, ?, ?)",
         (c["project"], c["hu"], ru, "", int(strict), c["source"], now, c["created_by"], user, now))
    db.x("update glossary set status='accepted', ru=?, decided_by=?, decided_at=? where id=?", (ru, user, now, cid))
    # тот же венгерский с другими вариантами перевода — решён этим выбором
    db.x("update glossary set status='rejected', decided_by=?, decided_at=? where project=? and layer='candidate' "
         "and status='active' and lower(hu)=lower(?)", (user, now, c["project"], c["hu"]))
    bump(c["project"])
    return c


# ------------------------------------------------------------------ импорт глоссария заказчика
def diff_customer(name, entries):
    """Что изменит импорт: новые, изменённые (тот же венгерский — другой русский), исчезнувшие."""
    ensure_project(name)
    cur = {r["hu"]: r["ru"] for r in db.q("select hu, ru from glossary where project=? and layer='customer' "
                                          "and status='active'", (name,))}
    new = {hu: ru for hu, ru, _, _ in entries}
    return {"added": [(hu, new[hu]) for hu in new if hu not in cur],
            "changed": [(hu, cur[hu], new[hu]) for hu in new if hu in cur and cur[hu] != new[hu]],
            "removed": [(hu, cur[hu]) for hu in cur if hu not in new]}


def replace_customer(name, rules, entries, user):
    now = time.time()
    with db.conn() as c:
        c.execute("delete from glossary where project=? and layer='customer'", (name,))
        c.executemany("insert into glossary(project, layer, hu, ru, note, section, source, created, created_by) "
                      "values (?, 'customer', ?, ?, ?, ?, 'import', ?, ?)",
                      [(name, hu, ru, note, sec, now, user) for hu, ru, note, sec in entries])
        if rules:
            c.execute("update projects set rules=? where name=?", (rules, name))
    bump(name)
