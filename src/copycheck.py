# -*- coding: utf-8 -*-
"""Сверка копии скана тремя чтениями (Максим, 12.11).

Копия (rucopy) — чтение листа порталом в HTML. На «Декларации» оно спутало Н/И, Т/Г, Л/1, 6/2 там, где простое чтение
того же листа и Surya прочли верно. Сверка: копия против простого чтения портала (copy_read/pNNNN.txt) и Surya
(tools/surya_read.py). По каждому расхождению:
- оба других чтения совпали, копия — нет: слово копии исправляется (примечание Word «было … стало …»);
- копия совпала хотя бы с одним: остаётся как есть;
- согласия нет: слово выделяется жёлтым и получает примечание Word с вариантами.
Выделение переходит в перевод своей группой (segments.run_sig), примечание остаётся по-русски (comments.xml не
переводится): сомнительное Максим видит и правит в английском Word.
Сравнение — без разницы кавычек «» и "" и латинских двойников кириллицы (C/С, P/Р, …): Surya пишет «04CEJI0»."""
import difflib
import json
import os
import re
import subprocess
import sys
import tempfile
from pathlib import Path

W = "{http://schemas.openxmlformats.org/wordprocessingml/2006/main}"
FOLD = str.maketrans("ABCEHKMOPTXYaceopxy", "АВСЕНКМОРТХУасеорху")
QUOTES = str.maketrans({"«": '"', "»": '"', "„": '"', "“": '"', "”": '"'})
MAX_SPAN = 8            # расхождение длиннее (порядок строк, пропуск абзаца) — не правится и не выделяется: шум
AUTHOR = "сверка"


def fold(t):
    return re.sub(r"\s+", " ", t.translate(QUOTES).translate(FOLD)).strip()


def merge_fix(cp, new):
    """Исправление слова копии буквами другого чтения, а кавычки и латиница/кириллица-двойники — как в копии: второе
    чтение пишет "…" вместо «…», Surya — латинские двойники."""
    out = []
    for tag, i1, i2, j1, j2 in difflib.SequenceMatcher(None, cp, new, autojunk=False).get_opcodes():
        if tag == "equal" or (tag == "replace" and i2 - i1 == j2 - j1 and fold(cp[i1:i2]) == fold(new[j1:j2])):
            out.append(cp[i1:i2])
        elif tag != "delete":
            out.append(new[j1:j2])
    return "".join(out)


def paragraphs(doc):
    """Абзацы копии по порядку (и абзацы ячеек таблиц) с номером листа: лист кончается разрывом листа или раздела."""
    from docx.text.paragraph import Paragraph
    page, out = 0, []
    for el in doc.element.body.iterchildren():
        if el.tag == W + "p":
            out.append((page, Paragraph(el, doc._body)))
            if el.find(f".//{W}br[@{W}type='page']") is not None or el.find(f"{W}pPr/{W}sectPr") is not None:
                page += 1
        elif el.tag == W + "tbl":
            out += [(page, Paragraph(p, doc._body)) for p in el.iter(W + "p")]
    return out


def copy_tokens(doc):
    """{лист: [(абзац, начало, конец, слово)]} — слова копии по сквозному тексту прогонов абзаца."""
    pages = {}
    for page, par in paragraphs(doc):
        text = "".join(r.text for r in par.runs)
        # заглушка копии «[логотип]», «[печать]» — не текст листа: второе чтение переписывает надписи логотипа, и
        # заглушка выходила «спорной» (жёлтое [logo] в переводе, Информационное письмо 12.11)
        pages.setdefault(page, []).extend((par, m.start(), m.end(), m.group()) for m in re.finditer(r"\S+", text)
                                          if not re.fullmatch(r"\[[^\]]+\]", m.group()))
    return pages


def surya_tokens(lines):
    """Строки Surya листа → [(слово, уверенность строки)] в порядке чтения: сверху вниз, в строке — слева направо."""
    lines = sorted(lines, key=lambda l: (round(l["box"][1] / 12), l["box"][0]))
    return [(w, l["c"]) for l in lines for w in l["t"].split()]


def aligned(copy, other):
    """Опкоды выравнивания слов копии со словами другого чтения (по свёрнутому виду)."""
    return difflib.SequenceMatcher(None, [fold(w) for w in copy], [fold(w) for w in other], autojunk=False).get_opcodes()


def span_text(ops, other, a, b):
    """Текст другого чтения на месте слов копии [a, b): у совпавших — слово в слово, у заменённых — весь кусок."""
    out = []
    for tag, i1, i2, j1, j2 in ops:
        if i2 <= a or i1 >= b:
            if tag == "insert" and a < i1 < b:                 # вставка внутри куска
                out += other[j1:j2]
            continue
        if tag == "equal":
            out += other[j1 + max(a, i1) - i1:j1 + min(b, i2) - i1]
        elif tag == "replace":
            out += other[j1:j2]
    return out


def disputes(copy, plain, sur):
    """Куски копии [a, b), где хоть одно чтение расходится с копией (слитые пересекающиеся)."""
    spans = []
    for ops in (aligned(copy, plain), aligned(copy, [w for w, _ in sur])):
        for tag, i1, i2, j1, j2 in ops:
            if tag in ("replace", "delete"):
                spans.append((i1, i2))
    spans.sort()
    out = []
    for a, b in spans:
        if out and a <= out[-1][1]:
            out[-1] = (out[-1][0], max(out[-1][1], b))
        else:
            out.append((a, b))
    return out


def isolate(par, start, end):
    """Прогоны абзаца ровно на знаках [start, end): прогоны на краях делятся. → [Run]."""
    from copy import deepcopy
    from docx.text.run import Run

    def split_at(pos):
        acc = 0
        for r in par.runs:
            t = r.text
            if acc < pos < acc + len(t):
                el = deepcopy(r._r)
                r._r.addnext(el)
                tail = Run(el, r._parent)
                r.text, tail.text = t[:pos - acc], t[pos - acc:]
                return
            acc += len(t)
    split_at(start)
    split_at(end)
    out, acc = [], 0
    for r in par.runs:
        t = r.text
        if t and start <= acc and acc + len(t) <= end:
            out.append(r)
        acc += len(t)
    return out


EDGE = "\"'«»„“”()[].,;:!?…"


def core(t):
    """Слово без знаков по краям: запятую от точки на скане не различить (после «630099» — запятая, после «Новосибирск»
    на вид — точка, Декларация 12.11); расхождение только в знаках препинания не правится и не выделяется."""
    return " ".join(w.strip(EDGE) for w in fold(t).split())


def decide(cp, pt, st):
    """Решение по куску: None — копия верна или спорить не о чем; ("fix", новое) — оба других чтения совпали, копия
    нет; ("mark", None) — согласия нет."""
    c, p, q = core(cp), core(pt), core(st)
    if c in (p, q):
        return None                                       # копия в большинстве (или разница только в знаках)
    if not p and not q:
        return None                                       # в других чтениях этих слов нет вовсе (под печатью, пропуск)
    if p and p == q:
        return "fix", pt
    return "mark", None


def uneven(copy, a, b, pts, sts):
    """Кусок разной длины в чтениях («включительно (подпись)» против «включительно» и «вк,почительно»): слова копии,
    совпавшие со вторым чтением, верны (вместе с похожим словом Surya); спорят только остальные — одним куском (от
    первого до последнего), без них — спора нет. → [(начало, конец, слова второго, слова Surya)]."""
    pl, sl, left = list(pts), list(sts), []
    for i in range(a, b):
        c = core(copy[i])
        k = next((k for k, w in enumerate(pl) if core(w) == c), None)
        if k is None:
            left.append(i)
            continue
        pl.pop(k)
        near = max(range(len(sl)), key=lambda k: difflib.SequenceMatcher(None, core(sl[k]), c).ratio(), default=None)
        if near is not None and difflib.SequenceMatcher(None, core(sl[near]), c).ratio() >= 0.75:
            sl.pop(near)
    return [(left[0], left[-1] + 1, pl, sl)] if left else []


def check(doc, plain_pages, surya_pages):
    """Сверка копии на месте. plain_pages — {лист: текст}, surya_pages — {лист: строки Surya}. → (исправлено, выделено)."""
    from docx.enum.text import WD_COLOR_INDEX
    fixed = marked = 0
    todo = []                                             # ([абзац, начало, конец], действие, примечание, замена)
    for page, toks in copy_tokens(doc).items():
        copy = [t[3] for t in toks]
        plain = (plain_pages.get(page) or "").split()
        sur = surya_tokens(surya_pages.get(page) or [])
        if not plain or not sur:
            continue
        sw = [w for w, _ in sur]
        conf_of = {}
        for w, c in sur:
            conf_of[w] = min(c, conf_of.get(w, 1.0))
        po, so = aligned(copy, plain), aligned(copy, sw)
        for a, b in disputes(copy, plain, sur):
            if b - a > MAX_SPAN:
                continue
            pts, sts = span_text(po, plain, a, b), span_text(so, sw, a, b)
            # кусок одной длины во всех чтениях — по слову: соседнее верное слово не выделяется вместе с сомнительным
            units = [(i, i + 1, [pts[i - a]], [sts[i - a]]) for i in range(a, b)] \
                if len(pts) == len(sts) == b - a else uneven(copy, a, b, pts, sts)
            for i, j, pu, su in units:
                cp, pt, st = " ".join(copy[i:j]), " ".join(pu), " ".join(su)
                act = decide(cp, pt, st)
                if act is None:
                    continue
                pars = {}
                for par, s0, e0, _ in toks[i:j]:
                    pars.setdefault(id(par), [par, s0, e0])[2] = e0
                parts = list(pars.values())
                if act[0] == "fix" and len(parts) == 1:
                    todo.append((parts[0], "fix", f"Сверка: было «{cp}», второе чтение и Surya — «{pt}».", pt))
                    continue
                conf = min((conf_of[w] for w in su if w in conf_of), default=None)
                sc = f" (уверенность Surya {conf:.2f})" if conf is not None else ""
                note = (f"Проверить по скану, лист {page + 1}. Чтения: копия «{cp}» · второе «{pt or '—'}» · "
                        f"Surya «{st or '—'}»{sc}.")
                todo += [(part, "mark", note, None) for part in parts]
    # с конца абзаца к началу: правка не сдвигает ещё не обработанные места
    todo.sort(key=lambda t: (id(t[0][0]), -t[0][1]))
    for (par, s0, e0), act, note, new in todo:
        runs = isolate(par, s0, e0)
        if not runs:
            continue
        if act == "fix":
            runs[0].text = merge_fix(runs[0].text + "".join(r.text for r in runs[1:]), new)
            for r in runs[1:]:
                r.text = ""
            doc.add_comment(runs[0], text=note, author=AUTHOR, initials="СВ")
            fixed += 1
        else:
            for r in runs:
                r.font.highlight_color = WD_COLOR_INDEX.YELLOW
            doc.add_comment(runs, text=note, author=AUTHOR, initials="СВ")
            marked += 1
    return fixed, marked


def host_path(p):
    """Путь из контейнера (/app/data/…) — на этой машине (data рядом с кодом или HUT_DATA)."""
    data = Path(os.environ.get("HUT_DATA", Path(__file__).resolve().parent.parent / "data"))
    return data / Path(p).relative_to("/app/data") if str(p).startswith("/app/data") else Path(p)


SURYA_DPI = 200


def run_surya(pdf, pages, python):
    """Листы PDF (200 dpi) → Surya в отдельном питоне (tools/surya_read.py). → {лист: строки}."""
    import fitz
    with tempfile.TemporaryDirectory() as td:
        pngs = []
        with fitz.open(str(pdf)) as d:
            for p in range(pages):
                f = Path(td) / f"p{p:04d}.png"
                d[p].get_pixmap(dpi=SURYA_DPI).save(str(f))
                pngs.append(str(f))
        out = Path(td) / "surya.json"
        r = subprocess.run([python, "-I", str(Path(__file__).resolve().parent.parent / "tools" / "surya_read.py"),
                            str(out), *pngs], capture_output=True, text=True, encoding="utf-8", errors="replace",
                           env=dict(os.environ, HF_HUB_OFFLINE="1"))
        if r.returncode != 0 or not out.exists():
            raise RuntimeError("Surya: " + (r.stderr or r.stdout).strip()[-400:])
        return dict(enumerate(json.loads(out.read_text(encoding="utf-8"))))


def main(copy_docx, read_dir, out_docx, surya_python):
    from docx import Document
    rd = Path(read_dir)
    meta = json.loads((rd / "pages.json").read_text(encoding="utf-8"))
    plain = {p: (rd / f"p{p:04d}.txt").read_text(encoding="utf-8") for p in range(meta["pages"])
             if (rd / f"p{p:04d}.txt").exists()}
    sur = run_surya(host_path(meta["pdf"]), meta["pages"], surya_python)
    import copylayout
    doc = Document(str(copy_docx))
    # раскладка по строкам Surya — до сверки текста: деление абзацев по переносам не сдвигает её правок
    laid, linked = copylayout.apply(doc, sur, SURYA_DPI)
    fixed, marked = check(doc, plain, sur)
    doc.save(str(out_docx))
    print(f"сверка: {Path(copy_docx).name} — раскладка со скана {laid}, в одну строку {linked}, "
          f"исправлено {fixed}, выделено {marked}")
    return fixed, marked


if __name__ == "__main__":
    sys.path.insert(0, str(Path(__file__).resolve().parent))         # и под python -I (свои модули — рядом)
    if len(sys.argv) != 5:
        sys.exit("usage: copycheck.py <копия.docx> <work/copy_read/<имя>> <выход.docx> <python с Surya>")
    main(*sys.argv[1:5])
