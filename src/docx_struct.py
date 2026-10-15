# -*- coding: utf-8 -*-
"""DOCX по логической структуре размеченного PDF (экспорт Word: StructTreeRoot) — «как будто документ писали
по-русски в Word, взяв за основу оформление оригинала» (требование Максима, 15.10).

pdf2docx восстанавливает положение строк на листе: разделы-страницы, отступы и интервалы под каждую строку, кегль
каждого прогона — переводчик видел «плавающий шрифт и непонятные отступы». Размеченный PDF несёт логику документа:
заголовки уровней, абзацы, списки с типом маркера, таблицы с шапками, рисунки, оглавление. По ней DOCX собирается
стилями (обычный текст, заголовки 1–4, подписи, таблицы, код), с настоящими списками, нумерацией заголовков,
оглавлением и списком рисунков полями Word, колонтитулами.

Модель (load): список блоков в порядке чтения —
  {"k": "h", "level", "runs"}            заголовок
  {"k": "p", "runs", "align", "indent"}  абзац; runs — [{"t": текст | "img": картинка, "b", "i", "font", "size", "color"}]
  {"k": "list", "kind", "items": [[блоки]]}   список (kind: Disc, Circle, Square, Decimal, LowerAlpha, …)
  {"k": "table", "rows": [[{"blocks", "th"}]]}
  {"k": "toc"}                           место оглавления (собирается заново полем)
Текст и картинки элемента — по MCID (страница, номер), упорядочены по номеру: MCID на листе идут в порядке отрисовки."""
import re
import unicodedata
from collections import Counter

HEAD = {"H1": 1, "H2": 2, "H3": 3, "H4": 4, "H5": 5, "H6": 6}
SKIP = {"Artifact"}
INLINE = {"Link", "Span", "Reference", "Annot", "Quote", "Code", "BibEntry", "Note"}


def _fmt(ch):
    name = ch["fontname"].split("+")[-1]
    fam = re.sub(r"(PSMT|PS|MT)$", "", re.split(r"[-,]", name)[0])
    col = ch.get("non_stroking_color")
    if isinstance(col, (list, tuple)) and len(col) == 3:
        color = "%02x%02x%02x" % tuple(int(round(float(v) * 255)) for v in col)
    elif isinstance(col, (list, tuple)) and len(col) == 1:
        color = "%02x%02x%02x" % ((int(round(float(col[0]) * 255)),) * 3)
    else:
        color = "000000"
    return {"font": fam, "size": round(float(ch["size"]) * 2) / 2,
            "b": bool(re.search(r"Bold|Black|Heavy", name)), "i": bool(re.search(r"Italic|Oblique", name)),
            "color": color}


def _runs(items):
    """Символы и картинки одного элемента → прогоны одинакового начертания. Перенос строки — пробел, широкий
    разрыв в строке (табуляция Word) — «\\t»."""
    runs, prev = [], None
    for it in items:
        if "img" in it:
            runs.append({"img": it["img"]}); prev = None; continue
        ch = it["ch"]
        t = ch["text"]
        if t and (0x1D400 <= ord(t[0]) <= 0x1D7FF or 0x2100 <= ord(t[0]) <= 0x214F):
            t = unicodedata.normalize("NFKC", t)        # формула Word (Cambria Math: «𝑏𝑖𝑛á𝑟𝑖𝑠») — обычными буквами, иначе не переводится
        if prev is not None:
            same_line = abs(ch["top"] - prev["top"]) < 0.5 * prev["size"]
            if not same_line and runs and "t" in runs[-1] and not runs[-1]["t"].endswith((" ", "\t")):
                t = " " + t
            elif same_line and ch["x0"] - prev["x1"] > 1.8 * prev["size"] and t != " " \
                    and runs and "t" in runs[-1] and not runs[-1]["t"].endswith("\t"):
                runs[-1]["t"] = runs[-1]["t"].rstrip(" ") + "\t"
                runs[-1].setdefault("tabs", []).append(float(ch["x0"]))     # позиция табуляции — по оригиналу
        f = _fmt(ch)
        if runs and "t" in runs[-1] and all(runs[-1][k] == f[k] for k in f):
            runs[-1]["t"] += t
        else:
            runs.append(dict(f, t=t))
        prev = ch
    # выключка по ширине даёт широкие пробелы в каждом промежутке строки — это не табуляции («Linux⇥alatt⇥például»):
    # табуляция — один-два разрыва в абзаце
    if sum(r.get("t", "").count("\t") for r in runs) >= 3:
        for r in runs:
            if "t" in r:
                r["t"] = r["t"].replace("\t", " ")
                r.pop("tabs", None)
    for r in runs:
        if "t" in r:
            r["t"] = re.sub(r" {2,}", " ", r["t"])
    while runs and "t" in runs[-1] and not runs[-1]["t"].strip():
        runs.pop()
    if runs and "t" in runs[-1]:
        runs[-1]["t"] = runs[-1]["t"].rstrip()
    if runs and "t" in runs[0]:
        runs[0]["t"] = runs[0]["t"].lstrip()
    return [r for r in runs if "img" in r or r["t"]]


def text_of(runs):
    return "".join(r.get("t", "") for r in runs)


NUMBERED = {"Decimal", "LowerAlpha", "UpperAlpha", "LowerRoman", "UpperRoman"}
LABEL = re.compile(r"^\s*(\d{1,3}|[A-Za-z]|[ivxIVX]{2,5})[.)](?=\s|$)\s*")
ROMAN = {"i": 1, "v": 5, "x": 10}


def label_value(lbl, kind):
    """«3» → 3, «c» → 3, «iv» → 4 (по виду списка)."""
    if not lbl:
        return None
    if lbl.isdigit():
        return int(lbl)
    if kind in ("LowerRoman", "UpperRoman") and set(lbl.lower()) <= set(ROMAN):
        vals = [ROMAN[c] for c in lbl.lower()]
        return sum(-v if i + 1 < len(vals) and v < vals[i + 1] else v for i, v in enumerate(vals))
    if len(lbl) == 1 and lbl.isalpha():
        return ord(lbl.lower()) - 96
    return None


def _strip_label(p):
    """Номер пункта, записанный текстом в начале абзаца («a. BER panel…»: Word кладёт номер в тело пункта), — прочь
    из прогонов; возвращает номер. Нумерацию даёт список Word, номер текстом давал «a. a. …»."""
    runs = p["runs"]
    m = LABEL.match(text_of(runs))
    if not m:
        return None
    cut = m.end()
    for r in runs:
        if "t" not in r:
            continue
        n = min(cut, len(r["t"]))
        r["t"] = r["t"][n:]
        cut -= n
        if cut <= 0:
            break
    p["runs"] = [r for r in runs if "img" in r or r["t"]]
    return m.group(1)


def _flatten(lst, depth, out):
    for it in lst["items"]:
        lbl = next((x["t"] for x in it if x["k"] == "lbl"), None)
        body = [x for x in it if x["k"] not in ("list", "lbl")]
        key = next(((x["geom"]["page"], x["geom"]["top"]) for x in body
                    if x["k"] in ("p", "h") and x.get("geom") and x["geom"].get("page")), None)
        out.append({"depth": depth, "kind": lst["kind"], "body": body, "key": key, "lbl": lbl})
        for x in it:
            if x["k"] == "list":
                _flatten(x, depth + 1, out)


def _nest(seq, pos, depth):
    """Плоские пункты (глубина, вид, номер) → вложенные списки. Новый список — при смене вида или когда номер пункта
    не следующий по счёту (список после рисунка продолжает нумерацию: start = номер первого пункта)."""
    lists, cur = [], None
    while pos < len(seq) and seq[pos]["depth"] >= depth:
        it = seq[pos]
        if it["depth"] > depth:
            subs, pos = _nest(seq, pos, it["depth"])
            if cur is None:
                cur = {"k": "list", "kind": it["kind"], "items": [[]], "start": 1}
                lists.append(cur)
            cur["items"][-1].extend(subs)
            continue
        lab = label_value(it["label"], it["kind"])
        if cur is None or cur["kind"] != it["kind"] or (lab is not None and lab != cur["start"] + len(cur["items"])):
            cur = {"k": "list", "kind": it["kind"], "items": [], "start": lab or 1}
            lists.append(cur)
        cur["items"].append(list(it["body"]))
        pos += 1
    return lists, pos


def relist(bs):
    """Списки — по месту пунктов на листе и номерам из текста. Дерево Word на стыке страниц путает вложенность:
    подпункты п. 3–6 лежали в подсписке п. 2, а сами п. 3–6 — отдельным списком следом (UTS, стр. 44–45: «e. a. …»)."""
    out, i = [], 0
    while i < len(bs):
        b = bs[i]
        if b["k"] == "table":
            b = dict(b, rows=[[dict(c, blocks=relist(c["blocks"])) for c in r] for r in b["rows"]])
        if b["k"] != "list":
            out.append(b); i += 1
            continue
        run = []
        while i < len(bs) and bs[i]["k"] == "list":
            run.append(bs[i]); i += 1
        flat = []
        for lst in run:
            _flatten(lst, 0, flat)
        last = (0, 0)
        for n, it in enumerate(flat):
            it["key"] = it["key"] or last
            last, it["n"] = it["key"], n
        flat.sort(key=lambda it: (it["key"], it["n"]))
        for it in flat:
            it["label"] = None
            if it["kind"] in NUMBERED:
                p = next((x for x in it["body"] if x["k"] == "p"), None)
                it["label"] = (_strip_label(p) if p else None) or \
                    (LABEL.match(it["lbl"]).group(1) if it["lbl"] and LABEL.match(it["lbl"]) else None)
            for x in it["body"]:
                if x["k"] == "table":
                    x["rows"] = [[dict(c, blocks=relist(c["blocks"])) for c in r] for r in x["rows"]]
        out += _nest(flat, 0, 0)[0]
    return out


def load(path):
    """Блоки размеченного PDF. Возвращает (блоки, сведения о листе: поля, колонтитулы)."""
    import pdfplumber
    from pdfplumber.structure import PDFStructTree
    pdf = pdfplumber.open(path)
    content, artifacts = {}, {"chars": [], "imgs": []}
    for p in pdf.pages:
        pn = p.page_number
        for ch in p.chars:
            if ch.get("mcid") is None:
                if ch.get("tag") in SKIP or ch.get("tag") is None:
                    artifacts["chars"].append(ch)
                continue
            content.setdefault((pn, ch["mcid"]), []).append({"ch": ch})
        for im in p.images:
            if im.get("mcid") is None:
                artifacts["imgs"].append(im)
                continue
            content.setdefault((pn, im["mcid"]), []).append({"img": im})
    tree = PDFStructTree(pdf)

    def mcids(el):
        out = [(el.page_number, m) for m in (el.mcids or [])]
        for c in el.children:
            out += mcids(c)
        return out

    def items(el):
        keys = sorted(set(mcids(el)), key=lambda k: (k[0] or 0, k[1]))
        out = []
        for k in keys:
            out += content.get(k, [])
        return out

    def geom(el):
        its = items(el)
        chs = [it["ch"] for it in its if "ch" in it and it["ch"]["text"].strip()]
        if not chs:
            ims = [it["img"] for it in its if "img" in it]
            if not ims:
                return None
            im = ims[0]                                     # абзац из картинки: значок в строке или рисунок
            return {"x0": float(im["x0"]), "x1": float(im["x1"]), "first_x0": float(im["x0"]),
                    "page": im["page_number"], "top": float(im["top"]), "bottom": float(im["bottom"]), "size": 10}
        first = [c for c in chs if c["page_number"] == chs[0]["page_number"]]
        return {"x0": min(c["x0"] for c in chs), "x1": max(c["x1"] for c in chs),
                "first_x0": chs[0]["x0"], "page": chs[0]["page_number"], "top": chs[0]["top"],
                "bottom": max(c["bottom"] for c in first), "lines": len({round(c["top"]) for c in first}),
                "size": Counter(round(c["size"]) for c in chs).most_common(1)[0][0]}

    def borders(g):
        """Есть ли у таблицы линии в оригинале (иначе — таблица-раскладка без рамок, как блок адреса на титуле)."""
        if not g:
            return True
        page = pdf.pages[g["page"] - 1]
        x0, x1, top, bot = g["x0"] - 8, g["x1"] + 8, g["top"] - 8, g["bottom"] + 8
        segs = [o for o in page.rects + page.lines if float(o["x0"]) >= x0 and float(o["x1"]) <= x1
                and float(o["top"]) >= top and float(o["bottom"]) <= bot
                and (float(o["width"]) < 3 or float(o["height"]) < 3)]          # линии, не заливки
        return len(segs) >= 2

    def frame(g, rows):
        """Рамка таблицы на первом листе: границы столбцов (x), заливка шапки, поля ячеек — для профиля и сборки."""
        page = pdf.pages[g["page"] - 1]
        top, bot = g["top"] - 30, g["bottom"] + 10
        thin = [o for o in page.rects + page.lines if float(o["bottom"]) >= top and float(o["top"]) <= bot
                and (float(o["width"]) < 3 or float(o["height"]) < 3)]
        # ширина — по верхней линии рамки (у пустого последнего столбца текст кончается задолго до рамки)
        tops = [o for o in thin if float(o["height"]) < 3 and float(o["top"]) <= g["top"] + 2
                and float(o["x1"]) > g["x0"] and float(o["x0"]) < g["x1"]]
        if tops:
            y = max(float(o["top"]) for o in tops)
            row = [o for o in tops if abs(float(o["top"]) - y) < 2]
            x0, x1 = min(float(o["x0"]) for o in row) - 2, max(float(o["x1"]) for o in row) + 2
        else:
            x0, x1 = g["x0"] - 12, g["x1"] + 12
        objs = [o for o in thin if float(o["x0"]) >= x0 and float(o["x1"]) <= x1]
        vert = [o for o in objs if float(o["width"]) < 3 and float(o["height"]) > 5]
        horiz = [o for o in objs if float(o["height"]) < 3 and float(o["width"]) > 20]
        xs = []
        for x in sorted((float(o["x0"]) + float(o["x1"])) / 2 for o in vert):
            if not xs or x - xs[-1] > 3:
                xs.append(x)
        ys = []
        for y in sorted((float(o["top"]) + float(o["bottom"])) / 2 for o in horiz):
            if not ys or y - ys[-1] > 2:
                ys.append(y)
        def hexcol(col):
            if isinstance(col, (int, float)):
                col = (col,)
            if isinstance(col, (list, tuple)) and len(col) in (1, 3):
                v = [float(c) for c in (tuple(col) * 3 if len(col) == 1 else col)]
                return "%02X%02X%02X" % tuple(int(round(c * 255)) for c in v)
            return None
        out = {"xs": [round(x, 1) for x in xs]}
        if vert:
            out["line"] = Counter(hexcol(o.get("non_stroking_color")) for o in vert).most_common(1)[0][0]
        head = rows[0] if rows else []
        hg = next((c["geom"] for c in head if c.get("geom") and c["geom"]["page"] == g["page"]), None)
        if hg and all(c["th"] for c in head):
            fills = [o for o in page.rects if o.get("fill") and float(o["width"]) > 20 and float(o["height"]) > 5
                     and float(o["top"]) <= hg["top"] + 2 <= float(o["bottom"])]
            col = hexcol(fills[0].get("non_stroking_color")) if fills else None
            if col and col != "FFFFFF":
                out["fill"] = col
        pads = {"left": [], "top": [], "bottom": []}
        for r in rows:
            for c in r:
                cg = c.get("geom")
                if not cg or cg["page"] != g["page"]:
                    continue
                lx = [x for x in xs if x <= cg["first_x0"] + 0.5]
                above = [y for y in ys if y <= cg["top"] + 0.5]
                below = [y for y in ys if y >= cg["bottom"] - 0.5]
                if lx:
                    pads["left"].append(cg["first_x0"] - lx[-1])
                if above and below:
                    pads["top"].append(cg["top"] - above[-1]); pads["bottom"].append(below[0] - cg["bottom"])
        for k, v in pads.items():
            if v:
                v.sort(); out["pad_" + k] = round(v[len(v) // 2], 1)
        return out

    def own(el):
        """Абзац из собственных MCID контейнера (текст прямо в LBody/TD, без вложенного P)."""
        its = []
        for m in el.mcids or []:
            its += content.get((el.page_number, m), [])
        return [{"k": "p", "runs": _runs(its), "geom": geom(el)}] if its else []

    def container(el, depth):
        """Содержимое LBody/TD/TH: свой текст, затем вложенные блоки."""
        return own(el) + sum((block(x, depth + 1) for x in el.children), [])

    def block(el, depth=0):
        t = el.type
        if t in HEAD:
            return [{"k": "h", "level": HEAD[t], "runs": _runs(items(el)), "geom": geom(el)}]
        if t in ("P", "Caption", "BlockQuote"):
            return [{"k": "p", "runs": _runs(items(el)), "geom": geom(el)}]
        if t == "Figure":
            return [{"k": "p", "runs": _runs(items(el)), "geom": geom(el), "figure": True}]
        if t == "L":
            kind = (el.attributes or {}).get("ListNumbering", "Disc")
            its = []
            for li in el.children:
                body = []
                for c in li.children:
                    if c.type == "Lbl":
                        body.append({"k": "lbl", "t": text_of(_runs(items(c)))})
                        continue
                    body += container(c, depth) if c.type == "LBody" else block(c, depth + 1)
                its.append(body)
            return [{"k": "list", "kind": str(kind), "items": its}]
        if t == "Table":
            rows = []
            for sec in el.children:
                trs = sec.children if sec.type in ("THead", "TBody", "TFoot") else [sec]
                for tr in trs:
                    if tr.type != "TR":
                        continue
                    row = []
                    for cell in tr.children:
                        if cell.type not in ("TD", "TH"):
                            continue
                        row.append({"blocks": container(cell, depth), "th": cell.type == "TH" or sec.type == "THead",
                                    "geom": geom(cell), "attrs": dict(cell.attributes or {})})
                    rows.append(row)
            cells = [c["geom"] for r in rows for c in r if c.get("geom")]
            g = None
            if cells:
                g = {"page": cells[0]["page"], "top": cells[0]["top"], "x0": min(c["x0"] for c in cells),
                     "x1": max(c["x1"] for c in cells), "first_x0": cells[0]["x0"], "size": cells[0]["size"],
                     "bottom": max(c.get("bottom", c["top"] + 12) for c in cells if c["page"] == cells[0]["page"])}
            return [{"k": "table", "rows": rows, "geom": g, "borders": borders(g), "frame": frame(g, rows) if g else {}}]
        if t in ("TOC", "TOCI"):
            return [{"k": "toc"}] if t == "TOC" else []
        out = []                                            # Document, Sect, Div, Span, Link, Art, Part…
        if t in INLINE:
            # ссылка/фрагмент рядом с абзацем, а не внутри него (перекрёстная ссылка Word «5.6.4.7 fejezet)» в пункте
            # списка) — часть соседнего абзаца, где бы ни оказалась на листе
            return [{"k": "p", "runs": _runs(items(el)), "geom": geom(el), "inline": True}]
        if el.mcids and t not in ("Document", "Sect", "Div", "Part", "Art"):
            out.append({"k": "p", "runs": _runs(items(el)), "geom": geom(el)})
            return out
        for c in el.children:
            out += block(c, depth + 1)
        return out

    blocks = []
    for root in tree:
        blocks += block(root)
    blocks = relist(blocks)
    page = pdf.pages[0]
    info = {"width": float(page.width), "height": float(page.height), "pages": len(pdf.pages),
            "artifacts": artifacts}
    return blocks, info


def layout(path, blocks):
    """Раскладка листа по оригиналу: поля тела и титула, шапка (картинки по краям, строки по центру) и штамп
    внизу (сетка по рамкам ячеек, строки по словам). Колонтитулы — артефакты PDF, в дереве структуры их нет."""
    import pdfplumber
    from statistics import median
    pdf = pdfplumber.open(path)
    W, H = float(pdf.pages[0].width), float(pdf.pages[0].height)
    def flat(bs):
        for b in bs:
            if b["k"] == "list":
                for it in b["items"]:
                    yield from flat(it)
            elif b["k"] == "table":
                continue
            else:
                yield b
    body = [b["geom"] for b in flat(blocks) if b.get("geom") and b["geom"]["page"] and b["geom"]["page"] > 2
            and b["k"] in ("p", "h")]
    left = median(sorted(g["x0"] for g in body)[: max(1, len(body) // 2)]) if body else 62
    right = W - max((g["x1"] for g in body), default=W - 59)
    title = [b["geom"] for b in flat(blocks) if b.get("geom") and b["geom"]["page"] == 1 and b["k"] == "p"
             and any("t" in r and r["t"].strip() for r in b["runs"])]
    page = pdf.pages[min(1, len(pdf.pages) - 1)]
    arts = [c for c in page.chars if c.get("mcid") is None]
    tops = sorted(g["top"] for g in body)
    body_top = min((g["top"] for g in body), default=110)
    head_chars = [c for c in arts if c["top"] < body_top]
    foot_chars = [c for c in arts if c["top"] > H * 0.8]
    head_imgs = [{"xref": i["stream"].objid, "x0": float(i["x0"]), "x1": float(i["x1"]), "top": float(i["top"]),
                  "bottom": float(i["bottom"])} for i in page.images if i.get("mcid") is None and float(i["top"]) < body_top]
    hw = pdfplumber.utils.extract_words(head_chars, x_tolerance=2.5, y_tolerance=2) if head_chars else []
    lines = {}
    for w in hw:
        lines.setdefault(round(float(w["top"]) / 4), []).append(w)
    head_lines = [{"t": " ".join(w["text"] for w in sorted(ws, key=lambda w: w["x0"])),
                   "size": round(float(ws[0].get("size", 10)) if "size" in ws[0] else 11)}
                  for _, ws in sorted(lines.items())]
    # штамп: столбцы — по рамкам ячеек, строки — по словам
    rects = [r for r in page.rects if float(r["top"]) > H * 0.8]
    xs = sorted({round(float(v)) for r in rects for v in (r["x0"], r["x1"])})
    cols = []
    for x in xs:
        if not cols or x - cols[-1] > 8:
            cols.append(x)
    words = pdfplumber.utils.extract_words(foot_chars, x_tolerance=2.5, y_tolerance=2) if foot_chars else []
    rows = {}
    for w in words:
        rows.setdefault(round(float(w["top"]) / 3), []).append(w)
    grid = []
    for _, ws in sorted(rows.items()):
        cells = [""] * max(len(cols) - 1, 1)
        bold = any("Bold" in c["fontname"] for c in foot_chars if abs(c["top"] - ws[0]["top"]) < 2)
        for w in sorted(ws, key=lambda w: w["x0"]):
            ci = max(0, min(len(cells) - 1, sum(1 for x in cols[1:-1] if float(w["x0"]) >= x - 2)))
            cells[ci] = (cells[ci] + " " + w["text"]).strip()
        grid.append({"cells": cells, "bold": bold})
    widths = [cols[i + 1] - cols[i] for i in range(len(cols) - 1)] if len(cols) > 1 else []
    foot_top = min((float(c["top"]) for c in foot_chars), default=H - 100)
    foot_bottom = max((float(c["bottom"]) for c in foot_chars), default=H - 40)
    head_top = min([i["top"] for i in head_imgs] + [float(c["top"]) for c in head_chars] or [36])
    return {"W": W, "H": H, "left": left, "right": max(right, 20), "top": body_top - 4,
            "bottom": H - foot_top + 12, "header": head_top, "footer": max(H - foot_bottom - 6, 18),
            "title_left": Counter(round(g["x0"]) for g in title).most_common(1)[0][0] if title else left,
            "head_imgs": head_imgs, "head_lines": head_lines, "stamp": grid, "stamp_widths": widths,
            "stamp_font": "Times New Roman"}
