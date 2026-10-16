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


# начертание в имени шрифта PDF — и по-английски, и по-местному: Word 2016 с венгерским интерфейсом пишет
# «Times New Roman félkövér» — это полужирный Times, а не шрифт с таким именем (ÜFK, 16.10)
BOLD_W = r"Bold|Black|Heavy|Semibold|Demibold|félkövér|Fett|полужирный|Gras|Negrita|Grassetto"
ITAL_W = r"Italic|Oblique|dőlt|Kursiv|курсив|Italique|Cursiva|Corsivo"


def _fmt(ch):
    name = ch["fontname"].split("+")[-1]
    fam = re.split(r"[-,]", name)[0]
    fam = re.sub(rf"\s+(?:{BOLD_W}|{ITAL_W}|Regular|Normál|Обычный)\b.*$", "", fam, flags=re.I).strip()
    fam = re.sub(r"(PSMT|PS|MT)$", "", fam)
    if " " not in fam:                          # «TimesNewRoman», «CourierNew» — имя шрифта Word с пробелами
        fam = re.sub(r"(?<=[a-z])(?=[A-Z])", " ", fam)
    col = ch.get("non_stroking_color")
    if isinstance(col, (list, tuple)) and len(col) == 3:
        color = "%02x%02x%02x" % tuple(int(round(float(v) * 255)) for v in col)
    elif isinstance(col, (list, tuple)) and len(col) == 1:
        color = "%02x%02x%02x" % ((int(round(float(col[0]) * 255)),) * 3)
    else:
        color = "000000"
    # курсив без курсивного начертания (у Tahoma его нет): Word наклоняет знаки матрицей — c ≠ 0 при b = 0
    m = ch.get("matrix") or (1, 0, 0, 1)
    slant = abs(float(m[1])) < 1e-6 and abs(float(m[2])) > 0.1 * abs(float(m[3]) or 1)
    return {"font": fam, "size": round(float(ch["size"]) * 2) / 2,
            "b": bool(re.search(BOLD_W, name, re.I)), "i": bool(re.search(ITAL_W, name, re.I)) or slant,
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
        else:                                   # начало прогона на листе — место для значка в строке (insert_imgs)
            runs.append(dict(f, t=t, x=float(ch["x0"]), top=float(ch["top"]), page=ch.get("page_number")))
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


def insert_imgs(runs, imgs):
    """Значки — на своё место в строке по листу: перед первым прогоном текста той же строки, который начинается
    правее значка; нет такого — в конец. Word рисует картинки строки после текста, и по порядку отрисовки значок
    пункта «● : не подключено» уезжал в конец: «: не подключено●» (UTS, 27 пунктов списков, 16.10)."""
    out = list(runs)
    for im in imgs:
        g = im["img"]
        cy = (float(g["top"]) + float(g["bottom"])) / 2
        pos = next((i for i, r in enumerate(out) if "x" in r and r.get("page") == g.get("page_number")
                    and r["top"] - 2 <= cy <= r["top"] + 1.5 * r["size"] and r["x"] >= float(g["x0"]) - 1), None)
        if pos is None:
            out.append(im)
        else:
            out.insert(pos, im)
    return out


def join_tables(bs):
    """Таблица, которую Word при экспорте разрезал по листам (продолжение — отдельная таблица сразу за ней, без шапки,
    с тем же числом столбцов, на следующем листе), — снова одна: в DOCX она переносится сама и повторяет шапку
    (UTS стр. 75–78: разрыв с пустой строкой и без шапки, 16.10)."""
    out = []
    for b in bs:
        if b["k"] == "list":
            b = dict(b, items=[join_tables(it) for it in b["items"]])
        # пометка о переносе («Folytatás a következő oldalon.») между таблицей и её продолжением без шапки на следующем
        # листе — прочь: в DOCX таблица переносится сама (ÜFK, 16.10)
        if b["k"] == "table" and len(out) >= 2 and out[-1]["k"] == "p" and out[-2]["k"] == "table" \
                and len(text_of(out[-1]["runs"]).strip()) < 60 and b["rows"] and out[-2]["rows"] \
                and not all(c["th"] for c in b["rows"][0]) \
                and (out[-2].get("geom") or {}).get("page", 0) <= (out[-1].get("geom") or {}).get("page", -1) \
                < (b.get("geom") or {}).get("page", 0):
            out.pop()
        prev = out[-1] if out else None
        if prev is not None and prev["k"] == "table" and b["k"] == "table" and prev["rows"] and b["rows"] \
                and not all(c["th"] for c in b["rows"][0]) \
                and max(map(len, b["rows"])) == max(map(len, prev["rows"])) \
                and (b.get("geom") or {}).get("page", 0) > (prev.get("geom") or {}).get("page", 0):
            out[-1] = dict(prev, rows=prev["rows"] + b["rows"])
            continue
        out.append(b)
    return out


def _paras(bs):
    for b in bs:
        if b["k"] == "p":
            yield b
        elif b["k"] == "list":
            for it in b["items"]:
                yield from _paras(it)
        elif b["k"] == "table":
            for r in b["rows"]:
                for c in r:
                    yield from _paras(c["blocks"])


def attach_icons(blocks, imgs, height):
    """Значок, который Word пометил оформлением (без MCID), но который стоит в первой строке абзаца, — в этот абзац
    на своё место (UTS: серый значок «nem meghatározható» пропадал из пункта). Колонтитулы, рисунки — не значки."""
    ps = [p for p in _paras(blocks) if p.get("geom") and any(r.get("t", "").strip() for r in p["runs"])]
    for im in imgs:
        top, bot, x0 = float(im["top"]), float(im["bottom"]), float(im["x0"])
        if bot - top > 20 or top < 0.12 * height or bot > 0.88 * height:
            continue
        cy = (top + bot) / 2
        cand = [p for p in ps if p["geom"]["page"] == im["page_number"]
                and p["geom"]["top"] - 3 <= cy <= p["geom"]["top"] + 3 + 1.5 * p["geom"]["size"]
                and p["geom"]["x0"] - 40 <= x0 <= p["geom"]["x1"]]
        if cand:
            p = min(cand, key=lambda p: abs(p["geom"]["x0"] - x0))
            p["runs"] = insert_imgs(p["runs"], [{"img": im}])


NUMBERED = {"Decimal", "LowerAlpha", "UpperAlpha", "LowerRoman", "UpperRoman"}
LABEL = re.compile(r"^\s*(\d{1,3}|[A-Za-z]|[ivxIVX]{2,5})(\.\)|[.)])(?=\s|$)\s*")
DASH = re.compile(r"^\s*[–—\-](?=\s)")


def text_labels(bs):
    """Маркированный по дереву список, у всех пунктов которого номер или тире стоят прямо в тексте (Word не вынес
    их в Lbl: ÜFK «a. kampányonként …», «– ha az …»), — буквенный, нумерованный или «тире» с тем же видом метки
    («a.» / «a.)»): иначе в DOCX «• a. …» — двойной маркер (16.10). Номера из текста снимает relist."""
    for b in bs:
        if b["k"] == "table":
            for r in b["rows"]:
                for c in r:
                    text_labels(c["blocks"])
        if b["k"] != "list":
            continue
        for it in b["items"]:
            text_labels(it)
        if b["kind"] in NUMBERED or not b["items"]:
            continue
        firsts = [next((x for x in it if x["k"] == "p"), None) for it in b["items"]]
        texts = [text_of(p["runs"]) if p else "" for p in firsts]
        if all(DASH.match(t) for t in texts):
            b["kind"] = "Dash"                              # тире из текста снимает сборка (LEAD)
            continue
        ms = [LABEL.match(t) for t in texts]
        if not all(ms):
            continue
        labs = [m.group(1) for m in ms]
        kind = ("Decimal" if all(x.isdigit() for x in labs) else
                "LowerAlpha" if all(len(x) == 1 and x.islower() for x in labs) else
                "UpperAlpha" if all(len(x) == 1 and x.isupper() for x in labs) else None)
        vals = [label_value(x, kind) for x in labs] if kind else []
        if kind and all(v2 == v1 + 1 for v1, v2 in zip(vals, vals[1:])):
            b["kind"], b["suffix"] = kind, ms[0].group(2)
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
        out.append({"depth": depth, "kind": lst["kind"], "body": body, "key": key, "lbl": lbl,
                    "suffix": lst.get("suffix")})
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
            if it.get("suffix"):
                cur["suffix"] = it["suffix"]
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
        return geom_of(items(el))

    def geom_of(its):
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

    def toc_info(el):
        """Оглавление оригинала, как его собрать заново полем Word: заголовок (текст, кегль, цвет, по центру ли), число
        уровней и какие из них полужирные (по отступам строк), пункты без номера раздела — абзацы вне заголовков,
        которым нужен уровень структуры (UTS: «Dokumentumhoz kapcsolódó szerepek» — шапка таблицы, 1-й уровень);
        список рисунков — если строки «N. ábra …»."""
        ents = []
        for e in el.children:
            its = [it for it in items(e) if "ch" in it]
            chs = [it["ch"] for it in its if it["ch"]["text"].strip()]
            if chs:
                txt = re.sub(r"\s+", " ", "".join(it["ch"]["text"] for it in its)).strip()
                ents.append({"t": txt, "x0": float(chs[0]["x0"]), "cx": (chs[0]["x0"] + chs[-1]["x1"]) / 2,
                             "f": _fmt(chs[0]), "num": bool(re.search(r"\d+\s*$", txt)),
                             "page": chs[0]["page_number"], "top": float(chs[0]["top"])})
        title = ents[0] if ents and not ents[0]["num"] else None
        body = [e for e in ents if e["num"]]
        lv = []
        for x in sorted(e["x0"] for e in body):
            if not lv or x - lv[-1] > 4:
                lv.append(x)
        level = lambda x: min(range(len(lv)), key=lambda i: abs(lv[i] - x)) + 1
        bold = []
        for i in range(len(lv)):
            at = [e for e in body if level(e["x0"]) == i + 1]
            bold.append(sum(e["f"]["b"] for e in at) * 2 > len(at))
        figs = sum(bool(re.match(r"\d+\.\s*ábra", e["t"], re.I)) for e in body) * 2 > len(body)
        extra = [] if figs else [(re.sub(r"\s*(\.{3,}\s*)?\d+\s*$", "", e["t"]).strip(), level(e["x0"]))
                                 for e in body if not re.match(r"\d+(\.\d+)*\s", e["t"])]
        width = float(pdf.pages[0].width)
        first_page = ents[0]["page"] if ents else None
        steps = sorted(b["top"] - a["top"] for a, b in zip(body, body[1:]) if a["page"] == b["page"] and b["top"] > a["top"])
        return {"title": title and dict(title["f"], t=title["t"], center=abs(title["cx"] - width / 2) < 30),
                "levels": len(lv), "bold": bold, "figures": figs, "extra": extra,
                "indent": [round(x - lv[0], 1) for x in lv],                      # отступ уровня от 1-го, pt
                "pitch": round(steps[len(steps) // 2], 1) if steps else None,     # шаг строк оглавления, pt
                "page": first_page}

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
            if t != "TOC":
                return []
            info = toc_info(el)                 # лист — чтобы сборка видела конец титула (ÜFK: оглавление на 2-м листе)
            return [{"k": "toc", "info": info, "geom": {"page": info["page"]} if info.get("page") else None}]
        out = []                                            # Document, Sect, Div, Span, Link, Art, Part…
        if t in INLINE:
            # ссылка/фрагмент рядом с абзацем, а не внутри него (перекрёстная ссылка Word «5.6.4.7 fejezet)» в пункте
            # списка) — часть соседнего абзаца, где бы ни оказалась на листе
            return [{"k": "p", "runs": _runs(items(el)), "geom": geom(el), "inline": True}]
        if el.mcids and t not in ("Document", "Sect", "Div", "Part", "Art"):
            out.append({"k": "p", "runs": _runs(items(el)), "geom": geom(el)})
            return out
        if el.mcids and t != "Document":
            # своё содержимое обёртки (Word кладёт значок пункта списка прямо в Sect внутри LBody) — абзацем перед
            # вложенными; раньше терялось (UTS: серый значок «nem meghatározható», 16.10)
            its = [x for m in el.mcids for x in content.get((el.page_number, m), [])]
            if its:
                out.append({"k": "p", "runs": _runs(its), "geom": geom_of(its)})
        for c in el.children:
            out += block(c, depth + 1)
        return out

    blocks = []
    for root in tree:
        blocks += block(root)
    page = pdf.pages[0]
    attach_icons(blocks, artifacts["imgs"], float(page.height))
    text_labels(blocks)
    blocks = join_tables(relist(blocks))
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
    body_top = min((g["top"] for g in body), default=110)
    page, page_tok = _typical_page(pdf, body_top, H)
    arts = [c for c in page.chars if c.get("mcid") is None]
    head_chars = [c for c in arts if c["top"] < body_top]
    ftop = _foot_top(arts, page.rects + page.lines, H)
    foot_chars = [c for c in arts if c["top"] >= ftop]
    head_imgs = [{"xref": i["stream"].objid, "x0": float(i["x0"]), "x1": float(i["x1"]), "top": float(i["top"]),
                  "bottom": float(i["bottom"])} for i in page.images if i.get("mcid") is None and float(i["top"]) < body_top]
    hw = pdfplumber.utils.extract_words(head_chars, x_tolerance=2.5, y_tolerance=2) if head_chars else []
    lines = {}
    for w in hw:
        lines.setdefault(round(float(w["top"]) / 4), []).append(w)
    head_lines = [{"t": " ".join(w["text"] for w in sorted(ws, key=lambda w: w["x0"])),
                   "size": round(float(ws[0].get("size", 10)) if "size" in ws[0] else 11)}
                  for _, ws in sorted(lines.items())]
    # штамп: столбцы — по рамкам ячеек (нет внутренних линий — по разрывам между группами слов), строки — по словам
    foot = _zone_grid(foot_chars, [r for r in page.rects + page.lines if float(r["top"]) >= ftop])
    # шапка без логотипов (ÜFK: строка-таблица «Üzemeltetési Feltételek … | 1.2. változat») — такой же сеткой
    head = _zone_grid(head_chars, [r for r in page.rects + page.lines if float(r["bottom"]) < body_top]) \
        if head_chars and not head_imgs else None
    foot_top = min((float(c["top"]) for c in foot_chars), default=H - 100)
    foot_bottom = max((float(c["bottom"]) for c in foot_chars), default=H - 40)
    head_top = min([i["top"] for i in head_imgs] + [float(c["top"]) for c in head_chars] or [36])
    return {"W": W, "H": H, "left": left, "right": max(right, 20), "top": body_top - 4,
            "bottom": H - foot_top + 12, "header": head_top, "footer": max(H - foot_bottom - 6, 18),
            "title_left": Counter(round(g["x0"]) for g in title).most_common(1)[0][0] if title else left,
            "head_imgs": head_imgs, "head_lines": head_lines, "head_grid": head,
            "stamp": foot["rows"], "stamp_widths": foot["widths"], "stamp_aligns": foot["aligns"],
            "stamp_font": foot["font"] or "Times New Roman", "stamp_size": foot["size"], "stamp_box": foot["box"],
            "stamp_inner": foot["inner"], "page_tok": page_tok}


def _typical_page(pdf, body_top, H, sample=15):
    """Лист для колонтитулов — с самыми частыми колонтитулами (текст без цифр), а не второй по счёту: у ÜFK
    второй лист — оглавление, и в подвал всех листов попадало «Tartalom» («Содержание», 16.10). Заодно — номер
    листа: число колонтитула, которое растёт вместе с листом («13. oldal» на 19-м листе) → (лист, его текст)."""
    import pdfplumber
    n = len(pdf.pages)
    idx = sorted({round(1 + k * (n - 2) / max(sample - 1, 1)) for k in range(sample)} & set(range(1, n))) or [0]
    sig, words = {}, {}
    for i in idx:
        pg = pdf.pages[i]
        art = [c for c in pg.chars if c.get("mcid") is None]
        ftop = _foot_top(art, pg.rects + pg.lines, H)
        arts = [c for c in art if c["top"] < body_top or c["top"] >= ftop]
        sig[i] = re.sub(r"\d+", "#", "".join(c["text"] for c in arts))
        words[i] = [w["text"] for w in pdfplumber.utils.extract_words(arts, x_tolerance=2.5, y_tolerance=2)] if arts else []
    common = Counter(s for s in sig.values() if s.strip()).most_common(1)
    same = [i for i in idx if common and sig[i] == common[0][0]] or idx
    rep = same[len(same) // 2]
    tok = None
    if len(same) >= 3 and len({len(words[i]) for i in same}) == 1:
        for j, w in enumerate(words[rep]):
            vals = [re.fullmatch(r"(\d+)\.?", words[i][j]) for i in same]
            if all(vals) and len({int(v[1]) - i for v, i in zip(vals, same)}) == 1:
                tok = w
                break
    return pdf.pages[rep], tok


def _foot_top(arts, rects, H):
    """Верх подвала: нижняя группа строк-артефактов и линий без разрыва больше 15 pt. Порог «ниже 80 % листа»
    захватывал низ таблицы и пометку «Folytatás a következő oldalon» — их линии резали подвал на лишние ячейки
    (ÜFK I, 16.10)."""
    ys = sorted({round(float(c["top"])) for c in arts if float(c["top"]) > H * 0.7}
                | {round(float(r["top"])) for r in rects if float(r["top"]) > H * 0.7}, reverse=True)
    if not ys:
        return H
    top = ys[0]
    for y in ys[1:]:
        if top - y > 15:
            break
        top = y
    return top - 1


def _zone_grid(chars, rects):
    """Сетка колонтитула: столбцы — по вертикальным краям рамок; если внутренних линий нет — по разрывам между
    группами слов (> 18 pt) внутри рамки; строки — по словам; выравнивание столбца — по положению текста; шрифт
    и кегль — по знакам колонтитула. {"rows": [{"cells", "bold"}], "widths", "aligns", "font", "size", "box"}."""
    import pdfplumber
    out = {"rows": [], "widths": [], "aligns": [], "font": None, "size": None, "box": bool(rects), "inner": False}
    words = pdfplumber.utils.extract_words(chars, x_tolerance=2.5, y_tolerance=2) if chars else []
    if not words:
        return out
    xs = sorted({round(float(r["x0"])) for r in rects} | {round(float(r["x1"])) for r in rects})
    cols = []
    for x in xs:
        if not cols or x - cols[-1] > 8:
            cols.append(x)
    out["inner"] = len(cols) >= 3
    if len(cols) < 3:                                       # нет внутренних линий — разрывы между словами
        box = len(cols) == 2                                # края рамки — или края текста
        lo = cols[0] if box else round(min(float(w["x0"]) for w in words)) - 4
        hi = cols[-1] if box else round(max(float(w["x1"]) for w in words)) + 4
        spans = sorted((float(w["x0"]), float(w["x1"])) for w in words)
        cols, end = [lo], spans[0][1]
        for a, b in spans[1:]:
            if a - end > 18:
                cols.append(round((a + end) / 2))
            end = max(end, b)
        cols.append(hi)
    rows = {}
    for w in words:
        rows.setdefault(round(float(w["top"]) / 3), []).append(w)
    ncell = max(len(cols) - 1, 1)
    side = [[] for _ in range(ncell)]
    for _, ws in sorted(rows.items()):
        cells = [""] * ncell
        bold = any("Bold" in c["fontname"] for c in chars if abs(c["top"] - ws[0]["top"]) < 2)
        for w in sorted(ws, key=lambda w: w["x0"]):
            ci = max(0, min(ncell - 1, sum(1 for x in cols[1:-1] if float(w["x0"]) >= x - 2)))
            cells[ci] = (cells[ci] + " " + w["text"]).strip()
            side[ci].append((float(w["x0"]) - cols[ci], cols[ci + 1] - float(w["x1"])))
        out["rows"].append({"cells": cells, "bold": bold})
    for ci, gaps in enumerate(side):
        left = min((g[0] for g in gaps), default=0)
        right = min((g[1] for g in gaps), default=0)
        out["aligns"].append("right" if right < 10 < left - 10 else "center" if abs(left - right) < 6 and left > 10 else "left")
    out["widths"] = [cols[i + 1] - cols[i] for i in range(ncell)]
    f = Counter((_fmt(c)["font"], _fmt(c)["size"]) for c in chars if c["text"].strip()).most_common(1)
    if f:
        out["font"], out["size"] = f[0][0]
    return out
