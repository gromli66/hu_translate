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
PAGE_MARK = "\ue000"           # знаки номера листа в колонтитуле образцового листа (layout) — там ставится поле PAGE
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
    just = _justified_gaps([it["ch"] for it in items if "ch" in it])
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
            elif same_line and ch["x0"] - prev["x1"] > 1.8 * prev["size"] and t != " " and id(ch) not in just \
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


def _justified_gaps(chars):
    """Знаки, перед которыми широкий промежуток — растяжка выключки по ширине, а не табуляция: строка не последняя,
    доходит до правого края абзаца, и все промежутки между словами в ней равны (или он один). В узкой ячейке строка из
    двух слов («Агрегат ⇥ работа») давала табуляцию — счёт «три разрыва в абзаце» её не ловил (синтетика, 16.10);
    табуляция после «1.» в такой строке шире прочих промежутков и остаётся. → множество id знаков."""
    ink = [c for c in chars if c["text"].strip()]
    if not ink:
        return set()
    right = max(float(c["x1"]) for c in ink)
    lines, cur = [], [ink[0]]
    for c in ink[1:]:
        if abs(float(c["top"]) - float(cur[-1]["top"])) < 0.5 * float(cur[-1]["size"]):
            cur.append(c)
        else:
            lines.append(cur); cur = [c]
    lines.append(cur)
    out = set()
    for ln in lines[:-1]:
        if float(ln[-1]["x1"]) < right - 1.5:
            continue
        gaps = [(float(b["x0"]) - float(a["x1"]), b) for a, b in zip(ln, ln[1:])
                if float(b["x0"]) - float(a["x1"]) > 0.15 * float(a["size"])]
        if gaps and max(g for g, _ in gaps) - min(g for g, _ in gaps) < 1.0:
            out |= {id(b) for g, b in gaps if g > 1.8 * float(b["size"])}
    return out


def _tail_pairs(last, r):
    """Хвост строки last на новом листе (см. _cut_rows) → [(столбец в last, ячейка хвоста)] или []."""
    cell_t = lambda c: " ".join(text_of(x["runs"]) for x in c.get("blocks", []) if x["k"] == "p").strip()
    page = lambda row: next((c["geom"]["page"] for c in row if c.get("geom")), None)
    if not (last and r and len(r) <= len(last) and (page(r) or 0) > (page(last) or 0)) \
            or any((c.get("attrs") or {}).get("ColSpan") for c in r):
        return []
    pairs = []
    for c in r:
        g = c.get("geom")
        j = next((j for j, d in enumerate(last) if g and d.get("geom") and abs(d["geom"]["x0"] - g["x0"]) < 3), None)
        t = cell_t(c)
        # продолжение фразы — со строчной: «ячейка выше без точки» не годится — у шапок точек нет, и строка «1 | …»
        # нового листа уходила в шапку (синтетика, 16.10)
        if j is None or t and not t[0].islower():
            return []
        pairs.append((j, c))
    return pairs if any(cell_t(c) for _, c in pairs) else []


def _cut_rows(rows):
    """Строка, разрезанная листом: Word переносит её хвост на следующий лист отдельной строкой — только с теми
    ячейками, где текст продолжается (шапка таблицы, начатая внизу листа: «…уровень» | «необходимо контроль» на
    следующем листе, синтетика 16.10). Хвост — строка на новом листе, где ячеек меньше, чем в строке перед ней, каждая
    стоит под своей ячейкой (по левому краю текста), без объединения столбцов, и текст продолжает фразу (непустые
    ячейки начаты строчной), — дописывается в ячейки строки выше. Хвост — только первая строка нового листа."""
    out, glued = [], -1
    for r in rows:
        last = out[-1] if out else None
        pairs = _tail_pairs(last, r) if last and glued != len(out) - 1 else []
        if not pairs:
            out.append(r)
            continue
        glued = len(out) - 1
        last = out[-1] = [dict(c) for c in last]
        for j, c in pairs:
            head, tail = list(last[j]["blocks"]), list(c["blocks"])
            if head and tail and head[-1]["k"] == "p" and tail[0]["k"] == "p":     # абзац, разрезанный листом, — один
                runs = list(head[-1]["runs"]) + [dict(x) for x in tail[0]["runs"]]
                k = len(head[-1]["runs"])
                if k < len(runs) and "t" in runs[k] and runs[k - 1].get("t", " ")[-1:] not in (" ", "-"):
                    runs[k]["t"] = " " + runs[k]["t"]
                head[-1] = dict(head[-1], runs=runs)
                tail = tail[1:]
            last[j]["blocks"] = head + tail
    return out


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
    (UTS стр. 75–78: разрыв с пустой строкой и без шапки, 16.10). Продолжение, начатое повтором шапки (строки
    «Повторять как заголовок» Word печатает на каждом листе), — тоже она: повтор прочь (синтетика, 16.10)."""
    cell_t = lambda c: " ".join(text_of(x["runs"]) for x in c.get("blocks", []) if x["k"] == "p").strip()
    row_t = lambda r: [cell_t(c) for c in r]

    def head_n(t):                          # число строк-шапок в начале таблицы
        n = 0
        while n < len(t["rows"]) and all(c["th"] for c in t["rows"][n]):
            n += 1
        return n
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
        rep = 0
        if prev is not None and prev["k"] == "table" and b["k"] == "table" and prev["rows"] and b["rows"]:
            n = head_n(b)
            if n and n < len(b["rows"]) and n <= head_n(prev) \
                    and [row_t(r) for r in b["rows"][:n]] == [row_t(r) for r in prev["rows"][:n]]                     and any(any(row_t(r)) for r in b["rows"][:n]):
                rep = n                     # повтор шапки на новом листе
        if prev is not None and prev["k"] == "table" and b["k"] == "table" and prev["rows"] and b["rows"] \
                and (rep or not all(c["th"] for c in b["rows"][0]) or _tail_pairs(prev["rows"][-1], b["rows"][0])) \
                and (max(map(len, b["rows"])) == max(map(len, prev["rows"])) or _tail_pairs(prev["rows"][-1], b["rows"][0])) \
                and (b.get("geom") or {}).get("page", 0) > (prev.get("geom") or {}).get("page", 0):
            out[-1] = dict(prev, rows=_cut_rows(prev["rows"] + b["rows"][rep:]))
            continue
        if b["k"] == "table":
            b = dict(b, rows=_cut_rows(b["rows"]))
        out.append(b)
    return out


def _pos(b):
    """(лист, верх) первого размещённого места блока (у списка — первого пункта) или None."""
    g = b.get("geom") or {}
    if g.get("page") and g.get("top") is not None:
        return g["page"], g["top"]
    for it in b.get("items", []):
        for x in it:
            q = _pos(x)
            if q:
                return q
    return None


def _flat_paras(b):
    """Блок без списков: список-обёртка (Word кладёт подпись в L/LI/LBody) — его абзацы по порядку."""
    if b["k"] != "list":
        return [b]
    return [y for it in b["items"] for x in it for y in _flat_paras(x)]


def evict(bs):
    """Чужое в ячейке: Word вкладывает в список последней ячейки таблицы текст со следующего листа (подпись «Рисунок 1
    – …» листа 13 — в ячейке таблицы листов 11–12, JPDS 16.10), и в DOCX он пропадал. Блок ячейки, стоящий на листе
    ниже того, что идёт после таблицы, — прочь из таблицы, в поток на своё место по листу. Хвост разрезанной строки
    стоит вверху следующего листа, выше всего после таблицы, и остаётся."""
    out = list(bs)
    i = 0
    while i < len(out):
        t = out[i]
        if t["k"] != "table":
            i += 1
            continue
        after = [q for q in (_pos(x) for x in out[i + 1:]) if q]
        lim = min(after) if after else None
        moved = []

        def clean(blocks, cpage):
            keep = []
            for x in blocks:
                q = _pos(x)
                if lim and q and q > lim and q[0] > cpage:      # и на более позднем листе, чем сама ячейка
                    moved.extend(_flat_paras(x))
                    continue
                if x["k"] == "list":
                    items = [clean(it, cpage) for it in x["items"]]
                    items = [it for it in items if it]
                    if not items:
                        continue
                    x = dict(x, items=items)
                keep.append(x)
            return keep
        rows = [[dict(c, blocks=clean(c.get("blocks", []), (c.get("geom") or {}).get("page") or 0)) for c in r]
                for r in t["rows"]]
        if not moved:
            i += 1
            continue
        out[i] = dict(t, rows=rows)
        for m in moved:                                 # на место по листу среди блоков после таблицы
            q = _pos(m)
            j = i + 1
            while j < len(out) and (_pos(out[j]) or (0, 0)) <= q:
                j += 1
            out.insert(j, m)
        i += 1
    return out


def join_pages(bs):
    """Абзац, разрезанный листом: Word размечает его половины двумя P — в DOCX выходил разрыв абзаца посреди фразы
    (с отступом первой строки у второй половины, синтетика 16.10). Соседние абзацы на соседних листах — один, если
    первый не кончен знаком конца фразы, а второй начат строчной или без отступа первой строки, который есть у
    первого. Внутри пунктов списков и ячеек — так же."""
    out = []
    for b in bs:
        if b["k"] == "list":
            b = dict(b, items=[join_pages(it) for it in b["items"]])
        elif b["k"] == "table":
            b = dict(b, rows=[[dict(c, blocks=join_pages(c.get("blocks", []))) for c in r] for r in b["rows"]])
        a = out[-1] if out else None
        ga, gb = (a or {}).get("geom") or {}, b.get("geom") or {}
        if a and a["k"] == b["k"] == "p" and ga.get("page") and gb.get("page") and gb["page"] == ga["page"] + 1 \
                and a["runs"] and b["runs"] and "t" in a["runs"][-1] and "t" in b["runs"][0]:
            ta, tb = text_of(a["runs"]).rstrip(), text_of(b["runs"]).lstrip()
            indent = ga.get("first_x0", 0) - ga.get("x0", 0) > 5 and abs(gb.get("first_x0", 0) - ga.get("x0", 0)) < 3
            if ta and tb and not re.search(r"[.!?:;]$", ta) and (tb[0].islower() or indent):
                runs = list(a["runs"]) + [dict(r) for r in b["runs"]]
                k = len(a["runs"])
                if not runs[k - 1]["t"].endswith((" ", "-", "\u00ad")):
                    runs[k]["t"] = " " + runs[k]["t"].lstrip()
                out[-1] = dict(a, runs=runs)
                continue
        out.append(b)
    return out


def reflow(bs):
    """Блоки не на своём месте — к своему листу: надпись Word (плавающая) лежит в дереве там, куда привязана, а нарисована
    далеко (руководство ТЭ357: «Текст строки» с листа 40 — между листами оглавления 2 и 3, оглавление рвалось на два
    поля, надписи печатались в нём). Блок, чей лист дальше чем на 2 от листов соседей с обеих сторон, — перед первым
    блоком своего листа или дальше."""
    page = lambda b: (b.get("geom") or {}).get("page")
    runs = []                                   # подряд идущие блоки одного листа: [(лист, [номера блоков])]
    for i, b in enumerate(bs):
        if page(b) is None:
            if runs:
                runs[-1][1].append(i)           # пустой абзац без места — с соседями
            continue
        if runs and runs[-1][0] == page(b):
            runs[-1][1].append(i)
        else:
            runs.append((page(b), [i]))
    stray = set()
    for k, (p, ids) in enumerate(runs):
        if 0 < k < len(runs) - 1:
            prev, nxt = runs[k - 1][0], runs[k + 1][0]
            if abs(p - prev) > 2 and abs(p - nxt) > 2 and prev <= nxt:
                stray |= set(ids)
    if not stray:
        return bs
    moved = [bs[i] for i in sorted(stray)]
    rest = [b for i, b in enumerate(bs) if i not in stray]
    at, last = None, None
    for b in moved:                             # группой, по порядку: место ищется только для нового листа
        if at is None or (page(b) is not None and page(b) != last):
            at = next((j for j, x in enumerate(rest) if (page(x) or 0) >= (page(b) or last or 0)), len(rest))
            last = page(b) or last
        rest.insert(at, b)
        at += 1
    return rest


def join_tocs(bs):
    """Оглавление, размеченное по листам (ÜFK: элемент TOC на каждом из листов 2–5), — одно: было четыре поля TOC
    подряд, 556 пунктов вместо 206 (16.10). Уровни — больший, пункты — все. Пустые абзацы между кусками не мешают."""
    out = []
    empty = lambda x: x["k"] == "p" and not text_of(x.get("runs") or []).strip() and \
        not any("img" in r for r in x.get("runs") or [])
    for b in bs:
        prev = next((x for x in reversed(out) if not empty(x)), None)
        if b["k"] == "toc" and prev is not None and prev["k"] == "toc" and \
                bool(b["info"].get("figures")) == bool(prev["info"].get("figures")):
            a, n = prev["info"], b["info"]
            a["levels"] = max(a["levels"], n["levels"])
            a["keys"] = a.get("keys", []) + n.get("keys", [])
            a["extra"] = a.get("extra", []) + [e for e in n.get("extra", []) if e not in a.get("extra", [])]
            if len(n.get("bold", [])) > len(a.get("bold", [])):
                a["bold"], a["indent"] = n["bold"], n["indent"]
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
        roman = lambda case: all(set(x) <= set("ivx" if case else "IVX") for x in labs) and             [label_value(x, "LowerRoman") for x in labs] == list(range(1, len(labs) + 1))
        kind = ("Decimal" if all(x.isdigit() for x in labs) else
                "LowerRoman" if roman(True) and len(labs) > 1 else          # «i.», «ii.» (ÜFK: «• i.» — двойной маркер)
                "UpperRoman" if roman(False) and len(labs) > 1 else
                "LowerAlpha" if all(len(x) == 1 and x.islower() for x in labs) else
                "UpperAlpha" if all(len(x) == 1 and x.isupper() for x in labs) else None)
        vals = [label_value(x, kind) for x in labs] if kind else []
        if kind and all(v2 == v1 + 1 for v1, v2 in zip(vals, vals[1:])):
            b["kind"], b["suffix"] = kind, ms[0].group(2)
PLABEL = re.compile(r"^\s*(\d{1,2}|[a-z]|[A-Z])(\.\)|\))\s+\S")   # «a.)», «1)»
DLABEL = re.compile(r"^\s*(\d{1,2}|[a-z])(\.)\s+\S")                  # «1.», «a.» — только с отступом от абзаца выше


def _label_kind(lbl):
    return "Decimal" if lbl.isdigit() else "LowerAlpha" if lbl.islower() else "UpperAlpha"


def para_lists(bs):
    """Абзацы подряд с метками по порядку прямо в тексте («a.)», «b.)»; «1)», «2)») — список того же вида: дерево Word
    дало их простыми абзацами, и в переводе пункты шли без отступа, метка — текстом (ÜFK I: 284 абзаца, 16.10).
    Абзац между пунктами правее метки — продолжение пункта. Метки снимает relist (номера даёт список Word)."""
    out, i = [], 0
    while i < len(bs):
        b = bs[i]
        if b["k"] == "table":
            b = dict(b, rows=[[dict(c, blocks=para_lists(c["blocks"])) for c in r] for r in b["rows"]])
        elif b["k"] == "list":
            b = dict(b, items=[para_lists(it) for it in b["items"]])
        m = PLABEL.match(text_of(b["runs"])) if b["k"] == "p" and b.get("runs") else None
        if not m and b["k"] == "p" and b.get("runs") and out and (out[-1].get("geom") or {}).get("x0") is not None                 and ((b.get("geom") or {}).get("x0") or 0) > out[-1]["geom"]["x0"] + 10:
            m = DLABEL.match(text_of(b["runs"]))            # «1. Az adott fejezetet …» под абзацем (ÜFK II)
        if not m:
            out.append(b)
            i += 1
            continue
        kind, suffix = _label_kind(m.group(1)), m.group(2)
        v, x0 = label_value(m.group(1), kind), (b.get("geom") or {}).get("x0")
        items, j = [[b]], i + 1
        while j < len(bs) and bs[j]["k"] == "p" and bs[j].get("runs"):
            c = bs[j]
            mm = (PLABEL if suffix != "." else DLABEL).match(text_of(c["runs"]))
            if mm and mm.group(2) == suffix and _label_kind(mm.group(1)) == kind and \
                    label_value(mm.group(1), kind) == v + 1:
                items.append([c])
                v += 1
            elif not mm and x0 is not None and ((c.get("geom") or {}).get("x0") or 0) > x0 + 8:
                items[-1].append(c)                         # продолжение пункта — с отступом текста пункта
            else:
                break
            j += 1
        prev = out[-1] if out else None
        px0 = ((prev or {}).get("geom") or {}).get("x0")
        if prev and prev["k"] == "list" and prev.get("kind") == kind and prev.get("suffix") == suffix and                 label_value(m.group(1), kind) == len(prev["items"]) + 1:
            prev["items"] += items                          # «f.)» после «a.)–e.)», разорванных другим элементом
            i = j
        elif len(items) >= 2 or (suffix != "." and label_value(m.group(1), kind) == 1 and x0 is not None
                                 and px0 is not None and x0 > px0 + 10):
            # и один пункт «a.)» с отступом под «Korlátozó feltétel:» (ÜFK: у условия бывает один пункт)
            out.append({"k": "list", "kind": kind, "suffix": suffix, "items": items})
            i = j
        else:
            out.append(b)
            i += 1
    return out


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
            # нумерованный и следом маркированный списки Word кладёт одним L (вид — по первому): у маркированных
            # пунктов «•» прямо в тексте, без метки — вид пункта по нему, смена вида открывает новый список
            # (синтетика, 16.10); сам маркер из текста снимает сборка (docx_build.LEAD)
            p0 = next((x for x in it["body"] if x["k"] == "p"), None)
            m = re.match(r"\s*([•▪◦·–-])\s", text_of(p0["runs"])) if p0 and p0["runs"] else None
            if it["kind"] in NUMBERED and m:
                it["kind"] = "Dash" if m.group(1) in "–-" else "Disc"
            # и наоборот: маркированный, следом нумерованный — у нумерованных номер («8.») в тексте
            ml = LABEL.match(text_of(p0["runs"])) if p0 and p0["runs"] else None
            if it["kind"] not in NUMBERED and ml and (ml.group(1).isdigit() or ml.group(2) in (")", ".)")):
                it["kind"] = _label_kind(ml.group(1))
                if ml.group(2) == ".)":
                    it["suffix"] = ".)"
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
        # уровень пункта — глубина его номера («4.2.1.», «DF 1.2.1.» — третий), если пункты в основном с номерами: у
        # руководства ТЭ357 номера всех уровней стоят в одном столбце, и по отступу выходил один уровень — поле TOC
        # собирало только главы (16.10); пункты без номеров — по отступу
        depth = lambda e: (lambda m: min(len(m.group(1).split(".")), 9) if m else None)(
            re.match(r"^\s*(?:[A-ZÁÉÍÓÖŐÚÜŰ]{1,4}\s+)?(\d+(?:\.\d+)*)\.?\s", e["t"]))
        ds = [depth(e) for e in body]
        lv = []
        if body and sum(d is not None for d in ds) >= 0.7 * len(body):
            for k in range(1, max(d for d in ds if d) + 1):
                at = sorted(e["x0"] for e, d in zip(body, ds) if d == k)
                lv.append(at[len(at) // 2] if at else (lv[-1] if lv else min(e["x0"] for e in body)))
        else:
            for x in sorted(e["x0"] for e in body):
                if not lv or x - lv[-1] > 4:
                    lv.append(x)
            ds = [None] * len(body)
        by_x = lambda x: min(range(len(lv)), key=lambda i: abs(lv[i] - x)) + 1
        lev = {id(e): d or by_x(e["x0"]) for e, d in zip(body, ds)}
        bold = []
        for i in range(len(lv)):
            at = [e for e in body if lev[id(e)] == i + 1]
            bold.append(sum(e["f"]["b"] for e in at) * 2 > len(at) if at else False)
        figs = sum(bool(re.match(r"\d+\.\s*ábra", e["t"], re.I)) for e in body) * 2 > len(body)
        extra = [] if figs else [(re.sub(r"\s*(\.{3,}\s*)?\d+\s*$", "", e["t"]).strip(), lev[id(e)])
                                 for e in body if not re.match(r"\d+(\.\d+)*\s", e["t"])]
        width = float(pdf.pages[0].width)
        first_page = ents[0]["page"] if ents else None
        steps = sorted(b["top"] - a["top"] for a, b in zip(body, body[1:]) if a["page"] == b["page"] and b["top"] > a["top"])
        return {"title": title and dict(title["f"], t=title["t"], center=abs(title["cx"] - width / 2) < 30),
                "levels": len(lv), "bold": bold, "figures": figs, "extra": extra,
                "indent": [round(x - lv[0], 1) for x in lv],                      # отступ уровня от 1-го, pt
                "pitch": round(steps[len(steps) // 2], 1) if steps else None,     # шаг строк оглавления, pt
                "page": first_page,
                # что в оглавлении было: заголовки не из него («DF 1.3.1.» у ÜFK) — без уровня структуры, иначе
                # поле TOC собирало 556 пунктов вместо 206 (16.10)
                "keys": [re.sub(r"\s+", "", re.sub(r"\s*(\.{3,}\s*)?\d+\s*$", "", e["t"])) for e in body]}

    def own(el):
        """Абзац из собственных MCID контейнера (текст прямо в LBody/TD, без вложенного P)."""
        its = []
        for m in el.mcids or []:
            its += content.get((el.page_number, m), [])
        return [{"k": "p", "runs": _runs(its), "geom": geom(el)}] if its else []

    def container(el, depth):
        """Содержимое LBody/TD/TH: свой текст, затем вложенные блоки."""
        return own(el) + sum((block(x, depth + 1) for x in el.children), [])

    thin_of = {}

    def thin(pn):
        """Линии листа (не заливки): (горизонтальные, вертикальные) как (поз., от, до)."""
        if pn not in thin_of:
            pg = pdf.pages[pn - 1]
            hs, vs = [], []
            for o in pg.rects + pg.lines:
                x0, x1, t0, t1 = float(o["x0"]), float(o["x1"]), float(o["top"]), float(o["bottom"])
                if t1 - t0 < 3 and x1 - x0 > 5:
                    hs.append(((t0 + t1) / 2, x0, x1))
                elif x1 - x0 < 3 and t1 - t0 > 5:
                    vs.append(((x0 + x1) / 2, t0, t1))
            thin_of[pn] = (hs, vs)
        return thin_of[pn]

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
            if g:
                cell_edges(rows, thin)
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
    blocks = evict(join_tocs(join_tables(relist(para_lists(join_pages(reflow(blocks)))))))
    # принудительный разрыв листа: заголовок первым на листе, а прошлый лист кончился заметно (больше дюйма) выше
    # самого низкого содержимого документа. Разрыв стилем — лишь когда с нового листа все заголовки уровня; у ÜFK
    # подразделы 4.4.1.x — каждый с нового листа, кроме идущего сразу под разделом, — и разрывы пропадали все:
    # 157 листов собирались в 128 (16.10)
    tops, bots = {}, {}
    for pg in pdf.pages:
        ys = [(float(c["top"]), float(c["bottom"])) for c in pg.chars if c.get("mcid") is not None and c["text"].strip()]
        ys += [(float(i["top"]), float(i["bottom"])) for i in pg.images if i.get("mcid") is not None]
        if ys:
            tops[pg.page_number], bots[pg.page_number] = min(y[0] for y in ys), max(y[1] for y in ys)
    low = max(bots.values(), default=0)
    for b in blocks:
        g = b.get("geom") or {}
        pn = g.get("page")
        if b["k"] == "h" and pn and pn - 1 in bots and pn in tops and abs(g["top"] - tops[pn]) < 2 \
                and low - bots[pn - 1] > 72:
            b["break"] = True
    info = {"width": float(page.width), "height": float(page.height), "pages": len(pdf.pages),
            "artifacts": artifacts}
    return blocks, info


def cell_edges(rows, thin):
    """Линии вокруг каждой ячейки в оригинале: c["edges"] = {"top", "bottom", "left", "right": bool}. У ÜFK линии —
    только между строками и между некоторыми столбцами (метка «E.1.» и текст — без линии), полная сетка Word
    («Сетка таблицы») рисовала лишние (16.10). thin(лист) → (горизонтальные, вертикальные) линии (поз., от, до).
    Ячейка без текста — по краям своего столбца в полных строках."""
    ncol = max((len(r) for r in rows), default=0)
    cols = [[c["geom"] for r in rows if len(r) == ncol for c in r[ci:ci + 1] if c.get("geom")] for ci in range(ncol)]
    span = [(min(g["x0"] for g in cs), max(g["x1"] for g in cs)) if cs else None for cs in cols]
    for ri, row in enumerate(rows):
        gs = [c.get("geom") for c in row]
        pn = next((g["page"] for g in gs if g), None)
        if pn is None:
            continue
        hs, vs = thin(pn)
        top = min(g["top"] for g in gs if g and g["page"] == pn)
        bot = max(g["bottom"] for g in gs if g and g["page"] == pn)
        prev = [c["geom"] for c in rows[ri - 1] if c.get("geom") and c["geom"]["page"] == pn] if ri else []
        nxt = [c["geom"] for c in rows[ri + 1] if c.get("geom") and c["geom"]["page"] == pn] if ri + 1 < len(rows) else []
        y_lo = max(g["bottom"] for g in prev) - 1 if prev else top - 16
        y_hi = min(g["top"] for g in nxt) + 1 if nxt else bot + 16
        near = [h for h in hs if y_lo - 2 <= h[0] <= y_hi + 2]
        tx0 = min((h[1] for h in near), default=None)
        tx1 = max((h[2] for h in near), default=None)
        xs = []
        for ci, g in enumerate(gs):
            if g and g["page"] == pn:
                xs.append((g["x0"], g["x1"]))
            elif len(row) == ncol and span[ci]:
                xs.append(span[ci])
            else:
                xs.append(None)
        for ci, c in enumerate(row):
            if xs[ci] is None:
                continue
            x0, x1 = xs[ci]
            w = max(x1 - x0, 1.0)
            cover = lambda h: min(h[2], x1) - max(h[1], x0) >= 0.5 * w
            l_lo = xs[ci - 1][1] - 1 if ci and xs[ci - 1] else (tx0 - 2 if tx0 is not None else x0 - 30)
            r_hi = xs[ci + 1][0] + 1 if ci + 1 < len(row) and xs[ci + 1] else (tx1 + 2 if tx1 is not None else x1 + 30)
            h = max(bot - top, 1.0)
            vcover = lambda v: min(v[2], bot) - max(v[1], top) >= 0.5 * h
            c["edges"] = {"top": any(y_lo <= y[0] <= top + 1 and cover(y) for y in hs),
                          "bottom": any(bot - 1 <= y[0] <= y_hi and cover(y) for y in hs),
                          "left": any(min(l_lo, x0 - 1) <= v[0] <= x0 + 1 and vcover(v) for v in vs),
                          "right": any(x1 - 1 <= v[0] <= max(r_hi, x1 + 1) and vcover(v) for v in vs)}
        for a, b in zip(row, row[1:]):                      # общая граница соседей — одна
            if "edges" in a and "edges" in b:
                a["edges"]["right"] = b["edges"]["left"] = a["edges"]["right"] or b["edges"]["left"]


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
    # верх тела — по первому блоку листов, 20-й процентиль, а не минимум: надпись у края одного листа (МТ-ЦАСУТП-14,
    # лист 19: 3 pt) давала верхнее поле −0.9 pt, и шапка (знаки выше верха тела) терялась целиком (16.10)
    firsts = {}
    for g in body:
        firsts[g["page"]] = min(firsts.get(g["page"], g["top"]), g["top"])
    tops_ = sorted(firsts.values())
    body_top = tops_[len(tops_) // 5] if tops_ else 110
    page, page_tok, page_off, tok_boxes = _typical_page(pdf, body_top, H)
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
    frects = [r for r in page.rects + page.lines if float(r["top"]) >= ftop]
    # номер листа — помечен в знаках: полем станет он, а не такое же число рядом («Блок 3» на листе 3, синтетика)
    mark = lambda chars: [dict(c, text=PAGE_MARK) if c["text"].isdigit() and any(
        abs(float(c["top"]) - bx["top"]) < 2 and bx["x0"] - 0.5 <= float(c["x0"]) and float(c["x1"]) <= bx["x1"] + 0.5
        for bx in tok_boxes) else c for c in chars]
    foot = _rule_grid(mark(foot_chars), frects) or _zone_grid(mark(foot_chars), frects)
    # шапка-таблица по линиям (с логотипом в своей ячейке); без внутренних линий и логотипов (ÜFK: строка-таблица
    # «Üzemeltetési Feltételek … | 1.2. változat») — сеткой по разрывам между словами
    hrects = _head_rects(page, body_top)
    head = _rule_grid(mark(head_chars), hrects, head_imgs) or \
        (_zone_grid(mark(head_chars), hrects) if head_chars and not head_imgs else None)
    if head:
        head["para"] = _as_para(head, left, W - right)
    # подвала нет — низ тела по последнему блоку листов (80-й процентиль): поле по умолчанию (100 pt) отнимало у
    # каждого листа МТ-ЦАСУТП-14 три строки (60 листов → 69, 16.10)
    lasts = {}
    for g in body:
        lasts[g["page"]] = max(lasts.get(g["page"], 0), g.get("bottom") or g["top"])
    bots_ = sorted(lasts.values())
    foot_top = min((float(c["top"]) for c in foot_chars),
                   default=bots_[len(bots_) * 4 // 5] + 14 if bots_ else H - 100)
    foot_bottom = max((float(c["bottom"]) for c in foot_chars), default=H - 40)
    # шапки нет — пустой абзац шапки (её расстояние + строка) должен уместиться в верхнее поле: при поле 24 pt и
    # расстоянии 36 pt тело уезжало на каждом листе вниз на 2–3 строки (ТЭ357: 66 листов → 85, 16.10)
    head_top = min([i["top"] for i in head_imgs] + [float(c["top"]) for c in head_chars] or
                   [max(6.0, min(36.0, body_top - 4 - 18))])
    sig = lambda chars: re.sub(r"\s+", " ", re.sub(r"\d+\.?", "", "".join(c["text"] for c in chars))).strip()
    body_sig = sig(foot_chars)
    out = {"W": W, "H": H, "left": left, "right": max(right, 20), "top": body_top - 4,
           "bottom": H - foot_top + 12, "header": head_top, "footer": max(H - foot_bottom - 6, 18),
           "title_left": Counter(round(g["x0"]) for g in title).most_common(1)[0][0] if title else left,
           "head_imgs": head_imgs, "head_lines": head_lines, "head_grid": head,
           "stamp": foot["rows"], "stamp_widths": foot["widths"], "stamp_aligns": foot["aligns"],
           "stamp_font": foot["font"] or "Times New Roman", "stamp_size": foot["size"], "stamp_box": foot["box"],
           "stamp_inner": foot["inner"], "stamp_merge": foot.get("merge"), "page_tok": page_tok,
           "page_marked": any(c["text"] == PAGE_MARK for c in mark(foot_chars + head_chars)),
           "stamp_para": _as_para(foot, left, W - right)}
    # первый лист — титул, только если его колонтитулы не такие, как у тела: документ, начатый прямо с текста под
    # той же шапкой и штампом, иначе получал «титул» без колонтитулов (синтетика, 16.10)
    p1 = pdf.pages[0]
    a1 = [c for c in p1.chars if c.get("mcid") is None]
    f1 = [c for c in a1 if c["top"] >= _foot_top(a1, p1.rects + p1.lines, H)]
    out["first_is_body"] = bool(body_sig) and sig(f1) == body_sig
    # вводная часть: листы после титула со своим подвалом (ÜFK: оглавление и лист изменений — «Tartalomjegyzék» без
    # номера) — своим разделом; нумерация тела — с числа на первом его листе («1. oldal» на 7-м листе, 16.10)
    front = []
    for pg in pdf.pages[1:]:
        art = [c for c in pg.chars if c.get("mcid") is None]
        ft = _foot_top(art, pg.rects + pg.lines, H)
        fc = [c for c in art if c["top"] >= ft]
        if not body_sig or sig(fc) == body_sig:
            break
        front.append((pg, fc, ft, [c for c in art if c["top"] < body_top]))
    # вводная часть короткая: подвал, меняющийся по главам (название главы), иначе съел бы листы тела
    if front and len(front) <= max(3, 0.15 * len(pdf.pages)):
        pg, fc, ft, hc = front[0]
        fg = _zone_grid(fc, [r for r in pg.rects + pg.lines if float(r["top"]) >= ft])
        out["front"] = {"to": front[-1][0].page_number, "stamp": fg["rows"], "stamp_widths": fg["widths"],
                        "stamp_aligns": fg["aligns"], "stamp_box": fg["box"], "stamp_inner": fg["inner"],
                        "stamp_merge": None, "stamp_para": _as_para(fg, left, W - right)}
        if sig(hc) != sig(head_chars):
            hg = _zone_grid(hc, _head_rects(pg, body_top))
            hg["para"] = _as_para(hg, left, W - right)
            out["front"]["head_grid"] = hg
    if page_tok is not None:
        first = (out["front"]["to"] + 1) if "front" in out else 1 if out["first_is_body"] else 2   # 1-й лист тела
        out["num_start"] = {"page": first, "value": first - 1 + page_off}
    return out


def _head_rects(page, body_top):
    """Линии зоны шапки: выше тела и выше первого содержимого листа с запасом — рамка таблицы, которой начат лист,
    лежит выше её первого знака и попадала в шапку лишними строками и столбцами (синтетика, 16.10)."""
    ctop = min((float(c["top"]) for c in page.chars if c.get("mcid") is not None), default=body_top)
    lim = min(body_top, ctop - 6)
    return [r for r in page.rects + page.lines if float(r["bottom"]) < lim]


def _as_para(g, x_left, x_right):
    """Колонтитул без рамки в одну строку — абзацем, а не таблицей: части — по листу (по центру, у правого поля,
    слева) табуляциями. {"parts": [(текст, вид, позиция от левого поля)], "align"}; None — таблицей."""
    if not g or g.get("box") or g.get("rules") or len(g.get("rows") or []) != 1:
        return None
    parts = [(t, sp) for t, sp in zip(g["rows"][0]["cells"], g.get("spans") or []) if isinstance(t, str) and t and sp]
    if not parts:
        return None
    mid = (x_left + x_right) / 2
    kind = lambda a, b: "center" if abs((a + b) / 2 - mid) < 10 else "right" if abs(b - x_right) < 10 else "left"
    out = []
    for t, (a, b) in parts:
        k = kind(a, b)
        out.append((t, k, round(b - x_left if k == "right" else (a + b) / 2 - x_left if k == "center" else a - x_left, 1)))
    return {"parts": out, "bold": g["rows"][0].get("bold", False)}


def _typical_page(pdf, body_top, H, sample=15):
    """Лист для колонтитулов — с самыми частыми колонтитулами (текст без цифр), а не второй по счёту: у ÜFK
    второй лист — оглавление, и в подвал всех листов попадало «Tartalom» («Содержание», 16.10). Заодно — номер
    листа: число колонтитула, которое растёт вместе с листом («13. oldal» на 19-м листе) → (лист, его текст,
    смещение: номер = индекс листа от 0 + смещение, слово номера на этом листе)."""
    import pdfplumber
    n = len(pdf.pages)
    # выборка — по всему документу, и 1-й лист: титул со своими колонтитулами в общую группу не попадёт, а короткий
    # документ без титула (два листа) иначе не давал номера листа — в штампе застывала цифра (синтетика, 16.10)
    idx = sorted({round(k * (n - 1) / max(sample - 1, 1)) for k in range(sample)})
    sig, words = {}, {}
    for i in idx:
        pg = pdf.pages[i]
        art = [c for c in pg.chars if c.get("mcid") is None]
        ftop = _foot_top(art, pg.rects + pg.lines, H)
        arts = [c for c in art if c["top"] < body_top or c["top"] >= ftop]
        sig[i] = re.sub(r"\d+", "#", "".join(c["text"] for c in arts))
        words[i] = pdfplumber.utils.extract_words(arts, x_tolerance=2.5, y_tolerance=2) if arts else []
    common = Counter(s for s in sig.values() if s.strip()).most_common(1)
    same = [i for i in idx if common and sig[i] == common[0][0]] or idx
    rep = same[len(same) // 2]
    # номер листа может стоять в нескольких местах (в шапке и в штампе): помечены все — первый задаёт смещение
    tok, off, boxes = None, 0, []
    if len(same) >= 2 and len({len(words[i]) for i in same}) == 1:
        for j, w in enumerate(words[rep]):
            vals = [re.fullmatch(r"(\d+)\.?", words[i][j]["text"]) for i in same]
            if all(vals) and len({int(v[1]) - i for v, i in zip(vals, same)}) == 1 \
                    and (tok is None or int(vals[0][1]) - same[0] == off and w["text"] == tok):
                if tok is None:
                    tok, off = w["text"], int(vals[0][1]) - same[0]      # номер листа = индекс (от 0) + off
                boxes.append({k: float(w[k]) for k in ("x0", "x1", "top")})
    return pdf.pages[rep], tok, off, boxes


def _foot_top(arts, rects, H):
    """Верх подвала: нижняя группа строк-артефактов без разрыва больше 15 pt — или связанных вертикалью таблицы
    (строка штампа-таблицы выше 15 pt: подвал обрезался до пустой строки под таблицей, синтетика 16.10), плюс линия
    рамки над ней. Порог «ниже 80 % листа» захватывал низ таблицы и пометку «Folytatás a következő oldalon» — их
    линии резали подвал на лишние ячейки (ÜFK I, 16.10)."""
    rows = sorted({round(float(c["top"])) for c in arts if float(c["top"]) > H * 0.7}, reverse=True)
    rr = [r for r in rects if float(r["top"]) > H * 0.7]
    if not rows:                                # подвал из одних линий (рамка без текста)
        ys = sorted({round(float(r["top"])) for r in rr}, reverse=True)
        if not ys:
            return H
        top = ys[0]
        for y in ys[1:]:
            if top - y > 15:
                break
            top = y
        return top - 1
    vs = [(float(r["top"]), float(r["bottom"])) for r in rr
          if float(r["x1"]) - float(r["x0"]) < 3 and float(r["bottom"]) - float(r["top"]) > 5]
    top = rows[0]
    for y in rows[1:]:
        linked = any(a <= y + 12 and b2 >= top - 6 for a, b2 in vs)
        if top - y > 15 and not linked:
            break
        top = y
    # линия рамки — ближайшая над текстом (и вплотную к ней — двойная/толстая): самая верхняя в 12 pt брала низ
    # таблицы тела, кончившейся над штампом, — лишняя пустая строка штампа (синтетика, 16.10)
    border = sorted({float(r["top"]) for r in rr if float(r["bottom"]) - float(r["top"]) < 3
                     and top - 12 <= float(r["top"]) <= top}, reverse=True)
    for i in range(1, len(border)):
        if border[i - 1] - border[i] > 2:           # низ таблицы тела в 2.5 pt над штампом — не его линия
            border = border[:i]
            break
    return round(border[-1] if border else top) - 1


def _rule_grid(chars, rects, imgs=()):
    """Колонтитул-таблица по линиям оригинала: столбцы — по вертикалям, строки — по горизонталям (не по строкам
    текста); ячейка без линии под ней — объединена с нижней. Шапка инструкций Печорской ГРЭС: логотип | название в
    три строки и номер листа над ним | код и «Для внутреннего использования» — строками текста выходило пять строк
    таблицы с линиями, где их нет, а с логотипом сетка не строилась вовсе и номер листа застывал числом (16.10).
    Внутренних вертикалей нет (ÜFK) — None, сетку строит _zone_grid. Ячейка — {"lines": [(текст, полужирный)],
    "align"}; out["merge"][r][c] — продолжение ячейки выше; out["imgs"] — {(строка, столбец): [картинка]}."""
    import pdfplumber
    thin = []
    for o in rects:
        a, b, t, u = float(o["x0"]), float(o["x1"]), float(o["top"]), float(o["bottom"])
        if (b - a >= 3 and u - t >= 3) and o.get("stroke") and o.get("object_type") == "rect":
            # обведённая ячейка (штамп UTS) — четыре стороны
            thin += [(a, a, t, u), (b, b, t, u), (a, b, t, t), (a, b, u, u)]
        else:
            thin.append((a, b, t, u))
    vs = [((a + b) / 2, t, u) for a, b, t, u in thin if b - a < 3 and u - t > 5]
    hs = [((t + u) / 2, a, b) for a, b, t, u in thin if u - t < 3 and b - a > 20]
    clu = lambda vals: [v for i, v in enumerate(sorted(vals)) if i == 0 or v - sorted(vals)[i - 1] > 2.5]
    xs, ys = clu([v[0] for v in vs]), clu([h[0] for h in hs])
    words = pdfplumber.utils.extract_words(chars, x_tolerance=2.5, y_tolerance=2) if chars else []
    if len(xs) < 3 or len(ys) < 2 or not (words or imgs):
        return None
    ncol, nrow = len(xs) - 1, len(ys) - 1
    hcov = lambda y, a, b: any(abs(h[0] - y) <= 2 and min(h[2], b) - max(h[1], a) >= 0.5 * (b - a) for h in hs)
    merge = [[r > 0 and not hcov(ys[r], xs[c], xs[c + 1]) for c in range(ncol)] for r in range(nrow)]
    owner = lambda r, c: owner(r - 1, c) if merge[r][c] else r
    where = lambda x, y: (next((i for i in range(nrow) if ys[i] <= y <= ys[i + 1]), None),
                          next((i for i in range(ncol) if xs[i] <= x <= xs[i + 1]), None))
    cells = [[[] for _ in range(ncol)] for _ in range(nrow)]
    for w in words:
        r, c = where((float(w["x0"]) + float(w["x1"])) / 2, (float(w["top"]) + float(w["bottom"])) / 2)
        if r is not None and c is not None:
            cells[owner(r, c)][c].append(w)
    bold = lambda w: any("Bold" in ch["fontname"] or _fmt(ch)["b"] for ch in chars
                         if float(w["x0"]) - 0.5 <= float(ch["x0"]) <= float(w["x1"]) and abs(ch["top"] - w["top"]) < 2)
    rows = []
    for r in range(nrow):
        row = []
        for c in range(ncol):
            ws = sorted(cells[r][c], key=lambda w: (round(float(w["top"]) / 3), float(w["x0"])))
            lines = {}
            for w in ws:
                lines.setdefault(round(float(w["top"]) / 3), []).append(w)
            out_lines = [(" ".join(w["text"] for w in sorted(lw, key=lambda w: float(w["x0"]))), all(bold(w) for w in lw))
                         for _, lw in sorted(lines.items())]
            align = "left"
            if ws:
                left = min(float(w["x0"]) for w in ws) - xs[c]
                right = xs[c + 1] - max(float(w["x1"]) for w in ws)
                tol = max(6, 0.05 * (xs[c + 1] - xs[c]))
                align = "right" if right < 10 < left - 10 else "center" if abs(left - right) < tol and left > 8 else "left"
            row.append({"lines": out_lines, "align": align})
        rows.append({"cells": row, "bold": False})
    pics = {}
    for im in imgs:
        r, c = where((im["x0"] + im["x1"]) / 2, (im["top"] + im["bottom"]) / 2)
        if r is not None and c is not None:
            pics.setdefault((owner(r, c), c), []).append(im)
    f = Counter((_fmt(ch)["font"], _fmt(ch)["size"]) for ch in chars if ch["text"].strip()).most_common(1)
    return {"rows": rows, "widths": [xs[i + 1] - xs[i] for i in range(ncol)], "aligns": ["left"] * ncol,
            "font": f[0][0][0] if f else None, "size": f[0][0][1] if f else None, "box": True, "inner": True,
            "merge": merge, "imgs": pics, "rules": True}


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
    # внутренние линии — только настоящие вертикали; концы отрезков рамки (у штампа UTS рамка сверху разбита на
    # куски в точках 160 и 448) — места столбцов, но не линии: сетка рисовалась там, где её нет (16.10)
    vx = [(float(r["x0"]) + float(r["x1"])) / 2 for r in rects
          if float(r["x1"]) - float(r["x0"]) < 3 and float(r["bottom"]) - float(r["top"]) > 5]
    out["inner"] = len(cols) >= 3 and any(abs(x - c) <= 3 for c in cols[1:-1] for x in vx)
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
    spans = [None] * ncell
    prev = None                                             # (верх, высота) прошлой строки
    for _, ws in sorted(rows.items()):
        cells = [""] * ncell
        bold = any("Bold" in c["fontname"] for c in chars if abs(c["top"] - ws[0]["top"]) < 2)
        for w in sorted(ws, key=lambda w: w["x0"]):
            ci = max(0, min(ncell - 1, sum(1 for x in cols[1:-1] if float(w["x0"]) >= x - 2)))
            cells[ci] = (cells[ci] + " " + w["text"]).strip()
            side[ci].append((float(w["x0"]) - cols[ci], cols[ci + 1] - float(w["x1"])))
            a, b = float(w["x0"]), float(w["x1"])
            spans[ci] = (min(spans[ci][0], a), max(spans[ci][1], b)) if spans[ci] else (a, b)
        top, h = float(ws[0]["top"]), float(ws[0]["bottom"]) - float(ws[0]["top"])
        # строка, где текст лишь в части ячеек строки выше и сразу под ней, — перенос текста этих ячеек (название в
        # штампе без внутренних линий в две строки давало лишнюю строку таблицы, синтетика 16.10): вторая строка
        # ячейки — своим абзацем, как в оригинале
        if out["rows"] and prev and top - prev[0] <= 1.7 * prev[1]:
            above = out["rows"][-1]["cells"]
            have, had = {i for i, c in enumerate(cells) if c}, {i for i, c in enumerate(above) if c}
            if have < had:
                out["rows"][-1]["cells"] = [a + "\n" + c if c else a for a, c in zip(above, cells)]
                prev = (top, h)
                continue
        prev = (top, h)
        out["rows"].append({"cells": cells, "bold": bold})
    out["spans"] = spans
    for ci, gaps in enumerate(side):
        left = min((g[0] for g in gaps), default=0)
        right = min((g[1] for g in gaps), default=0)
        # по центру — с допуском от ширины ячейки: «Tartalomjegyzék» в подвале ÜFK шириной 494 pt смещён на 7 pt
        tol = max(6, 0.05 * (cols[ci + 1] - cols[ci]))
        out["aligns"].append("right" if right < 10 < left - 10 else "center" if abs(left - right) < tol and left > 10 else "left")
    out["widths"] = [cols[i + 1] - cols[i] for i in range(ncell)]
    f = Counter((_fmt(c)["font"], _fmt(c)["size"]) for c in chars if c["text"].strip()).most_common(1)
    if f:
        out["font"], out["size"] = f[0][0]
    return out
