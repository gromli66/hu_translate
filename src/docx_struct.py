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
INLINE = {"Link", "Span", "Reference", "Annot", "Quote", "Code", "BibEntry"}


# начертание в имени шрифта PDF — и по-английски, и по-местному: Word 2016 с венгерским интерфейсом пишет
# «Times New Roman félkövér» — это полужирный Times, а не шрифт с таким именем (ÜFK, 16.10)
BOLD_W = r"Bold|Black|Heavy|Semibold|Demibold|félkövér|Fett|полужирный|Gras|Negrita|Grassetto"
ITAL_W = r"Italic|Oblique|dőlt|Kursiv|курсив|Italique|Cursiva|Corsivo"


def _words(chars):
    """Слова знаков колонтитула. Допуск по высоте — от кегля: «ő», «ű» запасного шрифта (TimesNewRoman при основном
    Times) стоят на 2,1 pt выше соседей, при допуске 2 pt они отрывались в «слово» «ő ő ő», а в названии не хватало
    букв (30GF118GT, 28.10). Верх слова с таким знаком выше, чем у соседей по строке, — строки колонтитула группируются
    по округлённому верху, и слова с «ő» уходили в соседнюю строку; верх слов одной строки (по низу) выравнивается."""
    import pdfplumber
    sizes = sorted(float(c["size"]) for c in chars if c["text"].strip())
    tol = max(2.0, 0.26 * sizes[len(sizes) // 2]) if sizes else 2.0
    # порядок знаков — как в потоке PDF: название бланка «TESZTELÉSI UTASÍTÁS» залезает под «Dátum:» и дату, по положению
    # знаки двух надписей перемешивались («áDtjááturmha:tóságának», 30RR0102, 28.10)
    words = pdfplumber.utils.extract_words(chars, x_tolerance=2.5, y_tolerance=tol, use_text_flow=True) if chars else []
    order = {id(w): i for i, w in enumerate(words)}
    words.sort(key=lambda w: float(w["bottom"]))
    groups, cur = [], []
    for w in words:
        if cur and float(w["bottom"]) - float(cur[-1]["bottom"]) > tol:
            groups.append(cur)
            cur = []
        cur.append(w)
    groups.append(cur)
    for g in (g for g in groups if g):
        tops = sorted(float(w["top"]) for w in g)
        med = tops[len(tops) // 2]
        for w in g:
            if abs(float(w["top"]) - med) <= tol:
                w["top"] = med
    return sorted(words, key=lambda w: order[id(w)])


def _upright(c):
    """Знак не повёрнут — по матрице (у повёрнутого на 90° текста JPDS pdfplumber даёт upright=True)."""
    m = c.get("matrix")
    return abs(float(m[0])) >= abs(float(m[1])) if m else c.get("upright", True)


def _underline(chars, objs):
    """Подчёркивание (ссылки на законы OAH 4.22: тонкий прямоугольник 0,6 pt под строкой) — ch["u"]: линия до 1,2 pt
    толщиной ниже базовой линии, но внутри рамки знака (от 0,7 кегля сверху до 0,8 pt над низом), под серединой знака.
    Граница ячейки идёт по низу рамки знака и ниже (ÜFK) — не подчёркивание, зачёркивание — выше."""
    from collections import defaultdict
    by_y = defaultdict(list)
    for o in objs:
        if float(o["height"]) <= 1.2 and float(o["width"]) >= 2:
            y = (float(o["top"]) + float(o["bottom"])) / 2
            by_y[int(y // 4)].append((y, float(o["x0"]), float(o["x1"])))
    if not by_y:
        return
    for ch in chars:
        top, size = float(ch["top"]), float(ch["size"])
        bottom = float(ch.get("bottom", top + size))
        cx = (float(ch["x0"]) + float(ch["x1"])) / 2
        lo, hi = top + 0.7 * size, bottom - 0.8
        if any(lo <= y <= hi and x0 - 1 <= cx <= x1 + 1
               for k in range(int(lo // 4), int(hi // 4) + 1) for y, x0, x1 in by_y.get(k, ())):
            ch["u"] = True


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
            "color": color, "u": bool(ch.get("u"))}


SUP = str.maketrans("0123456789+-oоº°", "⁰¹²³⁴⁵⁶⁷⁸⁹⁺⁻°°°°")
SUB = str.maketrans("0123456789+-", "₀₁₂₃₄₅₆₇₈₉₊₋")


def _runs(items):
    """Символы и картинки одного элемента → прогоны одинакового начертания. Перенос строки — пробел, широкий
    разрыв в строке (табуляция Word) — «\\t»."""
    runs, prev, base = [], None, None                     # base — последний знак обычного кегля (для индексов)
    hang = _hanging_tab([it["ch"] for it in items if "ch" in it])
    just = _justified_gaps([it["ch"] for it in items if "ch" in it], hang)
    bot = lambda c: float(c.get("bottom", float(c["top"]) + float(c["size"])))
    for it in items:
        if "img" in it:
            runs.append({"img": it["img"]}); prev = None; continue
        ch = it["ch"]
        t = ch["text"]
        if t and (0x1D400 <= ord(t[0]) <= 0x1D7FF or 0x2100 <= ord(t[0]) <= 0x214F)                 and len(unicodedata.normalize("NFKC", t)) == 1:
            # формула Word (Cambria Math: «𝑏𝑖𝑛á𝑟𝑖𝑠») — обычными буквами, иначе не переводится; «№», «℃» — как есть
            # (NFKC давал «No1» вместо «№1», ИН-1172, 17.10)
            t = unicodedata.normalize("NFKC", t)
        # индекс — мелкий знак выше или ниже строки рядом с обычным: знаком Юникода, он переживёт и перевод («545 ºс»
        # ИН-1065 печаталось «545 ос», «H₃BO₃» ÜFK II — «H3BO3», 17.10)
        if base is not None and prev is not None and t.strip() and float(ch["size"]) < 0.8 * float(base["size"])                 and abs(float(ch["x0"]) - float(prev["x1"])) < 2 * float(base["size"]):
            if bot(ch) < bot(base) - 0.2 * float(base["size"]):
                t = t.translate(SUP)
            elif bot(ch) > bot(base) + 0.1 * float(base["size"]):
                t = t.translate(SUB)
        elif t.strip():
            base = ch
        if prev is not None:
            # та же строка — и у индекса («H₃BO₃»: «3» ниже и мельче, после него вставлялся пробел, 17.10): знаки
            # перекрываются по высоте
            same_line = abs(ch["top"] - prev["top"]) < 0.5 * prev["size"] or min(bot(ch), bot(prev)) - max(
                float(ch["top"]), float(prev["top"])) > 0.3 * min(bot(ch) - float(ch["top"]), bot(prev) - float(prev["top"]))
            if not same_line and runs and "t" in runs[-1] and not runs[-1]["t"].endswith((" ", "\t")):
                t = " " + t
            elif same_line and (ch["x0"] - prev["x1"] > 1.8 * prev["size"] or id(ch) in hang) and t != " " \
                    and id(ch) not in just and runs and "t" in runs[-1] and not runs[-1]["t"].endswith("\t"):
                runs[-1]["t"] = runs[-1]["t"].rstrip(" ") + "\t"
                runs[-1].setdefault("tabs", []).append(float(ch["x0"]))     # позиция табуляции — по оригиналу
        f = _fmt(ch)
        if runs and "t" in runs[-1] and all(runs[-1][k] == f[k] for k in f):
            runs[-1]["t"] += t
            if t.strip():                       # правый край — по видимым знакам (хвостовой пробел потом срезается)
                runs[-1]["x1"] = float(ch["x1"])
        else:                                   # начало прогона на листе — место для значка в строке (insert_imgs)
            runs.append(dict(f, t=t, x=float(ch["x0"]), x1=float(ch["x1"]) if t.strip() else None, top=float(ch["top"]),
                             page=ch.get("page_number")))
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


def _justified_gaps(chars, hang=()):
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
        # Word растягивает только текст после последней табуляции: промежутки кода до табуляции к висячему
        # отступу («MSZ EN 571-1:2001 ⇥ Roncsolásmentes … vizsgálatok.», OAH 4.22) — обычные пробелы, и с ними
        # медианой выходила сама табуляция: табуляцией становилась растяжка, 19.10
        k = max((i for i, c in enumerate(ln) if id(c) in hang), default=0)
        gaps = [(float(b["x0"]) - float(a["x1"]), b) for a, b in zip(ln[k:], ln[k + 1:])
                if float(b["x0"]) - float(a["x1"]) > 0.15 * float(a["size"])]
        # растяжка — промежутки, равные медиане строки; заметно больший — настоящая табуляция («MSZ 27003-0:2013 ⇥
        # Nukleáris létesítmények…» перечня OAH: одна табуляция ломала прежнее «все равны», и растяжка шла
        # табуляциями, 16.10)
        med = sorted(g for g, _ in gaps)[len(gaps) // 2] if gaps else 0
        out |= {id(b) for g, b in gaps if g > 1.8 * float(b["size"]) and abs(g - med) < 1.0}
    return out


def _hanging_tab(chars):
    """Знак после промежутка в первой строке, стоящий ровно на висячем отступе (где начинаются прочие строки), —
    табуляция Word к отступу: «MSZ 27003-0:2013 ⇥ Nukleáris…» перечня OAH — промежуток 17,6 pt меньше порога
    «1,8 кегля», и строка шла растяжкой по ширине (18.10). → множество id знаков."""
    ink = [c for c in chars if c["text"].strip()]
    if len(ink) < 3:
        return set()
    first = [c for c in ink if abs(float(c["top"]) - float(ink[0]["top"])) < 0.5 * float(ink[0]["size"])]
    rest = [c for c in ink if c not in first and float(c["top"]) > float(ink[0]["top"])]
    if not rest:
        return set()
    lines, cur = [], [rest[0]]                              # начала прочих строк — один висячий отступ
    for c in rest[1:]:
        if abs(float(c["top"]) - float(cur[0]["top"])) < 0.5 * float(cur[0]["size"]):
            cur.append(c)
        else:
            lines.append(cur); cur = [c]
    lines.append(cur)
    x = min(float(ln[0]["x0"]) for ln in lines)
    if any(abs(float(ln[0]["x0"]) - x) > 1.5 for ln in lines) or x < float(first[0]["x0"]) + 20:
        return set()
    return {id(b) for a, b in zip(first, first[1:])
            if abs(float(b["x0"]) - x) < 1.5 and float(b["x0"]) - float(a["x1"]) > 0.5 * float(a["size"])}


def _tail_pairs(last, r):
    """Хвост строки last на новом листе (см. _cut_rows) → [(столбец в last, ячейка хвоста)] или []."""
    cell_t = lambda c: " ".join(text_of(x["runs"]) for x in c.get("blocks", []) if x["k"] == "p").strip()
    page = lambda row: next((c["geom"]["page"] for c in row if c.get("geom")), None)
    if not (last and r and len(r) <= len(last) and (page(r) or 0) > (page(last) or 0)) \
            or any((c.get("attrs") or {}).get("ColSpan") for c in r):
        return []
    # продолжение списка: ячейка выше кончается списком, хвост начат списком («– MOD …» | «– ADD …», JPDS табл. 15)
    lists = lambda d, c: bool(d.get("blocks") and c.get("blocks")) and d["blocks"][-1]["k"] == "list" \
        and c["blocks"][0]["k"] == "list"
    pairs = []
    for c in r:
        g = c.get("geom")
        j = next((j for j, d in enumerate(last) if g and d.get("geom") and abs(d["geom"]["x0"] - g["x0"]) < 3), None)
        t = cell_t(c)
        # продолжение фразы — со строчной: «ячейка выше без точки» не годится — у шапок точек нет, и строка «1 | …»
        # нового листа уходила в шапку (синтетика, 16.10)
        if j is None or t and not t[0].islower() and not lists(last[j], c):
            return []
        pairs.append((j, c))
    return pairs if any(cell_t(c) or lists(last[j], c) for j, c in pairs) else []


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
            elif head and tail and head[-1]["k"] == "list" == tail[0]["k"] and head[-1].get("kind") == tail[0].get("kind"):
                head[-1] = dict(head[-1], items=head[-1]["items"] + tail[0]["items"])     # список, разрезанный листом
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


def fit_rows(bs):
    """Строка, где ячеек больше, чем столбцов рамки таблицы: Word отдаёт вложенную в ячейку табличку («значок ⚠ + текст»
    в «Описании» JPDS табл. 15) лишними ячейками строки — вся таблица получала лишний столбец, текст сжимался в узкий
    (19.10). Ячейки, что лежат в одном столбце рамки (по середине текста), — одной ячейкой, блоки по порядку."""
    for b in bs:
        for it in b.get("items") or []:
            fit_rows(it)
        if b["k"] != "table":
            continue
        xs = (b.get("frame") or {}).get("xs") or []
        k = len(xs) - 1
        if k < 2:
            continue
        col = lambda c: (lambda g: None if not g or g.get("x1") is None else
                         sum((g["x0"] + g["x1"]) / 2 >= x for x in xs[1:-1]))(c.get("geom"))
        rows = []
        for r in b["rows"]:
            if len(r) <= k:
                rows.append(r)
                continue
            cs = [col(c) for c in r]
            if any(x is None for x in cs[1:]) or cs[1:] != sorted(cs[1:]):
                rows.append(r)
                continue
            out = []
            for c, ci in zip(r, cs):
                if out and ci is not None and out[-1][1] == ci:
                    prev = out[-1][0]
                    g0, g1 = prev.get("geom") or {}, c.get("geom") or {}
                    geom = dict(g0, x0=min(g0.get("x0", g1["x0"]), g1["x0"]), x1=max(g0.get("x1", g1["x1"]), g1["x1"]),
                                top=min(g0.get("top", g1["top"]), g1["top"]),
                                bottom=max(g0.get("bottom", g1["bottom"]), g1["bottom"])) if g1 else g0
                    out[-1] = (dict(prev, blocks=prev["blocks"] + c["blocks"], geom=geom), ci)
                else:
                    out.append((c, ci))
            rows.append([c for c, _ in out] if len(out) <= k else r)
        b["rows"] = rows
    return bs


def sub_rows(bs):
    """Метки, выровненные пустыми абзацами против своих пунктов в соседней ячейке (ÜFK, «Beavatkozások»: «A.1.»,
    «A.2.1.» … и сроки «A.2.1. 2 óra» — каждая против своего действия), — подстроками таблицы: пустые абзацы пропадали,
    метки вставали стопкой вверху ячейки и уезжали от своих пунктов (а в переводе длина пунктов другая — отбивкой не
    выровнять, 19.10). Режутся ячейки-метки (короткие абзацы с пустыми между ними) и ячейка-пара, у которой абзацы
    начинаются на тех же высотах; прочие ячейки строки — одной ячейкой на все подстроки (c["vcont"] — продолжение
    сверху). Линий между подстроками нет."""
    blank = lambda x: x["k"] == "p" and not text_of(x.get("runs") or []).strip()
    top = lambda x: (x.get("geom") or {}).get("top")
    for b in bs:
        for it in b.get("items") or []:
            sub_rows(it)
        if b["k"] != "table":
            continue
        rows = []
        for r in b["rows"]:
            labels, tops = [], None
            for ci, c in enumerate(r):
                bl = c.get("blocks") or []
                full = [x for x in bl if not blank(x)]
                if len(full) >= 2 and len(full) < len(bl) and all(x["k"] == "p" and top(x) is not None
                                                                  and len(text_of(x["runs"]).strip()) <= 20 for x in full):
                    ts = [top(x) for x in full]
                    if tops is None or len(ts) == len(tops) and all(abs(a - z) < 2 for a, z in zip(ts, tops)):
                        labels.append(ci)
                        tops = tops or ts
            # пара — ячейка сразу за первой ячейкой меток (метка стоит перед своим пунктом); условие слева, чьи абзацы
            # случайно начаты на тех же высотах, — не пара
            pair = [ci for ci, c in enumerate(r) if labels and ci == labels[0] + 1 and ci not in labels
                    and all(any(top(x) is not None and abs(top(x) - t) < 2 for x in c.get("blocks") or [] if not blank(x))
                            for t in tops)]
            if not labels or not pair:
                rows.append(r)
                continue
            cut = labels + pair
            n = len(tops)
            part = lambda c, k: [x for x in c["blocks"] if not blank(x) and top(x) is not None
                                 and (k == 0 or top(x) >= tops[k] - 2) and (k == n - 1 or top(x) < tops[k + 1] - 2)]
            edge = lambda c, k: dict(c.get("edges") or {}, **({"top": False} if k else {}),
                                     **({"bottom": False} if k < n - 1 else {}))
            for k in range(n):
                row = []
                for ci, c in enumerate(r):
                    if ci in cut:
                        g = dict(c.get("geom") or {}, top=tops[k]) if c.get("geom") else None
                        row.append(dict(c, blocks=part(c, k), geom=g, **({"edges": edge(c, k)} if "edges" in c else {})))
                    elif k == 0:
                        row.append(dict(c, **({"edges": edge(c, 0)} if "edges" in c else {})))
                    else:
                        row.append({"blocks": [], "th": c.get("th", False), "vcont": True, "geom": None,
                                    **({"edges": edge(c, k)} if "edges" in c else {})})
                rows.append(row)
        b["rows"] = rows
    return bs


def _same_cols(a, b, tol=3.0):
    """Две таблицы по линиям листа (docx_layout: frame.row_lines) — с теми же столбцами (на тех же местах: выравнивание
    столбцов склеенной таблицы голосует по рамке первой части, docx_build._col_align): лист 2, начатый другой таблицей
    с тем же числом столбцов, — не продолжение (ред-тим 20.10). Таблицы без линий листа (по разметке) — не
    проверяются."""
    fa, fb = a.get("frame") or {}, b.get("frame") or {}
    if not (fa.get("row_lines") and fb.get("row_lines")):
        return True
    xa, xb = fa.get("xs") or [], fb.get("xs") or []
    return len(xa) == len(xb) and all(abs(p - q) <= tol for p, q in zip(xa, xb))


def _other_grid(a, b):
    """Таблицы листов по линиям (docx_layout: frame.row_lines) — с разным числом столбцов сетки: это разные таблицы, даже
    если шапка повторена (30BOXFELTEL: форма на 36 листов — таблицы по 10, 12, 18 столбцов с одной шапкой; склеенные,
    они делили сетку первой, и ширин не было — столбцы поровну, описание в 40 pt, ячейка с галочкой на весь остаток)."""
    fa, fb = a.get("frame") or {}, b.get("frame") or {}
    return bool(fa.get("row_lines") and fb.get("row_lines") and fa.get("xs") and fb.get("xs")
                and len(fa["xs"]) != len(fb["xs"]))


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
            elif len(b["rows"]) > 1 and row_t(b["rows"][0]) == row_t(prev["rows"][0]) and all(row_t(b["rows"][0])):
                rep = 1                     # шапка без TH (JPDS), дословный повтор первой строки
        if prev is not None and prev["k"] == "table" and b["k"] == "table" and prev["rows"] and b["rows"] \
                and (rep or not all(c["th"] for c in b["rows"][0]) or _tail_pairs(prev["rows"][-1], b["rows"][0])) \
                and (rep or max(map(len, b["rows"])) == max(map(len, prev["rows"]))
                     or _tail_pairs(prev["rows"][-1], b["rows"][0])) \
                and ((rep and not _other_grid(prev, b)) or (not rep and _same_cols(prev, b))) \
                and (b.get("geom") or {}).get("page", 0) > (prev.get("geom") or {}).get("page", 0):
            rows = b["rows"][rep:]
            # картинки, что Word вложил в повтор шапки (значки строк прошлого листа, JPDS табл. 15), — не прочь:
            # их по месту на листе переселит adopt
            strays = [x for r in b["rows"][:rep] for c in r for x in c.get("blocks", [])
                      if x["k"] == "p" and not text_of(x.get("runs") or []).strip()
                      and any("img" in q for q in x.get("runs") or [])]
            # таблица через лист: повтор шапки — только если оригинал её повторял (сборщик, w:tblHeader)
            out[-1] = dict(prev, rows=_cut_rows(prev["rows"] + rows), split=True,
                           head_repeat=bool(rep) or prev.get("head_repeat", False),
                           # строки, что оригинал повторял на каждом листе, — шапка, даже без TH (сборщик: w:tblHeader)
                           repeat_rows=max(rep, prev.get("repeat_rows", 0)))
            # линии строк обеих частей — одной рамкой (по ним высота пустых строк, docx_build): у склеенной таблицы
            # их число не сходилось со строками, и пустые графы продолжения схлопывались
            ya, yb = (prev.get("frame") or {}).get("ys") or [], (b.get("frame") or {}).get("ys") or []
            if len(ya) == len(prev["rows"]) + 1 and len(yb) == len(b["rows"]) + 1 \
                    and len(out[-1]["rows"]) == len(prev["rows"]) + len(rows):
                # последняя строка части, растянутая рамкой до низа листа (формы RDGT: под текстом пусто до рамки
                # листа), — по своему тексту: пустота шла в DOCX высотой строки, и таблица разрасталась на листы (20.10)
                low = max((c["geom"]["bottom"] for c in prev["rows"][-1] if c.get("geom")), default=None)
                if low is not None and ya[-1] - low > 24:
                    ya = ya[:-1] + [round(max(low + 4, ya[-2] + 7), 1)]
                yb = yb[rep:]                       # без линий повтора шапки
                out[-1]["frame"] = dict(prev["frame"], ys=ya + [round(ya[-1] + y - yb[0], 1) for y in yb[1:]])
            if strays:
                out[-1]["strays"] = prev.get("strays", []) + strays
            continue
        if b["k"] == "table":
            b = dict(b, rows=_cut_rows(b["rows"]))
        out.append(b)
    return out


def split_at_zones(blocks, runs):
    """Раздел с другой шапкой (серия листов со своим колонтитулом, layout → variants) начинается между блоками, а таблица
    через лист (join_tables) — один блок: её режет там, где у листа другая серия; у следующей части — копия строк, что
    оригинал повторял на каждом листе (шапка таблицы)."""
    zone = lambda pn: next((i for i, z in enumerate(runs or []) if z["from"] <= pn <= z["to"]), None)
    out = []
    for b in blocks:
        if not runs or b["k"] != "table" or not b.get("split"):
            out.append(b)
            continue
        rows = b["rows"]
        ys = (b.get("frame") or {}).get("ys") or []
        rep = b.get("repeat_rows", 0) if b.get("head_repeat") else 0
        pages, last = [], (b.get("geom") or {}).get("page")
        for r in rows:                                      # лист строки — по её абзацам; пустая строка — как прошлая
            last = min((x["geom"]["page"] for c in r for x in c["blocks"] if (x.get("geom") or {}).get("page")),
                       default=last)
            pages.append(last)
        cuts = [i for i in range(max(rep, 1), len(rows)) if zone(pages[i]) != zone(pages[i - 1])]
        for a, z in zip([0] + cuts, cuts + [len(rows)]):
            head = rows[:rep] if a else []
            part = dict(b, rows=head + rows[a:z], geom=dict(b["geom"], page=pages[a]) if b.get("geom") else b.get("geom"))
            if len(ys) == len(rows) + 1:                    # линии строк — своей части; у повтора шапки — свои
                part["frame"] = dict(b["frame"], ys=ys[:z + 1] if not a else ys[:rep + 1] + [
                    round(ys[rep] + y - ys[a], 1) for y in ys[a + 1:z + 1]] if rep else ys[a:z + 1])
            out.append(part)
    return out


def attach_notes(bs):
    """Сноска Word (элемент Note: «1 Знаком «*» отмечены…» внизу листа) — к своему абзацу: тому, что перед ней, со
    знаком-ссылкой (мелкий знак с тем же номером). Прежде Note шла частью соседнего абзаца — текст сноски вливался в
    абзац (ТЗ ВШЭ, 16.10). Знак ссылки из абзаца прочь, место — долей длины текста; номер сноски даёт Word.
    Абзац получает "notes": [{"runs", "pos"}]."""
    out = []
    for b in bs:
        if b["k"] == "list":
            b = dict(b, items=[attach_notes(it) for it in b["items"]])
        elif b["k"] == "table":
            b = dict(b, rows=[[dict(c, blocks=attach_notes(c.get("blocks", []))) for c in r] for r in b["rows"]])
        if b["k"] != "note":
            out.append(b)
            continue
        runs = [dict(r) for r in b["runs"]]
        m = re.match(r"\s*(\d{1,3}|\*{1,3})\s*", text_of(runs))
        num = m.group(1) if m else None
        cut = m.end() if m else 0
        for r in runs:                                      # номер сноски — из текста прочь
            if cut <= 0 or "t" not in r:
                continue
            n = min(cut, len(r["t"]))
            r["t"], cut = r["t"][n:], cut - n
        runs = [r for r in runs if "img" in r or r.get("t")]
        if not runs:
            continue
        host, ref = None, None
        for x in reversed(out):                             # абзац со знаком ссылки — ближайший перед сноской
            for par in _host_paras(x):
                sizes = [r.get("size") or 0 for r in par.get("runs", []) if "t" in r and r["t"].strip()]
                big = max(sizes, default=0)
                k = next((i for i in range(len(par["runs"]) - 1, -1, -1)       # знак ссылки бывает и «¹» (индекс)
                          if num and "t" in par["runs"][i]
                          and unicodedata.normalize("NFKC", par["runs"][i]["t"].strip()) == num
                          and (par["runs"][i].get("size") or big) < 0.85 * big), None)
                if k is not None:
                    host, ref = par, k
                    break
            if host is not None or x["k"] in ("p", "list", "table"):
                break
        if host is None:                                    # знака не нашли — к абзацу перед сноской, в конец
            host = next((par for x in reversed(out) for par in _host_paras(x)), None)
        if host is None:                                    # абзаца перед сноской нет — сноска абзацем, не пропадает
            out.append({"k": "p", "runs": runs, "geom": b.get("geom")})
            continue
        before = len(text_of(host["runs"][:ref])) if ref is not None else None
        if ref is not None:
            host["runs"] = host["runs"][:ref] + host["runs"][ref + 1:]
        total = max(1, len(text_of(host["runs"])))
        host.setdefault("notes", []).append({"runs": runs, "pos": 1.0 if before is None else min(1.0, before / total)})
    return out


def _host_paras(b):
    """Абзацы блока с конца (для поиска знака сноски): сам абзац или последние абзацы пунктов/ячеек."""
    if b["k"] == "p":
        return [b]
    if b["k"] == "list":
        return [p for it in reversed(b["items"]) for x in reversed(it) for p in _host_paras(x)]
    if b["k"] == "table":
        return [p for r in reversed(b["rows"]) for c in reversed(r) for x in reversed(c.get("blocks", []))
                for p in _host_paras(x)]
    return []


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


def adopt(bs):
    """Абзац вне таблицы, что на листе лежит внутри её строки: Word разметил надписи и подложки ячеек (образцы «Текст
    строки» в таблице 16 ТЭ357) отдельными абзацами до таблицы — печатались строкой над рисунком, ячейки пустые
    (17.10), — в ячейку по строке и столбцу, по месту среди её блоков. Картинка в ячейке с листа раньше текста своей
    строки (значки строк прошлого листа в повторе шапки, JPDS табл. 15) — так же, по месту."""
    out = []
    for b in bs:
        out.append(b)
        if b["k"] != "table":
            continue
        out += b.pop("strays", [])                          # из повтора шапки (join_tables)
        for r in b["rows"]:
            home = min((x["geom"]["page"] for c in r for y in c.get("blocks", []) for x in _flat_paras(y)
                        if (x.get("geom") or {}).get("page") and text_of(x.get("runs") or []).strip()), default=None)
            for c in r:
                far = [x for x in c.get("blocks", []) if home and x["k"] == "p" and (x.get("geom") or {}).get("page")
                       and x["geom"]["page"] < home and not text_of(x.get("runs") or []).strip()
                       and any("img" in q for q in x.get("runs") or [])]
                if far:
                    c["blocks"] = [x for x in c["blocks"] if x not in far]
                    out += far
    for t in [b for b in out if b["k"] == "table"]:
        rows = t["rows"]

        def places(bs_):                                    # места блоков ячейки (геометрия ячейки — до evict)
            for y in bs_:
                if y["k"] == "list":
                    for it in y["items"]:
                        yield from places(it)
                elif (y.get("geom") or {}).get("bottom") and y["geom"].get("x1") is not None:
                    yield y["geom"]
        cg = [[list(places(c["blocks"])) for c in r] for r in rows]
        # (лист, верх, низ) текста строки на первом её листе: значки — не в счёт (Word кладёт значки нескольких строк
        # в первую ячейку продолжения таблицы, и они растягивали её строку на соседние, JPDS табл. 65, 18.10)
        icon = lambda y: y["k"] == "p" and not text_of(y.get("runs") or []).strip() and any(
            "img" in q for q in y.get("runs") or [])

        def texts(bs_):
            for y in bs_:
                if y["k"] == "list":
                    for it in y["items"]:
                        yield from texts(it)
                elif not icon(y) and (y.get("geom") or {}).get("bottom") and y["geom"].get("x1") is not None:
                    yield y["geom"]
        span = []
        for r in rows:
            gs = [g for c in r for g in texts(c["blocks"])]
            pn = min((g["page"] for g in gs), default=None)
            gs = [g for g in gs if g["page"] == pn]
            span.append((pn, min(g["top"] for g in gs), max(g["bottom"] for g in gs)) if gs else None)
        cols = {}                                           # число ячеек строки → {столбец: (лево, право)}
        for r, gr in zip(rows, cg):
            for k, gc in enumerate(gr):
                if gc:
                    lo, hi = cols.setdefault(len(r), {}).get(k, (gc[0]["x0"], gc[0]["x1"]))
                    cols[len(r)][k] = (min([lo] + [g["x0"] for g in gc]), max([hi] + [g["x1"] for g in gc]))
        if not cols:
            continue
        X0 = min(v[0] for cs in cols.values() for v in cs.values())
        X1 = max(v[1] for cs in cols.values() for v in cs.values())
        def place(x):
            """(строка, столбец) таблицы, в которой лежит на листе блок x, или None."""
            g = x.get("geom") or {}
            if x["k"] != "p" or not g.get("bottom") or g.get("x1") is None:
                return None
            cx, cy = (g["x0"] + g["x1"]) / 2, (g["top"] + g["bottom"]) / 2
            if not X0 - 2 <= cx <= X1 + 2:
                return None
            row = None
            for i, s in enumerate(span):
                if not s or s[0] != g["page"]:
                    continue
                prev = next((span[j] for j in range(i - 1, -1, -1) if span[j]), None)
                nxt = next((span[j] for j in range(i + 1, len(span)) if span[j]), None)
                y0 = (prev[2] + s[1]) / 2 if prev and prev[0] == s[0] else s[1] - 2
                y1 = (s[2] + nxt[1]) / 2 if nxt and nxt[0] == s[0] else s[2] + 2
                if y0 <= cy < y1:
                    row = i
                    break
            if row is None or len(rows[row]) < 1:
                return None
            r = rows[row]
            cs = dict(cols.get(len(r), {}))
            for k in range(len(r)):                         # столбец без места — между соседями
                if k not in cs:
                    lo = cs[k - 1][1] + 1 if k - 1 in cs else X0
                    hi = next((cs[q][0] for q in range(k + 1, len(r)) if q in cs), X1) - 1
                    cs[k] = (lo, max(lo, hi))
            bounds = [(cs[k][1] + cs[k + 1][0]) / 2 for k in range(len(r) - 1)]
            return row, sum(cx >= b for b in bounds)

        def put(x, row, k):
            g = x["geom"]
            blocks = rows[row][k]["blocks"]
            key = (g["top"], g["x0"])
            at = next((n for n, y in enumerate(blocks) if (y.get("geom") or {}).get("bottom")
                       and (y["geom"]["top"], y["geom"].get("x0", 0)) > key), len(blocks))
            blocks.insert(at, x)
        for x in list(out):
            rk = place(x)
            if rk is not None:
                put(x, *rk)
                out.remove(x)
        for i, r in enumerate(rows):                        # значок в чужой строке — в свою (по месту на листе)
            for k, c in enumerate(r):
                for x in [y for y in c.get("blocks", []) if icon(y)]:
                    rk = place(x)
                    if rk is not None and rk[0] != i:
                        c["blocks"].remove(x)
                        put(x, *rk)
    return out


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

        def clean(blocks, cpage, prev):
            keep = []
            for x in blocks:
                q = _pos(x)
                # и на более позднем листе, чем сама ячейка, или на том же, но между ним и прошлым блоком ячейки — то,
                # что идёт после таблицы (подпись «Рисунок 60» под рисунком — в последней ячейке таблицы над ним, JPDS
                # л. 69: печаталась в таблице, 17.10)
                if lim and q and q > lim and (q[0] > cpage or prev[0] and prev[0][0] == q[0] and any(
                        a[0] == q[0] and prev[0][1] <= a[1] < q[1] for a in after)):
                    moved.extend(_flat_paras(x))
                    continue
                if x["k"] == "list":
                    items = [clean(it, cpage, prev) for it in x["items"]]
                    items = [it for it in items if it]
                    if not items:
                        continue
                    x = dict(x, items=items)
                elif q and (x.get("geom") or {}).get("bottom"):
                    prev[0] = (q[0], x["geom"]["bottom"])
                keep.append(x)
            return keep
        rows = [[dict(c, blocks=clean(c.get("blocks", []), (c.get("geom") or {}).get("page") or 0, [None])) for c in r]
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


def join_pages(bs, cell=False):
    """Абзац, разрезанный листом: Word размечает его половины двумя P — в DOCX выходил разрыв абзаца посреди фразы
    (с отступом первой строки у второй половины, синтетика 16.10). Соседние абзацы на соседних листах — один, если
    первый не кончен знаком конца фразы, а второй начат строчной или без отступа первой строки, который есть у
    первого. Внутри пунктов списков и ячеек — так же. В ячейке слово прописными, перенесённое с дефисом в следующий
    абзац («ELLENŐR-» / «ZÉS HELYE», шапка таблицы 30YCGT), — один абзац: с текстом страницы («ELLENŐR- ZÉS HELYE») он
    совпадает, а половинами переводился «МЕСТО» / «КОНТРОЛЯ»."""
    out = []
    for b in bs:
        if b["k"] == "list":
            b = dict(b, items=[join_pages(it, cell) for it in b["items"]])
        elif b["k"] == "table":
            b = dict(b, rows=[[dict(c, blocks=join_pages(c.get("blocks", []), True)) for c in r] for r in b["rows"]])
        a = out[-1] if out else None
        ga, gb = (a or {}).get("geom") or {}, b.get("geom") or {}
        if cell and a and a["k"] == b["k"] == "p" and a["runs"] and b["runs"] and "t" in a["runs"][-1] \
                and "t" in b["runs"][0] and ga.get("page") == gb.get("page"):
            ta, tb = text_of(a["runs"]).rstrip(), text_of(b["runs"]).lstrip()
            if re.search(r"[A-ZÁÉÍÓÖŐÚÜŰ]{3,}-$", ta) and tb[:1].isupper() and len(tb) <= 40:
                runs = list(a["runs"]) + [dict(r) for r in b["runs"]]
                runs[len(a["runs"])]["t"] = " " + runs[len(a["runs"])]["t"].lstrip()
                out[-1] = dict(a, runs=runs, **join_notes(a, b))
                continue
        if a and a["k"] == b["k"] == "p" and ga.get("page") and gb.get("page") and gb["page"] == ga["page"] + 1 \
                and a["runs"] and b["runs"] and "t" in a["runs"][-1] and "t" in b["runs"][0]:
            ta, tb = text_of(a["runs"]).rstrip(), text_of(b["runs"]).lstrip()
            indent = ga.get("first_x0", 0) - ga.get("x0", 0) > 5 and abs(gb.get("first_x0", 0) - ga.get("x0", 0)) < 3
            # последняя строка на листе — во всю ширину (выключка по ширине): фраза продолжается и с прописной
            # («…az Országos | Atomenergia Hivatal …», OAH 1.1 л. 6–7: половины переводились порознь, фраза
            # повторялась, 19.10)
            full = (ga.get("lines") or 1) >= 2 and ga.get("last_x1") is not None and ga["x1"] - ga["last_x1"] < 4
            # начало — метка пункта («cb) …», «d) …», «2.», «–»): новый пункт, а не продолжение фразы, хоть и со строчной
            # (OAH N3a л. 44–45: «ca) …», «cb) …», «d) …» сливались в один абзац, 19.10)
            label = re.match(r"\(?[a-zа-яё]{1,3}\)|\d+(\.\d+)*[.)]\s|[–•\-]\s", tb)
            if ta and tb and not label and not re.search(r"[.!?:;]$", ta) \
                    and (tb[0].islower() or indent or full and tb[0].isalpha()):
                runs = list(a["runs"]) + [dict(r) for r in b["runs"]]
                k = len(a["runs"])
                if not runs[k - 1]["t"].endswith((" ", "-", "\u00ad")):
                    runs[k]["t"] = " " + runs[k]["t"].lstrip()
                out[-1] = dict(a, runs=runs, **join_notes(a, b))
                continue
        out.append(b)
    return out


def join_notes(a, b):
    """Сноски двух склеиваемых абзацев (a + b) — место пересчитано долей длины общего текста: {"notes": […]} или {}.
    Раньше при склейке оставались только сноски первой половины (ревью 17.10)."""
    na, nb = a.get("notes") or [], b.get("notes") or []
    if not na and not nb:
        return {}
    la, lb = len(text_of(a.get("runs") or [])), len(text_of(b.get("runs") or []))
    n = max(la + lb, 1)
    return {"notes": [dict(x, pos=x.get("pos", 1.0) * la / n) for x in na]
            + [dict(x, pos=(la + x.get("pos", 1.0) * lb) / n) for x in nb]}


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
        # заголовок оглавления вне элемента TOC (абзац «TARTALOMJEGYZÉK» перед ним на том же листе, OAH 4.22) — его
        # заголовком: сборка ставила свой «Содержание», и выходило два (18.10)
        if b["k"] == "toc" and not b["info"].get("title") and prev is not None and prev["k"] in ("p", "h"):
            t = text_of(prev.get("runs") or []).strip()
            ink = [r for r in prev.get("runs") or [] if "t" in r and r["t"].strip()]
            if t and len(t) <= 40 and not re.search(r"\d\s*$", t) and ink and \
                    (prev.get("geom") or {}).get("page") == b["info"].get("page"):
                r0 = ink[0]
                b["info"]["title"] = {"t": t, "font": r0.get("font"), "size": r0.get("size"), "b": r0.get("b"),
                                      "i": r0.get("i"), "color": r0.get("color"),
                                      "center": prev.get("align") == "center" or _centered_geom(prev)}
                out.remove(prev)
        out.append(b)
    return out


def _centered_geom(b):
    """Абзац по центру листа A4/Letter (по середине своей строки)."""
    g = b.get("geom") or {}
    return g.get("x1") is not None and abs((g["x0"] + g["x1"]) / 2 - 297.6) < 30


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
OWN_LABEL = re.compile(r"^\s*(?:\(\w{1,3}\)|[a-z]{2}\)|\d{1,3}[a-z]\))\s")      # «(a)», «(5a)», «ba)», «12a)»


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


def _continued(flat):
    """Пункт, перенесённый на следующий лист: Word размечает продолжение отдельным пунктом без метки (часто в новом
    списке «точками»: «csövek hegesztésénél…» на листе 20 после «d) … csak olyan csőelágazások vagy», OAH 16.10) —
    в DOCX выходил лишний пункт с «•», и нумерация дальше сбивалась. Пункт без метки, начатый строчной, на более
    позднем листе, чем недописанный (без знака конца фразы) пункт того же уровня перед ним, — его продолжение."""
    out = []
    for it in flat:
        prev = out[-1] if out else None
        p0 = next((x for x in it["body"] if x["k"] == "p"), None)
        last = next((x for x in reversed(prev["body"]) if x["k"] == "p"), None) if prev else None
        t0 = text_of(p0["runs"]).lstrip() if p0 and p0["runs"] else ""
        tl = text_of(last["runs"]).rstrip() if last and last["runs"] else ""
        pg = lambda x: (x.get("geom") or {}).get("page") or 0
        if prev and prev["depth"] == it["depth"] and not it["lbl"] and not it.get("label") and t0[:1].islower() \
                and tl and not re.search(r"[.!?:;]$", tl) and pg(p0) > pg(last) \
                and "t" in last["runs"][-1] and "t" in p0["runs"][0]:
            runs = list(last["runs"]) + [dict(r) for r in p0["runs"]]
            runs[len(last["runs"])]["t"] = " " + runs[len(last["runs"])]["t"].lstrip()
            body = [dict(x, runs=runs) if x is last else x for x in prev["body"]]
            prev["body"] = body + [x for x in it["body"] if x is not p0]
            continue
        out.append(it)
    return out


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
            # маркер шрифта Symbol/Wingdings в тексте — знаком частной области («\uf02d» — минус Symbol, OAH):
            # маркер или тире, из текста его снимает сборка (LEAD)
            m = re.match(r"\s*([•▪◦·–\-\uf020-\uf0ff])\s", text_of(p0["runs"])) if p0 and p0["runs"] else None
            dash = m and (m.group(1) in "–-" or "\uf020" <= m.group(1) <= "\uf0ff" and chr(ord(m.group(1)) - 0xF000) in "-–")
            if m and (it["kind"] in NUMBERED or it["kind"] == "Disc" and dash):
                it["kind"] = "Dash" if dash else "Disc"
            # и наоборот: маркированный, следом нумерованный — у нумерованных номер («8.», «e)») в тексте
            ml = LABEL.match(text_of(p0["runs"])) if p0 and p0["runs"] else None
            if it["kind"] not in NUMBERED and ml and (ml.group(1).isdigit() or ml.group(2) in (")", ".)")):
                it["kind"] = _label_kind(ml.group(1))
                if ml.group(2) in (")", ".)"):
                    it["suffix"] = ml.group(2)
            # своя метка в тексте, которую автонумерация не выразит («ba)», «(a)», «(5a)»; OAH T0.7, N3a), а метки Word
            # нет — пункт без маркера, метка остаётся текстом (был лишний «•» перед «ga)», 18.10)
            elif it["kind"] == "Disc" and not m and not it["lbl"] and p0 and p0["runs"] \
                    and OWN_LABEL.match(text_of(p0["runs"])):
                it["kind"] = "Label"
            if it["kind"] in NUMBERED:
                p = next((x for x in it["body"] if x["k"] == "p"), None)
                it["label"] = (_strip_label(p) if p else None) or \
                    (LABEL.match(it["lbl"]).group(1) if it["lbl"] and LABEL.match(it["lbl"]) else None)
            for x in it["body"]:
                if x["k"] == "table":
                    x["rows"] = [[dict(c, blocks=relist(c["blocks"])) for c in r] for r in x["rows"]]
        out += _nest(_continued(flat), 0, 0)[0]
    return out


def load(path):
    """Блоки размеченного PDF. Возвращает (блоки, сведения о листе: поля, колонтитулы)."""
    import pdfplumber
    from pdfplumber.structure import PDFStructTree
    pdf = pdfplumber.open(path)
    content, artifacts = {}, {"chars": [], "imgs": []}
    for p in pdf.pages:
        pn = p.page_number
        _underline(p.chars, p.rects + p.lines)
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
        # левый край строк после первой (висячий отступ: «MSZ EN 1593:2000 ⇥ Roncsolásmentes…», OAH 16.10)
        rest = [c for c in first if c["top"] > chs[0]["top"] + 0.6 * c["size"]]
        lt = max(c["top"] for c in first)                   # правый край последней строки (join_pages)
        return {"x0": min(c["x0"] for c in chs), "x1": max(c["x1"] for c in chs),
                "first_x0": chs[0]["x0"], "page": chs[0]["page_number"], "top": chs[0]["top"],
                "bottom": max(c["bottom"] for c in first), "lines": len({round(c["top"]) for c in first}),
                "last_x1": max(c["x1"] for c in first if c["top"] >= lt - 0.6 * c["size"]),
                "rest_x0": min(c["x0"] for c in rest) if rest else None,
                "rotated": sum(1 for c in chs if not _upright(c)) > len(chs) / 2,
                "size": Counter(round(c["size"]) for c in chs).most_common(1)[0][0]}

    def blank_geom(pn):
        """Место таблицы без текста (продолжение листа ознакомления ИН-1065: одни пустые графы) — по линиям листа между
        колонтитулами, если на листе нет текста тела. Без места она не склеивалась с началом на прошлом листе и шла
        отдельной таблицей с равными столбцами (19.10)."""
        page = pdf.pages[pn - 1]
        if any(c.get("mcid") is not None and c["text"].strip() for c in page.chars):
            return None
        H = float(page.height)
        ls = [o for o in page.rects + page.lines if 0.12 * H < float(o["top"]) and float(o["bottom"]) < 0.92 * H
              and (float(o["width"]) < 3 or float(o["height"]) < 3)]
        if sum(float(o["height"]) < 3 for o in ls) < 2 or sum(float(o["width"]) < 3 for o in ls) < 2:
            return None
        x0 = min(float(o["x0"]) for o in ls)
        return {"page": pn, "top": min(float(o["top"]) for o in ls), "bottom": max(float(o["bottom"]) for o in ls),
                "x0": x0, "x1": max(float(o["x1"]) for o in ls), "first_x0": x0, "size": 10}

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
        # пустые строки под текстом (графы листа ознакомления ИН-1065) — низ таблицы по её вертикалям, иначе линии
        # строк (высоты пустых строк) брались только до последнего текста
        allv = [o for o in page.rects + page.lines if float(o["width"]) < 3 and float(o["height"]) > 2
                and float(o["x0"]) >= x0 and float(o["x1"]) <= x1 and float(o["top"]) >= top]
        vbot = max((float(o["bottom"]) for o in vert), default=bot)
        while True:                                         # вертикали по строкам — цепочкой вниз
            nb = max((float(o["bottom"]) for o in allv if float(o["top"]) <= vbot + 2), default=vbot)
            if nb <= vbot:
                break
            vbot = nb
        if vbot > bot:
            objs += [o for o in page.rects + page.lines if bot < float(o["top"]) <= vbot + 2 and float(o["height"]) < 3
                     and float(o["x0"]) >= x0 and float(o["x1"]) <= x1]
        horiz = [o for o in objs if float(o["height"]) < 3 and float(o["width"]) > 20]
        xs = []
        for x in sorted((float(o["x0"]) + float(o["x1"])) / 2 for o in vert):
            if not xs or x - xs[-1] > 3:
                xs.append(x)
        # граница столбца — вертикаль хотя бы на половину высоты таблицы на листе: линия вложенной таблички в одной
        # строке («⚠ + текст» в «Описании» JPDS табл. 15) давала таблице лишний столбец, текст сжимался (19.10)
        span = max(1.0, min(float(g.get("bottom") or 0), bot) - max(float(g["top"]), top))
        cover = lambda x: sum(float(o["height"]) for o in vert if abs((float(o["x0"]) + float(o["x1"])) / 2 - x) <= 3)
        if len(xs) > 2:
            xs = [x for i, x in enumerate(xs) if i in (0, len(xs) - 1) or cover(x) >= 0.5 * span]
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
        out = {"xs": [round(x, 1) for x in xs], "ys": [round(y, 1) for y in ys]}   # линии строк — высота пустых строк
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
                # край текста — по самой левой строке: у многострочного текста по центру первая строка короче, и её
                # отступ шёл «полем» — 27,9 pt с каждой стороны, слова рвались («kötelezettsé g», N3a ABOS, 18.10)
                lx = [x for x in xs if x <= cg["x0"] + 0.5]
                rx = [x for x in xs if x >= cg["x1"] - 0.5]
                above = [y for y in ys if y <= cg["top"] + 0.5]
                below = [y for y in ys if y >= cg["bottom"] - 0.5]
                # поле слева — у ячеек с текстом влево: у текста по центру «поле» — половина свободного места (N3a:
                # 30 pt с каждой стороны, слова рвались посередине, 17.10)
                if lx and not (rx and abs((cg["x0"] - lx[-1]) - (rx[0] - cg["x1"])) < 4):
                    pads["left"].append(cg["x0"] - lx[-1])
                if above and below:
                    pads["top"].append(cg["top"] - above[-1]); pads["bottom"].append(below[0] - cg["bottom"])
        # больше 14 pt — отступ текста в ячейке (вложенный пункт, выравнивание), а не поле
        pads["left"] = [v for v in pads["left"] if v <= 14] or pads["left"][:0]
        for k, v in pads.items():
            if v:                                           # поле слева — нижняя четверть: отступ бывает и больше поля
                v.sort(); out["pad_" + k] = round(v[len(v) // 4 if k == "left" else len(v) // 2], 1)
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
                "keys": [re.sub(r"\s+", "", re.sub(r"\s*(\.{3,}\s*)?\d+\s*$", "", e["t"])) for e in body],
                # пункты с уровнем — текст заголовков, нарисованных картинками (toc_headings)
                "entries": [(re.sub(r"\s*(\.{3,}\s*)?\d+\s*$", "", e["t"]).strip(), lev[id(e)]) for e in body]}

    def own(el):
        """Абзац из собственных MCID контейнера (текст прямо в LBody/TD, без вложенного P)."""
        its = []
        for m in el.mcids or []:
            its += content.get((el.page_number, m), [])
        return [{"k": "p", "runs": _runs(its), "geom": geom(el)}] if its else []

    def container(el, depth):
        """Содержимое LBody/TD/TH: свой текст, затем вложенные блоки."""
        return own(el) + sum((block(x, depth + 1) for x in el.children), [])

    thin_of, fill_of = {}, {}

    def fills(pn):
        """Заливки листа (не линии, не белое, не во весь лист): [(x0, top, x1, bottom, цвет "RRGGBB")]."""
        if pn not in fill_of:
            pg = pdf.pages[pn - 1]
            area = float(pg.width) * float(pg.height)
            out = []
            for o in pg.rects:
                x0, x1, t0, t1 = float(o["x0"]), float(o["x1"]), float(o["top"]), float(o["bottom"])
                col = o.get("non_stroking_color")
                if not o.get("fill") or x1 - x0 < 5 or t1 - t0 < 5 or (x1 - x0) * (t1 - t0) > 0.3 * area:
                    continue
                v = [float(c) for c in ((col,) * 3 if isinstance(col, (int, float)) else
                                        tuple(col) * 3 if isinstance(col, (list, tuple)) and len(col) == 1 else
                                        col if isinstance(col, (list, tuple)) and len(col) == 3 else ())]
                if len(v) == 3 and min(v) < 0.97:
                    out.append((x0, t0, x1, t1, "%02X%02X%02X" % tuple(int(round(c * 255)) for c in v)))
            fill_of[pn] = out
        return fill_of[pn]

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
            rows, pgs = [], []
            for sec in el.children:
                trs = sec.children if sec.type in ("THead", "TBody", "TFoot") else [sec]
                for tr in trs:
                    if tr.type != "TR":
                        continue
                    row = []
                    for cell in tr.children:
                        if cell.type not in ("TD", "TH"):
                            continue
                        pgs.append(cell.page_number)
                        row.append({"blocks": container(cell, depth), "th": cell.type == "TH" or sec.type == "THead",
                                    "geom": geom(cell), "attrs": dict(cell.attributes or {})})
                    rows.append(row)
            cells = [c["geom"] for r in rows for c in r if c.get("geom")]
            g = None
            if cells:
                g = {"page": cells[0]["page"], "top": cells[0]["top"], "x0": min(c["x0"] for c in cells),
                     "x1": max(c["x1"] for c in cells), "first_x0": cells[0]["x0"], "size": cells[0]["size"],
                     "bottom": max(c.get("bottom", c["top"] + 12) for c in cells if c["page"] == cells[0]["page"])}
            elif pgs and pgs[0]:
                g = blank_geom(pgs[0])
            if g:
                cell_edges(rows, thin)
                cell_fills(rows, fills)
            return [{"k": "table", "rows": rows, "geom": g, "borders": borders(g), "frame": frame(g, rows) if g else {}}]
        if t in ("TOC", "TOCI"):
            if t != "TOC":
                return []
            info = toc_info(el)                 # лист — чтобы сборка видела конец титула (ÜFK: оглавление на 2-м листе)
            return [{"k": "toc", "info": info, "geom": {"page": info["page"]} if info.get("page") else None}]
        if t == "Note":                                     # сноска Word: текст внизу листа, ссылка — знаком в абзаце
            return [{"k": "note", "runs": _runs(items(el)), "geom": geom(el)}]
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
    blocks = inline_pictures(toc_headings(real_headings(adopt(evict(join_tocs(sub_rows(fit_rows(join_tables(figures_in_place(place_captions(unlist_captions(list_tails(relist(
        para_lists(join_pages(reflow(attach_notes(blocks))))))))))))))))))
    tops, bots = {}, {}
    for pg in pdf.pages:
        ys = [(float(c["top"]), float(c["bottom"])) for c in pg.chars if c.get("mcid") is not None and c["text"].strip()]
        ys += [(float(i["top"]), float(i["bottom"])) for i in pg.images if i.get("mcid") is not None]
        if ys:
            tops[pg.page_number], bots[pg.page_number] = min(y[0] for y in ys), max(y[1] for y in ys)
        else:
            # лист без текста тела (пустая таблица листа ознакомления, ИН-1065) — низ по линиям между колонтитулами,
            # иначе разрыв перед следующим листом не ставился (18.10)
            H = float(pg.height)
            ls = [float(o["bottom"]) for o in pg.rects + pg.lines if 0.12 * H < float(o["top"]) and float(o["bottom"]) < 0.92 * H]
            if ls:
                bots[pg.page_number] = max(ls)
    mark_breaks(blocks, tops, bots)
    page = pdf.pages[0]
    info = {"width": float(page.width), "height": float(page.height), "pages": len(pdf.pages),
            "artifacts": artifacts}
    return blocks, info


def _text_x(par):
    """Начало текста пункта из одной строки: первый прогон с буквой или цифрой, если он правее начала абзаца (маркера)
    на 3 pt и больше («•» ⇥ «Текст»); у пункта из нескольких строк место текста — по второй строке (rest_x0). Метка,
    снятая из текста (relist: «a.)»), оставляет прогону место метки — тогда места текста нет: следующий прогон брать
    нельзя (ÜFK I «a.) … 3,7 × 10⁶ Bq/dm³» — текст уезжал к верхнему индексу в конце строки, 20.10)."""
    g = par["geom"]
    if g.get("rest_x0") is not None or (g.get("lines") or 1) > 1:
        return None
    r = next((r for r in par.get("runs") or [] if re.search(r"\w", r.get("t") or "")), None)
    if r is None or r.get("x") is None or r["x"] < g["first_x0"] + 3 \
            or abs((r.get("top") or g.get("top") or 0) - (g.get("top") or 0)) >= 3:
        return None
    return r["x"]


def list_geo(blocks):
    """Место своего списка на листе: b["geo"] = {"marker": x метки, "text": x текста} (абсолютные, pt) — медиана по
    первым абзацам пунктов; текст — по строкам после первой (rest_x0). У ÜFK пункты «a.)» стоят и у поля, и с отступом
    30 pt: отступ вида списка (профиль) — медиана, и список у поля уезжал вправо (17.10). И в ячейках таблиц: там
    отступ по профилю замерен от поля листа, а не от края ячейки (ТЭ357, «•» почти все в таблицах)."""
    from statistics import median
    for b in blocks:
        if b["k"] == "table":
            for r in b["rows"]:
                for c in r:
                    list_geo(c["blocks"])
        if b["k"] == "list":
            heads = [it[0] for it in b["items"] if it and it[0]["k"] == "p" and it[0].get("geom")
                     and it[0]["geom"].get("first_x0") is not None]
            firsts = [x["geom"] for x in heads]
            texts = [g["rest_x0"] for g in firsts if g.get("rest_x0") and g["rest_x0"] > g["first_x0"] + 3]
            # пункт в одну строку — начало текста по первому прогону правее метки («•» ⇥ «Текст»): место списка из
            # однострочных пунктов брал профиль — медиана по документу, а список стоял иначе
            texts += [x for x in map(_text_x, heads) if x is not None]
            if firsts and texts:
                b["geo"] = {"marker": median(g["first_x0"] for g in firsts), "text": median(texts)}
            for it in b["items"]:
                list_geo(it)


HEAD_NUM = re.compile(r"^\s*(\d+(?:\.\d+)*)\.?(?=[\s\t])")


def real_headings(blocks):
    """Заголовки по смыслу, а не по стилю Word. Абзац тела с тегом заголовка (у ИН-1172 стиль обычного текста с уровнем
    структуры 2: 334 абзаца шли заголовками и в оглавление, 18.10) — абзацем: без номера, не полужирный и длинный (от
    80 знаков, от двух строк или с двоеточием в конце). Уровень — глубиной номера («1.1.» — 2-й), если уровни тегов с
    номерами расходятся у трети и больше заголовков (у ИН-1172 «1.1.» — H1, а «6.» — H3)."""
    hs = [b for b in blocks if b["k"] == "h"]
    for b in hs:
        t = text_of(b.get("runs") or []).strip()
        ink = [r for r in b.get("runs") or [] if "t" in r and r["t"].strip()]
        lines = (b.get("geom") or {}).get("lines") or 1
        # с номером — пункт текста («1.1. Настоящая инструкция предназначена …» на три строки, ИН-1065), если длинный
        long_ = (len(t) >= 150 or lines >= 3 or t.endswith(":")) if HEAD_NUM.match(t) else \
            (len(t) >= 80 or lines >= 2 or t.endswith(":"))
        if ink and not all(r.get("b") for r in ink) and long_:
            b["k"] = "p"
            b.pop("level", None)
    num = [(b, len(HEAD_NUM.match(text_of(b["runs"])).group(1).split("."))) for b in blocks
           if b["k"] == "h" and HEAD_NUM.match(text_of(b.get("runs") or []))]
    if num and sum(b["level"] != d for b, d in num) * 3 >= len(num):
        for b, d in num:
            b["level"] = min(d, 6)
    return blocks


def toc_headings(blocks):
    """Заголовок, нарисованный картинками (номер и текст — растры, текста в слое нет; JPDS: 71 заголовок), — текстом
    своего пункта оглавления: пункты идут по порядку заголовков, место — после пункта предыдущего заголовка с текстом,
    уровень должен совпасть. Раньше такой заголовок собирался пустым и пропадал вместе с картинками."""
    toc = next((b for b in blocks if b["k"] == "toc" and not b["info"].get("figures") and b["info"].get("entries")), None)
    if toc is None:
        return blocks
    ents = toc["info"]["entries"]
    keys = [re.sub(r"\s+", "", t) for t, _ in ents]
    cur = 0
    for b in blocks:
        if b["k"] != "h":
            continue
        t = text_of(b["runs"]).strip()
        if t:
            k = re.sub(r"\s+", "", t)
            j = next((j for j in range(cur, min(cur + 30, len(keys))) if keys[j] == k or keys[j].endswith(k)), None)
            if j is not None:
                cur = j + 1
        elif any("img" in r for r in b["runs"]) and cur < len(ents) and ents[cur][1] == b.get("level"):
            run = next((dict(r) for r in b["runs"] if "t" in r), {})
            b["drawn"] = [r["img"] for r in b["runs"] if "img" in r]     # растры заголовка — не «картинки без разметки»
            b["runs"] = [dict(run, t=ents[cur][0])]
            cur += 1
    return blocks


CAPTION_ANY = re.compile(r"^\s*(?:(?:Рис(?:унок|\.)|Таблица|Figure|Fig\.|Table|Abbildung|Abb\.|Tabelle)\s*\d"
                         r"|\d+\.\s*(?:ábra|táblázat))", re.I)


def list_tails(bs):
    """Хвост последнего пункта (рисунок, подпись), который на листе ниже абзаца, идущего за списком: Word разметил их
    в пункте, абзац между ними — после списка (JPDS л. 42: «– Файлы.», «Соответствие данных…:», рисунок 37 — фраза
    печаталась под подписью, 17.10), — за этот абзац, по месту на листе."""
    out = []
    for i, b in enumerate(bs):
        if b["k"] == "list":
            for it in b["items"]:
                it[:] = list_tails(it)
        out.append(b)
    i = 0
    while i < len(out):
        b = out[i]
        i += 1
        if b["k"] != "list" or not b["items"] or i >= len(out):
            continue
        last, n = b["items"][-1], out[i]
        gn = n.get("geom") or {}
        if n["k"] not in ("p", "h") or not gn.get("page") or not gn.get("bottom"):
            continue
        def geoms(x):                                       # места блока; вложенный список (подпись-пункт) — его блоков
            if x["k"] == "list":
                return [g for it in x["items"] for y in it for g in geoms(y)] or [{}]
            return [x.get("geom") or {}]
        # хвост — ниже абзаца на его листе или дальше (подпись следующего рисунка, вложенная в подпись-пункт)
        after = lambda x: all(g.get("page") and g.get("bottom") and (g["page"] > gn["page"] or g["page"] == gn["page"]
                                                                       and g["top"] >= gn["bottom"] - 2) for g in geoms(x))
        s = next((k for k in range(1, len(last)) if all(after(x) for x in last[k:])), None)
        if s is None or any(g.get("page") != gn["page"] or g.get("bottom", 1e9) > gn["top"] + 2
                            for g in geoms(last[s - 1])) or geoms(last[s])[0].get("page") != gn["page"]:
            continue
        top = min(g["top"] for g in geoms(last[s]) if g["page"] == gn["page"])
        j = i                                               # абзацы за списком, что выше хвоста
        while j < len(out) and out[j]["k"] in ("p", "h") and (out[j].get("geom") or {}).get("page") == gn["page"] \
                and (out[j]["geom"].get("bottom") or 1e9) <= top + 2:
            j += 1
        tail = last[s:]
        del last[s:]
        out[j:j] = tail
    return out


def unlist_captions(bs):
    """Подпись рисунка или таблицы, размеченная пунктом списка (JPDS: «Рисунок 42 – …» — пункт вложенного списка, в DOCX
    «a. Рисунок 42», «i. Рисунок 43», 17.10), — не пункт: абзацы пункта — на своё место, список рвётся вокруг них,
    нумерация продолжается."""
    out = []
    for b in bs:
        if b["k"] != "list":
            out.append(b)
            continue
        cur, n = [], 0
        start = b.get("start", 1)
        for it in b["items"]:
            it = unlist_captions(it)
            # пункт — подпись, перед ней разве что картинки (рисунок с подписью в пункте: «а. Рисунок 25», JPDS)
            first = next((x for x in it if x["k"] != "p" or text_of(x.get("runs") or []).strip()), None)
            pics = it[:it.index(first)] if first is not None else []
            if first is not None and first["k"] == "p" and CAPTION_ANY.match(text_of(first.get("runs") or []))                     and all(x["k"] == "p" and all("img" in r for r in x.get("runs") or []) for x in pics):
                if cur:
                    out.append(dict(b, items=cur, start=start + n - len(cur)))
                    cur = []
                out += it
            else:
                cur.append(it)
                n += 1
        if cur:
            out.append(dict(b, items=cur, start=start + n - len(cur)))
    return out


def inline_pictures(bs):
    """Значок в строке абзаца, который Word разметил отдельным рисунком после абзаца (клавиша ESC в ИН-1065: в тексте
    на его месте табуляция с позицией у края значка, значок шёл отдельным абзацем под текстом, 17.10), — в абзац на
    место этой табуляции: прогон {"img", "at_tab"}; сборка ставит его на первую табуляцию текста (нет её в переводе —
    в конец абзаца)."""
    out = []
    for b in bs:
        for it in b.get("items") or []:
            it[:] = inline_pictures(it)
        prev = out[-1] if out else None
        g, pg = b.get("geom") or {}, (prev or {}).get("geom") or {}
        runs = b.get("runs") or []
        if prev is not None and prev["k"] == "p" and b["k"] == "p" and len(runs) == 1 and "img" in runs[0]                 and g.get("page") and g.get("page") == pg.get("page") and pg.get("bottom") and g.get("bottom")                 and g["top"] >= pg["top"] - 2 and g["bottom"] <= pg["bottom"] + 2:
            im = runs[0]["img"]
            k = next((i for i, r in enumerate(prev["runs"]) if "	" in r.get("t", "") and any(
                min(abs(x - float(im["x1"])), abs(x - float(im["x0"]))) < 8 for x in r.get("tabs") or [])), None)
            if k is not None:
                r = dict(prev["runs"][k])
                r.pop("tabs", None)
                prev["runs"][k] = r
                # табуляция после маркера пункта («•⇥[значок] : …», UTS) — значок в начале текста пункта, не на
                # табуляции: маркер с табуляцией сборка снимает, и значок уезжал в конец пункта (18.10)
                lead = not re.search(r"\w", text_of(prev["runs"][:k]) + r.get("t", "").split("\t")[0])
                prev["runs"].insert(k + 1, {"img": im} if lead else {"img": im, "at_tab": True})
                continue
        out.append(b)
    return out


def place_captions(blocks):
    """Подпись рисунка, стоящая в потоке раньше своей картинки, а на листе — прямо под ней (до 30 pt): Word разметил её
    в предыдущем пункте списка (JPDS: «Рисунок 38» печаталась над меню, картинка — следующим листом, 17.10), — сразу
    за картинку. Пустые после этого пункты списков — прочь."""
    seq = []                                                # (контейнер, блок) в порядке чтения

    def walk(bs):
        for b in bs:
            seq.append((bs, b))
            for it in b.get("items") or []:
                walk(it)
    walk(blocks)
    order = {id(b): k for k, (_, b) in enumerate(seq)}
    pics = [(c, b) for c, b in seq if b["k"] == "p" and (b.get("geom") or {}).get("bottom")
            and any("img" in r for r in b.get("runs") or []) and len(text_of(b.get("runs") or []).strip()) <= 40]
    moved = False
    for cont, b in seq:
        g = b.get("geom") or {}
        if b["k"] != "p" or not g.get("page") or not FIG_CAPTION.match(text_of(b.get("runs") or [])):
            continue
        cx = (g.get("x0", 0) + g.get("x1", 0)) / 2
        host = next(((pc, pb) for pc, pb in pics if pb["geom"].get("page") == g["page"]
                     and 0 <= g["top"] - pb["geom"]["bottom"] <= 30
                     and pb["geom"].get("x0", 0) - 50 <= cx <= pb["geom"].get("x1", 0) + 50
                     and order[id(pb)] > order[id(b)]), None)
        if host is None:
            continue
        cont.remove(b)
        host[0].insert(host[0].index(host[1]) + 1, b)
        moved = True
    if moved:
        def prune(bs):
            for b in bs:
                if b.get("items") is not None:
                    b["items"] = [it for it in b["items"] if it]
                    for it in b["items"]:
                        prune(it)
        prune(blocks)
        blocks[:] = [b for b in blocks if b["k"] != "list" or b["items"]]
    return blocks


def figures_in_place(bs):
    """Рисунки и их подписи — по месту на листе среди текста. Word кладёт цепочку «рисунок, подпись» вложенными пунктами
    одного пункта списка, текст между ними — после списка (JPDS л. 109–111: три рисунка с подписями шли пачкой перед
    абзацами, что на листе выше них; л. 51 — рисунок раньше «Для переноса…», 17.10). Слияние двух последовательностей:
    текст (и таблицы) — в своём порядке, рисунок и подпись тоже; рисунок раньше текста, если целиком выше него на листе,
    позже — если целиком ниже; рядом по высоте (значок в строке, обтекание) — как в разметке."""
    for b in bs:
        for it in b.get("items") or []:
            it[:] = figures_in_place(it)
    # рисунок — картинка, разве что с короткой надписью (выноски «а 1» составного рисунка ТЭ357 — его же части)
    fig = lambda b: b["k"] == "p" and ((any("img" in r for r in b.get("runs") or [])
                                        and len(text_of(b.get("runs") or []).strip()) <= 40)
                                       or FIG_CAPTION.match(text_of(b.get("runs") or [])))
    placed = lambda b: (b.get("geom") or {}).get("page") and (b.get("geom") or {}).get("bottom") is not None
    # пустой абзац без места на листе участок не рвёт — идёт со своим предыдущим блоком
    empty = lambda b: b["k"] == "p" and not placed(b) and not text_of(b.get("runs") or []).strip() \
        and not any("img" in r for r in b.get("runs") or [])
    out, i = [], 0
    while i < len(bs):                                      # участки подряд идущих блоков с местом на листе
        if bs[i]["k"] not in ("p", "h", "table") or not placed(bs[i]):
            out.append(bs[i]); i += 1
            continue
        units = []
        while i < len(bs) and bs[i]["k"] in ("p", "h", "table") and placed(bs[i]):
            units.append([bs[i]]); i += 1
            while i < len(bs) and empty(bs[i]):
                units[-1].append(bs[i]); i += 1
        T = [u for u in units if not fig(u[0])]
        order = {id(u[0]): k for k, u in enumerate(units)}

        def above(u, v):                                    # u целиком выше v на листе (или на листе раньше)
            gu, gv = u[0]["geom"], v[0]["geom"]
            return gu["page"] < gv["page"] or gu["page"] == gv["page"] and gu["bottom"] <= gv["top"] + 2

        def follows(t, f):                                  # текст t — после рисунка f
            if above(f, t):
                return True
            return not above(t, f) and t[0]["geom"]["page"] == f[0]["geom"]["page"] and order[id(f[0])] < order[id(t[0])]
        slots = [[] for _ in range(len(T) + 1)]             # рисунки перед T[k]
        for u in units:
            if fig(u[0]):
                g = slots[next((k for k, t in enumerate(T) if follows(t, u)), len(T))]
                j = len(g)                                  # в своём промежутке — выше того, что целиком под ним
                while j and above(u, g[j - 1]):
                    j -= 1
                g.insert(j, u)
        for k in range(len(T) + 1):
            for u in slots[k]:
                out += u
            if k < len(T):
                out += T[k]
    return out


FIG_CAPTION = re.compile(r"^\s*(?:(?:Рис(?:унок|\.)|Figure|Fig\.|Abbildung|Abb\.)\s*\d|\d+\.\s*ábra)", re.I)


def compose_figures(blocks, path):
    """Составной рисунок: скриншот и фигуры Word поверх и рядом — рамка, стрелки, круги и рамки выносок (векторы),
    значки (картинки), надписи выносок (короткий текст). Порознь рисунок терял рамку и стрелки, а надписи рассыпались
    абзацами под ним (ТЭ357, рис. 35–36, 17.10). Рисунок прирастает всем, что его касается: векторами (кроме линий
    таблиц и рамок листа) и короткими (до 40 знаков, не подпись «Рисунок N») абзацами, — и идёт одной картинкой
    области листа: run["img"]["clip"]. Сборка рисует её с переведённого PDF (надписи выносок — переводом), без него —
    с оригинала. Область, задевшая обычный текст, — не рисунок: тогда всё как было."""
    import fitz
    with fitz.open(str(path)) as doc:
        return _compose_figures(blocks, doc)


def _compose_figures(blocks, doc):
    import fitz
    img_only = lambda b: b["k"] == "p" and b.get("runs") and all("img" in r for r in b["runs"])
    box = lambda g: fitz.Rect(g["x0"], g["top"], g["x1"], g["bottom"])
    by_page = {}
    for i, b in enumerate(blocks):
        g = b.get("geom") or {}
        if g.get("page") and g.get("x0") is not None and g.get("x1") is not None and g.get("bottom"):
            by_page.setdefault(g["page"], []).append(i)
    for b in blocks:                                        # сначала картинки ячеек — им свои картинки листа
        if b["k"] == "table":
            _cell_pictures(b, doc, blocks)
    drop, new = set(), {}                                   # new: место в потоке → рисунки не из разметки
    rep = Counter(x for pg in doc for x in {im[0] for im in pg.get_images()})   # картинка на многих листах — логотип
    for pn, idx in by_page.items():
        page = doc[pn - 1]
        W, H = page.rect.width, page.rect.height
        tables = [box(blocks[i]["geom"]) for i in idx if blocks[i]["k"] == "table"]
        draws = []                                          # (рамка, число элементов пути)
        for d in page.get_drawings():
            r = d["rect"]
            white = d.get("fill") and min(d["fill"]) > 0.95 and not d.get("color")
            if r.width <= 1.5 or r.height <= 1.5 or r.width > 0.8 * W or r.height > 0.8 * H or white                     or any(r.intersects(t) for t in tables):
                continue                                    # линии таблиц, рамка листа, белое на белом
            draws.append((r, len(d["items"])))
        # тонкие линии — только указатели выносок: начинаются на рисунке (проверка в grow), не во весь лист, не таблица
        # (ТЭ357 рис. 37: линия от снимка к рамке с мышкой толщиной 1 pt отбрасывалась как линейка, 19.10)
        leads = [d["rect"] + (-0.5, -0.5, 0.5, 0.5) for d in page.get_drawings()
                 if (d["rect"].width <= 1.5) != (d["rect"].height <= 1.5) and max(d["rect"].width, d["rect"].height) >= 10
                 and d["rect"].width < 0.6 * W and d["rect"].height < 0.6 * H and not any(d["rect"].intersects(t) for t in tables)]
        if not draws and not any(img_only(blocks[i]) for i in idx):
            continue
        def big_img(b):
            return any("img" in r and r["img"].get("x0") is not None
                       and (float(r["img"]["x1"]) - float(r["img"]["x0"])) >= 60
                       and (float(r["img"]["bottom"]) - float(r["img"]["top"])) >= 40 for r in b.get("runs") or [])
        # надпись или значок поверх рисунка: короткий текст, картинка — мелкая (мышки с номерами ТЭ357); абзац с
        # крупной картинкой — свой рисунок, не надпись (скриншот JPDS л. 96)
        sats = [j for j in idx if blocks[j]["k"] == "p" and not big_img(blocks[j])
                and len(text_of(blocks[j]["runs"]).strip()) <= 40 and not FIG_CAPTION.match(text_of(blocks[j]["runs"]))]

        def grow(R, me, near=2):
            """Рисунок прирастает касающимися векторами и короткими абзацами; (область, поглощённые, были ли
            векторы) — или None, если область задела обычный текст хоть строкой (абзац, пункт списка, заголовок,
            таблицу; ревью 17.10)."""
            got, shapes, changed = set(), False, True
            while changed:
                changed = False
                for r, _ in draws:
                    # 6 pt: линия-указатель кончается в 4 pt от рамки выноски (ТЭ357 рис. 37–38: мышки выносок
                    # оставались справа от снимка отдельными картинками, 19.10)
                    if r.intersects(R + (-6, -6, 6, 6)):
                        shapes = True
                        if not R.contains(r):
                            R |= r
                            changed = True
                for r in leads:                             # указатель, начатый внутри области
                    if not R.contains(r) and (R.contains(r.tl) or R.contains(r.br)):
                        before = fitz.Rect(R)
                        R |= r
                        changed = changed or R != before        # пустая рамка объединением не меняет — без вечного цикла
                for j in sats:
                    if j not in got and j not in drop and box(blocks[j]["geom"]).intersects(R + (-near, -near, near, near)):
                        R |= box(blocks[j]["geom"])
                        got.add(j)
                        changed = True
            others = [(box(x["geom"]), bool(FIG_CAPTION.match(text_of(x.get("runs") or []))))
                      for j, y in enumerate(blocks) if j != me and j not in got and j not in drop
                      for x in ([y] if y["k"] in ("h", "table") else _paras([y]))
                      if (x.get("geom") or {}).get("page") == pn and x["geom"].get("bottom")
                      and x["geom"].get("x1") is not None and not img_only(x)]
            # подпись рисунка, задетая низом области на несколько pt (буквы выносок «а б в» над «Рисунок 40», ТЭ357), —
            # не помеха: область обрезается по её верху
            for q, cap in others:
                if cap and min(q.x1, R.x1) - max(q.x0, R.x0) > 2 and 0 < R.y1 - q.y0 <= 6 and q.y0 > R.y0 + 0.5 * R.height:
                    R.y1 = q.y0
            if any(min(q.y1, R.y1) - max(q.y0, R.y0) > 2 and min(q.x1, R.x1) - max(q.x0, R.x0) > 2 for q, _ in others):
                return None
            return R & page.rect, got, shapes

        clip = lambda R: [{"img": {"clip": [R.x0, R.y0, R.x1, R.y1], "page_number": pn, "x0": R.x0, "x1": R.x1,
                                   "top": R.y0, "bottom": R.y1}}]
        done = []
        def pictured(b):
            """Картинка от 60×40 pt и текста до 40 знаков (номера выносок «1 23 4 5 67» в том же абзаце, JPDS л. 96:
            печатались строкой под скриншотом) — рисунок."""
            return b["k"] == "p" and len(text_of(b.get("runs") or []).strip()) <= 40 and big_img(b)
        for i in idx:                                       # рисунок-картинка с фигурами и надписями поверх
            b = blocks[i]
            if i in drop or not (img_only(b) or pictured(b)):
                continue
            R = box(b["geom"])
            for r in b["runs"]:
                if "img" in r and r["img"].get("x0") is not None and (r["img"].get("page_number") or pn) == pn:
                    R |= fitz.Rect(float(r["img"]["x0"]), float(r["img"]["top"]), float(r["img"]["x1"]),
                                   float(r["img"]["bottom"]))
            if R.width < 60 or R.height < 40:
                # мелкая картинка — рисунок, только если к ней сходятся выноски: фигуры-стрелки и хотя бы две надписи,
                # и вместе это не меньше 60×40 (индикатор 114×39 с номерами 1–6, ТЭ357 рис. 11–12: номера рассыпались
                # строками под картинкой, 19.10)
                res = grow(fitz.Rect(R), i) if R.width >= 30 and R.height >= 12 else None
                if res is None or not res[2] or len(res[1]) < 2 or res[0].width < 60 or res[0].height < 40:
                    done.append(R)
                    continue
                R, got, _ = res
                b["runs"] = clip(R)
                b["geom"] = dict(b["geom"], x0=R.x0, x1=R.x1, top=R.y0, bottom=R.y1)
                drop |= got - {i}                           # сама мелкая картинка — среди «надписей», её не выкидывать
                done.append(R)
                continue
            res = grow(fitz.Rect(R), i)
            if res is None or img_only(b) and not res[2] and not res[1]:
                done.append(R)
                continue
            R, got, _ = res
            b["runs"] = clip(R)
            b["geom"] = dict(b["geom"], x0=R.x0, x1=R.x1, top=R.y0, bottom=R.y1)
            drop |= got
            done.append(R)
        # картинка листа, которой нет в разметке (скриншот в группе фигур Word — Form XObject, pdfplumber его не
        # видит: JPDS л. 96 — оставалась одна подпись «Рисунок 88»), вне таблиц и колонтитулов — рисунком на своём месте
        shown = _img_rects(blocks, pn)                      # и блоки, начатые на другом листе
        done += shown                                       # и картинки ячеек: их фигуры — уже при деле
        tops = [blocks[i]["geom"]["top"] for i in idx]
        bots = [blocks[i]["geom"]["bottom"] for i in idx]
        area = W * H
        for im in page.get_image_info(xrefs=True):
            r = fitz.Rect(im["bbox"]) & page.rect
            cy = (r.y0 + r.y1) / 2
            if r.get_area() < 400 or r.get_area() > 0.5 * area or (im["xref"] and rep[im["xref"]] > 1)                     or not min(tops) - 20 <= cy <= max(bots) + 20 or any(r.intersects(t) for t in tables)                     or any((r & q).get_area() > 0.5 * r.get_area() for q in shown + done):
                continue
            res = grow(fitz.Rect(r), None)
            if res is None:
                continue
            R, got, _ = res
            at = next((i for i in idx if i not in drop and blocks[i]["geom"]["top"] >= R.y0 - 1), max(idx) + 1)
            new.setdefault(at, []).append({"k": "p", "runs": clip(R), "figure": True, "loose": True, "geom": {
                "page": pn, "top": R.y0, "bottom": R.y1, "x0": R.x0, "x1": R.x1, "first_x0": R.x0, "size": 10}})
            drop |= got
            done.append(R)
        # рисунок из одних векторов (график Word: области, рамки, подписи осей абзацами — ÜFK I, л. 48): скопление
        # фигур не у картинки — от трёх фигур или путь от 8 элементов, от 80×60 pt
        free = [(r, n) for r, n in draws if not any(r.intersects(q) for q in done)]
        groups = []
        for r, n in free:
            hit = [g for g in groups if g[0].intersects(r + (-12, -12, 12, 12))]   # узор из точек — одно скопление
            for g in hit:
                groups.remove(g)
            R = fitz.Rect(r)
            cnt, most = 1, n
            for g in hit:
                R |= g[0]
                cnt, most = cnt + g[1], max(most, g[2])
            groups.append((R, cnt, most))
        pics = [box(blocks[i]["geom"]) for i in idx if img_only(blocks[i])] + done
        for R, cnt, most in groups:
            # выноски поверх картинки (JPDS л. 96: номера в кружках на скриншоте) — дело картинки, не свой рисунок
            if R.width < 80 or R.height < 60 or cnt < 3 and most < 8 or any(R.intersects(q) for q in pics):
                continue
            # подписи осей стоят в нескольких pt от области графика
            res = grow(fitz.Rect(R), None, near=6)
            # без подписей и из немногих фигур — оформление (рамка, плашка); от 20 фигур — рисунок и без подписей
            # (узор из точек на обложке ONS, 184 фигуры, пропадал)
            if res is None or not res[1] and cnt < 20:
                continue
            R, got, _ = res
            at = min(got) if got else next((i for i in idx if i not in drop and blocks[i]["geom"]["top"] >= R.y0 - 1),
                                           max(idx) + 1)
            new.setdefault(at, []).append({"k": "p", "runs": clip(R), "figure": True, "vector": True, "geom": {
                "page": pn, "top": R.y0, "bottom": R.y1, "x0": R.x0, "x1": R.x1, "first_x0": R.x0,
                "size": (blocks[at].get("geom") or {}).get("size") or 10 if at < len(blocks) else 10}})
            drop |= got
    out = []
    for i, b in enumerate(blocks):
        out += new.get(i, [])
        if i not in drop:
            out.append(b)
    out += new.get(len(blocks), [])
    return out


def _img_rects(bs, pn):
    """Рамки картинок разметки листа pn во всех блоках (и в ячейках, пунктах)."""
    import fitz
    out = []
    for b in bs:
        for r in (b.get("runs") or []) + [{"img": x} for x in b.get("drawn") or []]:
            im = r.get("img")
            if im and (im.get("page_number") or pn) == pn and im.get("x0") is not None:
                out.append(fitz.Rect(float(im["x0"]), float(im["top"]), float(im["x1"]), float(im["bottom"])))
        for it in b.get("items") or []:
            out += _img_rects(it, pn)
        for row in b.get("rows") or []:
            for c in row:
                out += _img_rects(c["blocks"], pn)
    return out


def _cell_pictures(t, doc, blocks):
    """Картинка ячейки, которой нет в разметке: элементы интерфейса, собранные в Word из фигур и надписей (флажки,
    вкладки, полоса прокрутки, JPDS «Условные обозначения»), лежат в группе (Form XObject) — pdfplumber их не видит,
    в ячейке оставались одни подписи «Вкладка 2Вкладка 3Вкладка 1» (17.10). Картинки листа (PyMuPDF), которых нет среди
    картинок разметки, по полосе строки и столбцу ячейки — с касающимися их векторами и короткими надписями ячейки —
    одной картинкой области листа (img["clip"], как compose_figures). Задевает длинный текст ячейки — как было."""
    import fitz
    rows = t["rows"]
    span = []                                               # (лист, верх, низ) содержимого строки
    for r in rows:
        gs = [c["geom"] for c in r if c.get("geom") and c["geom"].get("bottom")]
        span.append((gs[0]["page"], min(g["top"] for g in gs), max(g["bottom"] for g in gs)) if gs else None)
    cols = {}                                               # число ячеек строки → [(лево, право)] столбцов
    for r in rows:
        for k, c in enumerate(r):
            g = c.get("geom")
            if g and g.get("x1") is not None:
                lo, hi = cols.setdefault(len(r), {}).get(k, (g["x0"], g["x1"]))
                cols[len(r)][k] = (min(lo, g["x0"]), max(hi, g["x1"]))
    pages, rep = {}, None
    for i, r in enumerate(rows):
        if span[i] is None or len(r) < 2:
            continue
        pn, top, bot = span[i]
        if pn not in pages:
            page = doc[pn - 1]
            shown = _img_rects(blocks, pn)
            if rep is None:                                 # картинка на нескольких листах — логотип, водяной знак
                rep = Counter(x for pg in doc for x in {im[0] for im in pg.get_images()})
            area = page.rect.get_area()
            imgs = [fitz.Rect(im["bbox"]) for im in page.get_image_info(xrefs=True)
                    if not (im["xref"] and rep[im["xref"]] > 1) and fitz.Rect(im["bbox"]).get_area() < 0.3 * area
                    and not any((fitz.Rect(im["bbox"]) & q).get_area() > 0.5 * fitz.Rect(im["bbox"]).get_area()
                                for q in shown)]
            dr = page.get_drawings()
            draws = [d["rect"] for d in dr if d["rect"].width > 1.5 and d["rect"].height > 1.5
                     and not (d.get("fill") and min(d["fill"]) > 0.95 and not d.get("color"))]
            # заливка цветом без текста в ячейке (легенда статусов JPDS: цветные ячейки пропадали) — тоже содержимое
            fills = [d["rect"] for d in dr if "f" in (d.get("type") or "") and d.get("fill") and min(d["fill"]) < 0.95
                     and d["rect"].width > 3 and d["rect"].height > 3 and d["rect"].get_area() < 0.3 * area]
            pages[pn] = (page, imgs, draws, fills)
        page, imgs, draws, fills = pages[pn]
        if not imgs and not fills:
            continue
        H = page.rect.height
        prev = next((span[j] for j in range(i - 1, -1, -1) if span[j] and span[j][0] == pn), None)
        nxt = next((span[j] for j in range(i + 1, len(rows)) if span[j] and span[j][0] == pn), None)
        # первая и последняя строки таблицы на листе — не до краёв листа (шапка, подвал, фигура над таблицей)
        y0 = (prev[2] + top) / 2 if prev else max(top - 40, 0.05 * H)
        y1 = (bot + nxt[1]) / 2 if nxt else min(bot + 60, 0.95 * H)
        cs = dict(cols.get(len(r), {}))
        if not cs:
            continue
        for k in range(len(r)):     # столбец без текста ни в одной строке (цветные ячейки легенды JPDS) — между соседями
            if k not in cs:
                lo = cs[k - 1][1] + 1 if k - 1 in cs else 0
                hi = next((cs[q][0] for q in range(k + 1, len(r)) if q in cs), page.rect.width) - 1
                cs[k] = (lo, max(lo, hi))
        bounds = [(cs[k][1] + cs[k + 1][0]) / 2 for k in range(len(r) - 1)]
        for k, c in enumerate(r):
            x0 = bounds[k - 1] if k else 0
            x1 = bounds[k] if k < len(bounds) else page.rect.width
            mine = [q for q in imgs if x0 <= (q.x0 + q.x1) / 2 < x1 and y0 <= (q.y0 + q.y1) / 2 < y1]
            tb = [fitz.Rect(g["x0"], g["top"], g["x1"], g["bottom"]) for g in
                  (x.get("geom") for x in c["blocks"]) if g and g.get("bottom") and g.get("page") == pn]
            if not mine and not tb:                         # пустая ячейка с заливкой — её цвет
                mine = [q for q in fills if x0 <= (q.x0 + q.x1) / 2 < x1 and y0 <= (q.y0 + q.y1) / 2 < y1]
            if not mine:
                continue
            G = fitz.Rect(mine[0])
            for q in mine[1:]:
                G |= q
            cb = [x for x in c["blocks"] if x.get("geom") and x["geom"].get("bottom") and x["geom"].get("page") == pn]
            got, changed, pic0 = [], True, None
            while changed:
                changed = False
                for q in draws:     # в своей полосе строки: заливки соседних строк легенды JPDS касаются друг друга —
                    # и каждая ячейка росла на весь столбец, 9 копий столбца по листу (+9 листов, 17.10)
                    if q.intersects(G + (-2, -2, 2, 2)) and not G.contains(q) and x0 <= (q.x0 + q.x1) / 2 < x1 \
                            and y0 <= (q.y0 + q.y1) / 2 < y1:
                        G |= q
                        changed = True
                if pic0 is None:
                    pic0 = fitz.Rect(G)                     # сама графика — до надписей
                for x in cb:
                    g = x["geom"]
                    # подпись рядом со значком — в 2–4 pt справа («☐ Отключено»)
                    if x not in got and fitz.Rect(g["x0"], g["top"], g["x1"], g["bottom"]).intersects(G + (-4, -2, 4, 2)):
                        got.append(x)
                        G |= fitz.Rect(g["x0"], g["top"], g["x1"], g["bottom"])
                        changed = True
            # надпись — короткая или целиком на самой графике (подписи вкладок «Вкладка 1 … Вкладка 3» — 80 знаков);
            # задет обычный текст ячейки — не картинка
            inside = lambda g: (fitz.Rect(g["x0"], g["top"], g["x1"], g["bottom"]) & pic0).get_area() >= 0.9 * max(
                fitz.Rect(g["x0"], g["top"], g["x1"], g["bottom"]).get_area(), 1)
            if any(x["k"] != "p" or len(text_of(x.get("runs") or []).strip()) > 40 and not inside(x["geom"]) for x in got) \
                    or pic0.x0 < x0 - 5 or pic0.x1 > x1 + 5:  # графика шире своей ячейки — не картинка ячейки
                continue
            G &= page.rect
            pic = {"k": "p", "runs": [{"img": {"clip": [G.x0, G.y0, G.x1, G.y1], "page_number": pn, "x0": G.x0,
                                               "x1": G.x1, "top": G.y0, "bottom": G.y1}}],
                   "geom": {"page": pn, "top": G.y0, "bottom": G.y1, "x0": G.x0, "x1": G.x1, "first_x0": G.x0,
                            "size": next((x["geom"]["size"] for x in got + cb if x["geom"].get("size")), 10)}}
            at = min((c["blocks"].index(x) for x in got), default=0)
            c["blocks"] = [x for x in c["blocks"] if x not in got]
            c["blocks"].insert(min(at, len(c["blocks"])), pic)
            imgs[:] = [q for q in imgs if q not in mine]


def mark_breaks(blocks, tops, bots):
    """Принудительный разрыв листа (b["break"]): заголовок или абзац первым на листе, а прошлый лист кончился заметно
    (больше дюйма) выше обычного низа содержимого. tops/bots — {лист от 1: верх/низ содержимого}. Разрыв стилем — лишь
    когда с нового листа все заголовки уровня; у ÜFK подразделы 4.4.1.x — каждый с нового листа, кроме идущего сразу под
    разделом, — и разрывы пропадали все: 157 листов собирались в 128 (16.10). Общее для PDF с разметкой и без
    (docx_layout, текстовый слой: у Eskom «Content» и «1. Introduction» шли следом за титульными сведениями, 17.10)."""
    # низ листа — обычный (80-й процентиль по листам), не самый низкий: текст у края одного листа (UTS ~790 pt при
    # обычных ~720) делал «пустым» всякий лист и ставил ложные разрывы
    bl_ = sorted(bots.values())
    low = bl_[len(bl_) * 4 // 5] if bl_ else 0
    img_only = lambda b: b["k"] == "p" and all("img" in r for r in b.get("runs") or [{"img": 1}])

    def first_row_bottom(t):
        """Низ первой строки таблицы на её листе: Word переносит таблицу по строкам, и с листа её уводит только
        неуместившаяся первая строка, а не вся таблица (лист согласования ИН-1065, 18.10)."""
        pn = t["geom"]["page"]
        for r in t.get("rows") or []:
            bs = [c["geom"]["bottom"] for c in r if (c.get("geom") or {}).get("page") == pn and c["geom"].get("bottom")]
            if bs:
                return max(bs)
        return t["geom"].get("bottom") or 0
    for i, b in enumerate(blocks):
        g = b.get("geom") or {}
        pn = g.get("page")
        if b["k"] == "toc":
            # оглавление первым на листе, а прошлый лист кончился заметно выше обычного низа, — с нового листа
            # («TARTALOMJEGYZÉK» OAH AKF1.32: лист 5 занят на четверть, а оглавление шло следом; E41 корпуса, 22.10)
            if pn and pn - 1 in bots and low - bots[pn - 1] > 72                     and not any((x.get("geom") or {}).get("page") == pn for x in blocks[:i]):
                b["break"] = True
            continue
        # и абзац («FŐIGAZGATÓI ELŐSZÓ» OAH — абзац, не заголовок), и таблица («УТВЕРЖДАЮ» листа изменений ИН-1065);
        # картинка — нет: её сдвигает нехватка места
        if not ((b["k"] in ("h", "table") or b["k"] == "p" and not img_only(b))
                and pn and pn - 1 in bots and pn in tops and abs(g["top"] - tops[pn]) < 2):
            continue
        gap = low - bots[pn - 1]
        # перенос сам собой: рисунок или таблица в начале листа вместе с абзацами перед ними (Word держит их вместе)
        # не влезали в остаток прошлого листа — не разрыв (UTS: «см. рисунок» + снимок экрана, 60 ложных разрывов)
        chain = next((x for x in blocks[i:i + 4] if (x.get("geom") or {}).get("page") == pn
                      and (x["k"] == "table" or img_only(x))), None)
        if chain is not None and ((first_row_bottom(chain) if chain["k"] == "table" else chain["geom"].get("bottom"))
                                  or 0) - tops[pn] > gap:
            continue
        if gap > 72:
            b["break"] = True
    # абзац, оформленный как заголовки первого уровня, идущие с нового листа (кегль и жирность те же, стоит первым на
    # листе на их высоте, коротко, без точки в конце): «ELŐSZÓ» OAH AKF1.32 — абзац, не заголовок, прошлый лист полон,
    # и по просвету разрыв не находился (E41 корпуса, 22.10)
    look = lambda b: next(((r.get("size"), bool(r.get("b"))) for r in b.get("runs") or [] if (r.get("t") or "").strip()),
                          None)
    h1 = [b for b in blocks if b["k"] == "h" and b.get("level") == 1 and b.get("break")
          and (b.get("geom") or {}).get("top") is not None]
    if len(h1) < 2:
        return
    sig, ys = {look(b) for b in h1}, [b["geom"]["top"] for b in h1]
    for b in blocks:
        g = b.get("geom") or {}
        pn, t = g.get("page"), text_of(b.get("runs") or []).strip() if b["k"] == "p" else ""
        if t and not b.get("break") and pn in tops and g.get("top") is not None and abs(g["top"] - tops[pn]) < 2                 and len(t.split()) <= 6 and t[-1] not in ".:;," and look(b) in sig and min(abs(g["top"] - y) for y in ys) < 6:
            b["break"] = True


def cell_fills(rows, fills):
    """Заливка ячейки, как в оригинале: c["fill"] = "RRGGBB" — цветной прямоугольник под всем текстом ячейки
    (наименьший такой). Серая шапка таблицы ICS OAH 4.22 — ячейки TD, не TH, и заливка шапки по TH её не видела;
    заливка строк и отдельных ячеек терялась совсем (18.10). fills(лист) → [(x0, top, x1, bottom, цвет)]."""
    for row in rows:
        for c in row:
            g = c.get("geom")
            if not g or g.get("bottom") is None or g.get("x1") is None:
                continue
            under = [f for f in fills(g["page"]) if f[0] - 1.5 <= g["x0"] and g["x1"] <= f[2] + 1.5
                     and f[1] - 1.5 <= g["top"] and g["bottom"] <= f[3] + 1.5]
            if under:
                c["fill"] = min(under, key=lambda f: (f[2] - f[0]) * (f[3] - f[1]))[4]


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


def _page_sizes(path):
    """Размеры листов оригинала (ширина, высота) с учётом поворота (/Rotate): альбомные листы — альбомными разделами
    (docx_build.Builder._page)."""
    import fitz
    with fitz.open(str(path)) as d:
        return [(round(pg.rect.width, 1), round(pg.rect.height, 1)) for pg in d]


def _zone_keys(texts):
    """Тексты колонтитулов листов без счётчика листов: число, что у большинства листов своё («Oldalszám:» 3, 4, 5 …), —
    «#»; число с немногими значениями (версия 0 или 1, идентификатор) остаётся — листы с разной версией различимы.
    Число — по тексту перед ним (до 10 знаков). Ключ — набор знаков без порядка: строки шапки на соседних листах
    читаются в разном порядке (разница в долю pt по высоте). {лист: текст} → ({лист: ключ}, {тексты перед счётчиками})."""
    parts = {pn: re.split(r"(\d+)", t) for pn, t in texts.items()}
    vals = {}
    for ps in parts.values():
        for i in range(1, len(ps), 2):
            vals.setdefault(ps[i - 1][-10:], []).append(ps[i])
    counter = {c for c, v in vals.items() if len(set(v)) >= max(3, 0.5 * len(v))}
    return {pn: "".join(sorted("".join("#" if i % 2 and ps[i - 1][-10:] in counter else x for i, x in enumerate(ps))))
            for pn, ps in parts.items()}, counter


def layout(path, blocks, content=None):
    """Раскладка листа по оригиналу: поля тела и титула, шапка (картинки по краям, строки по центру) и штамп
    внизу (сетка по рамкам ячеек, строки по словам). Колонтитулы — артефакты PDF, в дереве структуры их нет.
    content — PDF без разметки (docx_layout, текстовый слой): {лист от 1: [рамки текста тела]} — знаки в них считаются
    содержимым, прочие — колонтитулами, как артефакты у размеченного."""
    import pdfplumber
    from statistics import median
    pdf = pdfplumber.open(path)
    cbots = {}
    if content:
        for pg in pdf.pages:
            bx = content.get(pg.page_number) or []
            # знаки в рамках текста тела — содержимое, прочие — колонтитулы, даже с MCID: PDF с неполной разметкой (формы
            # RDGT — все знаки помечены P) идёт сюда как текстовый слой, и шапка с MCID считалась телом — колонтитула
            # не было, шапка 1-го листа терялась (20.10)
            for c in pg.chars:
                inside = any(b[0] - 1 <= float(c["x0"]) and float(c["x1"]) <= b[2] + 1
                             and b[1] - 2 <= float(c["top"]) and float(c["bottom"]) <= b[3] + 2 for b in bx)
                c["mcid"] = -1 if inside else None
            cbots[pg.page_number] = max((b[3] for b in bx), default=0.0)
    cbot = (lambda pg: cbots.get(pg.page_number, 0.0)) if content else (lambda pg: None)
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
            and b["k"] in ("p", "h") and not (content and b.get("figure"))]
    if content:
        # текстовый слой: у бланка текст — в таблицах, поля и верх тела — и по ним; снимок скана (во весь лист, от края)
        # — не мера поля: у 32RDGT левое поле выходило 0, верх тела — над штампом, и шапки не было (20.10)
        body += [b["geom"] for b in blocks if b["k"] == "table" and ((b.get("geom") or {}).get("page") or 0) > 2]
    # левое поле — самый левый край, повторяющийся у трёх и больше блоков (заголовки у поля), а не медиана: у OAH тело
    # сдвинуто от заголовков на 37 pt, поле выходило внутри сдвига, и центр шапки и логотипа не узнавался (16.10);
    # сдвиг тела — отступом абзаца (docx_build.paragraph)
    xs = Counter(round(g["x0"]) for g in body)
    rep = sorted(x for x, n in xs.items() if n >= 3)
    left = float(rep[0]) if rep else median(sorted(g["x0"] for g in body)[: max(1, len(body) // 2)]) if body else 62
    right = W - max((g["x1"] for g in body), default=W - 59)
    # верх тела — по первому блоку листов, 20-й процентиль, а не минимум: надпись у края одного листа (МТ-ЦАСУТП-14,
    # лист 19: 3 pt) давала верхнее поле −0.9 pt, и шапка (знаки выше верха тела) терялась целиком (16.10)
    firsts = {}
    for g in body:
        firsts[g["page"]] = min(firsts.get(g["page"], g["top"]), g["top"])
    tops_ = sorted(firsts.values())
    body_top = tops_[len(tops_) // 5] if tops_ else 110
    page, page_tok, page_off, tok_boxes = _typical_page(pdf, body_top, H, cbot=cbot)
    # номер листа — помечен в знаках: полем станет он, а не такое же число рядом («Блок 3» на листе 3, синтетика)
    mark = lambda chars: [dict(c, text=PAGE_MARK) if c["text"].isdigit() and any(
        abs(float(c["top"]) - bx["top"]) < 2 and bx["x0"] - 0.5 <= float(c["x0"]) and float(c["x1"]) <= bx["x1"] + 0.5
        for bx in tok_boxes) else c for c in chars]

    def edges(pg):
        """Знаки вне тела листа pg: (шапки — выше тела, подвала — от верха подвала, верх подвала)."""
        art = [c for c in pg.chars if c.get("mcid") is None and _upright(c)]
        ft = _foot_top(art, _drawn(pg.rects + pg.lines), H, cbot(pg))
        return [c for c in art if c["top"] < body_top], [c for c in art if c["top"] >= ft], ft

    # колонтитулы листов текстового слоя (PDF без разметки) — по тексту вместе с цифрами, кроме счётчика листов: версия
    # «1» на одних листах и «0» на других, «3PR10» и «3PR10/INFO» — не один колонтитул на документ. Образец тела — лист
    # с самым частым текстом, а не случайный из листов одной подписи без цифр (3PR10, 3SZ17, 30PR25GT, 30RVR*: версия и
    # идентификатор в шапке всех листов — с образца, аудит 29.10)
    def zone_chars(pg):
        hc, fc, _ = edges(pg)
        return sorted((c for c in hc + fc if c["text"].strip()), key=lambda c: (round(float(c["top"]) / 3), float(c["x0"])))
    zone_text = lambda pg: "".join(c["text"] for c in zone_chars(pg))
    zkeys, counter = {}, set()
    base_key = None
    if content and len(pdf.pages) > 2:
        zkeys, counter = _zone_keys({pg.page_number: zone_text(pg) for pg in pdf.pages})
        common = Counter(zkeys[pn] for pn in range(2, len(pdf.pages) + 1)).most_common(1)[0]
        # разных текстов — не больше половины листов; иначе счётчик листов не нашёлся (у каждого листа свой текст), и
        # раздельных колонтитулов нет
        if 1 < len({zkeys[pn] for pn in range(2, len(pdf.pages) + 1)}) <= 0.5 * (len(pdf.pages) - 1):
            base_key = common[0]
            if zkeys[page.page_number] != base_key:
                same_ = [pn for pn in range(2, len(pdf.pages) + 1) if zkeys[pn] == base_key]
                page = pdf.pages[same_[len(same_) // 2] - 1]
    head_chars, foot_chars, ftop = edges(page)
    # картинка больше четверти листа — сам лист (скан посреди текстового слоя) или фон, не логотип колонтитула: скан
    # 32RDGT шёл в шапку вводной части сырой картинкой без поворота листа — полосами (20.10)
    area = lambda i: (float(i["x1"]) - float(i["x0"])) * (float(i["bottom"]) - float(i["top"]))
    logo = lambda i: area(i) < 0.25 * W * H
    # и на листе-скане (картинки — полосами, каждая меньше четверти) картинок колонтитула нет: полоса скана 32RDGT
    # шла в шапку вводной части
    scanned = lambda pg: sum(area(i) for i in pg.images) >= 0.5 * W * H
    # на листе-скане не логотип — картинка, что касается других картинок листа (мозаика скана); логотип шапки над
    # сканом титула стоит отдельно (ИН-1065: логотип пропадал из шапки титула)
    near = lambda i, j: (float(j["x0"]) - 2 <= float(i["x1"]) and float(i["x0"]) <= float(j["x1"]) + 2
                         and float(j["top"]) - 2 <= float(i["bottom"]) and float(i["top"]) <= float(j["bottom"]) + 2)
    same = lambda i, j: all(abs(float(i[k]) - float(j[k])) < 0.01 for k in ("x0", "top", "x1", "bottom"))
    mosaic = lambda i, pg: scanned(pg) and any(not same(i, j) and near(i, j) for j in pg.images)
    img_of = lambda i: {"xref": i["stream"].objid, "x0": float(i["x0"]), "x1": float(i["x1"]), "top": float(i["top"]),
                        "bottom": float(i["bottom"])}

    def zones(pg, own=False, body=False):
        """Шапка (картинки, строки, таблица или абзацы) и штамп листа pg по знакам вне тела — поля раскладки; знаки шапки
        и подвала, верх подвала и картинки подвала — для полей страницы. own — листы с разными колонтитулами (серии):
        номер листа — не там, где у образца, и метки — на счётчике листов по тексту перед числом (_zone_keys), его число —
        page_tok листа; счётчика в колонтитуле нет — номера листа нет (у тела — как у образца, по месту)."""
        head_chars, foot_chars, ftop = edges(pg)
        mk, tok = mark, page_tok
        if own:
            zc = zone_chars(pg)
            txt, prev, hit = "".join(c["text"] for c in zc), 0, []
            for m in re.finditer(r"\d+", txt):
                if txt[prev:m.start()][-10:] in counter:
                    hit.append(m)
                prev = m.end()
            if not hit and not body:
                mk, tok = (lambda chars: chars), None
            if hit:
                at, ids = 0, set()
                for c in zc:                                # знак может быть из нескольких букв (лигатура)
                    if hit[0].start() <= at < hit[0].end():
                        ids.add(id(c))
                    at += len(c["text"])
                mk = lambda chars: [dict(c, text=PAGE_MARK) if id(c) in ids else c for c in chars]
                tok = hit[0].group()
        head_imgs = [img_of(i) for i in pg.images if i.get("mcid") is None and float(i["top"]) < body_top and logo(i)]
        head_imgs += _vector_logos(pg, body_top, head_chars)
        hw = _words(head_chars) if head_chars else []
        lines = {}
        for w in hw:
            lines.setdefault(round(float(w["top"]) / 4), []).append(w)
        head_lines = [{"t": " ".join(w["text"] for w in sorted(ws, key=lambda w: w["x0"])),
                       "size": round(float(ws[0].get("size", 10)) if "size" in ws[0] else 11),
                       "x0": min(float(w["x0"]) for w in ws), "x1": max(float(w["x1"]) for w in ws),
                       "top": min(float(w["top"]) for w in ws), "bottom": max(float(w["bottom"]) for w in ws)}
                      for _, ws in sorted(lines.items())]
        # штамп: столбцы — по рамкам ячеек (нет внутренних линий — по разрывам между группами слов), строки — по словам
        frects = [r for r in _drawn(pg.rects + pg.lines) if float(r["top"]) >= ftop]
        # картинки подвала (логотип OAH на каждом листе) — артефакты ниже тела
        fz = min(ftop, H * 0.8)
        foot_imgs = [img_of(i) for i in pg.images if i.get("mcid") is None and float(i["top"]) >= fz and logo(i)]
        foot = _rule_grid(mk(foot_chars), frects)
        foot_paras = None if foot else _zone_paras(mk(foot_chars), frects, left, W - right, foot_imgs, W)
        foot = foot or _zone_grid(mk(foot_chars), frects)
        # шапка-таблица по линиям (с логотипом в своей ячейке); без рамки — абзацами (_zone_paras); в рамке без
        # внутренних линий (ÜFK: строка-рамка «Üzemeltetési Feltételek … | 1.2. változat») — сеткой по разрывам
        hrects = _head_rects(pg, body_top)
        # линия под шапкой — только если есть в оригинале (у НГТУ её нет; прежний путь шапки с логотипами рисовал всегда)
        head_rule = any(float(r["bottom"]) - float(r["top"]) < 3 and float(r["x1"]) - float(r["x0"]) > 0.5 * (W - left - right)
                        for r in hrects)
        head = _rule_grid(mk(head_chars), hrects, head_imgs)
        if not head and not (head_imgs and head_chars):        # логотипы по краям и строки по центру — прежним путём (UTS)
            hp = _zone_paras(mk(head_chars), hrects, left, W - right, head_imgs, W)
            head = {"rows": [], "paras": hp, "font": hp["font"], "size": hp["size"]} if hp else (
                _zone_grid(mk(head_chars), hrects) if head_chars else None)
        fields = {"head_imgs": head_imgs, "head_lines": head_lines, "head_grid": head,
                  "stamp": foot["rows"], "stamp_widths": foot["widths"], "stamp_aligns": foot["aligns"],
                  "stamp_font": foot["font"] or "Times New Roman", "stamp_size": foot["size"], "stamp_box": foot["box"],
                  "stamp_inner": foot["inner"], "stamp_merge": foot.get("merge"),
                  "page_marked": any(c["text"] == PAGE_MARK for c in mk(foot_chars + head_chars)),
                  "stamp_paras": foot_paras, "head_rule": head_rule}
        fields["page_tok"] = tok
        return fields, head_chars, foot_chars, ftop, foot_imgs
    fields, head_chars, foot_chars, ftop, foot_imgs = zones(page, own=base_key is not None, body=True)
    head_imgs = fields["head_imgs"]
    est = Counter()                                         # текст листа по блокам — кандидаты в почти пустые
    for x in list(_paras(blocks)) + [b for b in blocks if b["k"] == "h"]:
        if (x.get("geom") or {}).get("page"):
            est[x["geom"]["page"]] += len(text_of(x.get("runs") or []).strip())
    # подвала нет — низ тела по последнему блоку листов (80-й процентиль): поле по умолчанию (100 pt) отнимало у
    # каждого листа МТ-ЦАСУТП-14 три строки (60 листов → 69, 16.10)
    lasts = {}
    for g in body:
        lasts[g["page"]] = max(lasts.get(g["page"], 0), g.get("bottom") or g["top"])
    bots_ = sorted(lasts.values())
    foot_top = min([float(c["top"]) for c in foot_chars] + [i["top"] for i in foot_imgs],
                   default=bots_[len(bots_) * 4 // 5] + 14 if bots_ else H - 100)
    foot_bottom = max([float(c["bottom"]) for c in foot_chars] + [i["bottom"] for i in foot_imgs], default=H - 40)
    # шапки нет — пустой абзац шапки (её расстояние + строка) должен уместиться в верхнее поле: при поле 24 pt и
    # расстоянии 36 pt тело уезжало на каждом листе вниз на 2–3 строки (ТЭ357: 66 листов → 85, 16.10)
    head_top = min([i["top"] for i in head_imgs] + [float(c["top"]) for c in head_chars] or
                   [max(6.0, min(36.0, body_top - 4 - 18))])
    sig = lambda chars: re.sub(r"\s+", " ", re.sub(r"\d+\.?", "", "".join(c["text"] for c in chars))).strip()
    body_sig = sig(foot_chars)
    out = {"page_sizes": _page_sizes(path),
           "W": W, "H": H, "left": left, "right": max(right, 20), "top": body_top - 4,
           "bottom": H - foot_top + 12, "header": head_top, "footer": max(H - foot_bottom - 6, 18),
           **fields,
           # почти пустые листы (меньше 300 знаков текста тела): пустоты между абзацами — по оригиналу (сборщик)
           "sparse": [pg.page_number for pg in pdf.pages if pg.page_number >= 2 and est[pg.page_number] < 400 and sum(
               1 for c in pg.chars if c.get("mcid") is not None and c["text"].strip()) < 300]}
    # вводная часть: листы после титула со своим подвалом (ÜFK: оглавление и лист изменений — «Tartalomjegyzék» без
    # номера) — своим разделом; нумерация тела — с числа на первом его листе («1. oldal» на 7-м листе, 16.10)
    # лист отличается колонтитулами — по шапке, подвалу и картинкам в них (у OAH вводные листы без шапки, а подвал —
    # тот же логотип без текста: по одному подвалу вводная часть не находилась, 16.10)
    def zone(pg):
        # повёрнутый текст на поле (JPDS: «Собственность… Запрещается…» вдоль края титула) — не колонтитул: абзацем
        # он рисовался столбиком по букве и растягивал титул на три листа (16.10)
        hc, fc, ft = edges(pg)
        him = [img_of(i) for i in pg.images if i.get("mcid") is None and float(i["top"]) < body_top and logo(i)
               and not mosaic(i, pg)]             + _vector_logos(pg, body_top, hc)
        fim = [img_of(i) for i in pg.images if i.get("mcid") is None and float(i["top"]) >= min(ft, H * 0.8)
               and logo(i) and not mosaic(i, pg)]
        return (sig(hc), sig(fc), len(him), len(fim)), hc, fc, ft, him, fim
    body_key = (sig(head_chars), body_sig, len(head_imgs), len(foot_imgs))

    def zone_layout(pg, hc, fc, ft, him, fim):
        """Колонтитулы отдельного листа (вводной части, первого листа): абзацами без рамки или сеткой."""
        frr = [r for r in _drawn(pg.rects + pg.lines) if float(r["top"]) >= ft]
        fg = _zone_grid(fc, frr)
        z = {"stamp": fg["rows"], "stamp_widths": fg["widths"], "stamp_aligns": fg["aligns"], "stamp_box": fg["box"],
             "stamp_inner": fg["inner"], "stamp_merge": None, "stamp_paras": _zone_paras(fc, frr, left, W - right, fim, W)}
        hr = _head_rects(pg, body_top)
        hp = _zone_paras(hc, hr, left, W - right, him, W)
        z["head_grid"] = {"rows": [], "paras": hp} if hp else (_zone_grid(hc, hr) if hc else None)
        z["head_lines"], z["head_imgs"] = [], []                   # не шапка тела прежним путём
        return z

    # первый лист — титул, если он по виду титул: текста меньше 40 % обычного листа или крупный шрифт при неполном
    # листе; одностраничный документ титула не имеет. Прежде титулом считался всякий первый лист с колонтитулами не
    # как у тела — и документ без колонтитулов (HSE, RS) или с особым первым листом (ONS) получал «титул» с полями
    # 20 pt и без выключки (16.10)
    ink = {pg.page_number: [c for c in pg.chars if c.get("mcid") is not None and c["text"].strip()]
           for pg in pdf.pages[:40]}
    counts = sorted(len(v) for k, v in ink.items() if k > 1)
    med = counts[len(counts) // 2] if counts else 0
    sizes = sorted(float(c["size"]) for v in ink.values() for c in v)
    msize = sizes[len(sizes) // 2] if sizes else 10
    t1 = ink.get(1, [])
    big = max((float(c["size"]) for c in t1), default=0) >= 1.5 * msize
    key1, hc1, fc1, ft1, him1, fim1 = zone(pdf.pages[0])
    # и лист со своими колонтитулами, отличными от тела (обложка UTS: текста много, но без шапки и штампа тела);
    # поля титула — как у тела, так что это то же, что «особый колонтитул первой страницы»
    # и обложка с фоном во весь лист: у UTS шапка и штамп тела нарисованы и на обложке, но под фоновой картинкой —
    # по колонтитулам лист 1 не отличался, обложка шла листом тела с шапкой и картинкой в потоке (16.10)
    p1 = pdf.pages[0]
    bg = any((float(i["x1"]) - float(i["x0"])) * (float(i["bottom"]) - float(i["top"])) > 0.5 * W * H for i in p1.images)
    # лист, с которого таблица уходит на следующий (бланк 30RDIT: на 1-м листе одна строка таблицы испытаний), — не
    # титул, хоть текста на нём мало: таблица целиком уходила в раздел титула, и у пяти листов не было шапки (20.10)
    runs_on = any(b["k"] == "table" and b.get("split") and (b.get("geom") or {}).get("page") == 1 for b in blocks)
    cover = len(pdf.pages) > 1 and (med > 0 and (len(t1) < 0.4 * med or big and len(t1) < 0.7 * med) and not runs_on
                                    or any(body_key) and key1 != body_key or bg)
    out["first_is_body"] = not cover
    if key1 != body_key:                    # свои колонтитулы первого листа (титул NSU — шапка с названием вуза)
        out["first"] = zone_layout(pdf.pages[0], hc1, fc1, ft1, him1, fim1)
    front = []
    for pg in pdf.pages[1:]:
        key, hc, fc, ft, him, fim = zone(pg)
        if not any(body_key) or key == body_key:
            break
        front.append((pg, fc, ft, hc, him, fim))
    # вводная часть короткая: подвал, меняющийся по главам (название главы), иначе съел бы листы тела
    if front and len(front) <= max(3, 0.15 * len(pdf.pages)):
        pg, fc, ft, hc, him, fim = front[0]
        z = zone_layout(pg, hc, fc, ft, him, fim)
        if (sig(hc), len(him)) == (sig(head_chars), len(head_imgs)):
            for k in ("head_grid", "head_lines", "head_imgs"):              # шапка — как у тела
                z.pop(k)
        out["front"] = dict(z, to=front[-1][0].page_number)
    if page_tok is not None:
        first = (out["front"]["to"] + 1) if "front" in out else 1 if out["first_is_body"] else 2   # 1-й лист тела
        out["num_start"] = {"page": first, "value": first - 1 + page_off}
    # листы тела, чей колонтитул отличается от колонтитула тела, — сериями подряд (docx_build: каждая серия — свой раздел);
    # шапка и подвал серии — по её первому листу, тем же разбором, что у тела
    if base_key is not None:
        first = (out["front"]["to"] + 1) if "front" in out else 1 if out["first_is_body"] else 2
        runs = []
        for pg in pdf.pages[first - 1:]:
            k = zkeys[pg.page_number]
            if k == base_key:
                continue
            if runs and runs[-1]["to"] == pg.page_number - 1 and runs[-1]["key"] == k:
                runs[-1]["to"] = pg.page_number
            else:
                runs.append({"from": pg.page_number, "to": pg.page_number, "key": k})
        if runs:
            out["variants"] = [dict({"from": r["from"], "to": r["to"]}, **zones(pdf.pages[r["from"] - 1], own=True)[0]) for r in runs]
    out["font_class"] = font_classes(path)
    return out


def _vector_logos(pg, body_top, chars):
    """Логотип шапки, нарисованный векторами (ONS: значок и название кривыми — в шапке обложки пропадал, 17.10):
    скопление от трёх фигур (не тонкие линии, не белое) выше тела листа, не на тексте шапки, — снимком области
    ({"clip"}, как составные рисунки docx_struct.compose_figures)."""
    W = float(pg.width)
    objs = []
    for o in list(pg.curves) + list(pg.rects):
        x0, x1, top, bottom = float(o["x0"]), float(o["x1"]), float(o["top"]), float(o["bottom"])
        col = o.get("non_stroking_color")
        white = isinstance(col, (list, tuple)) and col and min(float(c) for c in col) > 0.95 and not o.get("stroke")
        if bottom < body_top and x1 - x0 > 1.5 and bottom - top > 1.5 and x1 - x0 < 0.6 * W and not white:
            objs.append([x0, top, x1, bottom, 1])
    groups = []
    for b in objs:
        hit = [g for g in groups if g[0] - 6 <= b[2] and b[0] <= g[2] + 6 and g[1] - 6 <= b[3] and b[1] <= g[3] + 6]
        for g in hit:
            groups.remove(g)
            b = [min(b[0], g[0]), min(b[1], g[1]), max(b[2], g[2]), max(b[3], g[3]), b[4] + g[4]]
        groups.append(b)
    out = []
    for x0, top, x1, bottom, n in groups:
        if n < 3 or x1 - x0 < 15 or bottom - top < 8 or any(
                x0 <= float(c["x0"]) <= x1 and top <= float(c["top"]) <= bottom for c in chars if c["text"].strip()):
            continue
        out.append({"xref": None, "clip": [x0 - 1, top - 1, x1 + 1, bottom + 1], "page_number": pg.page_number,
                    "x0": x0 - 1, "x1": x1 + 1, "top": top - 1, "bottom": bottom + 1})
    return out


def font_classes(path, pages=12):
    """Класс шрифтов оригинала — по флагам шрифта PDF (с засечками, моноширинный), имя — как в DOCX (_fmt):
    {"Open Sans": "swiss", "Times New Roman": "roman", "Courier New": "modern"}. Без класса в таблице шрифтов DOCX
    неустановленный шрифт LibreOffice заменял шрифтом с засечками (руководства OAH набраны Open Sans, 16.10)."""
    import fitz
    votes = {}
    with fitz.open(path) as d:
        for pg in list(d)[:pages]:
            for bl in pg.get_text("dict")["blocks"]:
                for ln in bl.get("lines", []):
                    for sp in ln["spans"]:
                        if not sp["text"].strip():
                            continue
                        fam = _fmt({"fontname": sp["font"], "size": sp["size"]})["font"]
                        k = "modern" if sp["flags"] & 8 else "roman" if sp["flags"] & 4 else "swiss"
                        votes.setdefault(fam, Counter())[k] += len(sp["text"])
    return {f: c.most_common(1)[0][0] for f, c in votes.items()}


def _drawn(objs):
    """Линии и рамки, видные на листе: белая заливка без обводки (подложка поля бланка — у 30RDGT девять таких под
    словами штампа) — не край ячейки; по её краям штамп резался на 13 узких столбцов и слова шли столбиком (20.10)."""
    def white(c):
        v = [float(x) for x in (c if isinstance(c, (list, tuple)) else [c] if isinstance(c, (int, float)) else [])]
        return bool(v) and (max(v) < 0.05 if len(v) == 4 else min(v) > 0.95)
    return [o for o in objs if o.get("stroke", True) or o.get("fill") and not white(o.get("non_stroking_color"))]


def _head_rects(page, body_top):
    """Линии зоны шапки: выше тела и выше первого содержимого листа с запасом — рамка таблицы, которой начат лист,
    лежит выше её первого знака и попадала в шапку лишними строками и столбцами (синтетика, 16.10)."""
    ctop = min((float(c["top"]) for c in page.chars if c.get("mcid") is not None), default=body_top)
    lim = min(body_top, ctop - 6)
    return [r for r in _drawn(page.rects + page.lines) if float(r["bottom"]) < lim]


def _zone_paras(chars, rects, x_left, x_right, imgs=(), page_w=None):
    """Колонтитул без рамки (нет вертикалей и обведённых прямоугольников) — абзацами по строкам оригинала, а не
    таблицей: части строки (группы слов с разрывом больше 18 pt) — по листу (по центру / у правого поля / слева);
    горизонтальная линия — граница абзаца над ней; картинка (логотип) — своим абзацем по месту. Шапка OAH (номер,
    лист/всего, версия — строкой, название — второй строкой, линия снизу) собиралась таблицей в рамке со
    смешанными частями, логотип в подвале терялся (16.10). → {"rows": [{"parts": [(текст, вид, позиция от левого
    поля)], "bold"} | {"img", "kind"}], "font", "size"} или None — колонтитул в рамке (таблицей)."""
    import pdfplumber
    vs = [r for r in rects if float(r["x1"]) - float(r["x0"]) < 3 and float(r["bottom"]) - float(r["top"]) > 5]
    boxed = [r for r in rects if r.get("object_type") == "rect" and r.get("stroke")
             and float(r["x1"]) - float(r["x0"]) >= 3 and float(r["bottom"]) - float(r["top"]) >= 3]
    if vs or boxed:
        return None
    words = _words(chars) if chars else []
    if not words and not imgs:
        return None
    mid = (x_left + x_right) / 2
    # по центру — полосы набора или листа (логотип OAH стоит по середине листа, 16.10)
    kind = lambda a, b: "center" if min(abs((a + b) / 2 - mid), abs((a + b) / 2 - page_w / 2) if page_w else 99) < 10 \
        else "right" if abs(b - x_right) < 10 else "left"
    pos = lambda k, a, b: round(b - x_left if k == "right" else mid - x_left if k == "center" else a - x_left, 1)
    lines = {}
    for w in words:
        lines.setdefault(round(float(w["top"]) / 3), []).append(w)
    rows = []
    for _, ws in sorted(lines.items()):
        ws = sorted(ws, key=lambda w: float(w["x0"]))
        groups, cur = [], [ws[0]]
        for w in ws[1:]:
            if float(w["x0"]) - float(cur[-1]["x1"]) > 18:
                groups.append(cur)
                cur = [w]
            else:
                cur.append(w)
        groups.append(cur)
        parts = []
        for g in groups:
            a, b = float(g[0]["x0"]), float(g[-1]["x1"])
            k = kind(a, b)
            parts.append((" ".join(w["text"] for w in g), k, pos(k, a, b)))
        top, bot = min(float(w["top"]) for w in ws), max(float(w["bottom"]) for w in ws)
        ink = [c for c in chars if top - 1 <= float(c["top"]) <= bot and c["text"].strip()]
        # цвет строки — по большинству знаков (красное «CONTROLLED DISCLOSURE» в подвале Eskom выходило чёрным, 17.10)
        col = Counter(_fmt(c).get("color") or "000000" for c in ink).most_common(1)[0][0] if ink else "000000"
        rows.append({"parts": parts, "bold": bool(ink) and all(_fmt(c)["b"] for c in ink), "top": top, "bottom": bot,
                     "color": col})
    for im in imgs:
        rows.append({"img": im, "kind": kind(im["x0"], im["x1"]), "top": im["top"], "bottom": im["bottom"]})
    rows.sort(key=lambda r: r["top"])
    for r in rects:
        if float(r["bottom"]) - float(r["top"]) < 3 and float(r["x1"]) - float(r["x0"]) > 0.5 * (x_right - x_left):
            y = (float(r["top"]) + float(r["bottom"])) / 2
            above = [row for row in rows if row["bottom"] <= y + 2]
            (above[-1] if above else rows[0])["rule_below" if above else "rule_above"] = True
    f = Counter((_fmt(c)["font"], _fmt(c)["size"]) for c in chars if c["text"].strip()).most_common(1)
    return {"rows": rows, "font": f[0][0][0] if f else None, "size": f[0][0][1] if f else None}


PAGE_WORD = re.compile(r"(?:\D*[:.\-–#№])?(\d+)(?:\.|/\d+)?")


def _typical_page(pdf, body_top, H, sample=15, cbot=lambda pg: None):
    """Лист для колонтитулов — с самыми частыми колонтитулами (текст без цифр), а не второй по счёту: у ÜFK
    второй лист — оглавление, и в подвал всех листов попадало «Tartalom» («Содержание», 16.10). Заодно — номер
    листа: число колонтитула, которое растёт вместе с листом («13. oldal» на 19-м листе) → (лист, его текст,
    смещение: номер = индекс листа от 0 + смещение, слово номера на этом листе)."""
    import pdfplumber
    n = len(pdf.pages)
    # выборка — по всему документу, и 1-й лист: титул со своими колонтитулами в общую группу не попадёт, а короткий
    # документ без титула (два листа) иначе не давал номера листа — в штампе застывала цифра (синтетика, 16.10)
    idx = sorted({round(k * (n - 1) / max(sample - 1, 1)) for k in range(sample)})
    sig, words, arts_of = {}, {}, {}
    for i in idx:
        pg = pdf.pages[i]
        art = arts_of[i] = [c for c in pg.chars if c.get("mcid") is None]
        ftop = _foot_top(art, _drawn(pg.rects + pg.lines), H, cbot(pg))
        arts = [c for c in art if c["top"] < body_top or c["top"] >= ftop]
        # без пробелов: пробел в конце строки тела выходит за её рамку (PDF без разметки) и у каждого листа свой —
        # листы с одинаковым штампом не находились, номер листа застывал числом (32RDGT, 20.10)
        sig[i] = re.sub(r"\d+", "#", "".join(c["text"] for c in arts if c["text"].strip()))
        words[i] = _words(arts) if arts else []
    common = Counter(s for s in sig.values() if s.strip()).most_common(1)
    same = [i for i in idx if common and sig[i] == common[0][0]] or idx
    rep = same[len(same) // 2]
    # номер листа может стоять в нескольких местах (в шапке и в штампе): помечены все — первый задаёт смещение
    tok, off, boxes = None, 0, []
    if len(same) >= 2 and len({len(words[i]) for i in same}) == 1:
        for j, w in enumerate(words[rep]):
            # «13.», «13» и «13/57» (лист/всего листов, OAH): номер — первая часть; приклеенный к подписи двоеточием
            # или точкой («Oldalszám:7.», 30RR2126GT) — тоже: слово целиком в число не складывалось, и номер листа
            # застывал числом одного листа во всех колонтитулах
            vals = [PAGE_WORD.fullmatch(words[i][j]["text"]) for i in same]
            t = w["text"][vals[same.index(rep)].start(1):].split("/")[0] if all(vals) else ""
            if all(vals) and len({int(v[1]) - i for v, i in zip(vals, same)}) == 1 \
                    and (tok is None or int(vals[0][1]) - same[0] == off and t == tok):
                if tok is None:
                    tok, off = t, int(vals[0][1]) - same[0]      # номер листа = индекс (от 0) + off
                bx = {k: float(w[k]) for k in ("x0", "x1", "top")}
                if "/" in w["text"]:                             # метка — только цифры до «/»
                    sl = [c for c in arts_of[rep] if c["text"] == "/" and abs(float(c["top"]) - bx["top"]) < 2
                          and bx["x0"] - 0.5 <= float(c["x0"]) <= bx["x1"]]
                    if sl:
                        bx["x1"] = float(sl[0]["x0"])
                boxes.append(bx)
    # подпись номера листа разная на листах («Oldalszám:» / «Oldal» «:» — 30BLFKSZ), общих колонтитулов с одним набором
    # слов нет: число ищется по месту — то же место на листах, значение растёт вместе с листом
    if tok is None and len(idx) >= 3:
        for w in words[rep]:
            mw = PAGE_WORD.fullmatch(w["text"])
            if not mw:
                continue
            pairs = []
            for i in idx:
                q = next((v for v in words[i] if abs(float(v["x0"]) - float(w["x0"])) < 12 and abs(float(v["top"]) - float(w["top"])) < 4
                          and PAGE_WORD.fullmatch(v["text"])), None)
                if q:
                    pairs.append((i, int(PAGE_WORD.fullmatch(q["text"])[1])))
            if len(pairs) >= 3 and any(i == rep for i, _ in pairs) and len({v - i for i, v in pairs}) == 1:
                tok, off = w["text"][mw.start(1):].split("/")[0], pairs[0][1] - pairs[0][0]
                boxes = [{k: float(w[k]) for k in ("x0", "x1", "top")}]
                break
    return pdf.pages[rep], tok, off, boxes


def _foot_top(arts, rects, H, cbot=None):
    """Верх подвала: нижняя группа строк-артефактов без разрыва больше 15 pt — или связанных вертикалью таблицы
    (строка штампа-таблицы выше 15 pt: подвал обрезался до пустой строки под таблицей, синтетика 16.10), плюс линия
    рамки над ней. Порог «ниже 80 % листа» захватывал низ таблицы и пометку «Folytatás a következő oldalon» — их
    линии резали подвал на лишние ячейки (ÜFK I, 16.10). cbot — низ текста тела листа (PDF без разметки): подвал —
    всё ниже него, без счёта разрывов («CONTROLLED DISCLOSURE» Eskom в 20 pt над прочими строками терялся)."""
    if cbot is not None:
        # текстовый слой: пробел — не краска, пробел в конце строки тела выходит за её рамку и считался подвалом
        # (32RDGT: верх подвала на 509 pt, нижнее поле 344 pt, 20.10). У размеченного — как было: у HSE без подвала
        # поле считалось по пробелу-артефакту, без него лист переполнялся (регрессия A/B)
        low = [float(c["top"]) for c in arts if float(c["top"]) > max(cbot, H * 0.6) and c["text"].strip()]
        if low:
            top = min(low)
            border = [float(r["top"]) for r in rects if float(r["bottom"]) - float(r["top"]) < 3
                      and max(cbot, H * 0.6) < float(r["top"]) <= top]
            return round(min(border) if border else top) - 1
    # текстовый слой без знаков ниже тела — и здесь пробел не краска: пробел в конце строки таблицы (443 pt при низе
    # тела 470) становился «подвалом», нижнее поле — 165 pt вместо ~100, на листе 4–5 строк таблицы вместо 6
    # (таблица решений OAH: 154 листа → 211, 23.10)
    rows = sorted({round(float(c["top"])) for c in arts if float(c["top"]) > H * 0.7
                   and (cbot is None or c["text"].strip())}, reverse=True)
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
    words = _words(chars) if chars else []
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
    words = _words(chars) if chars else []
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
