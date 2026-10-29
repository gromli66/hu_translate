# -*- coding: utf-8 -*-
"""DOCX распознанного PDF (текст листа — контурами, перевод — по сегментам OCR): та же модель блоков, что у размеченного
PDF (docx_struct), но по самим сегментам листа — дальше тот же профиль оформления, сборщик и печать PDF из DOCX.

Раньше такой PDF («Печать в PDF», 3SZ19, 3PR42) получал DOCX из переведённого PDF: строки резались, налезали,
кегль у каждого сегмента свой (оценка по картинке) — Максим, 16.10: «плавающий шрифт». Здесь:
- абзац = сегмент (перевод готов, портал не нужен), кегль тела — один на документ (медиана), шрифт — по семейству;
- колонтитулы — тексты, повторяющиеся в верхней/нижней части листов (цифры не в счёт), номер листа — полем;
- заголовок — короткий полужирный абзац по центру, уровень — по глубине номера («1.1.» — 2);
- пункт с тире — по маркеру «-» в текстовом слое листа (он там есть, текст нарисован контурами); уровень пункта —
  по отступу (отступы всех списков документа — ступенями);
- полужирные и курсивные слова — по толщине и наклону штриха на картинке листа (команды «ZÁRNI», «INDÍTANI», номера
  пунктов, «Megjegyzés:» курсивом): абзац — с метками начертания, перевод с метками — шаг «структура»
  (pipeline.struct_fill), как у размеченного PDF;
- таблица — по векторным линиям листа (связная сетка с внутренней вертикалью; рамка листа и названия раздела —
  линии, повторяющиеся на половине листов, — не в счёт), ячейки — по сетке, линии ячеек — как в оригинале;
- рисунок (график, схема) — то, что на листе не текст и не линии рамок: картинкой с листа оригинала, сегменты
  OCR внутри (подписи осей, мусор распознавания) — не абзацами."""
import re
from collections import Counter
from statistics import median

FAMILY = {"monospace": "Courier New", "serif": "Times New Roman", "sans-serif": "Arial"}
NUM = re.compile(r"^\s*(\d+(?:\.\d+)*)\.?\s+\S")
SIGN = re.compile(r"\s*([•▪◦·●○■□–\-\uf020-\uf0ff])\s")        # знак маркера в начале строки текстового слоя
PLABEL = re.compile(r"^\s*(\d{1,2}|[a-z]|[A-Z])(\.\)|\))\s+\S")  # «a)», «1)», «a.)» — как docx_struct.PLABEL


def is_scan(page):
    """Лист — картинка (скан): растровые изображения закрывают хотя бы половину листа."""
    import fitz
    return sum(fitz.Rect(i["bbox"]).get_area() for i in page.get_image_info()) >= 0.5 * page.rect.get_area()


def suits(path, ocr_pages):
    """Распознанный PDF собирается по сегментам листа, если распознано больше половины листов и это не сканы, а
    листы с текстом контурами («Печать в PDF»: 3SZ19, 3PR42). Скан формы (титул и лист регистрации 30RDGT, 30RDIT)
    так терял таблицу формы, подписи и рукописные даты (линии — пиксели, рукопись в сегменты не попадает) и идёт
    прежним путём: перевод поверх скана, сделанный под формы (16.10, проверка на документах сервиса)."""
    import fitz
    ocr = list(ocr_pages or [])
    with fitz.open(str(path)) as doc:
        if not ocr or 2 * len(ocr) < doc.page_count:
            return False
        return sum(1 for p in ocr if p < doc.page_count and is_scan(doc[p])) * 2 < len(ocr)


def _zone(s, H):
    y0, y1 = s["loc"]["bbox"][1], s["loc"]["bbox"][3]
    return "head" if y1 < 0.14 * H else "foot" if y0 > 0.86 * H else None


def _sig(text):
    return re.sub(r"\d+", "#", text.strip())


def _running(segs, H, pages):
    """Колонтитулы: (зона, текст без цифр) или (зона, высота строки), что повторяется хотя бы на половине листов (и не
    меньше чем на двух). По высоте — потому что OCR режет шапку по-разному: «Megnevezés … UTASÍTÁSA 72» на одном листе
    и «Verziószám: 0 Oldalszám:» отдельно — на другом (3PR42, 16.10). Считаются листы, а не куски: строка таблицы из
    трёх кусков у верха одностраничного документа — не «три повтора» (30RDGT modj: шапка таблицы и первая строка
    уходили в шапку DOCX, в таблице на их месте — пустые строки, 20.10)."""
    cnt, pos = {}, {}
    for s in segs:
        z = _zone(s, H)
        if z:
            cnt.setdefault((z, _sig(s["text"])), set()).add(s["page"])
            pos.setdefault((z, round(s["loc"]["bbox"][1] / 4)), {}).setdefault(_sig(s["text"]), set()).add(s["page"])
    need = max(2, 0.5 * pages)
    # на высоте — и текст в основном тот же (≥ половины листов): название раздела в рамке на той же высоте меняется
    # от раздела к разделу — это текст, а не колонтитул
    return {k for k, ps in cnt.items() if len(ps) >= need} | \
        {k for k, c in pos.items() if len(set().union(*c.values())) >= need and max(map(len, c.values())) >= 0.5 * pages}


def _is_running(s, keys, H):
    z = _zone(s, H)
    return bool(z) and ((z, _sig(s["text"])) in keys or (z, round(s["loc"]["bbox"][1] / 4)) in keys)


def _portal_running(s):
    """Колонтитул распознанного листа по метке портала (roles.py): OCR читает шапку по-разному на разных листах («T0.7. sz.
    útmutató» / «3.7. sz. útmutató»), и повтор по тексту её не узнаёт — шапка печаталась посреди текста (замер на сканах
    18.10: колонтитулы правилами 0,27)."""
    return bool(s["loc"].get("ocr")) and s["loc"].get("role") == "header"


def _stable_running(s, pool, keys, H):
    """Колонтитул — строка, что стоит так же почти на всех листах: повтор по тексту (_is_running) или строка с меткой
    портала, если на 80 % листов с колонтитулами (pool — повторы и метки портала) на той же высоте похожий текст (OCR
    читает шапку по-разному). Название раздела в рамке (3SZ19: «1. A turbina üzemzavarai» / «2. …») портал метит
    колонтитулом, но оно меняется от раздела к разделу — это текст: в шапке всех листов стояло название второго раздела
    (шапка ломалась на три столбца), а в теле заголовок раздела пропадал в начале раздела и оставался посреди (19.10)."""
    if _is_running(s, keys, H):
        return True
    from difflib import SequenceMatcher
    z, y, k = _zone(s, H), s["loc"]["bbox"][1], _sig(s["text"])
    if not z:
        return False
    have = {t["page"] for t in pool if _zone(t, H) == z}
    hit = {t["page"] for t in pool if _zone(t, H) == z and abs(t["loc"]["bbox"][1] - y) < 4
           and SequenceMatcher(None, k, _sig(t["text"])).ratio() >= 0.8}
    return len(hit) >= 0.8 * len(have)


def _role(f):
    """Роль строки (абзаца) PDF без разметки — одна на все источники, по старшинству (договор правил, шаг 2; PDF с
    тегами решает docx_struct по тегам, сюда не доходит):
      1. факты листа оригинала — повтор у края листа (колонтитул), строка шапки таблицы, знак маркера (тире слоя
         скана, «•»/«–» в начале строки), номер листа оглавления (_toc_line), название раздела в рамке (скан: метка
         портала «колонтитул», не одинаковая на всех листах; повтор вверху следующего листа);
      2. метка портала (roles.py) — где лист размечен порталом: роль и границы абзацев решает он, правила вида строки
         не применяются; «колонтитул» текстового слоя — не роль строки (колонтитулы там — правилами по повтору,
         roles.py), «подпись» и «таблица» — абзац (таблицы — по линиям);
      3. правила вида строки — лист без меток: короткий полужирный по центру (или с номером у края в текстовом слое) —
         заголовок, «a)», «1)» в начале — пункт.
    Столкновения, что решались каждое своей заплаткой (опись RULES.md, Р1–Р10), решает этот порядок: строка
    оглавления с меткой «пункт» — оглавление (Р1); метка «колонтитул» текстового слоя — не заголовок (Р4); знак маркера
    старше метки «заголовок» (Р6: «- 1. Ismérvek:» 3PR42 — пункт с тире, как на листе); абзац, начатый порталом,
    правила к пункту не клеят (Р7; load).
    f — факты: table_head, repeat, ocr, label (метка портала), stable (метка «колонтитул» похожа на всех листах),
    title, sign, toc, pieces (кусков в строке), labeled (лист с метками), typo (вид заголовка), num_label, text.
    → "running" (колонтитул — из текста прочь), "heading", "item", "toc", "body"."""
    if f.get("table_head"):
        return "body"
    if f.get("repeat") or f.get("ocr") and f.get("label") == "header" and f.get("stable"):
        return "running"
    text = f.get("text") or ""
    named = len(text) < 200 and re.search(r"[^\W\d_]{3}", text)
    if f.get("sign"):
        return "item"
    if f.get("toc"):
        return "toc"
    # заголовок — строка из одного куска: строка из нескольких кусков («1.1. ⇥ ÜV-4 reaktorvédelmi működés» скана) —
    # абзац с табуляцией по месту оригинала; заголовком номер уезжал к полю, название — к табуляции стиля, без
    # полужирного (проверено на 3PR42 20.10)
    one = f.get("pieces", 1) == 1
    if one and (f.get("title") or f.get("ocr") and f.get("label") == "header") and named:
        return "heading"
    if one and f.get("label") == "heading" and named:
        return "heading"
    if f.get("label") == "list":
        return "item"
    if f.get("labeled"):
        return "body"
    if f.get("typo"):
        return "heading"
    if f.get("num_label"):
        return "item"
    return "body"


def _toc_line(group, right):
    """Строка оглавления: в конце — номер листа («13», «56/1») отдельным куском у правого поля или дальше 60 pt от
    текста строки (номер внутри рамки). Одно определение для роли строки: прежде их было два — по табуляции к полю
    (_runs_of, rtab) и по промежутку (Р10)."""
    if len(group) < 2 or not re.fullmatch(r"\d{1,4}(?:/\d+)?", group[-1]["text"].strip()):
        return False
    b = group[-1]["loc"]["bbox"]
    return right is None or b[2] > right - 40 or b[0] - group[-2]["loc"]["bbox"][2] > 60


def _layer(doc):
    """Маркеры пунктов «-»/«–» текстового слоя (текст нарисован контурами, маркеры — шрифтом): {лист: {"dash":
    [(x1, y_середина, высота)]}}."""
    out = {}
    for pno, page in enumerate(doc):
        out[pno] = {"dash": [(w[2], (w[1] + w[3]) / 2, w[3] - w[1]) for w in page.get_text("words")
                             if w[4] in ("-", "–", "—", "•")]}
    return out


def _slant(ink):
    """Наклон слова: во сколько раз резче вертикальная проекция краски, если строки сдвинуть влево (выпрямить курсив,
    наклон 0,15–0,25); по пролётам текстового слоя у курсива 1,15–1,65, у прямого — не больше 1,006 (3PR42; у
    Courier 3SZ19 прямой — до 1,033, курсив «FIGYELEM!» — 1,05: его порог 1,08 не берёт)."""
    import numpy as np
    ys, xs = np.nonzero(ink)
    if xs.size < 30:
        return 1.0
    h = ink.shape[0]
    base = float((np.bincount(xs).astype(float) ** 2).sum())
    best = base
    for sh in (0.15, 0.2, 0.25):
        idx = xs - np.round(sh * (h - 1 - ys)).astype(int) + int(sh * h) + 1
        best = max(best, float((np.bincount(idx).astype(float) ** 2).sum()))
    return best / base if base else 1.0


def _text_style(doc, segs):
    """Полужирные, курсивные и цветные слова по шрифтам текстового слоя (PDF без разметки, не контуры): то же, что
    _pixel_style, — {id сегмента: [(кусок, полужирный, курсив, цвет)]} у сегментов, где выделение или цвет есть
    (цвет — «rrggbb», чёрный и серый — «000000»: красное «Controlled Disclosure» Eskom выходило чёрным, 17.10). Знаки
    слоя ↔ буквы текста сегмента — по порядку непробельных знаков; слово — по большинству своих букв."""
    out = {}
    by_page = {}
    for s in segs:
        by_page.setdefault(s["page"], []).append(s)

    def hue(c):
        r, g, b = c >> 16 & 255, c >> 8 & 255, c & 255
        return f"{c:06x}" if max(r, g, b) - min(r, g, b) > 40 else "000000"
    for pno, ss in by_page.items():
        chars = []                                          # (знак, x-центр, y-центр, полужирный, курсив, цвет)
        for bl in doc[pno].get_text("rawdict")["blocks"]:
            for ln in bl.get("lines", []):
                for sp in ln["spans"]:
                    b = bool(sp["flags"] & 16) or bool(re.search(r"Bold|Black|Heavy|Semibold", sp["font"], re.I))
                    i = bool(sp["flags"] & 2) or bool(re.search(r"Italic|Oblique", sp["font"], re.I))
                    col = hue(sp.get("color") or 0)
                    for c in sp["chars"]:
                        if c["c"].strip():
                            x0, y0, x1, y1 = c["bbox"]
                            chars.append((c["c"], (x0 + x1) / 2, (y0 + y1) / 2, b, i, col))
        for s in ss:
            frags = s["loc"].get("frags") or [s["loc"]["bbox"]]
            mine = [c for c in chars if any(f[0] - 1 <= c[1] <= f[2] + 1 and f[1] - 1 <= c[2] <= f[3] + 1 for f in frags)]
            if not mine or not any(c[3] or c[4] or c[5] != "000000" for c in mine):
                continue
            seq = [(c[3], c[4], c[5]) for c in mine]
            words = re.findall(r"\S+\s*|\s+", s["text"])
            k, flags = 0, []
            for w in words:
                n = len(re.sub(r"\s", "", w))
                part = seq[k:k + n] or [seq[min(k, len(seq) - 1)]]
                k += n
                flags.append((sum(f[0] for f in part) * 2 > len(part), sum(f[1] for f in part) * 2 > len(part),
                              Counter(f[2] for f in part).most_common(1)[0][0]))
            pieces = []
            for w, f in zip(words, flags):
                if pieces and pieces[-1][1:] == f:
                    pieces[-1] = (pieces[-1][0] + w,) + f
                else:
                    pieces.append((w,) + f)
            if any(p[1] or p[2] or p[3] != "000000" for p in pieces if p[0].strip()):
                out[s["id"]] = pieces
    return out


def _pixel_style(doc, segs, layer=None):
    """Полужирные и курсивные слова по картинке листа (300 dpi). Текстовый слой для этого не годится: у 3SZ19
    полужирная строка оглавления «Turbina csapágy fehérfém…» лежит в нём только как «Turbina csapág», остальное
    нарисовано контурами (16.10).
    - Слово — пятно краски между самыми широкими разрывами строки, их — по числу слов текста (у Courier промежуток
      между буквами при 300 dpi доходит до трети высоты строки, у Times пробел — пятая её часть: порог на всех не
      годится, а число слов OCR знает); слова текста ↔ пятна — по длинам (_align).
    - Полужирное: толщина штриха — удвоенная площадь краски на длину её контура (не зависит от наклона и
      непрерывна), к высоте строки, но не меньше обычной; порог — 1,1 медианы документа или листа (что больше). Замер по пролётам слоя:
      жирный 0,096–0,103 против 0,049–0,051 (3SZ19), от 0,084 против 0,061–0,067 (3PR42); заглавные и коды с
      косыми («KM2/1») — не толще 0,069. По словам OCR у 3PR42 обычные — до 1,02 медианы, жирные — от 1,14
      («Megjegyzés:», полужирный курсив ссылок «ÜFK …»).
    - Курсив: _slant больше 1,08.
    Курсив короткого слова (до 3 букв: «Ha a», «és a») — от ближайших длинных слов (у «A», «7» косые штрихи);
    полужирных не меньше 65 % букв сегмента — полужирный весь (у полужирного курсива штрих тоньше, ссылка
    «ÜFK 4.6.3.3. A villamos …» выходила пятнами). {id сегмента: [(кусок,
    полужирный, курсив)]} — у сегментов, где выделение есть; layer — маркеры пунктов (_layer), их краска — не слово."""
    import fitz
    import numpy as np
    k = 300 / 72
    by_page, parts = {}, []
    for s in segs:
        by_page.setdefault(s["page"], []).append(s)
    for pno, ss in by_page.items():
        pix = doc[pno].get_pixmap(dpi=300, colorspace=fitz.csGRAY)
        img = (np.frombuffer(pix.samples, np.uint8).reshape(pix.height, pix.stride)[:, :pix.width] < 128).astype(np.uint8)
        marks = (layer or {}).get(pno, {}).get("dash", [])
        for s in ss:
            raw, lead = [], 0                                # строки: (краска, высота, пятна); lead — тире в начале
            for f in s["loc"].get("frags") or [s["loc"]["bbox"]]:
                y0, y1 = max(0, int(f[1] * k)), int(f[3] * k) + 1
                x0, x1 = max(0, int(f[0] * k)), int(f[2] * k) + 1
                a = img[y0:y1, x0:x1].copy()
                if a.size == 0 or y1 - y0 < 4:
                    continue
                for dx1, dy, _ in marks:                     # маркер пункта «-» в строке — не слово текста
                    if f[1] - 1 <= dy <= f[3] + 1 and f[0] - 2 <= dx1 <= f[0] + 30:
                        a[:, :max(0, int((dx1 + 1) * k) - x0)] = 0
                        if not raw and re.match(r"\s*[-–—•]\s", s["text"]):
                            lead = 1                         # тире есть и в тексте OCR — это слово не сопоставлять
                cols = np.nonzero(a.any(axis=0))[0]
                if not cols.size:
                    continue
                atoms, start = [], cols[0]
                for c0, c1 in zip(cols, cols[1:]):
                    if c1 - c0 > 1:
                        atoms.append((start, c0 + 1))
                        start = c1
                atoms.append((start, cols[-1] + 1))
                raw.append((a, y1 - y0, atoms))
            if not raw:
                continue
            need = len(s["text"].split()) - lead - len(raw)
            gaps = sorted(((atoms[i + 1][0] - atoms[i][1]) / h, li, i) for li, (_, h, atoms) in enumerate(raw)
                          for i in range(len(atoms) - 1))
            cut = {(li, i) for _, li, i in gaps[len(gaps) - need:]} if need > 0 else set()
            lines = []
            for li, (a, h, atoms) in enumerate(raw):
                runs, start = [], atoms[0][0]
                for i in range(len(atoms) - 1):
                    if (li, i) in cut:
                        runs.append((start, atoms[i][1]))
                        start = atoms[i + 1][0]
                runs.append((start, atoms[-1][1]))
                ink = a > 0
                q = np.pad(ink, 1)
                edge = ink & ~(q[:-2, 1:-1] & q[2:, 1:-1] & q[1:-1, :-2] & q[1:-1, 2:])    # краска у фона
                cl = []
                for a0, a1 in runs:
                    e = int(edge[:, a0:a1].sum())
                    cl.append((a0, a1, 2 * float(ink[:, a0:a1].sum()) / e if e else 0.0, h, _slant(ink[:, a0:a1])))
                lines.append(cl)
            parts.append((s, lines, lead))
    # толщина — к высоте строки, но не меньше обычной: строка из одних заглавных и цифр («KM2/1», 3PR42 л. 110) ниже
    # обычной на выносные элементы, и её штрих выходил толще
    med_h = float(np.median([c[3] for _, lines, _ in parts for cl in lines for c in cl])) if parts else 1.0
    # порог — по словам не короче полутора высот строки: обрывки («1», «/») тоньше и занижали медиану
    vals = [c[2] / max(c[3], med_h) for _, lines, _ in parts for cl in lines for c in cl
            if c[2] > 0 and c[1] - c[0] >= 1.5 * c[3]]
    if len(vals) < 20:
        return None                         # мерить не по чему — признак OCR
    med = float(np.median(vals))
    # и по листу: лист, пропечатанный толще (вставка другой редакции), — порог по его медиане, но только вверх и не
    # больше чем на 6 %
    per = {}
    for s, lines, _ in parts:
        per.setdefault(s["page"], []).extend(c[2] / max(c[3], med_h) for cl in lines for c in cl
                                             if c[2] > 0 and c[1] - c[0] >= 1.5 * c[3])
    # лист, где почти всё полужирное (оглавление 3SZ19), порог не поднимает — иначе выделение на нём пропадало
    lift = lambda v: float(np.median(v)) if len(v) >= 30 and float(np.median(v)) < 1.06 * med else med
    page_thr = {pg: 1.1 * max(med, lift(v)) for pg, v in per.items()}
    out = {}
    for s, lines, lead in parts:
        thr = page_thr.get(s["page"], 1.1 * med)
        text = s["text"]
        alltoks = [(m.start(), m.end()) for m in re.finditer(r"\S+", text)]
        toks = alltoks[lead:]
        cl = [(c[0], c[1], c[2] / max(c[3], med_h), c[4] > 1.08 and c[1] - c[0] >= 1.5 * c[3])
              for line in lines for c in line]                # курсив — по словам от трёх знаков: у «A», «7» косые
        # слова ↔ пятна краски по длинам (пятно бывает из двух слов, слово — из двух пятен: кавычка „ в Courier);
        # и при равном числе: подряд по порядку сдвигало выделение на соседнее слово (3SZ19, «- ÉLÉSÍTENI»)
        flags = [None] * len(toks)
        for ti, ci in _align([b - a for a, b in toks], [c[1] - c[0] for c in cl]):
            w = sum(cl[j][1] - cl[j][0] for j in ci) or 1
            # заглавные и цифры (коды «RTK», «31BB18») — порог выше: прямые штрихи без закруглений выходят толще
            caps = all(not re.search(r"[a-záéíóöőúüű]", text[slice(*toks[i])]) for i in ti)
            t = thr * (1.25 / 1.1 if caps else 1.0)
            st = (sum(cl[j][1] - cl[j][0] for j in ci if cl[j][2] > t) * 2 > w,
                  sum(cl[j][1] - cl[j][0] for j in ci if cl[j][3]) * 2 > w)
            for i in ti:
                flags[i] = st
        if any(f is None for f in flags):
            # не сопоставились — по доле символов: пятна подряд, разрыв между ними — в один знак средней ширины
            sp = sum(c[1] - c[0] for c in cl) / max(1, sum(b - a for a, b in toks))
            axis, pos = [], 0.0
            for c in cl:
                axis.append((pos, pos + c[1] - c[0], (c[2] > thr, c[3])))
                pos += c[1] - c[0] + sp
            span, n = max(pos - sp, 1.0), max(len(text), 1)
            flags = []
            for a, b in toks:
                votes = [next((st for a0, a1, st in axis if a0 <= (a + (b - a) * (j + 0.5) / 4) / n * span < a1), None)
                         for j in range(4)]
                votes = [v for v in votes if v is not None]
                flags.append(tuple(bool(votes) and sum(v[q] for v in votes) * 2 > len(votes) for q in (0, 1)))
        flags = [list(f) for f in [(False, False)] * lead + flags]
        short = [len(re.findall(r"\w", text[a:b])) <= 3 for a, b in alltoks]
        longs = [i for i in range(lead, len(flags)) if not short[i]]
        for i in range(lead, len(flags)):                     # курсив короткого («Ha a», «és a») — от длинных соседей
            if short[i]:
                nb = [flags[j][1] for j in ([x for x in longs if x < i][-1:] + [x for x in longs if x > i][:1])]
                flags[i][1] = bool(nb) and all(nb)
        size = [len(re.findall(r"\w", text[a:b])) for a, b in alltoks]
        if sum(z for z, f in zip(size, flags) if f[0]) >= 0.65 * max(1, sum(size)):
            for f in flags[lead:]:                            # почти всё полужирное (у курсива штрих тоньше) — всё
                f[0] = True
        flags = [tuple(f) for f in flags]
        if not any(b or i for b, i in flags):
            continue
        pieces, pos = [], 0
        for (a, b), f in zip(alltoks, flags):
            if pieces and pieces[-1][1] == f:
                pieces[-1][0] += text[pos:b]
            elif pieces:
                pieces[-1][0] += text[pos:a]
                pieces.append([text[a:b], f])
            else:
                pieces.append([text[pos:b], f])
            pos = b
        pieces[-1][0] += text[pos:]
        out[s["id"]] = [(t, f[0], f[1]) for t, f in pieces]
    return out


def _align(lens, widths):
    """Слова (длины в знаках) ↔ куски краски (ширины): разбиение обоих рядов на одинаковое число групп подряд — группа
    слов к группе кусков (1–3 к 1–3), чтобы ширина группы кусков в знаках средней ширины была ближе к длине слов.
    [(номера слов, номера кусков)]; [] — не сопоставить."""
    n, m = len(lens), len(widths)
    if not n or not m:
        return []
    cw = sum(widths) / max(1, sum(lens))
    band = 10 + n // 5
    inf = float("inf")
    D = {(0, 0): (0.0, None)}
    steps = ((1, 1), (2, 1), (1, 2), (3, 1), (1, 3), (2, 2))
    for i in range(n + 1):
        for j in range(m + 1):
            if (i, j) not in D or abs(j - i * m / n) > band + 3:
                continue
            c0 = D[(i, j)][0]
            for a, b in steps:
                if i + a > n or j + b > m:
                    continue
                c = c0 + abs(sum(widths[j:j + b]) / cw - sum(lens[i:i + a])) + 0.5 * (a + b - 2)
                if c < D.get((i + a, j + b), (inf,))[0]:
                    D[(i + a, j + b)] = (c, (i, j))
    if (n, m) not in D:
        return []
    out, cur = [], (n, m)
    while cur != (0, 0):
        prev = D[cur][1]
        out.append((list(range(prev[0], cur[0])), list(range(prev[1], cur[1]))))
        cur = prev
    return out[::-1]


def _seg_bold(s, style):
    """Сегмент полужирный целиком: по картинке (OCR жирность шапки 3SZ19 не видел); не измерено — по OCR."""
    if style is None:
        return bool(s["loc"].get("bold"))
    return s["id"] in style and all(b or not re.search(r"\w", t) for t, b, _ in style[s["id"]])


def _merge(rules):
    """Отрезки одной линии (положение ±1,5 pt, разрыв ≤ 2 pt) — одним: [(положение, от, до)]."""
    out, group = [], []
    for r in sorted(rules) + [None]:
        if r is not None and group and r[0] - group[-1][0] <= 1.5:
            group.append(r)
            continue
        if group:
            pos = sum(g[0] for g in group) / len(group)
            cur = None
            for _, a, b in sorted(group, key=lambda g: g[1]):
                if cur is not None and a <= cur[1] + 2:
                    cur[1] = max(cur[1], b)
                else:
                    cur = [a, b]
                    out.append((pos, cur))
        group = [r] if r is not None else []
    return [(pos, c[0], c[1]) for pos, c in out]


def _rules(drawings):
    """Горизонтальные и вертикальные линии листа (как layout.rules, но по get_cdrawings — втрое быстрее):
    ([(y, x0, x1)], [(x, y0, y1)]). Линия — до 3,5 pt толщиной: двойная вертикаль 3PR42 (л. 145) — залитый
    прямоугольник в 3,1 pt, с порогом 2,5 pt столбец пропадал."""
    hs, vs = [], []
    white = lambda c: c is not None and min(c) > 0.95
    for dr in drawings:
        x0, y0, x1, y1 = dr["rect"]
        # белое на белом листе — не линия: заливка ячеек белым (НГТУ, 22 прямоугольника) давала сетку таблицы
        # поверх всего листа (17.10)
        t = dr.get("type") or ""
        if t and ("f" not in t or white(dr.get("fill"))) and ("s" not in t or white(dr.get("color"))):
            continue
        if y1 - y0 <= 3.5 and x1 - x0 >= 15:
            hs.append(((y0 + y1) / 2, x0, x1))
        elif x1 - x0 <= 3.5 and y1 - y0 >= 8:
            vs.append(((x0 + x1) / 2, y0, y1))
        else:
            for it in dr["items"]:
                if it[0] == "l":
                    (ax, ay), (bx, by) = it[1], it[2]
                    if abs(ay - by) < 1 and abs(ax - bx) >= 15:
                        hs.append((ay, min(ax, bx), max(ax, bx)))
                    elif abs(ax - bx) < 1 and abs(ay - by) >= 8:
                        vs.append((ax, min(ay, by), max(ay, by)))
                elif it[0] in ("re", "qu"):
                    # рамка четырёхугольником (бланки 30RDGT: штамп, рамка таблицы — «qu»; без них таблица листа не
                    # находилась, шапка таблицы уходила в колонтитул, 15 листов → 68, 20.10) — как прямоугольник,
                    # если стороны по осям
                    if it[0] == "qu":
                        q = it[1]
                        if abs(q[0][1] - q[1][1]) > 1 or abs(q[0][0] - q[2][0]) > 1:
                            continue                        # повёрнутый
                        rx0, ry0 = min(p[0] for p in q), min(p[1] for p in q)
                        rx1, ry1 = max(p[0] for p in q), max(p[1] for p in q)
                    else:
                        rx0, ry0, rx1, ry1 = it[1]
                    if ry1 - ry0 <= 3.5 and rx1 - rx0 >= 15:  # тонкие прямоугольники одним путём (сетка 3PR42, л. 145)
                        hs.append(((ry0 + ry1) / 2, rx0, rx1))
                    elif rx1 - rx0 <= 3.5 and ry1 - ry0 >= 8:
                        vs.append(((rx0 + rx1) / 2, ry0, ry1))
                    elif rx1 - rx0 >= 15 and ry1 - ry0 > 3.5:  # рамка ячейки (и залитая, как layout.rules) — 4 стороны
                        hs += [(ry0, rx0, rx1), (ry1, rx0, rx1)]
                        vs += [(rx0, ry0, ry1), (rx1, ry0, ry1)]
    return _merge(hs), _merge(vs)


def _cluster(vals, tol=2.0):
    out = []
    for v in sorted(vals):
        if out and v - out[-1][-1] <= tol:
            out[-1].append(v)
        else:
            out.append([v])
    return [sum(g) / len(g) for g in out]


def _tables(hs, vs):
    """Таблицы листа: связные наборы линий (горизонталь касается вертикали) с хотя бы одной внутренней вертикалью —
    рамка с текстом внутри таблицей не считается. [(xs, ys, горизонтали, вертикали)]."""
    rules = [("h", r) for r in hs] + [("v", r) for r in vs]
    parent = list(range(len(rules)))

    def find(i):
        while parent[i] != i:
            parent[i] = parent[parent[i]]
            i = parent[i]
        return i
    across = lambda h, v: h[1] - 2 <= v[0] <= h[2] + 2
    # вертикаль ячейки начинается в 3–4 pt под линейкой строки и доходит до следующей линейки — так у двух и больше
    # вертикалей под одной линейкой или у одной вертикали в двух и больше строках (ÜFK л. 82–84 текстовым слоем:
    # таблицы «Beavatkozások», «Ellenőrzési követelmények» без боковых рамок распадались на куски по одной линейке и
    # уходили рисунком, 20.10) — тоже касание; одиночная вертикаль рамки чертежа (S2t) и засечки графика с зазором
    # снизу (T4) — нет
    # …и доходит до линейки СТРОКИ — не короче той, под которой начинается: рамка поля ввода в ячейке бланка (код в
    # рамке, окошко значения — 30RDGT) тоже начинается в 3–4 pt под линейкой строки, но доходит до своей нижней кромки
    # (100 pt против 530 у строки) — её вертикали становились столбцами таблицы: 3 столбца → 7, 15 листов → 29
    # (стенд вёрстки без языка, 22.10)
    width = lambda h: h[2] - h[1]
    gap = [(k, v) for k, h in enumerate(hs) for v in vs if across(h, v) and 2 < v[1] - h[0] <= 4
           and any(across(h2, v) and abs(v[2] - h2[0]) <= 2 and width(h2) >= 0.8 * width(h) for h2 in hs)]
    per_h, per_x = Counter(k for k, _ in gap), Counter(round(v[0]) for _, v in gap)
    below = {}
    for k, v in gap:
        if per_h[k] >= 2 or per_x[round(v[0])] >= 2:
            below.setdefault(k, []).append(v)
    for i, (oi, a) in enumerate(rules):
        if oi != "h":
            continue
        for j, (oj, b) in enumerate(rules):
            if oj == "v" and (b[1] - 2 <= a[0] <= b[2] + 2 and across(a, b) or b in below.get(i, ())):
                parent[find(i)] = find(j)
    groups = {}
    for i, r in enumerate(rules):
        groups.setdefault(find(i), []).append(r)
    out = []
    for g in groups.values():
        gh, gv = [r for o, r in g if o == "h"], [r for o, r in g if o == "v"]
        if len(gh) < 2 or len(gv) < 2:
            continue
        x0, x1 = min(v[0] for v in gv), max(v[0] for v in gv)
        y0, y1 = min(v[1] for v in gv), max(v[2] for v in gv)
        # таблица без боковых рамок (ÜFK, скан: внутренние вертикали есть, крайних нет — шла текстом, 303 строки из
        # 303) — края по горизонталям, если от двух они одинаково выходят за крайние вертикали с обеих сторон
        hh = [h for h in gh if y0 - 4 <= h[0] <= y1 + 4]                    # запас — как у касания выше
        lo, hi = [h[1] for h in hh if h[1] < x0 - 15], [h[2] for h in hh if h[2] > x1 + 15]
        edges = [min(lo), max(hi)] if len(lo) >= 2 and len(hi) >= 2 and max(lo) - min(lo) <= 6 \
            and max(hi) - min(hi) <= 6 else []
        if edges:
            x0, x1 = edges
        if x1 - x0 < 60 or y1 - y0 < 15 or not any(x0 + 3 < v[0] < x1 - 3 for v in gv):
            continue
        ys = _cluster([h[0] for h in hh] + [y0, y1])
        ys = [y for i, y in enumerate(ys) if i == 0 or y - ys[i - 1] >= 5]        # двойная линия — одна
        xs = _cluster([v[0] for v in gv] + edges)
        if len(ys) >= 2 and len(xs) >= 3:
            out.append((xs, ys, gh, gv))
    return out


def _grid_cells(xs, ys, gh, gv):
    """Ячейки сетки по строкам: [(первый столбец, столбец за последним, линии ячейки)]; столбцы без вертикали между
    ними в этой строке — одна ячейка."""
    def hcov(y, a, b):
        return any(abs(h[0] - y) <= 2 and min(h[2], b) - max(h[1], a) >= 0.5 * (b - a) for h in gh)

    def vcov(x, a, b):
        return any(abs(v[0] - x) <= 2 and min(v[2], b) - max(v[1], a) >= 0.5 * (b - a) for v in gv)
    rows = []
    for r in range(len(ys) - 1):
        ya, yb = ys[r], ys[r + 1]
        cells, start = [], 0
        for c in range(1, len(xs)):
            if c == len(xs) - 1 or vcov(xs[c], ya, yb):
                a, b = xs[start], xs[c]
                cells.append((start, c, {"top": hcov(ya, a, b), "bottom": hcov(yb, a, b),
                                         "left": vcov(a, ya, yb), "right": vcov(b, ya, yb)}))
                start = c
        rows.append(cells)
    return rows


def _on_lines(r, segs):
    """Рамка почти целиком (≥ 80 %) на строках сегментов — это сама строка текста, а не рисунок: светлый цветной текст,
    не вычтенный из краски листа, или растр с «запечённым» текстом под живой строкой (эффект InDesign). Картинкой
    такая строка оставалась в DOCX на языке оригинала (брошюра HERCA, 22.10)."""
    import fitz
    lines = [fitz.Rect(f) for s in segs for f in (s["loc"].get("frags") or [s["loc"]["bbox"]])]
    return sum((r & q).get_area() for q in lines if r.intersects(q)) >= 0.8 * r.get_area()


def _raster(page, segs, figs, reused, infos):
    """Картинки текстового слоя мельче рисунка-кластера (_figures: от 50×40 pt) — логотип в таблице титула, подписи
    (Eskom: логотип 121×37 pt и подписи пропадали): вставленные растры от 20×15 pt в пределах текста тела листа, не
    повторённые на других листах (логотип шапки — колонтитул), не под текстом и не фон во весь лист."""
    import fitz
    if not segs:
        return []
    ytop = min(s["loc"]["bbox"][1] for s in segs) - 2
    ybot = max(s["loc"]["bbox"][3] for s in segs) + 2
    area = page.rect.get_area()
    rects = []
    for im in infos:
        r = fitz.Rect(im["bbox"]) & page.rect
        if r.width < 20 or r.height < 15 or r.get_area() > 0.5 * area or r.y1 < ytop or r.y0 > ybot \
                or (im["xref"] and reused[im["xref"]] > 1) or any(r.intersects(f) for f, _ in figs) \
                or any(r.intersects(q) for q in rects) or any(r.contains(fitz.Point(*_center(s))) for s in segs)                 or _on_lines(r, segs):
            continue
        rects.append(r)
    # картинки одного ряда (четыре подписи титула Eskom) — одной картинкой по месту: порознь шли столбиком
    # и выталкивали титул на второй лист
    rows = []
    for r in sorted(rects, key=lambda q: q.y0):
        k = next((i for i, w in enumerate(rows) if min(w[0].y1, r.y1) - max(w[0].y0, r.y0) > 0.5 * min(w[0].height,
                                                                                                         r.height)), None)
        if k is None:
            rows.append([fitz.Rect(r), r])
        else:
            rows[k] = [rows[k][0] | r] + rows[k][1:] + [r]
    out = []
    for row in rows:                # между картинками ряда текст (ФИО, «Approved by») — тогда порознь
        own = row[1:] if any(row[0].contains(fitz.Point(*_center(s))) for s in segs) else row[:1]
        out += [(r, {"png": page.get_pixmap(dpi=150, clip=r).tobytes("png"), "x0": r.x0, "x1": r.x1, "top": r.y0,
                     "bottom": r.y1}) for r in own]
    return out


def _figures(page, segs, drawings, scan=False):
    """Рисунки листа (fitz.Rect): что на листе не текст сегментов и не длинные линии (рамки, таблицы, сетка графика),
    крупнее 50×40 pt — ядро; к нему — пересекающая его графика (рамка графика), кроме рамок во весь лист и знаков
    распознанного текста. На скане (scan) — и мелкое, от 25×8 pt: рукописные подпись и дата в текст не попадают
    (segments: рукопись не трогается), их место — картинкой (титул 30RDGT: «Aláírás:» без подписи)."""
    import fitz
    import numpy as np
    import cv2
    k = 50 / 72
    pix = page.get_pixmap(dpi=50, colorspace=fitz.csGRAY)
    ink = (np.frombuffer(pix.samples, np.uint8).reshape(pix.height, pix.stride)[:, :pix.width] < 200).astype(np.uint8)
    for s in segs:
        x0, y0, x1, y1 = s["loc"]["bbox"]
        ink[max(0, int((y0 - 1) * k)):int((y1 + 1) * k) + 1, max(0, int((x0 - 1) * k)):int((x1 + 1) * k) + 1] = 0
    n = int(40 * k)
    lines = cv2.morphologyEx(ink, cv2.MORPH_OPEN, cv2.getStructuringElement(cv2.MORPH_RECT, (n, 1))) | \
        cv2.morphologyEx(ink, cv2.MORPH_OPEN, cv2.getStructuringElement(cv2.MORPH_RECT, (1, n)))
    rest = ink.copy()
    rest[lines > 0] = 0
    rest = cv2.morphologyEx(rest, cv2.MORPH_OPEN, np.ones((2, 2), np.uint8))       # точки, обрезки линий
    num, lab, st, _ = cv2.connectedComponentsWithStats(cv2.dilate(rest, np.ones((int(10 * k), int(10 * k)), np.uint8)))
    W, H = page.rect.width, page.rect.height
    boxes = [fitz.Rect(s["loc"]["bbox"]) + (-1, -1, 1, 1) for s in segs]
    graph = [fitz.Rect(d["rect"]) for d in drawings]
    graph = [q for q in graph if q.width < 0.75 * W and q.height < 0.75 * H and not any(b.contains(q) for b in boxes)]
    out = []
    for i in range(1, num):
        x, y, w, h, _ = st[i]
        if (w / k < 25 or h / k < 8 or int(rest[lab == i].sum()) < 40) if scan else \
                (w / k < 50 or h / k < 40 or int(rest[lab == i].sum()) < 60):
            continue
        r = fitz.Rect(x / k, y / k, (x + w) / k, (y + h) / k)
        for _ in range(6):
            big = r + (-6, -6, 6, 6)
            nr = fitz.Rect(r)
            for q in graph:
                if q.intersects(big):
                    nr |= q
            if nr == r:
                break
            r = nr
        if not any(r.intersects(o) for o in out):
            out.append(r)
    return out


def _on_rule(r, rules, long=40):
    """Фигура — кусок длинной линии листа: тонкая и лежит на линии от long pt. Рамку и столбцы таблицы Word рисует по
    строкам — десятки отрезков в 11 pt один над другим; _rules склеивает их в одну линию, а рисунком они были «схемой»
    (3PR10: оглавление и тело листа с рамкой; 30PR132GT: бланк на 37 листах — снимком листа вместо таблицы)."""
    hs, vs = rules
    if r.height <= 4.5 and any(abs(y - (r.y0 + r.y1) / 2) <= 2 and a - 2 <= r.x0 and r.x1 <= b + 2 and b - a >= long
                               for y, a, b in hs):
        return True
    return r.width <= 4.5 and any(abs(x - (r.x0 + r.x1) / 2) <= 2 and a - 2 <= r.y0 and r.y1 <= b + 2 and b - a >= long
                                  for x, a, b in vs)


def _vector_figures(page, pno, drawings, segs, masked, tabs, figs, rules=((), ())):
    """Скопления векторных фигур (от 6, не во весь лист, не в таблицах и не в найденных рисунках, не куски длинных
    линий — _on_rule), от 80×60 pt, с двумя и больше подписями внутри и без обычного текста, задетого краем:
    [(рамка, img)], img — область листа ({"clip", "page_number"}: сборка снимет её с переведённого PDF, надписи —
    переводом)."""
    import fitz
    W, H = page.rect.width, page.rect.height
    off = [fitz.Rect(s["loc"]["bbox"]) for s in masked if s not in segs]       # колонтитулы
    rs = [fitz.Rect(d["rect"]) + (-0.5, -0.5, 0.5, 0.5) for d in drawings]   # линия без толщины — не пустая рамка
    rs = [r for r in rs if r.width < 0.75 * W and r.height < 0.75 * H and (r.width > 0.5 or r.height > 0.5)
          and not _on_rule(r, rules)
          and not any(r.intersects(t) for t in tabs) and not any(r.intersects(f) for f in figs)
          and not any(r.intersects(o) for o in off)]
    groups = []
    for r in rs:                                            # 20 pt: надписи-рамки стоят в 18 pt от скобки (Eskom)
        hit = [g for g in groups if g[0].intersects(r + (-20, -20, 20, 20))]
        R, n = fitz.Rect(r), 1
        for g in hit:
            groups.remove(g)
            R |= g[0]
            n += g[1]
        groups.append((R, n))
    out = []
    for R, n in groups:
        if n < 6 or R.width < 80 or R.height < 60:
            continue
        boxes = [fitz.Rect(s["loc"]["bbox"]) for s in segs]
        # подпись — серединой в области (с запасом 10 pt) и не шире её; длинная строка текста, задетая краем, — помеха
        inside = [b for b in boxes if (R + (-10, -10, 10, 10)).contains((b.tl + b.br) / 2) and b.width <= R.width + 20]
        cut = [b for b in boxes if b.intersects(R) and b not in inside]
        if len(inside) < 2 or cut:
            continue
        for b in inside:
            R |= b
        R = (R + (-3, -3, 3, 3)) & page.rect
        out.append((R, {"clip": [R.x0, R.y0, R.x1, R.y1], "page_number": pno + 1, "x0": R.x0, "x1": R.x1, "top": R.y0,
                        "bottom": R.y1}))
    return out


def _scheme_figures(page, pno, drawings, segs, tabs, figs):
    """Схема распознанного листа (текст — контурами): стрелки-связки с короткими подписями — областью листа с
    переведённого PDF, как _vector_figures. Связки разнесены дальше 20 pt и соединяются через подписи (стрелка
    упирается в подпись, следующая начинается под ней) — подпись в 10 pt от фигуры склеивает группы, подписи одного
    ряда (до 60 pt между ними) и строки одной подписи друг под другом — тоже; контуры букв (в рамке сегмента) и
    шапка/подвал — не фигуры. Схема — от 6 фигур, из них от 3 тонких линий длиннее 15 pt, и вертикальные, и
    горизонтальные (список условий 3SZ19 л. 4 с метками в рамочках «[3/2]» и тире — без вертикалей, это текст; поля
    бланка — одни горизонтали), от 2 подписей, ни одна не абзац (до 120 знаков: OCR сливает подписи ряда в одну —
    «Turbina főgépész Primerköri főgépész Villamos művezető»), и ни одна строка не задета краем. 3PR42 л. 113 (кто где
    при расхолаживании) рассыпался на абзацы «Turbi Reaktor operátor», «na operátor» (20.10)."""
    import fitz
    W, H = page.rect.width, page.rect.height
    sb = [fitz.Rect(s["loc"]["bbox"]) for s in segs]
    rs = [fitz.Rect(d["rect"]) + (-0.5, -0.5, 0.5, 0.5) for d in drawings]
    rs = [r for r in rs if r.width < 0.75 * W and r.height < 0.75 * H and (r.width > 0.5 or r.height > 0.5)
          and r.y1 >= 0.14 * H and r.y0 <= 0.86 * H and not any((b + (-2, -2, 2, 2)).contains(r) for b in sb)
          and not any(r.intersects(t) for t in tabs) and not any(r.intersects(f) for f in figs)]
    items = [(True, r) for r in rs] + [(False, b) for b in sb]
    up = list(range(len(items)))

    def root(i):
        while up[i] != i:
            up[i] = up[up[i]]
            i = up[i]
        return i
    for i, (fa, a) in enumerate(items):
        for j in range(i + 1, len(items)):
            fb, b = items[j]
            if fa or fb:
                m = 20 if fa and fb else 10
                near = (a + (-m, -m, m, m)).intersects(b)
            else:                                               # подписи — одного ряда или строки одной подписи
                h = min(a.height, b.height)
                near = min(a.y1, b.y1) - max(a.y0, b.y0) >= 0.5 * h and max(a.x0, b.x0) - min(a.x1, b.x1) <= 60 \
                    or min(a.x1, b.x1) > max(a.x0, b.x0) and max(a.y0, b.y0) - min(a.y1, b.y1) <= 0.6 * h
            if near:
                up[root(i)] = root(j)
    comps = {}
    for i, it in enumerate(items):
        comps.setdefault(root(i), []).append(it)
    out = []
    for c in comps.values():
        shapes = [r for f, r in c if f]
        vert = sum(r.width <= 3.5 and r.height >= 15 for r in shapes)
        horz = sum(r.height <= 3.5 and r.width >= 15 for r in shapes)
        if len(shapes) < 6 or vert + horz < 3 or not vert or not horz or len(c) - len(shapes) < 2:
            continue
        R = fitz.Rect(shapes[0])
        for _, r in c:
            R |= r
        inside = [s for s, b in zip(segs, sb) if R.contains((b.tl + b.br) / 2)]
        if any(len(s["text"].strip()) > 120 for s in inside) or any(b.intersects(R) and s not in inside
                                                                     for s, b in zip(segs, sb)):
            continue
        R = (R + (-3, -3, 3, 3)) & page.rect
        out.append((R, {"clip": [R.x0, R.y0, R.x1, R.y1], "page_number": pno + 1, "x0": R.x0, "x1": R.x1, "top": R.y0,
                        "bottom": R.y1}))
    return out


def _runs_of(group, bold, font, size, mono, right=None):
    """Прогоны абзаца из сегментов одной строки: части — через табуляцию по месту оригинала (оглавление «1.4.» ⇥
    «Tengelyelmozdulás …» ⇥ «13»: память переводит части по отдельности, 16.10); полужирные слова (по картинке,
    _pixel_style) — отдельными прогонами; «|» (вертикаль таблицы, прочитанная OCR) с краёв — прочь."""
    runs = []
    for i, g in enumerate(group):
        # нет в выделенных — обычный: признак полужирности OCR неверен (3PR42 л. 145, строка «haladja meg …»
        # выходила полужирной); OCR — только если по картинке не мерили
        pieces = list((bold or {}).get(g["id"]) or [(g["text"], bool(g["loc"].get("bold")) if bold is None else False,
                                                     False)])
        pieces[0] = (re.sub(r"^\s*\|\s*", "", pieces[0][0]),) + tuple(pieces[0][1:])
        pieces[-1] = (re.sub(r"\s*\|\s*$", "", pieces[-1][0]),) + tuple(pieces[-1][1:])
        pieces = [p for p in pieces if p[0]]
        text = "".join(p[0] for p in pieces)
        if len({p[1:] for p in pieces if re.search(r"\w", p[0])}) <= 1:   # весь сегмент одним начертанием
            st = next((p[1:] for p in pieces if re.search(r"\w", p[0])), (False, False))
            pieces = [(text,) + st]
        for pc in pieces:                                   # цвет — четвёртым, у текстового слоя (_text_style)
            t, b, it = pc[:3]
            runs.append({"t": t, "font": font, "size": size, "b": b, "i": it, "color": pc[3] if len(pc) > 3 else "000000",
                         "body_mono": mono, **({"box": f"b{g['id']}"} if g["loc"].get("box") else {})})
        last = runs[-1]
        if i + 1 < len(group):
            last["t"] += "\t"
            # номер листа в конце строки оглавления — к правому полю; число посреди строки («Business Level: ⇥ 3»
            # титульных сведений Eskom) — на своё место
            if i + 2 == len(group) and re.fullmatch(r"\d{1,4}", group[-1]["text"].strip()) and (
                    right is None or group[-1]["loc"]["bbox"][2] > right - 40):
                last["rtab"] = True
            else:
                last["tabs"] = [group[i + 1]["loc"]["bbox"][0]]
    return runs


def _table_block(t, segs, pno, bold, font, size, mono, imgs=(), pua=None, fields=()):
    """Блок таблицы (как у docx_struct): ячейки — по сетке, текст ячейки — сегменты с центром в ней, картинки
    (подпись от руки в форме) — в ячейку своего центра, после текста."""
    xs, ys, gh, gv = t
    grid = _grid_cells(xs, ys, gh, gv)
    pics = {}
    for rect, img in imgs:
        cx, cy = (rect.x0 + rect.x1) / 2, (rect.y0 + rect.y1) / 2
        r = next((i for i in range(len(ys) - 1) if ys[i] <= cy <= ys[i + 1]), None)
        c = next((i for i in range(len(xs) - 1) if xs[i] <= cx <= xs[i + 1]), None)
        if r is not None and c is not None:
            ci = next(k for k, (a, z, _) in enumerate(grid[r]) if a <= c < z)
            pics.setdefault((r, ci), []).append(img)
    # поля ввода бланка: текст в поле — в рамке (граница знаков), пустое поле — неразрывные пробелы его ширины в рамке
    inbox = lambda s, f: (f[0] - 1 <= (s["loc"]["bbox"][0] + s["loc"]["bbox"][2]) / 2 <= f[2] + 1
                          and f[1] - 1 <= (s["loc"]["bbox"][1] + s["loc"]["bbox"][3]) / 2 <= f[3] + 1)
    segs = [dict(s, loc=dict(s["loc"], box=True)) if any(inbox(s, f) for f in fields) else s for s in segs]
    for k, f in enumerate(fields):
        if not any(inbox(s, f) for s in segs):
            segs.append({"id": -10 ** 6 - 1000 * pno - k, "page": pno, "text": "\u00a0" * max(2, round((f[2] - f[0])
                         / (0.25 * size))), "loc": {"bbox": list(f), "frags": [list(f)], "box": True}})
    put = {}
    for s in segs:
        x0, y0, x1, y1 = s["loc"]["bbox"]
        cx, cy = (x0 + x1) / 2, (y0 + y1) / 2
        r = next((i for i in range(len(ys) - 1) if ys[i] <= cy <= ys[i + 1]), None)
        c = next((i for i in range(len(xs) - 1) if xs[i] <= cx <= xs[i + 1]), None)
        if r is None or c is None:
            continue
        ci = next(k for k, (a, z, _) in enumerate(grid[r]) if a <= c < z)
        put.setdefault((r, ci), []).append(s)
    # строка сетки без внутренних горизонталей, где в первом столбце — несколько меток пунктов («1.1.», «1.2.» … формы
    # RDGT: номер, условие и графа отметки — каждый против своего), — подстроками по верху меток: всё строкой шло
    # в одну ячейку, и номера, текст и отметки съезжали друг относительно друга (31RDGT л. 1, 20.10). Над первой
    # меткой — своя подстрока (легенда, заголовок раздела); линий между подстроками нет
    label = re.compile(r"\(?(?:\d+(?:\.\d+)*\.?|[a-zA-Z][.)])\)?|\d{2}[A-Z]{2}\d{2}[A-Z]\d{3}")     # номер пункта или код оборудования
    bands = []
    for r, cells in enumerate(grid):
        tops = sorted({round(s["loc"]["bbox"][1], 1) for s in put.get((r, 0), [])
                       if label.fullmatch(s["text"].strip())})
        if len(cells) >= 2 and len(tops) >= 2:
            above = any(s["loc"]["bbox"][3] <= tops[0] for ci in range(len(cells)) for s in put.get((r, ci), []))
            cuts = [ys[r]] + [t - 2 for t in (tops if above else tops[1:]) if t - 2 > ys[r] + 1] + [ys[r + 1]]
            for k in range(len(cuts) - 1):
                bands.append((r, cuts[k], cuts[k + 1], k == 0, k == len(cuts) - 2))
        else:
            bands.append((r, ys[r], ys[r + 1], True, True))
    rows = []
    for r, lo, hi, first, last in bands:
        cells = grid[r]
        row = []
        for ci, (a, z, e) in enumerate(cells):
            mid = lambda s: (s["loc"]["bbox"][1] + s["loc"]["bbox"][3]) / 2
            ss = sorted([s for s in put.get((r, ci), []) if lo <= mid(s) < hi or last and mid(s) >= hi],
                        key=lambda s: (round(s["loc"]["bbox"][1] / 3), s["loc"]["bbox"][0]))
            e = dict(e, **({} if first else {"top": False}), **({} if last else {"bottom": False}))
            paras, geom = [], None
            for group in _lines(ss):
                g0 = group[0]["loc"]["bbox"]
                gg = {"page": pno + 1, "top": g0[1], "bottom": max(g["loc"]["bbox"][3] for g in group), "x0": g0[0],
                      "x1": max(g["loc"]["bbox"][2] for g in group), "first_x0": g0[0], "size": size,
                      "lines": max(len(g["loc"].get("frags") or [1]) for g in group)}
                runs = _runs_of(group, bold, font, size, mono)
                for ru in runs if pua else ():               # значки частной области — знаком по виду (_pua_glyph)
                    ru["t"] = "".join(pua.get(ch, ch) for ch in ru["t"])
                paras.append({"k": "p", "runs": runs, "geom": gg, "seg": True})
                geom = gg if geom is None else dict(geom, x0=min(geom["x0"], gg["x0"]), x1=max(geom["x1"], gg["x1"]),
                                                     bottom=gg["bottom"], lines=geom["lines"] + gg["lines"])
            for img in (pics.get((r, ci), []) if first else []):
                paras.append({"k": "p", "runs": [{"img": img}], "seg": True})
            row.append({"blocks": paras, "th": False, "geom": geom, "edges": e, "span": z - a})
        rows.append(row)
    ys = [bands[0][1]] + [b[2] for b in bands] if bands else ys     # линии строк — по подстрокам
    head = [p for c in rows[0] for p in c["blocks"]] if rows else []
    # картинка в ячейке шапки (значок, JPDS текстовым слоем) — не текст: KeyError 't' ронял сборку (18.10)
    if head and len(rows) > 1 and all(r.get("b") for p in head for r in p["runs"] if r.get("t", "").strip()):
        for c in rows[0]:
            c["th"] = True                                  # шапка — первая строка, если вся полужирная
    # линии строк: строки таблицы листа и есть промежутки между ними — высота строки не меньше оригинала (docx_build)
    return {"k": "table", "rows": rows, "borders": True,
            "frame": {"xs": [round(x, 1) for x in xs], "ys": [round(y, 1) for y in ys], "row_lines": True},
            "geom": {"page": pno + 1, "top": ys[0], "bottom": ys[-1], "x0": xs[0], "x1": xs[-1], "first_x0": xs[0],
                     "size": size}}


def _grids(items, bold, font, size, mono, sign=None):
    """Сетка без линеек — таблицей без рамок: подряд от трёх строк из нескольких коротких кусков, куски которых встают в
    общие столбцы (перечень кодов 3PR42 л. 112: «04.TH10D001 ⇥ KM1/1 ⇥ 02.TK42D001 ⇥ …» в шесть столбцов). Строкой с
    табуляцией по месту оригинала последний столбец не влезал и переносился (кегль сборки крупнее), 17.10. Строка
    оглавления (в конце — номер листа) — не сетка. sign(строка) — знак маркера пункта (тот же факт, что у роли строки)."""
    from layout import code_line

    def row_ok(it):
        g = it[2]
        # строка — пункт со знаком маркера («– RN31S202 NYIT, RN35S201 ZÁR,» 3SZ19, куски через промежуток кода) — не
        # строка сетки: три пункта шли мелкой таблицей без рамок. По знаку на листе (факт), а не по метке портала
        # «пункт»: портал метит пунктами и строки перечня кодов 3PR42 («30RA04S201 - 37FY22 - BV IAB»), и перечень шёл
        # лесенкой абзацев (19.10); сетка (столбцы на листе) старше метки портала, знак маркера — старше сетки
        return it[1] == 0 and len(g) >= 2 and len({x["page"] for x in g}) == 1 \
            and all(len(x["text"]) <= 60 for x in g) and not re.fullmatch(r"\d{1,4}", g[-1]["text"].strip()) \
            and not (sign and sign(g))
    out, i = [], 0
    while i < len(items):
        j = i
        while j < len(items) and row_ok(items[j]) and items[j][2][0]["page"] == items[i][2][0]["page"]:
            j += 1
        run = items[i:j] if j - i >= 3 else []
        # столбец — края в пределах 20 pt (код «30BY07» под «04.TH10D001» сдвинут на 15 pt — тот же столбец)
        cols = _cluster([x["loc"]["bbox"][0] for it in run for x in it[2]], 20) if run else []
        col = lambda x: min(range(len(cols)), key=lambda k: abs(cols[k] - x["loc"]["bbox"][0]))
        good = run and 2 <= len(cols) <= 8 and all(
            len({col(x) for x in it[2]}) == len(it[2]) and all(abs(cols[col(x)] - x["loc"]["bbox"][0]) <= 20 for x in it[2])
            for it in run) and all(sum(any(col(x) == k for x in it[2]) for it in run) >= 2 for k in range(len(cols))) \
            and (len(cols) >= 3 or all(code_line(x["text"]) for it in run for x in it[2]))   # «подпись ⇥ текст» — не сетка
        if not good:
            out.append(items[i])
            i += 1
            continue
        pno = run[0][2][0]["page"]
        right = max(x["loc"]["bbox"][2] for it in run for x in it[2])
        xs = [round(c, 1) for c in cols] + [round(right + 4, 1)]
        rows = []
        for it in run:
            by = {col(x): x for x in it[2]}
            row = []
            for k in range(len(cols)):
                x = by.get(k)
                if x is None:
                    row.append({"blocks": [], "th": False, "geom": None, "span": 1,
                                "edges": dict.fromkeys(("top", "bottom", "left", "right"), False)})
                    continue
                b0 = x["loc"]["bbox"]
                # кегль — свой (перечень набран мельче текста: 9 pt против 11): таблица — мелким стилем, не переносится
                gg = {"page": pno + 1, "top": b0[1], "bottom": b0[3], "x0": b0[0], "x1": b0[2], "first_x0": b0[0],
                      "size": min(size, x["loc"].get("size") or size), "lines": len(x["loc"].get("frags") or [1])}
                row.append({"blocks": [{"k": "p", "runs": _runs_of([x], bold, font, size, mono), "geom": gg, "seg": True}],
                            "th": False, "geom": gg, "edges": dict.fromkeys(("top", "bottom", "left", "right"), False), "span": 1})
            rows.append(row)
        top = min(x["loc"]["bbox"][1] for x in run[0][2])
        bot = max(x["loc"]["bbox"][3] for x in run[-1][2])
        # кегль, при котором самый длинный кусок встаёт в свой столбец (знак ≈ 0,55 кегля, поля ячейки — по 1 pt):
        # кегль распознавания у кодов завышен (12 при 10 на листе), и ячейки переносились
        fit = min(((xs[col(x) + 1] - xs[col(x)] - 3) / (0.55 * len(x["text"])) for it in run for x in it[2]
                   if x["text"].strip()), default=size)
        out.append((run[0][0], 1, {"k": "table", "rows": rows, "borders": False, "frame": {"xs": xs}, "grid": True,
                                   "font_size": max(7.0, int(fit * 2) / 2) if fit < size else None,
                                   "geom": {"page": pno + 1, "top": top, "bottom": bot, "x0": cols[0], "x1": right,
                                            "first_x0": cols[0], "size": size}}))
        i = j
    return out


def _center(s):
    x0, y0, x1, y1 = s["loc"]["bbox"]
    return (x0 + x1) / 2, (y0 + y1) / 2


def _form_marks(drawings):
    """Отметки бланка, нарисованные векторами, а не знаком шрифта: кружок (контур из дуг) — «○», квадратик — «□»;
    6–20 pt, почти квадрат, только обводка. Графа «Teljesítés igazolás» форм 30RDGT — кружки кривыми: в DOCX графа
    выходила пустой (20.10). → [(рамка, знак)]."""
    out = []
    white = lambda c: c is not None and min(c) > 0.95
    for dr in drawings:
        x0, y0, x1, y1 = dr["rect"]
        w, h, t = x1 - x0, y1 - y0, dr.get("type") or ""
        if ("s" not in t or "f" in t and not white(dr.get("fill")) or not (6 <= w <= 20 and 6 <= h <= 20)
                or abs(w - h) > 2.5):
            continue
        ks = [it[0] for it in dr["items"]]
        if len(ks) >= 4 and all(k == "c" for k in ks):
            out.append(((x0, y0, x1, y1), "○"))
        elif ks in (["re"], ["qu"]) or len(ks) == 4 and all(k == "l" for k in ks):
            out.append(((x0, y0, x1, y1), "□"))
    if out:
        # внутри — другой путь (галочка, крестик, точка): отметка поставлена — пустой «□» говорил обратное (ред-тим 20.10)
        small = [o["rect"] for o in drawings if o["rect"][2] - o["rect"][0] <= 20 and o["rect"][3] - o["rect"][1] <= 20]
        inner = lambda b, r: (b[0] - 1 <= r[0] and r[2] <= b[2] + 1 and b[1] - 1 <= r[1] and r[3] <= b[3] + 1
                              and r[2] - r[0] < 0.9 * (b[2] - b[0]))
        out = [(b, {"○": "◉", "□": "☒"}[ch] if any(inner(b, r) for r in small) else ch) for b, ch in out]
    return out


def _in_cell(b, xs, ys, whole=False):
    """Рамка b — внутри одной ячейки сетки xs × ys, не на её линиях: клетка таблицы, нарисованная своим
    прямоугольником (график по дням, 16×16 pt), — не отметка бланка (ред-тим 20.10). whole — и не во всю ячейку:
    внутренняя рамка ячейки «двойной» таблицы (отступ до 3 pt со всех сторон) — край ячейки, а не поле ввода."""
    for c in range(len(xs) - 1):
        for q in range(len(ys) - 1):
            pad = (b[0] - xs[c], xs[c + 1] - b[2], b[1] - ys[q], ys[q + 1] - b[3])
            if min(pad) > 1:
                return not whole or max(pad) > 3
    return False


def _pua_glyph(page, bbox):
    """Знак частной области (U+F020–F0FF: значок Wingdings, вшитый в шрифт текста, — «» графы отметок 30RDIT) —
    по виду на листе: пустой квадрат «□», залитый «■», кружок «○»/«●»; иначе None. Знаком частной области он печатался
    пустым местом: шрифт текста этого знака не знает, а исходный шрифт по PDF не узнать (20.10)."""
    import fitz
    import numpy as np
    pix = page.get_pixmap(dpi=300, clip=fitz.Rect(bbox), colorspace=fitz.csGRAY)
    a = np.frombuffer(pix.samples, np.uint8).reshape(pix.height, pix.stride)[:, :pix.width] < 128
    ys, xs = np.nonzero(a)
    if len(xs) < 10:
        return None
    a = a[ys.min():ys.max() + 1, xs.min():xs.max() + 1]
    h, w = a.shape
    if min(h, w) < 8 or not 0.7 < w / h < 1.4:
        return None
    k = max(2, min(h, w) // 5)
    inner = a[k:-k, k:-k].mean()
    corner = (a[:k, :k].mean() + a[:k, -k:].mean() + a[-k:, :k].mean() + a[-k:, -k:].mean()) / 4
    square = corner > 0.3
    if inner < 0.15:
        return "□" if square else "○"
    if inner > 0.8:
        return "■" if square else "●"
    return None


def _form_fields(drawings):
    """Поля ввода бланка: обведённый прямоугольник (обводка «re»/«qu», без заливки краской) высотой 8–24 pt и шириной
    от полутора высот до 300 pt — рамка вокруг кода или пустое окошко для значения (30RDGT: «[31RD31S201] [   ]»).
    → [(x0, y0, x1, y1)]."""
    out = []
    white = lambda c: c is not None and min(c) > 0.95
    for dr in drawings:
        t = dr.get("type") or ""
        if ("s" not in t or "f" in t and not white(dr.get("fill")) or len(dr["items"]) != 1
                or dr["items"][0][0] not in ("re", "qu")):
            continue
        x0, y0, x1, y1 = dr["rect"]
        if 8 <= y1 - y0 <= 24 and 1.4 * (y1 - y0) <= x1 - x0 <= 300:
            out.append((x0, y0, x1, y1))
    return out


def _page_rules(doc, text=False):
    """Линии и рамки листов один раз на документ: (рамки путей, [(горизонтали, вертикали)] без рамки листа, [скан ли
    лист], [отметки и поля бланка], [низ рамки шапки или None]). text — текстовый слой (горизонтали, повторённые на
    листах, не снимаются)."""
    import fitz
    import numpy as np
    from layout import rules_from_image
    # от путей листа нужны линии (здесь) и рамки (для рисунков): все пути всех листов держали в памяти до 6,6 ГБ —
    # у 3PR42 текст нарисован контурами, тысячи путей на лист (17.10)
    draw, per, scans, marks = [], [], [], []
    for page in doc:
        d = page.get_cdrawings()
        # «solid» — залитая фигура не белым (полоса графика, ромб вехи, плашка) — для таблиц-графиков (_chart_table)
        draw.append([{"rect": x["rect"], "solid": "f" in (x.get("type") or "") and x.get("fill") is not None
                      and min(x["fill"]) < 0.95} for x in d])
        hs, vs = _rules(d)
        marks.append((_form_marks(d), _form_fields(d)))
        del d
        scan = is_scan(page)                                # отдельный скан в документе: линий-векторов нет
        if scan:
            pix = page.get_pixmap(dpi=200, colorspace=fitz.csGRAY)
            ih, iv = rules_from_image(np.frombuffer(pix.samples, np.uint8).reshape(pix.height, pix.stride)[:, :pix.width]
                                      .copy(), 200)
            hs = _merge(hs + [(y, a, b) for a, b, y in ih])
            vs = _merge(vs + [(x, a, b) for a, b, x in iv])
        per.append((hs, vs))
        scans.append(scan)
    cnt = Counter()
    key = lambda o, r: (o, round(r[0] / 3), round(r[1] / 6), round(r[2] / 6))
    for hs, vs in per:
        for k in {key("h", r) for r in hs} | {key("v", r) for r in vs}:
            cnt[k] += 1
    need = max(2, 0.5 * len(doc))

    def running(o, r):                                      # рамка листа, рамка названия раздела
        _, p, a, b = key(o, r)
        return sum(cnt.get((o, p + i, a + j, b + q), 0)
                   for i in (-1, 0, 1) for j in (-1, 0, 1) for q in (-1, 0, 1)) >= need
    lines = []
    for hs, vs in per:
        # текстовый слой: линии, повторённые на листах, не снимаются — ни горизонтали (верх таблицы титула Eskom на высоте
        # линии шапки: таблица не находилась), ни вертикали: бланк с одной и той же сеткой на каждом листе (30PR132GT,
        # 30PR25GT: столбцы «Sorsz. | Megnevezés | Ellenőrzés» на 37 листах) лишался столбцов вместе с рамкой, таблицы
        # не было, а линии без неё уходили снимком листа (_vector_figures); сама по себе рамка листа таблицы не даёт —
        # нужна внутренняя вертикаль (_tables). Лист-скан (линии по картинке) — как прежде: рамку листа и рамку названия
        # раздела снимают, оставляя вертикали, которых три и больше на одном пролёте (формы RDGT, 20.10)
        if not text:
            hs = [r for r in hs if not running("h", r)]
            span = Counter((round(r[1] / 6), round(r[2] / 6)) for r in vs if running("v", r))
            vs = [r for r in vs if not running("v", r) or span[(round(r[1] / 6), round(r[2] / 6))] >= 3]
        lines.append((hs, vs))
    # низ рамки шапки (рамка шапки и названия раздела 3SZ19, 3PR42) — что начинается ниже него, то тело листа: нижняя
    # горизонталь верхней пятой части листа той же ширины, что повторяющиеся на листах горизонтали рамки (от половины
    # листа); сама она может и не повторяться — под названием раздела в две строки рамка ниже (3SZ19: 119 вместо 110)
    frames = []
    for pg, (hs, _) in zip(doc, per):
        band = [r for r in hs if r[0] < 0.2 * pg.rect.height and r[2] - r[1] >= 0.5 * pg.rect.width]
        run = [r for r in band if running("h", r)]
        frames.append(max((r[0] for r in band if any(abs(r[1] - q[1]) < 6 and abs(r[2] - q[2]) < 6 for q in run)),
                          default=None))
    return draw, lines, scans, marks, frames


def _head_rows(doc, rules, body, keys, H):
    """Строки заголовков столбцов таблиц: первая строка сетки от двух столбцов, под которой в той же таблице есть текст
    тела (не повтор колонтитула). Повторяется на каждом листе (Word печатает шапку таблицы на каждом листе) — и
    снималась как колонтитул или бегущее название: у 30RDGT шапка таблицы уходила в шапку DOCX узкой сеткой, 15 листов
    → 68; у 31RDGT заголовки столбцов пропадали со второго листа (20.10). → {id сегмента}."""
    import fitz
    per = rules[1]
    by_page = {}
    for s in body:
        by_page.setdefault(s["page"], []).append(s)
    sure, maybe = set(), []
    for pno, (hs, vs) in enumerate(per):
        segs = by_page.get(pno, [])
        for xs, ys, _, _ in _tables(hs, vs):
            if len(xs) < 3 or len(ys) < 3:
                continue
            head = fitz.Rect(xs[0], ys[0], xs[-1], ys[1])
            rest = fitz.Rect(xs[0], ys[1], xs[-1], ys[-1])
            body_ = [s for s in segs if rest.contains(fitz.Point(*_center(s))) and not _is_running(s, keys, H)]
            if not body_:
                continue                                    # штамп: вся таблица повторяется на листах — колонтитул
            ids = [s for s in segs if head.contains(fitz.Point(*_center(s)))]
            # тело таблицы — и вне зон колонтитулов; таблица целиком в зоне шапки или подвала — штамп листа (ниже первой
            # строки меняется текст: номер, раздел, мусор слоя «���» разной длины), если её первая строка стоит на том
            # же месте и на других листах; тогда шапка таблицы — только если та же строка — шапка таблицы с телом на
            # другом листе. Иначе штамп шёл в тело один раз, у остальных листов шапки не было (ред-тим 20.10);
            # продолжение таблицы в одну-две строки вверху листа — по-прежнему таблица
            if any(not _zone(s, H) for s in body_):
                sure |= {s["id"] for s in ids}
            else:
                maybe.append(ids)
    place = lambda s: (_sig(s["text"]), round(s["loc"]["bbox"][0] / 6), round(s["loc"]["bbox"][1] / 6))
    seen = Counter(place(s) for p in by_page.values() for s in p)
    sigs = {_sig(s["text"]) for p in by_page.values() for s in p if s["id"] in sure}
    return sure | {s["id"] for ids in maybe if ids and (all(_sig(s["text"]) in sigs for s in ids)
                                                         or not any(seen[place(s)] > 1 for s in ids)) for s in ids}


def _graphic_page(page, segs, figs):
    """Лист-плакат: рисунки (краска вне текста и линий, _figures) — больше половины листа и на них ≥ 30 % подписей
    листа, или рисунок рассыпан на ≥ 5 кусков при малом тексте (обложка, разворот брошюры: картинки под маской,
    фоновые плашки). Такой лист собирается не абзацами (подписи схемы вперемешку с текстом, картинка — обрывками:
    плакат «Как работает АЭС» 1 лист → 2, брошюра 12 → 17), а целиком с переведённого PDF (стенд вёрстки без языка,
    22.10)."""
    import fitz
    # блок-схема Word: каждая фигура — десятки картинок (градиент, тень), надписи лежат на них; рисунок по тени не
    # находится (текст сегментов вычтен из краски), схема уходила в таблицы и абзацы россыпью (30GF(135)51GT, 28.10)
    if segs:
        imgs = [fitz.Rect(i["bbox"]) for i in page.get_image_info()]
        if len(imgs) >= 60 and sum(1 for s in segs if any(r.contains(fitz.Point(*_center(s))) for r in imgs)) >= 0.5 * len(segs):
            return True
    rects = [r for r, _ in figs]
    if not rects or not segs:
        return False
    area = sum(r.get_area() for r in rects) / page.rect.get_area()
    inside = sum(1 for s in segs if any(r.contains(fitz.Point(*_center(s))) for r in rects)) / len(segs)
    return area >= 0.5 and inside >= 0.3 or len(rects) >= 5 and area >= 0.05 and len(segs) <= 20


def _textless(doc, pno, boxes):
    """Копия листа pno, где снят текст в рамках boxes (сегменты), — картинки и графика целы: рисунок текстового слоя
    вырезается без текста, перевод которого идёт абзацами (22.10)."""
    import fitz
    c = fitz.open()
    c.insert_pdf(doc, from_page=pno, to_page=pno)
    for b in boxes:
        c[0].add_redact_annot(fitz.Rect(b), fill=False)
    c[0].apply_redactions(images=fitz.PDF_REDACT_IMAGE_NONE, graphics=fitz.PDF_REDACT_LINE_ART_NONE,
                          text=fitz.PDF_REDACT_TEXT_REMOVE)
    return c


def _ocr_layer(page):
    """Текст листа — невидимый слой распознавания поверх скана (газета 1813 г.: весь текст в режиме 3), а не настоящий
    текст поверх фона-картинки во весь лист (брошюра из InDesign: фон и фото — картинками, текст — видимый). Лист-плакат
    (_graphic_page) — не для первого: снимок брал скан без перевода (22.10); второй — плакат (HERCA 16 листов → 32)."""
    kinds = [k for k, _ in page.get_bboxlog() if "text" in k]
    return bool(kinds) and kinds.count("ignore-text") >= 0.5 * len(kinds)


def _chart_table(t, draw):
    """Таблица-график: в её рамке — от трёх залитых фигур, не совпадающих с ячейками (полосы диаграммы Ганта, ромбы
    вех, столбики): сетка таблицы — только поле графика, и таблицей собиралась пустая сетка без полос (22.10)."""
    import fitz
    xs, ys = t[0], t[1]
    R = fitz.Rect(xs[0], ys[0], xs[-1], ys[-1])
    near = lambda v, grid: any(abs(v - g) <= 2 for g in grid)
    n = 0
    for d in draw:
        if not d.get("solid"):
            continue
        q = fitz.Rect(d["rect"])
        if q.width < 2 or q.height < 2 or not R.contains(q) or q.get_area() > 0.5 * R.get_area():
            continue
        if near(q.x0, xs) and near(q.x1, xs) and near(q.y0, ys) and near(q.y1, ys):
            continue                                        # заливка ячейки
        n += 1
    return n >= 3


def _regions(doc, content, bold, font, size, mono, mask=None, rules=None):
    """Таблицы и рисунки листов: {лист: [{"rect", "block", "ids"}]} — сегменты ids идут не абзацами, а в них."""
    import fitz
    draw, per, scans, marks = (rules or _page_rules(doc, mask is not None))[:4]
    # картинки на нескольких листах (логотип шапки) — только текстовому слою; get_images не раскодирует картинки
    reused = Counter(x for page in doc for x in {im[0] for im in page.get_images()}) if mask is not None else Counter()
    by_page, masks = {}, {}
    for s in content:
        by_page.setdefault(s["page"], []).append(s)
    for s in mask or ():                                    # текстовый слой: и текст колонтитулов — не рисунок
        masks.setdefault(s["page"], []).append(s)
    out = {}
    for pno, page in enumerate(doc):
        segs = by_page.get(pno, [])
        rects = _figures(page, masks.get(pno, segs) if mask else segs, draw[pno], scans[pno])
        # «рисунок», почти целиком лежащий на строках сегментов, — сама строка текста (светлый цветной текст, не
        # вычтенный из краски листа): картинкой она оставалась в DOCX на языке оригинала (брошюра HERCA, 22.10)
        rects = [r for r in rects if not _on_lines(r, segs)]
        # текстовый слой: рисунок — без текста сегментов, задевающих его краем (их перевод идёт абзацем; в картинке
        # оставался оригинал: брошюра HERCA, «akkor nincs szükség…» поверх плашки)
        hit = [s["loc"]["bbox"] for s in segs if any(fitz.Rect(s["loc"]["bbox"]).intersects(r) for r in rects)]             if mask is not None else []
        bare = _textless(doc, pno, hit) if hit else None
        figs = [(r, {"png": (bare[0] if bare else page).get_pixmap(dpi=150, clip=r).tobytes("png"), "x0": r.x0,
                     "x1": r.x1, "top": r.y0, "bottom": r.y1}) for r in rects]
        hs, vs = per[pno]
        # скан — лист, распознанный OCR (своего текстового слоя нет); лист с крупным рисунком (> ½ листа) и текстом слоя —
        # не скан: текст — абзацами (ред-тим 20.10)
        # таблица-график на треть листа и больше (диаграмма Ганта во весь лист) — лист целиком, как плакат: частями
        # (график рисунком, названия задач абзацами) строки задач расходились с полосами на три листа (22.10)
        charts = sum(fitz.Rect(t[0][0], t[1][0], t[0][-1], t[1][-1]).get_area() for t in _tables(hs, vs)
                     if _chart_table(t, draw[pno])) if mask is not None else 0
        if mask is not None and segs and (scans[pno] and any(s["loc"].get("ocr") for s in segs)
                                          or not (scans[pno] and _ocr_layer(page)) and _graphic_page(page, segs, figs)
                                          or charts >= 0.3 * page.rect.get_area()):
            # скан посреди текстового слоя (лист регистрации, титул с подписями 32RDGT) — листом с переведённого PDF,
            # как отдельный скан формы (16.10): по сегментам сетка, подписи и штампы бланка рассыпались, 8 листов → 11
            # (20.10). Колонтитулы листа (повтор шапки других листов) посреди тела — не в снимке: их печатает шапка
            # DOCX; сканы в начале (за титулом — одни сканы) — вводная часть своим разделом без шапки, снимок целиком
            R = fitz.Rect()
            for q in [fitz.Rect(s["loc"]["bbox"]) for s in segs] + [f for f, _ in figs] + \
                    [fitz.Rect(a, y, b, y) for y, a, b in hs] + [fitz.Rect(x, a, x, b) for x, a, b in vs]:
                R |= q
            lead = all(scans[q] for q in range(1, pno))
            off = [s for s in masks.get(pno, []) if s not in segs] if not lead else []
            Hp = page.rect.height
            if any(_zone(s, Hp) == "head" for s in off):
                R.y0 = max(R.y0, max(s["loc"]["bbox"][3] for s in off if _zone(s, Hp) == "head") + 2)
            if any(_zone(s, Hp) == "foot" for s in off):
                R.y1 = min(R.y1, min(s["loc"]["bbox"][1] for s in off if _zone(s, Hp) == "foot") - 2)
            R = (R + (-4, -4, 4, 4)) & page.rect
            img = {"clip": [R.x0, R.y0, R.x1, R.y1], "page_number": pno + 1, "x0": R.x0, "x1": R.x1, "top": R.y0,
                   "bottom": R.y1}
            out[pno] = [{"rect": R, "ids": {s["id"] for s in segs},
                         "block": {"k": "p", "runs": [{"img": img}], "figure": True, "seg": True,
                                   "geom": {"page": pno + 1, "top": R.y0, "bottom": R.y1, "x0": R.x0, "x1": R.x1,
                                            "first_x0": R.x0, "size": size}}}]
            continue
        if mask is not None:
            infos = page.get_image_info(xrefs=True)          # один раз на лист
            figs += _raster(page, segs, figs, reused, infos)
            # рисунок — ровно картинка PDF: её xref (сборка вставит скриншот с переведёнными надписями, shots)
            for r, img in figs:
                for im in infos:
                    q = fitz.Rect(im["bbox"])
                    if im["xref"] and (q & r).get_area() >= 0.9 * max(r.get_area(), q.get_area()):
                        img["xref"] = im["xref"]
            # схема из линий с подписями (дерево «Layout format» Eskom: линии короче порога краски — рисунок не
            # находился, подписи перемешивались между разделами, 19.10) — областью листа с переведённого PDF
            tabs = [fitz.Rect(t[0][0], t[1][0], t[0][-1], t[1][-1]) for t in _tables(hs, vs)]
            figs += _vector_figures(page, pno, draw[pno], segs, masks.get(pno, []), tabs, [f for f, _ in figs], per[pno])
        elif not scans[pno]:
            tabs = [fitz.Rect(t[0][0], t[1][0], t[0][-1], t[1][-1]) for t in _tables(hs, vs)]
            figs += _scheme_figures(page, pno, draw[pno], segs, tabs, [f for f, _ in figs])
        regs, used = [], set()
        for t in _tables(hs, vs):
            r = fitz.Rect(t[0][0], t[1][0], t[0][-1], t[1][-1])
            if any((f & r).get_area() > 0.8 * r.get_area() for f, _ in figs):
                continue                                    # сетка внутри рисунка (график) — часть рисунка
            if _chart_table(t, draw[pno]):
                # таблица-график (диаграмма Ганта) — рисунком с переведённого PDF, подписи — переводом
                R = (r + (-2, -2, 2, 2)) & page.rect
                figs.append((R, {"clip": [R.x0, R.y0, R.x1, R.y1], "page_number": pno + 1, "x0": R.x0, "x1": R.x1,
                                 "top": R.y0, "bottom": R.y1}))
                continue
            inside = [s for s in segs if r.contains(fitz.Point(*_center(s)))]
            mine = [i for i, (f, _) in enumerate(figs) if r.contains(fitz.Point((f.x0 + f.x1) / 2, (f.y0 + f.y1) / 2))]
            # отметки бланка кривыми (кружок, квадратик) — знаком в своей ячейке, если там нет такого же текста
            xs_, ys_ = t[0], t[1]
            mid = lambda b: fitz.Point((b[0] + b[2]) / 2, (b[1] + b[3]) / 2)
            glyphs = [{"id": -1 - 1000 * pno - k, "page": pno, "text": ch, "loc": {"bbox": list(b), "frags": [list(b)]}}
                      for k, (b, ch) in enumerate(marks[pno][0]) if r.contains(mid(b)) and _in_cell(b, xs_, ys_)
                      and not any(fitz.Rect(s["loc"]["bbox"]).contains(mid(b)) for s in inside)]
            pua = {}
            for s in inside:                                # одинокий значок частной области — по виду на листе
                ch = s["text"].strip()
                if len(ch) == 1 and "\uf020" <= ch <= "\uf0ff" and ch not in pua:
                    pua[ch] = _pua_glyph(page, s["loc"]["bbox"]) or ch
            # поля ввода — внутри одной ячейки сетки, не на её линиях и не во всю ячейку
            fields = [f for f in marks[pno][1] if _in_cell(f, xs_, ys_, whole=True)]
            if inside:
                used |= set(mine)
                regs.append({"rect": r, "ids": {s["id"] for s in inside},
                             "block": _table_block(t, inside + glyphs, pno, bold, font, size, mono,
                                                   [figs[i] for i in mine], pua, fields)})
        for i, (r, img) in enumerate(figs):
            if i in used or any(r.intersects(g["rect"]) for g in regs):
                continue
            ids = {s["id"] for s in segs if r.contains(fitz.Point(*_center(s)))}
            if ids and mask is not None and "png" in img:
                # текст листа внутри рисунка (цитата поверх фото брошюры, подписи схемы) — сегменты рисунка, абзацами они
                # не идут: рисунок — с переведённого PDF, как снимок листа; из оригинала в DOCX оставался текст
                # оригинала картинкой, а перевод пропадал (брошюра HERCA, проверка порталом 22.10)
                img = {"clip": [r.x0, r.y0, r.x1, r.y1], "page_number": pno + 1, "x0": r.x0, "x1": r.x1, "top": r.y0,
                       "bottom": r.y1}
            regs.append({"rect": r, "ids": ids,
                         "block": {"k": "p", "runs": [{"img": img}], "figure": True, "seg": True,
                                   "geom": {"page": pno + 1, "top": r.y0, "bottom": r.y1, "x0": r.x0, "x1": r.x1,
                                            "first_x0": r.x0, "size": size}}})
        if regs:
            out[pno] = regs
    return out


def load(src, segs, text=False):
    """Блоки распознанного PDF по сегментам. Возвращает (блоки, сведения), как docx_struct.load. text — PDF с
    текстовым слоем без разметки: начертание — по шрифтам слоя, рисунки — без текста колонтитулов, в сведениях —
    рамки сегментов тела по листам (для раскладки колонтитулов docx_struct.layout по знакам)."""
    import fitz
    doc = fitz.open(str(src))
    W, H = doc[0].rect.width, doc[0].rect.height
    body = [s for s in segs if "shot" not in s["loc"] and s["loc"].get("bbox") and s["text"].strip()]
    run_keys = _running(body, H, len(doc))
    rules = _page_rules(doc, text)
    # шапка таблицы (строка заголовков столбцов) повторяется на каждом листе, но это не колонтитул и не бегущее название
    heads = _head_rows(doc, rules, body, run_keys, H)
    pool = [s for s in body if s["id"] not in heads and (_is_running(s, run_keys, H) or _portal_running(s))]
    content = [s for s in body if _role({
        "table_head": s["id"] in heads, "repeat": _is_running(s, run_keys, H), "ocr": bool(s["loc"].get("ocr")),
        "label": s["loc"].get("role"), "stable": _portal_running(s) and _stable_running(s, pool, run_keys, H)})
        != "running"]
    # меняющееся название раздела в рамке (метка портала, но не на всех листах одинаково) — один раз, первым появлением:
    # OCR читает его на каждом листе чуть по-разному, и точное сравнение с предыдущим листом ниже оставляло копии
    # посреди раздела (3SZ19, 19.10); похожие строки верхней зоны — одна группа
    from difflib import SequenceMatcher
    titles = {_sig(s["text"]) for s in pool if _portal_running(s) and not _stable_running(s, pool, run_keys, H)}
    titles = [k for k in titles if len(k) >= 6]
    like = lambda k, q: SequenceMatcher(None, k, q).ratio() >= 0.8
    seen, drop = [], set()
    for s in sorted(content, key=lambda s: (s["page"], s["loc"]["bbox"][1])):
        k = _sig(s["text"])
        if _zone(s, H) != "head" or s["id"] in heads or not any(like(k, t) for t in titles):
            continue
        if any(like(k, q) for q in seen):
            drop.add(id(s))
        else:
            seen.append(k)
    content = [s for s in content if id(s) not in drop]
    # бегущее название раздела в рамке сверху листа (3PR42: «1.15. Átfolyás …» на каждом листе раздела) — то же, что
    # вверху предыдущего листа, — не текст (на всех листах сразу его не найти: у каждого раздела своё)
    # название — от 6 знаков с буквами: короткое значение («3», «N/A» титульных сведений Eskom) повторяется само по себе
    # и пропадало (17.10)
    # и не строка, что начинается ниже рамки шапки листа: это тело — «1. Ismérvek:» 3SZ19 стоит под рамкой на одном
    # месте почти на каждом листе (раздел — с нового листа) и снималась как повтор с прошлого листа на 39 листах из 68
    # (20.10); лист без рамки — как прежде
    frames = rules[4]
    under = lambda s: frames[s["page"]] is not None and s["loc"]["bbox"][1] >= frames[s["page"]] - 1
    titled = lambda s: s["id"] not in heads and not under(s) and len(re.sub(r"\W", "", s["text"])) >= 6 \
        and re.search(r"[^\W\d_]{3}", s["text"])
    # пробелы — без счёта: OCR читает название то с двойным пробелом после номера, то с одним («1.1.  Turbina …»
    # / «1.1. Turbina …», 3SZ19), и повтор с прошлого листа оставался заголовком посреди раздела (19.10)
    key = lambda t: re.sub(r"\s+", " ", _sig(t)).strip()
    top = {}
    for s in content:
        if s["loc"]["bbox"][3] < 0.2 * H:
            top.setdefault(s["page"], set()).add(key(s["text"]))
    # первое появление бегущего названия (то же — вверху следующего листа) — заголовок раздела, даже без метки портала:
    # «1.8. Turbina csapágyak rezgése …» 3SZ19 шло обычным абзацем, а со следующего листа снималось как повтор (19.10)
    # (название — полужирное или по центру: строка тела вверху листа, повторённая на следующем, — не заголовок)
    first = {(s["page"], key(s["text"])) for s in content if s["loc"]["bbox"][3] < 0.2 * H and titled(s)
             and (s["loc"].get("bold") or s["loc"].get("align") == "center")
             and key(s["text"]) in top.get(s["page"] + 1, ()) and key(s["text"]) not in top.get(s["page"] - 1, ())}
    content = [s for s in content if not (s["loc"]["bbox"][3] < 0.2 * H and titled(s)
                                          and key(s["text"]) in top.get(s["page"] - 1, ()))]
    # (факт листа, старше метки портала — _role: метка «текст» у такой строки заголовок не отменяет)
    content = [dict(s, loc=dict(s["loc"], title=True)) if (s["page"], key(s["text"])) in first else s
               for s in content]
    # и на первом листе раздела: название в рамке сверху, а ниже на том же листе — тот же текст заголовком
    # («1.3/1. SZBV kazetta leesése» дважды подряд, 3PR42)
    norm = lambda t: re.sub(r"\s+", " ", _sig(t))
    lower = {}                                  # лист → [(верх, текст)]; заголовок бывает и в верхней пятой части листа
    for s in content:
        lower.setdefault(s["page"], []).append((s["loc"]["bbox"][1], norm(s["text"])))
    content = [s for s in content if not (s["loc"]["bbox"][3] < 0.2 * H and titled(s) and any(
        y > s["loc"]["bbox"][3] and t == norm(s["text"]) for y, t in lower.get(s["page"], ())))]
    if not content:
        return [], {"width": W, "height": H, "pages": len(doc), "artifacts": {"chars": [], "imgs": []}}
    size = round(median(s["loc"].get("size") or 10 for s in content) * 2) / 2
    font = FAMILY.get(Counter(s["loc"].get("family") for s in content).most_common(1)[0][0], "Arial")
    left = sorted(s["loc"]["bbox"][0] for s in content)[len(content) // 10]
    right = sorted(s["loc"]["bbox"][2] for s in content)[len(content) * 9 // 10]    # правый край текста
    mono = font == "Courier New"                            # моноширинный основной шрифт — не код (docx_build.is_code)
    text_mode = text
    layer = {} if text else _layer(doc)                     # текстовый слой: маркеры — знаками текста
    content_all = content
    bold = _text_style(doc, body) if text else _pixel_style(doc, body, layer)   # и колонтитулы: шапка 3SZ19 — жирная
    regions = _regions(doc, content, bold, font, size, mono, mask=body if text else None, rules=rules)
    # рисунок из краски колонтитула (шапка с линией под ней на скане): колонтитул снят с текста, а его краска уходила
    # картинкой посреди текста следующего листа (сканы, 18.10) — рисунок без своих сегментов, накрывающий колонтитул, — прочь
    import fitz as _f
    off = {}
    for s_ in body:
        if s_ not in content:
            off.setdefault(s_["page"], []).append(_f.Rect(s_["loc"]["bbox"]))
    for pno, regs in regions.items():
        regs[:] = [g for g in regs if g["ids"] or not any((g["rect"] & r).get_area() > 0.5 * max(r.get_area(), 1)
                                                          for r in off.get(pno, []))]
    # порядок: абзацы — по сегментам (порядок чтения OCR), таблица или рисунок — на месте первого своего сегмента;
    # рисунок без сегментов — перед первым сегментом листа ниже его верха
    taken = {i for regs in regions.values() for g in regs for i in g["ids"]}
    flow = [s for s in content if s["id"] not in taken]
    marks = {}          # текстовый слой: маркер пункта — свой сегмент слева от текста пункта (layout.frags_from_text)
    if text:            # только в потоке: «-» в ячейке таблицы — значение (НГТУ)
        bullet = [s for s in flow if re.fullmatch(r"\s*[•▪◦·●○■□o–\-\uf020-\uf0ff]\s*", s["text"])]
        # маркер — если справа на той же строке ближе 40 pt есть пункт; иначе «-» — значение («Revision: ⇥ -»);
        # строка пункта — первая строка его сегмента: у пункта в несколько строк середина всего сегмента ниже маркера,
        # и маркер Symbol Eskom не узнавался (текст пункта — по ширине знака, а не на своём месте, 20.10)
        head = lambda t: (t["loc"].get("frags") or [t["loc"]["bbox"]])[0]
        lone = {s["id"] for s in bullet if any(
            t["page"] == s["page"] and t["id"] != s["id"] and 0 < head(t)[0] - s["loc"]["bbox"][2] < 40
            and abs((head(t)[1] + head(t)[3]) - (s["loc"]["bbox"][1] + s["loc"]["bbox"][3])) / 2
            < 0.6 * max(s["loc"]["bbox"][3] - s["loc"]["bbox"][1], 6) for t in flow)}
        for s in flow:
            if s["id"] in lone:
                marks.setdefault(s["page"], []).append((s["loc"]["bbox"], s["text"].strip()))
        flow = [s for s in flow if s["id"] not in lone]
    items = [(group[0]["id"], 0, group) for group in _lines(sorted(flow, key=lambda s: s["id"]))]
    for pno, regs in regions.items():
        page_ids = sorted((s["id"], s["loc"]["bbox"][1]) for s in content if s["page"] == pno)
        for g in regs:
            if g["ids"]:
                k = min(g["ids"])
            else:
                k = next((i - 0.5 for i, y in page_ids if y > g["rect"].y0), (page_ids[-1][0] + 0.5) if page_ids else
                         max((s["id"] for s in content if s["page"] < pno), default=-1) + 0.5)
            items.append((k, 1, g["block"]))
    items.sort(key=lambda x: (x[0], x[1]))
    def sign_of(group, text=None):
        """Знак маркера пункта в начале строки — факт листа: скан — тире текстового слоя (чёрточка 0–30 pt левее начала
        строки; текст нарисован контурами, маркеры — шрифтом), текстовый слой — знак в начале текста («•», «–», маркер
        Symbol) или свой сегмент-маркер слева на той же строке (marks), а не чёрточка краски. → {"ch", "dash" | "tm" |
        "mk"} или None. Тот же факт решает роль строки (_role) и строку сетки без линеек (_grids)."""
        g0 = group[0]
        f = (g0["loc"].get("frags") or [g0["loc"]["bbox"]])[0]
        lh = max(f[3] - f[1], 6)
        if not text_mode:
            d = next((d for d in layer.get(g0["page"], {}).get("dash", [])
                      if abs(d[1] - (f[1] + f[3]) / 2) < 0.6 * lh and 0 < f[0] - d[0] < 30), None)
            return {"ch": None, "dash": d} if d else None
        tm = SIGN.match(g0["text"] if text is None else text)
        if tm:
            return {"ch": tm.group(1), "tm": tm}
        mk = next(((bb, t) for bb, t in marks.get(g0["page"], ())
                   if abs((bb[1] + bb[3]) / 2 - (f[1] + f[3]) / 2) < 0.6 * lh and 0 < f[0] - bb[2] < 40), None)
        return {"ch": mk[1], "mk": mk} if mk else None
    items = _grids(items, bold, font, size, mono, sign_of)
    blocks, cur = [], None                                  # cur — открытый список из пунктов с тире
    for _, kind, group in items:
        if kind == 1:
            blocks.append(group)
            cur = None
            continue
        s = group[0]
        loc = s["loc"]
        x0, y0, x1, y1 = loc["bbox"]
        first = (loc.get("frags") or [loc["bbox"]])[0]
        lh = max(first[3] - first[1], 6)
        geom = {"page": s["page"] + 1, "top": y0, "bottom": max(g["loc"]["bbox"][3] for g in group), "x0": x0,
                "x1": max(g["loc"]["bbox"][2] for g in group), "first_x0": first[0], "size": size,
                "lines": max(len(g["loc"].get("frags") or [1]) for g in group),
                "last_x0": (loc.get("frags") or [loc["bbox"]])[-1][0]}
        runs = _runs_of(group, bold, font, size, mono, right)
        text = "".join(r["t"] for r in runs).strip()
        m = NUM.match(text)
        sign = sign_of(group, text)
        # роль строки — по старшинству источников (_role): факты листа, метка портала, правила вида строки
        # (вид заголовка: короткий полужирный по центру; в текстовом слое — и у левого края, если с номером: «2.1.1
        # Purpose»; на листе с метками портала — не применяется: правило брало и подписи «5. ábra …» по центру, и
        # полужирные «Korlátozó feltétel:», 18.10; портал находил 57 % заголовков при точности 0,88, правила — 16 %)
        role = _role({"ocr": bool(loc.get("ocr")), "label": loc.get("role"), "labeled": bool(loc.get("roles")),
                      "title": bool(loc.get("title")), "sign": bool(sign), "toc": _toc_line(group, right),
                      "pieces": len(group),
                      "typo": len(group) == 1 and len(runs) == 1 and runs[0]["b"] and len(text) < 120 and (
                          loc.get("align") == "center" or text_mode and m and len(loc.get("frags") or [1]) == 1
                          and re.search(r"[^\W\d_]{3}", text[m.end(1):])),
                      "num_label": bool(PLABEL.match(text)), "text": text})
        if role == "heading":
            # «колонтитул» по метке портала, что остался в тексте (_stable_running: меняется от раздела к разделу), —
            # название раздела в рамке: заголовок (3SZ19 «1.8. Turbina csapágyak…» шёл пунктом списка над разделом)
            # (повтор того же текста заголовка уровня прежде выбрасывался — бегущее название раздела в начале каждого
            # листа, 3SZ19; его снимают правила повтора выше, а заплатка выбрасывала и настоящие подзаголовки разделов:
            # «1. Ismérvek:» 3SZ19 — 7 раз, 20.10)
            level = min(len(m.group(1).split(".")) if m else 1, 5)
            blocks.append({"k": "h", "level": level, "runs": runs, "geom": geom, "center": loc.get("align") == "center"})
            cur = None
            continue
        para = {"k": "p", "runs": runs, "geom": geom, "seg": True}    # seg — абзац собран, merge_inline не клеит
        if role == "item":
            ch = sign["ch"] if sign else None
            if sign and sign.get("mk"):
                # отступ пункта — от маркера, как у пункта с маркером в тексте: у Eskom один пункт с «•» отдельным
                # куском считался ступенью глубже соседнего («o» вместо «•», 19.10)
                x0 = sign["mk"][0][0]
            if sign:
                kind = ("Dash" if ch in "–-" or "\uf020" <= ch <= "\uf0ff" and chr(ord(ch) - 0xF000) in "-–"
                        else "Disc") if ch else "Dash"
                # знак маркера оригинала: «o»/«◦» — кружок, «▪»/«■» — квадрат (Symbol и Wingdings — частная область)
                if ch and kind == "Disc" and ch in "o◦○\uf06f":
                    kind = "Circle"
                elif ch and kind == "Disc" and ch in "▪■□\uf0a7\uf0a8\uf0a1":
                    kind = "Square"
            else:
                # пункт по метке портала без знака маркера («a) …», «1. …», «ba) …»), на листе без меток — «a)», «1)» в
                # начале: пункт без маркера, метка — в тексте
                kind = "Label"
            if cur is None or abs(cur["x0"] - x0) > 6 or cur["kind"] != kind:
                # знак — с листа (текстовый слой): сборка ставит его на любом уровне, не «o» второго уровня (Eskom)
                cur = {"k": "list", "kind": kind, "items": [], "x0": x0, "own": bool(text_mode and ch)}
                if text_mode and ch:
                    # место маркера и текста по оригиналу (как docx_struct.list_geo): тире НГТУ стояло на уровне
                    # текста абзацев, в оригинале — на 14 pt правее (19.10)
                    # знак в тексте и пункт в несколько строк — место текста по второй строке (висячий отступ
                    # оригинала: «•⇥Normal Environmental …» Eskom — 22 pt от маркера, по ширине «• » выходило 6)
                    import fitz
                    fr = loc.get("frags") or []
                    tx = first[0] if sign.get("mk") else fr[1][0] if len(fr) > 1 and fr[1][0] > x0 + 3 else \
                        x0 + fitz.get_text_length(sign["tm"].group(0), fontsize=geom.get("size") or 10)
                    cur["geo"] = {"marker": x0, "text": round(tx, 1)}
                elif sign:
                    # скан: тире — чёрточка слоя (её правый край), текст — начало куска; по профилю текст пункта стоял на
                    # 60 pt правее тире (3PR42: +23 листа), в оригинале — почти вплотную (19.10)
                    cur["geo"] = {"marker": round(sign["dash"][0] - 0.5 * (geom.get("size") or 10), 1),
                                  "text": round(first[0], 1)}
                else:
                    # пункт без знака (метка — в тексте: «1.3/1. Падение кассеты…» оглавления 3PR42): метка — с начала
                    # строки, продолжение — с начала второй строки; по профилю пункты стояли на 155 pt от поля узким
                    # столбцом (19.10)
                    fr = loc.get("frags") or []
                    rest = fr[1][0] if len(fr) > 1 and fr[1][0] >= x0 else x0
                    cur["geo"] = {"marker": round(x0, 1), "text": round(rest, 1)}
                blocks.append(cur)
            cur["items"].append([para])
            continue
        # продолжение — правило склейки, младше портала: на листе с метками границы абзацев уже его (roles.apply), и
        # абзац, начатый порталом, к абзацу выше и к пункту выше не клеится (Р7); строка оглавления не продолжает ничего
        # (строки оглавления 3PR42 уходили внутрь «пункта» над ними)
        last = cur["items"][-1][-1] if cur is not None and cur["items"] else (blocks[-1] if blocks else None)
        if role == "body" and not loc.get("roles") and len(group) == 1 and last is not None \
                and _continues(last, s, lh):
            _join(last, runs, geom)
            continue
        # продолжение пункта (OCR разбил пункт на сегменты)
        if role == "body" and not loc.get("roles") and cur is not None and cur["items"] \
                and abs(first[0] - cur["x0"]) < 8 and loc.get("align") != "center":
            cur["items"][-1].append(para)
            continue
        cur = None
        indent = round(max(0.0, first[0] - left) / 9) * 9
        if indent >= 9 and loc.get("align") != "center":
            para["indent"] = min(indent, 90)                # отступ по оригиналу — уровнем стиля (сборщик)
        blocks.append(para)
    # номер раздела из рамки отдельной строкой («1.1.») прямо перед заголовком с тем же номером («1.1. Срабатывание
    # АЗ…»): название из рамки снято как повтор, номер оставался строкой-сиротой над каждым разделом (3PR42, 19.10)
    txt = lambda b: "".join(r.get("t", "") for r in b.get("runs") or []).strip()
    num = re.compile(r"\d+(?:\.\d+)*(?:/\d+)?\.?")
    blocks = [b for i, b in enumerate(blocks)
              if not (b["k"] in ("p", "h") and num.fullmatch(txt(b)) and i + 1 < len(blocks)
                      and blocks[i + 1]["k"] in ("p", "h") and txt(blocks[i + 1]).startswith(txt(b).rstrip(".")))]
    # уровень пункта — ступень его отступа среди отступов всех списков документа (3PR42: «-» у края рамки и «-»
    # под пунктом 2.1.1. — раньше один уровень, вложенность пропадала)
    xs = [b["x0"] for b in blocks if b["k"] == "list"]
    # ступень — если так начинаются хотя бы 5 % списков
    steps = [c for c in _cluster(xs, 8) if sum(abs(x - c) <= 8 for x in xs) >= max(3, 0.05 * len(xs))] or \
        _cluster(xs, 8)[:1]
    for b in blocks:
        if b["k"] == "list":
            x = b.pop("x0")
            b["level"] = min(min(range(len(steps)), key=lambda i: abs(steps[i] - x)), 3)
    import docx_struct
    # таблица через лист (шапка повторена на каждом листе или продолжение без неё) — одной таблицей, как у размеченного:
    # в DOCX она переносится сама и повторяет шапку; частями по листам каждая часть тянула пустоту до рамки листа
    # (31RDGT: 5 листов → 7, 20.10)
    blocks = docx_struct.join_tables(blocks)
    # разрывы листа — правилом PDF с разметкой (docx_struct.mark_breaks): верх и низ тела листов; и для распознанных
    # листов — разделы 3SZ19 с нового листа шли подряд (18.10)
    if True:
        tops, bots = {}, {}
        ys = [(s["page"] + 1, s["loc"]["bbox"][1], s["loc"]["bbox"][3]) for s in content_all]
        ys += [(pno + 1, g["rect"].y0, g["rect"].y1) for pno, regs in regions.items() for g in regs]
        for pn, a, b in ys:
            tops[pn], bots[pn] = min(tops.get(pn, a), a), max(bots.get(pn, b), b)
        docx_struct.mark_breaks(blocks, tops, bots)
    boxes = {}
    for s in content_all:                                   # рамки текста тела по листам (от 1): колонтитулы — прочее
        for f in s["loc"].get("frags") or [s["loc"]["bbox"]]:
            boxes.setdefault(s["page"] + 1, []).append(tuple(f))
    return blocks, {"width": W, "height": H, "pages": len(doc), "artifacts": {"chars": [], "imgs": []},
                    "font": font, "size": size, "style": bold, "content_boxes": boxes}


def _continues(prev, s, lh):
    """Сегмент — следующая строка того же абзаца: OCR разрезал абзац по строкам («… hogy a Duna vízhőmérséklete ne» /
    «haladja meg a 29,5 °C-ot …», 3PR42 л. 145; было два абзаца, второй — с отступом). Тот же лист и левый край,
    вплотную ниже, предыдущий не кончается точкой, двоеточием, табуляцией, следующий начинается строчной буквой или
    числом, но не номером пункта: «… a következő táblázat … alapján a» / «2.2 és 2.3 pontok …» (край — как у последней
    строки абзаца: у пункта «2.1.» первая строка левее)."""
    g = prev.get("geom") or {}
    if prev.get("k") != "p" or prev.get("figure") or g.get("page") != s["page"] + 1 or g.get("x0") is None:
        return False
    t = "".join(r.get("t", "") for r in prev["runs"]).rstrip()
    x0, y0 = s["loc"]["bbox"][0], s["loc"]["bbox"][1]
    head = s["text"].lstrip()
    start = head[:1].islower() or (head[:1].isdigit() and not re.match(r"\d+(?:\.\d+)*\.\s", head))
    edge = min(abs(x0 - g["x0"]), abs(x0 - g.get("last_x0", g["x0"])))
    # первая строка абзаца с красной строкой: вторая левее (до 40 pt), обе — до одного правого края (выключка по
    # ширине): «2.2.2. Текст ПЗ … написан акку-» / «ратно от руки …» (NSU, 18.10) — было два абзаца
    indent = (g.get("lines") or 1) == 1 and 0 < g["x0"] - x0 <= 40 and g.get("x1") is not None \
        and abs(s["loc"]["bbox"][2] - g["x1"]) < 3
    return bool(t) and t[-1] not in ".:;!?" and "\t" not in t and (edge < 6 or indent) and \
        -2 <= y0 - g["bottom"] < 0.8 * lh and start


SUSPENDED = re.compile(r"(?:és|vagy|illetve|ill\.|и|или|либо|and|or)\s", re.I)   # «fűtő- és …»: дефис — не перенос


def _join(prev, runs, geom):
    """Строку — в конец абзаца: через пробел (после переноса «bar-» — без него и без дефиса, как layout.join_lines:
    «акку-» + «ратно» печаталось «акку-ратно», NSU 18.10; дефис составного слова и после прописной или цифры — на
    месте), прогоны одного начертания — одним."""
    tail = prev["runs"][-1]
    head = (runs[0].get("t") or "") if runs else ""
    if tail["t"].endswith("-") and head[:1].islower() and not SUSPENDED.match(head) \
            and not re.search(r"[A-ZÁÉÍÓÖŐÚÜŰА-ЯЁ0-9]-$", tail["t"]) and not re.search(r"\w-\w+-$", tail["t"]):
        tail["t"] = tail["t"][:-1]
    elif not tail["t"].endswith(" ") and (not tail["t"].endswith("-") or SUSPENDED.match(head)):
        tail["t"] += " "                                    # «fűtő- és hűtő…» — дефис с пробелом перед союзом
    for r in runs:
        tail = prev["runs"][-1]
        if (tail.get("b"), tail.get("i")) == (r.get("b"), r.get("i")) and not tail.get("tabs") and not tail.get("rtab"):
            tail["t"] += r["t"]
        else:
            prev["runs"].append(r)
    g = prev["geom"]
    g["last_x0"] = geom.get("last_x0", g.get("last_x0"))
    g["bottom"] = max(g["bottom"], geom["bottom"])
    g["x1"] = max(g["x1"], geom["x1"])
    g["lines"] = (g.get("lines") or 1) + (geom.get("lines") or 1)


def _lines(segs):
    """Подряд идущие сегменты одной строки листа — группой, слева направо: тот же лист, середины первых строк — в
    пределах половины высоты, по ширине не налезают друг на друга. OCR отдаёт их и не по порядку: «Turbina
    üzemzavari …» раньше «1.1.» (оглавление 3SZ19, 16.10). Номер листа у перенесённого названия — на одной из
    следующих его строк: «Generátor teherledobás … / üresjáratig ⇥ 11»."""
    out = []
    mid = lambda f: (f[1] + f[3]) / 2
    for s in segs:
        fr = s["loc"].get("frags") or [s["loc"]["bbox"]]
        f = fr[0]
        if out and out[-1][-1]["page"] == s["page"]:
            grp = out[-1]
            pfr = grp[-1]["loc"].get("frags") or [grp[-1]["loc"]["bbox"]]
            pf = pfr[0]
            same = abs(mid(pf) - mid(f)) < 0.5 * max(pf[3] - pf[1], 6) or (
                len(fr) == 1 and re.fullmatch(r"\d{1,4}", s["text"].strip()) and any(
                    abs(mid(pl) - mid(f)) < 0.5 * max(pl[3] - pl[1], 6) and f[0] > pl[2] - 2 for pl in pfr[1:]))
            free = all(f[0] > q["loc"]["bbox"][2] - 2 or f[2] < q["loc"]["bbox"][0] + 2 for q in grp)
            if same and free:
                grp.append(s)
                continue
        out.append([s])
    return [sorted(g, key=lambda q: q["loc"]["bbox"][0]) for g in out]


def _grid(zsegs, bold=None):
    """Сетка колонтитула по сегментам одного листа: строки — по высоте, столбцы — по левым краям (±12 pt)."""
    rows = {}
    for s in sorted(zsegs, key=lambda s: (s["loc"]["bbox"][1], s["loc"]["bbox"][0])):
        key = next((k for k in rows if abs(k - s["loc"]["bbox"][1]) < 5), s["loc"]["bbox"][1])
        rows.setdefault(key, []).append(s)
    xs = []
    for x in sorted(s["loc"]["bbox"][0] for s in zsegs):
        if not xs or x - xs[-1] > 12:
            xs.append(x)
    right = max(s["loc"]["bbox"][2] for s in zsegs) + 6
    cols = [x - 4 for x in xs] + [right]
    out = {"rows": [], "widths": [cols[i + 1] - cols[i] for i in range(len(xs))], "aligns": ["left"] * len(xs),
           "inner": False, "box": True, "font": None, "size": None}
    for _, ss in sorted(rows.items()):
        cells = [""] * len(xs)
        for s in ss:
            ci = max(i for i, x in enumerate(xs) if s["loc"]["bbox"][0] >= x - 12)
            cells[ci] = (cells[ci] + " " + s["text"].strip()).strip()
        # полужирность — по текстовому слою: шапка 3SZ19 набрана жирным, OCR этого не видел
        out["rows"].append({"cells": cells, "bold": all(_seg_bold(s, bold) for s in ss)})
    # столбец из одних чисел («Oldalszám: | 41») — к предыдущему: узкий столбец разносил номер листа по строкам
    for ci in range(len(xs) - 1, 0, -1):
        if all(re.fullmatch(r"\d{1,4}", r["cells"][ci]) or not r["cells"][ci] for r in out["rows"]) \
                and any(r["cells"][ci] for r in out["rows"]):
            for r in out["rows"]:
                r["cells"][ci - 1] = (r["cells"][ci - 1] + " " + r["cells"][ci]).strip()
                del r["cells"][ci]
            out["widths"][ci - 1] += out["widths"][ci]
            del out["widths"][ci], out["aligns"][ci]
    return out


def layout(src, blocks, segs, info):
    """Раскладка листа (как docx_struct.layout): поля — по тексту тела, колонтитулы — сеткой по повторяющимся
    сегментам типичного листа, номер листа — число, растущее с листом."""
    import fitz
    doc = fitz.open(str(src))
    W, H = doc[0].rect.width, doc[0].rect.height
    body = [s for s in segs if "shot" not in s["loc"] and s["loc"].get("bbox") and s["text"].strip()]
    run_keys = _running(body, H, len(doc))
    runs = [s for s in body if _is_running(s, run_keys, H) or _portal_running(s)]
    geoms = [b["geom"] for b in blocks if b.get("geom") and b["k"] in ("p", "h")] + \
        [p["geom"] for b in blocks if b["k"] == "list" for it in b["items"] for p in it if p.get("geom")]
    left = sorted(g["x0"] for g in geoms)[len(geoms) // 10] if geoms else 62
    right = W - sorted(g["x1"] for g in geoms)[-max(1, len(geoms) // 10)] if geoms else 59
    top = min((g["top"] for g in geoms), default=72)
    bottom = max((g["bottom"] for g in geoms), default=H - 72)
    by_page = {}
    for s in runs:
        by_page.setdefault(s["page"], []).append(s)
    pages = sorted(by_page)
    # образец — лист с самым частым числом сегментов колонтитулов (полный набор, а не случайно слитый OCR)
    common = Counter(len(by_page[p]) for p in pages).most_common(1)[0][0] if pages else 0
    same = [p for p in pages if len(by_page[p]) == common]
    rep = same[len(same) // 2] if same else None
    page_tok, num_start = None, None
    if rep is not None:
        for s in by_page[rep]:                               # число, растущее с листом: «Oldalszám: 7»
            m = re.search(r"\d+", s["text"])
            if not m:
                continue
            same = [t for p in pages for t in by_page[p] if _sig(t["text"]) == _sig(s["text"])
                    and abs(t["loc"]["bbox"][1] - s["loc"]["bbox"][1]) < 4]       # то же место на листе
            vals = Counter(int(re.search(r"\d+", t["text"]).group()) - t["page"] for t in same if re.search(r"\d+", t["text"]))
            # смещение номера от листа — одно у большинства (3PR42: нумерация сдвигается посреди документа)
            if len(same) >= 3 and vals and vals.most_common(1)[0][1] >= 0.6 * len(same):
                page_tok = m.group()
                # нумерация — с числа оригинала: у 3SZ19 первый лист PDF — «Oldalszám: 2» (титул не в файле); с титулом
                # тело начинается со 2-го листа — и число его (раньше выходило на единицу меньше)
                first = 1 if any(s["page"] == 0 for s in runs) else 2
                num_start = {"page": first, "value": vals.most_common(1)[0][0] + first - 1}
                break
    bold = info.get("style")
    # в сетку колонтитула — только строки, одинаковые почти на всех листах (_stable_running)
    zone_of = lambda z: [s for s in by_page.get(rep, []) if _zone(s, H) == z and _stable_running(s, runs, run_keys, H)] \
        if rep is not None else []
    headz = zone_of("head")
    head = _grid(headz, bold) if headz else None
    footz = zone_of("foot")
    foot = _grid(footz, bold) if footz else {"rows": [], "widths": [], "aligns": [], "inner": False, "box": False}
    if head:                                # шапка — шрифтом документа (стиль шапки по умолчанию — Tahoma, 3SZ19 — Courier)
        head["font"], head["size"] = info.get("font"), info.get("size")
    head_bottom = max((s["loc"]["bbox"][3] for s in headz), default=36)
    foot_top = min((s["loc"]["bbox"][1] for s in footz), default=H - 36)
    return {"page_sizes": __import__("docx_struct")._page_sizes(src),
            "W": W, "H": H, "left": left, "right": max(right, 20), "top": max(top - 4, head_bottom + 6),
            "bottom": max(H - bottom - 4, H - foot_top + 6), "header": 24, "footer": 18,
            "head_imgs": [], "head_lines": [], "head_grid": head,
            "stamp": foot["rows"], "stamp_widths": foot["widths"], "stamp_aligns": foot["aligns"],
            "stamp_font": info.get("font") or "Times New Roman", "stamp_size": info.get("size"),
            "stamp_box": foot["box"], "stamp_inner": False, "page_tok": page_tok, "num_start": num_start,
            "first_is_body": any(s["page"] == 0 for s in runs)}    # колонтитулы и на 1-м листе — титула нет
