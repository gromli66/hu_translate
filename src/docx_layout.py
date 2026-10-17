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
    и «Verziószám: 0 Oldalszám:» отдельно — на другом (3PR42, 16.10)."""
    cnt, pos = Counter(), {}
    for s in segs:
        z = _zone(s, H)
        if z:
            cnt[(z, _sig(s["text"]))] += 1
            pos.setdefault((z, round(s["loc"]["bbox"][1] / 4)), Counter())[_sig(s["text"])] += 1
    need = max(2, 0.5 * pages)
    # на высоте — и текст в основном тот же (≥ половины листов): название раздела в рамке на той же высоте меняется
    # от раздела к разделу — это текст, а не колонтитул
    return {k for k, n in cnt.items() if n >= need} | \
        {k for k, c in pos.items() if sum(c.values()) >= need and c.most_common(1)[0][1] >= 0.5 * pages}


def _is_running(s, keys, H):
    z = _zone(s, H)
    return bool(z) and ((z, _sig(s["text"])) in keys or (z, round(s["loc"]["bbox"][1] / 4)) in keys)


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
                elif it[0] == "re":
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
    for i, (oi, a) in enumerate(rules):
        if oi != "h":
            continue
        for j, (oj, b) in enumerate(rules):
            if oj == "v" and b[1] - 2 <= a[0] <= b[2] + 2 and a[1] - 2 <= b[0] <= a[2] + 2:
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
        if x1 - x0 < 60 or y1 - y0 < 15 or not any(x0 + 3 < v[0] < x1 - 3 for v in gv):
            continue
        ys = _cluster([h[0] for h in gh if y0 - 3 <= h[0] <= y1 + 3] + [y0, y1])
        ys = [y for i, y in enumerate(ys) if i == 0 or y - ys[i - 1] >= 5]        # двойная линия — одна
        xs = _cluster([v[0] for v in gv])
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
                or any(r.intersects(q) for q in rects) or any(r.contains(fitz.Point(*_center(s))) for s in segs):
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
                         "body_mono": mono})
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


def _table_block(t, segs, pno, bold, font, size, mono, imgs=()):
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
    rows = []
    for r, cells in enumerate(grid):
        row = []
        for ci, (a, z, e) in enumerate(cells):
            ss = sorted(put.get((r, ci), []), key=lambda s: (round(s["loc"]["bbox"][1] / 3), s["loc"]["bbox"][0]))
            paras, geom = [], None
            for group in _lines(ss):
                g0 = group[0]["loc"]["bbox"]
                gg = {"page": pno + 1, "top": g0[1], "bottom": max(g["loc"]["bbox"][3] for g in group), "x0": g0[0],
                      "x1": max(g["loc"]["bbox"][2] for g in group), "first_x0": g0[0], "size": size,
                      "lines": max(len(g["loc"].get("frags") or [1]) for g in group)}
                paras.append({"k": "p", "runs": _runs_of(group, bold, font, size, mono), "geom": gg, "seg": True})
                geom = gg if geom is None else dict(geom, x0=min(geom["x0"], gg["x0"]), x1=max(geom["x1"], gg["x1"]),
                                                     bottom=gg["bottom"], lines=geom["lines"] + gg["lines"])
            for img in pics.get((r, ci), []):
                paras.append({"k": "p", "runs": [{"img": img}], "seg": True})
            row.append({"blocks": paras, "th": False, "geom": geom, "edges": e, "span": z - a})
        rows.append(row)
    head = [p for c in rows[0] for p in c["blocks"]] if rows else []
    # картинка в ячейке шапки (значок, JPDS текстовым слоем) — не текст: KeyError 't' ронял сборку (18.10)
    if head and len(rows) > 1 and all(r.get("b") for p in head for r in p["runs"] if r.get("t", "").strip()):
        for c in rows[0]:
            c["th"] = True                                  # шапка — первая строка, если вся полужирная
    return {"k": "table", "rows": rows, "borders": True, "frame": {"xs": [round(x, 1) for x in xs]},
            "geom": {"page": pno + 1, "top": ys[0], "bottom": ys[-1], "x0": xs[0], "x1": xs[-1], "first_x0": xs[0],
                     "size": size}}


def _grids(items, bold, font, size, mono):
    """Сетка без линеек — таблицей без рамок: подряд от трёх строк из нескольких коротких кусков, куски которых встают в
    общие столбцы (перечень кодов 3PR42 л. 112: «04.TH10D001 ⇥ KM1/1 ⇥ 02.TK42D001 ⇥ …» в шесть столбцов). Строкой с
    табуляцией по месту оригинала последний столбец не влезал и переносился (кегль сборки крупнее), 17.10. Строка
    оглавления (в конце — номер листа) — не сетка."""
    from layout import code_line

    def row_ok(it):
        g = it[2]
        return it[1] == 0 and len(g) >= 2 and len({x["page"] for x in g}) == 1 \
            and all(len(x["text"]) <= 60 for x in g) and not re.fullmatch(r"\d{1,4}", g[-1]["text"].strip())
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


def _regions(doc, content, bold, font, size, mono, mask=None):
    """Таблицы и рисунки листов: {лист: [{"rect", "block", "ids"}]} — сегменты ids идут не абзацами, а в них."""
    import fitz
    import numpy as np
    from layout import rules_from_image
    # от путей листа нужны линии (здесь) и рамки (для рисунков): все пути всех листов держали в памяти до 6,6 ГБ —
    # у 3PR42 текст нарисован контурами, тысячи путей на лист (17.10)
    draw, per, scans = [], [], []
    for page in doc:
        d = page.get_cdrawings()
        draw.append([{"rect": x["rect"]} for x in d])
        hs, vs = _rules(d)
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
        figs = [(r, {"png": page.get_pixmap(dpi=150, clip=r).tobytes("png"), "x0": r.x0, "x1": r.x1, "top": r.y0,
                     "bottom": r.y1}) for r in _figures(page, masks.get(pno, segs) if mask else segs, draw[pno],
                                                         scans[pno])]
        # текстовый слой: горизонталь как у линии под шапкой других листов — не рамка листа, если таблицу держат свои
        # вертикали (верх таблицы титула Eskom на высоте линии шапки: таблица не находилась); рамку листа отсекают
        # её же повторяющиеся вертикали
        hs = [r for r in per[pno][0] if mask is not None or not running("h", r)]
        vs = [r for r in per[pno][1] if not running("v", r)]
        if mask is not None:
            infos = page.get_image_info(xrefs=True)          # один раз на лист
            figs += _raster(page, segs, figs, reused, infos)
            # рисунок — ровно картинка PDF: её xref (сборка вставит скриншот с переведёнными надписями, shots)
            for r, img in figs:
                for im in infos:
                    q = fitz.Rect(im["bbox"])
                    if im["xref"] and (q & r).get_area() >= 0.9 * max(r.get_area(), q.get_area()):
                        img["xref"] = im["xref"]
        regs, used = [], set()
        for t in _tables(hs, vs):
            r = fitz.Rect(t[0][0], t[1][0], t[0][-1], t[1][-1])
            if any((f & r).get_area() > 0.8 * r.get_area() for f, _ in figs):
                continue                                    # сетка внутри рисунка (график) — часть рисунка
            inside = [s for s in segs if r.contains(fitz.Point(*_center(s)))]
            mine = [i for i, (f, _) in enumerate(figs) if r.contains(fitz.Point((f.x0 + f.x1) / 2, (f.y0 + f.y1) / 2))]
            if inside:
                used |= set(mine)
                regs.append({"rect": r, "ids": {s["id"] for s in inside},
                             "block": _table_block(t, inside, pno, bold, font, size, mono, [figs[i] for i in mine])})
        for i, (r, img) in enumerate(figs):
            if i in used or any(r.intersects(g["rect"]) for g in regs):
                continue
            regs.append({"rect": r, "ids": {s["id"] for s in segs if r.contains(fitz.Point(*_center(s)))},
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
    content = [s for s in body if not _is_running(s, run_keys, H)]
    # бегущее название раздела в рамке сверху листа (3PR42: «1.15. Átfolyás …» на каждом листе раздела) — то же, что
    # вверху предыдущего листа, — не текст (на всех листах сразу его не найти: у каждого раздела своё)
    # название — от 6 знаков с буквами: короткое значение («3», «N/A» титульных сведений Eskom) повторяется само по себе
    # и пропадало (17.10)
    titled = lambda s: len(re.sub(r"\W", "", s["text"])) >= 6 and re.search(r"[^\W\d_]{3}", s["text"])
    top = {}
    for s in content:
        if s["loc"]["bbox"][3] < 0.2 * H:
            top.setdefault(s["page"], set()).add(_sig(s["text"]))
    content = [s for s in content if not (s["loc"]["bbox"][3] < 0.2 * H and titled(s)
                                          and _sig(s["text"]) in top.get(s["page"] - 1, ()))]
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
    regions = _regions(doc, content, bold, font, size, mono, mask=body if text else None)
    # порядок: абзацы — по сегментам (порядок чтения OCR), таблица или рисунок — на месте первого своего сегмента;
    # рисунок без сегментов — перед первым сегментом листа ниже его верха
    taken = {i for regs in regions.values() for g in regs for i in g["ids"]}
    flow = [s for s in content if s["id"] not in taken]
    marks = {}          # текстовый слой: маркер пункта — свой сегмент слева от текста пункта (layout.frags_from_text)
    if text:            # только в потоке: «-» в ячейке таблицы — значение (НГТУ)
        bullet = [s for s in flow if re.fullmatch(r"\s*[•▪◦·●○■□o–\-\uf020-\uf0ff]\s*", s["text"])]
        # маркер — если справа на той же строке ближе 40 pt есть пункт; иначе «-» — значение («Revision: ⇥ -»)
        lone = {s["id"] for s in bullet if any(
            t["page"] == s["page"] and t["id"] != s["id"] and 0 < t["loc"]["bbox"][0] - s["loc"]["bbox"][2] < 40
            and abs((t["loc"]["bbox"][1] + t["loc"]["bbox"][3]) - (s["loc"]["bbox"][1] + s["loc"]["bbox"][3])) / 2
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
    items = _grids(items, bold, font, size, mono)
    blocks, cur = [], None                                  # cur — открытый список из пунктов с тире
    last_head = {}                                          # уровень → текст последнего заголовка
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
        # заголовок: короткий полужирный по центру; в текстовом слое — и у левого края, если с номером («2.1.1 Purpose»);
        # или по метке портала (roles.py: правила находили 16 % заголовков, портал — 57 % при точности 0,88)
        # (лист с метками портала — заголовок только по метке: правило брало и подписи «5. ábra …» по центру, и
        # полужирные «Korlátozó feltétel:», 18.10)
        if len(group) == 1 and (not loc.get("roles") and len(runs) == 1 and (loc.get("align") == "center" or text_mode
                and m and len(loc.get("frags") or [1]) == 1 and re.search(r"[^\W\d_]{3}", text[m.end(1):]))
                and runs[0]["b"] and len(text) < 120
                or loc.get("role") == "heading" and len(text) < 200 and re.search(r"[^\W\d_]{3}", text)):
            level = min(len(m.group(1).split(".")) if m else 1, 5)
            norm = re.sub(r"\s+", " ", text)
            if last_head.get(level) == norm:
                continue        # бегущее название раздела в начале каждого листа (3SZ19: 41 и 42 раза) — не новый раздел
            last_head[level] = norm
            for lv in range(level + 1, 6):
                last_head.pop(lv, None)
            blocks.append({"k": "h", "level": level, "runs": runs, "geom": geom, "center": loc.get("align") == "center"})
            cur = None
            continue
        mark = next((d for d in layer.get(s["page"], {}).get("dash", [])
                     if abs(d[1] - (first[1] + first[3]) / 2) < 0.6 * lh and 0 < first[0] - d[0] < 30), None)
        # текстовый слой: маркер пункта — знаком текста («•», «–», маркер Symbol) в начале строки или своим сегментом
        # слева на той же строке, а не чёрточкой краски
        ch = None
        if text_mode:
            tm = re.match(r"\s*([•▪◦·●○■□–\-\uf020-\uf0ff])\s", text)
            ch = tm.group(1) if tm else next((t for bb, t in marks.get(s["page"], ())
                                              if abs((bb[1] + bb[3]) / 2 - (first[1] + first[3]) / 2) < 0.6 * lh
                                              and 0 < first[0] - bb[2] < 40), None)
            mark = ch
        kind = ("Dash" if ch in "–-" or "\uf020" <= ch <= "\uf0ff" and chr(ord(ch) - 0xF000) in "-–" else "Disc") \
            if ch else "Dash"
        if not mark and loc.get("role") == "list":
            # пункт по метке портала без знака маркера («a) …», «1. …», «ba) …») — пункт без маркера, метка — в тексте
            mark, kind = True, "Label"
        para = {"k": "p", "runs": runs, "geom": geom, "seg": True}    # seg — абзац собран, merge_inline не клеит
        last = cur["items"][-1][-1] if cur is not None and cur["items"] else (blocks[-1] if blocks else None)
        # границы абзацев листа по меткам портала (loc["roles"]) уже верные: строки по правилам не клеим
        if not mark and len(group) == 1 and last is not None and not loc.get("roles") and _continues(last, s, lh):
            _join(last, runs, geom)
            continue
        if mark:
            if cur is None or abs(cur["x0"] - x0) > 6 or cur["kind"] != kind:
                cur = {"k": "list", "kind": kind, "items": [], "x0": x0}
                blocks.append(cur)
            cur["items"].append([para])
            continue
        if cur is not None and cur["items"] and abs(first[0] - cur["x0"]) < 8 and loc.get("align") != "center":
            cur["items"][-1].append(para)                   # продолжение пункта (OCR разбил пункт на сегменты)
            continue
        cur = None
        indent = round(max(0.0, first[0] - left) / 9) * 9
        if indent >= 9 and loc.get("align") != "center":
            para["indent"] = min(indent, 90)                # отступ по оригиналу — уровнем стиля (сборщик)
        blocks.append(para)
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
    if text:            # разрывы листа — правилом PDF с разметкой (docx_struct.mark_breaks): верх и низ тела листов
        import docx_struct
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
    runs = [s for s in body if _is_running(s, run_keys, H)]
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
    head = _grid([s for s in by_page.get(rep, []) if _zone(s, H) == "head"], bold) if rep is not None and \
        any(_zone(s, H) == "head" for s in by_page.get(rep, [])) else None
    footz = [s for s in by_page.get(rep, []) if _zone(s, H) == "foot"] if rep is not None else []
    foot = _grid(footz, bold) if footz else {"rows": [], "widths": [], "aligns": [], "inner": False, "box": False}
    if head:                                # шапка — шрифтом документа (стиль шапки по умолчанию — Tahoma, 3SZ19 — Courier)
        head["font"], head["size"] = info.get("font"), info.get("size")
    head_bottom = max((s["loc"]["bbox"][3] for s in by_page.get(rep, []) if _zone(s, H) == "head"), default=36)
    foot_top = min((s["loc"]["bbox"][1] for s in footz), default=H - 36)
    return {"W": W, "H": H, "left": left, "right": max(right, 20), "top": max(top - 4, head_bottom + 6),
            "bottom": max(H - bottom - 4, H - foot_top + 6), "header": 24, "footer": 18,
            "head_imgs": [], "head_lines": [], "head_grid": head,
            "stamp": foot["rows"], "stamp_widths": foot["widths"], "stamp_aligns": foot["aligns"],
            "stamp_font": info.get("font") or "Times New Roman", "stamp_size": info.get("size"),
            "stamp_box": foot["box"], "stamp_inner": False, "page_tok": page_tok, "num_start": num_start,
            "first_is_body": any(s["page"] == 0 for s in runs)}    # колонтитулы и на 1-м листе — титула нет
