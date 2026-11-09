# -*- coding: utf-8 -*-
"""Где перевод встаёт на листе: правила «в своё место оригинала», общие для сканов форм и схем (аудит 29.10, причина №5).

1. Повёрнутые подписи (стрелки схем: «sikeres», «sikertelen») — rotated_lines: строки текстового слоя с вертикальным
   направлением; перевод рисуется в том же повороте (assemble.build_pdf, rotate=).
2. Блок схемы — container: перевод вписывается в самую малую картинку или заливку, внутри которой стоит надпись
   (не вылезает за блок, на стрелки и соседние блоки); надписи одного блока — одним текстом (blocks_by_container).
3. Скан формы — pixels_only: абзац из фамилий, сокращений из заглавных, чисел и дат не перерисовывается — остаётся
   пикселями оригинала (рукопись, штамп, подпись и печатная фамилия рядом; перевод фамилии и расшифровка «ÜBO» — не
   задача формы).
4. Скан — mark_tilt / tilted: блок слов, повёрнутый к строкам листа (резиновый штамп наискось), — пиксели оригинала."""
import re

import fitz

SHAPE_MIN, SHAPE_MAX = 400.0, 0.25        # блок схемы: от 400 pt² до четверти листа (фон и рамка листа — не блок)
SCHEME_SHARE = 0.4                       # и это не меньше 40 % надписей листа (поля бланка с рамками — не схема)
SCHEME_MIN = 4                            # лист — схема из блоков, если в блоках стоят от 4 надписей (заливка шапки таблицы — не схема)
ACR_MAX = 4                               # сокращение из заглавных: до 4 букв (ÜBO, ÜVFO, RTO, TO, ÁPR, JAN)


# ------------------------------------------------------------------ повёрнутые подписи
def rotated_lines(page, skip=None):
    """Вертикальные строки текстового слоя: [{t, bbox (как страница видна), size, bold, family, rot}]. rot — как у
    insert_htmlbox(rotate=): 270 — читается сверху вниз, 90 — снизу вверх. skip(bbox) — спрятанный текст (hidden_text)."""
    from layout import span_text, family
    d = page.get_text("rawdict", flags=fitz.TEXT_PRESERVE_WHITESPACE | fitz.TEXT_MEDIABOX_CLIP)
    rot = page.rotation_matrix
    out = []
    for b in d["blocks"]:
        if b.get("type") != 0:
            continue
        for ln in b["lines"]:
            v = fitz.Point(ln["dir"]) * rot - fitz.Point(0, 0) * rot
            if abs(v.x) > 0.1 or abs(v.y) < 0.9:
                continue
            t = "".join(span_text(dict(s, text="".join(c["c"] for c in s.get("chars") or []))) for s in ln["spans"]).strip()
            bb = fitz.Rect(ln["bbox"]) * rot
            if not re.search(r"\w", t) or (skip and skip(bb)):
                continue
            s0 = ln["spans"][0]
            out.append({"t": t, "bbox": [round(x, 2) for x in bb], "size": float(s0["size"]),
                        "bold": bool(s0["flags"] & 16) or "Bold" in s0["font"], "family": family(s0["font"]),
                        "rot": 270 if v.y > 0 else 90})
    return out


def rotated_avail(bbox, rot, page_rect, shapes=(), grow=0.4):
    """Место подписи: вдоль чтения запас в долю длины с обоих концов (перевод длиннее источника, подпись стоит у
    середины стрелки — растёт в обе стороны), но не дальше блоков (shapes) на пути строки; поперёк — как есть."""
    x0, y0, x1, y1 = bbox
    e = (y1 - y0) * grow
    top, bot = max(y0 - e, page_rect.y0), min(y1 + e, page_rect.y1)
    for r in shapes:
        if min(r.x1, x1) - max(r.x0, x0) <= 0:
            continue
        if top < r.y1 <= y0 + 1:
            top = r.y1 + 1
        if y1 - 1 <= r.y0 < bot:
            bot = r.y0 - 1
    return [x0, min(top, y0), x1, max(bot, y1)]


# ------------------------------------------------------------------ блоки схемы
def _outline(dr):
    """Замкнутая обводка (блок блок-схемы, ромб решения): (Rect, ромб?) или None. Контур — один прямоугольник,
    четырёхугольник или четыре отрезка; ромб — вершины не в углах охватывающей рамки."""
    kinds = "".join(i[0] for i in dr["items"])
    if kinds not in ("re", "qu", "llll"):
        return None
    r = dr["rect"]
    if kinds == "re":
        return r, False
    pts = [i[1] for i in dr["items"]] if kinds == "llll" else [getattr(dr["items"][0][1], k) for k in ("ul", "ur", "ll", "lr")]
    corners = [fitz.Point(r.x0, r.y0), fitz.Point(r.x1, r.y0), fitz.Point(r.x0, r.y1), fitz.Point(r.x1, r.y1)]
    far = min(r.width, r.height) * 0.25
    return r, all(min(abs(p - c) for c in corners) > far for p in pts)


def shape_rects(page, kinds=False):
    """Блоки схемы страницы — [Rect]: картинки-блоки (плитки заливки: схема, собранная из картинок) и, если на листе есть
    ромб решения (блок-схема линиями), залитые и обведённые фигуры. Без ромба заливки-векторы не блоки: так выглядят
    ячейки таблиц и поля ввода бланков (30RL2126, 30BOXFELTEL) — их место определяет сама строка. kinds=True — ещё и
    множество ромбов {tuple(Rect)}: текст ромба умещается в его середину."""
    area = page.rect.get_area()
    rot = page.rotation_matrix
    ok = lambda r: SHAPE_MIN <= r.get_area() <= SHAPE_MAX * area and min(r.width, r.height) > 12
    out, vector, outlines = [], [], []
    for ii in page.get_image_info():
        r = fitz.Rect(ii["bbox"]) * rot
        if ok(r):
            out.append(r)
    for dr in page.get_drawings():
        r = dr["rect"] * rot
        if dr.get("fill") is not None and ok(r):
            vector.append(r)
        o = _outline(dr)
        if o is not None and ok(o[0] * rot):
            outlines.append((o[0] * rot, o[1]))
    diamonds = {tuple(r) for r, d in outlines if d}
    if diamonds:
        out += vector + [r for r, _ in outlines]
    return (out, diamonds) if kinds else out


def clamp_right(rect, bbox, shapes):
    """Место надписи вправо — до ближайшего блока на её строке: подпись у стрелки между блоками («успешный» длиннее
    «sikeres») не заезжает на следующий блок; сама надпись (bbox) остаётся целиком."""
    bb = fitz.Rect(bbox)
    xs = [r.x0 for r in shapes if r.x0 >= bb.x1 - 1 and min(r.y1, bb.y1) - max(r.y0, bb.y0) > 0.3 * bb.height]
    if not xs or min(xs) - 1 >= rect.x1:
        return rect
    return fitz.Rect(rect.x0, rect.y0, max(min(xs) - 1, bb.x1), rect.y1)


def is_scheme(n_in_blocks, n_all):
    """Лист — схема из блоков: в блоках стоят от SCHEME_MIN надписей и не меньше SCHEME_SHARE всех надписей листа
    (заливка шапки таблицы, поля ввода бланка с рамками — единицы на листе текста; в схеме текст стоит в блоках)."""
    return n_in_blocks >= SCHEME_MIN and n_in_blocks >= SCHEME_SHARE * max(n_all, 1)


def container(shapes, bbox, ratio=25.0, floor=12000.0):
    """Самый малый блок, в котором по центру стоит надпись bbox, или None. Надпись — в блоке, если её середина внутри и
    внутри не меньше половины её площади (строка «és» на нижней кромке блока, подпись у края тени блока — тоже в блоке).
    Блок не больше ratio площадей надписи (но не меньше floor pt²): рамка во весь столбец и заливка колонки — не «блок
    подписи»; надпись у края блока (подпись слева в заливке ячейки) — не блок схемы, её место определяет сама строка."""
    bb = fitz.Rect(bbox)
    mid = fitz.Point((bb.x0 + bb.x1) / 2, (bb.y0 + bb.y1) / 2)
    best = None
    for r in shapes:
        if r.contains(mid) and (r & bb).get_area() >= 0.5 * bb.get_area()                 and r.get_area() <= max(ratio * bb.get_area(), floor)                 and (best is None or r.get_area() < best.get_area()):
            best = r
    return best


def blocks_by_container(shapes, bboxes):
    """{номер блока формы: [индексы надписей]} — надписи одного блока идут вместе (строки подписи блока схемы — отдельные
    фрагменты); None-ключ не используется — надписи вне блоков не возвращаются."""
    groups = {}
    for i, bb in enumerate(bboxes):
        c = container(shapes, bb)
        if c is not None:
            groups.setdefault(tuple(round(v, 1) for v in c), []).append(i)
    return groups


def fit_rect(box, labels, used, pad=1.5, diamond=False):
    """Куда писать перевод блока: (рамка, выключка). Внутри блока box с полем pad; у ромба — в его середине (вписанный
    прямоугольник половинного размера). Выключка по центру, если подписи (labels, рамки оригинала) стояли по центру блока,
    иначе влево от левого края подписей; по высоте — по середине блока в том же случае, иначе от верха подписей.
    used(rect) — высота перевода, вписанного в rect."""
    if diamond:
        inner = fitz.Rect(box.x0 + box.width / 4, box.y0 + box.height / 4, box.x1 - box.width / 4, box.y1 - box.height / 4)
    else:
        inner = fitz.Rect(box.x0 + pad, box.y0 + pad, box.x1 - pad, box.y1 - pad)
    u = fitz.Rect(labels[0])
    for b in labels[1:]:
        u |= fitz.Rect(b)
    centered = diamond or abs((u.x0 + u.x1) / 2 - (box.x0 + box.x1) / 2) <= 0.2 * box.width
    if not centered:
        inner.x0 = max(inner.x0, u.x0 - 0.5)
    h = used(inner)
    cy = (box.y0 + box.y1) / 2
    align = "center" if centered else "left"
    if diamond or abs((u.y0 + u.y1) / 2 - cy) < 0.25 * box.height:
        return fitz.Rect(inner.x0, max(inner.y0, cy - h / 2 - 0.5), inner.x1, min(inner.y1, cy + h / 2 + 0.5)), align
    return fitz.Rect(inner.x0, max(inner.y0, u.y0 - 0.5), inner.x1, inner.y1), align


# ------------------------------------------------------------------ скан формы: что остаётся пикселями
_EDGE = ".,;:()[]{}|/-–—_«»\"'“”„\\"
_known = {}


def known_words():
    """Слова словаря проекта (первая колонка глоссария и словаря комплекта, строчными, от 2 букв): метки форм, их
    сокращения и термины. Файлы — из project.json проектов рядом с кодом; нет файлов — пусто (тогда метка опознаётся
    только по виду слова)."""
    import json
    from pathlib import Path
    if "w" in _known:
        return _known["w"]
    words = set()
    for pj in sorted((Path(__file__).resolve().parent.parent / "projects").glob("*/project.json")):
        try:
            prj = json.loads(pj.read_text(encoding="utf-8"))
        except (OSError, ValueError) as e:
            print(f"pagefit: {pj}: {type(e).__name__}: {e}"[:200], flush=True)
            continue
        for k in ("glossary", "terms"):
            f = pj.parent / prj[k] if prj.get(k) else None
            if f is None or not f.exists():
                continue
            for ln in f.read_text(encoding="utf-8").splitlines():
                if ln.startswith("|") and not ln.startswith("|---"):
                    cell = ln.strip().strip("|").split("|")[0]
                    words |= {w.lower() for w in re.findall(r"[^\W\d_]{2,}", cell)}
    _known["w"] = words
    return words


def _kind(tok, known):
    """Слово абзаца: num — с цифрой (число, дата, код); label — строчное слово, подпись с «:», длинные заглавные (5+
    букв) или слово словаря; acr — до 4 заглавных; name — Имя с прописной (3+ букв); '' — знаки."""
    core = tok.strip(_EDGE)
    letters = [c for c in core if c.isalpha()]
    # подпись, слипшаяся со значением («Verziószám:1», «Dátum:2025.05.»), и обрывок переноса («Ver-») — метки
    if re.match(r"[^\W\d_]{3,}:", tok) or tok.endswith("-") and len(letters) >= 2:
        return "label"
    if re.search(r"\d", core):
        return "num"
    if not letters:
        return ""
    if tok.rstrip(".,;)").endswith(":") or len(letters) != len(re.sub(r"[-’']", "", core)):
        return "label"
    low = core.lower()
    upper = all(c.isupper() for c in letters)
    if upper and len(letters) > ACR_MAX or not letters[0].isupper():
        return "label"
    if upper:
        return "label-acr" if low in known else "acr"
    if len(letters) >= 3 and all(c.islower() for c in letters[1:]):
        return "label" if low in known else "name"
    return "label"


def pixels_only(text, known=None):
    """Абзац скана формы, который не перерисовывается: фамилии (Имена с прописной, которых нет в словаре), числа,
    даты, неизвестные сокращения из заглавных. Метка — любое слово словаря проекта, строчное слово, подпись с «:»,
    длинное слово заглавными; известное сокращение (ÜBO, TO) — метка, если рядом нет фамилии («Bárdos Zoltán ÜBO» —
    фамилия с должностью, остаётся). Правило по тексту абзаца: словарь проекта — тот же, что у перевода."""
    known = known_words() if known is None else known
    kinds = [k for k in (_kind(t, known) for t in text.split()) if k]
    if "label" in kinds:
        return False
    return not ("label-acr" in kinds and "name" not in kinds)


INK_RED = 0.15      # доля красных пикселей в рамке слова, от которой краска — красная печать (у слов красной плашки 0,2–0,38)
COL_GAP_LINE = 1.5  # в строке прозы (от трёх пробелов) зазор-столбец — и во столько раз больше её самого большого пробела,
COL_GAP_MAX = 14.0  # но от 14 pt — столбец и там (граничные 12–14 pt в строке прозы — пробел машинописи)


def ink_red(rgb, box, dpi):
    """Доля красных пикселей в рамке слова. В сером изображении красная печать (251, 61, 49) светлее порога тёмной краски (segments._ink),
    и уверенно прочитанное слово шло как бледная рукопись (красное предупреждение «… FIGYELEMBEVÉTELÉVEL KÉSZÜLT!», красные
    заголовки форм: 30YCGT, 30RYBSZGT, 30RVRYTSBIGT). Красный — насыщенный и яркий: R выше G и B не меньше чем на 90; ручка — синяя
    или чёрная, красной краски в ней нет."""
    import numpy as np
    k = dpi / 72
    x0, y0, x1, y1 = (int(v * k) for v in box)
    cut = rgb[max(y0, 0):max(y1, y0 + 1), max(x0, 0):max(x1, x0 + 1)].astype(np.int16)
    if not cut.size:
        return 0.0
    return float(((cut[:, :, 0] - np.maximum(cut[:, :, 1], cut[:, :, 2]) >= 90) & (cut[:, :, 0] >= 150)).mean())


def mark_red_ink(w, rgb, dpi):
    """Слову Tesseract — доля красной краски (ink_c); от INK_RED красная печать — краска наравне с тёмной: плотность слова (ink) не ниже неё,
    и слово не уходит в «бледную рукопись» (segments.ocr_words)."""
    red = round(ink_red(rgb, w["box"], dpi), 3)
    w["ink_c"] = red
    if red >= INK_RED:
        w["ink"] = max(w.get("ink", 0), red)


TILT_DEG = 5.0      # блок слов скана, повёрнутый к строкам листа на столько градусов и больше, — штамп или рукопись наискось
TILT_GAIN = 1.2     # ... если строки блока при его угле и вправду резче, чем при угле листа (у штампа 30RYBSZGT 8° и 1,76 раза)
TILT_MAX_PTS = 20000  # точек краски блока на замер не больше (через одну-две — угол тот же, считается быстрее)


def _tilt(ys, xs):
    """Углы −12…12° через 0,25° и резкость проекции точек краски на вертикаль при каждом (сумма квадратов столбцов
    гистограммы: строки, повёрнутые на этот угол, дают узкие высокие пики)."""
    import numpy as np
    step = max(1, len(ys) // TILT_MAX_PTS)
    ys, xs = ys[::step].astype(float), xs[::step].astype(float)
    x = xs - xs.mean()
    angs = np.arange(-12, 12.01, 0.25)
    sharp = []
    for a in angs:
        b = np.round(ys - x * np.tan(np.radians(a))).astype(int)
        sharp.append(float((np.bincount(b - b.min()).astype(float) ** 2).sum()))
    return angs, np.array(sharp)


def mark_tilt(words, gray, dpi):
    """Словам скана — наклон их блока Tesseract (key: блок, абзац) к строкам листа: tilt = [разница углов, °; во сколько
    раз проекция краски блока при его угле резче, чем при угле листа] или None (краски на замер мало). Угол — тот, при
    котором проекция печатных слов блока на вертикаль резче всего; угол листа — медиана углов блоков по количеству краски
    (скан листа целиком бывает скошен на 1–2°). Резиновый штамп ставят наискось («Paksi Atomerőmű Zrt. Minőségfelügyeleti
    Osztály 1.», 30RYBSZGT PG-144760: 8°): Tesseract читает его мусором («ömű Zrt. Minőségfelügyeleti 97», «pakol Atomer
    tégy»), и два перевода ложились на стёртый штамп друг на друга (tilted)."""
    import numpy as np
    k = dpi / 72
    blocks = {}
    for w in words:
        w["tilt"] = None
        if len(re.sub(r"\W", "", w["t"])) >= 2 and w.get("ink", 0) >= 0.1 and w["h"] >= 2:
            blocks.setdefault(tuple(w["key"]), []).append(w)
    meas = {}
    for key, ws in blocks.items():
        ys, xs = [], []
        for w in ws:
            x0, y0, x1, y1 = (max(int(v * k), 0) for v in w["box"])
            yy, xx = np.nonzero(gray[y0:max(y1, y0 + 1), x0:max(x1, x0 + 1)] < 128)
            ys.append(yy + y0); xs.append(xx + x0)
        ys, xs = np.concatenate(ys), np.concatenate(xs)
        if len(ys) >= 200 * (dpi / 300) ** 2:
            angs, sharp = _tilt(ys, xs)
            meas[key] = (angs, sharp, len(ys))
    if not meas:
        return
    best = sorted((m[0][int(np.argmax(m[1]))], m[2]) for m in meas.values())
    acc, page = 0, best[-1][0]
    for a, n in best:                                   # медиана углов по количеству краски
        acc += n
        if acc >= sum(n for _, n in best) / 2:
            page = a
            break
    for w in words:
        m = meas.get(tuple(w["key"]))
        if m:
            angs, sharp, _ = m
            i = int(np.argmax(sharp))
            w["tilt"] = [round(float(angs[i] - page), 2), round(float(sharp[i] / max(sharp[int(np.argmin(abs(angs - page)))], 1.0)), 3)]


def tilted(w):
    """Слово из блока, повёрнутого к строкам листа (mark_tilt), — штамп или рукопись наискось: пиксели оригинала, не
    переводится и не закрывается плашкой (segments.ocr_words — как рукопись)."""
    t = w.get("tilt")
    return bool(t) and abs(t[0]) >= TILT_DEG and t[1] >= TILT_GAIN


CELL_WORD = re.compile(r"[A-ZÁÉÍÓÖŐÚÜŰ][a-záéíóöőúüű]{2,}")


def wide_gap(g, gaps, col_gap, word):
    """Зазор g перед словом word — пустое место перед столбцом: от col_gap и, в строке с пробелами машинописи (от трёх зазоров меньше
    col_gap — gaps, зазоры своей строки), не меньше полутора самого большого пробела и COL_GAP_MAX. Машинопись с пробелами по 10–12 pt:
    зазор в 12,0 pt перед словом — пробел, а не столбец; он совпадал с левым краем у чужих строк листа, и строка рвалась («A KU 2.6. |
    pontjában…», 3PR10 стр. 2). Но если после зазора слово с прописной (CELL_WORD: подпись поля «Aláírás:», «[Aláírási», заголовок
    «Verzió»), это начало ячейки: в строке подписей формы («Dátum: … Név: Balogh Béla Szervezet: ÜIO | Aláírás:») и в шапке таблицы
    версий («Oldal Verzió Oldal …» через 11,5–12,2 pt) пробелы между ячейками — те же 9–12 pt. Иначе ячейки склеивались через границу
    («Szervezet: ÜIO Aláírás:», 3SZ23, 30RVRW/YVTBIGT), а «Verzió» переставало быть краем столбца, и «TU» ⇥ «-ban» под ним шли одним
    куском вместе с номером строки ниже. Значение после подписи (дата, номер, сокращение) ячейку не начинает: «Dátum: 2026.04.16.» —
    один кусок, дата — пиксели (value_tail)."""
    spaces = [x for x in gaps if x < col_gap]
    return g >= col_gap and (len(spaces) < 3 or bool(CELL_WORD.search(word))
                             or g >= min(COL_GAP_LINE * max(spaces), COL_GAP_MAX))


def ruleword(w):
    """Кусок линейки ячейки, прочитанный знаками («[d», «|1»): два знака и больше в рамке уже 0,3 высоты на знак (букв такой ширины нет)
    и плотная краска — линия, не слово: картинка, не переводится и не закрывается плашкой («[d Страница» в шапке таблицы версий 3PR10).
    Знаки — те, что в рамке (inbox): у склейки переноса рамка первой половины («mű-» + «ködik»), у куска слова, разрезанного по линейке,
    raw — чтение целого слова; по ним жирное слово уходило бы в картинку целиком."""
    x0, y0, x1, y1 = w["box"]
    n = len(w["inbox"]) if "inbox" in w else len(w.get("raw", w["t"]))
    return n >= 2 and (x1 - x0) < 0.3 * (y1 - y0) * n and w.get("ink", 0) > 0.45


def split_values(frags, pix, union):
    """Фрагменты скана без значений посреди строки: значение между двумя подписями (value_runs) — рамка в pix, подпись и остаток строки —
    свои фрагменты (union — объединение рамок, layout._union). Фрагмент — словарь с t, bbox, line, ws."""
    out = []
    for f in frags:
        runs = value_runs(f["ws"])
        if not runs:
            out.append(f)
            continue
        pos = 0
        for a, b in runs:
            if pix is not None:
                pix.append(union([w["box"] for w in f["ws"][a:b]]))
            if a > pos:
                out.append(dict(f, ws=f["ws"][pos:a]))
            pos = b
        if pos < len(f["ws"]):
            out.append(dict(f, ws=f["ws"][pos:]))
    for f in out:
        f["t"] = " ".join(w["t"] for w in f["ws"])
        f["bbox"] = union([w["box"] for w in f["ws"]])
    return out


def value_runs(words, known=None):
    """[(a, b)] — куски words[a:b] посреди фрагмента скана, что остаются пикселями: значение между двумя подписями («Dátum: 71127 JAN 2
    Név: …»). Начало — слово после подписи с «:», конец — перед следующей подписью с «:»; внутри числа, неизвестные сокращения из
    заглавных и знаки, и хотя бы одно число (дата, номер) — одно сокращение («Szervezet: ÜVFO |») переводится, как прежде. Хвост строки —
    value_tail. Неверно прочитанное значение штампа («2022 JAN 25» → «71127 JAN 2») не рисуется заново."""
    known = known_words() if known is None else known
    kinds = [_kind(w["t"], known) for w in words]
    runs, i, n = [], 1, len(words)
    while i < n:
        if kinds[i - 1] == "label" and words[i - 1]["t"].rstrip().endswith(":"):
            j = i
            while j < n and kinds[j] in ("num", "acr", ""):
                j += 1
            if i < j < n and kinds[j] == "label" and words[j]["t"].rstrip().endswith(":") and "num" in kinds[i:j]:
                runs.append((i, j))
                i = j
                continue
        i += 1
    return runs


def value_tail(words, known=None):
    """Сколько первых слов фрагмента скана остаётся текстом: значение после подписи («Dátum: 2025 JAN 31», «Verziószám: 2»)
    — числа, даты и неизвестные сокращения из заглавных — не читается и не перерисовывается, это пиксели оригинала
    (рукопись, штамп, номер версии: OCR путает «04» с «94», «2023» с «2025»). Только хвост после слова с «:»; слова
    значения со строчной буквой («51 bar-os») и известные сокращения («RTO») остаются текстом."""
    known = known_words() if known is None else known
    k = len(words)
    while k > 1 and _kind(words[k - 1]["t"], known) in ("num", "acr", ""):
        k -= 1
    return k if k < len(words) and words[k - 1]["t"].rstrip().endswith(":") else len(words)
