# -*- coding: utf-8 -*-
"""Где перевод встаёт на листе: правила «в своё место оригинала», общие для сканов форм и схем (аудит 29.10, причина №5).

1. Повёрнутые подписи (стрелки схем: «sikeres», «sikertelen») — rotated_lines: строки текстового слоя с вертикальным
   направлением; перевод рисуется в том же повороте (assemble.build_pdf, rotate=).
2. Блок схемы — container: перевод вписывается в самую малую картинку или заливку, внутри которой стоит надпись
   (не вылезает за блок, на стрелки и соседние блоки); надписи одного блока — одним текстом (blocks_by_container).
3. Скан формы — pixels_only: абзац из фамилий, сокращений из заглавных, чисел и дат не перерисовывается — остаётся
   пикселями оригинала (рукопись, штамп, подпись и печатная фамилия рядом; перевод фамилии и расшифровка «ÜBO» — не
   задача формы)."""
import re

import fitz

SHAPE_MIN, SHAPE_MAX = 400.0, 0.25        # блок схемы: от 400 pt² до четверти листа (фон и рамка листа — не блок)
NEAR = 1.5                                # допуск «надпись внутри блока», pt
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
def shape_rects(page):
    """Блоки страницы: картинки и залитые фигуры — [Rect], без фона листа и мелочи (значки, линии, тени-стрелки)."""
    area = page.rect.get_area()
    out = []
    for ii in page.get_image_info():
        r = fitz.Rect(ii["bbox"]) * page.rotation_matrix
        if SHAPE_MIN <= r.get_area() <= SHAPE_MAX * area and min(r.width, r.height) > 12:
            out.append(r)
    for dr in page.get_drawings():
        r = dr["rect"] * page.rotation_matrix
        if dr.get("fill") is not None and SHAPE_MIN <= r.get_area() <= SHAPE_MAX * area and min(r.width, r.height) > 12:
            out.append(r)
    return out


def clamp_right(rect, bbox, shapes):
    """Место надписи вправо — до ближайшего блока на её строке: подпись у стрелки между блоками («успешный» длиннее
    «sikeres») не заезжает на следующий блок; сама надпись (bbox) остаётся целиком."""
    bb = fitz.Rect(bbox)
    xs = [r.x0 for r in shapes if r.x0 >= bb.x1 - 1 and min(r.y1, bb.y1) - max(r.y0, bb.y0) > 0.3 * bb.height]
    if not xs or min(xs) - 1 >= rect.x1:
        return rect
    return fitz.Rect(rect.x0, rect.y0, max(min(xs) - 1, bb.x1), rect.y1)


def container(shapes, bbox, ratio=25.0):
    """Самый малый блок, внутри которого по центру стоит надпись bbox, или None. Блок не больше ratio площадей надписи:
    рамка во весь столбец и заливка колонки — не «блок подписи»; надпись у края блока (подпись слева в заливке ячейки) —
    не блок схемы, её место определяет сама строка."""
    bb = fitz.Rect(bbox)
    big = fitz.Rect(bb.x0 - NEAR, bb.y0 - NEAR, bb.x1 + NEAR, bb.y1 + NEAR)
    best = None
    for r in shapes:
        if r.contains(big) and r.get_area() <= ratio * max(bb.get_area(), 1.0)                 and abs((bb.x0 + bb.x1) / 2 - (r.x0 + r.x1) / 2) <= 0.2 * r.width                 and (best is None or r.get_area() < best.get_area()):
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


def fit_rect(box, labels, used, pad=1.5):
    """Куда писать перевод блока: внутри блока box с полем pad; по высоте — по середине блока, если подписи (labels,
    рамки оригинала) стояли по середине, иначе от верха подписей. used(rect) — высота перевода, вписанного в rect."""
    inner = fitz.Rect(box.x0 + pad, box.y0 + pad, box.x1 - pad, box.y1 - pad)
    u = fitz.Rect(labels[0])
    for b in labels[1:]:
        u |= fitz.Rect(b)
    h = used(inner)
    cy = (box.y0 + box.y1) / 2
    if abs((u.y0 + u.y1) / 2 - cy) < 0.25 * box.height:
        return fitz.Rect(inner.x0, max(inner.y0, cy - h / 2 - 0.5), inner.x1, min(inner.y1, cy + h / 2 + 0.5))
    return fitz.Rect(inner.x0, max(inner.y0, u.y0 - 0.5), inner.x1, inner.y1)


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
