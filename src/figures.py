# -*- coding: utf-8 -*-
"""Фигуры листа — один классификатор вместо россыпи признаков (UNIFIED_DESIGN.md, П6, этап E7).

Раньше решение «что на листе — рисунок» жило в четырёх местах разными признаками: лист-скан (docx_layout.is_scan,
_ocr_layer), лист-плакат (_graphic_page), таблица-график (_chart_table) и две схемы — `_vector_figures` (текстовый слой: от
шести фигур, от 80×60 pt) и `_scheme_figures` (распознанный: стрелки, подписи-связки). Здесь:

* `figure_kind(...)` — что за лист: "scan" | "chart" | "scheme" | None. Лист целиком одним снимком (с переведённого PDF) —
  только для этих трёх видов; иначе рисунки — по отдельности (растр — картинкой PDF, схема — областью листа).
* `scheme_regions(...)` — схема: замкнутые фигуры и линии-связки с подписями — ОДИН набор признаков для текстового слоя
  и распознанного листа (у каждого из старых наборов было своё, чего нет у другого).
* `sheet_rect(...)` — рамка снимка листа: вся краска, без колонтитула оригинала.
* формы-шаблоны (7б) и сверка рисованного перевода вторым распознаванием (7в) — ниже.

Включение — HU_U_FIG (по умолчанию включено; 0 — прежнее поведение побайтно, старые функции в docx_layout остаются)."""
import os

import geom as GM


def enabled():
    return os.environ.get("HU_U_FIG", "1") != "0"


def _mid(s):
    x0, y0, x1, y1 = s["loc"]["bbox"]
    return (x0 + x1) / 2, (y0 + y1) / 2


# ------------------------------------------------------------------ лист: скан, плакат, график
def is_scan(page):
    """Лист — картинка (скан): растровые изображения закрывают хотя бы половину листа."""
    import fitz
    return sum(fitz.Rect(i["bbox"]).get_area() for i in page.get_image_info()) >= 0.5 * page.rect.get_area()


def ocr_layer(page):
    """Текст листа — невидимый слой распознавания поверх скана (газета 1813 г.: весь текст в режиме 3), а не настоящий
    текст поверх фона-картинки во весь лист (брошюра из InDesign: фон и фото — картинками, текст — видимый)."""
    kinds = [k for k, _ in page.get_bboxlog() if "text" in k]
    return bool(kinds) and kinds.count("ignore-text") >= 0.5 * len(kinds)


def graphic_page(page, segs, figs):
    """Лист-плакат: рисунки (краска вне текста и линий) — больше половины листа и на них ≥ 30 % подписей листа, или
    рисунок рассыпан на ≥ 5 кусков при малом тексте (обложка, разворот брошюры: картинки под маской, фоновые плашки);
    блок-схема Word — десятки картинок (градиент, тень), надписи лежат на них. figs — [(рамка, img)]."""
    import fitz
    if segs:
        imgs = [fitz.Rect(i["bbox"]) for i in page.get_image_info()]
        if len(imgs) >= GM.POSTER_PICTURES and sum(1 for s in segs if any(r.contains(fitz.Point(*_mid(s))) for r in imgs)) \
                >= 0.5 * len(segs):
            return True
    rects = [r for r, _ in figs]
    if not rects or not segs:
        return False
    area = sum(r.get_area() for r in rects) / page.rect.get_area()
    inside = sum(1 for s in segs if any(r.contains(fitz.Point(*_mid(s))) for r in rects)) / len(segs)
    return area >= GM.POSTER_AREA and inside >= GM.POSTER_CAPTIONS or len(rects) >= GM.POSTER_PIECES \
        and area >= GM.POSTER_PIECES_AREA and len(segs) <= GM.POSTER_TEXT


def nested(t, tables):
    """Рамки таблиц, лежащих внутри таблицы t (вложенные: рамка листа вокруг таблицы с заливкой). Таблица — (xs, ys, горизонтали,
    вертикали), как docx_layout._tables."""
    import fitz
    R = fitz.Rect(t[0][0], t[1][0], t[0][-1], t[1][-1])
    out = []
    for u in tables:
        if u != t:
            r = fitz.Rect(u[0][0], u[1][0], u[0][-1], u[1][-1])
            if R.contains(r) and r != R:
                out.append(r)
    return out


def chart_table(t, draw, inner=()):
    """Таблица-график: в её рамке — от трёх залитых фигур, не совпадающих с ячейками (полосы диаграммы Ганта, ромбы вех,
    столбики). t — (xs, ys, горизонтали, вертикали) таблицы, draw — [{"rect", "solid"}] листа. inner — рамки вложенных
    таблиц (nested): заливка их ячеек — не столбики внешней (рамка листа вокруг бланка с серыми клетками на 68 % листа была
    «графиком», и лист уходил снимком, 30TH50-80ÁTVGT л. 6–9)."""
    import fitz
    xs, ys = t[0], t[1]
    R = fitz.Rect(xs[0], ys[0], xs[-1], ys[-1])
    near = lambda v, grid: any(abs(v - g) < GM.DOUBLE for g in grid)
    xs, ys = list(xs) + [v[0] for v in t[3]], list(ys) + [h[0] for h in t[2]]
    n = 0
    for d in draw:
        if not d.get("solid"):
            continue
        q = fitz.Rect(d["rect"])
        if q.width < 2 or q.height < 2 or not R.contains(q) or q.get_area() > 0.5 * R.get_area():
            continue
        if near(q.x0, xs) and near(q.x1, xs) and near(q.y0, ys) and near(q.y1, ys):
            continue                                        # заливка ячейки
        if any(r.contains(q) for r in inner):
            continue                                        # заливка вложенной таблицы — её дело
        n += 1
    return n >= GM.CHART_SHAPES


def figure_kind(page, segs, figs, scan, charts, text_layer):
    """Вид листа для снимка целиком: "scan" — лист-скан с распознанными надписями; "scheme" — лист-плакат или
    блок-схема (graphic_page: не для скана с невидимым слоем распознавания — снимок брал скан без перевода); "chart" —
    таблицы-графики на треть листа и больше; None — рисунки по отдельности. Снимок — только у текстового слоя (text_layer:
    у листа распознанного PDF таблицы и рисунки собираются иначе)."""
    if not (text_layer and segs):
        return None
    if scan and any(s["loc"].get("ocr") for s in segs):
        return "scan"
    if not (scan and ocr_layer(page)) and graphic_page(page, segs, figs):
        return "scheme"
    if charts >= GM.CHART_SHARE * page.rect.get_area():
        return "chart"
    return None


# ------------------------------------------------------------------ схема: один набор признаков
def on_rule(r, rules, long=40):
    """Фигура — кусок длинной линии листа: тонкая и лежит на линии от long pt. Рамку и столбцы таблицы Word рисует по
    строкам — десятки отрезков в 11 pt один над другим; склеенные в одну линию, они — не «схема»."""
    hs, vs = rules
    if r.height <= 4.5 and any(abs(y - (r.y0 + r.y1) / 2) <= 2 and a - 2 <= r.x0 and r.x1 <= b + 2 and b - a >= long
                               for y, a, b in hs):
        return True
    return r.width <= 4.5 and any(abs(x - (r.x0 + r.x1) / 2) <= 2 and a - 2 <= r.y0 and r.y1 <= b + 2 and b - a >= long
                                  for x, a, b in vs)


def _linked(fa, a, fb, b):
    """Фигуры (fa, fb — «фигура», а не подпись) соединены: фигура с фигурой — ближе SCHEME_LINK, подпись с фигурой — ближе
    SCHEME_CAPTION; подписи — одного ряда (до SCHEME_ROW_GAP между ними) или строки одной подписи друг под другом."""
    if fa or fb:
        m = GM.SCHEME_LINK if fa and fb else GM.SCHEME_CAPTION
        return (a + (-m, -m, m, m)).intersects(b)
    h = min(a.height, b.height)
    return min(a.y1, b.y1) - max(a.y0, b.y0) >= 0.5 * h and max(a.x0, b.x0) - min(a.x1, b.x1) <= GM.SCHEME_ROW_GAP         or min(a.x1, b.x1) > max(a.x0, b.x0) and max(a.y0, b.y0) - min(a.y1, b.y1) <= 0.6 * h


def _candidates(items):
    """Пары индексов (i < j) возможных соседей — сеткой: фигура ищет соседей в радиусе SCHEME_LINK (подпись рядом найдёт
    её сама), подпись — в радиусе SCHEME_ROW_GAP (подписи одного ряда). Прямой перебор всех пар на листе с тысячами фигур
    (карта, схема КИП) — десятки секунд."""
    cell = 32.0
    grid = {}
    for i, (_, r) in enumerate(items):
        for cx in range(int(r.x0 // cell), int(r.x1 // cell) + 1):
            for cy in range(int(r.y0 // cell), int(r.y1 // cell) + 1):
                grid.setdefault((cx, cy), []).append(i)
    pairs = set()
    for i, (shape, r) in enumerate(items):
        m = GM.SCHEME_LINK if shape else max(GM.SCHEME_ROW_GAP, GM.SCHEME_CAPTION)
        for cx in range(int((r.x0 - m) // cell), int((r.x1 + m) // cell) + 1):
            for cy in range(int((r.y0 - m) // cell), int((r.y1 + m) // cell) + 1):
                pairs.update((min(i, j), max(i, j)) for j in grid.get((cx, cy), ()) if j != i)
    return sorted(pairs)


def scheme_regions(page, pno, drawings, segs, masked, tabs, figs, rules=((), ())):
    """Схемы листа: [(рамка, img)], img — область листа ({"clip", "page_number"}: сборка снимет её с переведённого PDF,
    надписи — переводом). Схема — связная группа фигур (замкнутые рамки, ромбы, линии-связки) и коротких подписей:
    фигуры соединяются, если ближе 20 pt, подпись — с фигурой ближе 10 pt, подписи одного ряда (до 60 pt между ними)
    и строки одной подписи друг под другом — между собой. Группа — схема, если
      * от 6 фигур и от 2 подписей, ни одна подпись не абзац (до 120 знаков: распознавание сливает подписи ряда в одну —
        «Turbina főgépész Primerköri főgépész»), и ни одна строка листа не задета краем группы;
      * есть связки — от 3 тонких линий длиннее 15 pt, и вертикальные, и горизонтальные («Layout format» Eskom,
        блок-схемы 3FI18; поля бланка — одни горизонтали, список условий с метками в рамочках — без вертикалей: это не
        схема), и группа не меньше 80×60 pt.
    Фигуры не берутся: во всю ширину или высоту листа (рамка листа), не рамки таблиц (tabs), не найденные рисунки (figs),
    не куски длинных линий листа (rules), не колонтитул (masked — сегменты вне тела: у распознанного листа его нет, тогда
    — полосы листа HEAD_EDGE/FOOT_EDGE) и не контуры букв внутри рамки сегмента."""
    import fitz
    W, H = page.rect.width, page.rect.height
    sb = [fitz.Rect(s["loc"]["bbox"]) for s in segs]
    off = [fitz.Rect(s["loc"]["bbox"]) for s in masked if s not in segs]
    rs = [fitz.Rect(d["rect"]) + (-0.5, -0.5, 0.5, 0.5) for d in drawings]       # линия без толщины — не пустая рамка
    rs = [r for r in rs if r.width < GM.PAGE_SHARE * W and r.height < GM.PAGE_SHARE * H and (r.width > 0.5 or r.height > 0.5)
          and not on_rule(r, rules)
          and not any(r.intersects(t) for t in tabs) and not any(r.intersects(f) for f in figs)
          and not any(r.intersects(o) for o in off)
          and not any((b + (-2, -2, 2, 2)).contains(r) for b in sb)
          and (bool(masked) or r.y1 >= GM.HEAD_EDGE * H and r.y0 <= GM.FOOT_EDGE * H)]
    # подпись-абзац (длиннее SCHEME_PARAGRAPH знаков) и строка шире полулиста — текст рядом со схемой: в группу не входит
    # (подводка над схемой и тело под ней не втягиваются), но мешает, если задевает её край или стоит внутри (ниже)
    items = [(True, r) for r in rs] + [(False, b) for s, b in zip(segs, sb)
                                       if len(s["text"].strip()) <= GM.SCHEME_PARAGRAPH and b.width <= GM.SCHEME_CAPTION_W * W]
    up = list(range(len(items)))

    def root(i):
        while up[i] != i:
            up[i] = up[up[i]]
            i = up[i]
        return i
    for i, j in _candidates(items):
        if _linked(*items[i], *items[j]):
            up[root(i)] = root(j)
    comps = {}
    for i, it in enumerate(items):
        comps.setdefault(root(i), []).append(it)
    out = []
    for c in comps.values():
        shapes = [r for f, r in c if f]
        if len(shapes) < GM.SCHEME_SHAPES or len(c) - len(shapes) < GM.SCHEME_CAPTIONS:
            continue
        vert = sum(r.width <= GM.ARROW_THIN and r.height >= GM.ARROW_LEN for r in shapes)
        horz = sum(r.height <= GM.ARROW_THIN and r.width >= GM.ARROW_LEN for r in shapes)
        R = fitz.Rect(shapes[0])
        for _, r in c:
            R |= r
        if not (vert and horz and vert + horz >= GM.SCHEME_ARROWS and R.width >= GM.SCHEME_SIZE[0]
                and R.height >= GM.SCHEME_SIZE[1]):
            continue
        inside = [s for s, b in zip(segs, sb) if R.contains((b.tl + b.br) / 2)]
        if any(len(s["text"].strip()) > GM.SCHEME_PARAGRAPH for s in inside) or any(b.intersects(R) and s not in inside
                                                                     for s, b in zip(segs, sb)):
            continue
        R = (R + (-3, -3, 3, 3)) & page.rect
        out.append((R, {"clip": [R.x0, R.y0, R.x1, R.y1], "page_number": pno + 1, "x0": R.x0, "x1": R.x1, "top": R.y0,
                        "bottom": R.y1}))
    return out


# ------------------------------------------------------------------ снимок листа
def clamp_zones(R, head, foot, hs, W, body=()):
    """Рамка снимка листа не заходит в колонтитулы оригинала: их печатает DOCX сам, а нижняя линия рамки шапки и обрывок
    «Oldalszám: 17.» под ней оставались в снимке лишней линией, и снимок был выше поля листа — ужимался (схема GF).
    head, foot — рамки сегментов колонтитулов (вне тела) в зонах листа; hs — горизонтали листа [(y, x0, x1)]; W — ширина
    листа; body — рамки сегментов тела листа. Линия рамки — ниже текста шапки (или выше текста подвала) не дальше
    GM.FRAME_GAP, от половины листа шириной, и между ней и текстом колонтитула нет текста тела (подчёркнутый заголовок
    рисунка под шапкой — тело, а не рамка шапки). Вызывается ПОСЛЕ полей снимка: поле в 4 pt иначе возвращало бы текст
    колонтитула. От снимка осталась полоска (< GM.CLAMP_MIN) — обрезки нет."""
    wide = [y for y, a, b in hs if b - a >= 0.5 * W]
    y0, y1 = R.y0, R.y1
    if head:
        bottom = max(h[3] for h in head)
        frame = [y + 2 for y in wide if bottom - 1 <= y <= bottom + GM.FRAME_GAP
                 and not any(b[3] > bottom and b[1] < y for b in body)]
        y0 = max(y0, bottom + 2, *frame)
    if foot:
        top = min(f[1] for f in foot)
        frame = [y - 2 for y in wide if top - GM.FRAME_GAP <= y <= top + 1
                 and not any(b[1] < top and b[3] > y for b in body)]
        y1 = min(y1, top - 2, *frame)
    if y1 - y0 >= GM.CLAMP_MIN:
        R.y0, R.y1 = y0, y1
    return R


# ------------------------------------------------------------------ формы-шаблоны (7б)
import json
import re
from pathlib import Path

# форма опознаётся по своему заголовку — короткому сегменту: лист регистрации, титул тестовой инструкции, лист изменений
FORM_TITLES = {
    "nyilv": re.compile(r"NYILV[ÁA]NT[ÁA]RT[ÓO]\s+LAP", re.I),
    "ciml": re.compile(r"OPERAT[ÍI]V\s+[ÜU]ZEMVITELI\s+TESZTEL", re.I),
    "modj": re.compile(r"M[ÓO]DOS[ÍI]T[ÁA]S\w*\s+JEGYZ[ÉE]K", re.I),
}
TITLE_MAX = 60                      # заголовок формы — короткая строка, а не ссылка на неё в тексте
LABELS_FILE = Path(__file__).resolve().parent / "form_labels.json"
_labels = {}


FORM_PAGES = 2                      # бланк — отдельный документ в один-два листа; раздел с тем же заголовком в длинном — нет


def known_form(segs):
    """Форма документа — "nyilv" | "ciml" | "modj" или None: заголовок формы — отдельный короткий сегмент документа не
    длиннее FORM_PAGES листов (раздел «MÓDOSÍTÁS JEGYZÉK» внутри инструкции — не бланк: его строки — текст инструкции).
    Листы должны быть известны (у сегментов DOCX page = None — длину не проверить, формы нет) и форма — одна: титул одной
    формы и лист другой в одном документе — не бланк."""
    if not segs or any(s["page"] is None for s in segs) or max(s["page"] for s in segs) >= FORM_PAGES:
        return None
    found = {form for s in segs if len(s["text"].strip()) <= TITLE_MAX
             for form, rx in FORM_TITLES.items() if rx.search(s["text"])}
    return found.pop() if len(found) == 1 else None


def strip_bars(text):
    """Без пробелов и «|» по краям: линии ячеек, прочитанные распознаванием как знак (посреди строки «|» — текст)."""
    while True:
        t = text.strip().strip("|")
        if t == text:
            return t
        text = t


def label_key(text):
    """Подпись без «|» по краям и лишних пробелов — ключ словаря форм."""
    return re.sub(r"\s+", " ", strip_bars(text))


def is_label(text):
    """Печатная подпись формы, а не значение бланка: от двух букв, без цифр, фамилий и неизвестных сокращений из заглавных
    (значения — пиксели оригинала: pagefit.pixels_only), без точек-заполнителей."""
    import pagefit as PF
    if len(text) > 70 or re.search(r"\d|[.…]{3,}", text) or len(re.findall(r"[^\W\d_]", text)) < 2:
        return False
    known = PF.known_words()
    return not any(PF._kind(t, known) in ("name", "num", "acr") for t in text.split())


def form_labels():
    """Словарь подписей форм {форма: {подпись: перевод}} (form_labels.json, собирается tools/form_labels.py)."""
    if "d" not in _labels:
        try:
            _labels["d"] = json.loads(LABELS_FILE.read_text(encoding="utf-8"))
        except (OSError, ValueError) as e:
            print(f"figures: словарь форм не прочитан: {type(e).__name__}: {e}"[:200], flush=True)
            _labels["d"] = {}
    return _labels["d"]


PROTECTED = ("edited", "tm", "copied", "polish", "ru_pdf")      # перевод редактора, память, без перевода, шлифовка, сокращение для PDF


def apply_form_labels(segs, res):
    """Подписи известной формы — по словарю: один перевод на подпись во всех документах формы (портал давал «Да»/«да»/
    «ДА», «Подпись:»/«Подпись: |»). Не трогаются перевод редактора (edited), память переводов (tm), сегмент без
    изменения (copied), шлифованный (polish: словарь заказчика) и сокращённый для PDF (ru_pdf). → (res, сколько подписей заменено, форма); res — новый словарь, исходный не меняется."""
    form = known_form(segs)
    labels = form_labels().get(form) or {} if form else {}
    if not labels:
        return res, 0, form
    out, n = dict(res), 0
    for s in segs:
        r = res.get(str(s["id"]))
        k = label_key(s["text"])
        if not r or not r.get("ru") or k not in labels or any(r.get(f) for f in PROTECTED):
            continue
        if r["ru"] != labels[k]:
            out[str(s["id"])] = dict(r, ru=labels[k])
            n += 1
    return out, n, form


# ------------------------------------------------------------------ перевод поверх скана (7в)
def drawn_text(ru):
    """Перевод для рисования поверх скана без «|» по краям: линия ячейки, прочитанная как знак, рисовалась буквой
    («Дата: 2022-01-11 |»), хотя сама линия стёрта плашкой строки."""
    return strip_bars(ru) or ru


# флажок бланка («☐ igen» / «☒ nem»), прочитанный распознаванием буквами: «CI нет», «X да», «Fr да» — плашка стирала
# флажок, и вместо него рисовались буквы (57 надписей в 36 документах корпуса). Флажок — пиксели оригинала: плашка и
# перевод начинаются после него, а перевод — само «да»/«нет»
CHECKBOX_WORD = re.compile(r"(?P<mark>\S{1,2})\s+(?P<word>igen|nem)", re.I)
ARTICLES = {"a", "az"}                                          # «a nem», «Az igen» — слова, не флажок
SHORT_WORDS = {"és", "is", "ha", "ne", "de", "ki", "be", "el", "le", "fel"}   # «Ki nem»; ЗАГЛАВНЫМИ («KI igen») — флажок
YES_NO = {"igen": "да", "nem": "нет"}


YES_NO_RU = re.compile(r"(?:\S{1,6}\s+)?(да|нет)", re.I)       # перевод надписи — «<значок> да|нет», а не фраза («Это не»)


def split_checkboxes(segs, res):
    """(segs, res) без флажков-букв: у надписи «<1–2 знака> igen|nem» (распознавание, одна строка), перевод которой —
    «<значок> да|нет», рамка начинается после флажка — на середине пробела между ним и словом (по доле знаков), а перевод —
    «да»/«нет» (регистр как у слова). Не трогаются: перевод не на русский, редактора (edited), память (tm), слова-фразы
    («a nem», «Ki nem»; заглавными «KI igen» — флажок). Новые списки; исходные не меняются."""
    out_s, out_r = list(segs), dict(res)
    for i, s in enumerate(segs):
        loc, r = s["loc"], res.get(str(s["id"]))
        if not loc.get("ocr") or len(loc.get("frags") or []) != 1 or not r or r.get("edited") or r.get("tm")                 or not YES_NO_RU.fullmatch((r.get("ru") or "").strip()):
            continue                                                  # «да»/«нет» — только для перевода на русский
        t = s["text"].strip()
        m = CHECKBOX_WORD.fullmatch(t)
        low = m["mark"].lower() if m else ""
        if not m or low in ARTICLES or low in SHORT_WORDS and m["mark"] != m["mark"].upper():
            continue
        x0, y0, x1, y1 = loc["frags"][0]
        cut = x0 + (x1 - x0) * (len(m["mark"]) + 0.5) / len(t) + 1      # середина пробела + 1 pt: плашка (cut − 2) флажка не касается
        moved = lambda b: [max(cut, b[0])] + list(b[1:])
        out_s[i] = dict(s, text=m["word"], loc=dict(loc, frags=[moved(loc["frags"][0])], bbox=moved(loc["bbox"]),
                                                    **({"avail": moved(loc["avail"])} if "avail" in loc else {})))
        ru = YES_NO[m["word"].lower()]
        w = m["word"]
        out_r[str(s["id"])] = dict(r, ru=ru.upper() if w.isupper() else ru.capitalize() if w[0].isupper() else ru)
    return out_s, out_r


TOKEN = re.compile(r"[^\W_](?:[^\s]*[^\W_])?")
_CONFUSED = str.maketrans({"O": "0", "I": "1", "L": "1", "S": "5", "B": "8", "Z": "2", "|": "1"})   # путаница распознавания


_LOOKALIKE = str.maketrans("АВСЕНКМОРТХ", "ABCEHKMOPTX")        # кириллица, похожая на латиницу: «В2» вместо «B2» (rus+eng)


def _tok(w):
    """Слово для сверки: заглавные, без знаков по краям, кириллические двойники латиницы — латиницей; в слове с цифрой
    «O», «I», «L», «S», «B», «Z» — как «0», «1», «1», «5», «8», «2» (путаница распознавания), «,» — как «.»,
    «−», «–», «—» — как «-»."""
    w = w.strip(".,;:()[]{}|/-–—−_«»\"'“”„\\").upper()
    if re.search(r"\d", w):
        w = re.sub("[–—−]", "-", w.translate(_LOOKALIKE).translate(_CONFUSED).replace(",", "."))
    return w


def code_map(text):
    """Коды и числа текста (слова с цифрой, от двух знаков) — то, что при рисовании не должно потеряться:
    {слово для сверки: слово как в тексте} — для сообщения глазами находимо."""
    out = {}
    for w in TOKEN.findall(text):
        t = _tok(w)
        if re.search(r"\d", t) and len(t) >= 2:
            out.setdefault(t, w.strip(".,;:()[]{}|")[:40])
    return out


def codes_of(text):
    return set(code_map(text))


def _near(a, b):
    """Слова совпадают, или различаются одним знаком (слова от 5 знаков): распознавание мелкой печати ошибается."""
    if a == b:
        return True
    if len(a) < 5 or len(b) < 5 or abs(len(a) - len(b)) > 1:
        return False
    i = next((k for k, (x, y) in enumerate(zip(a, b)) if x != y), min(len(a), len(b)))
    return a[i + 1:] == b[i + 1:] or a[i:] == b[i + 1:] or a[i + 1:] == b[i:]


def missing_codes(expected, seen_words):
    """Какие из expected (codes_of) не найдены среди слов seen_words (второе распознавание листа после рисования): слово
    совпало, или различается одним знаком, или длинный код (от 5 знаков) распознан частями («30YCGT» → «30Y» «CGT»)."""
    seen = [_tok(w) for w in seen_words]
    blob = "".join(seen)
    return sorted(e for e in expected if e not in seen and not (len(e) >= 5 and e in blob)
                  and not any(_near(e, w) for w in seen))


def ocr_words(page, dpi=200, lang="rus+eng"):
    """Слова листа вторым распознаванием (Tesseract, разреженный текст): список слов; сбой — None (сверка пропущена)."""
    import subprocess
    import tempfile
    try:
        with tempfile.TemporaryDirectory() as td:
            png = Path(td) / "p.png"
            page.get_pixmap(dpi=dpi, colorspace=__import__("fitz").csGRAY).save(str(png))
            r = subprocess.run(["tesseract", str(png), "stdout", "-l", lang, "--psm", "11", "tsv"],
                               capture_output=True, text=True, encoding="utf-8", timeout=180)
    except (OSError, subprocess.SubprocessError) as e:
        print(f"figures: второе распознавание листа не выполнено: {type(e).__name__}: {e}"[:200], flush=True)
        return None
    if r.returncode:
        print(f"figures: tesseract: {(r.stderr or '').strip()[:160]}", flush=True)
        return None
    out = []
    for ln in r.stdout.splitlines()[1:]:
        c = ln.split("\t")
        if len(c) >= 12 and c[0] == "5" and c[11].strip() and float(c[10]) >= 0:
            out.append(c[11])
    return out


def verify_page(page, drawn, pno):
    """Сверка листа после рисования: коды и числа нарисованного перевода (drawn — [текст]) должны читаться на листе
    вторым распознаванием. → (строка, громко) или None (всё на месте, нечего сверять, нет Tesseract или сверка не
    удалась: справочная проверка не роняет сборку). Громко — не найдено от VERIFY_LOST кодов и от VERIFY_SHARE
    проверенных (info["verify"] → «ВНИМАНИЕ» сборки); меньше — просто сведения (info["verify_notes"]): Tesseract
    ошибается на мелкой печати и сам, 37 из 65 листов корпуса теряли по одному слову. Лист не меняется: расхождение —
    повод посмотреть глазами (закрыто рамкой, налезло)."""
    want = {}
    for t in drawn:
        for k, w in code_map(t).items():
            want.setdefault(k, w)
    if not want:
        return None
    try:
        words = ocr_words(page)
        if words is None:
            return None
        lost = missing_codes(set(want), words)
    except Exception as e:                                           # noqa: BLE001 — справочная проверка, сборку не роняет
        print(f"figures: сверка листа {pno + 1} не удалась: {type(e).__name__}: {e}"[:200], flush=True)
        return None
    if not lost:
        return None
    shown = [want[k] for k in lost]
    msg = (f"лист {pno + 1} скана: после рисования перевода не найдено при втором распознавании {len(lost)} из "
           f"{len(want)} кодов и чисел — {', '.join(shown[:12])}{' …' if len(shown) > 12 else ''}")
    return msg, len(lost) >= GM.VERIFY_LOST and len(lost) >= GM.VERIFY_SHARE * len(want)
