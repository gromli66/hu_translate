# -*- coding: utf-8 -*-
"""Раскладка страницы PDF: фрагменты строк → абзацы (единицы перевода) с рамками.

Один алгоритм для текстового слоя и для слов Tesseract (фрагменты сканов — segments.ocr_frags):
- строка режется на фрагменты по горизонтальному разрыву > GAP (столбцы таблиц без вертикальных линий);
- фрагмент присоединяется к открытому абзацу, если он ниже вплотную, перекрывается по x,
  между ними нет горизонтальной линейки, начертание то же и фрагмент не начинается с маркера пункта;
- для вставки перевода рамка абзаца расширяется вправо/вниз до соседей, линеек и полей."""
import re
import fitz

GAP = 30.0          # выключка по ширине в узких ячейках даёт пробелы до ~25 pt


def family(font):
    """Семейство шрифта для вставки перевода: моноширинный / без засечек / с засечками (по имени шрифта PDF)."""
    f = font.lower()
    if any(k in f for k in ("courier", "mono", "consol", "lucidaconsole")):
        return "monospace"
    if any(k in f for k in ("arial", "helvetica", "sans", "verdana", "tahoma", "calibri")):
        return "sans-serif"
    return "serif"
MARK = re.compile(r"^\s*(\(?\d+(\.\d+)*\.?\)?|[a-zA-Z]\.?\)|[A-Z]\.\d*\.?|[-–•▪]|\d+/\d+\.|\d+\./|\[\d+\])\s")


def code_line(t):
    """Строка из кодов («04.TH10D001 KM1/1», «30BW07», «03/A TJ22D001»): каждое слово — с цифрой или аббревиатура до
    4 прописных, слов с цифрами не меньше половины. Такие строки — ячейки перечня без линеек, а не строки абзаца:
    столбец кодов склеивался в один абзац, код и его значение — в один кусок (3PR42 л. 112, 17.10)."""
    toks = [x for x in t.split() if re.search(r"\w", x)]
    dig = [x for x in toks if re.search(r"\d", x)]
    return bool(toks) and 2 * len(dig) >= len(toks) and all(
        re.search(r"\d", x) or (x.isupper() and len(re.sub(r"\W", "", x)) <= 4) for x in toks)


CODE_GAP = 8.0      # код и значение в строке перечня («TK41D001 ⇥ KM2/1»: 10–24 pt, слова прозы — 3–4 pt) — разные куски


def _union(bbs):
    return [min(b[0] for b in bbs), min(b[1] for b in bbs), max(b[2] for b in bbs), max(b[3] for b in bbs)]


def _ink_x1(f):
    """Правый край кусочка по видимым знакам: хвостовой пробел моноширинной строки («…kézzel tudja ») доходил до линии
    столбца и закрывал её от проверки склейки (3SZ17: строки двух столбцов склеивались в один кусок)."""
    t, bb = f["t"], f["bbox"]
    k = len(t) - len(t.rstrip())
    return bb[2] - k * (bb[2] - bb[0]) / len(t) if t and k else bb[2]


def _vrule_in(vs, xa, xb, y0, y1):
    return any(xa - 1 <= x <= xb + 1 and min(y1, ry1) - max(y0, ry0) > 0.5 * (y1 - y0) for ry0, ry1, x in vs)


BULLET = re.compile(r"^[•▪◦·o\-–]$")


def hidden_text(page):
    """Рамки текста, которого не видно: невидимый слой (режим 3) и текст, нарисованный раньше картинки, которая его
    целиком накрывает. Титул UTS (15.10): шапка и колонтитул лежат под фоном во весь лист — перевод печатался поверх
    титула. Порядок отрисовки — по номеру записи get_bboxlog; невидимый слой там — «ignore-text».
    get_texttrace не годится: в PyMuPDF 1.27 на сотне страниц роняет процесс (none_dealloc, 15.10).
    Координаты — как у get_text.
    Прячет только крупная (от половины листа) картинка без маски прозрачности; если «спрятанным» вышло больше 70 %
    текста страницы — это водяной знак поверх всего, не фон: прозрачный знак, полупрозрачная печать и картинка,
    обрезанная в Word, съедали текст молча (красная команда 15.10)."""
    log = page.get_bboxlog()
    area = page.rect.get_area()
    info = page.get_image_info()

    def opaque_big(r):
        if r.get_area() < 0.5 * area:
            return False
        same = [ii for ii in info if abs(fitz.Rect(ii["bbox"]).x0 - r.x0) < 1 and abs(fitz.Rect(ii["bbox"]).y1 - r.y1) < 1]
        return not any(ii.get("has-mask") for ii in same)

    imgs = [(i, fitz.Rect(b)) for i, (k, b) in enumerate(log) if k == "fill-image" and opaque_big(fitz.Rect(b))]
    texts = [(i, k, fitz.Rect(b)) for i, (k, b) in enumerate(log) if "text" in k]
    under = [r for i, k, r in texts if k != "ignore-text" and any(j > i and ir.contains(r) for j, ir in imgs)]
    shown = [r for i, k, r in texts if k != "ignore-text"]
    if shown and len(under) > 0.7 * len(shown):
        under = []
    return [r for i, k, r in texts if k == "ignore-text"] + under


def _covered(rects, bb, share=0.5):
    r = fitz.Rect(bb); a = r.get_area() or 1
    return sum((r & h).get_area() for h in rects if h.intersects(r)) >= share * a


def _lead(t):
    """Перед значком только маркер пункта или 1–3 знака («• ✓: minden…», «10.5 : még…»): строка рвётся на значке.
    Значок посреди фразы («az indító (▶) gombbal») строку не рвёт: два куска одного абзаца с продолжениями
    ложились друг на друга (UTS стр. 39, 15.10) — значок под переводом меньшее зло."""
    t = t.strip()
    return len(t) <= 3 or bool(BULLET.match(t))


def _image_between(imgs, xa, xb, y0, y1):
    """Значок (✓/✗) в разрыве строки рвёт строку: перевод не наезжает на значок."""
    return any(r.x0 < xb + 1 and r.x1 > xa - 1 and min(y1, r.y1) - max(y0, r.y0) > 0.3 * (y1 - y0) for r in imgs)


def _cut_spans(spans, vs, rot):
    """Куски строки текстового слоя (rawdict) с текстом; кусок поперёк вертикали таблицы («Индекс ПЗОбразец
    Обозначения» — одна строка PDF через линию столбца, шапка таблицы НГТУ: два столбца сливались в один, 19.10) —
    частями по линии, по серединам знаков."""
    out = []
    for s in spans:
        chars = s.get("chars") or []
        bb = fitz.Rect(s["bbox"]) * rot
        xs = sorted(x for y0, y1, x in vs if bb.x0 + 1 < x < bb.x1 - 1
                    and min(bb.y1, y1) - max(bb.y0, y0) > 0.5 * (bb.y1 - bb.y0)) if chars else []
        groups, cur, k = [], [], 0
        for c in chars:
            r = fitz.Rect(c["bbox"]) * rot
            while k < len(xs) and (r.x0 + r.x1) / 2 > xs[k]:
                groups.append(cur)
                cur, k = [], k + 1
            cur.append(c)
        groups.append(cur)
        if len([g for g in groups if g]) <= 1:
            out.append(dict(s, text="".join(c["c"] for c in chars)))
            continue
        for g in groups:
            if g:
                # рамка куска — по знакам без пробелов: пробел у линии столбца делал зазор нулевым, линия выпадала из
                # проверки склейки (_vrule_in), и куски склеивались обратно («47,05281 19,11048 Kunszentmiklós» —
                # три столбца таблицы HAKSER в одной ячейке, 22.10)
                ink = [c for c in g if c["c"].strip()] or g
                r = fitz.Rect(ink[0]["bbox"])
                for c in ink[1:]:
                    r |= fitz.Rect(c["bbox"])
                out.append(dict(s, text="".join(c["c"] for c in g), bbox=tuple(r)))
    return out


NUMBERISH = re.compile(r"^(?:[\dIVXivx]+[.)]?)+$|^\S+:$")


def _split_columns(frags, words):
    """Кусок строки, склеенный через просвет столбца таблицы без линий («Dátum ⇥ Határozat száma»: 27 pt при пороге
    склейки GAP 30 pt) — надвое. Просвет — столбцовый, если он не меньше полутора кеглей и хотя бы вчетверо больше
    всех прочих просветов куска (выключка по ширине растягивает их все поровну); в куске от трёх слов. Номер пункта,
    номер листа и подпись с двоеточием по краям просвета — не столбец (E43 корпуса, 22.10).
    words — [(рамка как видна, слово)] листа."""
    out = []
    for f in frags:
        x0, y0, x1, y1 = f["bbox"]
        mid = (y0 + y1) / 2
        ws = sorted(((r, t) for r, t in words if r.y0 <= mid <= r.y1 and x0 - 1 <= (r.x0 + r.x1) / 2 <= x1 + 1),
                    key=lambda w: w[0].x0)
        size = max((c[1] for c in f["chars"]), default=10)
        if len(ws) < 3 or " ".join(t for _, t in ws) != " ".join(f["t"].split()):
            out.append(f)
            continue
        gaps = [ws[i + 1][0].x0 - ws[i][0].x1 for i in range(len(ws) - 1)]
        j = max(range(len(gaps)), key=gaps.__getitem__)
        rest = max(g for i, g in enumerate(gaps) if i != j)
        left, right = " ".join(t for _, t in ws[:j + 1]), " ".join(t for _, t in ws[j + 1:])
        if gaps[j] < 1.5 * size or gaps[j] < 4 * max(rest, 1) or NUMBERISH.match(left) or NUMBERISH.match(right):
            out.append(f)
            continue
        out.append(dict(f, t=left, bbox=list(_union([list(r) for r, _ in ws[:j + 1]])), chars=list(f["chars"])))
        out.append(dict(f, t=right, bbox=list(_union([list(r) for r, _ in ws[j + 1:]])), chars=list(f["chars"])))
    return out


SYMBOLIC = re.compile(r"Symbol|Wingdings|Webdings", re.I)


def span_text(s):
    """Текст куска; пустой глиф шрифта Symbol/Wingdings (код 0, в текстовом слое — «�») — флажок «□»: в Word он
    рисуется рамкой, а модель и сборка получали мусорный знак (30RR2126GT: 18 флажков бланка, 28.10)."""
    t = s["text"]
    return t.replace("�", "□") if "�" in t and SYMBOLIC.search(s["font"]) else t


def frags_from_text(page, vs=()):
    """Фрагменты из текстового слоя: [{t, bbox, size, bold}]. vs — вертикальные линейки.
    Координаты — как страница видна (с учётом /Rotate), как у слов OCR с картинки: PyMuPDF отдаёт текст в системе
    неповёрнутой страницы (32RDGT, листы регистрации с поворотом 270°, 07.10).
    - Невидимый текст (hidden_text) пропускается.
    - Маленькие картинки внутри строки (значки) рвут строку.
    - Маркер пункта (•, o, –), отделённый от текста отступом, — свой фрагмент без перевода: перевод встаёт на место
      текста пункта с висячим отступом (UTS, 15.10: перевод начинался с маркера, отступ терялся). Текст такого пункта
      помечен item — к абзацу выше не приклеивается, как строка с маркером."""
    d = page.get_text("rawdict", flags=fitz.TEXT_PRESERVE_WHITESPACE | fitz.TEXT_MEDIABOX_CLIP)
    rot = page.rotation_matrix
    hidden = hidden_text(page)
    area = page.rect.get_area()
    icons = [fitz.Rect(i["bbox"]) * rot for i in page.get_image_info() if fitz.Rect(i["bbox"]).get_area() < 0.01 * area]
    out = []
    for b in d["blocks"]:
        if b.get("type") != 0:
            continue
        for ln in b["lines"]:
            if abs((fitz.Point(ln["dir"]) * rot - fitz.Point(0, 0) * rot).y) > 0.1:   # повёрнутый текст не трогаем
                continue
            cur = None
            for s in _cut_spans(ln["spans"], vs, rot):
                txt = span_text(s)
                if hidden and _covered(hidden, s["bbox"]):
                    continue
                if not txt.strip():
                    if cur is not None:
                        cur["t"] += txt
                    continue
                bb = list(fitz.Rect(s["bbox"]) * rot)
                bold = bool(s["flags"] & 16) or "Bold" in s["font"]
                marker = cur is not None and bool(BULLET.match(cur["t"].strip())) and bb[0] - cur["bbox"][2] > 3
                if cur is not None and bb[0] - cur["bbox"][2] <= GAP and not marker \
                        and not (bb[0] - cur["bbox"][2] >= CODE_GAP and code_line(cur["t"]) and code_line(txt)) \
                        and not _vrule_in(vs, _ink_x1(cur), bb[0], bb[1], bb[3]) \
                        and not (_lead(cur["t"]) and _image_between(icons, cur["bbox"][2], bb[0], bb[1], bb[3])):
                    cur["t"] += txt
                    cur["bbox"] = _union([cur["bbox"], bb])
                    cur["chars"].append((len(txt.strip()), s["size"], bold, family(s["font"])))
                else:
                    if cur is not None:
                        out.append(cur)
                    cur = {"t": txt, "bbox": bb, "chars": [(len(txt.strip()), s["size"], bold, family(s["font"]))]}
                    if marker:
                        cur["item"] = True
            if cur is not None:
                out.append(cur)
    # выключенная строка: PyMuPDF отдаёт каждое слово отдельной «строкой» — склеить по базовой линии
    out.sort(key=lambda f: (round((f["bbox"][1] + f["bbox"][3]) / 4), f["bbox"][0]))
    merged = []
    for f in out:
        m = merged[-1] if merged else None
        if m is not None and abs(m["bbox"][1] - f["bbox"][1]) < 2.5 and abs(m["bbox"][3] - f["bbox"][3]) < 2.5 \
                and 0 <= f["bbox"][0] - m["bbox"][2] <= GAP and not _vrule_in(vs, _ink_x1(m), f["bbox"][0], f["bbox"][1], f["bbox"][3]) \
                and not (f["bbox"][0] - m["bbox"][2] >= CODE_GAP and code_line(m["t"]) and code_line(f["t"])) \
                and not (_lead(m["t"]) and _image_between(icons, m["bbox"][2], f["bbox"][0], f["bbox"][1], f["bbox"][3])) \
                and not (BULLET.match(m["t"].strip()) and f["bbox"][0] - m["bbox"][2] > 3):
            m["t"] = m["t"].rstrip() + " " + f["t"].lstrip()
            m["bbox"] = _union([m["bbox"], f["bbox"]]); m["chars"] += f["chars"]
        else:
            if m is not None and BULLET.match(m["t"].strip()) and abs(m["bbox"][3] - f["bbox"][3]) < 2.5 \
                    and 3 < f["bbox"][0] - m["bbox"][2] <= 2 * GAP:
                f["item"] = True                        # маркер и текст пункта — разные «строки» PyMuPDF
            merged.append(f)
    out = _split_columns(merged, [(fitz.Rect(w[:4]) * rot, w[4]) for w in page.get_text("words")])
    for f in out:
        n = sum(c[0] for c in f["chars"]) or 1
        f["size"] = max(f["chars"], key=lambda c: c[0])[1]
        f["bold"] = sum(c[0] for c in f["chars"] if c[2]) > n / 2
        fam = {}
        for c in f["chars"]:
            fam[c[3]] = fam.get(c[3], 0) + c[0]
        f["family"] = max(fam, key=fam.get)
        f["t"] = f["t"].strip()
        del f["chars"]
    return [f for f in out if f["t"]]




def rules(page):
    """Горизонтальные и вертикальные линейки (тонкие залитые прямоугольники и отрезки); координаты — как страница видна."""
    hs, vs = [], []
    rot = page.rotation_matrix
    for dr in page.get_drawings():
        r = dr["rect"] * rot
        if r.height <= 2.5 and r.width >= 15:
            hs.append((r.x0, r.x1, (r.y0 + r.y1) / 2))
        elif r.width <= 2.5 and r.height >= 8:
            vs.append((r.y0, r.y1, (r.x0 + r.x1) / 2))
        else:
            for it in dr["items"]:
                if it[0] == "l":
                    p, q = it[1] * rot, it[2] * rot
                    if abs(p.y - q.y) < 1 and abs(p.x - q.x) >= 15:
                        hs.append((min(p.x, q.x), max(p.x, q.x), p.y))
                    elif abs(p.x - q.x) < 1 and abs(p.y - q.y) >= 8:
                        vs.append((min(p.y, q.y), max(p.y, q.y), p.x))
                elif it[0] == "re":
                    # залитый прямоугольник тоже даёт рамку: в формах 30RDGT заливка ячеек — единственный разделитель
                    # («<= 2400» / «mm»); цена — шапка с заливкой по строкам рвётся («Fejlesztő / környezet», UTS)
                    rr = it[1] * rot
                    if rr.width >= 15 and rr.height > 2.5:     # рамка ячейки: четыре стороны
                        hs += [(rr.x0, rr.x1, rr.y0), (rr.x0, rr.x1, rr.y1)]
                        vs += [(rr.y0, rr.y1, rr.x0), (rr.y0, rr.y1, rr.x1)]
    return hs, vs


def _rule_between(hs, x0, x1, ya, yb):
    for rx0, rx1, y in hs:
        if ya - 0.5 <= y <= yb + 0.5 and min(x1, rx1) - max(x0, rx0) > 5:
            return True
    return False


def _in_cell(vs, bb):
    """Фрагмент зажат вертикальными линейками слева и справа (ячейка таблицы)."""
    x0, y0, x1, y1 = bb
    left = any(x <= x0 + 1 and x0 - x < 250 and ry0 <= y1 and ry1 >= y0 for ry0, ry1, x in vs)
    right = any(x >= x1 - 1 and x - x1 < 250 and ry0 <= y1 and ry1 >= y0 for ry0, ry1, x in vs)
    return left and right


def _right_limit(frags, vs, bb, page_right):
    """Где кончается место справа от строки: вертикальная линейка или соседний фрагмент на той же высоте."""
    x0, y0, x1, y1 = bb
    lim = page_right
    for ry0, ry1, x in vs:
        if x >= x1 - 1 and min(y1, ry1) - max(y0, ry0) > 0.5 * (y1 - y0):
            lim = min(lim, x)
    for f in frags:
        fx0, fy0, fx1, fy1 = f["bbox"]
        if fx0 >= x1 + 1 and min(y1, fy1) - max(y0, fy0) > 0.5 * (y1 - y0):
            lim = min(lim, fx0)
    return lim


def _partner(frags, vs, bb, left=False):
    """Справа на той же линии есть другой фрагмент и линейки между ними нет (вторая колонка списка без таблицы);
    left — слева (значение против своей подписи)."""
    x0, y0, x1, y1 = bb
    same = lambda f: min(y1, f["bbox"][3]) - max(y0, f["bbox"][1]) > 0.5 * (y1 - y0)
    if left:
        return any(f["bbox"][2] <= x0 - 5 and same(f) and not _vrule_in(vs, f["bbox"][2], x0, y0, y1) for f in frags)
    return any(f["bbox"][0] >= x1 + 5 and same(f) and not _vrule_in(vs, x1, f["bbox"][0], y0, y1) for f in frags)


ANCHOR = re.compile(r"^\(?\d+(?:\.\d+)*\.?\)?$")


def _row_anchor(frags, bb):
    """Якорь строки таблицы без горизонтальных линий: ближайший слева на той же линии кусок — код оборудования или номер
    пункта (через вертикальные линейки колонок — они делят колонки, а не строки). Возвращает его рамку или None."""
    x0, y0, x1, y1 = bb
    best = None
    for f in frags:
        fb = f["bbox"]
        if fb[2] <= x0 - 5 and min(y1, fb[3]) - max(y0, fb[1]) > 0.5 * (y1 - y0) and (best is None or fb[2] > best[2]):
            best = fb
            best_t = f["t"].strip()
    if best is None:
        return None
    ok = bool(ANCHOR.match(best_t)) or (code_line(best_t) and len(best_t.split()) <= 3 and bool(re.search(r"\d", best_t)))
    return tuple(best) if ok else None


def mid_phrase(prev, nxt):
    """Строка оборвалась на середине фразы: в ней ≥ 3 слов, она кончается буквой или запятой, следующая начинается
    строчной буквой. Одиночные значения столбика («rendben» ×4) и пункты с точкой не подходят."""
    prev, nxt = prev.rstrip(), nxt.lstrip()
    return len(prev.split()) >= 3 and bool(re.search(r"[^\W\d_]$|,$", prev)) and nxt[:1].isalpha() and nxt[:1].islower()


def group(frags, hs, vs=(), form=False):
    """form=True (скан формы): строка, кончающаяся двоеточием, — подпись поля, к ней не приклеивается следующая."""
    page_right = max((f["bbox"][2] for f in frags), default=0)
    frags = sorted(frags, key=lambda f: (round(f["bbox"][1] / 3), f["bbox"][0]))
    # слева на той же строке — номер пункта отдельным куском («1.6.» ⇥ «FOT szintcsökkenés» в оглавлении 3SZ19): кусок
    # начинает новый абзац — иначе название приклеивалось к названию пункта выше, а у «1.6.» оставалось пустым (17.10)
    label = re.compile(r"^\(?\d+(\.\d+)*\.?\)?$|^[a-zA-Z][.)]$")
    nums = [f["bbox"] for f in frags if label.match(f["t"].strip())]
    used = set()                              # номер открывает абзац один раз — для первой строки против него

    def labeled(bb):
        return any(n[2] < bb[0] and bb[0] - n[2] < 120 and min(n[3], bb[3]) - max(n[1], bb[1]) > 0.5 * (bb[3] - bb[1])
                   for n in nums)

    def continues(f):
        """Предыдущая строка этого же столбца оборвана на середине фразы, начинается без номера и выровнена с этой."""
        x0, y0 = f["bbox"][0], f["bbox"][1]
        for q in reversed(pars[-12:]):
            last = q["frags"][-1]["bbox"]
            size = max(q["size"], f["size"])
            if -0.35 * size <= y0 - last[3] <= 0.6 * size and min(f["bbox"][2], q["bbox"][2]) - max(x0, q["bbox"][0]) > 0:
                return abs(x0 - last[0]) < 3 and mid_phrase(q["frags"][-1]["t"], f["t"]) and not labeled(q["frags"][0]["bbox"])
        return False

    def numbered(bb):
        """Номер слева на той же строке — этот кусок начинает абзац. Номер по центру ячейки в две строки («2.8.» против
        строк «…участниками и» / «blokkvezénylő között.») лежит против обеих — начало только первой, вторая продолжает
        (30RR2126GT: предложение переводилось двумя обрывками, 28.10)."""
        for k, n in enumerate(nums):
            if n[2] < bb[0] and bb[0] - n[2] < 120 and min(n[3], bb[3]) - max(n[1], bb[1]) > 0.5 * (bb[3] - bb[1]):
                if k in used:
                    continue
                used.add(k)
                return True
        return False
    pars = []
    for f in frags:
        x0, y0, x1, y1 = f["bbox"]
        best = None
        # маркер пункта отдельным фрагментом — сам по себе: маркеры соседних пунктов слипались в один «• •», и пункт
        # оставался без своего (Eskom 331-165, 17.10)
        plain = not MARK.match(f["t"]) and not f.get("item") and not BULLET.match(f["t"].strip())
        starts = plain and numbered(f["bbox"])
        # номер строки таблицы напротив последней строки ячейки («0.» против «…szintváltozásának módosítása…»): у первой
        # строки абзаца своего номера нет — продолжение фразы, а не новый пункт (30TH50, 28.10)
        if starts and continues(f):
            starts = False
        if plain and not starts:
            for p in reversed(pars[-12:]):
                if len(p["frags"]) == 1 and BULLET.match(p["frags"][0]["t"].strip()):
                    continue
                px0, py0, px1, py1 = p["bbox"]
                last = p["frags"][-1]["bbox"]
                gap = y0 - last[3]
                size = max(p["size"], f["size"])
                # рамка строки выше шага строк (моноширинный шрифт 3SZ17: рамка 15 pt при шаге 10 pt) — рамки соседних
                # строк перекрываются; строка ниже верха предыдущей на шаг — продолжение, а не та же строка
                tall = gap < 0 and -0.7 * size <= gap and y0 - last[1] >= 0.85 * size
                if not (-0.35 * size <= gap <= 0.6 * size or tall):
                    continue
                if min(x1, px1) - max(x0, px0) <= 0:
                    continue
                if p["bold"] != f["bold"] or abs(p["size"] - f["size"]) > 1.5:
                    continue
                if _rule_between(hs, x0, x1, last[3], y0):
                    continue
                if form and p["frags"][-1]["t"].rstrip().endswith(":"):
                    continue
                if code_line(p["frags"][-1]["t"]) and code_line(f["t"]):
                    continue                                # строки перечня кодов — каждая своим куском
                # у каждой из двух строк слева на своей линии свой якорь (код, номер пункта) — это строки таблицы, а не
                # строки абзаца: колонка «GÉPHÁZ "A" sor 28KÁ» ×38 против колонки кодов шла одним абзацем и переводилась
                # повтором (3SZ23, ещё 3PR10, 3SZ17, 30HU, 30PR, 30RVR: аудит 29.10)
                a1, a2 = _row_anchor(frags, last), _row_anchor(frags, f["bbox"])
                if a1 is not None and a2 is not None and a1 != a2:
                    continue
                # фраза оборвана на середине: длинная строка без знака конца, следующая — со строчной (центрированный
                # текст в ячейке: «…A hidroakkumulátorok megengedett» / «szintváltozásának módosítása…» шли порознь, и
                # перевод выходил «Допустимые гидроемкостей САОЗ изменение изменения уровня», 30TH50, 28.10)
                if abs(x0 - last[0]) < 3 and mid_phrase(p["frags"][-1]["t"], f["t"]):
                    best = p
                    break
                # короткая строка (до преграды справа далеко) закрывает абзац — если первое слово следующей строки
                # в остаток влезло бы: в узкой ячейке строка коротка, потому что слово не поместилось
                # («Teszt / szüneteltetése», таблица 1 UTS — ячейка переводилась по строкам, 15.10)
                lim = _right_limit(frags, vs, last, page_right)
                width = lim - px0
                word = f["t"].split()[0] if f["t"].split() else ""
                fits = (x1 - x0) * (len(word) + 1) / max(len(f["t"]), 1) + 2
                # список «подпись — значение» без линеек (титул UTS: «RRP3 azonosító / EVP azonosító …» слипались):
                # у обеих строк справа на той же линии есть значение — строка закрывается и при малом остатке
                # (оценка ширины слова тут не годится: «EVP azonosító» — остаток 64,3 pt, «Dokumentum» — 64,4 pt);
                # в двух колонках сплошного текста строка доходит до промежутка, остаток меньше 25 pt
                pair = _partner(frags, vs, last) and _partner(frags, vs, f["bbox"])
                if pair and lim - last[2] > 25:
                    continue
                # столбик коротких значений против своих подписей слева («BSZG állásjelzések … rendben» ×4, 31RDGT:
                # «rendben rendben rendben rendben» шло одним абзацем); сплошной текст второй колонки — длинные строки
                if len(p["frags"][-1]["t"].split()) <= 3 and len(f["t"].split()) <= 3 \
                        and _partner(frags, vs, last, left=True) and _partner(frags, vs, f["bbox"], left=True):
                    continue
                # строка по центру ячейки: свободное место — с двух сторон («Fejlesztő / környezet» в шапке
                # таблицы 1 UTS переводились порознь: «Разработчик / среда»)
                room = lim - last[2]
                left = max([x for ry0, ry1, x in vs if x <= last[0] + 1 and min(last[3], ry1) - max(last[1], ry0) > 0],
                           default=None)
                if left is not None and room > 3 and abs((last[0] - left) - room) < 4:
                    room += last[0] - left
                if width > 0 and room > fits and lim - last[2] > 0.45 * width:
                    continue
                best = p
                break
        if best is None:
            pars.append({"frags": [f], "bbox": list(f["bbox"]), "size": f["size"], "bold": f["bold"]})
        else:
            best["frags"].append(f); best["bbox"] = _union([best["bbox"], f["bbox"]])
    for p in pars:
        p["text"] = join_lines([f["t"] for f in p["frags"]])
        fam = {}
        for f in p["frags"]:
            fam[f.get("family", "serif")] = fam.get(f.get("family", "serif"), 0) + len(f["t"])
        p["family"] = max(fam, key=fam.get) if fam else "serif"
    return pars


def gutter(pars, page_w):
    """Две колонки сплошного текста (Danfoss л. 9, E46 корпуса) → x межколоночного просвета, иначе None. Просвет —
    сквозной вертикальный, в средней половине листа, не уже 5 pt по рамкам кусков строк (они с хвостовыми пробелами:
    по словам просвет шире); его задевают только абзацы во всю ширину (заголовок, колонтитул), не больше 10 % кусков.
    Таблица без линий (абзацы слева и справа начинаются на одной высоте) — не колонки (22.10)."""
    frs = [f["bbox"] for p in pars for f in p["frags"]]
    if len(pars) < 6 or not frs:
        return None
    lo, hi = int(0.25 * page_w), int(0.75 * page_w)
    cover = [sum(1 for b in frs if b[0] - 1 <= x <= b[2] + 1) for x in range(lo, hi)]
    least = min(cover)
    if least > max(2, 0.1 * len(frs)):            # заголовок-другой во всю ширину — и на коротком листе
        return None
    runs, start = [], None                          # отрезки x с наименьшим покрытием
    for i, c in enumerate(cover + [least + 1]):
        if c == least and start is None:
            start = i
        elif c != least and start is not None:
            runs.append((start, i)); start = None
    a, z = max(runs, key=lambda r: r[1] - r[0])
    if z - a < 5:
        return None
    g = lo + (a + z) / 2
    left = [p for p in pars if p["bbox"][2] < g]
    right = [p for p in pars if p["bbox"][0] > g]
    if len(left) < 3 or len(right) < 3:
        return None
    # колонка текста — не уже четверти листа: узкий столбец слева (сокращения ⇥ расшифровки, u_T3) — таблица
    if g - min(p["bbox"][0] for p in left) < 0.25 * page_w or max(p["bbox"][2] for p in right) - g < 0.25 * page_w:
        return None
    rows = sum(1 for p in left if any(abs(p["bbox"][1] - q["bbox"][1]) < 2 for q in right))
    return None if rows >= 0.5 * min(len(left), len(right)) else g


def column_order(pars, g):
    """Абзацы двух колонок (просвет g, gutter) — колонка за колонкой, а не поперёк листа по высоте; абзацы во всю
    ширину делят лист на полосы, колонки читаются в каждой полосе («2.1 …» левой колонки и «2.3 …» правой шли
    подряд, текст колонок вперемешку, E46)."""
    side = lambda p: 0 if p["bbox"][2] < g else 1 if p["bbox"][0] > g else 2
    out, zone = [], []
    for p in sorted(pars, key=lambda p: (p["bbox"][1], p["bbox"][0])):
        if side(p) == 2:
            out += sorted(zone, key=lambda p: (side(p), p["bbox"][1])) + [p]
            zone = []
        else:
            zone.append(p)
    return out + sorted(zone, key=lambda p: (side(p), p["bbox"][1]))


def join_lines(lines):
    out = ""
    for ln in lines:
        ln = ln.strip()
        if not ln:
            continue
        if not out:
            out = ln
        elif out.endswith("-") and re.match(r"(?:és|vagy|illetve|ill\.|и|или|либо|and|or)\s", ln, re.I):
            out += " " + ln                                 # «fűtő- és hűtőrendszer»: дефис перед союзом — не перенос
        elif out.endswith("-") and ln[:1].islower():
            # дефис — часть слова: после прописной или цифры, и в составном через дефис («loss-of-» + «coolant», Eskom)
            keep = re.search(r"[A-ZÁÉÍÓÖŐÚÜŰ0-9]-$", out) or re.search(r"\w-\w+-$", out)
            out = (out if keep else out[:-1]) + ln
        else:
            out += " " + ln
    return out


def avail_rects(pars, hs, vs, page_rect):
    """Рамка под перевод: вправо до соседа/вертикальной линейки/правого поля, вниз до соседа/линейки."""
    if not pars:
        return
    wide = sorted(p["bbox"][2] for p in pars if p["bbox"][2] - p["bbox"][0] > 0.5 * page_rect.width)
    right_margin = wide[len(wide) // 2] if wide else max(p["bbox"][2] for p in pars)
    for p in pars:
        x0, y0, x1, y1 = p["bbox"]
        lim_x = max(right_margin, x1)
        for q in pars:
            if q is p:
                continue
            qx0, qy0, qx1, qy1 = q["bbox"]
            if qx0 >= x1 - 1 and min(y1, qy1) - max(y0, qy0) > 0:
                lim_x = min(lim_x, qx0 - 3)
        for ry0, ry1, x in vs:
            if x >= x1 - 0.5 and min(y1, ry1) - max(y0, ry0) > 0:
                lim_x = min(lim_x, x - 2)
        lim_y = page_rect.y1 - 20
        last = p["frags"][-1]["bbox"]
        for q in pars:
            if q is p:
                continue
            qx0, qy0, qx1, qy1 = q["bbox"]
            # сосед снизу — начинается ниже середины своей последней строки: рамки строк в тесном списке
            # перекрываются, и следующая строка не считалась соседом — перенос наезжал на неё (32RDGT, п. 2.2, 07.10)
            if qy0 >= min(y1 - 1, (last[1] + last[3]) / 2) and min(lim_x, qx1) - max(x0, qx0) > 0:
                lim_y = min(lim_y, qy0 - 1)
        for rx0, rx1, y in hs:
            if y >= y1 - 0.5 and min(lim_x, rx1) - max(x0, rx0) > 5:
                lim_y = min(lim_y, y - 1)
        p["avail"] = [x0, y0, max(x1, lim_x), max(y1, lim_y)]


def aligns(pars, vs):
    """Выключка абзаца: по центру — если он посередине своей ячейки (между ближайшими вертикальными линейками,
    без них — между краями текста страницы) и не делит левый край с соседом той же ячейки (строки списка).
    Края текста всей страницы сбивал штамп у кромки листа — заголовок титула не центрировался (32RDGT, 07.10)."""
    if not pars:
        return
    fl, fr = min(p["bbox"][0] for p in pars), max(p["bbox"][2] for p in pars)
    for p in pars:
        x0, y0, x1, y1 = p["bbox"]
        over = lambda ry0, ry1: min(y1, ry1) - max(y0, ry0) > 0.5 * (y1 - y0)
        lefts = [x for ry0, ry1, x in vs if x <= x0 + 1 and over(ry0, ry1)]
        rights = [x for ry0, ry1, x in vs if x >= x1 - 1 and over(ry0, ry1)]
        p["cell"] = [round(max(lefts), 2) if lefts else fl, round(min(rights), 2) if rights else fr]
    for p in pars:
        x0, y0, x1, y1 = p["bbox"]
        cl, cr = p["cell"]
        lm, rm = x0 - cl, cr - x1
        center = lm > 15 and rm > 15 and abs(lm - rm) < max(8, 0.06 * (cr - cl))
        if center and any(q is not p and q["cell"] == p["cell"] and abs(q["bbox"][0] - x0) < 1.5
                          and abs((q["bbox"][0] + q["bbox"][2]) / 2 - (x0 + x1) / 2) > 8 for q in pars):
            center = False
        p["align"] = "center" if center else "left"


def rules_from_image(gray, dpi, min_h_pt=15, min_v_pt=20):
    """Сплошные линейки таблицы на скане (у скана нет векторных линий): морфологическое «открытие» длинным ядром
    оставляет только длинные прямые штрихи — буквы, пунктир под подпись и флажки отсеиваются.
    gray — numpy-массив оттенков серого; возвращает (hs, vs) в пунктах, как rules()."""
    import cv2
    k = dpi / 72
    bw = cv2.adaptiveThreshold(gray, 255, cv2.ADAPTIVE_THRESH_MEAN_C, cv2.THRESH_BINARY_INV, 31, 15)
    out = {}
    for name, ksize, minlen in (("h", (int(min_h_pt * k), 1), min_h_pt), ("v", (1, int(min_v_pt * k)), min_v_pt)):
        lines = cv2.morphologyEx(bw, cv2.MORPH_OPEN, cv2.getStructuringElement(cv2.MORPH_RECT, ksize))
        _, _, stats, _ = cv2.connectedComponentsWithStats(lines, 8)
        res = []
        for x, y, w, h, _area in stats[1:]:
            if name == "h" and w / k >= minlen and h / k <= 4:
                res.append((x / k, (x + w) / k, (y + h / 2) / k))
            elif name == "v" and h / k >= minlen and w / k <= 4:
                res.append((y / k, (y + h) / k, (x + w / 2) / k))
        out[name] = res
    return out["h"], out["v"]
