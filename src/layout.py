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
                r = fitz.Rect(g[0]["bbox"])
                for c in g[1:]:
                    r |= fitz.Rect(c["bbox"])
                out.append(dict(s, text="".join(c["c"] for c in g), bbox=tuple(r)))
    return out


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
                txt = s["text"]
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
                        and not _vrule_in(vs, cur["bbox"][2], bb[0], bb[1], bb[3]) \
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
                and 0 <= f["bbox"][0] - m["bbox"][2] <= GAP and not _vrule_in(vs, m["bbox"][2], f["bbox"][0], f["bbox"][1], f["bbox"][3]) \
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
    out = merged
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


def _partner(frags, vs, bb):
    """Справа на той же линии есть другой фрагмент и линейки между ними нет (вторая колонка списка без таблицы)."""
    x0, y0, x1, y1 = bb
    return any(f["bbox"][0] >= x1 + 5 and min(y1, f["bbox"][3]) - max(y0, f["bbox"][1]) > 0.5 * (y1 - y0)
               and not _vrule_in(vs, x1, f["bbox"][0], y0, y1) for f in frags)


def group(frags, hs, vs=(), form=False):
    """form=True (скан формы): строка, кончающаяся двоеточием, — подпись поля, к ней не приклеивается следующая."""
    page_right = max((f["bbox"][2] for f in frags), default=0)
    frags = sorted(frags, key=lambda f: (round(f["bbox"][1] / 3), f["bbox"][0]))
    # слева на той же строке — номер пункта отдельным куском («1.6.» ⇥ «FOT szintcsökkenés» в оглавлении 3SZ19): кусок
    # начинает новый абзац — иначе название приклеивалось к названию пункта выше, а у «1.6.» оставалось пустым (17.10)
    label = re.compile(r"^\(?\d+(\.\d+)*\.?\)?$|^[a-zA-Z][.)]$")
    nums = [f["bbox"] for f in frags if label.match(f["t"].strip())]
    numbered = lambda bb: any(n[2] < bb[0] and bb[0] - n[2] < 120 and min(n[3], bb[3]) - max(n[1], bb[1]) > 0.5 * (
        bb[3] - bb[1]) for n in nums)
    pars = []
    for f in frags:
        x0, y0, x1, y1 = f["bbox"]
        best = None
        # маркер пункта отдельным фрагментом — сам по себе: маркеры соседних пунктов слипались в один «• •», и пункт
        # оставался без своего (Eskom 331-165, 17.10)
        if not MARK.match(f["t"]) and not f.get("item") and not BULLET.match(f["t"].strip()) and not numbered(f["bbox"]):
            for p in reversed(pars[-12:]):
                if len(p["frags"]) == 1 and BULLET.match(p["frags"][0]["t"].strip()):
                    continue
                px0, py0, px1, py1 = p["bbox"]
                last = p["frags"][-1]["bbox"]
                gap = y0 - last[3]
                size = max(p["size"], f["size"])
                if not (-0.35 * size <= gap <= 0.6 * size):
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
