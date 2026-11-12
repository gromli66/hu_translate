# -*- coding: utf-8 -*-
"""Копия скана чистым набором (режим задания «только копия оригинала», 12.11).

Скан без текстового слоя прежде шёл так: Tesseract режет лист на строки, портал лишь уточняет слова, перевод ставится
поверх картинки. Раскладка рвётся (обрывки, наложения, мусор из логотипов), Word выходит картинкой с надписями поверх.
Здесь основа — чтение листа порталом целиком (текст портала на сканах почти без ошибок): портал переписывает лист сразу
в HTML — заголовки, абзацы, полужирное, таблицы, места печатей, подписей и логотипов с рамкой. Из HTML собирается
обычный Word на языке оригинала: текст абзацами, таблицы — таблицами Word, печати, подписи, логотипы — вырезками скана.
Копию смотрит человек (и правит в Word), потом она переводится обычным заданием по пути DOCX.
Прежний путь перевода сканов не меняется."""
import difflib
import io
import json
import re
import time
from concurrent.futures import ThreadPoolExecutor
from html.parser import HTMLParser
from pathlib import Path

DPI_READ = 150          # картинка листа для портала: A4 — 1240×1754
DPI_CROP = 200          # вырезки печатей и логотипов — чуть чётче
FIGURES = ("печать", "штамп", "подпись", "логотип", "герб", "QR")


def prompt(lang, w, h):
    return f"""Это скан одной страницы документа, язык — {lang} (картинка {w}×{h} пикселей). Перепиши страницу в HTML, чтобы набрать её заново в Word.
Правила:
1. Текст — дословно, на языке оригинала, как напечатано: ничего не переводи, не исправляй, не пропускай и не добавляй. Номера, коды, даты, адреса, e-mail — посимвольно.
2. Строки одного абзаца склей в один <p>; переносы слов убери. Новый абзац — только где он в оригинале.
3. Заголовки — <h1> или <h2>; полужирное — <b>; подчёркнутое — <u>; текст по центру — <p align="center">, по правому краю — <p align="right">.
4. Таблица с линейками или ровными столбцами — <table border="1"> с <tr>/<td>; объединённые ячейки — colspan/rowspan.
5. Печать, штамп, подпись от руки, логотип, герб, QR-код не переписывай: на их место в потоке поставь <img alt="печать|штамп|подпись|логотип|герб|QR" data-box="x0,y0,x1,y1">, где data-box — рамка на картинке в пикселях.
6. Рукописный текст (вписанные от руки даты, фамилии) — перепиши в <i>.
Выведи только HTML от <body> до </body>, без пояснений."""


def read_page(pdf, pno, lang):
    """Лист → (HTML портала, ширина и высота картинки для портала, ошибка или None)."""
    import fitz
    from llm import chat, img_b64
    with fitz.open(str(pdf)) as doc:
        pix = doc[pno].get_pixmap(dpi=DPI_READ)
    r = chat([{"role": "user", "content": prompt(lang, pix.width, pix.height)}],
             images=[img_b64(pix.tobytes("png"))], max_tokens=12000, tag=f"copy{pno}")
    return r.get("text") or "", pix.width, pix.height, r.get("error")


def read_text(pdf, pno):
    """Второе чтение листа — простым текстом, тем же запросом, что OCR-путь перевода (segments.pdf_segments: серый
    лист 200 dpi, ocr_prompt): листы, уже читанные переводом, берутся из кэша. Нужно сверке (copycheck): чтение HTML
    путало Н/И, Т/Г, а простое — нет («Авангард-НСК», «СИТИ-СЕРТ», Декларация 12.11). → (текст, ошибка или None)."""
    import fitz
    from llm import chat, img_b64
    import segments as S
    with fitz.open(str(pdf)) as doc:
        png = doc[pno].get_pixmap(dpi=200, colorspace=fitz.csGRAY).tobytes("png")
    r = chat([{"role": "user", "content": S.ocr_prompt()}], images=[img_b64(png)], max_tokens=8000, tag=f"ocr{pno}")
    return r.get("text") or "", r.get("error")


BOX_PROMPT = """На картинке — страница документа ({w}×{h} пикселей). Найди на ней: {names}.
Для каждого — одна строка: название: x0, y0, x1, y1 — рамка в пикселях этой картинки (x вправо, y вниз от левого
верхнего угла); несколько одинаковых — каждый своей строкой, сверху вниз и слева направо. Чего нет — название: нет.
Больше ничего не пиши."""


def read_boxes(pdf, pno, alts):
    """Рамки картинок листа — отдельным коротким вопросом порталу. В задании «перепиши лист в HTML» портал место
    логотипа в тексте ставит, а рамку не даёт; поиск по скану (figures) логотип из букв («АВАНГАРД-НСК», «PLANET») от
    текста не отличает (Информационное письмо 12.11). → ([(название, доли листа x0, y0, x1, y1)], ошибка, ответ)."""
    import fitz
    from llm import chat, img_b64
    with fitz.open(str(pdf)) as doc:
        pix = doc[pno].get_pixmap(dpi=DPI_READ)
    r = chat([{"role": "user", "content": BOX_PROMPT.format(w=pix.width, h=pix.height, names=", ".join(alts))}],
             images=[img_b64(pix.tobytes("png"))], max_tokens=1000, tag=f"boxes{pno}")
    out = []
    num = r"(\d+(?:\.\d+)?)"
    for ln in (r.get("text") or "").splitlines():
        m = re.match(rf"\s*[-*•]?\s*(.+?)\s*:\s*{num}\D+{num}\D+{num}\D+{num}", ln)
        if not m:
            continue
        x0, y0, x1, y1 = (float(v) for v in m.groups()[1:])
        fr = (min(x0, x1) / pix.width, min(y0, y1) / pix.height, max(x0, x1) / pix.width, max(y0, y1) / pix.height)
        if 0 <= fr[0] < fr[2] <= 1.02 and 0 <= fr[1] < fr[3] <= 1.02:
            out.append((m.group(1).strip().strip("*").lower(), tuple(min(v, 1.0) for v in fr)))
    return out, r.get("error"), r.get("text") or ""


def placeholders(html):
    """Подписи заглушек картинок (<img alt>) HTML листа по порядку."""
    out = []

    def walk(n):
        for k in n.kids:
            if isinstance(k, Node):
                if k.tag == "img":
                    out.append((k.attrs.get("alt") or "рисунок").strip())
                walk(k)
    walk(parse(html))
    return out


def has_ink(page, frac):
    """В рамке (доли листа) есть краска — рамка портала не пустое место."""
    import fitz
    import numpy as np
    r = page.rect
    clip = fitz.Rect(r.x0 + frac[0] * r.width, r.y0 + frac[1] * r.height,
                     r.x0 + frac[2] * r.width, r.y0 + frac[3] * r.height) & r
    if clip.is_empty:
        return False
    pix = page.get_pixmap(dpi=50, clip=clip)
    a = np.frombuffer(pix.samples, np.uint8).reshape(pix.height, pix.width, pix.n)[:, :, :3]
    return int((a.min(2) < 150).sum()) >= 30


# ------------------------------------------------------------------ HTML → дерево
class Node:
    def __init__(self, tag, attrs, parent):
        self.tag, self.attrs, self.kids, self.parent = tag, dict(attrs), [], parent


class _Tree(HTMLParser):
    VOID = {"br", "img", "hr", "meta", "col", "input"}

    def __init__(self):
        super().__init__(convert_charrefs=True)
        self.root = Node("root", [], None)
        self.cur = self.root

    def handle_starttag(self, tag, attrs):
        n = Node(tag, attrs, self.cur)
        self.cur.kids.append(n)
        if tag not in self.VOID:
            self.cur = n

    def handle_startendtag(self, tag, attrs):
        self.cur.kids.append(Node(tag, attrs, self.cur))

    def handle_endtag(self, tag):
        n = self.cur
        while n is not self.root and n.tag != tag:
            n = n.parent
        if n is not self.root:                  # закрывающий без открывающего — мимо
            self.cur = n.parent

    def handle_data(self, data):
        self.cur.kids.append(data)


def parse(html):
    """HTML портала → корень дерева (body, если есть). Обёртка ```html и текст до первого тега — прочь."""
    html = re.sub(r"^.*?(?=<)", "", html.strip(), flags=re.S)
    html = re.sub(r"```\s*$", "", html)
    t = _Tree()
    t.feed(html)
    t.close()

    def find(n):
        if n.tag == "body":
            return n
        for k in n.kids:
            if isinstance(k, Node) and (b := find(k)):
                return b
        return None
    return find(t.root) or t.root


def box_of(node, w, h):
    """data-box картинки → (x0, y0, x1, y1) в долях листа или None."""
    nums = [float(v) for v in re.findall(r"-?\d+(?:\.\d+)?", node.attrs.get("data-box") or "")]
    if len(nums) != 4:
        return None
    x0, y0, x1, y1 = nums
    x0, x1 = sorted((x0, x1)); y0, y1 = sorted((y0, y1))
    r = (max(x0 / w, 0), max(y0 / h, 0), min(x1 / w, 1), min(y1 / h, 1))
    return r if r[2] - r[0] > 0.005 and r[3] - r[1] > 0.005 else None


def crop_rect(page, rect, ink=None):
    """Вырезка по рамке в pt (figures) — та же, что по рамке портала."""
    r = page.rect
    return crop(page, ((rect.x0 - r.x0) / r.width, (rect.y0 - r.y0) / r.height,
                       (rect.x1 - r.x0) / r.width, (rect.y1 - r.y0) / r.height), ink)


def crop(page, frac, ink=None):
    """Вырезка скана по рамке (доли листа) с запасом 5 %, поджатая по краске → (PNG, ширина и высота в pt).
    ink — круги печатей в pt (может быть пустым): вырезка печати, штампа, подписи — только цветная краска и то, что
    внутри круга печати, остальное белое: иначе в копию шёл и текст оригинала под печатью и рядом с ней («(подпись)»,
    «Регистрационный номер…», Декларация 12.11). None — как есть (логотип, фото)."""
    import fitz
    import numpy as np
    r = page.rect
    x0, y0, x1, y1 = frac
    dx, dy = 0.05 * (x1 - x0), 0.05 * (y1 - y0)
    if ink is not None:                     # бледный край штампа («К» у «КОПИЯ ВЕРНА») — лишнее вычистит маска цвета
        dx, dy = max(dx, 12 / r.width), max(dy, 12 / r.height)
    clip = fitz.Rect(r.x0 + (x0 - dx) * r.width, r.y0 + (y0 - dy) * r.height,
                     r.x0 + (x1 + dx) * r.width, r.y0 + (y1 + dy) * r.height) & r
    pix = page.get_pixmap(dpi=DPI_CROP, clip=clip)
    a = np.frombuffer(pix.samples, np.uint8).reshape(pix.height, pix.width, pix.n)[:, :, :3].copy()
    k = DPI_CROP / 72
    if ink is not None:
        # тот же порог, что у поиска печатей (figures): с 40 бледная синяя печать «Авангард LLC» находилась, но из
        # вырезки вычищалась (Информационное письмо 12.11)
        colored = (a.max(2).astype(np.int16) - a.min(2) >= 25) & (a.min(2) <= 200)
        keep = colored.copy()
        yy, xx = np.ogrid[:pix.height, :pix.width]
        for cx, cy, rad in ink:
            disc = (xx - (cx - clip.x0) * k) ** 2 + (yy - (cy - clip.y0) * k) ** 2 <= ((rad + 2) * k) ** 2
            # круг целиком — только у серой печати (ксерокопия): у цветной внутри круга чёрный текст оригинала
            # (номер декларации под печатью) повторял бы набранный
            if colored[disc].mean() < 0.03:
                keep |= disc
        a[~keep] = 255
    m = (a.min(2) < 200) | (a.max(2).astype(np.int16) - a.min(2) > 60)          # тёмное или цветное
    ys, xs = np.nonzero(m)
    if len(xs) > 20:
        pad = 6
        cx0, cx1 = max(xs.min() - pad, 0), min(xs.max() + pad + 1, pix.width)
        cy0, cy1 = max(ys.min() - pad, 0), min(ys.max() + pad + 1, pix.height)
        a = a[cy0:cy1, cx0:cx1]
        clip = fitz.Rect(clip.x0 + cx0 / k, clip.y0 + cy0 / k, clip.x0 + cx1 / k, clip.y0 + cy1 / k)
    if ink is not None:
        # бумага — прозрачная: печать ложится поверх текста краской, а не белым кругом, закрывающим строки
        alpha = np.where(a.min(2) > 225, 0, 255).astype(np.uint8)
        rgba = np.ascontiguousarray(np.dstack([a, alpha]))
        return (fitz.Pixmap(fitz.csRGB, rgba.shape[1], rgba.shape[0], rgba.tobytes(), True).tobytes("png"),
                clip.width, clip.height)
    a = np.ascontiguousarray(a)
    return fitz.Pixmap(fitz.csRGB, a.shape[1], a.shape[0], a.tobytes(), False).tobytes("png"), clip.width, clip.height


def float_shape(shape, x, y, rel=("column", "paragraph")):
    """Картинка в строке → поверх текста без обтекания (wp:anchor): x, y в pt от rel (по умолчанию — от левого края
    колонки и верха абзаца; ("page", "page") — от края листа). Печать и подпись на скане стоят поверх текста и места не
    занимают; в строке они раздвигали лист (Декларация 1 → 2)."""
    from docx.oxml import parse_xml
    from docx.oxml.ns import nsdecls
    inl = shape._inline
    emu = lambda v: str(int(round(v * 12700)))
    anc = parse_xml(
        f'<wp:anchor {nsdecls("wp")} distT="0" distB="0" distL="0" distR="0" simplePos="0" relativeHeight="251659264" '
        f'behindDoc="0" locked="0" layoutInCell="1" allowOverlap="1"><wp:simplePos x="0" y="0"/>'
        f'<wp:positionH relativeFrom="{rel[0]}"><wp:posOffset>{emu(x)}</wp:posOffset></wp:positionH>'
        f'<wp:positionV relativeFrom="{rel[1]}"><wp:posOffset>{emu(y)}</wp:posOffset></wp:positionV></wp:anchor>')
    anc.append(inl.extent)
    anc.append(parse_xml(f'<wp:effectExtent {nsdecls("wp")} l="0" t="0" r="0" b="0"/>'))
    anc.append(parse_xml(f'<wp:wrapNone {nsdecls("wp")}/>'))
    for el in list(inl):
        anc.append(el)                                      # docPr, cNvGraphicFramePr, graphic
    inl.getparent().replace(inl, anc)


# ------------------------------------------------------------------ печати, подписи, логотипы на скане
# Портал место картинки в тексте ставит, а рамку не даёт (вместо data-box — src="stamp.png", Декларация 12.11): где
# вырезать — находится по самому скану. Вырезки идут в копию картинками как есть и не переводятся (NO_TRANSLATE).
NO_TRANSLATE = "не переводить"
DARK_ALT = {"логотип", "герб", "qr", "фото", "рисунок", "изображение"}       # остальные (печать, штамп, подпись) — краска


def _merge(rects, tol=8):
    """Пересекающиеся (с запасом tol pt) рамки — одной: печать поверх подписи, «КОПИЯ ВЕРНА» рядом с печатью."""
    import fitz
    rects = [fitz.Rect(r) for r in rects]
    i = 0
    while i < len(rects):
        a = fitz.Rect(rects[i].x0 - tol, rects[i].y0 - tol, rects[i].x1 + tol, rects[i].y1 + tol)
        j = next((j for j in range(i + 1, len(rects)) if a.intersects(rects[j])), None)
        if j is None:
            i += 1
        else:
            rects[i] = rects[i] | rects.pop(j)
    return rects


def figures(page, dpi=100):
    """Печати, штампы, подписи, логотипы, гербы, QR, фото на скане → [(вид, рамка в pt, круги печатей в ней)] по
    высоте листа.
    - «цвет»: явная цветная краска (синяя, фиолетовая, красная: max−min ≥ 60) — не оттенок бумаги бланка и не шум у
      чёрных букв; куски ближе 6 pt — одна вырезка; больше 260 pt — фон бланка, не печать;
    - «круг»: круглая печать любой краски (серая на ксерокопии) — краска лежит по ободу и её нет чуть внутри и снаружи
      (иначе окружность «находится» в таблице с текстом);
    - «фото»: сплошь не белое от 60×60 pt (фото изделия, заливка);
    - «тёмное»: тёмное плотное пятно выше букв вчетверо (логотип, герб, QR), от 20×20 pt."""
    import cv2
    import fitz
    import numpy as np
    pix = page.get_pixmap(dpi=dpi)
    a = np.frombuffer(pix.samples, np.uint8).reshape(pix.height, pix.width, pix.n)[:, :, :3]
    k, W, H = dpi / 72, pix.width, pix.height
    gray = cv2.cvtColor(np.ascontiguousarray(a), cv2.COLOR_RGB2GRAY)
    mx, mn = a.max(2).astype(np.int16), a.min(2).astype(np.int16)
    # краска (не бумага: min ≤ 200) хоть немного цветная: у бледной печати на ксерокопии цвет едва виден (у половины
    # её краски max−min 19–25), у чёрного текста — 0 (Декларация о соответствии, 12.11; прежде порог 60 её терял)
    color = cv2.morphologyEx(((mx - mn >= 25) & (mn <= 200)).astype(np.uint8), cv2.MORPH_OPEN,
                             np.ones((2, 2), np.uint8))
    found = []
    n, _, st, _ = cv2.connectedComponentsWithStats(cv2.dilate(color, np.ones((int(6 * k), int(6 * k)), np.uint8)), 8)
    for i in range(1, n):
        x, y, w, h, _ = st[i]
        if 15 * k < w < 260 * k and 10 * k < h < 260 * k and color[y:y + h, x:x + w].sum() > 40 * k * k:
            found.append(("цвет", fitz.Rect(x / k, y / k, (x + w) / k, (y + h) / k)))

    def ring(cx, cy, r):
        t = np.linspace(0, 2 * np.pi, 180, endpoint=False)
        xs = np.clip((cx + r * np.cos(t)).astype(int), 0, W - 1)
        ys = np.clip((cy + r * np.sin(t)).astype(int), 0, H - 1)
        return float(np.mean([gray[max(y - 1, 0):y + 2, max(x - 1, 0):x + 2].min() < 160
                              or color[max(y - 1, 0):y + 2, max(x - 1, 0):x + 2].any() for x, y in zip(xs, ys)]))
    cs = cv2.HoughCircles(cv2.medianBlur(gray, 5), cv2.HOUGH_GRADIENT, dp=1.2, minDist=int(60 * k), param1=100,
                          param2=45, minRadius=int(35 * k), maxRadius=int(80 * k))
    for cx, cy, r in (cs[0] if cs is not None else []):
        on, f = max((ring(cx, cy, r * f), f) for f in np.linspace(0.92, 1.08, 9))
        if on >= 0.7 and on - (ring(cx, cy, r * f * 0.82) + ring(cx, cy, r * f * 1.18)) / 2 >= 0.25:
            R = r * f
            found.append(("круг", fitz.Rect((cx - R) / k, (cy - R) / k, (cx + R) / k, (cy + R) / k)))
    win = int(16 * k)
    dense = (cv2.blur((gray < 215).astype(np.float32), (win, win)) > 0.5).astype(np.uint8)
    n, _, st, _ = cv2.connectedComponentsWithStats(dense, 8)
    for i in range(1, n):
        x, y, w, h, area = st[i]
        if w >= 60 * k and h >= 60 * k and area >= 3600 * k * k:
            found.append(("фото", fitz.Rect(x / k, y / k, (x + w) / k, (y + h) / k)))
    dark = ((gray < 110) & (color == 0)).astype(np.uint8)
    n, lab, st, _ = cv2.connectedComponentsWithStats(dark, 8)
    hs = sorted(st[i][3] for i in range(1, n) if 3 <= st[i][3] <= 40 and st[i][2] <= 40)
    med = hs[len(hs) // 2] if hs else 10
    big = np.zeros_like(dark)
    for i in range(1, n):
        x, y, w, h, area = st[i]
        if h >= 4 * med and area / (w * h) >= 0.15 and w < 0.6 * W:
            big[lab == i] = 1
    n, _, st, _ = cv2.connectedComponentsWithStats(cv2.dilate(big, np.ones((int(12 * k), int(12 * k)), np.uint8)), 8)
    for i in range(1, n):
        x, y, w, h, _ = st[i]
        if h >= 20 * k and w >= 20 * k:
            found.append(("тёмное", fitz.Rect(x / k, y / k, (x + w) / k, (y + h) / k)))
    rings = lambda R: [((r.x0 + r.x1) / 2, (r.y0 + r.y1) / 2, r.width / 2) for kd, r in found
                       if kd == "круг" and R.contains(fitz.Point((r.x0 + r.x1) / 2, (r.y0 + r.y1) / 2))]

    def qr(R):
        """QR-код: квадрат из множества мелких тёмных пятен (у логотипа ЕАС их три)."""
        x0, y0, x1, y1 = (int(v * k) for v in R)
        cnt = cv2.connectedComponents((gray[y0:y1, x0:x1] < 110).astype(np.uint8), connectivity=8)[0] - 1
        return 0.8 <= R.width / max(R.height, 1) <= 1.25 and cnt >= 30
    # печати сливаются с печатями, картинки — с картинками: фото изделия на обложке не съедает печати под ним
    ink = [("круг" if any(kd == "круг" and R.intersects(r) for kd, r in found) else "цвет", R)
           for R in _merge([r for kd, r in found if kd in ("цвет", "круг")])]
    pics = [("фото" if any(kd == "фото" and R.intersects(r) for kd, r in found) else "qr" if qr(R) else "тёмное", R)
            for R in _merge([r for kd, r in found if kd in ("фото", "тёмное")])]
    # тёмное внутри печати (середина серой печати «М.П.») — часть печати, не логотип
    pics = [(kd, R) for kd, R in pics if not any((R & I).get_area() > 0.5 * R.get_area() for _, I in ink)]
    return sorted(((kd, R, rings(R) if kd in ("цвет", "круг") else []) for kd, R in ink + pics),
                  key=lambda f: (f[1].y0 + f[1].y1) / 2)


def text_lines(page):
    """Слова скана по Tesseract: [(слово, рамка в pt, строка)] — выравнивание абзацев копии по листу (портал
    выравнивание указывает не всегда: заголовок «ЕВРАЗИЙСКИЙ ЭКОНОМИЧЕСКИЙ СОЮЗ» вышел слева, Декларация о соответствии
    12.11). Нет Tesseract — пусто (выравнивание — только от портала)."""
    import tempfile
    import fitz
    import segments as S
    with tempfile.TemporaryDirectory() as td:
        f = Path(td) / "p.png"
        page.get_pixmap(dpi=150, colorspace=fitz.csGRAY).save(str(f))
        try:
            ws = S._tess_words(f, 150)
        except FileNotFoundError:
            return []
    return [(w["t"], fitz.Rect(w["box"]), w["line"]) for w in ws if w["c"] >= 40]


def align_of(text, words):
    """Выравнивание короткого абзаца (до 12 слов) по скану: края — по первому и последнему его слову на скане (лист
    бывает повёрнут на пару градусов, и Tesseract рвёт заголовок на две строки, а к строке липнет мусор логотипа), поля —
    по словам листа. По центру — поля слева и справа равны; вправо — прижат к правому краю; иначе None."""
    norm = lambda t: re.sub(r"\s+", " ", t.lower()).strip()
    toks = norm(text).split()
    if not toks or len(toks) > 12 or not words:
        return None
    if len(words) < 20:
        return None
    # поля — по всем словам листа (5-й и 95-й процентили): строки Tesseract рвёт на куски, и края «длинных строк»
    # уходили внутрь текста (правый 474 pt вместо 555, Декларация о соответствии); QR и мусор — в хвостах
    x0s, x1s = sorted(r.x0 for _, r, _ in words), sorted(r.x1 for _, r, _ in words)
    left, right = x0s[len(x0s) // 20], x1s[len(x1s) * 19 // 20]
    W = right - left

    def best(tok):
        sc = lambda w: difflib.SequenceMatcher(None, norm(w[0]), tok).ratio()
        w = max(words, key=sc)
        return w if sc(w) >= 0.7 else None
    first, last = best(toks[0]), best(toks[-1])
    if W <= 0 or first is None or last is None:
        return None
    h = max(first[1].height, last[1].height)
    if abs(first[1].y0 - last[1].y0) > 2.5 * h or last[1].x1 <= first[1].x0:
        return None
    lg, rg = first[1].x0 - left, right - last[1].x1
    if lg > 0.08 * W and abs(lg - rg) < 0.08 * W:
        return "center"
    if rg < 0.04 * W and lg > 0.3 * W:
        return "right"
    return None


# ------------------------------------------------------------------ дерево → Word
BLOCK = {"p", "div", "h1", "h2", "h3", "h4", "h5", "h6", "table", "ul", "ol", "li", "hr", "section", "article",
         "header", "footer", "center", "blockquote"}
ALIGN = {"center": 1, "right": 2, "justify": 3}           # WD_ALIGN_PARAGRAPH: 1 центр, 2 право, 3 по ширине


def own_align(node):
    """Выравнивание узла: атрибут align или text-align в style; нет — None."""
    m = re.search(r"text-align\s*:\s*(\w+)", node.attrs.get("style") or "")
    a = (m.group(1) if m else node.attrs.get("align") or "").lower()
    return a if a in ALIGN or a == "left" else None


class Builder:
    def __init__(self, doc, page, w, h, figs=(), size=11, lines=(), boxes=()):
        self.doc, self.page, self.w, self.h, self.size, self.lines = doc, page, w, h, size, list(lines)
        self.boxes, self.bused = list(boxes), set()                  # рамки портала (read_boxes)
        self.paras = 0
        self.figs, self.used, self.slots = list(figs), set(), []      # slots: (абзац, рамка)
        self.inked = False                                             # печать листа уже вставлена
        self.last = None                                               # последний абзац листа

    def text_width_pt(self):
        s = self.doc.sections[-1]
        return (s.page_width - s.left_margin - s.right_margin) / 12700

    def picture(self, run, png, wpt, cap, alt):
        """Вырезка в строку: ширина как на скане, не шире колонки; замещающий текст — «не переводить» (shots)."""
        from docx.shared import Pt
        shape = run.add_picture(io.BytesIO(png), width=Pt(min(wpt, cap or self.text_width_pt())))
        shape._inline.docPr.set("descr", f"{alt} — {NO_TRANSLATE}")
        return shape

    def margin_pt(self):
        return self.doc.sections[-1].left_margin / 12700

    def place(self, run, fig, alt, cap):
        """Картинка скана в абзац прогона. Печать, штамп, подпись (краска) — поверх текста там же на листе, что на скане
        (от края листа: копия кеглем вписана лист в лист, и место печати по абзацу-заглушке уводило её на чужие строки,
        Декларация о соответствии 12.11). Логотип, фото, QR — в строке, у того края, что на скане. В ячейке таблицы — в
        строке: поверх текста ячейки она уехала бы из ячейки."""
        from docx.table import _Cell
        kd, r, rings = fig
        ink = kd in ("цвет", "круг")
        png, wpt, hpt = crop_rect(self.page, r, rings if ink else None)
        shape = self.picture(run, png, wpt, cap, alt)
        par = run._parent
        if ink and not isinstance(par._parent, _Cell):
            float_shape(shape, r.x0 - self.page.rect.x0, r.y0 - self.page.rect.y0, rel=("page", "page"))
            self.inked = True
        elif not par.text.strip() and not isinstance(par._parent, _Cell):
            # картинка своей строкой — у того края, что на скане (QR справа, логотип слева)
            mid, W = (r.x0 + r.x1) / 2 - self.page.rect.x0, self.page.rect.width
            par.alignment = 0 if mid < W / 3 else 2 if mid > 2 * W / 3 else 1
        self.slots.append((par, r))

    def image(self, run, node, width_cap=None):
        """Заглушка портала (<img alt="печать">) → вырезка скана: по рамке портала, если дал, иначе — первая по
        высоте листа ещё не вставленная картинка того же рода (логотип, герб, QR — тёмная; печать, штамп, подпись —
        краска); нечего вставить — пометка курсивом «[печать]»."""
        alt = (node.attrs.get("alt") or "рисунок").strip()
        frac = box_of(node, self.w, self.h) if self.page is not None else None
        if frac is not None:
            png, wpt, _ = crop(self.page, frac)
            self.picture(run, png, wpt, width_cap, alt)
            r = self.page.rect
            box = type(r)(r.x0 + frac[0] * r.width, r.y0 + frac[1] * r.height,
                          r.x0 + frac[2] * r.width, r.y0 + frac[3] * r.height)
            hit = [i for i, f in enumerate(self.figs) if f[1].intersects(box)]      # та же печать — не второй раз
            self.used.update(hit)
            self.slots.append((run._parent, box))
            return
        a = alt.lower()
        if a in DARK_ALT and self.page is not None and self.portal_box(run, a, width_cap, alt):
            return
        want = ("qr",) if a == "qr" else ("тёмное", "фото") if a in DARK_ALT else ("цвет", "круг")
        i = next((i for i, f in enumerate(self.figs) if i not in self.used and f[0] in want), None)
        if i is None or self.page is None:
            if not (self.inked and want == ("цвет", "круг")):    # печати листа уже стоят одной склеенной вырезкой
                run.text = f"[{alt}]"; run.italic = True
            return
        self.used.add(i)
        self.place(run, self.figs[i], alt, width_cap)

    def portal_box(self, run, a, cap, alt):
        """Логотип, герб, QR, фото — по рамке портала (read_boxes), если она есть и в ней краска: вырезка как есть, в
        строке, у того края, что на скане; найденное поиском внутри неё — та же картинка (ромб логотипа). → вставлено ли."""
        from docx.table import _Cell
        j = next((j for j, (n, fr) in enumerate(self.boxes) if j not in self.bused and n == a
                  and has_ink(self.page, fr)), None)
        if j is None:
            return False
        self.bused.add(j)
        fr = self.boxes[j][1]
        png, wpt, _ = crop(self.page, fr)
        self.picture(run, png, wpt, cap, alt)
        r = self.page.rect
        box = type(r)(r.x0 + fr[0] * r.width, r.y0 + fr[1] * r.height, r.x0 + fr[2] * r.width, r.y0 + fr[3] * r.height)
        self.used.update(i for i, f in enumerate(self.figs) if f[1].intersects(box) and f[0] not in ("цвет", "круг"))
        par = run._parent
        if not par.text.strip() and not isinstance(par._parent, _Cell):
            mid = (box.x0 + box.x1) / 2 - r.x0
            par.alignment = 0 if mid < r.width / 3 else 2 if mid > 2 * r.width / 3 else 1
        self.slots.append((par, box))
        return True

    def finish(self, container):
        """Картинки скана, которым портал места не дал: печать — поверх текста там же на листе (привязана к последнему
        абзацу листа); логотип, фото, QR — отдельным абзацем в конце листа."""
        for i, fig in enumerate(self.figs):
            if i in self.used or self.page is None:
                continue
            kd = fig[0]
            self.used.add(i)
            if kd in ("цвет", "круг") and self.last is not None:
                self.place(self.last.add_run(), fig, "печать", None)
            else:
                self.place(container.add_paragraph().add_run(), fig, {"тёмное": "логотип", "qr": "QR"}.get(kd, "фото"),
                           None)

    def inline(self, par, node, fmt, cap=None):
        """Строчное содержимое узла → прогоны абзаца (b/i/u/sup/sub, br, img)."""
        for k in node.kids:
            if isinstance(k, str):
                t = re.sub(r"\s+", " ", k)
                if not par.runs and not par.text:
                    t = t.lstrip()
                if t:
                    r = par.add_run(t)
                    r.bold, r.italic, r.underline = fmt.get("b"), fmt.get("i"), fmt.get("u")
                    r.font.superscript, r.font.subscript = fmt.get("sup"), fmt.get("sub")
                continue
            tag = k.tag
            if tag == "br":
                par.add_run().add_break()
            elif tag == "img":
                self.image(par.add_run(), k, cap)
            else:
                f = dict(fmt)
                if tag in ("b", "strong", "th"):
                    f["b"] = True
                elif tag in ("i", "em"):
                    f["i"] = True
                elif tag == "u":
                    f["u"] = True
                elif tag in ("sup", "sub"):
                    f[tag] = True
                self.inline(par, k, f, cap)

    def paragraph(self, container, node, fmt, align=None, size=None, cap=None, prefix=""):
        par = container.add_paragraph()
        a = (node.attrs.get("align") or "").lower() or align
        m = re.search(r"text-align\s*:\s*(\w+)", node.attrs.get("style") or "")
        a = m.group(1).lower() if m else a
        if a in ALIGN:
            par.alignment = ALIGN[a]
        if prefix:
            r = par.add_run(prefix); r.bold = fmt.get("b")
        self.inline(par, node, fmt, cap)
        if par.runs and par.runs[-1].text:
            par.runs[-1].text = par.runs[-1].text.rstrip()      # перевод строки в конце абзаца HTML — не пробел
        if a not in ALIGN and a != "left" and cap is None and self.lines and (got := align_of(par.text, self.lines)):
            par.alignment = ALIGN[got]                         # портал не сказал — по строке скана (не в ячейке)
        if size:
            for r in par.runs:
                r.font.size = size
        self.paras += 1
        self.last = par
        return par

    def blocks(self, container, node, fmt=None, cap=None, align=None):
        """Блочное содержимое узла → абзацы и таблицы контейнера (тело документа или ячейка). align — выравнивание
        обёртки: text-align у <div> наследуют вложенные абзацы, как в браузере (портал оборачивает заголовок в
        <div style="text-align: center">, Декларация о соответствии 12.11)."""
        from docx.shared import Pt
        fmt = fmt or {}
        loose = Node("p", [], node)                         # строчное между блоками — своим абзацем
        for k in list(node.kids) + [None]:
            is_block = isinstance(k, Node) and k.tag in BLOCK
            if k is None or is_block:
                if any(isinstance(x, Node) or x.strip() for x in loose.kids):
                    self.paragraph(container, loose, fmt, align=align, cap=cap)
                loose = Node("p", [], node)
            if k is None:
                break
            if not is_block:
                loose.kids.append(k)
                continue
            tag = k.tag
            if tag in ("h1", "h2", "h3", "h4", "h5", "h6"):
                self.paragraph(container, k, dict(fmt, b=True), size=Pt(self.size + {"h1": 3, "h2": 1}.get(tag, 0)),
                               align=align, cap=cap)
            elif tag == "table":
                self.table(container, k, cap)
            elif tag in ("ul", "ol"):
                for i, li in enumerate([x for x in k.kids if isinstance(x, Node) and x.tag == "li"], 1):
                    self.paragraph(container, li, fmt, cap=cap, prefix=f"{i}. " if tag == "ol" else "– ")
            elif tag == "hr":
                continue
            elif tag == "li":
                self.paragraph(container, k, fmt, cap=cap, prefix="– ")
            elif any(isinstance(x, Node) and x.tag in BLOCK for x in k.kids):
                self.blocks(container, k, fmt, cap, own_align(k) or align)       # div с абзацами внутри
            else:
                self.paragraph(container, k, fmt, align=align, cap=cap)

    def table(self, container, node, cap=None):
        """<table> → таблица Word: сетка по colspan/rowspan, объединённые ячейки — merge."""
        from docx.shared import Pt
        rows = []

        def collect(n):
            for k in n.kids:
                if isinstance(k, Node):
                    if k.tag == "tr":
                        rows.append([c for c in k.kids if isinstance(c, Node) and c.tag in ("td", "th")])
                    elif k.tag != "table":
                        collect(k)
        collect(node)
        rows = [r for r in rows if r]
        if not rows:
            return
        span = lambda c, a: max(int(re.sub(r"\D", "", c.attrs.get(a) or "1") or 1), 1)
        occ, placed = set(), []
        for ri, tr in enumerate(rows):
            ci = 0
            for c in tr:
                while (ri, ci) in occ:
                    ci += 1
                cs, rs = span(c, "colspan"), min(span(c, "rowspan"), len(rows) - ri)
                occ |= {(ri + i, ci + j) for i in range(rs) for j in range(cs)}
                placed.append((ri, ci, cs, rs, c))
                ci += cs
        ncols = max(c for _, c in occ) + 1
        t = container.add_table(rows=len(rows), cols=ncols)
        # рамки — только у таблицы с линейками: шапку бланка портал даёт таблицей с «border: none», и копия рисовала
        # рамки, которых на скане нет (Информационное письмо 12.11); ничего не сказано — с рамками, как прежде
        style = (node.attrs.get("style") or "").replace(" ", "").lower()
        border = node.attrs.get("border")
        if border is not None and border.strip() not in ("", "0"):
            t.style = "Table Grid"
        elif not re.search(r"border:(none|0)", style):
            t.style = "Table Grid"
        full = cap or self.text_width_pt()
        # ширины столбцов — из width="25%" ячеек без colspan (первая такая в столбце), прочие поровну
        pct = {}
        for ri, ci, cs, rs, c in placed:
            m = re.fullmatch(r"\s*(\d+(?:\.\d+)?)\s*%\s*", c.attrs.get("width") or "")
            if cs == 1 and m and ci not in pct:
                pct[ci] = float(m.group(1)) / 100
        rest = [j for j in range(ncols) if j not in pct]
        left = max(1 - sum(pct.values()), 0)
        widths = [full * (pct[j] if j in pct else (left / len(rest) if rest else 0)) for j in range(ncols)]
        if pct and all(w > 0 for w in widths):
            from docx.shared import Pt as _Pt
            t.autofit = False
            for row in t.rows:
                for j, cell in enumerate(row.cells):
                    cell.width = _Pt(widths[j])
        else:
            widths = [full / ncols] * ncols
        for ri, ci, cs, rs, c in placed:
            cell = t.cell(ri, ci)
            if cs > 1 or rs > 1:
                cell = cell.merge(t.cell(ri + rs - 1, ci + cs - 1))
            box = _Cell(cell)
            self.blocks(box, c, {"b": True} if c.tag == "th" else {}, cap=sum(widths[ci:ci + cs]), align=own_align(c))
            for p in cell.paragraphs:
                for r in p.runs:
                    r.font.size = Pt(max(self.size - 2, 7))
        if not hasattr(container, "_take"):
            container.add_paragraph()                       # таблицы подряд не склеиваются в одну


class _Cell:
    """Ячейка как контейнер: первый абзац — уже есть в ячейке, дальше — новые."""

    def __init__(self, cell):
        self.cell, self._first = cell, True

    def _take(self):
        self._first = False
        return self.cell.paragraphs[0]

    def add_paragraph(self):
        if self._first:
            return self._take()
        return self.cell.add_paragraph()

    def add_table(self, rows, cols):
        self._first = False
        return self.cell.add_table(rows=rows, cols=cols)


def new_document(size=11):
    from docx import Document
    from docx.shared import Pt
    doc = Document()
    st = doc.styles["Normal"]
    st.font.name, st.font.size = "Times New Roman", Pt(size)
    st.element.rPr.rFonts.set("{http://schemas.openxmlformats.org/wordprocessingml/2006/main}eastAsia", "Times New Roman")
    st.paragraph_format.space_after = Pt(1)
    st.paragraph_format.line_spacing = 1.0
    return doc


def margins(src):
    """Поля копии — по краске скана (самые узкие по всем листам), 20–72 pt: у бланков поля уже 2 см, и копия с полями
    шаблона уходила на лист больше оригинала."""
    import numpy as np
    out = [72.0] * 4                                          # слева, сверху, справа, снизу
    for page in src:
        pix = page.get_pixmap(dpi=50)
        a = np.frombuffer(pix.samples, np.uint8).reshape(pix.height, pix.width, pix.n)[:, :, :3].min(2)
        cols, rows = np.nonzero((a < 150).sum(0) >= 3)[0], np.nonzero((a < 150).sum(1) >= 3)[0]
        if not len(cols) or not len(rows):
            continue
        k = 72 / 50
        got = [cols[0] * k, rows[0] * k, (pix.width - 1 - cols[-1]) * k, (pix.height - 1 - rows[-1]) * k]
        out = [min(o, g) for o, g in zip(out, got)]
    return [min(max(v - 2, 20.0), 72.0) for v in out]


def page_setup(section, rect, m=(56.7, 42.5, 56.7, 42.5)):
    from docx.shared import Pt
    from docx.enum.section import WD_ORIENT
    section.page_width, section.page_height = Pt(rect.width), Pt(rect.height)
    section.orientation = WD_ORIENT.LANDSCAPE if rect.width > rect.height else WD_ORIENT.PORTRAIT
    section.left_margin, section.top_margin, section.right_margin, section.bottom_margin = (Pt(v) for v in m)


def build(pdf, pages, out, size=11, lines=None, boxes=None):
    """pages — [(HTML, ширина, высота, ошибка)] по листам → Word с кеглем size; lines — строки скана по листам
    (text_lines) для выравнивания; → число абзацев."""
    import fitz
    from docx.enum.section import WD_SECTION
    doc = new_document(size)
    n = 0
    with fitz.open(str(pdf)) as src:
        m = margins(src)
        for pno, (html, w, h, err) in enumerate(pages):
            page = src[pno]
            if pno == 0:
                page_setup(doc.sections[0], page.rect, m)
            else:
                prev = src[pno - 1].rect
                if (prev.width > prev.height) != (page.rect.width > page.rect.height):
                    page_setup(doc.add_section(WD_SECTION.NEW_PAGE), page.rect, m)
                else:
                    doc.add_page_break()
            if err or not html.strip():
                p = doc.add_paragraph(f"[лист {pno + 1} не прочитан: {err or 'пустой ответ портала'}]")
                p.runs[0].italic = True
                print(f"копия: {Path(pdf).name}, лист {pno + 1} — не прочитан: {err or 'пустой ответ'}"[:300], flush=True)
                continue
            if "</body>" not in html and "</html>" not in html:
                print(f"копия: {Path(pdf).name}, лист {pno + 1} — ответ портала оборван (нет </body>)", flush=True)
            b = Builder(doc, page, w, h, figures(page), size, lines[pno] if lines else (), boxes[pno] if boxes else ())
            b.blocks(doc, parse(html))
            b.finish(doc)
            n += b.paras
    out.parent.mkdir(parents=True, exist_ok=True)
    doc.save(str(out))
    return n


SIZES = (11, 10.5, 10, 9.5, 9, 8.5, 8)


def fit(pdf, pages, out, n_pages, tmp_dir, boxes=None):
    """Кегль копии — самый крупный из SIZES, при котором листов копии не больше, чем у оригинала (лист копии — лист
    оригинала: печати и разрывы листов — на своих листах). Листы считает LibreOffice; нет его — кегль 11.
    → (абзацев, кегль, листов копии или None)."""
    import fitz
    from assemble import lo_convert
    tmp_dir.mkdir(parents=True, exist_ok=True)
    got = None
    with fitz.open(str(pdf)) as d:
        lines = [text_lines(pg) for pg in d]
    for size in SIZES:
        k = build(pdf, pages, out, size, lines, boxes)
        try:
            lo_convert(out, tmp_dir / f"{out.stem}.pdf", timeout=300)
        except (OSError, RuntimeError) as e:
            print(f"копия: {out.name} — листы не сверены: {type(e).__name__}: {e}"[:300], flush=True)
            return k, size, None
        with fitz.open(str(tmp_dir / f"{out.stem}.pdf")) as d:
            got = d.page_count
        if got <= n_pages:
            break
    return k, size, got


def run(in_dir, work, lang, progress=None):
    """Каждый PDF папки → <имя>.docx в work/out: копия оригинала чистым набором. → число абзацев всех копий."""
    import fitz
    import langs as L
    import segments as S
    files = sorted(f for f in Path(in_dir).iterdir() if f.is_file())
    pdfs = [f for f in files if f.suffix.lower() == ".pdf"]
    for f in files:
        if f.suffix.lower() != ".pdf":
            print(f"копия: {f.name} — не PDF, пропущен (копия нужна только сканам)", flush=True)
    with_pages = []
    for f in pdfs:
        up = S.upright(f, Path(work) / "upright")       # повёрнутые листы — прямо
        if up != f:
            print(f"копия: {f.name} — листы с поворотом выпрямлены", flush=True)
        with fitz.open(str(up)) as d:
            with_pages.append((f, up, d.page_count))
    total, done, n = sum(c for *_, c in with_pages) or 1, 0, 0
    if progress:
        progress("копия", 0, total)
    for f, up, cnt in with_pages:
        t0 = time.time()
        S.SRC_LANG = lang                               # ocr_prompt — по языку источника
        with ThreadPoolExecutor(4) as ex:
            pages = list(ex.map(lambda p: read_page(up, p, L.NAME_RU.get(lang, lang)), range(cnt)))
            texts = list(ex.map(lambda p: read_text(up, p), range(cnt)))
            # рамки картинок — только у листов с заглушками логотипа, герба, QR, фото (печати поиск находит сам)
            alts = [placeholders(html) for html, *_ in pages]
            asked = list(ex.map(lambda p: read_boxes(up, p, alts[p]) if any(a.lower() in DARK_ALT for a in alts[p])
                                else ([], None, ""), range(cnt)))
        # оба чтения и путь к листам (выпрямленным) — для сверки копии (copycheck): Surya читает те же листы
        rd = Path(work) / "copy_read" / f.stem.rstrip(". ")      # «X..pdf»: папку с точкой на конце Windows не открывает
        rd.mkdir(parents=True, exist_ok=True)
        for pno, ((html, *_), (txt, terr)) in enumerate(zip(pages, texts)):
            (rd / f"p{pno:04d}.html").write_text(html, encoding="utf-8")
            (rd / f"p{pno:04d}.txt").write_text(txt, encoding="utf-8")
            if terr:
                print(f"копия: {f.name}, лист {pno + 1} — второе чтение не получено: {terr}"[:300], flush=True)
            bx, berr, btxt = asked[pno]
            if btxt or berr:
                (rd / f"p{pno:04d}.boxes.txt").write_text(btxt or f"ошибка: {berr}", encoding="utf-8")
        (rd / "pages.json").write_text(json.dumps({"pdf": str(up), "pages": cnt}, ensure_ascii=False), encoding="utf-8")
        done += cnt
        if progress:
            progress("копия", done, total)
        out = Path(work) / "out" / f"{f.stem}.docx"
        k, size, got = fit(up, pages, out, cnt, Path(work) / "copy_fit", [a[0] for a in asked])
        n += k
        print(f"копия: {f.name} — листов {cnt}, абзацев {k}, кегль {size}"
              + (f", листов копии {got}" if got is not None else ", без сверки листов (нет LibreOffice)")
              + f", {time.time() - t0:.0f} с", flush=True)
    return n
