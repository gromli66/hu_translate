# -*- coding: utf-8 -*-
"""Копия скана чистым набором (режим задания «только копия оригинала», 12.11).

Скан без текстового слоя прежде шёл так: Tesseract режет лист на строки, портал лишь уточняет слова, перевод ставится
поверх картинки. Раскладка рвётся (обрывки, наложения, мусор из логотипов), Word выходит картинкой с надписями поверх.
Здесь основа — чтение листа порталом целиком (текст портала на сканах почти без ошибок): портал переписывает лист сразу
в HTML — заголовки, абзацы, полужирное, таблицы, места печатей, подписей и логотипов с рамкой. Из HTML собирается
обычный Word на языке оригинала: текст абзацами, таблицы — таблицами Word, печати, подписи, логотипы — вырезками скана.
Копию смотрит человек (и правит в Word), потом она переводится обычным заданием по пути DOCX.
Прежний путь перевода сканов не меняется."""
import io
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


def crop(page, frac):
    """Вырезка скана по рамке портала (доли листа) с запасом 5 %, поджатая по краске → (PNG, ширина и высота в pt)."""
    import fitz
    import numpy as np
    r = page.rect
    x0, y0, x1, y1 = frac
    dx, dy = 0.05 * (x1 - x0), 0.05 * (y1 - y0)
    clip = fitz.Rect(r.x0 + (x0 - dx) * r.width, r.y0 + (y0 - dy) * r.height,
                     r.x0 + (x1 + dx) * r.width, r.y0 + (y1 + dy) * r.height) & r
    pix = page.get_pixmap(dpi=DPI_CROP, clip=clip)
    a = np.frombuffer(pix.samples, np.uint8).reshape(pix.height, pix.width, pix.n)[:, :, :3].astype(np.int16)
    ink = (a.min(2) < 200) | (a.max(2) - a.min(2) > 60)          # тёмное или цветное (синяя печать, подпись)
    ys, xs = np.nonzero(ink)
    if len(xs) > 20:
        pad = 6
        cx0, cx1 = max(xs.min() - pad, 0), min(xs.max() + pad, pix.width - 1)
        cy0, cy1 = max(ys.min() - pad, 0), min(ys.max() + pad, pix.height - 1)
        k = 72 / DPI_CROP
        clip = fitz.Rect(clip.x0 + cx0 * k, clip.y0 + cy0 * k, clip.x0 + cx1 * k, clip.y0 + cy1 * k)
        pix = page.get_pixmap(dpi=DPI_CROP, clip=clip)
    return pix.tobytes("png"), clip.width, clip.height


# ------------------------------------------------------------------ дерево → Word
BLOCK = {"p", "div", "h1", "h2", "h3", "h4", "h5", "h6", "table", "ul", "ol", "li", "hr", "section", "article",
         "header", "footer", "center", "blockquote"}
ALIGN = {"center": 1, "right": 2, "justify": 3}           # WD_ALIGN_PARAGRAPH: 1 центр, 2 право, 3 по ширине


class Builder:
    def __init__(self, doc, page, w, h):
        self.doc, self.page, self.w, self.h = doc, page, w, h
        self.paras = 0

    def text_width_pt(self):
        s = self.doc.sections[-1]
        return (s.page_width - s.left_margin - s.right_margin) / 12700

    def image(self, run, node, width_cap=None):
        """Картинка-заглушка портала → вырезка скана в строку; без рамки — пометка курсивом."""
        from docx.shared import Pt
        frac = box_of(node, self.w, self.h)
        alt = node.attrs.get("alt") or "рисунок"
        if frac is None or self.page is None:
            run.text = f"[{alt}]"; run.italic = True
            return
        png, wpt, _ = crop(self.page, frac)
        cap = width_cap or self.text_width_pt()
        run.add_picture(io.BytesIO(png), width=Pt(min(wpt, cap)))

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
        if size:
            for r in par.runs:
                r.font.size = size
        self.paras += 1
        return par

    def blocks(self, container, node, fmt=None, cap=None):
        """Блочное содержимое узла → абзацы и таблицы контейнера (тело документа или ячейка)."""
        from docx.shared import Pt
        fmt = fmt or {}
        loose = Node("p", [], node)                         # строчное между блоками — своим абзацем
        for k in list(node.kids) + [None]:
            is_block = isinstance(k, Node) and k.tag in BLOCK
            if k is None or is_block:
                if any(isinstance(x, Node) or x.strip() for x in loose.kids):
                    self.paragraph(container, loose, fmt, cap=cap)
                loose = Node("p", [], node)
            if k is None:
                break
            if not is_block:
                loose.kids.append(k)
                continue
            tag = k.tag
            if tag in ("h1", "h2", "h3", "h4", "h5", "h6"):
                self.paragraph(container, k, dict(fmt, b=True), size=Pt({"h1": 14, "h2": 12}.get(tag, 11)), cap=cap)
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
                self.blocks(container, k, fmt, cap)          # div с абзацами внутри
            else:
                self.paragraph(container, k, fmt, cap=cap)

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
        t.style = "Table Grid"
        cw = (cap or self.text_width_pt()) / ncols
        for ri, ci, cs, rs, c in placed:
            cell = t.cell(ri, ci)
            if cs > 1 or rs > 1:
                cell = cell.merge(t.cell(ri + rs - 1, ci + cs - 1))
            box = _Cell(cell)
            self.blocks(box, c, {"b": True} if c.tag == "th" else {}, cap=cw * cs)
            for p in cell.paragraphs:
                for r in p.runs:
                    r.font.size = Pt(9)
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


def new_document():
    from docx import Document
    from docx.shared import Pt
    doc = Document()
    st = doc.styles["Normal"]
    st.font.name, st.font.size = "Times New Roman", Pt(11)
    st.element.rPr.rFonts.set("{http://schemas.openxmlformats.org/wordprocessingml/2006/main}eastAsia", "Times New Roman")
    st.paragraph_format.space_after = Pt(3)
    return doc


def page_setup(section, rect):
    from docx.shared import Pt, Cm
    from docx.enum.section import WD_ORIENT
    section.page_width, section.page_height = Pt(rect.width), Pt(rect.height)
    section.orientation = WD_ORIENT.LANDSCAPE if rect.width > rect.height else WD_ORIENT.PORTRAIT
    section.left_margin = section.right_margin = Cm(2)
    section.top_margin = section.bottom_margin = Cm(1.5)


def build(pdf, pages, out):
    """pages — [(HTML, ширина, высота, ошибка)] по листам → Word; → число абзацев."""
    import fitz
    from docx.enum.section import WD_SECTION
    doc = new_document()
    n = 0
    with fitz.open(str(pdf)) as src:
        for pno, (html, w, h, err) in enumerate(pages):
            page = src[pno]
            if pno == 0:
                page_setup(doc.sections[0], page.rect)
            else:
                prev = src[pno - 1].rect
                if (prev.width > prev.height) != (page.rect.width > page.rect.height):
                    page_setup(doc.add_section(WD_SECTION.NEW_PAGE), page.rect)
                else:
                    doc.add_page_break()
            if err or not html.strip():
                p = doc.add_paragraph(f"[лист {pno + 1} не прочитан: {err or 'пустой ответ портала'}]")
                p.runs[0].italic = True
                print(f"копия: {Path(pdf).name}, лист {pno + 1} — не прочитан: {err or 'пустой ответ'}"[:300], flush=True)
                continue
            if "</body>" not in html and "</html>" not in html:
                print(f"копия: {Path(pdf).name}, лист {pno + 1} — ответ портала оборван (нет </body>)", flush=True)
            b = Builder(doc, page, w, h)
            b.blocks(doc, parse(html))
            n += b.paras
    out.parent.mkdir(parents=True, exist_ok=True)
    doc.save(str(out))
    return n


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
        with ThreadPoolExecutor(4) as ex:
            pages = list(ex.map(lambda p: read_page(up, p, L.NAME_RU.get(lang, lang)), range(cnt)))
        done += cnt
        if progress:
            progress("копия", done, total)
        k = build(up, pages, Path(work) / "out" / f"{f.stem}.docx")
        n += k
        print(f"копия: {f.name} — листов {cnt}, абзацев {k}, {time.time() - t0:.0f} с", flush=True)
    return n
