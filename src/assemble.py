# -*- coding: utf-8 -*-
"""Сборка перевода обратно в формат источника + двуязычная таблица для вычитки.

- DOCX: текст каждой группы прогонов (одинаковое начертание) пишется в первый прогон группы,
  остальные прогоны группы опустошаются; метки <fN> раскладывают перевод по группам.
  Если метки сломаны — весь перевод в первую группу (начертание первой группы).
- PDF: исходный текст переведённых абзацев убирается (текстовый слой — redaction текста;
  OCR-страница — белая заливка + удаление нарисованных букв), перевод вписывается в рамку
  абзаца, расширенную до соседей (insert_htmlbox ужимает кегль, если не влезает)."""
import re, json, html, zipfile
from pathlib import Path

from segments import W, _own_runs
from layout import rules

TAG = re.compile(r"</?f\d+>")


def _split_groups(ru, n):
    if n <= 1:
        return [TAG.sub("", ru)]
    parts = re.findall(r"<f(\d+)>(.*?)</f\1>", ru, re.S)
    nums = [int(k) for k, _ in parts]
    if sorted(nums) == list(range(1, n + 1)):
        d = {int(k): t for k, t in parts}
        return [d[k] for k in range(1, n + 1)]
    return [TAG.sub("", ru)] + [""] * (n - 1)


def _set_run_text(r, text):
    from lxml import etree
    for ch in list(r):
        if ch.tag in (W + "t", W + "tab", W + "br", W + "cr", W + "noBreakHyphen", W + "softHyphen"):
            r.remove(ch)
    if not text:
        return
    for k, line in enumerate(text.split("\n")):
        if k:
            etree.SubElement(r, W + "br")
        for j, chunk in enumerate(line.split("\t")):
            if j:
                etree.SubElement(r, W + "tab")
            if chunk:
                t = etree.SubElement(r, W + "t")
                t.text = chunk
                t.set("{http://www.w3.org/XML/1998/namespace}space", "preserve")


def build_docx(src, out, segs, res):
    from lxml import etree
    by_part = {}
    for s in segs:
        r = res.get(str(s["id"]))
        if r and not r.get("copied") and r["ru"]:
            by_part.setdefault(s["loc"]["part"], []).append((s, r["ru"]))
    broken = 0
    with zipfile.ZipFile(src) as zin, zipfile.ZipFile(out, "w", zipfile.ZIP_DEFLATED) as zout:
        for item in zin.infolist():
            data = zin.read(item.filename)
            if item.filename in by_part:
                root = etree.fromstring(data)
                ps = list(root.iter(W + "p"))
                for s, ru in by_part[item.filename]:
                    p = ps[s["loc"]["p"]]
                    runs = _own_runs(p)
                    groups = s["loc"]["groups"]
                    texts = _split_groups(ru, len(groups))
                    texts[0] = s["loc"].get("lead", "") + texts[0]
                    if len(groups) > 1 and not any(texts[1:]) and texts[0]:
                        broken += 1
                    for g, t in zip(groups, texts):
                        _set_run_text(runs[g[0]], t)
                        for ri in g[1:]:
                            _set_run_text(runs[ri], "")
                data = etree.tostring(root, xml_declaration=True, encoding="UTF-8", standalone=True)
            zout.writestr(item, data)
    return {"broken_tags": broken}


def build_pdf(src, out, segs, res):
    import fitz
    doc = fitz.open(src)
    pages = sorted({s["page"] for s in segs})
    shrink = []
    sub = fitz.open()
    for pno in pages:
        page = doc[pno]
        todo = [(s, res[str(s["id"])]) for s in segs if s["page"] == pno
                and not res[str(s["id"])].get("copied") and res[str(s["id"])]["ru"]]
        ocr = any(s["loc"].get("ocr") for s, _ in todo)
        for s, r in todo:
            if ocr:
                # OCR-страница: буквы нарисованы кривыми — белая заливка по общей рамке абзаца,
                # влево с запасом под тире/номер пункта, которые Tesseract мог не включить в рамку
                x0, y0, x1, y1 = s["loc"]["bbox"]
                left = 16 if re.match(r"\s*[-–•]", s["text"]) else 2
                page.add_redact_annot(fitz.Rect(x0 - left, y0 - 1.5, x1 + 2, y1 + 1.5), fill=(1, 1, 1))
            else:
                for bb in s["loc"]["frags"]:
                    page.add_redact_annot(fitz.Rect(bb), fill=False)
        page.apply_redactions(images=fitz.PDF_REDACT_IMAGE_NONE,
                              graphics=fitz.PDF_REDACT_LINE_ART_REMOVE_IF_COVERED if ocr else fitz.PDF_REDACT_LINE_ART_NONE,
                              text=fitz.PDF_REDACT_TEXT_REMOVE)
        boxes = [s["loc"]["bbox"] for s in segs if s["page"] == pno]
        fl, fr = min(b[0] for b in boxes), max(b[2] for b in boxes)
        items = []
        for s, r in todo:
            loc = s["loc"]
            x0, y0, x1, y1 = loc["bbox"]
            ax1, ay1 = loc["avail"][2], loc["avail"][3]
            rect = fitz.Rect(x0, y0 - 0.5, max(ax1, x1), max(ay1, y1) + 0.5)
            align = "left"
            # центр — только при равных полях слева и справа от рамки текста страницы
            if abs((x0 - fl) - (fr - x1)) < 12 and (x1 - x0) < 0.7 * (fr - fl) and x0 - fl > 20:
                half = min((x0 + x1) / 2 - fl, fr - (x0 + x1) / 2)
                rect = fitz.Rect((x0 + x1) / 2 - half, y0 - 0.5, (x0 + x1) / 2 + half, max(ay1, y1) + 0.5)
                align = "center"
            txt = html.escape(TAG.sub("", r["ru"])).replace("\n", "<br>").replace("\t", "&emsp;")
            if loc.get("bold"):
                txt = f"<b>{txt}</b>"
            items.append((s, rect, txt, align, loc["size"]))
        # общий масштаб страницы: пробная вставка на пустой странице, медиана нужных масштабов
        trial = fitz.open(); tp = trial.new_page(width=page.rect.width, height=page.rect.height)
        need = []
        for s, rect, txt, align, size in items:
            css = f"* {{font-family: serif; font-size: {size:.1f}pt; line-height: 1.15; text-align: {align};}}"
            need.append(tp.insert_htmlbox(rect, txt, css=css, scale_low=0)[1])
        k = sorted(need)[int(len(need) * 0.3)] if need else 1.0
        for (s, rect, txt, align, size), nk in zip(items, need):
            sz = size * min(k, 1.0)
            css = f"* {{font-family: serif; font-size: {sz:.1f}pt; line-height: 1.15; text-align: {align};}}"
            spare, scale = page.insert_htmlbox(rect, txt, css=css, scale_low=0)
            if scale * min(k, 1.0) < 0.7:
                shrink.append({"page": pno + 1, "id": s["id"], "scale": round(scale * min(k, 1.0), 2)})
        sub.insert_pdf(doc, from_page=pno, to_page=pno)
    # отдаём только обработанные страницы: и для выборки, и для полного прогона это вся работа
    # каждая вставка абзаца встраивает свой экземпляр шрифта (5275 объектов на 157 стр. ÜFK I, 43,8 МБ);
    # garbage=4 склеивает одинаковые объекты: 3,3 МБ, рендер совпадает пиксель в пиксель (замер 01.10)
    sub.save(out, garbage=4, deflate=True, deflate_fonts=True)
    return {"shrunk": shrink, "pages": len(pages)}


def _page_blocks(psegs):
    """Сегменты страницы → полосы по вертикали: одиночный сегмент — абзац, несколько рядом — строка таблицы.
    Полоса растёт, пока следующий сегмент перекрывает её по высоте (ячейка на несколько строк тянет полосу)."""
    bands = []
    for s in sorted(psegs, key=lambda s: (s["loc"]["bbox"][1], s["loc"]["bbox"][0])):
        y0, y1 = s["loc"]["bbox"][1], s["loc"]["bbox"][3]
        if bands and y0 < bands[-1]["y1"] - 2:
            bands[-1]["segs"].append(s); bands[-1]["y1"] = max(bands[-1]["y1"], y1)
        else:
            bands.append({"segs": [s], "y1": y1})
    # подряд идущие полосы из 2+ сегментов — одна таблица
    blocks = []
    for b in bands:
        if len(b["segs"]) == 1:
            blocks.append(("p", b["segs"][0]))
        elif blocks and blocks[-1][0] == "t":
            blocks[-1][1].append(b["segs"])
        else:
            blocks.append(("t", [b["segs"]]))
    return blocks


def _cluster(vals, tol):
    out = []
    for v in sorted(vals):
        if out and v - out[-1][-1] <= tol:
            out[-1].append(v)
        else:
            out.append([v])
    return [sum(c) / len(c) for c in out]


def _grid(rows, hs, vs):
    """Сетка таблицы: границы столбцов — вертикальные линейки PDF внутри таблицы (иначе кластеры левых краёв),
    границы строк — горизонтальные линейки (иначе полосы). Сегмент — в ячейку по центру своей рамки
    (заголовки ячеек центрированы, содержимое по левому краю: по x0 столбцы двоились, замер 01.10)."""
    segs = [s for row in rows for s in row]
    x0 = min(s["loc"]["bbox"][0] for s in segs); x1 = max(s["loc"]["bbox"][2] for s in segs)
    y0 = min(s["loc"]["bbox"][1] for s in segs); y1 = max(s["loc"]["bbox"][3] for s in segs)
    vx = [x for ry0, ry1, x in vs if x0 + 4 < x < x1 - 4 and min(y1, ry1) - max(y0, ry0) > 4]
    if vx:
        edges = [x0] + _cluster(vx, 3) + [x1]
    else:
        lefts = _cluster([s["loc"]["bbox"][0] for s in segs], 14)
        edges = lefts + [x1]
    # линия строки часто нарисована отрезками по ячейкам — суммарное покрытие на одной высоте
    cover = {}
    for rx0, rx1, y in hs:
        if y0 + 2 < y < y1 - 2:
            key = round(y / 2)
            cover[key] = cover.get(key, 0) + max(0, min(x1, rx1) - max(x0, rx0))
    hy = [k * 2 for k, c in cover.items() if c > 0.3 * (x1 - x0)]
    if hy:
        ys = [y0] + _cluster(hy, 2) + [y1 + 1]
    else:
        ys = [min(s["loc"]["bbox"][1] for s in row) for row in rows] + [y1 + 1]
    ncol, nrow = len(edges) - 1, len(ys) - 1
    cells = {}
    for s in segs:
        bx0, by0, bx1, by1 = s["loc"]["bbox"]
        cx = (bx0 + bx1) / 2 if vx else bx0 + 1
        ci = max(0, min(ncol - 1, next((i for i in range(ncol) if cx < edges[i + 1]), ncol - 1)))
        ri = max(0, min(nrow - 1, next((i for i in range(nrow) if by0 + 1 < ys[i + 1]), nrow - 1)))
        cells.setdefault((ri, ci), []).append(s)
    # пустые строки (между линейками нет текста) — выкинуть
    used = sorted({r for r, _ in cells})
    remap = {r: k for k, r in enumerate(used)}
    return edges, len(used), {(remap[r], c): v for (r, c), v in cells.items()}


def pdf_to_docx(src, out, segs, res):
    """Редактируемый DOCX из PDF: абзацы с отступом/кеглем/жирностью, строки рядом — таблица, страница PDF —
    страница Word. Вёрстка приблизительная (в отличие от PDF «на месте»), зато правится в Word."""
    import fitz
    from docx import Document
    from docx.shared import Pt, Mm
    from docx.enum.text import WD_BREAK
    from docx.oxml.ns import qn
    doc_pdf = fitz.open(src)
    d = Document()
    st = d.styles["Normal"]
    st.font.name = "Times New Roman"; st.font.size = Pt(10)
    st.element.rPr.rFonts.set(qn("w:eastAsia"), "Times New Roman")
    sec = d.sections[0]
    sec.page_width, sec.page_height = Mm(210), Mm(297)
    sec.left_margin = sec.right_margin = Mm(15); sec.top_margin = sec.bottom_margin = Mm(12)
    usable = (210 - 30) / 25.4 * 72                       # ширина набора, pt
    text = lambda s: TAG.sub("", (res.get(str(s["id"])) or {}).get("ru") or s["text"])
    pages = sorted({s["page"] for s in segs})
    tables = 0
    for k, pno in enumerate(pages):
        psegs = [s for s in segs if s["page"] == pno]
        if not psegs:
            continue
        left = min(s["loc"]["bbox"][0] for s in psegs)
        right = max(s["loc"]["bbox"][2] for s in psegs)
        scale = min(1.0, usable / max(right - left, 1))
        hs, vs = rules(doc_pdf[pno])
        for kind, obj in _page_blocks(psegs):
            grid = _grid(obj, hs, vs) if kind == "t" else None
            if kind == "p" or len(grid[0]) == 2:
                # абзац; «таблица» из одного столбца (строки OCR чуть перекрылись по высоте) — тоже абзацы
                for s in ([obj] if kind == "p" else sorted((s for row in obj for s in row),
                                                            key=lambda s: s["loc"]["bbox"][1])):
                    p = d.add_paragraph()
                    p.paragraph_format.left_indent = Pt(max(0, (s["loc"]["bbox"][0] - left) * scale))
                    p.paragraph_format.space_after = Pt(2)
                    run = p.add_run(text(s))
                    run.font.size = Pt(max(6, min(14, s["loc"]["size"])))
                    run.bold = bool(s["loc"].get("bold"))
            else:
                edges, nrow, cells = grid
                t = d.add_table(rows=nrow, cols=len(edges) - 1)
                t.style = "Table Grid"
                tables += 1
                for ci in range(len(edges) - 1):
                    w = Pt(max(20, (edges[ci + 1] - edges[ci]) * scale))
                    for cell in t.columns[ci].cells:
                        cell.width = w
                for (ri, ci), ss in cells.items():
                    cell = t.cell(ri, ci)
                    for j, s in enumerate(sorted(ss, key=lambda s: (s["loc"]["bbox"][1], s["loc"]["bbox"][0]))):
                        p = cell.paragraphs[0] if j == 0 else cell.add_paragraph()
                        run = p.add_run(text(s))
                        run.font.size = Pt(max(6, min(12, s["loc"]["size"] - 0.5)))
                        run.bold = bool(s["loc"].get("bold"))
                d.add_paragraph().paragraph_format.space_after = Pt(0)
        if k < len(pages) - 1:
            d.add_paragraph().add_run().add_break(WD_BREAK.PAGE)
    d.save(out)
    return {"pages": len(pages), "tables": tables}


def _h_note(r):
    """Колонка H: подозрение смысловой проверки, автоисправление или перевод по образцу правки редактора."""
    if (r.get("review") or {}).get("serious"):
        return r["review"].get("reason", "")
    if (r.get("autofix") or {}).get("accepted"):
        return ("Автоисправление (проверить): " + (r["autofix"].get("reason") or "")[:200] + " | было: " +
                TAG.sub("", r["autofix"].get("before") or "")[:300])
    if r.get("rag"):
        return ("По образцу правки редактора (проверить): " + r["rag"].get("example", "")[:200] + " | было: " +
                TAG.sub("", r["rag"].get("before") or "")[:300])
    return ""


def review_xlsx(out, segs, res, title):
    from openpyxl import Workbook
    from openpyxl.styles import Alignment, Font, PatternFill
    wb = Workbook(); ws = wb.active; ws.title = "вычитка"
    ws.append(["№", "стр.", "Венгерский (источник)", "Русский (портал)", "Замечания проверки", "Предупреждения (варианты расходились; термины не найдены)",
               "Исправление (редактор)", "Смысловая проверка (портал)"])
    for c in ws[1]:
        c.font = Font(bold=True)
    # редактор пишет правку в G; команда edits разносит её по одинаковым и похожим фразам, пересобирает документ
    # и кладёт в память переводов проекта — следующие документы возьмут её без модели
    ws.cell(1, 7).fill = PatternFill("solid", fgColor="E2F0D9")
    red = PatternFill("solid", fgColor="FFE0E0"); yel = PatternFill("solid", fgColor="FFF6D0")
    orange = PatternFill("solid", fgColor="FFD8A8"); blue = PatternFill("solid", fgColor="DDEBF7")
    for s in segs:
        r = res.get(str(s["id"]))
        if not r or r.get("copied"):
            continue
        ws.append([s["id"], (s["page"] + 1) if s["page"] is not None else "", TAG.sub("", s["text"]),
                   TAG.sub("", r["ru"]), "; ".join(r["issues"]),
                   "; ".join(r.get("warn", []) + r.get("gloss_miss", [])), None, _h_note(r)])
        row = ws.max_row
        # автоисправление (autofix.py): слепая вычитка 05.10 — серьёзных 32 → 14 на 123, но 12 внесено заново → голубым
        if (r.get("autofix") or {}).get("accepted") or r.get("rag"):
            for col in (4, 8):
                ws.cell(row, col).fill = blue
        # смысловая проверка агентом портала (review_run.py): подозрение на искажение смысла — оранжевым
        if (r.get("review") or {}).get("serious"):
            for col in (4, 8):
                ws.cell(row, col).fill = orange
        if r["issues"]:
            for c in ws[row]:
                c.fill = red
        elif r.get("warn") or r.get("gloss_miss"):
            ws.cell(row, 6).fill = yel
    for col, w in zip("ABCDEFGH", (6, 6, 70, 70, 30, 30, 70, 50)):
        ws.column_dimensions[col].width = w
    for row in ws.iter_rows(min_row=2):
        for c in row:
            c.alignment = Alignment(wrap_text=True, vertical="top")
    ws.freeze_panes = "A2"
    wb.save(out)


def source_path(d):
    """Путь к исходнику из JSON прогона (конвейер пишет абсолютный путь)."""
    return Path(d["file"])


def assemble_all(jf, d, out_dir=None):
    """Перевод в исходном формате (DOCX на месте / PDF на месте + DOCX по раскладке) и таблица вычитки.
    out_dir — куда класть файлы (по умолчанию рядом с JSON)."""
    src = source_path(d)
    outd, stem = Path(out_dir or Path(jf).parent), Path(jf).stem
    outd.mkdir(parents=True, exist_ok=True)
    if src.suffix.lower() == ".docx":
        info = build_docx(src, outd / f"{stem}_RU.docx", d["segs"], d["res"])
    else:
        info = build_pdf(src, outd / f"{stem}_RU.pdf", d["segs"], d["res"])
        info["docx"] = pdf_to_docx(src, outd / f"{stem}_RU.docx", d["segs"], d["res"])
    review_xlsx(outd / f"{stem}_вычитка.xlsx", d["segs"], d["res"], stem)
    return info
