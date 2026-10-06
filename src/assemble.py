# -*- coding: utf-8 -*-
"""Сборка перевода обратно в формат источника + двуязычная таблица для вычитки.

- DOCX: текст каждой группы прогонов (одинаковое начертание) пишется в первый прогон группы,
  остальные прогоны группы опустошаются; метки <fN> раскладывают перевод по группам.
  Если метки сломаны — весь перевод в первую группу (начертание первой группы).
- PDF: исходный текст переведённых абзацев убирается (текстовый слой — redaction текста;
  OCR-страница — белая заливка + удаление нарисованных букв), перевод вписывается в рамку
  абзаца, расширенную до соседей (insert_htmlbox ужимает кегль, если не влезает).
  DOCX для PDF-исходника — из переведённого PDF конвертером pdf2docx (docx_from_pdf)."""
import re, html, zipfile
from pathlib import Path

from segments import W, _own_runs

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


def _free(rect, busy, own):
    """Рамка под перевод без мест, уже занятых соседями: занятое ниже начала — обрезает рамку снизу, правее —
    справа. Свою исходную строку (own) рамка не теряет: перевод «Номер версии», перенесённый на вторую строку,
    больше не накрывается длинным переводом строки ниже (30RDGT nyilv, 06.10)."""
    import fitz
    r = fitz.Rect(rect)
    for b in busy:
        if not r.intersects(b):
            continue
        if b.y0 > r.y0 + 1 and b.y0 >= own[3] - 0.5:
            r.y1 = min(r.y1, b.y0 - 0.5)
        elif b.x0 > r.x0 + 1 and b.x0 >= own[2] - 0.5:
            r.x1 = min(r.x1, b.x0 - 1)
    return r


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
                # OCR-страница: белая заливка по рамкам строк абзаца, а не по общей рамке — общая рамка формы
                # захватывала соседние поля и рукописные подписи (30RDGT nyilv, 06.10); у первой строки — запас
                # влево под тире/номер пункта, которые Tesseract мог не включить в рамку
                for i, (x0, y0, x1, y1) in enumerate(s["loc"]["frags"]):
                    left = 16 if i == 0 and re.match(r"\s*[-–•]", s["text"]) else 2
                    page.add_redact_annot(fitz.Rect(x0 - left, y0 - 1.5, x1 + 2, y1 + 1.5), fill=(1, 1, 1))
            else:
                for bb in s["loc"]["frags"]:
                    page.add_redact_annot(fitz.Rect(bb), fill=False)
        # скан: пиксели под белой плашкой стираются в самой картинке — иначе её вытаскивает конвертер в DOCX
        # и под русским текстом проступает венгерский (30RDGT nyilv, 06.10)
        page.apply_redactions(images=fitz.PDF_REDACT_IMAGE_PIXELS if ocr else fitz.PDF_REDACT_IMAGE_NONE,
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
            items.append((s, rect, txt, align, loc["size"], loc.get("family", "serif")))
        # общий масштаб страницы: пробная вставка на пустой странице, медиана нужных масштабов
        trial = fitz.open(); tp = trial.new_page(width=page.rect.width, height=page.rect.height)
        need = []
        for s, rect, txt, align, size, fam in items:
            css = f"* {{font-family: {fam}; font-size: {size:.1f}pt; line-height: 1.15; text-align: {align};}}"
            need.append(tp.insert_htmlbox(rect, txt, css=css, scale_low=0)[1])
        k = sorted(need)[int(len(need) * 0.3)] if need else 1.0
        if ocr:
            # скан формы: тесные ячейки не должны ужимать всю страницу (подписи полей уходили в 3 pt) —
            # общий масштаб не ниже 0,85, дальше каждая ячейка ужимается сама
            k = max(k, 0.85)
        busy = []                     # занятое уже вписанными абзацами: следующие его не получают
        order = sorted(range(len(items)), key=lambda i: (round(items[i][1].y0 / 3), items[i][1].x0))
        for i in order:
            s, rect, txt, align, size, fam = items[i]
            rect = _free(rect, busy, s["loc"]["bbox"])
            sz = size * min(k, 1.0)
            # семейство шрифта — как в оригинале (моноширинный Courier 3SZ19 вписывался засечным шрифтом)
            css = f"* {{font-family: {fam}; font-size: {sz:.1f}pt; line-height: 1.15; text-align: {align};}}"
            spare, scale = page.insert_htmlbox(rect, txt, css=css, scale_low=0)
            busy.append(fitz.Rect(rect.x0, rect.y0, rect.x1, rect.y1 - max(spare, 0)))
            if scale * min(k, 1.0) < 0.7:
                shrink.append({"page": pno + 1, "id": s["id"], "scale": round(scale * min(k, 1.0), 2)})
        sub.insert_pdf(doc, from_page=pno, to_page=pno)
    # отдаём только обработанные страницы: и для выборки, и для полного прогона это вся работа
    # каждая вставка абзаца встраивает свой экземпляр шрифта (5275 объектов на 157 стр. ÜFK I, 43,8 МБ);
    # garbage=4 склеивает одинаковые объекты: 3,3 МБ, рендер совпадает пиксель в пиксель (замер 01.10)
    sub.save(out, garbage=4, deflate=True, deflate_fonts=True)
    return {"shrunk": shrink, "pages": len(pages)}


P2D = {"line_break_width_ratio": 1.0, "line_overlap_threshold": 1.0}


def _merge_docx(paths, out):
    """Склейка постраничных DOCX в один: каждая страница — свой раздел (размер и поля как в PDF), картинки
    переносятся с новыми связями (get_or_add_image — без конфликтов имён частей)."""
    import io
    import copy
    from docx import Document
    from docx.oxml.ns import qn
    target = Document(paths[0])
    body = target.element.body
    for path in paths[1:]:
        src = Document(path)
        last_sect = body.find(qn("w:sectPr"))
        # конец текущего раздела: его sectPr уходит в пустой абзац, разрыв — со следующей страницы
        p = body.makeelement(qn("w:p"), {})
        ppr = p.makeelement(qn("w:pPr"), {})
        ppr.append(copy.deepcopy(last_sect))
        p.append(ppr)
        last_sect.addprevious(p)
        for el in list(src.element.body):
            if el.tag == qn("w:sectPr"):
                continue
            el = copy.deepcopy(el)
            for node in el.iter():
                for attr in (qn("r:embed"), qn("r:link"), qn("r:id")):
                    rid = node.get(attr)
                    if rid and rid in src.part.rels and "image" in src.part.rels[rid].reltype:
                        new_rid, _ = target.part.get_or_add_image(io.BytesIO(src.part.rels[rid].target_part.blob))
                        node.set(attr, new_rid)
            last_sect.addprevious(el)
        body.replace(last_sect, copy.deepcopy(src.element.body.find(qn("w:sectPr"))))
    target.save(out)


def docx_from_pdf(pdf, out):
    """DOCX по переведённому PDF — конвертером pdf2docx: таблицы и разбивка по страницам как в PDF.
    Прежняя самодельная сборка по блокам страницы разваливала таблицы с колонками разной ширины (30RDGT feladat,
    06.10: 129 таблиц вместо одной на страницу, 32 страницы вместо 15).
    - Каждая страница конвертируется отдельно и потом склеивается: при разборе нескольких страниц сразу pdf2docx 0.5.13
      молча теряет часть мелких надписей (подписи «Примечание:» — 57 из 116; по одной странице — все).
    - line_break_width_ratio=1 — переносы строк как в PDF (иначе строки абзаца склеивались без пробела:
      «уполномоченныйпредставитель»); line_overlap_threshold=1 — выбрасывать только полные дубли строк.
    Сбой конвертера не валит перевод: DOCX нет, PDF и таблица вычитки есть."""
    import logging
    import tempfile
    try:
        import fitz
        from pdf2docx import Converter
    except ImportError:
        return {"error": "pdf2docx не установлен"}
    level = logging.getLogger().level
    logging.getLogger().setLevel(logging.ERROR)            # pdf2docx пишет по строке на страницу в INFO
    try:
        n = fitz.open(str(pdf)).page_count
        with tempfile.TemporaryDirectory() as tmp:
            parts = []
            for pno in range(n):
                part = Path(tmp) / f"p{pno:04d}.docx"
                # pdf2docx падает на некоторых таблицах («Failed to merge docx_cell», титул 30RDGT, 06.10) и
                # молча отдаёт пустую страницу — тогда повтор без распознавания таблиц по тексту, потом совсем без таблиц
                for extra in ({}, {"parse_stream_table": False}, {"parse_stream_table": False, "parse_lattice_table": False}):
                    cv = Converter(str(pdf))
                    try:
                        cv.convert(str(part), pages=[pno], raw_exceptions=True, **P2D, **extra)
                        break
                    except Exception:                      # noqa: BLE001 — следующий набор настроек
                        continue
                    finally:
                        cv.close()
                parts.append(part)
            _merge_docx(parts, out)
        return {"converter": "pdf2docx", "pages": n}
    except Exception as e:                                  # noqa: BLE001 — сторонний конвертер, перевод уже собран
        Path(out).unlink(missing_ok=True)
        return {"error": f"{type(e).__name__}: {e}"[:200]}
    finally:
        logging.getLogger().setLevel(level)


def _h_note(r):
    """Колонка H: подозрение смысловой проверки, автоисправление или перевод по образцу правки редактора."""
    if (r.get("review") or {}).get("serious"):
        return r["review"].get("reason", "")
    if (r.get("autofix") or {}).get("accepted"):
        return ("Автоисправление (проверить): " + (r["autofix"].get("reason") or "")[:200] + " | было: " +
                TAG.sub("", r["autofix"].get("before") or "")[:300])
    if r.get("rag") and r["rag"].get("numbers"):
        return "Правка редактора, подставлены свои числа: " + r["rag"].get("example", "")[:200]
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
        if (r.get("autofix") or {}).get("accepted") or (r.get("rag") and not r["rag"].get("numbers")):
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
    """Перевод в исходном формате (DOCX на месте / PDF на месте + DOCX из переведённого PDF) и таблица вычитки.
    out_dir — куда класть файлы (по умолчанию рядом с JSON)."""
    src = source_path(d)
    outd, stem = Path(out_dir or Path(jf).parent), Path(jf).stem
    outd.mkdir(parents=True, exist_ok=True)
    if src.suffix.lower() == ".docx":
        info = build_docx(src, outd / f"{stem}_RU.docx", d["segs"], d["res"])
    else:
        info = build_pdf(src, outd / f"{stem}_RU.pdf", d["segs"], d["res"])
        info["docx"] = docx_from_pdf(outd / f"{stem}_RU.pdf", outd / f"{stem}_RU.docx")
    review_xlsx(outd / f"{stem}_вычитка.xlsx", d["segs"], d["res"], stem)
    return info
