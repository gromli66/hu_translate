# -*- coding: utf-8 -*-
"""Сборка перевода обратно в формат источника + двуязычная таблица для вычитки.

- DOCX: текст каждой группы прогонов (одинаковое начертание) пишется в первый прогон группы,
  остальные прогоны группы опустошаются; метки <fN> раскладывают перевод по группам.
  Если метки сломаны — весь перевод в первую группу (начертание первой группы).
- PDF: исходный текст переведённых абзацев убирается (текстовый слой — redaction текста;
  OCR-страница — белая заливка + удаление нарисованных букв), перевод вписывается в рамку
  абзаца, расширенную до соседей (insert_htmlbox ужимает кегль, если не влезает).
  DOCX для PDF-исходника — из переведённого PDF конвертером pdf2docx (docx_from_pdf); для размеченного PDF
  (экспорт Word) — по дереву разметки и профилю оформления оригинала (build_struct_docx)."""
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


TABLE_FLOOR = 0.7


def pdf_text(r):
    """Текст для PDF: сокращённая формулировка (pipeline.fit), пока перевод тот же, из которого она сделана —
    правка редактора её отменяет."""
    return r["ru_pdf"] if r.get("ru_pdf") and r.get("ru_pdf_of") == r["ru"] else r["ru"]


TOC = re.compile(r"^(?P<t>.*?\S)[\s.…]*?\.{3,}[\s.…]*(?P<n>\d{1,4})\s*$", re.S)


def _width(txt, css):
    """Ширина строки в pt тем же шрифтом, каким её впишет insert_htmlbox (у базового Helvetica нет кириллицы)."""
    import fitz
    import numpy as np
    doc = fitz.open(); pg = doc.new_page(width=2000, height=80)
    pg.insert_htmlbox(fitz.Rect(0, 0, 2000, 80), txt, css=css)
    pix = pg.get_pixmap(dpi=216, colorspace=fitz.csGRAY, clip=fitz.Rect(0, 0, 2000, 80))
    a = np.frombuffer(pix.samples, np.uint8).reshape(pix.height, pix.width)
    cols = np.nonzero((a < 160).any(axis=0))[0]
    return (cols.max() + 1) * 72 / 216 if len(cols) else 0.0


def _toc_line(page, s, r, rect, fam, sz, busy):
    """Строка оглавления «текст ….. 19»: номер — у правого края, как в оригинале, точки — ровно до номера.
    Точки-заполнители переводились буквально, их число не менялось — номера страниц стояли вразнобой, длинные
    строки уносили номер на следующую строку (оглавление UTS, 15.10). Не влезает и при 75 % кегля — обычная вставка."""
    import fitz
    m_src, m_ru = TOC.match(TAG.sub("", s["text"])), TOC.match(TAG.sub("", pdf_text(r)).strip())
    if not (m_src and m_ru) or len(s["loc"]["frags"]) != 1:
        return False
    title, num = m_ru.group("t").strip(" .…"), m_ru.group("n")
    right = s["loc"]["bbox"][2]
    bold = "font-weight: bold;" if s["loc"].get("bold") else ""
    for k in (1.0, 0.9, 0.8, 0.75):
        css = f"* {{font-family: {fam}; font-size: {sz * k:.2f}pt; white-space: nowrap; {bold}}}"
        w_num, w_dot, w_t = _width(num, css), _width("." * 20, css) / 20, _width(html.escape(title), css)
        room = right - rect.x0 - w_num - 2
        if w_t + 4 * w_dot <= room:
            break
    else:
        return False
    line = html.escape(title) + " " + "." * max(int((room - w_t) / w_dot) - 2, 2)
    page.insert_htmlbox(fitz.Rect(rect.x0, rect.y0, right - w_num - 1, rect.y1), line, css=css, scale_low=0)
    page.insert_htmlbox(fitz.Rect(right - w_num - 6, rect.y0, right + 1, rect.y1), num,
                        css=css.replace("white-space: nowrap;", "white-space: nowrap; text-align: right;"), scale_low=0)
    busy.append(fitz.Rect(rect.x0, rect.y0, right + 1, rect.y0 + sz * 1.3))
    return True


def build_pdf(src, out, segs, res, progress=None, shots_out=None):
    """shots_out — куда сохранить оригинал с переведёнными скриншотами (без правки текста): из него собирается DOCX."""
    import fitz
    import shots as SH
    doc = fitz.open(src)
    pages = sorted({s["page"] for s in segs if s["page"] is not None and 0 <= s["page"] < len(doc)})
    n_shots, shot_err = SH.apply(doc, segs, res)
    if shots_out:
        doc.save(shots_out, garbage=3, deflate=True)
    # текст страницы — сегменты с рамками строк; надписи скриншотов и абзацы DOCX вписываются не здесь
    segs = [s for s in segs if "frags" in s["loc"]]
    shrink = []
    sub = fitz.open()
    for k, pno in enumerate(pages):
        if progress:
            progress(k / len(pages))
        page = doc[pno]
        # координаты сегментов — как страница видна; заливка и вставка — в системе неповёрнутой страницы,
        # текст — с поворотом страницы (32RDGT: листы с /Rotate 270 получали перевод поперёк листа, 07.10)
        der = page.derotation_matrix
        todo = [(s, res[str(s["id"])]) for s in segs if s["page"] == pno
                and not res[str(s["id"])].get("copied") and res[str(s["id"])]["ru"]]
        ocr = any(s["loc"].get("ocr") for s, _ in todo)
        # таблицы — до вычистки текста (find_tables смотрит и на текст); на сканах и повёрнутых листах — нет
        # таблица больше половины листа — рамка формы (TU-документы): для неё прежний масштаб страницы, иначе
        # минимум по всей форме мельчил её целиком (30RDIT_TU: медиана 10 → 9 pt, красная команда 15.10)
        tabs = [] if ocr or page.rotation else [r for r in (fitz.Rect(t.bbox) for t in page.find_tables().tables)
                                                 if r.get_area() < 0.5 * page.rect.get_area()]
        for s, r in todo:
            if ocr:
                # OCR-страница: белая заливка по рамкам строк абзаца, а не по общей рамке — общая рамка формы
                # захватывала соседние поля и рукописные подписи (30RDGT nyilv, 06.10); у первой строки — запас
                # влево под тире/номер пункта, которые Tesseract мог не включить в рамку
                for i, (x0, y0, x1, y1) in enumerate(s["loc"]["frags"]):
                    left = 16 if i == 0 and re.match(r"\s*[-–•]", s["text"]) else 2
                    page.add_redact_annot(fitz.Rect(x0 - left, y0 - 1.5, x1 + 2, y1 + 1.5) * der, fill=(1, 1, 1))
            else:
                for bb in s["loc"]["frags"]:
                    page.add_redact_annot(fitz.Rect(bb) * der, fill=False)
        # скан: пиксели под белой плашкой стираются в самой картинке — иначе её вытаскивает конвертер в DOCX
        # и под русским текстом проступает венгерский (30RDGT nyilv, 06.10)
        page.apply_redactions(images=fitz.PDF_REDACT_IMAGE_PIXELS if ocr else fitz.PDF_REDACT_IMAGE_NONE,
                              graphics=fitz.PDF_REDACT_LINE_ART_REMOVE_IF_COVERED if ocr else fitz.PDF_REDACT_LINE_ART_NONE,
                              text=fitz.PDF_REDACT_TEXT_REMOVE)
        boxes = [s["loc"]["bbox"] for s in segs if s["page"] == pno] or [[0, 0, page.rect.width, 0]]
        fl, fr = min(b[0] for b in boxes), max(b[2] for b in boxes)
        items = []
        for s, r in todo:
            loc = s["loc"]
            x0, y0, x1, y1 = loc["bbox"]
            ax1, ay1 = loc["avail"][2], loc["avail"][3]
            rect = fitz.Rect(x0, y0 - 0.5, max(ax1, x1), max(ay1, y1) + 0.5)
            align = "left"
            if "align" in loc:                      # выключка и ячейка определены при извлечении (layout.aligns)
                if loc["align"] == "center":
                    cl, cr = loc["cell"]
                    half = min((x0 + x1) / 2 - cl, cr - (x0 + x1) / 2) - 2
                    rect = fitz.Rect((x0 + x1) / 2 - half, y0 - 0.5, (x0 + x1) / 2 + half, max(ay1, y1) + 0.5)
                    align = "center"
            # старые задачи: центр — при равных полях слева и справа от рамки текста страницы
            elif abs((x0 - fl) - (fr - x1)) < 12 and (x1 - x0) < 0.7 * (fr - fl) and x0 - fl > 20:
                half = min((x0 + x1) / 2 - fl, fr - (x0 + x1) / 2)
                rect = fitz.Rect((x0 + x1) / 2 - half, y0 - 0.5, (x0 + x1) / 2 + half, max(ay1, y1) + 0.5)
                align = "center"
            # ru_pdf — сокращённая порталом формулировка для тесного места (pipeline.fit); в DOCX — полный перевод
            txt = html.escape(TAG.sub("", pdf_text(r))).replace("\n", "<br>").replace("\t", "&emsp;")
            if loc.get("bold"):
                txt = f"<b>{txt}</b>"
            items.append((s, rect, txt, align, loc["size"], loc.get("family", "serif")))
        # общий масштаб страницы: пробная вставка на пустой странице, медиана нужных масштабов
        trial = fitz.open(); tp = trial.new_page(width=page.rect.width, height=page.rect.height)
        need = []
        for s, rect, txt, align, size, fam in items:
            css = f"* {{font-family: {fam}; font-size: {size:.1f}pt; line-height: 1.15; text-align: {align};}}"
            need.append(tp.insert_htmlbox(rect, txt, css=css, scale_low=0)[1])
        # таблица — один масштаб на все ячейки (минимум нужных, не ниже TABLE_FLOOR; ниже — ячейка ужимается сама):
        # каждая ячейка ужималась отдельно — рядом стояли 9 и 5 pt (UTS, таблица 1, 15.10); масштаб страницы —
        # только по тексту вне таблиц, тесная таблица не ужимает абзацы
        tab_of = [next((ti for ti, tb in enumerate(tabs) if tb.contains((rect.tl + rect.br) / 2)), None)
                  for _, rect, *_ in items]
        free = [n for n, t in zip(need, tab_of) if t is None]
        k = sorted(free)[int(len(free) * 0.3)] if free else 1.0
        if ocr:
            # скан формы: тесные ячейки не должны ужимать всю страницу (подписи полей уходили в 3 pt) —
            # общий масштаб не ниже 0,85, дальше каждая ячейка ужимается сама
            k = max(k, 0.85)
        groups = {}
        for n, t in zip(need, tab_of):
            if t is not None:
                groups.setdefault(t, []).append(n)
        # не мельче 30-го перцентиля таблицы минус 0,15: одна совсем тесная ячейка ужимается сама, не тянет всю таблицу
        kt = {t: max(min(v), sorted(v)[int(len(v) * 0.3)] - 0.15, TABLE_FLOOR) for t, v in groups.items()}
        busy = []                     # занятое уже вписанными абзацами: следующие его не получают
        order = sorted(range(len(items)), key=lambda i: (round(items[i][1].y0 / 3), items[i][1].x0))
        for i in order:
            s, rect, txt, align, size, fam = items[i]
            rect = _free(rect, busy, s["loc"]["bbox"])
            ki = kt[tab_of[i]] if tab_of[i] is not None else k
            sz = size * min(ki, 1.0)
            if not ocr and not page.rotation and _toc_line(page, s, res[str(s["id"])], rect, fam, sz, busy):
                continue
            # семейство шрифта — как в оригинале (моноширинный Courier 3SZ19 вписывался засечным шрифтом)
            css = f"* {{font-family: {fam}; font-size: {sz:.1f}pt; line-height: 1.15; text-align: {align};}}"
            spare, scale = page.insert_htmlbox(rect * der, txt, css=css, scale_low=0, rotate=page.rotation)
            busy.append(fitz.Rect(rect.x0, rect.y0, rect.x1, rect.y1 - max(spare, 0)))
            if scale * min(ki, 1.0) < 0.7:
                shrink.append({"page": pno + 1, "id": s["id"], "scale": round(scale * min(ki, 1.0), 2)})
        sub.insert_pdf(doc, from_page=pno, to_page=pno)
    # отдаём только обработанные страницы: и для выборки, и для полного прогона это вся работа
    # каждая вставка абзаца встраивает свой экземпляр шрифта (5275 объектов на 157 стр. ÜFK I, 43,8 МБ);
    # garbage=4 склеивает одинаковые объекты: 3,3 МБ, рендер совпадает пиксель в пиксель (замер 01.10)
    sub.save(out, garbage=4, deflate=True, deflate_fonts=True)
    info = {"shrunk": shrink, "pages": len(pages)}
    if n_shots or shot_err:
        info["shots"] = n_shots
        if shot_err:
            info["shot_errors"] = shot_err
    return info


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
    pages_done = 0
    for path in paths[1:]:
        src = Document(path)
        last_sect = body.find(qn("w:sectPr"))
        # конец текущего раздела: его sectPr уходит в пустой абзац, разрыв — со следующей страницы
        p = body.makeelement(qn("w:p"), {})
        ppr = p.makeelement(qn("w:pPr"), {})
        ppr.append(copy.deepcopy(last_sect))
        p.append(ppr)
        # закладка «_hutPage» — граница страницы PDF: pdf2docx ставит свои разрывы разделов и внутри страницы,
        # по ним номер страницы абзаца уезжал (ÜFK I: 45 разрывов на 40 страниц, красная команда 15.10)
        bid = str(900000 + pages_done)
        p.append(p.makeelement(qn("w:bookmarkStart"), {qn("w:id"): bid, qn("w:name"): f"_hutPage{pages_done}"}))
        p.append(p.makeelement(qn("w:bookmarkEnd"), {qn("w:id"): bid}))
        pages_done += 1
        last_sect.addprevious(p)
        for el in list(src.element.body):
            if el.tag == qn("w:sectPr"):
                continue
            el = copy.deepcopy(el)
            for node in el.iter():
                for attr in (qn("r:embed"), qn("r:link"), qn("r:id")):
                    rid = node.get(attr)
                    if not rid or rid not in src.part.rels:
                        continue
                    rel = src.part.rels[rid]
                    # связь страницы переносится в общий документ: картинка — частью, внешняя ссылка (почта, сайт,
                    # ссылки оглавления) — новой связью; иначе r:id висит в пустоте и Word считает файл повреждённым
                    # (DOCX из оригинала UTS: ссылки на стр. 90 и 99, 15.10)
                    if "image" in rel.reltype:
                        new_rid, _ = target.part.get_or_add_image(io.BytesIO(rel.target_part.blob))
                        node.set(attr, new_rid)
                    elif rel.is_external:
                        node.set(attr, target.part.relate_to(rel.target_ref, rel.reltype, is_external=True))
                    else:
                        del node.attrib[attr]
            last_sect.addprevious(el)
        body.replace(last_sect, copy.deepcopy(src.element.body.find(qn("w:sectPr"))))
    target.save(out)


def docx_from_pdf(pdf, out, progress=None):
    """DOCX по PDF — конвертером pdf2docx: таблицы и разбивка по страницам как в PDF.
    Прежняя самодельная сборка по блокам страницы разваливала таблицы с колонками разной ширины (30RDGT feladat,
    06.10: 129 таблиц вместо одной на страницу, 32 страницы вместо 15).
    - Конвертер работает в отдельном процессе: pdf2docx вызывает page.get_texttrace, который в PyMuPDF 1.27 портит
      счётчик ссылок на None — после двух конвертаций (DOCX из оригинала: при извлечении и при сборке) процесс
      падал на выходе (none_dealloc), и сервер считал готовую задачу сбойной (красная команда, 15.10).
      Успех — по готовому файлу и отметке «OK» дочернего процесса, не по коду выхода.
    - Каждая страница конвертируется отдельно и потом склеивается: при разборе нескольких страниц сразу pdf2docx 0.5.13
      молча теряет часть мелких надписей (подписи «Примечание:» — 57 из 116; по одной странице — все).
    - line_break_width_ratio=1 — переносы строк как в PDF (иначе строки абзаца склеивались без пробела:
      «уполномоченныйпредставитель»); line_overlap_threshold=1 — выбрасывать только полные дубли строк.
    Сбой конвертера не валит перевод: DOCX нет, PDF и таблица вычитки есть."""
    import subprocess
    import sys
    Path(out).unlink(missing_ok=True)
    cmd = [sys.executable, "-X", "utf8", "-c",
           "import sys; sys.path.insert(0, sys.argv[1]); import assemble; assemble._p2d_child(sys.argv[2], sys.argv[3])",
           str(Path(__file__).resolve().parent), str(pdf), str(out)]
    try:
        proc = subprocess.Popen(cmd, stdout=subprocess.PIPE, stderr=subprocess.PIPE, text=True, encoding="utf-8",
                                errors="replace")
    except OSError as e:
        return {"error": f"{type(e).__name__}: {e}"[:200]}
    ok, err = None, ""
    for ln in proc.stdout:
        ln = ln.strip()
        if ln.startswith("PAGE ") and progress:
            i, n = map(int, ln.split()[1:3])
            progress(i / max(n, 1))
        elif ln.startswith("OK "):
            ok = int(ln.split()[1])
        elif ln.startswith("ERR "):
            err = ln[4:]
    proc.wait()
    if ok is not None and Path(out).exists():
        return {"converter": "pdf2docx", "pages": ok}
    Path(out).unlink(missing_ok=True)
    return {"error": (err or (proc.stderr.read() if proc.stderr else "") or f"код {proc.returncode}")[:200]}


def _p2d_child(pdf, out):
    """Дочерний процесс docx_from_pdf: постраничная конвертация и склейка; ход — строками «PAGE i n» в stdout."""
    import logging
    import tempfile
    try:
        import fitz
        from pdf2docx import Converter
    except ImportError:
        print("ERR pdf2docx не установлен", flush=True); return
    logging.getLogger().setLevel(logging.ERROR)            # pdf2docx пишет по строке на страницу в INFO
    try:
        n = fitz.open(str(pdf)).page_count
        with tempfile.TemporaryDirectory() as tmp:
            parts = []
            for pno in range(n):
                print(f"PAGE {pno} {n}", flush=True)
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
        print(f"OK {n}", flush=True)
    except Exception as e:                                  # noqa: BLE001 — сторонний конвертер, перевод уже собран
        Path(out).unlink(missing_ok=True)
        print(" ".join(f"ERR {type(e).__name__}: {e}"[:200].split()), flush=True)


def build_struct_docx(src, out, segs, res, rules=None, progress=None):
    """DOCX размеченного PDF (экспорт Word со StructTreeRoot, pipeline.extract → d["struct"]): блоки — по дереву
    разметки (docx_struct), оформление — профилем оригинала (docx_profile), сборка — docx_build.Builder; перевод — из
    сегментов задачи (после проверки, шлифовки и правок редактора), скриншоты — переведённые (shots.render).
    Готовый DOCX проверяется: оформление против профиля, языковые соглашения (rules — d["polish"]["rules"]) по
    русскому тексту. Возвращает (сведения, строки листа «нарушения»)."""
    import fitz
    import conventions as C
    import docx_build as B
    import docx_profile as F
    import docx_struct as D
    import shots as SH
    from docx import Document
    from docx.text.paragraph import Paragraph
    tick = progress or (lambda x: None)
    blocks, _ = D.load(str(src))
    L = D.layout(str(src), blocks)
    prof = F.extract(str(src), blocks, L)
    tick(0.4)
    mem = B.job_memory(segs, res)
    doc = fitz.open(src)
    shot_items = {}                     # те же надписи, что в PDF (shots.apply)
    for s in segs:
        sh, r = s["loc"].get("shot"), res.get(str(s["id"])) or {}
        ru = TAG.sub("", r.get("ru") or "").strip()
        if sh and ru and not r.get("copied") and ru != s["text"]:
            shot_items.setdefault(sh["xref"], []).append((sh, ru))
    cache = {}

    def image(img):
        x = img["stream"].objid
        if x in cache:
            return cache[x]
        data = None
        try:
            masked = any(doc.xref_get_key(x, k)[0] != "null" for k in ("SMask", "Mask")) \
                or doc.xref_get_key(x, "ImageMask")[1] == "true"
            if x in shot_items and not masked:
                data = SH.render(doc, x, shot_items[x]).tobytes("png")          # переведённый скриншот
            elif masked:
                pg = doc[(img.get("page_number") or 2) - 1]
                clip = fitz.Rect(float(img["x0"]), float(img["top"]), float(img["x1"]), float(img["bottom"]))
                data = pg.get_pixmap(dpi=300, clip=clip).tobytes("png")         # прозрачность — снимок с листа на белом
            else:
                pix = fitz.Pixmap(doc, x)
                if pix.alpha or (pix.colorspace and pix.colorspace.n not in (1, 3)):
                    pix = fitz.Pixmap(fitz.csRGB, pix)
                data = pix.tobytes("png")
        except Exception as e:                                                  # noqa: BLE001 — картинка не валит сборку
            print(f"DOCX по разметке: картинка {x} не вставлена: {type(e).__name__}: {e}"[:200], flush=True)
        cache[x] = data
        return data

    B.Builder(mem, image, prof).build(blocks, str(out), L)
    tick(0.8)
    viol = F.check(str(out), prof, blocks, L)
    dx, lang = Document(str(out)), []
    for el in dx.element.body.iter(B.W + "p"):
        p = Paragraph(el, dx)
        if p.style.name not in ("Code UTS", "Code Center UTS"):              # код — дословно, соглашения не к нему
            lang += [(r, why, p.text[:90]) for r, why in C.check_text(p.text, rules or {})]
    rows = [("", v["rule"], f"оформление: {v['where']} — {v['was']} (нужно: {v['need']})", "") for v in viol] \
        + [("", r, f"в DOCX: {why}", t) for r, why, t in lang]
    tick(1.0)
    return {"converter": "разметка PDF", "missed": len(set(mem.missing)), "format": len(viol), "lang": len(lang)}, rows


def _h_note(r):
    """Колонка H: подозрение смысловой проверки, автоисправление или перевод по образцу правки редактора;
    правка шлифовки (pipeline.polish), пока перевод тот же; сокращённая для PDF формулировка (pipeline.fit)."""
    note = _h_note_main(r)
    p = r.get("polish")
    if p and p.get("of") == r["ru"]:
        note = (note + " | " if note else "") + f"Шлифовка ({p['rules']}; {p['who']}), было: " + TAG.sub("", p["before"])[:300]
    if pdf_text(r) != r["ru"]:
        note = (note + " | " if note else "") + "В PDF — сокращённо (не помещалось): " + r["ru_pdf"]
    return note


def _h_note_main(r):
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


def review_xlsx(out, segs, res, title, polish=None, extra_rows=()):
    """Лист «вычитка»; после шлифовки (polish — d["polish"]: журнал и оставшиеся нарушения) или проверки DOCX по
    разметке (extra_rows — нарушения оформления и соглашений в готовом DOCX) — ещё листы «шлифовка» и «нарушения»."""
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
                   "; ".join(r.get("warn", []) + r.get("gloss_miss", [])), None,
                   " | ".join(x for x in ("Надпись на скриншоте" if s["loc"].get("shot") else "", _h_note(r)) if x)])
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
    if polish is not None or extra_rows:
        import polish as PL
        PL.add_sheets(wb, (polish or {}).get("edits", []), (polish or {}).get("remaining", []), extra_rows)
    wb.save(out)


def source_path(d):
    """Путь к исходнику из JSON прогона (конвейер пишет абсолютный путь)."""
    return Path(d["file"])


def assemble_all(jf, d, out_dir=None, progress=None):
    """Перевод в исходном формате (DOCX на месте / PDF на месте + DOCX по разметке, из оригинала или из переведённого
    PDF) и таблица вычитки.
    out_dir — куда класть файлы (по умолчанию рядом с JSON). progress(доля 0..1) — ход по страницам PDF и DOCX."""
    tick = progress or (lambda x: None)
    src = source_path(d)
    outd, stem = Path(out_dir or Path(jf).parent), Path(jf).stem
    outd.mkdir(parents=True, exist_ok=True)
    rows = []
    if src.suffix.lower() == ".docx":
        info = build_docx(src, outd / f"{stem}_RU.docx", d["segs"], d["res"])
    else:
        if d.get("struct"):
            # размеченный PDF: DOCX по дереву разметки и профилю оформления оригинала
            info = build_pdf(src, outd / f"{stem}_RU.pdf", d["segs"], d["res"], progress=lambda x: tick(0.4 * x))
            try:
                info["docx"], rows = build_struct_docx(src, outd / f"{stem}_RU.docx", d["segs"], d["res"],
                                                       (d.get("polish") or {}).get("rules"),
                                                       progress=lambda x: tick(0.4 + 0.6 * x))
            except Exception as e:                      # noqa: BLE001 — DOCX прежним путём, причина — в отчёт сборки
                info["docx"] = docx_from_pdf(outd / f"{stem}_RU.pdf", outd / f"{stem}_RU.docx",
                                             progress=lambda x: tick(0.4 + 0.6 * x))
                info["docx"]["fallback"] = f"{type(e).__name__}: {e}"[:200]
        elif d.get("docx"):
            # PDF с текстовым слоем: DOCX — из оригинала с переведёнными скриншотами, перевод — в абзацы Word
            import docx_src
            shots_pdf = outd / f"{stem}_shots.tmp.pdf"
            try:
                info = build_pdf(src, outd / f"{stem}_RU.pdf", d["segs"], d["res"], progress=lambda x: tick(0.4 * x),
                                 shots_out=shots_pdf)
                try:
                    info["docx"] = docx_src.build(shots_pdf, outd / f"{stem}_RU.docx", d["segs"], d["res"], d["docx"])
                except Exception as e:                  # noqa: BLE001 — DOCX прежним путём, причина — в отчёт сборки
                    info["docx"] = docx_from_pdf(outd / f"{stem}_RU.pdf", outd / f"{stem}_RU.docx",
                                                 progress=lambda x: tick(0.4 + 0.6 * x))
                    info["docx"]["fallback"] = f"{type(e).__name__}: {e}"[:200]
            finally:
                shots_pdf.unlink(missing_ok=True)
        else:
            info = build_pdf(src, outd / f"{stem}_RU.pdf", d["segs"], d["res"], progress=lambda x: tick(0.4 * x))
            info["docx"] = docx_from_pdf(outd / f"{stem}_RU.pdf", outd / f"{stem}_RU.docx",
                                         progress=lambda x: tick(0.4 + 0.6 * x))
    review_xlsx(outd / f"{stem}_вычитка.xlsx", d["segs"], d["res"], stem, polish=d.get("polish"), extra_rows=rows)
    return info
