# Дефекты сборки DOCX по просмотру глазами 06.11 (EYES_2026-11-06.md, E9_DESIGN.md §5): один дефект — один тест.
import fitz
from docx import Document
from docx.oxml.ns import qn

import docx_build as B
import docx_layout as DL
import docx_struct as D

from test_assembly_sheets import _sheets_pdf
from test_struct_docx import LAYOUT, PROFILE, png

KEYS = ("head_grid", "variants", "page_marked", "page_tok", "stamp", "stamp_widths", "stamp_box", "stamp_inner",
        "stamp_paras", "stamp_aligns", "stamp_merge", "head_rule", "head_imgs", "head_lines", "first_is_body",
        "sheet_no", "num_start")


def _build_sheets(tmp_path, pages=3, sizes=None, name="sheets", **kw):
    src = tmp_path / f"{name}.pdf"
    segs = _sheets_pdf(src, pages=pages, **kw)
    blocks, info = DL.load(str(src), segs, text=True)
    L = D.layout(str(src), blocks, content=info.get("content_boxes"))
    blocks = D.split_at_zones(blocks, L.get("variants"), D.sheet_closed(L, blocks))
    out = tmp_path / f"{name}.docx"
    bld = B.Builder(B.Memory([(s["text"], s["text"]) for s in segs]), lambda img: png(), PROFILE)
    lay = dict(LAYOUT, **{k: v for k, v in L.items() if k in KEYS})
    if sizes:
        lay["page_sizes"] = sizes
    bld.build(blocks, str(out), lay)
    return bld, Document(str(out)), L


def _text(el):
    return "".join(t.text or "" for t in el.iter(qn("w:t")))


def test_section_break_after_table_takes_no_room(tmp_path):
    """Разрыв раздела после таблицы (лист изменений 30THHAGT, каждый лист пути А): python-docx ставит разрыв в новый
    пустой абзац кеглем тела, а после таблицы уже стоит пустой абзац Word. Два пустых абзаца не влезали под таблицу,
    кончавшуюся у низа листа (722 pt при нижнем поле 733), и уходили на свой лист — пустая стр. 2 (просмотр 06.11).
    Теперь разрыв несёт сам абзац после таблицы, в 1 pt; а там, где таблицы нет, — абзац разрыва в 1 pt."""
    _, doc, _ = _build_sheets(tmp_path, pages=3)
    body = doc.element.body
    brks = [p for p in body.iterchildren(qn("w:p")) if p.find(qn("w:pPr") + "/" + qn("w:sectPr")) is not None]
    assert len(brks) >= 2
    for p in brks:
        assert not _text(p).strip()
        sp = p.find(qn("w:pPr") + "/" + qn("w:spacing"))
        assert sp is not None and sp.get(qn("w:lineRule")) == "exact" and int(sp.get(qn("w:line"))) <= 20
        assert int(sp.get(qn("w:before")) or 0) == 0 and int(sp.get(qn("w:after")) or 0) == 0
        prev = p.getprevious()
        # перед разрывом — не ещё один пустой абзац (тот, что Word ставит после таблицы)
        assert not (prev is not None and prev.tag == qn("w:p") and not _text(prev).strip()
                    and prev.find(".//" + qn("w:drawing")) is None)



def _sec_head(sec):
    return "".join(t.text for t in sec.header._element.iter(qn("w:t")))


def test_last_sheet_before_blank_tail_page_is_one_sheet(tmp_path):
    """Последний лист перед пустой страницей в конце оригинала (30THHAGT: стр. 19 пустая, без номера): граница
    последнего раздела — «листов в документе + 1», и пустой хвост считался листом раздела — раздел выходил «в два
    листа»: без шапки «N (прод.)» и без ужатия docx_fit; перелив последнего листа печатался с номером «17» (просмотр
    06.11). Пустой хвост теперь не лист раздела: раздел в один лист, как у остальных."""
    bld, doc, L = _build_sheets(tmp_path, pages=3, sizes=[(595, 842)] * 4)     # 4-я страница оригинала — пустая
    assert L["sheet_no"] == {1: 1, 2: 2, 3: 3}
    last = len(doc.sections) - 1
    assert bld.fit_want.get(last) == 1
    sec = doc.sections[last]
    assert sec.different_first_page_header_footer and "(прод.)" in _sec_head(sec)


def _frame_doc(tmp_path, name, pages):
    """Лист с рамкой шапки (линии 35/68/82/110) и сегментами pages[pno] = [(текст, x0, y0, x1, align, role, bold)]."""
    doc = fitz.open()
    segs = []
    for pno, lines in enumerate(pages):
        page = doc.new_page(width=595, height=842)
        sh = page.new_shape()
        for y in (35, 68, 82, 110):
            sh.draw_line((70, y), (559, y))
        sh.finish(color=(0, 0, 0), width=0.6)
        sh.commit()
        for text, x0, y0, x1, align, role, bold in lines:
            segs.append({"id": len(segs), "page": pno, "text": text,
                         "loc": {"bbox": [x0, y0, x1, y0 + 9], "frags": [[x0, y0, x1, y0 + 9]], "size": 9.7,
                                 "bold": bold, "align": align, "family": "serif", "ocr": False, "roles": True,
                                 "role": role}})
    path = tmp_path / f"{name}.pdf"
    doc.save(path)
    return str(path), segs


def _texts(blocks):
    return [(b["geom"]["page"], B.text_of(b["runs"]).strip()) for b in B._paras(blocks)]


def test_heading_just_under_header_frame_with_other_code_is_kept(tmp_path):
    """30RR2126GT стр. 10–12 (E9_DESIGN §5, F-1): заголовки «A 30RR24S002 / 25S002 / 26S002 ellenőrzése» стоят сразу
    под рамкой шапки, верх строки на 1,2 pt выше её низа (91,4 при рамке 92,6), — допуск «под рамкой» был по верху
    строки, 1 pt: строка считалась «в рамке», и бегущее название (тот же текст вверху прошлого листа, цифры — маской)
    снимало два заголовка из трёх. «Под рамкой» — по середине строки."""
    head = [("Azonosító: 30RR2126GT", 77, 40, 220, "left", None, False),
            ("Oldalszám: {n}", 460, 57, 538, "left", None, False)]
    pages = []
    codes = {3: "30RR24S002", 4: "30RR25S002", 5: "30RR26S002"}        # три листа из восьми — не колонтитул
    for k in range(8):
        lines = [(t.format(n=k + 7), *rest) for t, *rest in head]
        if k in codes:
            lines.append((f"A {codes[k]} ellenőrzése", 90, 108.8, 300, "left", None, True))
        lines.append((f"A {k + 1}. lépés: a szelepet nyitni, a mérést elvégezni és rögzíteni.", 90, 130, 480,
                      "left", None, False))
        pages.append(lines)
    src, segs = _frame_doc(tmp_path, "rr", pages)
    blocks, _ = DL.load(src, segs, text=True)
    texts = [t for _, t in _texts(blocks)]
    for code in ("30RR24S002", "30RR25S002", "30RR26S002"):
        assert f"A {code} ellenőrzése" in texts



def test_top_line_repeated_lower_is_dropped_only_before_its_heading(tmp_path):
    """L33 (E9_DESIGN §5, F-2): строка вверху листа, чей текст ниже на том же листе повторён, снималась при любом
    повторе — «Sziasztok!» письма 31/32SDTÖMGT, «Schnepp Péter» (30RYBSZGT), «BLOKKINDÍTÁSKOR» (титул 30TKBSZ36BIGT)
    молча пропадали и писались «колонтитулом». Правило писалось под 3PR42: название раздела в рамке сверху и тот же
    текст ниже заголовком. Теперь снимается только перед заголовком (полужирный, по центру или метка «заголовок»)."""
    pages = [[("Sziasztok!", 72, 60, 140, "left", None, False),
              ("A mellékelt utasítást átnéztük, észrevételünk nincs.", 72, 90, 420, "left", None, False),
              ("Sziasztok!", 72, 300, 140, "left", None, False),
              ("Schnepp Péter", 72, 330, 170, "left", None, False)],
             [("1.3/1. SZBV kazetta leesése", 160, 75, 400, "center", None, False),   # не на высоте «Sziasztok!»
              ("1.3/1. SZBV kazetta leesése", 72, 200, 300, "left", "heading", True),
              ("A kazetta leesése esetén a reaktort leállítani.", 72, 230, 400, "left", None, False)]]
    doc = fitz.open()
    segs = []
    for pno, lines in enumerate(pages):
        doc.new_page(width=595, height=842)
        for text, x0, y0, x1, align, role, bold in lines:
            segs.append({"id": len(segs), "page": pno, "text": text,
                         "loc": {"bbox": [x0, y0, x1, y0 + 9], "frags": [[x0, y0, x1, y0 + 9]], "size": 9.7,
                                 "bold": bold, "align": align, "family": "serif", "ocr": False, "roles": True,
                                 "role": role}})
    doc.save(tmp_path / "l33.pdf")
    blocks, _ = DL.load(str(tmp_path / "l33.pdf"), segs, text=True)
    texts = _texts(blocks)
    assert sum(t == "Sziasztok!" for _, t in texts) == 2                 # обычный повтор ниже — оба на месте
    assert sum(t == "1.3/1. SZBV kazetta leesése" for _, t in texts) == 1  # название в рамке перед заголовком — прочь


def _printed_pdf(path, pages=7, ver=None):
    """Листы бланка с версией в шапке (серия листов со своей версией) и подвалом печати «Nyomtatva: 2026-06-01 08:28»
    и номером листа «N/46» правее на той же строке (30RL2126GT, 30RR0102ÁTJGT); номера листа в шапке нет."""
    from test_text_forms import _segs
    ver = ver or {}
    doc = fitz.open()
    for pno in range(pages):
        pg = doc.new_page(width=595, height=842)
        sh = pg.new_shape()
        sh.draw_rect(fitz.Rect(60, 90, 540, 700))
        sh.draw_line((60, 130), (540, 130))
        for x in (140, 440):
            sh.draw_line((x, 130), (x, 700))
        sh.finish(color=(0, 0, 0), width=0.6)
        sh.commit()
        put = lambda x, y, t: pg.insert_text((x, y), t, fontsize=9, fontname="tiro")
        put(62, 40, "Azonosító: 30XYZ")
        put(400, 40, f"Verziószám: {ver.get(pno, 0)}")
        put(62, 56, "Megnevezés: Próba")
        put(64, 124, "Sorsz.")
        put(200, 124, "Művelet, ellenőrzés")
        for k in range(14):
            put(64, 150 + 22 * k, f"{pno + 1}.{k + 1}.")
            put(200, 150 + 22 * k, f"A {pno + 1}. lap {k + 1}. lépése megtörtént.")
        put(62, 800, "Nyomtatva: 2026-06-01 08:28")
        put(298, 800, f" {pno + 2}/46")
    doc.save(str(path))
    return _segs(path)


def _zone_texts(z):
    rows = (z.get("stamp_paras") or {}).get("rows") or []
    return [part[0] for r in rows for part in r["parts"]] + [c for r in z.get("stamp") or [] for c in r["cells"]]


def test_series_counter_is_a_word_not_digits_glued_to_print_time(tmp_path):
    """Счётчик листа серии (раздел листов со своей версией, zones(own=True)) метился по знакам подряд: «08:28» и
    номер «2/46» правее на той же строке склеивались в «08:282/46», число «282» — счётчик, и поле PAGE съедало минуты:
    в подвале «Nyomtatva: 2026-06-01 08:» + PAGE + PAGE + «/46» (30RL2126GT, 30RR0102ÁTJGT; E9_DESIGN §5, F-3).
    Счётчик — по словам: минуты остаются текстом, полем — только номер листа."""
    src = tmp_path / "printed.pdf"
    segs = _printed_pdf(src, ver={3: 1, 4: 1})
    blocks, info = DL.load(str(src), segs, text=True)
    L = D.layout(str(src), blocks, content=info.get("content_boxes"))
    assert [(v["from"], v["to"]) for v in L["variants"]] == [(4, 5)]
    for z in [L] + L["variants"]:
        texts = _zone_texts(z)
        assert any(t.endswith("08:28") for t in texts), texts           # минуты — текстом
        marked = [t for t in texts if D.PAGE_MARK in t]
        assert marked and all(t == D.PAGE_MARK + "/46" for t in marked), texts   # полем — только номер листа


def test_number_slash_number_is_a_field_only_for_the_marked_counter():
    """«число/число» в колонтитуле становилось PAGE/NUMPAGES всегда (штамп UTS): у титулов 30RL2126GT и 30RR0102ÁTJGT
    «Nyomtatva: … 08:28 1/1» на печати выходило «1/<листов перевода>» (E9_DESIGN §5, F-4). Теперь полем — только если в
    документе счётчик листа помечен (тогда номер приходит меткой); где счётчика нет вовсе — как прежде (UTS)."""
    from docx import Document as Doc
    bld = B.Builder(B.Memory([]), lambda img: png(), PROFILE)
    p = Doc().add_paragraph()
    bld._zone_text(p, "Nyomtatva: 2026-06-01 08:28 1/1", False, "6", page_marked=True)
    xml = p._p.xml
    assert "NUMPAGES" not in xml and "PAGE" not in xml and p.text.endswith("1/1")
    p2 = Doc().add_paragraph()
    bld._zone_text(p2, "12/40", False, None, page_marked=False)                 # счётчика в документе нет — штамп UTS
    assert "NUMPAGES" in p2._p.xml and " PAGE" in p2._p.xml


def test_sheet_after_whole_page_picture_starts_on_new_page(tmp_path):
    """3PR10 стр. 3 (NIGHT2_EYES п. 6): лист 2 (оглавление) не открывает раздела — он не закрыт подписью, — и его
    первые строки («Azonosító: 3PR10», «Megnevezés: …») текли прямо под картинку скана титула и печатались на листе
    скана. Картинка во весь лист и есть лист оригинала: следующий лист — с новой страницы."""
    from test_struct_docx import p as para
    fig = {"k": "p", "figure": True, "runs": [{"img": {"png": png(), "x0": 40, "x1": 555, "top": 30, "bottom": 800}}],
           "geom": {"page": 2, "top": 30, "bottom": 800, "x0": 40, "x1": 555, "first_x0": 40}}
    small = {"k": "p", "figure": True, "runs": [{"img": {"png": png(), "x0": 62, "x1": 200, "top": 500, "bottom": 560}}],
             "geom": {"page": 3, "top": 500, "bottom": 560, "x0": 62, "x1": 200, "first_x0": 62}}
    blocks = [para("Lapszám: 1", 1, 100), fig, para("Azonosító: 3PR10", 3, 44), para("TARTALOMJEGYZÉK", 3, 125), small,
              para("A rajz alatt a szöveg folytatódik.", 4, 90)]
    out = tmp_path / "t.docx"
    B.Builder(B.Memory([]), lambda img: png(), PROFILE).build(blocks, str(out), LAYOUT)
    doc = Document(str(out))
    pb = {x.text: bool(x.paragraph_format.page_break_before) for x in doc.paragraphs if x.text}
    assert pb["Azonosító: 3PR10"]                                   # лист после скана — с новой страницы
    assert not pb["TARTALOMJEGYZÉK"] and not pb["A rajz alatt a szöveg folytatódik."]   # после мелкой картинки — поток


def test_sub_line_with_own_mark_keeps_indent_under_its_item():
    """30HU675GT стр. 33 (п. 9.4.4, 9.4.7), 30HU375GT стр. 12 (п. 4.1.4): строки под пунктом — «30TE61S201 tömören zár»,
    «30TE32S001 visszacsapó szelep nyitva», каждая со своей отметкой △ в столбце отметок, — в оригинале с отступом от
    текста пункта (+20…35 pt), в DOCX шли с края текста столбца: вложенный перечень знали только после строки с
    двоеточием («1-es hurok:»). Подпункт с кода оборудования, со своей отметкой в ряду, правее строки над ним в
    столбце, — с отступом от неё. Строка без своей отметки — как прежде."""
    import tempfile, os
    g = lambda x0, top: {"page": 1, "top": top, "bottom": top + 10, "x0": x0, "x1": x0 + 200, "first_x0": x0, "size": 10,
                         "lines": 1}
    run = lambda t: {"t": t, "font": "Times New Roman", "size": 10, "b": False, "i": False, "color": "000000"}
    para = lambda t, x0, top: {"k": "p", "runs": [run(t)], "geom": g(x0, top), "seg": True, "align": "left"}
    edges = {"top": True, "bottom": True, "left": True, "right": True}
    cell = lambda ps: {"blocks": ps, "th": False, "geom": ps[0]["geom"] if ps else None, "span": 1, "edges": edges}
    rows = [[cell([para("9.4.4.", 84, 290)]),
             cell([para("5 perc elteltével ellenőrizni, hogy a nyomáscsökkenés kisebb, mint 5,4 bar", 128, 290)]),
             cell([para("△", 480, 290)])],
            [cell([]), cell([para("30TE61S201 tömören zár", 163, 304)]), cell([para("△", 480, 304)])],
            [cell([]), cell([para("30TE62S201 tömören zár", 163, 316)]), cell([para("\uf072", 480, 316)])],   # △ Wingdings
            [cell([]), cell([para("30TE63S201 nyitva", 150, 330)]), cell([])]]           # без своей отметки
    tab = {"k": "table", "rows": rows, "borders": True, "frame": {"xs": [71, 120, 453, 520], "ys": [288, 302, 314, 328, 342]},
           "geom": {"page": 1, "top": 288, "bottom": 342, "x0": 71, "x1": 520, "first_x0": 71, "size": 10}}
    out = os.path.join(tempfile.mkdtemp(), "sub.docx")
    B.Builder(B.Memory([]), lambda img: png(), PROFILE).build([tab], out, dict(LAYOUT, first_is_body=True))
    t = Document(out).tables[0]
    ind = {p.text: p.paragraph_format.left_indent for c in t._cells for p in c.paragraphs if p.text}
    assert ind["30TE61S201 tömören zár"] is not None and round(ind["30TE61S201 tömören zár"].pt) == 35
    assert round(ind["30TE62S201 tömören zár"].pt) == 35
    assert ind["30TE63S201 nyitva"] is None


def test_minus_sign_is_list_marker_for_memory_and_text_layer():
    """«−» (U+2212, минус символьного шрифта Word) — маркер пунктов TU пп. 1.16–1.17 (SEG_QUALITY_REPORT, «Нужно от
    сборки» п. 1). Память знала метки «-», «–», «•», но не «−»: дерево разметки, отдающее метку отдельно, спрашивало
    «30YP15S101 zárva (…)» и не находило; перевод находился с «−» текстом — на печати двойной маркер «• − 30YP15S101 …».
    И текстовый слой (docx_layout.SIGN) не узнавал «−» маркером. «−5 °C» (минус числа) — не маркер."""
    import docx_layout as DLx
    m = B.Memory([("− 30YP15S101 zárva (lámpajelzés)", "− 30YP15S101 закрыта (световая сигнализация)")])
    assert m("30YP15S101 zárva (lámpajelzés)") == "30YP15S101 закрыта (световая сигнализация)"
    assert DLx.SIGN.match("− 30YP15S101 zárva") and not DLx.SIGN.match("−5 °C")


def test_blank_line_between_paragraphs_of_whole_sheet_cell_is_kept():
    """01-02FMP0BIGT л. 2 (NIGHT2_EYES п. 15, просмотр 06.11 п. 11): лист бланка — одна большая ячейка, между абзацами —
    пустая строка (зазор 9–15 pt при обычном −2…+4), в DOCX её нет — лист плотнее оригинала. Зазоры всех ячеек раньше
    сознательно не переносили (30RL2126GT: 449 границ из 738 — 49 листов при 46): теперь пустая строка — только в ячейке
    во весь лист (абзацы ячейки — больше половины высоты листа); в обычной ячейке — как прежде."""
    import tempfile, os
    def para(t, top, lines=1):
        g = {"page": 1, "top": top, "bottom": top + 13.8 * lines, "x0": 136, "x1": 460, "first_x0": 136, "size": 10,
             "lines": lines}
        return {"k": "p", "runs": [{"t": t, "font": "Times New Roman", "size": 10, "b": False, "i": False,
                                    "color": "000000"}], "geom": g, "seg": True, "align": "left"}
    edges = {"top": True, "bottom": True, "left": True, "right": True}
    big = [para("A próba terjedelme:", 140.6), para("4.1 - MET-re hozás.", 163.6), para("4.3 - Az SZBV hajtások.", 175.1),
           para("A program érvényessége:", 234.2), para("A programban leírt információk érvényesek.", 257.3, 2),
           para("A mérési program ütemezése:", 578.7), para("A zóna ellenőrző rendszer hőmérői.", 601.9, 3)]
    small = [para("Első sor a kis cellában.", 120), para("Második sor a kis cellában.", 145)]
    cell = lambda ps: {"blocks": ps, "th": False, "geom": ps[0]["geom"], "span": 1, "edges": edges}
    rows = [[cell(big), cell(small)]]
    tab = {"k": "table", "rows": rows, "borders": True, "frame": {"xs": [71, 470, 540], "ys": [100, 700]},
           "geom": {"page": 1, "top": 100, "bottom": 700, "x0": 71, "x1": 540, "first_x0": 71, "size": 10}}
    out = os.path.join(tempfile.mkdtemp(), "gap.docx")
    B.Builder(B.Memory([]), lambda img: png(), PROFILE).build([tab], out, dict(LAYOUT, first_is_body=True))
    t = Document(out).tables[0]
    sb = {p.text: p.paragraph_format.space_before for c in t._cells for p in c.paragraphs if p.text}
    assert sb["4.1 - MET-re hozás."] is not None and round(sb["4.1 - MET-re hozás."].pt) == 9      # пустая строка
    assert not sb["4.3 - Az SZBV hajtások."]                                     # строка за строкой — без зазора
    assert round(sb["A mérési program ütemezése:"].pt) >= 9
    assert not sb["Második sor a kis cellában."]                                 # обычная ячейка — как прежде


def _nested_pdf(path, grey=True):
    """Лист: рамка в 3 столбца (Sorsz. | Feltétel | Igazolás) во весь лист, в средней ячейке — бланк 4 × 4 с серыми
    клетками (заливка ячейки и внутри неё — заливка абзаца тем же цветом, как рисует Word) — 30TH50-80ÁTVGT л. 6–9."""
    from test_text_forms import _segs
    doc = fitz.open()
    pg = doc.new_page(width=595, height=842)
    sh = pg.new_shape()
    sh.draw_rect(fitz.Rect(65, 78, 570, 753))
    for x in (113, 520):
        sh.draw_line((x, 78), (x, 753))
    sh.draw_line((65, 100), (570, 100))
    xs, ys = [150, 230, 310, 390, 470], [360, 380, 400, 420, 440]
    for x in xs:
        sh.draw_line((x, ys[0]), (x, ys[-1]))
    for y in ys:
        sh.draw_line((xs[0], y), (xs[-1], y))
    sh.finish(color=(0, 0, 0), width=0.6)
    sh.commit()
    if grey:
        for (x0, y0) in ((230, 400), (310, 400), (390, 400), (230, 420)):
            pg.draw_rect(fitz.Rect(x0 + 0.5, y0 + 0.5, x0 + 79.5, y0 + 19.5), color=None, fill=(0.85, 0.85, 0.85))
            pg.draw_rect(fitz.Rect(x0 + 5, y0 + 6, x0 + 74, y0 + 14), color=None, fill=(0.85, 0.85, 0.85))
    put = lambda x, y, t: pg.insert_text((x, y), t, fontsize=9, fontname="tiro")
    put(70, 94, "Sorsz.")
    put(250, 94, "Feltétel, művelet, ellenőrzés")
    put(525, 94, "Igazolás")
    put(70, 130, "4.4.")
    put(120, 130, "A kimérés az adatok felvételéből áll, a program vezetője rögzíti.")
    for i, row in enumerate((("Mérés", "óra/perc", "L101 (mm)", "L102 (mm)"), ("0 perc", "...", "", ""),
                             ("20 perc", "...", "", ""), ("40 perc", "...", "", ""))):
        for j, t in enumerate(row):
            if t:
                put(xs[j] + 4, ys[i] + 13, t)
    put(120, 480, "Amennyiben a szint csökkenése eléri a 2,5 mm-t, a nyomást be kell állítani.")
    doc.save(str(path))
    return _segs(path)


def test_paragraph_shading_inside_grey_cell_and_nested_table_fills_are_not_chart_bars():
    """Таблица-график — от трёх залитых фигур не по сетке. Word рисует в серой ячейке ещё и заливку абзаца того же цвета
    (не по сетке) — бланк «Время измерения» сам был «графиком»; а серые клетки вложенного бланка считались столбиками
    внешней рамки листа — лист 30TH50-80ÁTVGT уходил снимком (E7 §5, просмотр 06.11 п. 12)."""
    import figures as FIG
    xs, ys = [150, 230, 310, 390, 470], [360, 380, 400, 420, 440]
    inner = (xs, ys, [], [])
    cells = [((x0 + 0.5, y0 + 0.5, x0 + 79.5, y0 + 19.5), (x0 + 5, y0 + 6, x0 + 74, y0 + 14))
             for x0, y0 in ((230, 400), (310, 400), (390, 400), (230, 420))]
    draw = [{"rect": r, "solid": True, "fill": (0.85, 0.85, 0.85)} for c in cells for r in c]
    assert not FIG.chart_table(inner, draw)                         # заливка абзаца в ячейке — не столбик
    other = [{"rect": c[0], "solid": True, "fill": (0.85, 0.85, 0.85)} for c in cells] + \
        [{"rect": c[1], "solid": True, "fill": (0.2, 0.4, 0.8)} for c in cells]
    assert FIG.chart_table(inner, other)                            # полоса другого цвета в ячейке — столбик (Гант)
    outer = ([65, 113, 520, 570], [78, 100, 753], [], [])
    assert FIG.chart_table(outer, draw)                             # для внешней клетки вложенной — «столбики»…
    assert not FIG.chart_table(outer, draw, [fitz.Rect(150, 360, 470, 440)])   # …но не своей таблицы


def test_nested_table_in_cell_is_a_table_not_a_sheet_snapshot(tmp_path):
    """Бланк «Время измерения» в ячейке рамки листа (30TH50-80ÁTVGT л. 6–9): лист уходил снимком (картинка — не
    текст), а откат E7 показал, что абзацами внешней ячейки бланк разваливался в строки. Теперь вложенная таблица —
    своим блоком по своей сетке в ячейке внешней, на своём месте по высоте, её текст в ячейке внешней не повторяется."""
    src = tmp_path / "nested.pdf"
    segs = _nested_pdf(src)
    blocks, _ = DL.load(str(src), segs, text=True)
    assert not [b for b in blocks if b.get("figure")]
    outer = [b for b in blocks if b["k"] == "table" and any(x["k"] == "table" for row in b["rows"] for c in row
                                                              for x in c["blocks"])]
    assert len(outer) == 1
    cell = next(c for row in outer[0]["rows"] for c in row if any(x["k"] == "table" for x in c["blocks"]))
    kinds = [(x["k"], B.text_of(x.get("runs") or [])[:12]) for x in cell["blocks"]]
    assert [k for k, _ in kinds] == ["p", "table", "p"], kinds      # текст над бланком, бланк, текст под ним
    sub = next(x for x in cell["blocks"] if x["k"] == "table")
    texts = [" ".join(B.text_of(x["runs"]) for x in c["blocks"]) for row in sub["rows"] for c in row]
    assert "20 perc" in texts and "L101 (mm)" in texts
    assert not any("20 perc" in B.text_of(x.get("runs") or []) for x in cell["blocks"] if x["k"] == "p")
