# -*- coding: utf-8 -*-
"""Решатель «колонтитул листа» (zones.decide, E9b-Z) на сегментах без PDF: линии листов и рамки шапки — заданы."""
import zones as ZN

H = 842


def _seg(segs, page, text, y, x0=60, x1=300, **loc):
    segs.append({"id": len(segs), "page": page, "text": text,
                 "loc": dict({"bbox": [x0, y, x1, y + 10], "size": 10, "bold": False, "align": "left"}, **loc)})


def _rules(pages, lines=None):
    """(—, линии листов [(горизонтали, вертикали)], —, —, низ рамки шапки) — как docx_layout._page_rules."""
    return None, [(lines or {}).get(p, ([], [])) for p in range(pages)], None, None, [None] * pages


def test_series_header_fields_leave_body_from_first_sheet_of_series():
    """Серия листов со своей шапкой (3PR10 л. 20–34 «Azonosító: 3PR10/INFO», «verziószám: 0»): строки-поля шапки —
    не тело и на первом листе серии (оставались текстом в начале листа 20, у 1SZ01 «Oldalszám:1» — заголовком);
    бегущее название раздела (полужирное, по центру, не поле) — заголовком один раз, на первом листе."""
    segs = []
    for p in range(6):
        # шапка серии ниже шапки тела (3PR10: 45,8 pt против 34,4) — повтор по высоте её не берёт
        _seg(segs, p, *(("Azonosító: 3PR10", 40) if p < 4 else ("Azonosító: 3PR10/INFO", 52)))
        _seg(segs, p, f"Oldalszám: {p + 1}", 40, x0=400, x1=520)
        if p in (2, 3):
            _seg(segs, p, "1.15. Átfolyás a primerkörből", 80, x0=200, x1=400, bold=True, align="center")
        _seg(segs, p, f"Szöveg a lapon, sor {p}.", 300)
    content, why = ZN.decide(segs, H, 6, _rules(6))
    kept = {s["id"]: s for s in content}
    info = [s for s in segs if s["text"] == "Azonosító: 3PR10/INFO"]
    assert [why.get(s["id"]) for s in info] == ["шапка серии", "повтор с прошлого листа"]
    title = [s for s in segs if s["text"].startswith("1.15.")]
    assert title[0]["id"] in kept and kept[title[0]["id"]]["loc"].get("title") and title[1]["id"] not in kept
    assert all(s["id"] in kept for s in segs if s["text"].startswith("Szöveg"))


def test_repeat_from_previous_sheet_is_not_a_table_cell():
    """Строка вверху листа, чья подпись на прошлом листе стоит в ячейке таблицы с телом (таблица изменений), — тело:
    «Érintett fejezetek:» листа регистрации 30RL2126GT пропадала как «повтор с прошлого листа» (E9b-Z, Z4)."""
    segs = []
    _seg(segs, 0, "Verzió", 104, x1=190)
    _seg(segs, 0, "Érintett fejezetek:", 124, x0=210, x1=400)
    _seg(segs, 0, "Új TU kiadása, alkalmasság ellenőrzésre", 144, x0=210, x1=480)
    _seg(segs, 1, "Érintett fejezetek:", 134, x0=65, x1=180)
    _seg(segs, 1, "Ciklikusan felülvizsgált TU kiadása.", 300)
    xs, ys = [50, 200, 500], [100, 120, 140, 160]
    table = ([(y, 50, 500) for y in ys], [(x, 100, 160) for x in xs])
    content, why = ZN.decide(segs, H, 2, _rules(2, {0: table}))
    assert {s["id"] for s in content} == {s["id"] for s in segs}
    assert 1 in why.heads and 3 not in why.heads


def test_ocr_sheets_numbered_from_header_counter(tmp_path):
    """Путь А и для распознанного документа (E9b-Z, Z8): номер листа оригинала — счётчик шапки («Oldalszám: N»; первый
    лист PDF — «2», титула в файле нет), лист — свой раздел со своим номером, номер в шапке — поле PAGE на месте
    счётчика (метка), а не первое число перевода; лист, где OCR номера не прочёл, — без номера (не угадывается): он
    продолжает раздел прошлого листа."""
    from docx import Document
    from docx.oxml.ns import qn
    import docx_build as B
    import docx_struct as D
    from test_struct_docx import LAYOUT, PROFILE, ocr_doc, png
    src, segs = ocr_doc(tmp_path)
    for s in segs:
        if s["page"] == 4 and s["text"].startswith("Oldalszám"):
            s["text"] = "Oldalszám:"
    blocks, L = B.blocks_layout(src, segs, "layout")
    assert L["sheet_no"] == {p: p + 1 for p in range(1, 9) if p != 5} and L["page_marked"]
    assert any(D.PAGE_MARK in c for r in L["head_grid"]["rows"] for c in r["cells"])
    out = tmp_path / "ocr.docx"
    B.Builder(B.Memory([(s["text"], s["text"]) for s in segs]), lambda img: png(), PROFILE).build(
        blocks, str(out), dict(LAYOUT, **L))
    doc = Document(str(out))
    starts = [(s._sectPr.find(qn("w:pgNumType")).get(qn("w:start")) if s._sectPr.find(qn("w:pgNumType")) is not None
               else None) for s in doc.sections]
    assert starts == ["2", "3", "4", "5", "7", "8", "9"]
    # номер — полем на первой странице раздела-листа; его перелив — «N … (прод.)» текстом (путь А)
    assert all("PAGE" in sec.first_page_header._element.xml for sec in doc.sections)


def test_inserted_ocr_sheet_numbered_as_text(tmp_path):
    """Вставной лист «5/1» (3PR42 «56/1–56/3»): номер не целый, полем PAGE не выразить — раздел листа без начала
    нумерации, номер текстом в шапке (первая страница — «5/1», перелив — «5/1 (прод.)»); следующий лист — своё «7»."""
    from docx import Document
    from docx.oxml.ns import qn
    import docx_build as B
    from test_struct_docx import LAYOUT, PROFILE, ocr_doc, png
    src, segs = ocr_doc(tmp_path)
    for s in segs:
        if s["page"] == 4 and s["text"].startswith("Oldalszám"):
            s["text"] = "Oldalszám: 5/1"
    blocks, L = B.blocks_layout(src, segs, "layout")
    assert L["sheet_no"][5] == "5/1" and L["sheet_no"][6] == 7
    out = tmp_path / "ocr.docx"
    B.Builder(B.Memory([(s["text"], s["text"]) for s in segs]), lambda img: png(), PROFILE).build(
        blocks, str(out), dict(LAYOUT, **L))
    doc = Document(str(out))
    starts = [(s._sectPr.find(qn("w:pgNumType")).get(qn("w:start")) if s._sectPr.find(qn("w:pgNumType")) is not None
               else None) for s in doc.sections]
    assert starts == ["2", "3", "4", "5", None, "7", "8", "9"]
    ins = doc.sections[4]
    assert "5/1" in ins.first_page_header._element.xml and " PAGE " not in ins.first_page_header._element.xml
    assert "5/1 (прод.)" in "".join(p.text for p in ins.header.paragraphs) + "".join(
        c.text for t in ins.header.tables for r in t.rows for c in r.cells)
    assert " PAGE " in doc.sections[5].first_page_header._element.xml


def test_ocr_header_sample_without_merged_lines(tmp_path):
    """Образец шапки OCR — лист без сегмента в две строки, если такие листы есть: OCR склеил «Verziószám: 0» над
    «Oldalszám: N» в один сегмент на большинстве листов (3PR42: 58 из 146), и шапка перевода шла «Номер версии: 0 Номер
    страницы:» одной строкой, а номер — строкой ниже; на листах, где строки прочитаны по отдельности, — как в оригинале."""
    import docx_build as B
    import docx_struct as D
    from test_struct_docx import ocr_doc
    src, segs = ocr_doc(tmp_path)
    sid = max(s["id"] for s in segs) + 1
    out = []
    for s in segs:
        if s["text"].startswith("Azonosító"):     # вторая строка шапки слева — на всех листах (как «Megnevezés» 3PR42)
            out.append(dict(s, id=sid, text="Megnevezés: TURBINA",
                            loc=dict(s["loc"], bbox=[77, 48, 300, 57], frags=[[77, 48, 300, 57]])))
            sid += 1
        if s["text"].startswith("Oldalszám"):
            n = s["text"].split()[-1]
            if s["page"] < 5:                       # склейка двух строк шапки (два фрагмента — две строки)
                s = dict(s, text=f"Verziószám: 0 Oldalszám: {n}",
                         loc=dict(s["loc"], bbox=[460, 35, 530, 57], frags=[[460, 35, 530, 44], [460, 48, 530, 57]]))
            else:
                out.append(dict(s, id=sid, text="Verziószám: 0",
                                loc=dict(s["loc"], bbox=[460, 35, 530, 44], frags=[[460, 35, 530, 44]])))
                sid += 1
        out.append(s)
    blocks, L = B.blocks_layout(src, out, "layout")
    rows = [" ".join(c if isinstance(c, str) else " ".join(t for t, _ in c["lines"]) for c in r["cells"])
            for r in L["head_grid"]["rows"]]
    assert any("Verziószám: 0" in r and "Oldalszám" not in r for r in rows), rows
    assert any("Oldalszám:" in r and D.PAGE_MARK in r for r in rows), rows


def test_header_cell_label_and_number_translated():
    """Ячейка шапки «подпись + число» из двух сегментов («Verziószám:» и «0»: сетка колонтитула приклеивает столбец
    чисел к подписи) — перевод подписи и число; в памяти целиком её нет, и она шла по-венгерски (3PR42)."""
    import docx_build as B
    mem = B.Memory([("Verziószám:", "Номер версии:"), ("0", "0"), ("Azonosító: 3PR42", "Идентификатор: 3PR42")])
    assert mem("Verziószám: 0") is None
    assert mem.zone("Verziószám: 0") == "Номер версии: 0"
    assert mem.zone("Megnevezés 7") is None                  # подпись не из памяти — как было


def test_sheet_numbers_only_read_and_monotonic():
    """Номер листа OCR — только прочитанный и в монотонном ряду: повтор/скачок назад — лист без номера; сбоев много —
    номеров нет у документа (прежнее поведение)."""
    def zs(nums):
        return {p: [{"page": p - 1, "text": f"Oldalszám: {n}" if n is not None else "Oldalszám:",
                     "loc": {"bbox": [400, 40, 500, 50]}}] for p, n in enumerate(nums, 1)}
    no, _ = ZN.sheet_numbers(zs([2, 3, 4, 5, 5, 7, 8, 9, 10, 11]))
    assert no == {1: 2, 2: 3, 3: 4, 6: 7, 7: 8, 8: 9, 9: 10, 10: 11}
    no, _ = ZN.sheet_numbers(zs([2, 3, None, None, 6, None, 8, 9]))
    assert no == {}
    # вставные листы «56/1», «56/2» (3PR42) — номер строкой, если строго между соседями; сосед «56» свой номер не теряет
    z = zs([54, 55, 56, 56, 56, 57, 58, 59])
    for p in (4, 5):
        z[p][0]["text"] = f"Oldalszám: 56/{p - 3}"
    no, _ = ZN.sheet_numbers(z)
    assert no == {1: 54, 2: 55, 3: 56, 4: "56/1", 5: "56/2", 6: 57, 7: 58, 8: 59}
    # OCR прочёл «56/7» вместо «56/2»: ряд «56/1, 56/7, 56/3» не растёт по порядку — у «56/7» и «56/3» номера нет, без
    # угадывания; и у «56/1»: сосед без номера продолжил бы его раздел под «56/1 (прод.)»
    z[5][0]["text"] = "Oldalszám: 56/7"
    z[6][0]["text"] = "Oldalszám: 56/3"
    no, _ = ZN.sheet_numbers(z)
    assert no == {1: 54, 2: 55, 3: 56, 7: 58, 8: 59}


def test_height_key_takes_header_row_not_registry_sheet_text():
    """По высоте строки шапки снимается строка той же шапки (склеенная иначе: «Azonosító: 3PR10 Megnevezés: …»;
    строка-поле «Lapszám: 1»; штамп, что портал пометил колонтитулом), но не текст листа без шапки на той же высоте —
    «NYILVÁNTARTÓ LAP», «2./ A kiadás oka:» листа регистрации пропадали (30PR132GT л. 1, E9b-Z Z1)."""
    segs = []
    _seg(segs, 0, "NYILVÁNTARTÓ LAP", 41, x0=230, x1=360)
    _seg(segs, 0, "2./ A kiadás oka:", 61)
    _seg(segs, 0, "Lapszám: 1", 61, x0=430, x1=520)
    _seg(segs, 0, "ELLENŐRZÖTT PÉLDÁNY", 41, x0=10, x1=90, ocr=True, role="header")
    for p in range(1, 8):
        if p == 3:
            _seg(segs, p, "Azonosító: 3PR10 Megnevezés: Főkeringtető szivattyú", 40, x1=520)
        else:
            _seg(segs, p, "Azonosító: 3PR10", 40)
            _seg(segs, p, "Megnevezés: Főkeringtető szivattyú", 40, x0=320, x1=520)
        _seg(segs, p, f"Oldalszám: {p}", 60, x0=430, x1=520)
        _seg(segs, p, f"Szöveg a lapon, sor {p}.", 300)
    content, why = ZN.decide(segs, H, 8, _rules(8))
    kept = {s["text"] for s in content}
    assert {"NYILVÁNTARTÓ LAP", "2./ A kiadás oka:"} <= kept
    assert not kept & {"Lapszám: 1", "ELLENŐRZÖTT PÉLDÁNY", "Azonosító: 3PR10 Megnevezés: Főkeringtető szivattyú"}


def test_portal_header_field_leaves_body_text_layer_title_stays():
    """Строка-поле шапки, что портал пометил колонтитулом на скане первого листа тела («Verziószám: 2.», «Oldalszám:
    1.» 30HU*75GT), — колонтитул, не «название раздела» (первое появление оставалось в тексте заголовком); строка
    текстового слоя, похожая на метку OCR-листа, — не название раздела («10/6. Melléklet» 1SZ01 под рамкой), E9b-Z Z2."""
    segs = []
    _seg(segs, 0, "Verziószám: 2.", 66, x0=430, x1=505, ocr=True, role="header")
    _seg(segs, 0, "Oldalszám: 1.", 80, x0=430, x1=500, ocr=True, role="header")
    _seg(segs, 0, "10/6. Melléklet", 30, x0=250, x1=360, ocr=True, role="header")
    for p in range(1, 6):
        _seg(segs, p, "Azonosító: 30HU375GT", 40)
        _seg(segs, p, f"Oldalszám: {p + 1}.", 40, x0=430, x1=520)
        _seg(segs, p, f"Szöveg a lapon, sor {p}.", 300)
    _seg(segs, 4, "10/6. Melléklet", 90, x0=250, x1=360, bold=True, align="center")
    content, why = ZN.decide(segs, H, 6, _rules(6))
    kept = {(s["page"], s["text"]) for s in content}
    assert (0, "Verziószám: 2.") not in kept and (0, "Oldalszám: 1.") not in kept
    assert (4, "10/6. Melléklet") in kept


def test_series_header_in_frame_below_body_top_is_header(tmp_path):
    """Шапка серии листов — рамкой ниже верха тела документа, выше неё знаков нет (3FI18 л. 31, 33–39: рамка 57–83 pt
    при верхе тела 53): путь А берёт в шапку серии знаки внутри рамки — прежде шапка была пустой, а её строки после Z4
    сняты из тела и пропадали."""
    import fitz
    import docx_layout as DL
    import docx_struct as D
    from test_text_forms import _segs
    src = tmp_path / "frame.pdf"
    doc = fitz.open()
    words = ("elsődleges tömítettség szivattyú indítás szelep zárás nyomás hőmérséklet teljesítmény áramlás szint "
             "vezérlés védelem mérés kapcsoló betáplálás hűtés fűtés csővezeték").split()
    for pno in range(6):
        pg = doc.new_page(width=595, height=842)
        put = lambda x, y, t: pg.insert_text((x, y), t, fontsize=9, fontname="tiro")
        top = 70
        if pno < 4:
            put(62, 40, "Azonosító: 30XYZ")
            put(400, 40, f"Oldalszám: {pno + 1}")
        else:                                               # серия: шапка в рамке ниже верха тела, тело — ниже
            sh = pg.new_shape()
            sh.draw_line((60, 78), (540, 78))
            sh.draw_line((60, 104), (540, 104))
            sh.finish(color=(0, 0, 0), width=0.6)
            sh.commit()
            put(62, 88, "Azonosító: 30XYZ/IB")
            put(400, 88, "Verziószám: 0")
            put(400, 100, f"Oldalszám: {pno - 3}")
            top = 130
        for k in range(12):
            put(64, top + 22 * k, f"A {words[(pno * 5 + k) % len(words)]} és a {words[(pno * 3 + k * 7 + 2) % len(words)]}"
                                  f" ellenőrzése megtörtént ({words[(pno + k * 11) % len(words)]}).")
    doc.save(str(src))
    segs = _segs(src)
    blocks, info = DL.load(str(src), segs, text=True)
    L = D.layout(str(src), blocks, content=info.get("content_boxes"))
    v = [z for z in L.get("variants") or [] if z["from"] == 5]
    assert v and "30XYZ/IB" in str(v[0].get("head_grid")) + str(v[0].get("head_lines"))
