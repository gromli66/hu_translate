"""E9b-S: путь А колонтитулов и у размеченного PDF (S) — номер листа оригинала по счётчику подписи, серии, метка."""
import fitz

import docx_struct as D


def _tagged_like(path, pages, amend=()):
    """PDF без разметки знаков (все знаки — вне дерева, как артефакты размеченного): шапка вверху, абзац тела, подвал
    «Módosítás sorszáma: [k.] N. oldal» внизу (ÜFK I: номер изменения заполнен на части листов). → блоки тела."""
    d = fitz.open()
    blocks = []
    for pno in range(pages):
        pg = d.new_page(width=595.3, height=841.9)
        pg.insert_text((70, 40), "Kezelesi utasitas", fontsize=9)
        pg.insert_text((70, 160), f"Torzs szoveg {pno}", fontsize=11)
        k = "2. " if pno + 1 in amend else ""
        pg.insert_text((70, 800), f"Modositas sorszama: {k}{pno + 1}. oldal", fontsize=9)
        blocks.append({"k": "p", "runs": [{"t": f"Torzs szoveg {pno}"}],
                       "geom": {"page": pno + 1, "top": 150, "bottom": 162, "x0": 70, "x1": 200}})
    d.save(str(path))
    return blocks


def test_tagged_pdf_gets_sheet_numbers_by_counter(tmp_path):
    """Размеченный PDF (без content) — путь А, как у текстового слоя: номер листа оригинала из счётчика подвала.
    Прежде у S путь А не строился (sheet_no не было): нумерация шла подряд по страницам перевода."""
    src = tmp_path / "s.pdf"
    blocks = _tagged_like(src, 7)
    L = D.layout(str(src), blocks)
    assert L.get("sheet_no") == {pn: pn for pn in range(1, 8)}


def test_counter_with_filled_field_before_it_takes_sheet_number(tmp_path):
    """Номер изменения «2.» перед «N. oldal» на листах 4–6 из 12 (ÜFK I: л. 14–16, 34, 64, 65 из 157): номер листа —
    число, что продолжает ряд соседей, а не первое по подписи; метка поля PAGE — на нём, «2.» — текстом в подвале серии.
    Ряд, которого нет по первым числам (изменение на 3 листах из 7), подбором не создаётся — номеров нет, как прежде."""
    src = tmp_path / "amend.pdf"
    blocks = _tagged_like(src, 12, amend={4, 5, 6})
    L = D.layout(str(src), blocks)
    assert L.get("sheet_no") == {pn: pn for pn in range(1, 13)}
    v = [z for z in L["variants"] if z["from"] <= 4 <= z["to"]]
    assert v, "листы с номером изменения — своя серия"
    stamp = " ".join(c for r in v[0]["stamp"] for c in r["cells"])
    assert stamp.replace(" ", "") == "Modositassorszama:2." + D.PAGE_MARK + ".oldal"
    dense = tmp_path / "dense.pdf"
    assert not D.layout(str(dense), _tagged_like(dense, 7, amend={3, 4, 5})).get("sheet_no")


def test_foot_top_border_is_nearest_rule_not_body_table_bottom():
    """Верх подвала — по ближайшей линии над его строкой (и вплотную к ней), а не по любой линии ниже тела: низ таблицы
    тела в 12 pt над подвалом давал штампу лишнюю пустую строку (JPDS, 57 листов; E9b-S Z10)."""
    import docx_struct as D
    foot = {"text": "Д11", "top": 780.0, "bottom": 790.0, "x0": 70.0, "x1": 100.0}
    rule = lambda y: {"top": y, "bottom": y + 0.5, "x0": 70.0, "x1": 560.0}
    assert D._foot_top([foot], [rule(763.0), rule(775.0)], 842.0, 752.0) == 774
    assert D._foot_top([foot], [rule(763.0)], 842.0, 752.0) == 779          # линии рядом нет — по строке
    assert D._foot_top([foot], [rule(774.0), rule(775.5)], 842.0, 752.0) == 773   # двойная линия — целиком


def test_ocr_without_path_a_marks_old_page_number(tmp_path, monkeypatch):
    """Распознанный документ, где путь А не сработал (счётчик решателя не найден): номер листа прежнего поиска («число,
    растущее с листом») — меткой на его цифрах, поле PAGE — по ней (E9b-S Z11: ветки «первое такое число перевода»
    в docx_build._zone_text больше нет — без метки номер застыл бы числом образца)."""
    import docx_layout as DL
    import zones as ZN
    from test_struct_docx import ocr_doc
    monkeypatch.setattr(ZN, "sheet_numbers", lambda zsegs: ({}, set()))
    src, segs = ocr_doc(tmp_path)
    blocks, info = DL.load(src, segs)
    L = DL.layout(src, blocks, segs, info)
    assert L["page_tok"] and L.get("page_marked") and not L.get("sheet_no")
    assert any(D.PAGE_MARK in c for r in L["head_grid"]["rows"] for c in r["cells"])


def test_unmarked_number_in_zone_is_text_not_page_field(tmp_path):
    """Поле PAGE — только по метке решателя: число листа без метки в колонтитуле — текст (прежде — «первое такое число
    перевода» становилось полем, E9b-S Z11)."""
    from docx import Document
    import docx_build as B
    from test_struct_docx import LAYOUT, PROFILE, png
    blocks = [{"k": "p", "runs": [{"t": "Bevezetés"}], "geom": {"page": 1, "top": 100, "bottom": 110, "x0": 70, "x1": 200}}]
    mem = B.Memory([("Bevezetés", "Введение"), ("85. oldal", "Страница 85")])
    for cell, marked, field in (("85. oldal", False, False), (D.PAGE_MARK * 2 + ". oldal", True, True)):
        out = tmp_path / f"z{int(marked)}.docx"
        L = dict(LAYOUT, stamp=[{"cells": [cell], "bold": False}], stamp_widths=[400], stamp_aligns=["right"],
                 page_tok="85.", page_marked=marked, first_is_body=True)
        B.Builder(mem, lambda img: png(), PROFILE).build(blocks, str(out), L)
        xml = Document(str(out)).sections[0].footer._element.xml
        assert ("PAGE" in xml) == field and "85" in xml


def test_fit_layout_takes_section_mark_not_its_toc_copy(tmp_path):
    """docx_fit._layout: LibreOffice при печати обновляет оглавление, и заголовок с меткой раздела (1 pt) попадает в
    него кеглем оглавления. Началом раздела бралась страница оглавления — у ÜFK I 151 раздел «на стр. 2–3», подгонка не
    видела ужатия и откатывала его (листов 162 → 175). Метка ≤ 2 pt важнее копии; без мелкой — любая, как прежде."""
    import docx_fit as F
    pdf = tmp_path / "m.pdf"
    d = fitz.open()
    pg = d.new_page()                                   # л. 1: раздел 0 и оглавление с копиями меток 1 и 2
    pg.insert_text((72, 100), F.MARK.format(0), fontsize=1)
    pg.insert_text((72, 200), F.MARK.format(1) + " 1. Fejezet ....... 3", fontsize=11)
    pg.insert_text((72, 230), F.MARK.format(2) + " 2. Fejezet ....... 4", fontsize=11)
    d.new_page().insert_text((72, 100), "folytatás", fontsize=11)
    d.new_page().insert_text((72, 100), F.MARK.format(1) + " 1. Fejezet", fontsize=1)
    d.new_page().insert_text((72, 100), F.MARK.format(2) + " 2. Fejezet", fontsize=1)
    d.new_page().insert_text((72, 100), "folytatás", fontsize=11)
    pg = d.new_page()                                   # раздел 3 — метка только обычным кеглем (поддельная печать)
    pg.insert_text((72, 100), F.MARK.format(3), fontsize=10)
    d.save(str(pdf))
    L = F._layout(pdf)
    assert {i: v[:2] for i, v in L.items()} == {0: (0, 2), 1: (2, 1), 2: (3, 2), 3: (5, 1)}
