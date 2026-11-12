# -*- coding: utf-8 -*-
"""Раскладка копии скана по строкам Surya (src/copylayout.py, 12.11): выравнивание и отступ коротких кусков, пункты
списка из одного <p> с переносами — отдельными абзацами с отступом, красная строка длинных абзацев, куски одной строки
скана — одним абзацем через табуляцию; не сопоставилось — как у портала; текст не меняется."""
import math

from docx import Document
from docx.enum.text import WD_TAB_ALIGNMENT

import copycheck as CC
import copylayout as CL

DPI = 200
PT = 72 / DPI
LONG = ("ООО «Авангард-НСК» ИНН 5407471468 информирует Вас о том, что с 01 марта 2023 года реализацию всей продукции "
        "иностранных брендов ввезенных на территорию РФ будет производить дочернее предприятие ООО «АВАНГАРД» "
        "ИНН 5405490514 (организовано с 2014 г.).")


def ln(t, x0, y0, x1, h=30, a=0.0):
    """Строка Surya: рамка и многоугольник (повёрнутый на a радиан вокруг начала — перекос листа)."""
    c, s = math.cos(a), math.sin(a)
    pts = [[x * c - y * s, x * s + y * c] for x, y in ((x0, y0), (x1, y0), (x1, y0 + h), (x0, y0 + h))]
    xs, ys = [p[0] for p in pts], [p[1] for p in pts]
    return {"t": t, "box": [min(xs), min(ys), max(xs), max(ys)], "poly": pts, "c": 0.97}


def page_lines(a=0.0, sign="Сливкин Д.Н."):
    return [ln("Информационное письмо", 700, 100, 1000, a=a),
            ln("ООО «Авангард-НСК» ИНН 5407471468 информирует Вас о том, что с 01 марта", 280, 200, 1500, a=a),
            ln("2023 года реализацию всей продукции иностранных брендов ввезенных на", 200, 240, 1500, a=a),
            ln("территорию РФ будет производить дочернее предприятие ООО «АВАНГАРД»", 200, 280, 1490, a=a),
            ln("ИНН 5405490514 (организовано с 2014 г.).", 200, 320, 800, a=a),
            ln("Данное решение обусловлено рядом причин:", 200, 400, 900, a=a),
            ln("- защита канала поставки от санкционных рисков,", 260, 440, 1100, a=a),
            ln("- разделение направлений бизнеса,", 260, 480, 900, a=a),
            ln("По месту требования", 1100, 560, 1500, a=a),
            ln("(подпись)", 300, 700, 450, a=a),
            ln(sign, 900, 700, 1150, a=a),
            ln("(noint", 200, 800, 330, a=a)]


def make_doc():
    doc = Document()
    doc.add_paragraph("Информационное письмо")
    doc.add_paragraph(LONG)
    p = doc.add_paragraph("Данное решение обусловлено рядом причин:")
    p.add_run().add_break()
    p.add_run(" - защита канала поставки от санкционных рисков,")
    p.add_run().add_break()
    p.add_run(" - разделение направлений бизнеса,")
    doc.add_paragraph("По месту требования").alignment = 1
    doc.add_paragraph("(подпись)").alignment = 1
    doc.add_paragraph("Сливкин Д.Н.").alignment = 1
    doc.add_paragraph("Директор").alignment = 1
    return doc


def words(doc):
    return " ".join(p.text for p in doc.paragraphs).split()


def test_layout_from_scan():
    doc = make_doc()
    before = words(doc)
    laid, linked = CL.apply(doc, {0: page_lines()}, DPI)
    ps = doc.paragraphs
    assert [p.text for p in ps] == ["Информационное письмо", LONG, "Данное решение обусловлено рядом причин:",
                                    "- защита канала поставки от санкционных рисков,", "- разделение направлений бизнеса,",
                                    "По месту требования", "(подпись)\tСливкин Д.Н.", "Директор"]
    assert words(doc) == before                                          # текст скан не меняет
    assert ps[0].alignment == 1                                          # заголовок — по центру
    assert abs(ps[1].paragraph_format.first_line_indent.pt - 80 * PT) < 0.2      # красная строка длинного абзаца
    assert ps[1].alignment is None
    assert not ps[2].paragraph_format.left_indent                         # «Данное решение…» — от поля
    for p in ps[3:5]:                                                     # пункты списка — свои абзацы с отступом
        assert abs(p.paragraph_format.left_indent.pt - 60 * PT) < 0.2
    assert ps[2].paragraph_format.space_after.pt == 0 and ps[3].paragraph_format.space_after.pt == 0
    assert ps[5].alignment == 2
    pf = ps[6].paragraph_format                                           # (подпись) ⇥ ФИО — одной строкой
    assert ps[6].alignment == 0 and abs(pf.left_indent.pt - 100 * PT) < 0.2
    tab = list(pf.tab_stops)
    assert len(tab) == 1 and tab[0].alignment == WD_TAB_ALIGNMENT.LEFT and abs(tab[0].position.pt - 700 * PT) < 0.2
    assert ps[7].alignment == 1                                           # «(noint» не сопоставилось — как у портала
    assert (laid, linked) == (7, 1)


def test_skewed_sheet_same_layout():
    """Лист с перекосом 2° (Декларация о соответствии — 2,5°): строки поворачиваются обратно, раскладка та же."""
    doc = make_doc()
    CL.apply(doc, {0: page_lines(a=math.radians(2))}, DPI)
    ps = doc.paragraphs
    assert ps[0].alignment == 1 and ps[5].alignment == 2
    assert abs(ps[3].paragraph_format.left_indent.pt - 60 * PT) < 1.5
    assert ps[6].text == "(подпись)\tСливкин Д.Н."


def test_right_piece_gets_right_tab():
    doc = make_doc()
    lines = page_lines()
    lines[10] = ln("Сливкин Д.Н.", 1250, 700, 1500)
    CL.apply(doc, {0: lines}, DPI)
    tab = list(doc.paragraphs[6].paragraph_format.tab_stops)
    assert len(tab) == 1 and tab[0].alignment == WD_TAB_ALIGNMENT.RIGHT


def test_half_read_line_left_as_portal():
    """Строка наполовину под печатью: Surya прочла правую половину — начала куска в ней нет, кусок не берётся
    (иначе отступ 354 px, «Регистрационный номер…» Декларации о соответствии)."""
    doc = make_doc()
    lines = page_lines()
    lines[8] = ln("требования", 1300, 560, 1500)
    doc.paragraphs[3].alignment = 1
    CL.apply(doc, {0: lines}, DPI)
    assert doc.paragraphs[5].text == "По месту требования" and doc.paragraphs[5].alignment == 1


def test_no_lines_nothing_changes():
    doc = make_doc()
    assert CL.apply(doc, {0: []}, DPI) == (0, 0)
    assert len(doc.paragraphs) == 7 and doc.paragraphs[2].text.count("\n") == 2


def test_table_cells_untouched():
    """Строки ячеек таблицы уходят в ячейки, а не в абзацы, и раскладка ячеек не меняется."""
    doc = Document()
    doc.add_paragraph("Информационное письмо")
    t = doc.add_table(rows=1, cols=1)
    t.cell(0, 0).paragraphs[0].text = "По месту требования"
    lines = page_lines()
    CL.apply(doc, {0: lines}, DPI)
    assert doc.paragraphs[0].alignment == 1
    assert t.cell(0, 0).paragraphs[0].alignment is None


def test_layout_then_check(tmp_path):
    """Сверка текста после раскладки: деление абзаца по переносам не мешает ей."""
    doc = make_doc()
    CL.apply(doc, {0: page_lines()}, DPI)
    sur = [{"t": p.text, "box": [0, 0, 10, 10], "c": 0.9} for p in doc.paragraphs]
    plain = " ".join(p.text for p in doc.paragraphs).replace("бизнеса", "бизнесa")
    assert CC.check(doc, {0: plain}, {0: sur}) == (0, 0)
    f = tmp_path / "c.docx"
    doc.save(f)
    assert Document(str(f)).paragraphs[6].text == "(подпись)\tСливкин Д.Н."
