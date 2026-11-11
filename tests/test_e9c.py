"""E9c: средние дефекты вида (11.11) — «N/NN» пути А, выравнивание столбца, список «код — описание», JPDS."""
from docx.oxml.ns import qn

import docx_build as B
import docx_struct as D

from test_docx_defects import _build_sheets
from test_struct_docx import PROFILE, png


def _hdr_xml(doc):
    out = []
    for s in doc.sections:
        for part in (s.header, s.footer, s.first_page_header, s.first_page_footer):
            if not part.is_linked_to_previous:
                out.append(part._element.xml)
    return "\n".join(out)


def test_path_a_sheet_count_is_original_text_not_numpages(tmp_path):
    """Путь А: номер листа — оригинала (раздел на лист, нумерация с его числа), и «из NN» — число листов оригинала
    текстом. Полем NUMPAGES LibreOffice печатал «10/36» при 30 листах перевода и 24 оригинала (OAH 1.1; UTS «33/103»
    при 101 листе и 99 оригинала): считал свои невыводимые пустые листы при перезапуске нумерации."""
    bld, doc, L = _build_sheets(tmp_path, pages=4, head=lambda pno: (0, f"{pno + 3}/24"))
    assert L.get("sheet_no") == {1: 3, 2: 4, 3: 5, 4: 6}, L.get("sheet_no")
    xml = _hdr_xml(doc)
    assert "NUMPAGES" not in xml and " PAGE" in xml
    assert "/24" in "".join(t.text or "" for t in doc.sections[0].header._element.iter(qn("w:t")))


def test_without_path_a_sheet_count_stays_numpages_field():
    """Без пути А (счётчик листа не помечен решателем, «N/NN» без номера по оригиналу) — как прежде, NN полем."""
    from docx import Document as Doc
    bld = B.Builder(B.Memory([]), lambda img: png(), PROFILE)
    p = Doc().add_paragraph()
    bld._zone_text(p, D.PAGE_MARK + "/24", False, "7", page_marked=True)
    assert "NUMPAGES" in p._p.xml and " PAGE" in p._p.xml
    bld.L = {"sheet_no": {1: 7}}
    p2 = Doc().add_paragraph()
    bld._zone_text(p2, D.PAGE_MARK + "/24", False, "7", page_marked=True)
    assert "NUMPAGES" not in p2._p.xml and " PAGE" in p2._p.xml and p2.text.endswith("/24")


def _cell(x0, x1, lines=1, page=5, th=False):
    return {"th": th, "blocks": [], "geom": {"page": page, "x0": x0, "x1": x1, "lines": lines}}


def test_full_width_line_is_no_vote_for_right():
    """30RR2126GT л. 5, столбец «Feltétel, művelet, ellenőrzés» (319 pt): единственная однострочная ячейка — заголовок
    «2. A TESZT ELŐKÉSZÍTÉS INTÉZKEDÉSEI» во всю ширину (слева 25,7 pt — отступ столбца, справа 4,6) — голосовала «вправо»
    и прижимала вправо всю прозу столбца (16 абзацев). Строка во всю ширину — не голос за «вправо»."""
    xs = [65.5, 113.5, 432.5, 526.1]
    rows = [[_cell(75.2, 103.4, th=True), _cell(208.1, 340.3, th=True), _cell(458.0, 500.0, th=True)],
            [_cell(81.2, 100.6), _cell(139.2, 427.9), _cell(508.2, 516.0)],
            [_cell(81.2, 100.6), _cell(136.8, 409.6, lines=2), _cell(508.2, 516.0)]]
    al = B._col_align(rows, xs, 5, any_page=True)
    assert ("body", 1) not in al, al


def test_short_right_line_still_votes_right():
    """Номера пунктов, прижатые вправо в узком столбце (30RTKIT: «4.3.» в столбце 42,7 pt, слева 24,5, справа 0,7) — как
    прежде «вправо»: строка не во всю ширину."""
    xs = [70.0, 112.7, 430.0]
    rows = [[_cell(94.5, 112.0), _cell(120.0, 300.0)] for _ in range(4)]
    assert B._col_align(rows, xs, 5)[("body", 0)] == "right"
