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


def _p(page, top, text, bottom=None, x0=85.0, x1=520.0):
    return {"k": "p", "runs": [{"t": text, "x": x0, "x1": x1, "top": top, "size": 12.0}],
            "geom": {"page": page, "top": top, "bottom": bottom or top + 12, "x0": x0, "x1": x1}}


def _pic(page, top, bottom, x0, x1):
    return {"k": "p", "runs": [{"img": {"x0": x0, "x1": x1, "top": top, "bottom": bottom}}],
            "geom": {"page": page, "top": top, "bottom": bottom, "x0": x0, "x1": x1}}


def test_list_item_keeps_no_blocks_from_later_sheets():
    """JPDS л. 20: Word вложил подписи «Рисунок 7, 9–13» листов 21–25 в последний пункт списка листа 20 — они
    печатались стопкой посреди текста листа 20. Блок пункта с более позднего листа, чем сам пункт, — в поток на своё
    место по листу (как «чужое в ячейке» у таблицы); пункт и его текст на своём листе — на месте."""
    cap7, cap9 = _p(21, 275.0, "Рисунок 7 – Ярлык для запуска программы"), _p(22, 395.0, "Рисунок 9 – Элементы формы")
    lst = {"k": "list", "items": [[_p(20, 317.0, "- виртуальная машина Java")],
                                  [_p(20, 333.0, "- офисное приложение (Excel)"), cap7, cap9]]}
    tbl = {"k": "table", "rows": [[{"blocks": [_pic(21, 184.0, 250.0, 164.0, 230.0)], "geom": {"page": 21}}]],
           "geom": {"page": 21, "top": 184.0}}
    flow = [lst, _p(20, 377.0, "2.2 Подготовка к работе"), tbl, _p(21, 297.0, "При успешном запуске"),
            _pic(22, 81.0, 390.0, 100.0, 500.0), _p(22, 417.0, "Таблица 3 – Элементы")]
    out = D.evict(flow)
    texts = [B.text_of(b["runs"]) if b["k"] == "p" else b["k"] for b in out]
    assert texts == ["list", "2.2 Подготовка к работе", "table", "Рисунок 7 – Ярлык для запуска программы",
                     "При успешном запуске", "", "Рисунок 9 – Элементы формы", "Таблица 3 – Элементы"], texts
    assert [len(it) for it in out[0]["items"]] == [1, 1]


def test_list_item_paragraph_on_next_sheet_before_flow_stays():
    """Пункт, переходящий на следующий лист (абзац пункта вверху листа 21, выше блока после списка), — на месте."""
    lst = {"k": "list", "items": [[_p(20, 700.0, "- первый абзац пункта"), _p(21, 60.0, "его продолжение")]]}
    out = D.evict([lst, _p(21, 100.0, "Следующий абзац")])
    assert [len(it) for it in out[0]["items"]] == [2] and len(out) == 2


def test_cell_icon_in_gap_of_its_line_goes_into_that_item():
    """JPDS л. 142, табл. 62, «Кнопки:»: значки (x 157–179,5) стоят в строке пункта между маркером «- » и
    « – добавить файл;», а Word разметил их рисунками после списка ячейки — печатались стопкой в конце ячейки, три — на
    листе «(прод.)». Значок в просвете строки — в этот пункт, в начало его текста."""
    def item(top, tail):
        return {"k": "p", "geom": {"page": 142, "top": top, "bottom": top + 12.07, "x0": 149.66, "x1": 270.0},
                "runs": [{"t": "- ", "x": 149.66, "x1": 153.656, "top": top, "size": 12.0},
                         {"t": tail, "x": 179.42, "x1": 270.0, "top": top + 0.07, "size": 12.0}]}
    cell = [_p(142, 656.2, "Кнопки:", x0=149.66, x1=192.2),
            {"k": "list", "items": [[item(681.22, " – добавить файл;")], [item(731.4, " – обновить данные;")]]},
            _pic(142, 668.14, 690.64, 156.98, 179.48), _pic(142, 718.33, 740.83, 156.98, 179.48),
            _pic(142, 760.0, 790.0, 300.0, 340.0)]                          # не в просвете строки — как было
    tbl = {"k": "table", "rows": [[{"blocks": cell}]]}
    out = D.inline_pictures([tbl])[0]["rows"][0][0]["blocks"]
    assert len(out) == 3 and out[2]["runs"][0]["img"]["x0"] == 300.0
    for it in out[1]["items"]:
        rs = it[0]["runs"]
        assert rs[0]["t"] == "- " and "img" in rs[1] and not rs[1].get("at_tab") and rs[2]["t"].startswith(" –")
