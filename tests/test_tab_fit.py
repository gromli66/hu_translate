# -*- coding: utf-8 -*-
"""E9b-Z3: значение после табуляции не отрывается от подписи (docx_build.Builder._tab_fit)."""
import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parents[1] / "src"))

from docx import Document  # noqa: E402
from docx.shared import Pt, Twips  # noqa: E402

import docx_build as B  # noqa: E402

W = B.W


def _cell_doc(lines, stop=368.2, width=9212):
    """Документ: одна ячейка шириной width twip, абзацы «до ⇥ после» с левой табуляцией stop pt."""
    doc = Document()
    cell = doc.add_table(rows=1, cols=1).cell(0, 0)
    cell.width = Twips(width)
    for i, (before, after) in enumerate(lines):
        p = cell.paragraphs[0] if i == 0 else cell.add_paragraph()
        p.paragraph_format.tab_stops.add_tab_stop(Pt(stop))
        for r in (p.add_run(before), p.add_run("\t" + after)):
            r.font.size = Pt(10)
    b = B.Builder.__new__(B.Builder)
    b.doc = doc
    return b, cell


def _stops(p):
    return [int(t.get(W + "pos")) / 20 for t in p._p.find(f"{W}pPr/{W}tabs")]


def test_value_after_tab_kept_on_line_whole_cell_column_moves_together():
    """30RA+SGT TU: «Наименование: … ⇥ Номер страницы: 3.» — табуляция по месту «Oldalszám» оригинала (368 pt), русская
    подпись длиннее, и «3.» уходило строкой ниже к левому краю. Табуляция — левее, на «край − значение», у обеих строк
    шапки ячейки («Номер версии» над «Номер страницы» — столбцом)."""
    b, cell = _cell_doc([("Идентификатор: 30RA*SGT", "Номер версии: 1."),
                         ("Наименование: Испытание страгивания задвижек второго контура ", "Номер страницы: 3.")])
    b._tab_fit({})
    s1, s2 = _stops(cell.paragraphs[0]), _stops(cell.paragraphs[1])
    lim = 9212 / 20 - 2 * 5.4
    after = B._print_w("Номер страницы: 3.", 10)
    assert s1 == s2 and s1[0] < 368.2 and s1[0] + after <= lim
    assert B._print_w("Наименование: Испытание страгивания задвижек второго контура ", 10) + 4 <= s1[0]


def test_value_fits_nothing_changes():
    b, cell = _cell_doc([("Идентификатор: 30RA*SGT", "Номер версии: 1.")])
    b._tab_fit({})
    assert _stops(cell.paragraphs[0]) == [368.2]


def test_no_room_value_not_torn_bound_together():
    """Текст до табуляции занимает строку до самого стопа: табуляцию не сдвинуть — значение неразрывно (переносится
    целиком, а не «300.» | «0s»)."""
    before = "Следующая арматура закрыта, время хода:"
    while B._print_w(before, 10) < 385:             # текст до табуляции кончается у самого стопа (400 pt)
        before += " а"
    b, cell = _cell_doc([(before, "300.0 с")], stop=400, width=8600)
    b._tab_fit({})
    p = cell.paragraphs[0]
    assert _stops(p) == [400]
    assert "300.⁠0 с" in "".join(t.text for t in p._p.iter(W + "t"))
