# E9b-R «ряд в сборке» (E9_DESIGN.md §6): одно определение ряда, ячейки и столбца для сборки — тест на шаг.
from docx import Document

import docx_build as B
from test_struct_docx import LAYOUT, PROFILE, W, png


def _cell(text, x0, x1, top, tabs=None):
    run = {"t": text, "font": "Tahoma", "size": 10, "b": False, "i": False, "color": "000000"}
    if tabs:
        run["tabs"] = tabs
    g = {"page": 1, "top": top, "bottom": top + 10, "x0": x0, "x1": x1, "first_x0": x0, "size": 10, "lines": 1}
    return {"blocks": [{"k": "p", "runs": [run], "geom": g, "seg": True}], "th": False, "geom": g}


def test_short_row_cells_take_their_frame_columns_in_word_fit_too(tmp_path):
    """«nyitva» 30RA+SGT TU (DEFECTS4 §4): таблица по разметке 48 | 319 | 94 pt на 28 листов; ряд оглавления раздела на
    листе 4 — две ячейки на три столбца (текст на столбцы 0–1, номер листа — столбец 2). table() кладёт ячейки
    короткого ряда по рамке (_span_starts), а _fit_words — по порядку: длинное слово оглавления меряло столбец 0, он рос
    48 → 105,8 pt за счёт столбца 1, и табуляция «30RA01S201 ⇥ nyitva» по месту оригинала (248 pt от края текста)
    уходила за край ячейки — LibreOffice «ОТКРЫТО» не печатал (0 из 93). Ячейка → столбец — одним правилом (по рамке)."""
    xs = [65.4, 113.4, 432.4, 526.0]
    rows = [[_cell("1.", 75, 90, 100), _cell("A tesztelés feltételei", 118, 300, 100), _cell("x", 457, 470, 100)],
            [_cell("Tartalomjegyzék technológiai-paraméter-beállítás", 70.8, 430.6, 115), _cell("6.", 472.9, 485.5, 115)],
            [_cell("4.7.10.", 75, 104, 130), _cell("30RA01S201\tnyitva", 118.8, 400, 130, tabs=[366.6]),
             _cell("x", 457, 470, 130)]]
    tab = {"k": "table", "rows": rows, "borders": True, "frame": {"xs": xs, "ys": [95, 112, 127, 142]},
           "geom": {"page": 1, "top": 95, "bottom": 142, "x0": xs[0], "x1": xs[-1], "first_x0": xs[0], "size": 10}}
    mem = B.Memory([("Tartalomjegyzék technológiai-paraméter-beállítás",
                     "Оглавление технологических-параметров-настройки"), ("30RA01S201\tnyitva", "30RA01S201\tОТКРЫТО")])
    out = tmp_path / "t.docx"
    B.Builder(mem, lambda img: png(), PROFILE).build([tab], str(out), dict(LAYOUT, first_is_body=True))
    tb = Document(str(out)).tables[0]
    widths = [int(gc.get(W + "w")) / 20 for gc in tb._tbl.tblGrid.findall(W + "gridCol")]
    assert abs(widths[0] - 48) < 1 and abs(widths[1] - 319) < 1, widths
    # табуляция «⇥ ОТКРЫТО» — внутри текста своей ячейки
    stops = [int(t.get(W + "pos")) / 20 for t in tb._tbl.iter(W + "tab") if t.get(W + "pos")]
    assert stops and max(stops) < widths[1] - 2 * 5.4, (stops, widths)

