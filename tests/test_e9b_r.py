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


def test_row_anchor_is_one_definition_for_cutting_and_assembly():
    """Шаг 0: якорь ряда — rows.is_anchor и в сборке (docx_layout._table_block). Прежде в сборке было своё
    (patterns.ROW_ANCHOR: метка пункта с буквой, код — только один KKS «31UJ21S601»), при нарезке — номер без буквы и
    любой код до трёх слов. Столбец кодов «YA-K7 / YA-K8 / YA-K9» (3PR10 INFO) против значений, слитых нарезкой в один
    абзац, — ряды, как при нарезке: сборка не считала «YA-K8» якорем, и три ряда шли одной строкой."""
    import docx_layout as DL
    import rows as RW
    assert RW.is_anchor("a)") and RW.is_anchor("(3)") and RW.is_anchor("1.2.") and RW.is_anchor("YA-K7")
    assert RW.is_anchor("30BW07 ZÁR") and not RW.is_anchor("nyitva") and not RW.is_anchor("A 30RA01S201 nyitása")
    gh = [(90, 60, 560), (300, 60, 560)]
    gv = [(60, 90, 300), (110, 90, 300), (560, 90, 300)]
    t = (DL._cluster([v[0] for v in gv]), [90, 300], gh, gv)

    def seg(text, x0, y0, y1, x1):
        n = round((y1 - y0) / 12)
        return {"id": 0, "page": 0, "text": text, "loc": {"bbox": [x0, y0, x1, y1], "size": 10, "bold": False,
                "align": "left", "frags": [[x0, y0 + 12 * k, x1, y0 + 12 * (k + 1)] for k in range(n)]}}
    segs = [seg("YA-K7", 70, 100, 112, 100), seg("YA-K8", 70, 124, 136, 100), seg("YA-K9", 70, 148, 160, 100),
            seg("126 bar 128 bar 130 bar", 120, 100, 160, 180)]
    b = DL._table_block(t, segs, 0, None, "Times New Roman", 10, False)
    text = lambda c: ["".join(r.get("t", "") for r in p["runs"]) for p in c["blocks"]]
    assert [text(r[0]) for r in b["rows"]] == [["YA-K7"], ["YA-K8"], ["YA-K9"]], [[text(c) for c in r] for r in b["rows"]]


def _p(text, top, x0=80.0):
    g = {"page": 1, "top": top, "bottom": top + 10, "x0": x0, "x1": x0 + 60, "first_x0": x0, "size": 10, "lines": 1}
    return {"k": "p", "runs": [{"t": text, "font": "Tahoma", "size": 10, "b": False, "i": False, "color": "000000"}],
            "geom": g}


def _blank():
    return {"k": "p", "runs": [{"t": "", "font": "Tahoma", "size": 10, "b": False, "i": False, "color": "000000"}]}


def test_form_marks_cell_is_cut_against_its_lines():
    """П. 11 просмотра DOCX 06.11 (TU «Пакш», 30RVRXTSBIGT_TU л. 1–4): столбец отметок «O» в Word — одна ячейка на весь
    ряд, «O» выровнены пустыми абзацами против строк пункта. sub_rows резал ряд по меткам «1.1.», «1.2.» и паре, а
    ячейка отметок была «прочей» — все «O» стопкой в первой подстроке. Отметка на высоте абзаца пары — своя подстрока."""
    import docx_struct as D
    cell = lambda *bl: {"blocks": list(bl), "th": False, "geom": {"page": 1, "top": 100, "bottom": 160, "x0": 70,
                                                                  "x1": 100}}
    row = [cell(_p("1.1.", 100, 70), _blank(), _blank(), _p("1.2.", 140, 70)),
           cell(_p("Az alábbi armatúrák zárva:", 100, 120), _p("30TK30S201", 120, 120), _p("A 31RD31P702 kör", 140, 120)),
           cell(_blank(), _p("O", 120, 480), _blank(), _p("O", 140, 480))]
    out = D.sub_rows([{"k": "table", "rows": [row]}])[0]["rows"]
    text = lambda c: [D.text_of(x["runs"]) for x in c["blocks"]]
    assert [[text(c) for c in r] for r in out] == [
        [["1.1."], ["Az alábbi armatúrák zárva:"], []], [[], ["30TK30S201"], ["O"]], [["1.2."], ["A 31RD31P702 kör"], ["O"]]]
    assert not any(c.get("vcont") for r in out for c in r)
    # без меток (номер пункта один, «1.15.»): пара — ячейка слева, с абзацами на высотах отметок; текст выше — своя подстрока
    row = [cell(_p("1.15.", 100, 70)),
           cell(_p("TO szakterület által:", 100, 120), _p("30RR21S201", 110, 120), _p("30RR21S401", 120, 120)),
           cell(_blank(), _p("O", 110, 480), _p("O", 120, 480))]
    out = D.sub_rows([{"k": "table", "rows": [row]}])[0]["rows"]
    assert [[text(c) for c in r] for r in out] == [
        [["1.15."], ["TO szakterület által:"], []], [[], ["30RR21S201"], ["O"]], [[], ["30RR21S401"], ["O"]]]
