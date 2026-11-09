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


def test_form_mark_is_one_definition_for_cell_valign():
    """Шаг 0б: ячейка по низу (отметка против нижней строки многострочного абзаца рядом) — только ячейка из отметок
    rows.FORM_MARK, как при нарезке. Прежде своё правило «до двух знаков без букв и цифр»: прочерк «-» (нет значения,
    02VE25 л. 41) вставал по низу ячейки; «X» (крестик 30BL150FÜT 4.2.2) и греческая «Ο» (30SZBV*IT TU) — отметки."""
    import docx_layout as DL
    gh = [(90, 60, 560), (400, 60, 560)]
    gv = [(60, 90, 400), (300, 90, 400), (560, 90, 400)]
    t = (DL._cluster([v[0] for v in gv]), [90, 400], gh, gv)

    def seg(text, x0, y0, y1):
        return {"id": 0, "page": 0, "text": text, "loc": {"bbox": [x0, y0, x0 + 200, y1], "frags": [[x0, y0, x0 + 200, y1]],
                                                          "size": 10, "bold": False, "align": "left"}}
    for mark, want in (("-", None), ("X", "bottom"), ("Ο", "bottom"), ("O", "bottom")):
        segs = [seg("YP10T401, YP10T701, YP14T001XQ01, YP24T001XQ01, YP29T001XQ01.", 70, 100, 160), seg(mark, 310, 150, 160)]
        b = DL._table_block(t, segs, 0, None, "Times New Roman", 10, False)
        assert [c.get("valign") for c in b["rows"][0]] == [None, want], mark


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


def test_sub_rows_have_no_cell_margins_inside_their_row(tmp_path):
    """R2а-доводка: подстроки одного ряда оригинала (строка кода и её «O») — без полей ячейки по сторонам разреза и с
    зазором оригинала до следующей подстроки (не больше строки), как полосы docx_layout._table_block. С полями таблицы
    (TU «Пакш»: 6,7 + 10,2 pt) строка кода выходила втрое выше оригинала — 30RVRXTSBIGT_TU 27 → 32 листа."""
    import docx_struct as D
    cell = lambda *bl: {"blocks": list(bl), "th": False, "geom": {"page": 1, "top": 100, "bottom": 160, "x0": 70,
                                                                  "x1": 100}}
    row = [cell(_p("1.15.", 100, 70)),
           cell(_p("TO szakterület által:", 100, 120), _p("30RR21S201", 120, 120), _p("30RR21S401", 130.5, 120)),
           cell(_blank(), _p("O", 120, 480), _p("O", 130.5, 480))]
    tab = D.sub_rows([{"k": "table", "rows": [row], "borders": True,
                       "geom": {"page": 1, "top": 95, "bottom": 165, "x0": 65, "x1": 526, "first_x0": 65, "size": 10}}])[0]
    assert [[c.get("cut"), c.get("pad")] for c in (r[1] for r in tab["rows"])] == [
        [["bottom"], 10], [["top", "bottom"], None], [["top"], None]]
    out = tmp_path / "t.docx"
    B.Builder(B.Memory([]), lambda img: png(), PROFILE).build([tab], str(out), dict(LAYOUT, first_is_body=True))
    tb = Document(str(out)).tables[0]

    def mar(ri, ci):
        m = tb.rows[ri].cells[ci]._tc.tcPr.find(W + "tcMar")
        return {} if m is None else {el.tag.split("}")[1]: el.get(W + "w") for el in m}
    assert [mar(ri, 1) for ri in range(3)] == [{"bottom": "0"}, {"top": "0", "bottom": "0"}, {"top": "0"}]
    assert [mar(ri, 2) for ri in range(3)] == [{"bottom": "0"}, {"top": "0", "bottom": "0"}, {"top": "0"}]
    assert tb.rows[0].cells[1].paragraphs[-1].paragraph_format.space_after.pt == 10


def _grid1(xs, y0, y1):
    """Сетка таблицы листа без внутренних горизонталей: одна строка сетки, вертикали по xs."""
    import docx_layout as DL
    gh = [(y0, xs[0], xs[-1]), (y1, xs[0], xs[-1])]
    gv = [(x, y0, y1) for x in xs]
    return DL._cluster([v[0] for v in gv]), [y0, y1], gh, gv


def _seg(text, x0, y0, x1, lines=1, step=12):
    return {"id": 0, "page": 0, "text": text, "loc": {"bbox": [x0, y0, x1, y0 + step * lines], "size": 10, "bold": False,
            "align": "left", "frags": [[x0, y0 + step * k, x1, y0 + step * (k + 1)] for k in range(lines)]}}


def test_band_anchor_is_the_rows_anchor():
    """R6: якорь полосы _table_block — тот же, что при нарезке (rows.Rows): «YA-K7 ⇥ условие» (3PR10 л. 22–23) — код не
    KKS, прежнее правило (patterns.ROW_ANCHOR в первом столбце) его не брало: абзац условия, заходящий на уровень
    следующего кода, склеивал ряды — «YA-K7 YA-K8» одной ячейкой, условия мимо своих кодов."""
    import docx_layout as DL
    t = _grid1([60, 110, 400], 90, 200)
    segs = [_seg("YA-K7", 70, 100, 100), _seg("YA-K8", 70, 124, 100), _seg("YA-K9", 70, 160, 100),
            _seg("Megengedett a hidegvizes indítás 0,8 U névl tápfeszültségnél, melegvizes indítás", 120, 100, 390,
                 lines=3), _seg("Tilos az FKSZ indítása, ha merev a primerkör", 120, 160, 390)]
    b = DL._table_block(t, segs, 0, None, "Times New Roman", 10, False)
    text = lambda c: ["".join(r.get("t", "") for r in p["runs"]) for p in c["blocks"]]
    assert [text(r[0]) for r in b["rows"]] == [["YA-K7"], ["YA-K8"], ["YA-K9"]], [[text(c) for c in r] for r in b["rows"]]


def test_code_under_a_phrase_in_phrase_column_is_not_a_band_anchor():
    """R6, ограда (3SZ23 л. 6, шаг 0 сессии 1): код не-KKS под фразой в столбце «Tevékenység» («Az oltás indítását
    követően a» / «31(32)NB11S601 NYIT.», значение рядом — на обе строки) — содержание, а не якорь: в столбце фразы
    преобладают, ряд не режется (якорь в первом столбце по rows.is_anchor отрывал фразу от кода)."""
    import docx_layout as DL
    t = _grid1([60, 300, 400], 90, 200)
    segs = [_seg("Az oltás indítását követően a", 70, 100, 290), _seg("31(32)NB11S601 NYIT.", 70, 112, 200),
            _seg("A tartály töltése megkezdődik", 70, 124, 290), _seg("Indítás után nyitva", 310, 100, 390, lines=2)]
    b = DL._table_block(t, segs, 0, None, "Times New Roman", 10, False)
    assert len(b["rows"]) == 1
