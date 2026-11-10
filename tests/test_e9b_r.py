# E9b-R «ряд в сборке» (E9_DESIGN.md §6): одно определение ряда, ячейки и столбца для сборки — тест на шаг.
from docx import Document
from docx.shared import Pt

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
    prof = dict(PROFILE, table=dict(PROFILE.get("table") or {}, pad_top=8.7, pad_bottom=12.2))   # поля TU
    B.Builder(B.Memory([]), lambda img: png(), prof).build([tab], str(out), dict(LAYOUT, first_is_body=True))
    tb = Document(str(out)).tables[0]

    # LibreOffice поля ячейки сверху и снизу (tcMar) не берёт (печать 10.11: 30RVRXTSBIGT_TU 27 → 31 листа): поля таблицы
    # сверху и снизу — нулевые, отбивка — у настоящих границ ряда, по сторонам разреза — нет
    m = {el.tag.split("}")[1]: el.get(W + "w") for el in tb._tbl.tblPr.find(W + "tblCellMar")}
    assert m["top"] == m["bottom"] == "0"
    sp = lambda ri, ci: ((tb.rows[ri].cells[ci].paragraphs[0].paragraph_format.space_before or Pt(0)).pt,
                         (tb.rows[ri].cells[ci].paragraphs[-1].paragraph_format.space_after or Pt(0)).pt)
    top, bot = sp(0, 2)[0], sp(2, 2)[1]
    assert top > 0 and bot > 0
    for ci in (1, 2):                                       # зазор оригинала до следующей подстроки — 10 pt
        assert [sp(ri, ci) for ri in range(3)] == [(top, 10), (0, 0), (0, bot)], ci
    # у ячеек нет своих линий (Table Grid) — по сторонам разреза линий нет (30RA+SGT TU 4.1.x, печать 10.11)
    bd = lambda ri, ci: {el.tag.split("}")[1]: el.get(W + "val") for el in tb.rows[ri].cells[ci]._tc.tcPr.find(W + "tcBorders")}
    assert [(bd(ri, 1)["top"], bd(ri, 1)["bottom"]) for ri in range(3)] == [
        ("single", "nil"), ("nil", "nil"), ("nil", "single")]


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
    # ряд делится, только если у каждой полосы свой текст: условие K7 накрывает и K8 (нарезка слила их) — K7 и K8 стопкой
    # при общем тексте; K9 со своим условием — свой ряд (печать 10.11: иначе строка K7 пустая, её условие — у K8)
    assert [[text(c) for c in r] for r in b["rows"]] == [
        [["YA-K7", "YA-K8"], ["Megengedett a hidegvizes indítás 0,8 U névl tápfeszültségnél, melegvizes indítás"]],
        [["YA-K9"], ["Tilos az FKSZ indítása, ha merev a primerkör"]]], [[text(c) for c in r] for r in b["rows"]]


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


def test_codes_of_nearly_one_width_flush_right_are_right(tmp_path):
    """П. 11 просмотра DOCX 06.11 (EPM 30PR132GT л. 35, 30PR25GT л. 14): перечень кодов столбца «Megnevezés» выключен
    вправо — правые края 444,7 у всех, левые 390,5…394; середины расходятся меньше 10 pt, и правило «столбец вправо»
    (середины ≥ 10 pt — против столбца по центру) их не брало: коды шли влево. По центру разошлись бы на ту же величину и
    правые края; вправо — правые ровно (≤ 0,5 pt), левые — нет (≥ 1,5 pt, три разных начала). Заголовок группы у
    левого края — как был; шапка по центру и два одинаковых «A 203/3» (30PR132GT л. 25) — не вправо."""
    def para(t, x0, x1, top):
        g = {"page": 1, "top": top, "bottom": top + 12, "x0": x0, "x1": x1, "first_x0": x0, "size": 10, "lines": 1,
             "last_x1": x1}
        return {"k": "p", "runs": [{"t": t, "font": "Times New Roman", "size": 10, "b": False, "i": False,
                                    "color": "000000"}], "geom": g, "seg": True, "align": "left"}
    edges = {"top": True, "bottom": True, "left": True, "right": True}
    cell = lambda ps: {"blocks": ps, "th": False, "span": 1, "edges": edges, "geom": dict(ps[0]["geom"]) if ps else None}

    def build(body):
        rows = [[cell([]), cell([para(t, a, z, 160 + 20 * k)]), cell([para("O", 507, 515, 160 + 20 * k)])]
                for k, (t, a, z) in enumerate(body)]
        tab = {"k": "table", "rows": rows, "borders": True, "frame": {"xs": [71, 137, 463, 530], "ys": [
            150 + 20 * k for k in range(len(body) + 1)]}, "geom": {"page": 1, "top": 150, "bottom": 150 + 20 * len(body),
                                                                    "x0": 71, "x1": 530, "first_x0": 71, "size": 10}}
        out = tmp_path / "e.docx"
        B.Builder(B.Memory([]), lambda img: png(), PROFILE).build([tab], str(out), dict(LAYOUT, first_is_body=True))
        return [p.alignment for c in Document(str(out)).tables[0].columns[1].cells for p in c.paragraphs if p.text]
    epm = [("30YA53D001", 167.6, 220.9), ("30TK65S201", 394.0, 444.7), ("30TK65N001", 392.4, 444.7),
           ("30TK65W001", 390.5, 444.7), ("30TK65T101", 393.5, 444.7), ("30TK65P001", 394.0, 444.7)]
    assert build(epm) == [None, 2, 2, 2, 2, 2]
    head = [("Reaktor Osztály ellenőrzési terjedelme", 140.0, 359.0), ("Megnevezés", 268.7, 321.1),
            ("A 203/3", 300.0, 321.4), ("A 203/3", 300.0, 321.4)]
    assert 2 not in build(head)


def test_bands_have_no_cell_margins_inside_their_grid_row():
    """R2 ↔ полосы: полоса строки сетки (_table_block) — как подстрока sub_rows: без полей ячейки по сторонам разреза
    (c["cut"] → tcMar 0 в docx_build, тест выше), зазор до следующей — pad. Поля таблицы (TU: 6,7 + 10,2 pt) у каждой
    полосы раздували ряды «код ⇥ условие» сверх оригинала."""
    import docx_layout as DL
    t = _grid1([60, 110, 400], 90, 200)
    segs = [_seg("YA-K7", 70, 100, 100), _seg("YA-K8", 70, 130, 100), _seg("YA-K9", 70, 160, 100),
            _seg("Megengedett a hidegvizes indítás", 120, 100, 390), _seg("Tilos az FKSZ indítása", 120, 130, 390),
            _seg("Megengedett 20 óra időtartamra", 120, 160, 390)]
    b = DL._table_block(t, segs, 0, None, "Times New Roman", 10, False)
    assert [[c.get("cut") for c in r] for r in b["rows"]] == [[["bottom"]] * 2, [["top", "bottom"]] * 2, [["top"]] * 2]


def test_row_tail_on_next_sheet_continues_a_phrase_as_in_segmentation():
    """R3: хвост ряда на новом листе (docx_struct._cut_rows) — продолжение фразы не только «со строчной», но и как строки
    фразы при нарезке (lines.mid_phrase, hyphen_break): код после служебного слова («…szelepet a» / «30UJ97B001 felé
    NYITNI.», 3SZ17 л. 9) и перенос слова прописными («KI-» / «KAPCSOLVA»). Новый пункт с прописной — свой ряд."""
    import docx_struct as D

    def cell(t, page, x0):
        g = {"page": page, "top": 100.0, "bottom": 110.0, "x0": x0, "x1": x0 + 90, "first_x0": x0, "size": 10, "lines": 1}
        return {"blocks": [_p(t, 100.0, x0) | {"geom": g}], "th": False, "geom": g, "attrs": {}}
    for up, down, glued in (("Kinyitni a háromutú szelepet a", "30UJ97B001 felé NYITNI.", True),
                            ("A szivattyút KI-", "KAPCSOLVA állapotba kell hozni", True),
                            ("Kinyitni a háromutú szelepet.", "Szivattyú indítása", False)):
        got = D._cut_rows([[cell("1.", 1, 57), cell(up, 1, 177)], [cell(down, 2, 177)]])
        assert (len(got) == 1) is glued, (up, down)


def test_band_without_text_is_not_a_row():
    """R6: якорь ниже начала абзаца соседней ячейки («@0YATBEATL» против «Átlagos zóna belépő hőmérséklet» в две строки,
    01-02FMP0BIGT л. 24) резал полосу над собой, абзац уходил в полосу своей середины — над рядом вставал пустой ряд.
    Полоса без текста — не ряд."""
    import docx_layout as DL
    t = _grid1([60, 110, 400], 90, 200)
    segs = [_seg("QA1121", 70, 112, 100), _seg("QA1122", 70, 150, 100),
            _seg("Átlagos zóna belépő hőmérséklet", 120, 100, 390, lines=2), _seg("Bórsavkonc. (kézi)", 120, 150, 390)]
    b = DL._table_block(t, segs, 0, None, "Times New Roman", 10, False)
    text = lambda c: ["".join(r.get("t", "") for r in p["runs"]) for p in c["blocks"]]
    assert [[text(c) for c in r] for r in b["rows"]] == [
        [["QA1121"], ["Átlagos zóna belépő hőmérséklet"]], [["QA1122"], ["Bórsavkonc. (kézi)"]]]


def test_sub_rows_cut_like_bands_and_split_lists_by_item():
    """R2: места разреза подстрок sub_rows — rows.bands, как полосы _table_block (якоря — метки и отметки): «O» против
    пунктов списка «− 30EV …» (30RVRXTSBIGT_TU 1.15) — каждая в подстроке своего пункта, список делится по пунктам;
    прежде список был одним блоком в одной подстроке, «O» — стопкой или в пустых подстроках под ним."""
    import docx_struct as D
    cell = lambda *bl: {"blocks": list(bl), "th": False, "geom": {"page": 1, "top": 100, "bottom": 160, "x0": 70,
                                                                  "x1": 100}}
    lst = {"k": "list", "kind": "bullet", "items": [[_p("− 30EV felől táplálás", 120, 120)],
                                                    [_p("− be kell kapcsolni", 130, 120)]]}
    row = [cell(_p("1.15.", 100, 70)), cell(_p("Megjegyzés:", 110, 120), lst),
           cell(_blank(), _p("O", 120, 480), _p("O", 130, 480))]
    out = D.sub_rows([{"k": "table", "rows": [row]}])[0]["rows"]
    text = lambda c: [D.text_of(x["runs"]) if x["k"] == "p" else [D.text_of(i[0]["runs"]) for i in x["items"]]
                      for x in c["blocks"]]
    assert [[text(c) for c in r] for r in out] == [
        [["1.15."], ["Megjegyzés:"], []], [[], [["− 30EV felől táplálás"]], ["O"]], [[], [["− be kell kapcsolni"]], ["O"]]]
def test_nested_table_over_many_rows_keeps_its_columns(tmp_path):
    """R1: граница столбца рамки — вертикаль на половину высоты таблицы на листе (против вложенной таблички в одну строку,
    JPDS табл. 15) или через три ряда и больше: вложенная таблица «Megnevezés | Állapot» ×2 на много рядов в ячейке
    «Megnevezés» (30BOXVILGT_TU л. 3) короче половины листа — рамка была в два столбца, fit_rows сворачивал ряд в одну
    ячейку «30XL11S201 ZÁRVA 30XL53S201 ZÁRVA», отметки пропадали."""
    import docx_struct as D
    from test_cloud_sidefixes import _tagged_pdf, _tj
    line = lambda x0, y0, x1, y1: f"0.5 w {x0} {y0} m {x1} {y1} l S "
    c = "".join(line(70, y, 520, y) for y in (760, 740, 720, 300))                       # линии внешней таблицы
    c += "".join(line(130, y, 520, y) for y in (700, 680, 660, 640))                    # ряды вложенной
    c += "".join(line(x, 760, x, 300) for x in (70, 130, 520))
    c += "".join(line(x, 720, x, 640) for x in (230, 330, 430))                         # её столбцы — 4 ряда из 6
    texts = [(75, 745, "Sorszam"), (140, 745, "Megnevezes"), (75, 725, "6.1"), (140, 725, "GB12")]
    for k, y in enumerate((705, 685, 665, 645)):
        texts += [(140, y, f"30XL1{k}S201"), (240, y, "ZARVA"), (340, y, f"30XL5{k}S201"), (440, y, "ZARVA")]
    texts += [(140, 310, "A lapot kitoltotte")]
    c += "".join(f"/P <</MCID {m}>> BDC " + _tj(x, y, t) + "EMC " for m, (x, y, t) in enumerate(texts))
    td = lambda m: ("TD", [("P", [m])])
    tree = [("Table", [("TR", [td(0), td(1)]), ("TR", [td(2), td(3)])]
             + [("TR", [("TD", [])] + [td(4 + 4 * k + j) for j in range(4)]) for k in range(4)]
             + [("TR", [("TD", []), td(20)])])]
    _tagged_pdf(tmp_path / "t.pdf", c, tree)
    blocks, _ = D.load(str(tmp_path / "t.pdf"))
    tab = next(b for b in blocks if b["k"] == "table")
    assert len(tab["frame"]["xs"]) == 6, tab["frame"]["xs"]
    assert [len(r) for r in tab["rows"]][2:6] == [4, 4, 4, 4]                            # не свёрнуты в две


def test_row_asked_as_one_line_before_r1_is_not_lost_when_cells_are_taken():
    """R1 + учёт: ряд вложенной таблицы, что сборка прежде спрашивала одной строкой (fit_rows сворачивал его в ячейку, и
    добор struct_fill завёл struct_only «30XL11S201 ZÁRVA □ 30XL53S201 ZÁRVA □», 30BOXVILGT_TU), теперь берётся по ячейкам:
    его текст ровно склеивается из подряд идущих взятых сегментов листа — устарел, не потерян. Ряд, где ячейка не взята
    («Ellenőrzés» без перевода до добора), — потерян."""
    import assemble as A
    seg = lambda i, t, extra=None: {"id": i, "text": t, "src": t, "page": 2, "loc": extra or {"bbox": [100, 100 + i, 150, 110 + i]}}
    # отметки «□» на листе — после обоих кодов, в строке разметки — после своего «ZÁRVA»
    segs = [seg(1, "30XL11S201"), seg(2, "ZÁRVA"), seg(3, "30XL53S201"), seg(4, "ZÁRVA"), seg(7, ""),
            seg(8, ""), seg(5, "Megnevezés"), seg(6, "Ellenőrzés ROP"),
            seg(9, "30XL11S201 ZÁRVA  30XL53S201 ZÁRVA ", {"struct_only": True}),
            seg(10, "Megnevezés Ellenőrzés", {"struct_only": True})]
    res = {str(s["id"]): {"ru": s["text"]} for s in segs}
    lost = A.lost_ids(segs, res, used={1, 2, 3, 4, 5, 7, 8}, exempt={}, tree=True)
    assert [s["id"] for s, _ in lost] == [10]


def test_paragraph_after_list_in_cell_stays_after_it(tmp_path):
    """Ячейка, начатая списком: абзац после списка писался в пустой начальный абзац ячейки — над списком («Ha nem
    állnak…» вставал над пунктом «− Az RVR …» своей подстроки, 30RVRXTSBIGT_TU 1.15, печать 10.11)."""
    lst = {"k": "list", "kind": "bullet", "items": [[_p("− Az RVR turbina gyorszárak", 100, 120)]]}
    cell = lambda *bl: {"blocks": list(bl), "th": False, "geom": {"page": 1, "top": 100, "bottom": 130, "x0": 70,
                                                                  "x1": 300}}
    tab = {"k": "table", "rows": [[cell(_p("1.15.", 100, 70)), cell(lst, _p("Ha nem állnak rendelkezésre", 112, 130))]],
           "borders": True, "geom": {"page": 1, "top": 95, "bottom": 135, "x0": 65, "x1": 526, "first_x0": 65, "size": 10}}
    out = tmp_path / "l.docx"
    B.Builder(B.Memory([]), lambda img: png(), PROFILE).build([tab], str(out), dict(LAYOUT, first_is_body=True))
    texts = [p.text for p in Document(str(out)).tables[0].rows[0].cells[1].paragraphs if p.text.strip()]
    assert [t[:8] for t in texts] == ["Az RVR t", "Ha nem á"], texts          # «−» — нумерацией списка


# ------------------------------------------------------------------ стопка кодов (E9B_R2, печать 10.11): rows.stack_bands
def _rows_text(b):
    text = lambda c: ["".join(r.get("t", "") for r in p["runs"]) for p in c["blocks"]]
    return [[text(c) for c in r] for r in b["rows"]]


def test_code_band_without_text_joins_the_stack_under_common_text():
    """3PR10 л. 22–23: нарезка слила условия YA-K7…K9 в один абзац; полосы резали ряд по каждому коду — строки K7 и K9
    пустые, их условия у K8. Полоса с кодом без своего текста (у K7 только «1» в столбце «M»), которую накрывает абзац, —
    в стопку кодов при общем тексте; K10 со своим условием — свой ряд."""
    import docx_layout as DL
    t = _grid1([60, 110, 400, 430], 90, 200)
    segs = [_seg("YA-K7", 70, 100, 100), _seg("1", 410, 100, 420), _seg("YA-K8", 70, 112, 100),
            _seg("YA-K9", 70, 124, 100), _seg("YA-K10", 70, 140, 100),
            _seg("Megengedett a hidegvizes indítás. Tilos az FKSZ beszakaszolása. Tilos az FKSZ indítása", 120, 100, 390,
                 lines=3), _seg("Tilos az FKSZ indítása, ha merev a primerkör", 120, 140, 390)]
    b = DL._table_block(t, segs, 0, None, "Times New Roman", 10, False)
    assert _rows_text(b) == [
        [["YA-K7", "YA-K8", "YA-K9"], ["Megengedett a hidegvizes indítás. Tilos az FKSZ beszakaszolása. Tilos az FKSZ "
                                       "indítása"], ["1"]],
        [["YA-K10"], ["Tilos az FKSZ indítása, ha merev a primerkör"], []]], _rows_text(b)


def test_mark_band_without_text_joins_the_item_it_hangs_over():
    """01-02FMP0BIGT 4.10.13 (лист 17): «O» пункта стоит на первой строке его текста, но выше номера — своя полоса без
    текста, которую накрывает абзац пункта: «O» в ряду своего пункта, а не отдельной строкой над ним (E9B_R2: «лучше»)."""
    import docx_layout as DL
    t = _grid1([80, 125, 520, 560], 470, 570)
    box = lambda text, bb, fr=None: {"id": 0, "page": 0, "text": text, "loc": {"bbox": bb, "size": 10, "bold": False,
                                                                              "align": "left", "frags": fr or [bb]}}
    segs = [box("4.10.12", [84, 484, 117, 497]), box("ÜM engedélyezte a MET-re hozást", [131, 484, 274, 497]),
            box("O", [533, 483, 543, 497]), box("O", [533, 506, 543, 520]),
            box("A MET-re hozást a PR04 szerint végre kell hajtani", [131, 507, 507, 532],
                [[131, 507, 507, 519], [131, 520, 400, 532]]), box("4.10.13", [84, 518, 117, 532]),
            box("4.10.14", [84, 541, 117, 555]), box("A reaktor MET-en", [131, 541, 208, 555]), box("O", [533, 541, 543, 555])]
    b = DL._table_block(t, segs, 0, None, "Times New Roman", 10, False)
    assert _rows_text(b) == [
        [["4.10.12"], ["ÜM engedélyezte a MET-re hozást"], ["O"]],
        [["4.10.13"], ["A MET-re hozást a PR04 szerint végre kell hajtani"], ["O"]],
        [["4.10.14"], ["A reaktor MET-en"], ["O"]]], _rows_text(b)


def test_code_band_joins_the_text_that_covers_it_not_the_item_above():
    """30RR2126GT л. 5 (прогон h13): номер «2.3.» стоит по середине своего текста в три строки — текст начат выше номера
    и сам по себе полоса ниже. Полоса без текста встаёт к тексту, что её накрывает, а не к хвосту пункта 2.2 выше."""
    import docx_layout as DL
    t = _grid1([75, 130, 465, 500], 200, 340)
    box = lambda text, bb, fr=None: {"id": 0, "page": 0, "text": text, "loc": {"bbox": bb, "size": 11, "bold": False,
                                                                              "align": "left", "frags": fr or [bb]}}
    segs = [box("2.2.", [81, 211, 101, 226]), box("O", [475, 211, 486, 226]),
            box("− A tömörség ellenőrzést egyszerre csak egy visszacsapó szelepnél lehet elvégezni.", [137, 208, 400, 238],
                [[137, 208, 400, 223], [137, 223, 300, 238]]),
            box("2.3.", [81, 261, 101, 276]),
            box("− A blokkvezénylő és a helyszín közötti közvetlen telefonkapcsolat rendelkezésre áll", [137, 255, 418, 301],
                [[137, 255, 418, 270], [137, 270, 418, 285], [137, 286, 200, 301]]), box("O", [475, 274, 486, 289]),
            box("2.4.", [81, 301, 101, 316]), box("− A teszt végrehajtásához", [137, 302, 428, 317]),
            box("O", [475, 302, 486, 317])]
    b = DL._table_block(t, segs, 0, None, "Times New Roman", 10, False)
    assert [r[0] for r in _rows_text(b)] == [["2.2."], ["2.3."], ["2.4."]], _rows_text(b)
    assert _rows_text(b)[1][1] == ["− A blokkvezénylő és a helyszín közötti közvetlen telefonkapcsolat rendelkezésre áll"]


def test_code_without_text_does_not_stack_over_an_item_with_its_own_value():
    """1SZ01 л. 90 (прогон h13): у 11SA10P101 своего текста нет (значения ушли в шапку), у 11SA10P501 — своё «11KF11» и
    часть слитого «jelzés 11GB02 RM 2.1.». Весь текст 11SA10P501 полосу 11SA10P101 не накрывает — стопки нет:
    стопкой «11KF11» вставало против чужого кода."""
    import docx_layout as DL
    t = _grid1([72, 140, 228, 300, 400, 540], 110, 185)
    box = lambda text, bb, fr=None: {"id": 0, "page": 0, "text": text, "loc": {"bbox": bb, "size": 10, "bold": False,
                                                                              "align": "left", "frags": fr or [bb]}}
    segs = [box("11SA10P101", [78, 115, 137, 130]), box("11SA10P501", [78, 128, 137, 143]), box("11KF11", [157, 128, 195, 143]),
            box("jelzés 11GB02 RM 2.1. -- --", [316, 115, 386, 166], [[316, 115 + 13 * k, 386, 128 + 13 * k] for k in range(4)]),
            box("11SA10P801", [78, 141, 137, 156]), box("helyszín", [152, 141, 195, 156])]
    b = DL._table_block(t, segs, 0, None, "Times New Roman", 10, False)
    assert [r[0] for r in _rows_text(b)] == [["11SA10P101"], ["11SA10P501"], ["11SA10P801"]], _rows_text(b)


def test_continuation_without_code_does_not_take_next_item_number():
    """30THHAGT 4.23 (E9B_R2): строка продолжения пункта 4.22 без номера («30TH50B001 nyomása … bar ○ ○») задевает полосу
    пункта 4.23 со своим текстом — граница остаётся: номер «4.23.» против своего «Nyitni …», а не против чужого давления."""
    import docx_layout as DL
    t = _grid1([80, 120, 530, 560], 370, 460)
    box = lambda text, bb, fr=None: {"id": 0, "page": 0, "text": text, "loc": {"bbox": bb, "size": 12, "bold": False,
                                                                              "align": "left", "frags": fr or [bb]}}
    # геометрия листа 8 (orig): отметки «□ □» пункта 4.22 — один сегмент на две строки, задевает полосу 4.23
    segs = [box("4.22.", [85, 382, 109, 398]), box("Rögzíteni az alábbi nyomásértéket:", [125, 382, 297, 398]),
            box("30TH50B001 nyomása a 30TH50P101 alapján", [111, 409, 336, 426]),
            box("○ ○", [540, 411, 550, 444], [[540, 411, 550, 426], [540, 429, 550, 444]]),
            box("4.23.", [85, 429, 109, 445]), box("Nyitni a 30TH56S201, 30TH56S202-t", [125, 429, 312, 445])]
    b = DL._table_block(t, segs, 0, None, "Times New Roman", 10, False)
    assert _rows_text(b) == [
        [["4.22."], ["Rögzíteni az alábbi nyomásértéket:"], []],
        [[], ["30TH50B001 nyomása a 30TH50P101 alapján"], ["○ ○"]],
        [["4.23."], ["Nyitni a 30TH56S201, 30TH56S202-t"], []]], _rows_text(b)


def test_state_paragraph_does_not_glue_the_next_action_group():
    """3SZ17 л. 7 (E9B_R2): состояние «Armatúrák ZÁRVA …» на четыре строки свисает до следующей группы действий — кодов
    в ряду нет, граница остаётся: состояние не встаёт вровень с чужой командой «… NYITNI.»."""
    import docx_layout as DL
    t = _grid1([60, 320, 570], 265, 340)
    box = lambda text, bb, fr=None: {"id": 0, "page": 0, "text": text, "loc": {"bbox": bb, "size": 10, "bold": False,
                                                                              "align": "left", "frags": fr or [bb]}}
    # геометрия листа 6: состояние на две строки — своя полоса между группами действий, задевает следующую
    segs = [box("A 30UJ04D001 BÉNÍTANI, KISZAKASZOLNI.", [71, 274, 271, 290]),
            box("Armatúrák ZÁRVA, KIKAPCSOLVA, LETILTVA", [323, 286, 528, 312], [[323, 286, 528, 299],
                                                                                [323, 299, 528, 312]]),
            box("A 30UJ04S007,S008,S016 NYITNI.", [71, 303, 238, 320])]
    b = DL._table_block(t, segs, 0, None, "Times New Roman", 10, False)
    assert _rows_text(b) == [
        [["A 30UJ04D001 BÉNÍTANI, KISZAKASZOLNI."], []], [[], ["Armatúrák ZÁRVA, KIKAPCSOLVA, LETILTVA"]],
        [["A 30UJ04S007,S008,S016 NYITNI."], []]], _rows_text(b)


def test_items_with_own_text_stay_apart_under_a_hanging_paragraph():
    """1SZ01 л. 120 «Üzemviteli korlátozások» (E9B_R2): условия пункта 1 на пять строк свисают до пункта 2 — у каждого
    пункта свой номер и текст, граница остаётся (слитые «1.» и «2.» ставили условия пункта 1 против «Завершения»)."""
    import docx_layout as DL
    t = _grid1([60, 110, 250, 520], 90, 230)
    segs = [_seg("1.", 70, 100, 100), _seg("A gyorszárók felmelegítése", 120, 100, 240, lines=2),
            _seg("A felmelegítést akkor kell kezdeni, ha a gőznyomás p > 34 bar.", 260, 100, 510, lines=5),
            _seg("2.", 70, 154, 100), _seg("Felmelegítés befejezése", 120, 154, 240),
            _seg("A gyorszáró karimák külső falhőmérséklete nagyobb mint 170 °C.", 260, 154, 510)]
    b = DL._table_block(t, segs, 0, None, "Times New Roman", 10, False)
    assert [r[0] for r in _rows_text(b)] == [["1."], ["2."]], _rows_text(b)


def test_codes_with_own_values_stay_apart_under_a_merged_value():
    """1SZ01 л. 102, коды 11SE (E9B_R2): нарезка слила значения двух датчиков в один сегмент — у каждого кода своё место
    («helyszín»), граница остаётся: склейка ставила значения стопкой с первой строки — «16 бар» у чужого кода."""
    import docx_layout as DL
    t = _grid1([60, 160, 260, 400], 90, 200)
    segs = [_seg("11SE03P401", 70, 100, 150), _seg("helyszín", 170, 100, 250),
            _seg("11SE03P402", 70, 112, 150), _seg("helyszín", 170, 112, 250),
            _seg("1 ÷ 1,2 bar < 0,4 bar", 270, 100, 390, lines=2)]
    b = DL._table_block(t, segs, 0, None, "Times New Roman", 10, False)
    assert [r[0] for r in _rows_text(b)] == [["11SE03P401"], ["11SE03P402"]], _rows_text(b)


def test_item_number_stays_at_its_own_text_under_a_two_line_item():
    """01-02FMP0BIGT 4.13.1.2 (E9B_R2): текст пункта 4.13.1.1 в две строки задевает полосу 4.13.1.2 со своим текстом —
    граница остаётся: номер не встаёт против второй строки чужого пункта."""
    import docx_layout as DL
    t = _grid1([60, 130, 520], 90, 200)
    segs = [_seg("4.13.1.1", 70, 100, 120), _seg("- Tiszta kondenzátum beadása során a bórsav-koncentráció", 140, 100, 510,
                                                lines=2),
            _seg("4.13.1.2", 70, 120, 120), _seg("- Bármelyik mérés végrehajtása során a periódus", 140, 120, 510)]
    b = DL._table_block(t, segs, 0, None, "Times New Roman", 10, False)
    assert [r[0] for r in _rows_text(b)] == [["4.13.1.1"], ["4.13.1.2"]], _rows_text(b)


def test_marks_of_list_items_stay_with_their_items():
    """30RVRWTSBIGT л. 6 (E9B_R2): пункт «− 30YP25S601 …» в две строки задевает следующий пункт — кодов в первом столбце
    нет, у каждого пункта своя «O»: граница остаётся (стопкой вторая «O» вставала против «обесточен)»)."""
    import docx_layout as DL
    t = _grid1([60, 110, 400, 430], 90, 200)
    segs = [_seg("− 30YP25S601 mágnes szelep zárva (terhelő mágnes feszmentes)", 120, 100, 390, lines=2),
            _seg("O", 410, 100, 418), _seg("− 30YP25S201 lefúvató szakaszoló szelep", 120, 120, 390),
            _seg("O", 410, 120, 418)]
    b = DL._table_block(t, segs, 0, None, "Times New Roman", 10, False)
    assert [r[1:] for r in _rows_text(b)] == [
        [["− 30YP25S601 mágnes szelep zárva (terhelő mágnes feszmentes)"], ["O"]],
        [["− 30YP25S201 lefúvató szakaszoló szelep"], ["O"]]], _rows_text(b)


def test_merged_values_do_not_jump_over_a_single_value():
    """3PR10 л. 33 (E9B_R2): нарезка слила значения «4,5±0,25 bar 50-80 bar …» на несколько рядов, между ними — своё
    «20-270 °C» у YA12÷62T103P. У каждого кода своё место («BSZ-2B», «GA01», «BSZG») — ряды не сливаются, и «20-270 °C»
    остаётся у своего кода, а не после слитого абзаца."""
    import docx_layout as DL
    t = _grid1([60, 200, 300, 420], 90, 200)
    segs = [_seg("YA13÷63P411, 441, 471", 70, 100, 190), _seg("BSZ-2B", 210, 100, 290),
            _seg("YA12÷62T103P", 70, 112, 190), _seg("GA01", 210, 112, 290), _seg("20-270 °C", 310, 112, 410),
            _seg("YA13P001", 70, 124, 190), _seg("BSZG", 210, 124, 290),
            _seg("YA23P001", 70, 136, 190), _seg("BSZG", 210, 136, 290),
            _seg("4,5±0,25 bar 50-80 bar 50-80 bar", 310, 100, 410, lines=4)]
    b = DL._table_block(t, segs, 0, None, "Times New Roman", 10, False)
    got = _rows_text(b)
    assert [r[0] for r in got] == [["YA13÷63P411, 441, 471"], ["YA12÷62T103P"], ["YA13P001"], ["YA23P001"]], got
    assert got[1][2] == ["20-270 °C"], got


def test_sub_rows_marks_of_items_with_own_text_stay_apart():
    """30RVRYTSBIGT_TU 4.2.3.8 (E9B_R2), путь разметки: абзац «BLÜGY …» в две строки задевает следующий пункт со своей
    «O» — меток (кодов) в ряду нет, подстроки остаются: обе «O» стопкой оставляли следующий пункт без отметки."""
    import docx_struct as D
    cell = lambda *bl: {"blocks": list(bl), "th": False, "geom": {"page": 1, "top": 100, "bottom": 160, "x0": 70,
                                                                  "x1": 100}}
    two = _p("BLÜGY ellenőrzi a GFDP működését", 110, 120)
    two["geom"] = dict(two["geom"], bottom=131.5, lines=2)                 # задевает следующий пункт на 1,5 pt
    row = [cell(_p("4.2.3.8.", 100, 70)), cell(_p("Megjegyzés:", 100, 120), two, _p("A lefutás után", 130, 120)),
           cell(_blank(), _p("O", 110, 480), _p("O", 130, 480))]
    out = D.sub_rows([{"k": "table", "rows": [row]}])[0]["rows"]
    text = lambda c: [D.text_of(x["runs"]) for x in c["blocks"]]
    assert [[text(c) for c in r][1:] for r in out] == [
        [["Megjegyzés:"], []], [["BLÜGY ellenőrzi a GFDP működését"], ["O"]], [["A lefutás után"], ["O"]]], \
        [[text(c) for c in r] for r in out]
