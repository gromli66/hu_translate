# -*- coding: utf-8 -*-
"""Нарезка: слово неделимо (E4, lines.py) и ряд — единица таблицы (E5, rows.py). Таблица «строка × строка → склеено/нет»
и контрпример на каждое правило (UNIFIED_DESIGN.md, П3, П1–П2, П9)."""
import pytest

import lines as LN


@pytest.mark.parametrize("prev, nxt, want", [
    # перенос слова: строка кончается буквой и «-», продолжение со строчной — без дефиса и пробела
    ("végre-", "hajtani", "végrehajtani"),
    ("написан акку-", "ратно от руки", "написан аккуратно от руки"),
    # дефис — часть слова: после прописной или цифры, составное через дефис
    ("ÜFK-", "s rendszer", "ÜFK-s rendszer"),
    ("a 30-", "as szelep", "a 30-as szelep"),
    ("АЭС-", "ов", "АЭС-ов"),
    ("loss-of-", "coolant", "loss-of-coolant"),
    # дефис перед союзом — не перенос, в том числе союз — последнее слово (слово Tesseract)
    ("fűtő-", "és hűtőrendszer", "fűtő- és hűtőrendszer"),
    ("fűtő-", "és", "fűtő- és"),
    ("FŰTŐ-", "ÉS HŰTŐ", "FŰTŐ- ÉS HŰTŐ"),
    # слово из прописных: одна из частей от четырёх букв — перенос; пара коротких аббревиатур — составная
    ("A tábla KI-", "KAPCSOLVA feliratú", "A tábla KIKAPCSOLVA feliratú"),
    ("Az eredmény REND-", "BEN.", "Az eredmény RENDBEN."),
    ("RO-", "ROP", "RO- ROP"),
    # не перенос: продолжение с прописной у слова из строчных, длинное тире, строка без дефиса
    ("a mosatás-", "Ürítés", "a mosatás- Ürítés"),
    ("a próba –", "folytatás", "a próba – folytatás"),
    ("a próba", "folytatás", "a próba folytatás"),
])
def test_join_two_lines(prev, nxt, want):
    assert LN.join_lines([prev, nxt]) == want


def test_glue_is_the_only_rule_for_tesseract_words():
    """Слова Tesseract склеиваются тем же правилом, что строки слоя: прежде «fűtő-» + «és» → «fűtőés», «KI-» +
    «KAPCSOLVA» не склеивалось, «loss-of-» + «coolant» терял дефис."""
    import segments as S
    w = lambda t, line, key=(1, 1): {"t": t, "line": line, "key": key}
    got = lambda ws: [x["t"] for x in S._dehyphen(ws)]
    assert got([w("végre-", (1, 1, 1)), w("hajtani", (1, 1, 2))]) == ["végrehajtani"]
    assert got([w("fűtő-", (1, 1, 1)), w("és", (1, 1, 2))]) == ["fűtő-", "és"]
    assert got([w("KI-", (1, 1, 1)), w("KAPCSOLVA", (1, 1, 2))]) == ["KIKAPCSOLVA"]
    assert got([w("ÜFK-", (1, 1, 1)), w("s", (1, 1, 2))]) == ["ÜFK-s"]
    # та же строка или другой абзац Tesseract — не перенос
    assert got([w("végre-", (1, 1, 1)), w("hajtani", (1, 1, 1))]) == ["végre-", "hajtani"]
    assert got([w("végre-", (1, 1, 1)), w("hajtani", (2, 1, 1), key=(2, 1))]) == ["végre-", "hajtani"]


def test_layout_uses_the_same_functions():
    import layout as LY
    assert LY.join_lines is LN.join_lines and LY.hyphen_break is LN.hyphen_break


# --- E5: ряд — единица таблицы (rows.py) ---------------------------------------------------------------------------

def _fr(t, bb, size=9.96, bold=False):
    return {"t": t, "bbox": list(bb), "size": size, "bold": bold, "family": "mono"}


def _texts(frags, hs=(), vs=()):
    import layout as LY
    return sorted(p["text"] for p in LY.group(frags, list(hs), list(vs)))


def test_values_column_is_cut_by_rows_of_codes_column():
    """INFO-таблица (3SZ23 л. 24): значения столбца без кодов против столбца кодов — каждое своим сегментом, хотя ближайший
    кусок слева («GÉPHÁZ … 28KÁ») не якорь; рамки строк 17 pt при шаге 11,4. Прежде «-3,6m helyszín … 0,0m helyszín» —
    шесть рядов одним сегментом. Без столбца кодов — правило молчит."""
    import rows as RW
    ys = [142.0, 153.4, 164.8, 176.2]
    codes = [_fr(c, (114, y, 180, y + 17)) for c, y in zip(("30UJ03S601", "31UJ20S601", "31UJ21S601", "31UJ21S602"), ys)]
    place = [_fr('GÉPHÁZ "A" sor 28KÁ', (234, y, 354, y + 17)) for y in ys]
    vals = [_fr(v, (392, y, 482, y + 17)) for v, y in zip(("-3,6m helyszín", "0,0m helyszín", "0,0m helyszín",
                                                             "9,6m helyszín"), ys)]
    got = _texts(codes + place + vals, vs=[(130, 200, x) for x in (75, 220, 361, 506, 540)])     # линейки столбцов
    assert got.count("0,0m helyszín") == 2 and "-3,6m helyszín" in got and "9,6m helyszín" in got
    rw = RW.Rows(codes + place + vals, [])
    assert rw.apart(vals[0], vals[1]) == "якорь" and rw.apart(vals[1], vals[2]) == "якорь"
    assert RW.Rows(place + vals, []).apart(vals[0], vals[1]) is None


def test_row_rule_hidden_under_tall_frames_splits_title_from_column_header():
    """Название таблицы и шапка столбца под ним (3PR10 INFO): линейка ячейки шапки (y 84,5) лежит внутри обеих рамок
    (70,6–87,5 и 83,4–100,4) — прежде выпадала из проверки, и «Segédenergiával működtetett armatúrák Beépítés helye» шли
    одним сегментом. Линейка обрамляет «Beépítés helye» с обеих сторон — граница, хотя она в пределах ширины названия."""
    import rows as RW
    a = _fr("Segédenergiával működtetett armatúrák", (178.1, 70.6, 396.3, 87.5), bold=True)
    b = _fr("Beépítés helye", (249.2, 83.4, 331.7, 100.4), bold=True)
    assert RW.Rows([a, b], [(220.3, 360.6, 84.5)]).apart(a, b) == "линейка"
    assert RW.Rows([a, b], []).apart(a, b) is None


def test_underline_and_line_highlight_are_not_row_rules():
    """Подчёркивание слова в строке и край заливки строки — в пределах ширины текста и рамки своей строки: абзац не рвут
    (3FI18 л. 8: строки с заливкой, края — между строками абзаца). Та же линейка шире текста — граница ряда."""
    import rows as RW
    a = _fr("A tűzjelzést az adott helyiséghez tartozó", (100, 100, 400, 117))
    b = _fr("piros led jelzi az oltóközpontban.", (100, 111, 330, 128))
    assert len(_texts([a, b], hs=[(150, 260, 113)])) == 1                # подчёркивание части строки a
    assert len(_texts([a, b], hs=[(100, 400, 112)])) == 1                # край заливки строки a
    assert len(_texts([a, b], hs=[(90, 420, 112)])) == 2                 # линейка ряда шире текста
    assert RW.rule_between([(90, 420, 112)], a["bbox"], b["bbox"], soft=False)
    # подчёркнутый заголовок по центру в две строки (30PR132GT л. 42): подчёркивание первой строки шире второй, но
    # выровнено с текстом первой — одна фраза
    h1 = _fr("Reaktor és Készülék Karbantartó Osztály ellenőrzési", (155, 162, 430, 179), bold=True)
    h2 = _fr("terjedelme", (264, 176, 321, 192), bold=True)
    assert RW.Rows([h1, h2], [(155.3, 426.8, 176.5)]).apart(h1, h2) is None
    # серая полоса под каждой строкой заголовка (3SZ23 л. 18): сторона заливки без обводки — не линейка ряда
    import layout as LY
    t1 = _fr("10.1 UJ21 vízköddel oltó berendezések üzemeltetése N2 nyomás alacsony jelzés", (74, 80, 536, 99))
    t2 = _fr("esetén", (74, 91, 110, 110))
    edge = LY._mark((74.4, 559.9, 95.0), False, edge=True)
    assert RW.Rows([t1, t2], [edge]).apart(t1, t2) is None
    assert RW.Rows([t1, t2], [(74.4, 559.9, 95.0)]).apart(t1, t2) == "линейка"


def test_codes_inside_prose_column_are_not_row_anchors():
    """Коды строками внутри текста столбца «Tevékenység» (3SZ23 л. 6: «31(32)UJ21S603» под фразой) — содержание, а не
    якоря: якорь — только в столбце, где номера и коды преобладают. Абзац третьего столбца против них не рвётся."""
    import rows as RW
    left = [_fr("Az alábbi armatúrák közül az oltási", (74, 432, 319, 449)),
            _fr("helyhez tartozók NYITNAK:", (74, 443, 260, 460)),
            _fr("31(32)UJ21S603", (74, 454, 164, 471)), _fr("31(32)UJ21S604", (74, 466, 164, 483)),
            _fr("Az UJ01/1 CERBERUS kezelőn a jelzéseket", (74, 520, 319, 537)),
            _fr("nyugtázni a műszakvezetővel.", (74, 531, 260, 548))]
    mid = [_fr("lásd a táblázatban", (200, 454, 300, 471)), _fr("lásd a táblázatban", (200, 466, 300, 483))]
    right = [_fr("Az oltási körhöz tartozó mágnes szelep", (326, 454, 566, 471)),
             _fr("nyitva, az oltás rendben működik.", (326, 466, 540, 483))]
    rw = RW.Rows(left + mid + right, [])
    assert rw.anchors == [] and rw.apart(right[0], right[1]) is None
    assert "Az oltási körhöz tartozó mágnes szelep nyitva, az oltás rendben működik." in _texts(left + mid + right)


def test_anchor_centred_in_multiline_cell_does_not_split_it():
    """Номер пункта по центру ячейки в три строки («4.1.9.» против второй строки, 30HU375GT л. 12) строку не делит: у первой
    строки своего якоря нет, а зазор между строками — обычный шаг, не зазор рядов."""
    import rows as RW
    nums = [_fr("4.1.8.", (84, 440, 109, 453)), _fr("4.1.9.", (84, 486, 109, 499)), _fr("4.1.10.", (84, 530, 112, 543))]
    lines = [_fr("Amennyiben olyan hibát észleltek, amely mellett a", (123, 474, 439, 488)),
             _fr("próba nem végezhető el, abban az esetben a hibát", (123, 486, 438, 499)),
             _fr("el kell hárítani.", (123, 497, 352, 511))]
    rw = RW.Rows(nums + lines, [])
    assert rw.apart(lines[0], lines[1]) is None and rw.apart(lines[1], lines[2]) is None


def test_rows_with_gap_split_even_where_only_first_line_has_anchor():
    """Ряды «YA-K7 ⇥ …» / «YA-K8 ⇥ …» (3PR10 л. 23): у второй строки ячейки своего якоря нет — она в ряду якоря выше; у
    следующей строки — свой якорь, и между ними зазор рядов (23 pt при шаге 12): новый ряд. Шапка над первым рядом
    («Elosztó helye» над «E601/II», 3FI18 л. 38) — тоже через зазор."""
    import rows as RW
    anc = [_fr("YA-K7", (71, 95, 107, 112)), _fr("YA-K8", (71, 130, 101, 147)), _fr("YA-K9", (71, 165, 107, 182))]
    txt = [_fr("Megengedett a hidegvizes indítás 0,8 U névl tápfeszültség", (124, 95, 457, 112)),
           _fr("esetén, melegvizes indítás 0,75 Unévl tápfesz. esetén, és a", (124, 107, 446, 124)),
           _fr("tilos a reaktor indítása, ha merev a primerkör és a", (124, 130, 466, 147)),
           _fr("térfogatkompenzátorban nincs gőz vagy nitrogén párna.", (124, 143, 444, 160))]
    rw = RW.Rows(anc + txt, [])
    assert rw.apart(txt[1], txt[2]) == "якорь" and rw.apart(txt[0], txt[1]) is None and rw.apart(txt[2], txt[3]) is None
    assert len(_texts(anc + txt, vs=[(85, 190, x) for x in (66, 118, 480)])) == 3 + 2
    head = [_fr("Elosztó", (698, 111, 746, 127)), _fr("helye", (698, 122, 734, 139))]
    vals = [_fr("E601/II", (698, 140, 746, 157)), _fr("E602/II", (698, 153, 746, 170))]
    codes = [_fr("30UJ47S201", (380, 140, 445, 157)), _fr("30UJ48S201", (380, 153, 445, 170))]
    rw = RW.Rows(head + vals + codes, [])
    assert rw.apart(head[1], vals[0]) == "якорь" and rw.apart(head[0], head[1]) is None


def test_form_marks_never_merge():
    """Столбик отметок (○ □ O ✓ ☒ ◉) — каждая отметка своя ячейка, к соседней строке не приклеивается."""
    marks = [_fr(m, (500, 100 + 11 * i, 508, 113 + 11 * i)) for i, m in enumerate("○○☒")]
    assert len(_texts(marks)) == 3


def test_portal_label_does_not_merge_across_row_boundary():
    """Метка портала «продолжение» склеивает абзацы правил, но не через границу ряда (apart): иначе портал сводил на нет
    нарезку по рядам. Сторона белой подложки — не граница: склейка по метке через неё остаётся (30RL2126GT)."""
    import layout as LY
    import roles as R
    import rows as RW
    from lines import join_lines
    a = _fr("Az oltási körhöz tartozó mágnes szelep", (326, 100, 566, 112))
    b = _fr("nyitva, az oltás rendben működik.", (326, 113, 540, 125))
    pars = lambda: [{"frags": [a], "bbox": a["bbox"], "size": 9.96, "bold": False},
                    {"frags": [b], "bbox": b["bbox"], "size": 9.96, "bold": False}]
    lines = [{"id": 0, "x0": 326, "x1": 566, "top": 100, "bottom": 112},
             {"id": 1, "x0": 326, "x1": 540, "top": 113, "bottom": 125}]
    labels = {0: {"new": True, "role": "body"}, 1: {"new": False, "role": "body"}}
    assert len(R.apply(pars(), lines, labels, join_lines)) == 1
    assert len(R.apply(pars(), lines, labels, join_lines, apart=RW.Rows([a, b], [(300, 600, 112.5)]).apart)) == 2
    soft = LY.Rule((326, 566, 112.5))
    soft.soft = True
    assert len(R.apply(pars(), lines, labels, join_lines, apart=RW.Rows([a, b], [soft]).apart)) == 1


def test_wide_row_keeps_cell_continuation():
    """Строка ячейки в широком ряду (десяток столбцов, 3FI18 л. 35) находит свой абзац: прежде кандидатами были только 12
    последних абзацев, и «29A-30 KÁ» отрывалась от «E412/II. csőfolyosó ’V’ sor»."""
    cells = [_fr(f"érték{i}", (200 + 45 * i, 100, 240 + 45 * i, 117)) for i in range(14)]
    cell = [_fr("E412/II.", (110, 100, 164, 117)), _fr("csőfolyosó", (110, 113, 170, 130))]
    more = [_fr(f"tovább{i}", (200 + 45 * i, 113, 240 + 45 * i, 130)) for i in range(14)]
    vs = [(95, 135, 197 + 45 * i) for i in range(15)]                    # линейки столбцов
    assert "E412/II. csőfolyosó" in _texts(cell + cells + more, vs=vs)
