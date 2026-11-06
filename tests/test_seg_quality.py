# -*- coding: utf-8 -*-
"""Нарезка по просмотру глазами (NIGHT2_EYES пп. 12, 14; E6 «нет перевода» у TU): правило и контрпример на каждое.
1. Перечень кодов, растянутый выключкой по ширине, — одна строка абзаца, а не ячейки перечня.
2. Значок легенды («□ sikerességi kritérium») и маркер «−» (U+2212) начинают свой абзац.
3. Просвет двух колонок текста — не между столбцами таблицы с линейками рядов."""
from pathlib import Path

import fitz

import layout as LY

# шрифт с «−» (U+2212) и «ő»: у базовых шрифтов PDF их нет (как в test_audit_out)
FONT = next((f for f in (r"C:\Windows\Fonts\arial.ttf", "/usr/share/fonts/truetype/dejavu/DejaVuSans.ttf")
             if Path(f).exists()), None)


def _fr(t, bb, size=9.0, bold=False):
    return {"t": t, "bbox": list(bb), "size": size, "bold": bold, "family": "serif"}


def _texts(frags, hs=(), vs=()):
    return [p["text"] for p in LY.group(frags, list(hs), list(vs))]


# --- 1. перечень кодов в абзаце с выключкой по ширине --------------------------------------------------------------

def _mono_page():
    """Лист моноширинного текста (Courier 11): абзац «код + условие», продолжение — строки кодов через запятую,
    растянутые выключкой по ширине (просвет 22 pt: CODE_GAP 8 ≤ просвет ≤ GAP 30), как 02VE25 л. 44, раздел 5.2;
    ниже — перечень «код ⇥ значение» без запятых (два столбца без линеек)."""
    d = fitz.open()
    pg = d.new_page(width=595, height=842)
    put = lambda x, y, t: pg.insert_text((x, y), t, fontname="cour", fontsize=11)
    put(70.8, 100, "02TD96S202 Nyitási feltétel, 4. VT útvonal kiválasztva, 30TD97S201,")
    for row, codes in ((114.4, ("30TD97S202,", "30TD98S201,", "30TD98S202,", "40TD97S201,", "40TD97S202,")),
                       (128.8, ("40TD98S201,", "40TD98S202,", "40TD98S203,", "40TD98S204,", "02TD97S601,"))):
        for k, c in enumerate(codes):
            put(70.8 + 95.3 * k, row, c)
    put(70.8, 143.2, "02TD98S601 zárva.")
    for k, (code, val) in enumerate((("TK41D001", "KM2/1"), ("TK42D001", "KM2/2"), ("TK43D001", "KM2/3"))):
        put(70.8, 300 + 24 * k, code)
        put(70.8 + 95.3, 300 + 24 * k, val)
    return d, pg


def test_justified_code_enumeration_is_one_line_of_its_paragraph():
    """Строка «30TD97S202,  30TD98S201,  …» — продолжение фразы (запятая: перечисление идёт дальше), пробелы растянула
    выключка: одним куском и одним абзацем с началом «02TD96S202 Nyitási…» и концом «… zárva.». Прежде правило «код ⇥
    значение» (CODE_GAP) резало её на пять «ячеек», и коды разлетались по сегментам вразнобой (в переводе — строки,
    начинающиеся с середины, NIGHT2_EYES п. 14)."""
    d, pg = _mono_page()
    fr = LY.frags_from_text(pg)
    line = [f["t"] for f in fr if f["bbox"][1] < 110 < f["bbox"][3] and f["t"].startswith("30TD97S202")]
    assert line and line[0].split() == ["30TD97S202,", "30TD98S201,", "30TD98S202,", "40TD97S201,", "40TD97S202,"]
    hs, vs = LY.rules(pg)
    pars = [p["text"] for p in LY.group(fr, hs, vs) if p["bbox"][1] < 200]
    assert pars == ["02TD96S202 Nyitási feltétel, 4. VT útvonal kiválasztva, 30TD97S201, 30TD97S202, 30TD98S201, 30TD98S202, "
                    "40TD97S201, 40TD97S202, 40TD98S201, 40TD98S202, 40TD98S203, 40TD98S204, 02TD97S601, "
                    "02TD98S601 zárva."]
    # контрпример: «код ⇥ значение» без запятой — ячейки перечня, каждая своим куском
    low = sorted(f["t"] for f in fr if f["bbox"][1] > 250)
    assert low == ["KM2/1", "KM2/2", "KM2/3", "TK41D001", "TK42D001", "TK43D001"]


def test_justified_code_enumeration_on_scan_words():
    """То же для слов Tesseract (ocr_frags): «30TD97S202,» ⇥ «30TD98S201,» — одна строка, «TK41D001» ⇥ «KM2/1» — две."""
    import segments as S
    w = lambda t, x, line: {"t": t, "box": [x, 100 + 14 * line, x + 70, 110 + 14 * line], "h": 10, "c": 95, "ink": 0.2,
                            "stroke": 0.8, "line": (1, 1, line), "key": (1, 1)}
    ws = [w("30TD97S202,", 70, 0), w("30TD98S201,", 165, 0), w("30TD98S202,", 260, 0),
          w("TK41D001", 70, 3), w("KM2/1", 165, 3)]
    got = sorted(f["t"] for f in S.ocr_frags(ws))
    assert got == ["30TD97S202, 30TD98S201, 30TD98S202,", "KM2/1", "TK41D001"]


# --- 2. значок легенды и маркер «−» начинают абзац ------------------------------------------------------------------

def test_legend_glyph_line_is_not_glued_to_heading_above():
    """Шапка бланка EPM (30PR132GT л. 35): заголовок по центру рамки, под ним легенда «□ sikerességi kritérium» /
    «O nem sikerességi kritérium» от левого края. Строка, начинающаяся значком символьного шрифта (Wingdings ,
    Юникод □ ○ …), — пункт легенды: свой абзац. Прежде она приклеивалась к заголовку («… (EPM) — критерий
    успешности»), и значок пропадал: портал его не возвращает, а вернуть его на место можно только в начале сегмента
    (segments.keep_lead_glyph)."""
    head = _fr("ELLENŐRIZENDŐ PARAMÉTEREK, MŰKÖDÉSEK (EPM)", (183.1, 91.4, 442.3, 104.8), 9.96)
    no = _fr("O nem sikerességi kritérium", (83.8, 114.4, 203.9, 127.9), 9.96)
    frame = ([(66.8, 556.2, 86.0), (66.8, 556.2, 133.3)], [(86.4, 133.0, 66.8), (86.4, 133.0, 556.2)])
    for glyph in ("", "□"):
        leg = _fr(glyph + " sikerességi kritérium", (83.8, 102.9, 184.2, 116.3), 9.96)
        assert _texts([head, leg, no], *frame) == [head["t"], leg["t"], no["t"]]
    # контрпример: значок посреди строки абзац не рвёт, обычное продолжение фразы клеится как прежде
    a = _fr("A jelölések: □ sikerességi kritérium, ○ nem sikerességi kritérium, a többi", (70, 200, 520, 212))
    b = _fr("jelölés a táblázat végén található.", (70, 211, 260, 223))
    assert _texts([a, b]) == [a["t"] + " " + b["t"]]


def test_minus_sign_marker_starts_list_item():
    """Пункты списка с маркером «−» (U+2212, знак минуса символьного шрифта Word) в ячейке таблицы (TU пп. 1.16–1.17:
    «− 30YP15S101 zárva (…)» / «− 30YP13S101, S102 zárva (…)» …) — каждый свой абзац, как с «-» и «–». Прежде строки,
    длиннее половины ячейки, склеивались в один сегмент страницы; дерево разметки делит их по пунктам — пар у пунктов не
    было, и у старого JSON 14 пунктов остались в DOCX по-венгерски («нет перевода», E6)."""
    items = ["− 30YP15S101 zárva (lámpajelzés, analóg állásjelzés)",
             "− 30YP13S101, S102 zárva (terhelő mágnes feszültség alatt)",
             "− 30YP13S201, S202 zárva, kikapcsolva (lámpajelzés, analóg állásjelzés)"]
    fr = [_fr(t, (131.4, 360.9 + 11 * k, x1, 373.0 + 11 * k)) for k, (t, x1) in enumerate(zip(items, (332.2, 355.7, 402.7)))]
    marks = [_fr("O", (510.2, 360.6 + 11 * k, 519.9, 374.0 + 11 * k), 9.96) for k in range(3)]
    vs = [(100.1, 765.7, 100.1), (100.1, 765.7, 474.96), (100.1, 765.7, 550)]
    got = _texts(fr + marks, vs=vs)
    assert [t for t in got if t != "O"] == items
    # контрпример: минус перед числом («−5 °C») — не маркер, фраза продолжается
    a = _fr("A hőmérséklet a próba teljes ideje alatt sem csökkenhet a megadott érték alá, vagyis", (70, 200, 520, 212))
    b = _fr("−5 °C alá, különben a szivattyú leáll.", (70, 211, 260, 223))
    assert _texts([a, b]) == [a["t"] + " " + b["t"]]


def test_minus_sign_with_gap_stays_with_item_text():
    """Маркер «−» отдельным куском строки PDF (отступ до текста, как в TU пп. 1.16–1.17) — в одном сегменте с текстом
    пункта, и каждый пункт — свой абзац: дерево разметки Word отдаёт пункт целиком («− 30YP15S101 zárva (…)»), и сборка
    DOCX по разметке ищет перевод именно этого текста (docx_build.missing). Отдельный кусок «−» (как «-», «–» — BULLET)
    давал текст пункта без знака: пары с деревом всё равно не было (76 → 81 текстов без пары у TU X), а сборка по
    сегментам листа «−» маркером не знает (patterns.MARKERS) — вышел бы абзац из одного «−»."""
    d = fitz.open()
    pg = d.new_page(width=595, height=842)
    items = ("30YP15S101 zárva (lámpajelzés, analóg állásjelzés), a próba előtt ellenőrizni kell",
             "30YP13S101, S102 zárva (terhelő mágnes feszültség alatt), a próba előtt is")
    for k, t in enumerate(items):
        pg.insert_text((131, 100 + 11 * k), "−", fontname="F0", fontfile=FONT, fontsize=9)
        pg.insert_text((145, 100 + 11 * k), t, fontname="F0", fontfile=FONT, fontsize=9)
    fr = LY.frags_from_text(pg)
    got = [" ".join(p["text"].split()) for p in LY.group(fr, [], [])]
    assert got == ["− " + t for t in items]


# --- 3. просвет колонок текста — не между столбцами таблицы ---------------------------------------------------------

def _par(x0, y0, x1, y1, t, n=1):
    return {"bbox": [x0, y0, x1, y1], "frags": [{"bbox": [x0, y0 + k * (y1 - y0) / n, x1, y0 + (k + 1) * (y1 - y0) / n],
                                                  "t": t} for k in range(n)], "text": t}


def test_gutter_is_not_found_between_columns_of_a_ruled_table():
    """Бланк EPM (30PR132GT л. 25–62, 30PR25GT; так же «Tevékenység | Állapot» 1SZ01, 3SZ17, 3SZ23, 3PR10): уровни
    перечня в ячейке «Megnevezés» («30TK52» слева, коды справа) и столбцы таблицы стоят на разной высоте — просвет между
    ними принимался за просвет двух колонок сплошного текста (E46). Невидимая линейка просвета резала ячейку, абзацы
    шли «колонка за колонкой» (все «30TK52» листа, потом все коды), ячейка «30TK52» — [123,6; 237,5] вместо столбца
    таблицы. Таблица: линейки рядов поперёк просвета, над ними и под ними — абзацы колонок (не меньше трёх; у колонок
    текста — только черта колонтитула вверху и внизу)."""
    head = _par(76, 38, 520, 68, "Azonosító: 30PR132GT Megnevezés: Primerkör 132 bar-os tömörségi próbája Oldalszám: 35.")
    title = _par(183, 91, 442, 105, "ELLENŐRIZENDŐ PARAMÉTEREK, MŰKÖDÉSEK (EPM)")
    left = [_par(84, 103, 184, 116, "□ sikerességi kritérium"), _par(84, 114, 204, 128, "O nem sikerességi kritérium"),
            _par(83, 136, 110, 149, "Sorsz.")] + [_par(168, 172 + 62 * k, 200, 184 + 62 * k, "30TK5%d" % k) for k in range(8)]
    right = [_par(269, 136, 321, 149, "Megnevezés"), _par(489, 136, 533, 149, "Ellenőrzés")]
    for k in range(40):
        y = 182 + 12.4 * k
        right += [_par(394, y, 445, y + 12, "30TK45S%03d" % k), _par(507, y, 515, y + 12, "O")]
    pars = [head, title] + left + right + [_par(300, 718, 550, 730, "A lapot kitöltötte (dátum, aláírás): ....")]
    hs = [(67.2, 555.8, 86.0), (124.0, 463.4, 133.3), (124.0, 463.4, 151.6), (67.2, 555.8, 715.7), (67, 556, 745.3)]
    assert LY.gutter(pars, 595) is not None                 # без линеек — как прежде (правило E46)
    assert LY.gutter(pars, 595, hs) is None
    # контрпример: две колонки сплошного текста с чертой колонтитула вверху и внизу — колонки
    top, foot = _par(57, 20, 300, 30, "Biztonság Kezelési kézikönyv"), _par(57, 812, 540, 822, "MG16D547 Danfoss A/S 5")
    cols = [_par(57, 80 + 120 * k, 290, 180 + 120 * k, f"bal {k}", 8) for k in range(5)] + \
        [_par(310, 95 + 120 * k, 540, 195 + 120 * k, f"jobb {k}", 8) for k in range(5)]
    g = LY.gutter([top] + cols + [foot], 595, [(57, 540, 35.0), (57, 540, 805.0)])
    assert g is not None and 290 < g < 310


def test_line_overflowing_column_rule_is_full_and_phrase_continues():
    """Моноширинная строка ячейки кончается на 1,8 pt правее линии столбца (последний знак «3» заходит за линию, 1SZ01
    л. 23): до края ячейки места нет — строка полная, следующая строка того же столбца продолжает абзац. Прежде линия не
    считалась пределом справа (она левее конца строки), строка выходила «короткой», и фраза рвалась («Az 1*RA11,12,14
    turbina gőzvezetékek, az S203» / «és az SE10S801÷S804 közötti szakasz …»); на листах 1SZ01 это прятал ложный
    просвет колонок (п. 3). Контрпример: строка, до линии которой далеко, абзац закрывает, как прежде."""
    a = _fr("Az 1*RA11,12,14 turbina gőzvezetékek, az S203", (74.4, 359.5, 324.4, 374.8), 9.24)
    b = _fr("és az SE10S801÷S804 közötti szakasz nyomásmen-", (74.4, 370.1, 330.0, 385.4), 9.24)
    b["bbox"][2] = 316.0
    c = _fr("tességét ELLENŐRIZNI.", (74.4, 380.7, 190.0, 396.0), 9.24)
    right = _fr("Az alábbi armatúrák ZÁRVA, a turbina gőzvezetékek nyomásmentesek.", (326.0, 359.5, 565.0, 374.8), 9.24)
    vs = [(290, 361, 322.6), (361, 453, 322.6), (290, 453, 74.0), (290, 453, 566.0)]
    got = _texts([a, b, c, right], vs=vs)
    assert "Az 1*RA11,12,14 turbina gőzvezetékek, az S203 és az SE10S801÷S804 közötti szakasz nyomásmentességét " \
           "ELLENŐRIZNI." in got
    short = _fr("Az 1*RA11,12,14 turbina", (74.4, 359.5, 200.0, 374.8), 9.24)
    nxt = _fr("Új bekezdés a táblázat alatt.", (74.4, 370.1, 230.0, 385.4), 9.24)
    got = _texts([short, nxt, right], vs=vs)
    assert short["t"] in got and nxt["t"] in got


def test_paragraph_does_not_skip_a_line_of_its_column():
    """Строки абзаца идут подряд: строка не продолжает абзац, если между его последней строкой и ею в том же столбце
    стоит другая строка. 3SZ17 л. 9 (моноширинный, рамки строк 15–18 pt при шаге 10): «A 31,32SZ93S004 háromutú szelepet
    a» (выступает за линию столбца — полная), «30UJ97B001 felé NYITNI.», «A 30UJ05,06,07,08S201 NYITNI.» — третья
    приклеивалась к первой через вторую («…szelepet a A 30UJ05…»), окно кандидатов near доставало строку через одну.
    Контрпример: три строки одной фразы подряд склеиваются, как прежде."""
    l1 = _fr("A 31,32SZ93S004 háromutú szelepet a", (71.0, 152.0, 321.0, 167.0), 9.0)
    l2 = _fr("30UJ97B001 felé NYITNI.", (71.0, 160.0, 201.0, 178.0), 9.0)
    l3 = _fr("A 30UJ05,06,07,08S201 NYITNI.", (71.0, 171.0, 233.0, 188.0), 9.0)
    right = _fr("A mosóvíz habmentes.", (323.0, 193.0, 436.0, 208.0), 9.0)
    vs = [(96, 480, 319.0), (96, 480, 71.0), (96, 480, 562.0)]
    got = _texts([l1, l2, l3, right], vs=vs)
    assert l3["t"] in got and not any("szelepet a A 30UJ05" in t for t in got)
    a = _fr("A turbina védelmi hatáslánc armatúrái zárnak, és a", (71.0, 152.0, 300.0, 167.0), 9.0)
    b = _fr("blokk teljesítménye csökken, a kezelő a táblán", (71.0, 162.0, 300.0, 177.0), 9.0)
    c = _fr("ellenőrzi az állapotot.", (71.0, 172.0, 180.0, 187.0), 9.0)
    assert _texts([a, b, c], vs=vs) == [a["t"] + " " + b["t"] + " " + c["t"]]


def test_code_enumeration_continues_after_short_line():
    """Перечень кодов через запятую в ячейке (1SZ01 л. 20, под «ZÁRNAK:»): строки при обычном шаге с одного края —
    один абзац, короткая строка «1*RB11S801, RB12S801,» его не закрывает (запятая: перечень идёт дальше). Прежде шёл
    кусками по 1–2 строки. Контрпример: коды через пустую строку (ряды псевдотаблицы «02TD95S201,» / «02TD95S202,»,
    02VE25 л. 44, шаг 24 pt при кегле 11) — каждый свой абзац."""
    rows_ = ["1*SE10S801, S802, S803, S804,", "1*SE10S401, S402, S403, S404,", "1*RB11S801, RB12S801,",
             "1*RA11S203, RA12S203, RA14S203, RA15S201,", "1*RM21S201."]
    ends = (470, 470, 435, 540, 385)
    fr = [_fr(t, (325, 300 + 10.3 * k, x1, 315 + 10.3 * k), 9.2) for k, (t, x1) in enumerate(zip(rows_, ends))]
    vs = [(280, 420, 320.0), (280, 420, 562.0)]
    assert _texts(fr, vs=vs) == [" ".join(rows_)]
    codes = ["02TD95S201,", "02TD95S202,", "02TD96S201,"]
    fr = [_fr(c, (70.8, 270.7 + 24.3 * k, 143.4, 289.5 + 24.3 * k), 11.04) for k, c in enumerate(codes)]
    assert _texts(fr) == codes
