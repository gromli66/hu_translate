# -*- coding: utf-8 -*-
"""Нарезка, второй круг (остатки SEG_QUALITY_REPORT, EYES_2026-11-06 N5, E9.7г): правило и контрпример на каждое.
1. Фраза продолжается с кода или числа после запятой или служебного слова; голое число со словом — количество, не номер.
2. Строка, где жирное — только часть, абзац начертанием не рвёт; куски одной строки — по линиям, не по округлённой высоте.
3. Пустая строка — граница и для меток портала (кроме фразы, что продолжается).
4. Пункты бланка со своими отметками справа — разные абзацы.
5. Перенос «цифра-» + цифра — диапазон без пробела."""
from pathlib import Path

import fitz
import pytest

import layout as LY
import lines as LN
import rows as RW

FONTS = [(r"C:\Windows\Fonts\arial.ttf", r"C:\Windows\Fonts\arialbd.ttf"),
         ("/usr/share/fonts/truetype/dejavu/DejaVuSans.ttf", "/usr/share/fonts/truetype/dejavu/DejaVuSans-Bold.ttf")]
FONT, FONT_B = next((r, b) for r, b in FONTS if Path(r).exists() and Path(b).exists())


def _fr(t, bb, size=9.0, bold=False, **kw):
    return {"t": t, "bbox": list(bb), "size": size, "bold": bold, "family": "monospace", **kw}


def _texts(frags, hs=(), vs=()):
    return [p["text"] for p in LY.group(frags, list(hs), list(vs))]


# --- 1. продолжение с кода или числа; количество — не номер пункта --------------------------------------------------

@pytest.mark.parametrize("prev, nxt, want", [
    ("A 31,32SZ93S004 háromutú szelepet a", "30UJ97B001 felé NYITNI.", True),           # артикль в конце строки
    ("A rendszer vízoldali leürítéséhez a", "30UJ05,06,07,08S002 NYITNI.", True),
    ("(RL), 7 bar-os gőzrendszer (RX),", "5 bar-os gőzrendszer (RY)", True),               # запятая
    ("Az alábbi armatúrák zárva vannak és", "30RR21S001 nyitva", True),                    # союз
    # строка кончается словом или точкой — перед кодом обычно новая строка перечня
    ("A 30UJ97D001 LEÁLLÍTANI, kiszakaszoltatni.", "30UJ05S201 ZÁRNI.", False),
    ("Az alábbi armatúrák kézi szelepei", "30RR21S001", False),
    ("A turbina fordulatszáma", "1000 felett", False),                                     # меньше трёх слов
    ("Ez egy mondat része,", "Következő mondat", False),                                   # прописная — как прежде
])
def test_mid_phrase_continues_with_code_after_comma_or_function_word(prev, nxt, want):
    assert LN.mid_phrase(prev, nxt) is want


def test_phrase_continues_with_code_line_in_cell():
    """3SZ17 л. 9 (моноширинный, рамки выше шага): «A 31,32SZ93S004 háromutú szelepet a» / «30UJ97B001 felé NYITNI.» —
    одна фраза; следующая строка «A 30UJ05,06,07,08S201 NYITNI.» — своя. Прежде вторая строка шла отдельным сегментом
    (продолжение ждали только со строчной буквы)."""
    l1 = _fr("A 31,32SZ93S004 háromutú szelepet a", (71.0, 152.0, 321.0, 167.0))
    l2 = _fr("30UJ97B001 felé NYITNI.", (71.0, 160.0, 201.0, 178.0))
    l3 = _fr("A 30UJ05,06,07,08S201 NYITNI.", (71.0, 171.0, 233.0, 188.0))
    vs = [(96, 480, 319.0), (96, 480, 70.0), (96, 480, 562.0)]
    assert _texts([l1, l2, l3], vs=vs) == [l1["t"] + " " + l2["t"], l3["t"]]


def test_bare_number_with_lowercase_word_is_quantity_mid_phrase():
    """«5 bar-os gőzrendszer (RY) …» после «…7 bar-os gőzrendszer (RX),» (1SZ01 л. 8) — продолжение фразы: голое число и
    слово со строчной после запятой или служебного слова — количество, не номер пункта. Прежде строка с «5 » начинала свой
    абзац. Контрпримеры: «1 Előzmények» — номер; значения столбца «AE-re jóváhagyott ütemben» / «3 havonta egyszer»
    (1SZ01 л. 97, каждое против своего пункта) — раздельно, как прежде."""
    assert LY.is_mark("1 Előzmények") and LY.is_mark("1. Bevezetés") and LY.is_mark("2) a szelep")
    assert LY.is_mark("2 db szelep zárva") and LY.is_mark("3 havonta egyszer", "AE-re jóváhagyott ütemben")
    assert not LY.is_mark("5 bar-os gőzrendszer (RY)", "(RL), 7 bar-os gőzrendszer (RX),")
    a = _fr("blokki főgőz rendszer (RA), tápvíz rendszer (RL), 7 bar-os gőzrendszer (RX),", (145.3, 350.0, 565.3, 365.3))
    b = _fr("5 bar-os gőzrendszer (RY) víztelenítéseinek névleges állapothoz tartozó", (145.3, 360.2, 559.9, 375.5))
    assert _texts([a, b]) == [a["t"] + " " + b["t"]]
    v = [_fr("AE-re", (350, 100, 380, 115)), _fr("jóváhagyott", (342, 110, 388, 125)), _fr("ütemben", (348, 120, 382, 135)),
         _fr("3 havonta", (345, 130, 385, 145)), _fr("egyszer", (349, 140, 381, 155))]
    vs = [(80, 200, 335.0), (80, 200, 395.0)]
    assert _texts(v, vs=vs) == ["AE-re jóváhagyott ütemben", "3 havonta egyszer"]


# --- 2. смешанное начертание; куски строки — по линиям ---------------------------------------------------------------

def _mixed_page(y):
    """«Ha » обычным, «n ≥ 1000 fordulat/perc,» жирным на 0,6 pt выше, «akkor a» обычным; ниже — продолжение фразы."""
    d = fitz.open()
    pg = d.new_page(width=595, height=842)
    put = lambda x, yy, t, f: pg.insert_text((x, yy), t, fontname="F1" if f == FONT_B else "F0", fontfile=f, fontsize=9)
    put(326, y, "Ha", FONT)
    put(350.6, y - 0.6, "n ≥ 1000 fordulat/perc,", FONT_B)
    put(350.6 + fitz.Font(fontfile=FONT_B).text_length("n ≥ 1000 fordulat/perc, ", 9), y, "akkor a", FONT)
    put(326, y + 10.2, "1*SE10S801, S804 tömörtelen. Ebben az esetben a gyorszárókat", FONT)
    return d, pg


def test_partly_bold_line_does_not_break_paragraph():
    """1SZ01 л. 26: «Ha **n ≥ 1000 fordulat/perc**, akkor a» / «1*SE10S801, S804 tömörtelen. …» — одна фраза. Прежде:
    жирный кусок на 1,8 pt выше попадал в ряд сортировки раньше «Ha», и куски строки не склеивались; строку с жирной
    частью начертание отрывало от следующей. Проверка на сетке высот строки. Абзац с жирной частью — не жирный целиком."""
    bad = []
    for k in range(30):
        d, pg = _mixed_page(400 + 0.1 * k)
        pars = LY.group(LY.frags_from_text(pg), [], [])
        got = [" ".join(p["text"].split()) for p in pars]
        if got != ["Ha n ≥ 1000 fordulat/perc, akkor a 1*SE10S801, S804 tömörtelen. Ebben az esetben a gyorszárókat"] \
                or pars[0]["bold"]:
            bad.append((400 + 0.1 * k, got))
    assert not bad, bad[:2]


def test_bold_heading_still_apart_from_body():
    """Контрпример: заголовок жирным целиком и строка текста под ним — разные абзацы, как прежде."""
    h = _fr("Az alábbi armatúrák zárva vannak és", (70, 100, 300, 112), bold=True)
    b = _fr("30RR21S001 nyitva van, a turbina üzemel.", (70, 111, 280, 123))
    assert _texts([h, b]) == [h["t"], b["t"]]
    h2 = _fr("Az alábbi armatúrák zárva vannak és", (70, 100, 300, 112), bold=True, mixed=True)
    assert _texts([h2, b]) == [h2["t"] + " " + b["t"]]   # строка с жирной частью и фраза продолжается — не рвёт
    # фраза кончилась: «…3.8.-tól FOLYTATNI.» (жирная часть) и метка «MEGJEGYZÉS:» под ней (1SZ01 л. 26) — раздельно
    f1 = _fr("A 1SZ01 KU 3.8.-tól FOLYTATNI.", (326, 663.5, 488, 680.9), mixed=True)
    f2 = _fr("MEGJEGYZÉS:", (326, 673.7, 391, 690.9), bold=True)
    assert _texts([f1, f2]) == [f1["t"], f2["t"]]
    # строка с жирной частью, но не с того же левого края: «Verziószám: 0.» справа над строкой шапки во всю ширину
    # (30SZBVTIT TU) — разные абзацы
    v = _fr("Verziószám: 0.", (454.3, 70.1, 530.2, 86.2), 12.0)
    m = _fr("Megnevezés: SZBV hajtások ejtési próbája és tartalék járat mérése Oldalszám: 9.", (76.9, 83.7, 529.7, 100.3),
            12.0, bold=True, mixed=True)
    assert _texts([v, m]) == [v["t"], m["t"]]


# --- 3. пустая строка — и для меток портала ---------------------------------------------------------------------------

def test_blank_line_stops_portal_merge_unless_phrase_continues():
    """Метки портала склеивали пункты через пустую строку: «A, fordulatszám … üzemképtelen,» / пусто / «B, főszervó …»
    (1SZ01 л. 112). Пустая строка — граница и для roles.apply (Rows.apart), кроме фразы, что продолжается: «Várakozási
    idő: 10 perc,» / пусто / «vagy ha a kezelő …» (02VE25 л. 13) — одна фраза."""
    import roles as R
    a = _fr("A, fordulatszám távadók közül egy db. üzemképtelen,", (145, 300, 450, 315))
    b = _fr("B, főszervó rögzítő (blokkoló) köri és üzemi szabályzó", (145, 320.5, 560, 335.5))
    c = _fr("Várakozási idő: 10 perc,", (160, 400, 300, 418.8), 11.04)
    e = _fr("vagy ha a kezelő a TOVÁBBLÉPÉS gombra kattint, akkor a", (108, 424.4, 531, 443.2), 11.04)
    rw = RW.Rows([a, b, c, e], [])
    assert rw.apart(a, b) == "пустая строка" and rw.apart(c, e) is None
    # слово, разорванное переносом, в шапке с просторным межстрочием (1SZ01 л. 90: шаг 20,4 pt при кегле 9) — не граница
    h1, h2 = _fr("Megjelení-", (146.6, 82.5, 206.0, 97.8)), _fr("tés helye", (149.3, 102.9, 197.9, 118.2))
    assert RW.Rows([h1, h2], []).apart(h1, h2) is None and _texts([h1, h2]) == ["Megjelenítés helye"]
    lines = [{"id": k, "x0": f["bbox"][0], "x1": f["bbox"][2], "top": f["bbox"][1], "bottom": f["bbox"][3]}
             for k, f in enumerate([a, b, c, e])]
    labels = {0: {"new": True, "role": "list"}, 1: {"new": False, "role": "list"},
              2: {"new": True, "role": "body"}, 3: {"new": False, "role": "body"}}
    pars = [{"frags": [f], "bbox": f["bbox"], "size": f["size"], "bold": False, "text": f["t"]} for f in (a, b, c, e)]
    got = [p["text"] for p in R.apply(pars, lines, labels, LN.join_lines, apart=rw.apart)]
    assert got == [a["t"], b["t"], c["t"] + " " + e["t"]]


# --- 4. пункты бланка со своими отметками ------------------------------------------------------------------------------

def test_form_items_with_own_marks_are_separate():
    """30PR25GT л. 13 (EYES N5): «30TJ13 ág: TZ26S003,S004 - TZ27S001,S002 - TH88S201,S202» и «30TJ23 ág: …» — у каждой
    строки справа своя отметка «O»: разные пункты. Прежде длинная первая строка абзац не закрывала, «30TJ23 ág:» уходил в
    её конец. Контрпример: пункт в две строки с одной отметкой — один абзац."""
    l1 = _fr("30TJ13 ág: TZ26S003,S004 - TZ27S001,S002 - TH88S201,S202", (105.1, 557.1, 348.5, 569.6))
    l2 = _fr("30TJ23 ág: TJ24S004,S005 - TJ24S006,S007", (105.1, 567.6, 276.6, 580.0))
    o1 = _fr("O", (498.5, 557.3, 507.2, 569.4))
    o2 = _fr("O", (498.5, 567.7, 507.2, 579.8))
    vs = [(150, 700, 470.0), (150, 700, 95.0), (150, 700, 540.0)]
    assert [t for t in _texts([l1, o1, l2, o2], vs=vs) if t != "O"] == [l1["t"], l2["t"]]
    m1 = _fr("A TH, TJ vezetékek fővízkör felől kinyitott armatúrák zárva vannak, a", (105.1, 598.7, 450, 610.8))
    m2 = _fr("szivárgás ellenőrzése után.", (105.1, 609.0, 240, 621.1))
    om = _fr("O", (498.5, 609.0, 507.2, 621.0))
    assert [t for t in _texts([m1, m2, om], vs=vs) if t != "O"] == [m1["t"] + " " + m2["t"]]
    # отметки — значок символьного шрифта (□ Wingdings, частная область): «30YA62P401ZQ35 p = 20 - 25 bar,» и «30YP10L101
    # L ≈ 7100 mm,» (30PR132GT л. 5) — каждая строка со своей отметкой, хотя первая кончается запятой
    p1 = _fr("30YA62P401ZQ35  p = 20 - 25 bar,", (145.8, 245.4, 276.3, 257.5))
    p2 = _fr("30YP10L101  L ≈ 7100 mm,", (145.8, 256.4, 250.3, 268.6))
    b1, b2 = _fr("", (505, 246, 512, 257)), _fr("", (505, 257, 512, 268))
    assert [t for t in _texts([p1, b1, p2, b2], vs=vs) if t != ""] == [p1["t"], p2["t"]]


# --- 5. перенос «цифра-» + цифра -----------------------------------------------------------------------------------------

@pytest.mark.parametrize("prev, nxt, want", [
    ("nyomás 1000-", "1200 bar között", "nyomás 1000-1200 bar között"),     # диапазон: дефис на месте, без пробела
    ("a 30-", "as szelep", "a 30-as szelep"),                               # как прежде
    ("a TH10-", "20 vezetékek", "a TH10-20 vezetékek"),                     # код с цифрой перед дефисом
    ("szekrény -", "30XM:4 kapcsain", "szekrény -30XM:4 kapcsain"),        # клемма «-30XM:4» (30SZBVBIT TU), прежде «- 30XM:4»
    ("a mosatás-", "12 perc", "a mosatás- 12 perc"),                        # буква перед дефисом — как прежде
    ("a 30TH50L101 szint felé -", "1,2 mm", "a 30TH50L101 szint felé -1,2 mm"),  # отдельный «-» после слова + число — минус (3SZ19)
    ("p = 20 -", "25 bar", "p = 20 - 25 bar"),                            # «число -» + число — диапазон через пробелы
])
def test_digit_hyphen_digit_is_range(prev, nxt, want):
    """E9.7г: «цифра-» + цифра — диапазон, склеивается без пробела (сборка, docx_layout._join, давала «x-12», нарезка —
    «x- 12»; после этого правила копию в сборке можно свести к lines.glue)."""
    assert LN.join_lines([prev, nxt]) == want


def test_tesseract_words_do_not_glue_number_ranges():
    """Слова Tesseract: «118-» в конце строки и номер штампа «10135772» в начале следующей — разные числа. Склеенное
    «118-10135772» не сверялось с чтением портала и пропадало (титулы-сканы 30HU*75GT, 30SZBV*, 30GF*: −8 знаков на
    документ). Диапазоны — только для строк текстового слоя (lines.glue(ranges=True)); перенос слова — как прежде."""
    import segments as S
    w = lambda t, line: {"t": t, "line": line, "key": (1, 1)}
    assert [x["t"] for x in S._dehyphen([w("118-", (1, 1, 1)), w("10135772", (1, 1, 2))])] == ["118-", "10135772"]
    assert [x["t"] for x in S._dehyphen([w("végre-", (1, 1, 1)), w("hajtani", (1, 1, 2))])] == ["végrehajtani"]
