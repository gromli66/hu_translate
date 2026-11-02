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
