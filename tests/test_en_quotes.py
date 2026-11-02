# -*- coding: utf-8 -*-
"""Кавычки в переводе на английский (Guide to Creating Exercises, ru-en): русские ёлочки «Select» оставались в тексте.
«…» → “…”, вложенные „…“ → ‘…’; кавычки внутри пути, ключа команды, адреса и прямые " ' — как есть; перевод на
русский (hu-ru и др.) — без изменений."""
import pytest

import translate as T
from test_pairs import gloss  # noqa: F401 — фикстура глоссария


@pytest.mark.parametrize("src, want", [
    ("Click «Select».", "Click “Select”."),
    ("Open the «Settings» window (see «Help»): press «OK»!", "Open the “Settings” window (see “Help”): press “OK”!"),
    ("The «„Start“ button» is grey.", "The “‘Start’ button” is grey."),
    ("<f1>«Save»</f1> and <f2>«Exit»</f2>", "<f1>“Save”</f1> and <f2>“Exit”</f2>"),
    ("«Select»/«Cancel» — choose one", "“Select”/“Cancel” — choose one"),
    ("Log file «C:\\Temp\\log.txt».", "Log file “C:\\Temp\\log.txt”."),
    ("Status:\t«Ready»\nNext: «Run»", "Status:\t“Ready”\nNext: “Run”"),
])
def test_russian_quotes_become_english(src, want):
    assert T.en_quotes(src) == want


@pytest.mark.parametrize("src", [
    'run.exe --title=«Report» --mode="fast"',          # ключ команды
    "C:\\Data\\«Old»\\report.doc",                     # путь
    "/opt/«app»/bin/start.sh",
    "http://host/«page»",
    'Press "OK" and type \'yes\'',                     # прямые кавычки — не трогать
    "Already “English” and ‘nested’ quotes",
    "No quotes at all, 0.4 kV",
])
def test_code_paths_and_other_quotes_untouched(src):
    assert T.en_quotes(src) == src


def test_post_converts_only_for_english_target(gloss):  # noqa: F811
    for pair in ("ru-en", "hu-en"):
        assert T.Translator(glossary=str(gloss), pair=pair).post("Нажмите «Выбрать».", "Click «Select».") == "Click “Select”."
    for pair in ("hu-ru", "en-ru"):                                      # перевод на русский — ёлочки остаются
        ru = T.Translator(glossary=str(gloss), pair=pair).post("Kattintson a „Kiválaszt” gombra.", "Нажмите «Выбрать».")
        assert ru == "Нажмите «Выбрать»."
