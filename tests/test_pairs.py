# -*- coding: utf-8 -*-
"""Пара языков проекта "pair": hu, ru, en в обе стороны (22.10): разбор пары, разворот глоссария, промпты и проверки
по языкам, постобработка, суффикс файла, OCR. hu-ru — прежнее поведение (его проверяют остальные тесты набора)."""
import pytest

import glossary as G
import langs as L
import review_portal as RP
import segments as S
import translate as T

GLOSS = """# Глоссарий

## Правила применения

- Венгерские аббревиатуры: при первом употреблении — русский вариант и оригинал в скобках.

## Оборудование

| Венгерский | Русский | Комментарий |
|---|---|---|
| szivattyú | насос | |
| biztonsági szelep | предохранительный клапан | строгий |
"""


@pytest.fixture
def gloss(tmp_path):
    f = tmp_path / "g.md"
    f.write_text(GLOSS, encoding="utf-8")
    return f


def test_parse_pair_and_suffix():
    assert L.parse_pair(None) == ("hu", "ru") and L.parse_pair("RU-EN") == ("ru", "en")
    assert L.suffix(None) == "_RU" and L.suffix("en-hu") == "_HU" and L.suffix("hu-en") == "_EN"
    for bad in ("ru-ru", "de-ru", "ru"):
        with pytest.raises(ValueError):
            L.parse_pair(bad)


def test_glossary_flips_for_reverse_pair(gloss):
    rules, ents = G.load(gloss)                                   # как записано (hu-ru)
    assert [(e.hu, e.ru) for e in ents] == [("szivattyú", "насос"), ("biztonsági szelep", "предохранительный клапан")]
    assert "аббревиатуры" in rules
    rules, ents = G.load(gloss, ("ru", "hu"))                     # ru-hu: развёрнут, правила о другом направлении — нет
    assert [(e.hu, e.ru) for e in ents] == [("насос", "szivattyú"), ("предохранительный клапан", "biztonsági szelep")]
    assert rules == "" and ents[1].note == "строгий"
    assert [e.hu for e in G.find("Предохранительного клапана насоса нет", ents)] == ["предохранительный клапан", "насос"]
    assert G.load(gloss, ("hu", "en"))[1] == []                   # таблица hu-ru для пары hu-en — не берётся


def test_translator_ru_hu_prompt_check_and_post(gloss):
    tr = T.Translator(glossary=str(gloss), pair="ru-hu")
    assert tr.system.startswith("You are a professional translator of technical documentation from Russian into Hungarian.")
    assert "„…”" in tr.system and "decimal comma stays" in tr.system
    p = tr._prompt([{"text": "Насос включён"}], [], "")
    assert "Glossary terms for this fragment:\n- насос → szivattyú" in p and "Translate the segments:" in p
    assert tr.full_check("Давление 0,4 МПа", "Nyomás 0,4 MPa") == []
    assert tr.full_check("Давление 0,4 МПа", "Nyomás 0.4 MPa") == []          # десятичный знак — по любой стороне
    assert tr.full_check("Насос Grundfos включён", "A szivattyú bekapcsolva") == ["латиница потеряна: Grundfos"]
    assert tr.full_check("Насос включён полностью", "A насос включён") == ["не переведено"]
    assert tr.post("Насос", "A szivattyú 10 kV") == "A szivattyú 10 kV"       # единицы в венгерском — как есть



def test_strict_terms_flip(gloss, tmp_path):
    terms = tmp_path / "t.md"
    terms.write_text("## Проверенные\n\n| Венгерский | Русский | Комментарий |\n|---|---|---|\n"
                     "| üzemi nyomás | рабочее давление | строгий |\n", encoding="utf-8")
    tr = T.Translator(glossary=str(gloss), extra=str(terms), pair="ru-hu")
    assert tr.term_issues("Рабочее давление в норме", "A nyomás rendben") == [
        "термин документа: рабочее давление → üzemi nyomás"]
    assert tr.term_issues("Рабочее давление в норме", "Az üzemi nyomás rendben") == []


def test_translator_en_ru_numbers_leftover_and_units(gloss):
    tr = T.Translator(glossary=str(gloss), pair="en-ru")
    assert "from English into Russian" in tr.system and "decimal point becomes a decimal comma" in tr.system
    assert T.check_generic("Price 1,250.50 rubles", "Цена 1 250,50 рублей", "en", "ru") == []
    assert T.check_generic("The pump is stopped", "Насос stopped", "en", "ru") == ["английский остался: stopped"]
    assert T.check_generic("The pump is stopped", "The pump is stopped", "en", "ru") == ["не переведено"]
    assert tr.post("Voltage 10 kV", "Напряжение 10 kV") == "Напряжение 10 кВ"     # единицы — по-русски
    assert tr.post("PUMP STOPPED", "НАСОС ОСТАНОВЛЕН") == "НАСОС ОСТАНОВЛЕН"        # fix_caps — только hu-ru


def test_translator_hu_en_and_en_hu(gloss):
    tr = T.Translator(glossary=str(gloss), pair="hu-en")
    assert "from Hungarian into English" in tr.system and "decimal comma becomes a decimal point" in tr.system
    assert T.check_generic("A szivattyú leállt", "The pump szivattyú stopped", "hu", "en") == ["венгерский остался: szivattyú"]
    assert "tömör" in tr._prompt([{"text": "A tömörtelenség megszűnt"}], [], "")   # разбор отрицаний — для венгерского
    tr2 = T.Translator(glossary=str(gloss), pair="en-hu")
    assert "„…”" in tr2.system
    assert T.check_generic("The valve is closed now", "The valve is closed now", "en", "hu") == ["не переведено"]


def test_hu_ru_unchanged(gloss):
    tr = T.Translator(glossary=str(gloss))
    assert tr.hu_ru and tr.system == T.SYSTEM.format(rules=tr.rules)
    assert "Термины глоссария для этого фрагмента:" in tr._prompt([{"text": "szivattyú"}], [], "")


def test_review_generic_prompts():
    sysm, user = RP.generic_prompts("ru", "hu")
    assert "from Russian into Hungarian" in sysm and "должен" not in sysm
    assert "Original (RU):" in user and "Translation (HU):" in user and "corrected Hungarian translation" in user


def test_ocr_language(monkeypatch):
    monkeypatch.setattr(S, "SRC_LANG", "hu")
    assert S.ocr_prompt() == S.OCR_PROMPT
    monkeypatch.setattr(S, "SRC_LANG", "ru")
    assert "(русский)" in S.ocr_prompt() and "ő" not in S.ocr_prompt()
    assert L.TESS["ru"] == "rus+eng" and L.TESS["hu"] == "hun"
