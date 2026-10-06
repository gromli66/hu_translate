# -*- coding: utf-8 -*-
"""Проверки без портала: разбор глоссария, проверки перевода, постобработка, ключи памяти.
Каждый случай — ловушка, найденная на комплекте «Пакш» (01–06.10.2026)."""
import json

import glossary as G
import morph
import segments as S
import translate as T
import pipeline as P


def entry(hu, ru="перевод", note=""):
    e = G.Entry(hu, ru, note, "")
    e.alts = G._alts(hu, ru)
    return e


# ------------------------------------------------------------------ глоссарий
def test_ellipsis_is_gapped_term_not_two_words():
    # многоточие — пропуск внутри термина; раньше голое «hűtővíz» срабатывало на любой строке
    assert G._alts("biztonsági … hűtővíz") == [("biztonsági", "hűtővíz")]


def test_slash_word_choice_when_russian_has_slash():
    # «zárt / nyitott primerkör → уплотненный / разуплотненный первый контур»: косая — выбор слова
    assert G._alts("zárt / nyitott primerkör", "уплотненный / разуплотненный первый контур") == \
        [["zárt", "primerkör"], ["nyitott", "primerkör"]]


def test_negated_variant_drops_positive_twin():
    # «nem oldott / oldott → нерастворенный»: вариант без «nem» — противоположный смысл
    assert G._alts("nem működik / működik") == [["nem", "működik"]]


def test_parenthesized_note_is_not_part_of_term():
    assert G._alts("turbógenerátor (TG)") == [["turbógenerátor"]]


def test_capitalized_abbreviation_matches_uppercase_entry():
    # «Blügy» в тексте = «BLÜGY» глоссария (аббревиатуры от 4 букв)
    assert [e.hu for e in G.find("A Blügy értesíti", [entry("BLÜGY")])] == ["BLÜGY"]


def test_load_reads_rules_and_tables(tmp_path):
    md = tmp_path / "g.md"
    md.write_text("# Глоссарий\n\n## Правила применения\n\n- коды не переводятся\n\n## Системы\n\n"
                  "| Венгерский | Русский | Комментарий |\n|---|---|---|\n| hűtővíz | охлаждающая вода | |\n",
                  encoding="utf-8")
    rules, entries = G.load(md)
    assert rules == "- коды не переводятся"
    assert [(e.hu, e.ru, e.section) for e in entries] == [("hűtővíz", "охлаждающая вода", "Системы")]


def test_gloss_misses_by_stems():
    e = entry("biztonsági hűtővíz", "охлаждающая вода ответственных потребителей")
    assert T.gloss_misses("A biztonsági hűtővíz rendszer", "система", [e]) == ["biztonsági hűtővíz"]
    assert T.gloss_misses("A biztonsági hűtővíz rendszer", "система охлаждающей воды ответственных потребителей", [e]) == []


# ------------------------------------------------------------------ проверки перевода
def test_code_lost_is_flagged_unit_translation_is_not():
    assert T.check("400KV táp 30TL02GT", "питание 400 кВ 30TL02GT") == []
    assert T.check("400KV táp 30TL02GT", "питание 400 кВ") == ["нет кодов/чисел: 30TL02GT"]


def test_action_flip():
    assert T.action_flip("A szelepet nyitni kell", "Клапан закрыть")
    assert not T.action_flip("A szelepet nyitni kell", "Клапан открыть")
    # «le és felterhelés» — это «разгрузка и нагрузка», а не перевёрнутое действие
    assert not T.action_flip("le és felterhelés", "разгрузка и нагрузка")


def test_badness_counts_list_items():
    assert T.badness([]) == 0
    assert T.badness(["нет кодов/чисел: 1, 2", "длина"]) == 3


def test_negation_morphemes():
    assert morph.neg_words("tömörtelenség esetén") == [("tömörtelenség", "tömör", "telen", "ség")]


def test_fix_units_and_service_marks():
    # «[уточнить]» из подсказок глоссария не должно попадать в перевод
    assert T.fix_units("5 mg/dm3 [уточнить]") == "5 мг/дм³"


def test_ocr_leader_noise_removed():
    assert S._denoise("AláíráS:,...sszsseseeeeeeeeseeeeseeeseezeszeezesen") == "AláíráS…"
    assert S._denoise("Dátum:....................") == "Dátum…"
    for word in ("elszállítása", "teszteseteket", "kiszerelése", "Ellenőrizte"):   # обычные слова не трогаем
        assert S._denoise(word) == word


def test_needs_translation():
    assert not S.needs_translation("30TL02GT")
    assert S.needs_translation("Szelep zárva")


# ------------------------------------------------------------------ конвейер
def test_memory_key_keeps_tabs_and_breaks():
    # схлопывание табуляций давало колонтитулам чужую раскладку (560 строк [TAB]/[BR], v4 05.10)
    assert P.norm("Oldalszám:\t95") != P.norm("Oldalszám:\t\t95")
    assert P.norm("a  b ") == "a b"


def test_tm_key_matches_apply_review():
    import apply_review as AR
    t = "<f1>Szelep</f1>\t zárva  "
    assert P.tmkey(t) == AR.norm_hu(t) == "Szelep zárva"


def test_load_project_resolves_paths(tmp_path):
    (tmp_path / "data").mkdir()
    (tmp_path / "project.json").write_text(json.dumps({"glossary": "data/g.md"}), encoding="utf-8")
    prj = P.load_project(tmp_path / "project.json")
    assert prj["glossary"] == str((tmp_path / "data" / "g.md").resolve())
    assert prj["tm"].endswith("tm.json")


def test_doc_files_skip_review_backups(tmp_path):
    (tmp_path / "json").mkdir()
    for n in ("a.json", "a.before_review.json"):
        (tmp_path / "json" / n).write_text("{}", encoding="utf-8")
    assert [f.name for f in P.doc_files(tmp_path)] == ["a.json"]
