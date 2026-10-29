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
    # наречие «fokozatosan» («постепенно») — не «fokozat» («ступень»): к верному переводу замечания нет
    st = entry("fokozat", "ступень")
    assert T.gloss_misses("csak fokozatosan nyitható", "открывается только постепенно", [st]) == []
    assert T.gloss_misses("a fokozat nő", "растёт", [st]) == ["fokozat"]


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
    assert S._denoise("Aláírás:..............ggsseeseseeeteeeéeteeíéeéséééétésései") == "Aláírás…"
    for word in ("elszállítása", "teszteseteket", "kiszerelése", "Ellenőrizte", "szükségessége:"):   # слова не трогаем
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


def test_check_does_not_flag_glued_codes_and_numbered_words():
    """Коды через запятую без пробела, слово в скобку и номер пункта, слипшийся со словом («5.Teszt»), — не потеря кодов
    и не лишнее число в верном переводе (30RR2126GT: 8 ложных замечаний, 28.10)."""
    assert T.check("ZÁRNI a 30RS61S004,30RS61S005 és a 30RS61S006 kézi szelepeket.",
                   "ЗАКРЫТЬ ручные вентили 30RS61S004, 30RS61S005 и 30RS61S006.") == []
    assert T.check("A blokk M1 üzemállapotban vagy leállási üzemállapotban(M2) van.",
                   "Блок находится в состоянии M1 или в состоянии останова (M2).") == []
    assert T.check("5.Teszt utáni állapot", "5. Состояние после тестирования") == []
    assert T.check("A 30RS61S004 szelep.", "Вентиль 30RS61S004 и 77.") == ["лишнее число: 77"]       # настоящее лишнее — по-прежнему


def test_hourly_units_hungarian_and_flat_digit_variant():
    """«5 dm3/óra» → «5 дм3/ч» (цифра без степени) — единица переведена, не потеря кода (30BL150FÜT: 10 замечаний)."""
    assert T.check("Szivárgás < 5 dm3/óra.", "Протечка < 5 дм3/ч.") == []
    assert T.fix_units("5 dm3/óra") == "5 дм³/ч"


def test_invented_noun_vyvesivanie_is_fixed_in_every_case():
    """«ВЫВЕСИВАНИЕ» — слова нет: модель собрала его от «вывесить плакат» глоссария (107 строк партии «Пакш», 28.10)."""
    assert T.fix_units("ПЛОМБИРОВАНИЕ, ВЫВЕСИВАНИЕ ПЛАКАТА") == "ПЛОМБИРОВАНИЕ, ВЫВЕШИВАНИЕ ПЛАКАТА"
    assert T.fix_units("Вывесивание плаката") == "Вывешивание плаката" and T.fix_units("вывесивания") == "вывешивания"


def test_names_in_signature_block_and_mail_header_are_not_hungarian_leftover():
    """Имена остаются латиницей по глоссарию: «Név: Bajnai Péter Krisztián», «Juhász Antal <juhasza@npp.hu>» — шум более
    половины замечаний партии (28.10). Настоящий венгерский остаток рядом по-прежнему ловится."""
    assert T.check("Név: Bajnai Péter Krisztián | Szervezet: RTO Aláírás:",
                   "Ф. И. О.: Bajnai Péter Krisztián | Подразделение: RTO Подпись:") == []
    assert T.check("Név: DOLHA GÁBOR | Szervezet: RTO", "Ф. И. О.: DOLHA GÁBOR | Подразделение: RTO") == []
    assert T.check("Cc: Bucsi Gábor <bucsi@npp.hu>; Kaszás István <kaszasi@npp.hu>",
                   "Копия: Bucsi Gábor <bucsi@npp.hu>; Kaszás István <kaszasi@npp.hu>") == []
    assert T.check("Név: Bárdos Zoltán Aláírás: rendben", "Ф. И. О.: Bárdos Zoltán Подпись: rendben") == ["венгерский остался: rendben"]


def test_filler_line_hyphenated_caps_and_numbered_abbreviation_are_not_lost():
    """Линия-заполнитель «20.......év.....hó.....nap» не код; «MENNYI- SÉG» — одно слово, перенесённое в ячейке;
    «1.VT» → «СВО-1» (номер впереди, аббревиатура переводится) — всё это верный перевод (30RXBSZGT, 30YCGT)."""
    assert T.check("P a k s , 20.......év.........................hó............nap",
                   "Пакш, 20.......г.........................мес............день") == []
    assert T.check("MENNYI- SÉG (liter)", "КОЛИЧЕСТВО (литр)") == []
    assert T.check("30TE01N001 TÖMÖRTELENSÉG AZ 1.VT. SZŰRŐ FEDÉLNÉL", "30TE01N001 НЕГЕРМЕТИЧНОСТЬ НА КРЫШКЕ ФИЛЬТРА СВО-1",
                   gloss_abbr=frozenset({"VT"})) == []
    assert T.check("30TE01N001 TÖMÖRTELENSÉG AZ 1.VT. SZŰRŐ", "30TE01N001 НЕГЕРМЕТИЧНОСТЬ ФИЛЬТРА СВО-2",
                   gloss_abbr=frozenset({"VT"})) == ["нет кодов/чисел: 1.VT"]


def test_retry_hides_equipment_designator_behind_a_mark(monkeypatch):
    """«SZIVÁRGÁS A 1.GF.H.ÁG.SZEK.FEDÉL TÖMÍTÉSEN» модель «расшифровывала» каждый раз по-своему (48 строк 30YCGT);
    в повторе обозначение уходит под меткой X1 и возвращается дословно."""
    src = "SZIVÁRGÁS A 1.GF.H.ÁG.SZEK.FEDÉL TÖMÍTÉSEN (↑P 30YA12P501B1)"
    assert T.mask_designators(src) == ("SZIVÁRGÁS A X1 TÖMÍTÉSEN (↑P 30YA12P501B1)", {"X1": "1.GF.H.ÁG.SZEK.FEDÉL"})
    assert T.mask_designators("A 30YA12P501B1 szelep, 5.Teszt.") == ("A 30YA12P501B1 szelep, 5.Teszt.", {})
    asked = []

    def call(self, units, *a, **k):
        asked.append(units[0]["text"])
        # первый проход (пачка) «расшифровывает», повтор видит метку
        ru = "ПРОТЕЧКА КРЫШКИ ГОРЯЧЕГО ОТВОДА 1-ГО ПГ (↑P 30YA12P501B1)" if "X1" not in units[0]["text"] \
            else "ПРОТЕЧКА НА УПЛОТНЕНИИ X1 (↑P 30YA12P501B1)"
        return {0: ru}, {}
    monkeypatch.setattr(T.Translator, "_call", call)
    tr = T.Translator.__new__(T.Translator)
    tr.hu_ru, tr.ru_en, tr.src, tr.tgt = True, False, "hu", "ru"
    tr.allowed_latin, tr.gloss_abbr, tr.vocab, tr.abbrs, tr.strict_entries, tr.entries = frozenset(), frozenset(), frozenset(), {}, [], []
    tr.think, tr.retry_think = False, False
    res = tr.run_batch({"units": [{"key": "1", "text": src}], "first_abbr": [], "context": ""})["1"]
    assert asked[0] == src and "X1" in asked[1]
    assert res["ru"] == "ПРОТЕЧКА НА УПЛОТНЕНИИ 1.GF.H.ÁG.SZEK.FEDÉL (↑P 30YA12P501B1)" and res["issues"] == []


def test_numpy_values_in_segments_do_not_break_saving(tmp_path):
    """Из распознавания листов в сегменты попадали значения numpy (bool_, int64): json.dumps падал и обрывал весь набор
    на извлечении (30GF118GT_nyilv, 28.10)."""
    import numpy as np
    (tmp_path / "json").mkdir()
    P.save_doc(tmp_path, "d", {"segs": [{"id": np.int64(1), "loc": {"ocr": np.bool_(True), "size": np.float32(10.5)}}]})
    d = json.loads((tmp_path / "json" / "d.json").read_text(encoding="utf-8"))
    assert d["segs"][0] == {"id": 1, "loc": {"ocr": True, "size": 10.5}}


def test_glued_tokens_codes_units_and_brands_are_not_false_issues():
    """Ложные замечания партии 4 (28.10): слово или артикль, приставшие к номеру («2.1.Személyi», «a30TH60S201»,
    «bar-os(123», «sebességét-5», «mivoltát.(30»); «INFO» в «3SZ17/INFO» — часть кода; STORZ, ASME — марки; «1,5 bar/10 мин»
    и «1/min» — единицы по глоссарию."""
    assert T.check("2.1.Személyi és adminisztratív feltételek", "2.1. Персональные и административные условия") == []
    assert T.check("A 30GE01 panelen a30TH60S201 üzemmód kapcsolót", "На панели 30GE01 переключатель режима 30TH60S201") == []
    assert T.check("132 bar-os(123", "при 132 бар (123") == []
    assert T.check("a nyomáscsökkentés sebességét-5 bar/perc- betartva", "соблюдая скорость 5 бар/мин") == []
    assert T.check("ellenőrizni a stabil mivoltát.(30 perc időtartamig stabil.)", "проверить стабильность (30 мин)") == []
    vocab = frozenset({"info"})
    assert T.check("Azonosító: 3SZ17/INFO", "Идентификатор: 3SZ17/INFO", vocab=vocab) == []
    assert T.check("ZÁRVA, STORZ KAPOCS ZÁRÓKUPAKOK FELSZERELVE", "ЗАКРЫТО, ЗАГЛУШКИ НА СОЕДИНЕНИЯХ STORZ УСТАНОВЛЕНЫ") == []
    assert T.fix_units("Перепад давления < 1,5 bar/10 мин.") == "Перепад давления < 1,5 бар/10 мин."
    assert T.fix_units("Частота вращения (1/min), 2900 1/min") == "Частота вращения (1/мин), 2900 1/мин"
    assert T.fix_units("5 bar/h и 1,5 bar") == "5 bar/h и 1,5 бар"
    # единица внутри «2bar/10 perc» переводится: коды — «2» и «10»
    assert {"2", "10"} <= T.codes_of("dp/dt< 2bar/10 perc") and "2bar/10" not in T.codes_of("dp/dt< 2bar/10 perc")
    assert T.check("ha a dp/dt< 2bar/10 perc.", "если dp/dt< 2 бар/10 мин.") == []


def test_word_glued_to_number_is_separated_for_the_model():
    """«ugrás a(z) 4.1.2.29.fejezetre» слитным модель оставляла по-венгерски (28 строк); в запросе слово отделено,
    а номер файла «3SZ17.pdf» и десятичное «2.5mm» не трогаются."""
    assert T.to_model("ugrás a(z) 4.1.2.29.fejezetre") == "ugrás a(z) 4.1.2.29. fejezetre"
    assert T.to_model("a 4.6.fejezetre és 2.5mm, 3SZ17.pdf, 5.Teszt") == "a 4.6. fejezetre és 2.5mm, 3SZ17.pdf, 5.Teszt"


def test_fix_caps_keeps_capitals_of_glossary_abbreviations():
    """«KÜTSZ» (одна гласная — не слово) в оригинале: ДАПЭН из глоссария понижался до «дапэн» (25 строк 28.10);
    лишние прописные модели («НЕМЕДЛЕННО») по-прежнему понижаются."""
    assert T.fix_caps("ÜTSZ, KÜTSZ, A204", "АПЭН, ДАПЭН, A204", frozenset({"ДАПЭН"})) == "АПЭН, ДАПЭН, A204"
    assert T.fix_caps("ÜTSZ, KÜTSZ, A204", "АПЭН, ДАПЭН, A204") == "АПЭН, дапэн, A204"
    assert T.fix_caps("azonnal leállítani", "НЕМЕДЛЕННО остановить", frozenset({"ДАПЭН"})) == "Немедленно остановить"


def test_batch5_false_issues_and_units():
    """Партия 5 (28.10): «Tmax» → «Tмакс» (буква величины + русская пометка), «max.»/«pmax» — не венгерский, «50oC» → «50 °C»,
    «1.LÉPÉS», «információ3», «idő10», «/6FKSz-enként/», «Turb.géph/31(32)UJ21», «1/kampány», «1500/min» и «(1/min)» с
    «об/мин» — не потеря кода; «INFObank» — название системы, а не непереведённый текст."""
    ok = [("T≈50oC, Tmax=120oC", "T≈50 °C, Tмакс=120 °C"), ("1.LÉPÉS Áttárolás", "ШАГ 1. Перекачка"),
          ("♦ Megszűnik a szöveges információ3:", "♦ Исчезает текстовая информация 3:"),
          ("02TD91S006 ZÁRNI Figyelő idő10 perc", "02TD91S006 ЗАКРЫТЬ Время наблюдения 10 мин"),
          ("Villamos betáplálás helye /6FKSz-enként/", "Место электропитания /6FKSz/"),
          ("Tartály szint Turb.géph/31(32)UJ21 oltóberendezés", "Уровень Турб. маш. зал/31(32)UJ21 пожаротушение"),
          ("Fordulatszám  (1/min)", "Частота вращения  (об/мин)"), ("1500/min", "1500/мин"), ("(1/kampány)", "(1/кампания)"),
          ("INFObank | - | - | -", "INFObank | - | - | -"), ("max.30 óra", "max.30 ч"), ("pmax. = 3,7 bar", "pmax. = 3,7 бар")]
    for src, ru in ok:
        assert T.check(src, ru) == [], (src, T.check(src, ru))
    assert T.check("nyomás 1500/min", "давление 1500/мин") == [] and T.check("nyomás 1500/min", "давление") == ["нет кодов/чисел: 1500"]
    assert T.fix_units("T≈50oC, 120 oC; 1500/min; 2900 1/min") == "T≈50 °C, 120 °C; 1500/мин; 2900 1/мин"
    assert T.check("Szöveges információ", "Szöveges információ") != []                      # непереведённое по-прежнему ловится


def test_it_document_identifiers_and_numbers_are_not_false_issues():
    """IS_reko_09 (ПО тренажёра, 29.10): идентификаторы кода («get_ stat_name», «Char16», «MarkTrigger»), имена файлов и
    ключевые слова (.dat, SQL, FIFO, char) остаются латиницей — не «не переведено»/«венгерский остался»; «Sorszám1»,
    «30RW72S201(1.szinten)», «31,32 SA» и пробелы в конце ячейки — не потеря кода/лишнее число/длина; «ASME OM CODE»."""
    ok = [("get_ stat_name", "get_ stat_name"), ("Char16", "Char16"), ("MarkTrigger", "MarkTrigger"), ("l3=ipar4", "l3=ipar4"),
          ("IS_TERM-nek", "IS_TERM"), ("7.2.3\tMain_form\t110", "7.2.3\tMain_form\t110"),
          ("8.4.7\tA paramsxx.dat file-ok szerkezete\t146", "8.4.7\tСтруктура файлов paramsxx.dat\t146"),
          ("1.2.1\tAz SQL adatbázis tartalma\t9", "1.2.1\tСодержание базы данных SQL\t9"),
          ("4.6.3\tA CDQ FIFO puffer\t51", "4.6.3\tБуфер CDQ FIFO\t51"),
          ("  char TRIG_ID[3];\t// a trigger fv. sorszáma", "  char TRIG_ID[3];\t// порядковый номер функции триггера"),
          ("malfun\t\t(integer*2)\tMérés meghibásodás státusza", "malfun\t\t(integer*2)\tСтатус отказа измерения"),
          ("Sorszám1", "№ п/п 1"), ("30RW72S201(1.szinten)", "30RW72S201 (на 1-м уровне)"),
          ("30TV61S001(fent a podiumon)", "30TV61S001 (на подиуме)"), ("31,32 SA üzemen kívül.", "31, 32 SA вне работы."),
          ("Módosított oldalak:1,11", "Измененные страницы: 1, 11"), ("Szivattyú indulás                          ", "Пуск насоса"),
          ("Elektronikus tesztelési utasítások az ASME OM CODE", "Электронные инструкции по ASME OM CODE")]
    for src, ru in ok:
        assert T.check(src, ru) == [], (src, T.check(src, ru))
    assert T.check("Szöveges információ", "Szöveges információ") != []                      # венгерский текст по-прежнему ловится
    assert T.check("Szelep áteresztés", "Szelep áteresztés") != []
    assert T.fix_units("переход к 4.1.0.26.fejezetre") == "переход к разделу 4.1.0.26"
    assert T.fix_units("переход к 4.3. fejezetre") == "переход к разделу 4.3"


def test_program_identifiers_are_not_translated():
    """ПО тренажёра (IS_reko_09, 29.10): имена из кода и типы данных — не текст к переводу."""
    for t in ("get_ stat_name", "Char16", "MarkTrigger", "real*4", "int*4", "Integer*2", "float", "mf_load(loa, meas, stb..)",
              "Check_clients", "<f1>valve_up</f1>", "switch_n,\tswitch_r", "ipar4", "boardMouseMove", "SNPTimeToStep",
              "is_cmd0xx.dat", "Trgxxx.dat", "IS_TERM\\VMS\\comment", "  char PAR2[2];"):
        assert not S.needs_translation(t), t
    for t in ("Sorszám1", "Parancsok", "Szelep zárva", "Integer paraméter címe", "Real output", "Egyedi azonosító"):
        assert S.needs_translation(t), t


def test_it_doc_decimal_dot_and_single_are_not_false_issues():
    """IS_reko_09 (29.10): «0.2 mp.» → «0,2 с» (десятичная запятая) не потеря кода; «egyetlen» — «единственный», не отрицание;
    «int *icon_adr» — код."""
    assert T.check("amely 0.2 mp.-ként adatot küld", "которая каждые 0,2 с отправляет данные") == []
    assert T.check("Egyetlen szimulációs ciklus alatt", "В рамках одного цикла") == []
    assert not S.needs_translation("int *icon_adr") and not S.needs_translation("IS_reko_09.doc")


def test_abbreviation_slash_code_and_liter_per_hour():
    """30HU (29.10): «BSZG/30GA01» → «ИВС/30GA01» — аббревиатура переводится, код на месте (по замечанию на 50 строк);
    «……liter /h» → «л/ч»."""
    ga = frozenset({"BSZG"})
    assert T.check("30TK61P101 BSZG/30GA01", "30TK61P101 ИВС/30GA01", gloss_abbr=ga) == []
    assert T.check("30TK61P101 BSZG/30GA01", "30TK61P101 ИВС", gloss_abbr=ga) == ["нет кодов/чисел: BSZG/30GA01"]
    assert T.fix_units("ΔF = FTK – FTX = ……………liter /h") == "ΔF = FTK – FTX = ……………л/ч"
