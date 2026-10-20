# -*- coding: utf-8 -*-
"""Обратное направление RU→EN (проект "pair": "ru-en"; ТЗ тренажёра ТЭС, 20.10.2026) и поле номера страницы в DOCX."""
import zipfile

from lxml import etree

import glossary as G
import segments as S
import translate as T

W = S.W


def entry(ru, en="term", note=""):
    e = G.Entry(ru, en, note, "")
    e.alts = G._alts(ru, en)
    return e


# ------------------------------------------------------------------ глоссарий с русским источником
def test_russian_term_found_in_inflected_text():
    # прежде слова искались только латиницей — русские термины не находились вовсе
    found = G.find("При вводе отказа происходит межфазное короткое замыкание",
                   [entry("ввод отказа"), entry("межфазное короткое замыкание"), entry("аварийный останов")])
    assert [e.hu for e in found] == ["межфазное короткое замыкание", "ввод отказа"]


def test_russian_word_matches_by_whole_word_and_stem():
    # «останов»: основа «остан» + «ова» — не окончание, засчитывается всё слово + «а»; «станция» → «станции»
    assert G._word_hit("останов", "останова") and G._word_hit("станция", "станции")
    assert not G._word_hit("останов", "остановка") and not G._word_hit("тренажер", "тренажерный")
    assert G._word_hit("тренажер", "тренажёра")                   # ё = е


def test_russian_header_row_is_not_an_entry(tmp_path):
    f = tmp_path / "g.md"
    f.write_text("## Сокращения\n\n| Русский | Английский | Комментарий |\n|---|---|---|\n| АТ | AS | Analytical Simulator |\n",
                 encoding="utf-8")
    assert [(e.hu, e.ru) for e in G.load(f)[1]] == [("АТ", "AS")]


def test_cyrillic_segment_needs_translation():
    assert S.needs_translation("Техническое задание") and not S.needs_translation("1.1.1.3 ")
    assert S.needs_translation("№ п/п ")                          # шапка таблицы оставалась по-русски


# ------------------------------------------------------------------ проверки английского перевода
def test_check_en_decimal_point_and_units_are_not_issues():
    assert T.check_en("После ввода отказа, насос отключается. 0,4 кВ, 600МВт",
                      "After the malfunction is inserted, the pump trips. 0.4 kV, 600 MW") == []


def test_check_en_flags_cyrillic_lost_latin_and_numbers():
    iss = T.check_en("Тренажер Thai Binh 2 должен работать 15 минут", "Тренажер Thai shall operate for minutes")
    assert "латиница потеряна: Binh" in iss and "нет кодов/чисел: 15, 2" in iss
    assert any(i.startswith("кириллица осталась") for i in iss)
    assert T.check_en("Тренажер должен работать", "Тренажер должен работать") == ["не переведено"]


def test_translator_ru_en_uses_english_prompt_and_checks(tmp_path):
    g = tmp_path / "g.md"
    g.write_text("## Правила применения\n\n- shall\n\n## Т\n\n| Русский | Английский |\n|---|---|\n| тренажер | simulator |\n",
                 encoding="utf-8")
    tr = T.Translator(glossary=str(g), pair="ru-en", domain="You translate a simulator specification.")
    assert tr.system.startswith("You translate a simulator specification.") and "Венгер" not in tr.system
    assert "Glossary terms for this fragment:\n- тренажер → simulator" in tr._prompt([{"text": "Тренажер готов"}], [], "")
    assert tr.full_check("Тренажер готов", "The simulator is ready") == []


# ------------------------------------------------------------------ поле номера страницы
def _docx(tmp_path, body):
    f = tmp_path / "f.docx"
    xml = (f'<w:document xmlns:w="{W[1:-1]}"><w:body>{body}</w:body></w:document>')
    with zipfile.ZipFile(f, "w") as z:
        z.writestr("word/document.xml", xml)
    return f


def test_page_field_result_is_not_text(tmp_path):
    # «Стр.» + PAGE: результат поля «24» шёл в сегмент, перевод «Page 24» + пересчитанное поле = «Page 2402»
    f = _docx(tmp_path, '<w:p><w:r><w:t>Стр.</w:t></w:r><w:r><w:fldChar w:fldCharType="begin"/></w:r>'
                        '<w:r><w:instrText> PAGE \\* MERGEFORMAT </w:instrText></w:r>'
                        '<w:r><w:fldChar w:fldCharType="separate"/></w:r><w:r><w:t>24</w:t></w:r>'
                        '<w:r><w:fldChar w:fldCharType="end"/></w:r></w:p>'
                        '<w:p><w:r><w:t>см. стр. </w:t></w:r><w:r><w:fldChar w:fldCharType="begin"/></w:r>'
                        '<w:r><w:instrText> PAGEREF _Toc1 \\h </w:instrText></w:r>'
                        '<w:r><w:fldChar w:fldCharType="separate"/></w:r><w:r><w:t>7</w:t></w:r>'
                        '<w:r><w:fldChar w:fldCharType="end"/></w:r></w:p>')
    segs = S.docx_segments(f)
    assert [s["text"] for s in segs] == ["Стр.", "см. стр. 7"]       # ссылка на страницу (PAGEREF) — как раньше
