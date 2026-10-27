# -*- coding: utf-8 -*-
"""Пара языков проекта: "pair": "<источник>-<цель>" в project.json, языки hu, ru, en в обе стороны (22.10).

hu-ru — исходное направление (проверено на комплекте «Пакш»): его промпты, проверки, шлифовка, автоисправления
смысловой проверки — как были. ru-en — проверено на ТЗ тренажёра и договоре (свой промпт и check_en). Остальные пары
(ru-hu, en-ru, hu-en, en-hu) — общие промпты и проверки по языкам пары; автоисправления портала по умолчанию выключены.
Глоссарий проекта, записанный в обратную сторону (шапка «| Венгерский | Русский |» для пары ru-hu), разворачивается."""
import re

LANGS = ("hu", "ru", "en")
NAME_EN = {"hu": "Hungarian", "ru": "Russian", "en": "English"}
NAME_RU = {"hu": "венгерский", "ru": "русский", "en": "английский"}
TESS = {"hu": "hun", "ru": "rus+eng", "en": "eng"}             # языки Tesseract (русский — с латиницей кодов)
DEFAULT = ("hu", "ru")
# шапка таблицы глоссария: слово в первых двух ячейках → язык колонки
HEADER = {"венгерский": "hu", "magyar": "hu", "hungarian": "hu", "русский": "ru", "orosz": "ru", "russian": "ru",
          "английский": "en", "angol": "en", "english": "en"}
CYR = re.compile(r"[А-Яа-яЁё]")


def parse_pair(pair):
    """"ru-en" → ("ru", "en"); пусто → ("hu", "ru")."""
    if not pair:
        return DEFAULT
    m = re.fullmatch(r"\s*([a-z]{2})\s*-\s*([a-z]{2})\s*", str(pair).lower())
    if not m or m[1] not in LANGS or m[2] not in LANGS or m[1] == m[2]:
        raise ValueError(f"пара языков «{pair}»: нужна «источник-цель» из {', '.join(LANGS)}")
    return m[1], m[2]


def pair_of(prj):
    """Пара языков проекта (словарь project.json)."""
    return parse_pair((prj or {}).get("pair"))


def suffix(pair):
    """Суффикс файла перевода: _RU, _EN, _HU."""
    return "_" + parse_pair(pair)[1].upper()


def header_langs(cells):
    """Шапка таблицы глоссария → (язык 1-й колонки, язык 2-й) или None (не шапка)."""
    if len(cells) < 2:
        return None
    a, b = (HEADER.get(c.strip().lower()) for c in cells[:2])
    return (a, b) if a and b and a != b else None
