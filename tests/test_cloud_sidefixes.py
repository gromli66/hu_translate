# -*- coding: utf-8 -*-
"""Побочные дефекты, найденные чтением кода (облачная ветка cloud-sidefixes, 01.11): каждый тест падал на прежнем коде."""
import inspect
import sys

import pipeline as P


# ------------------------------------------------------------------ 1. CLI: язык источника — параметр src_lang
def test_cli_extract_runs_with_pipeline_signature(tmp_path, monkeypatch):
    """translate_komplekt run/extract/terms вызывал P.extract(..., lang=…) — у extract параметр src_lang: TypeError
    на первом же шаге любой команды с извлечением (сервер вызывал верно)."""
    import translate_komplekt as TK
    (tmp_path / "in").mkdir()
    monkeypatch.setattr(sys, "argv", ["translate_komplekt.py", "extract", str(tmp_path / "in"),
                                      "--work", str(tmp_path / "w")])
    TK.main()
    assert (tmp_path / "w" / "json").is_dir()


def test_cli_terms_passes_project_source_language(tmp_path, monkeypatch):
    import translate_komplekt as TK
    real, seen = inspect.signature(P.extract), {}

    def extract(*a, **kw):
        seen.update(real.bind(*a, **kw).arguments)        # чужое имя параметра — TypeError, как у настоящего
    monkeypatch.setattr(P, "load_project", lambda path: {"pair": "ru-en", "glossary": "g.md"})
    monkeypatch.setattr(P, "extract", extract)
    monkeypatch.setattr(P, "terms", lambda *a, **kw: None)
    monkeypatch.setattr(sys, "argv", ["translate_komplekt.py", "terms", str(tmp_path), "--project", "p.json",
                                      "--work", str(tmp_path / "w")])
    TK.main()
    assert seen["src_lang"] == "ru"


# ------------------------------------------------------------------ 2. надписи на картинках: источник en
def test_every_source_language_has_shot_settings():
    """Пара с источником en (en-ru, en-hu): shots.LANGS без «en» — KeyError в Tesseract, ловился как сбой картинки,
    и надписи скриншотов и рисунков молча оставались непереведёнными."""
    import langs as L
    import shots as SH
    for lang in L.LANGS:
        assert SH.LANGS[lang]["tess"] == L.TESS[lang]
        assert set(SH.LANGS[lang]["other"]) <= set(SH.LANG_NAMES) and lang not in SH.LANGS[lang]["other"]
    pr = SH.check_prompt("en", SH.PICTURE)
    assert "английским словарём" in pr and '"lang": "en"|' in pr


def test_english_source_labels_read_and_only_english_translated(tmp_path, monkeypatch):
    import re
    import llm
    import shots as SH
    from test_docx_shots import _echo, _picture
    img = SH._open_raster(_picture(["Power supply", "Съемный кожух"], size=18))[0]
    labels = SH.pic_labels(img, tmp_path / "t.png", "en")                 # Tesseract со словарём eng
    assert any("Power supply" in lb["t"] for lb in labels)
    monkeypatch.setattr(llm, "chat", _echo(lambda t: "ru" if re.search("[а-яА-Я]", t) else "en"))
    good = SH.check_labels(b"png", labels, "t", "en", SH.PICTURE, merge=True)
    assert good and all(not re.search("[а-яА-Я]", lb["t"]) for lb in good)
    assert any("Power supply" in lb["t"] for lb in good)
