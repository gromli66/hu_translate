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
