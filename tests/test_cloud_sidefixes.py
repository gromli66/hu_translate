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


# ------------------------------------------------------------------ 3. портфолио и одноимённые документы комплекта
def _pdf(text):
    import fitz
    d = fitz.open()
    d.new_page().insert_text((72, 72), text)
    return d.tobytes()


def _portfolio(path, attachments):
    """attachments: [(имя файла вложения, байты)] — имена могут повторяться (ключи вложений — свои)."""
    import fitz
    d = fitz.open()
    d.new_page().insert_text((72, 72), "A legjobb eredmény érdekében a PDF-portfóliót érdemes az Acrobat-ban megnyitni.")
    for i, (name, data) in enumerate(attachments):
        d.embfile_add(f"k{i}", data, filename=name, ufilename=name)
    d.xref_set_key(d.pdf_catalog(), "Collection", "<< /Type /Collection >>")
    d.save(path)


def _inner_texts(path):
    import fitz
    d = fitz.open(path)
    return [fitz.open(stream=d.embfile_get(i), filetype="pdf")[0].get_text().strip() for i in range(d.embfile_count())]


def test_rebuilt_portfolio_is_idempotent(tmp_path, monkeypatch):
    """Повторная сборка (правки редактора пересобирают комплект) находила `30X_RU.pdf` — оболочку прошлой сборки —
    и клала вторую копию в `30X_RU_портфолио.pdf`; лишняя копия от прежней сборки убирается."""
    monkeypatch.setenv("HU_ROLES", "0")
    src, work = tmp_path / "in", tmp_path / "work"
    src.mkdir()
    _portfolio(src / "30X.PDF", [("30X feladat.pdf", _pdf("A szivattyu indul."))])
    P.extract([src], work)
    out = work / "out"
    out.mkdir()
    (out / "30X feladat_RU.pdf").write_bytes(_pdf("Nasos zapuskaetsya."))
    P.rebuild_portfolios(work, out)
    P.rebuild_portfolios(work, out)
    assert sorted(f.name for f in out.iterdir()) == ["30X feladat_RU.pdf", "30X_RU.pdf"]
    assert P._is_portfolio(out / "30X_RU.pdf") and _inner_texts(out / "30X_RU.pdf") == ["Nasos zapuskaetsya."]
    (out / "30X_RU_портфолио.pdf").write_bytes((out / "30X_RU.pdf").read_bytes())    # след прежней пересборки
    P.rebuild_portfolios(work, out)
    assert sorted(f.name for f in out.iterdir()) == ["30X feladat_RU.pdf", "30X_RU.pdf"]


def test_rebuilt_portfolio_beside_same_named_attachment_stays_beside(tmp_path, monkeypatch):
    """Вложение с именем оболочки (30Y.PDF ⊃ 30Y.pdf): перевод вложения — `30Y_RU.pdf`, оболочка — рядом; и при
    повторной сборке, и если перевод вложения в этот раз не собрался."""
    monkeypatch.setenv("HU_ROLES", "0")
    src, work = tmp_path / "in", tmp_path / "work"
    src.mkdir()
    _portfolio(src / "30Y.PDF", [("30Y.pdf", _pdf("A szivattyu indul."))])
    P.extract([src], work)
    out = work / "out"
    out.mkdir()
    (out / "30Y_RU.pdf").write_bytes(_pdf("Nasos."))
    for _ in range(2):
        P.rebuild_portfolios(work, out)
        assert sorted(f.name for f in out.iterdir()) == ["30Y_RU.pdf", "30Y_RU_портфолио.pdf"]
        assert not P._is_portfolio(out / "30Y_RU.pdf") and P._is_portfolio(out / "30Y_RU_портфолио.pdf")
    assert P.portfolio_target(tmp_path / "empty", "30Y", "_RU", ["30Y.pdf"]).name == "30Y_RU_портфолио.pdf"


def test_same_stem_docx_and_pdf_both_extracted(tmp_path, monkeypatch, capsys):
    """Doc.docx и Doc.pdf в одном комплекте давали один Doc.json, второй молча пропускался как «уже есть»."""
    from docx import Document
    monkeypatch.setenv("HU_ROLES", "0")
    src, work = tmp_path / "in", tmp_path / "work"
    src.mkdir()
    dx = Document()
    dx.add_paragraph("A DOCX szövege.")
    dx.save(src / "Doc.docx")
    (src / "Doc.pdf").write_bytes(_pdf("A PDF szovege."))
    P.extract([src], work)
    docs = P.load_docs(work)
    assert sorted(docs) == ["Doc", "Doc (2)"]                         # порядок прежний: первым — Doc.docx
    assert docs["Doc"]["file"].endswith("Doc.docx") and "DOCX" in docs["Doc"]["segs"][0]["text"]
    assert docs["Doc (2)"]["file"].endswith("Doc (2).pdf") and "PDF" in docs["Doc (2)"]["segs"][0]["text"]
    assert "имя «Doc» уже занято" in capsys.readouterr().out
    P.extract([src], work)                                            # повторный прогон — из готового, без копий
    assert sorted(P.load_docs(work)) == ["Doc", "Doc (2)"] and capsys.readouterr().out.count("уже есть") == 2


def test_same_named_attachments_do_not_overwrite_each_other(tmp_path, monkeypatch, capsys):
    monkeypatch.setenv("HU_ROLES", "0")
    src, work = tmp_path / "in", tmp_path / "work"
    src.mkdir()
    _portfolio(src / "30Z.PDF", [("30Z a.pdf", _pdf("Elso.")), ("30Z a.pdf", _pdf("Masodik."))])
    P.extract([src], work)
    docs = P.load_docs(work)
    assert sorted(docs) == ["30Z a", "30Z a (2)"]
    assert "Elso" in docs["30Z a"]["segs"][0]["text"] and "Masodik" in docs["30Z a (2)"]["segs"][0]["text"]
    assert "не одно с таким именем" in capsys.readouterr().out
    out = work / "out"
    out.mkdir()
    (out / "30Z a_RU.pdf").write_bytes(_pdf("Pervyj."))
    (out / "30Z a (2)_RU.pdf").write_bytes(_pdf("Vtoroj."))
    P.rebuild_portfolios(work, out)
    assert _inner_texts(out / "30Z_RU.pdf") == ["Pervyj.", "Vtoroj."]


def test_attachment_named_like_set_document_gets_own_translation(tmp_path, monkeypatch):
    """Вложение «30X feladat.pdf» и документ комплекта с тем же именем: оба переводятся, в портфолио — перевод вложения."""
    monkeypatch.setenv("HU_ROLES", "0")
    src, work = tmp_path / "in", tmp_path / "work"
    src.mkdir()
    (src / "30X feladat.pdf").write_bytes(_pdf("Kulon dokumentum."))
    _portfolio(src / "30X.PDF", [("30X feladat.pdf", _pdf("Belso dokumentum."))])
    P.extract([src], work)
    docs = P.load_docs(work)
    assert sorted(docs) == ["30X feladat", "30X feladat (2)"]
    assert "Kulon" in docs["30X feladat"]["segs"][0]["text"] and "Belso" in docs["30X feladat (2)"]["segs"][0]["text"]
    out = work / "out"
    out.mkdir()
    (out / "30X feladat_RU.pdf").write_bytes(_pdf("Otdelnyj."))
    (out / "30X feladat (2)_RU.pdf").write_bytes(_pdf("Vlozhennyj."))
    P.rebuild_portfolios(work, out)
    assert _inner_texts(out / "30X_RU.pdf") == ["Vlozhennyj."]
