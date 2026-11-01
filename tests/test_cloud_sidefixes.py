# -*- coding: utf-8 -*-
"""Побочные дефекты, найденные чтением кода (облачная ветка cloud-sidefixes, 01.11): каждый тест падал на прежнем коде."""
import inspect
import sys
from pathlib import Path

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


# ------------------------------------------------------------------ 4. сборка: запасной DOCX и короткий PDF
def _pages_pdf(path, texts):
    """PDF: лист на каждый текст; пустой текст — лист без текста (только рамка)."""
    import fitz
    d = fitz.open()
    for t in texts:
        pg = d.new_page()
        if t:
            pg.insert_text((72, 72), t, fontname="helv")
        else:
            pg.draw_rect(fitz.Rect(100, 100, 300, 300))
    d.save(path)


def test_p2d_page_failing_every_setting_is_blank_page_not_lost_docx(tmp_path, monkeypatch, capsys):
    """Лист, на котором pdf2docx падает при всех трёх настройках, шёл в склейку несозданным файлом: Document() падал,
    и запасной DOCX терялся целиком («ERR …»). Теперь лист — пустым разделом и громкой отметкой SKIP."""
    import pdf2docx
    import assemble as A
    from docx import Document
    pdf, out = tmp_path / "a.pdf", tmp_path / "a.docx"
    _pages_pdf(pdf, ["Oldal egy", "Oldal ketto", "Oldal harom"])
    Real = pdf2docx.Converter

    class Conv(Real):
        def convert(self, docx_filename=None, *a, pages=None, **kw):
            if pages == [1]:
                raise ValueError("Failed to merge docx_cell")
            return super().convert(docx_filename, *a, pages=pages, **kw)
    monkeypatch.setattr(pdf2docx, "Converter", Conv)
    A._p2d_child(str(pdf), str(out))
    said = capsys.readouterr().out
    assert "OK 3" in said and "SKIP 1 ValueError: Failed to merge docx_cell" in said
    xml = Document(str(out)).element.body.xml
    assert "Oldal egy" in xml and "Oldal harom" in xml and "Oldal ketto" not in xml
    assert xml.count("<w:sectPr") == 3                           # границы листов — на месте


def test_docx_from_pdf_reports_skipped_pages(tmp_path, monkeypatch, capsys):
    import subprocess
    import assemble as A
    out = tmp_path / "a.docx"

    class Proc:
        def __init__(self, cmd, **kw):
            Path(cmd[-1]).write_bytes(b"docx")
            self.stdout, self.stderr, self.returncode = iter(["PAGE 0 2\n", "SKIP 1 ValueError: boom\n", "OK 2\n"]), None, 0

        def wait(self):
            return 0
    monkeypatch.setattr(subprocess, "Popen", Proc)
    info = A.docx_from_pdf(tmp_path / "a.pdf", out)
    assert info["pages"] == 2 and info["failed_pages"] == [2] and "boom" in info["failed_why"]
    assert "листы 2 из 2 pdf2docx не сконвертировал" in capsys.readouterr().out


def test_pdf_in_place_shorter_than_original_is_reported(tmp_path, monkeypatch, capsys):
    """PDF «на месте» отдаёт только листы с сегментами: лист без текста выпадал, и без LibreOffice (или для скана) PDF
    перевода молча выходил короче оригинала. Поведение то же — сигнал в сведениях сборки и в логе."""
    import fitz
    import assemble as A
    import segments as S
    monkeypatch.setattr(S, "OCR_DIR", tmp_path / "ocr")
    monkeypatch.setattr(S, "_tess_words", lambda png, dpi: [])          # листы с текстовым слоем: Tesseract не нужен
    pdf = tmp_path / "Doc.pdf"
    _pages_pdf(pdf, ["Elso oldal szovege.", "", "Harmadik oldal szovege."])
    segs, _ = S.pdf_segments(pdf, ocr=False, shots=False)
    res = {str(s["id"]): {"ru": "Перевод листа.", "issues": []} for s in segs}
    info = A.build_pdf(pdf, tmp_path / "o.pdf", segs, res)
    assert fitz.open(tmp_path / "o.pdf").page_count == 2                # поведение прежнее
    assert info["pages"] == 2 and info["pages_of"] == 3 and info["pages_missing"] == [2]
    d = {"file": str(pdf), "segs": segs, "res": res, "info": {}, "struct": None, "docx": None}
    A.assemble_all(tmp_path / "Doc.json", d, tmp_path / "out")
    assert "Doc_RU.pdf — 2 из 3 листов оригинала, без текста не вошли: 2" in capsys.readouterr().out
    full = tmp_path / "Full.pdf"
    _pages_pdf(full, ["Elso oldal.", "Masodik oldal."])
    segs, _ = S.pdf_segments(full, ocr=False, shots=False)
    assert "pages_missing" not in A.build_pdf(full, tmp_path / "f.pdf", segs, {str(s["id"]): {"ru": "Лист."} for s in segs})


# ------------------------------------------------------------------ 5. проверка оформления: таблица в ячейке
def test_table_checks_pair_top_level_tables_after_nested_one(tmp_path):
    """Ф9/Ф10: блоки таблиц шли вместе с вложенными, таблицы DOCX — только верхнего уровня: после первой таблицы
    в ячейке сверка сдвигалась на одну (ложные «ширина»/«заливка шапки» у следующей), ширина внешней складывалась
    с ячейками вложенной."""
    import docx_build as B
    import docx_profile as F
    from test_struct_docx import LAYOUT, PROFILE, p, png
    prof = dict(PROFILE, page={"left": LAYOUT["left"], "right": PROFILE["page"]["right"]})
    inner = {"k": "table", "frame": {"xs": [100, 200]},
             "rows": [[{"blocks": [p("Belső")], "th": True}], [{"blocks": [p("Belső 2")], "th": False}]]}
    outer = {"k": "table", "frame": {"xs": [62.4, 535.6]},
             "rows": [[{"blocks": [p("Külső"), inner], "th": False}, {"blocks": [p("Jobb")], "th": False}]]}
    second = {"k": "table", "frame": {"xs": [62.4, 535.6], "fill": "D9D9D9"},
              "rows": [[{"blocks": [p("Fej")], "th": True}], [{"blocks": [p("Test")], "th": False}]]}
    mem = B.Memory([(t, t) for t in ("Belső", "Belső 2", "Külső", "Jobb", "Fej", "Test")])
    out = tmp_path / "t.docx"
    B.Builder(mem, lambda img: png(), prof).build([outer, second], str(out), LAYOUT)
    tab = lambda v: [x for x in v if x["rule"] in ("Ф9", "Ф10")]
    assert tab(F.check(str(out), prof, [outer, second], LAYOUT)) == []
    # настоящее расхождение по-прежнему видно — и у той таблицы, где оно есть
    other = dict(second, frame={"xs": [62.4, 535.6], "fill": "FFFF00"})
    assert [x["where"] for x in tab(F.check(str(out), prof, [outer, other], LAYOUT))] == ["таблица 2: заливка шапки"]
