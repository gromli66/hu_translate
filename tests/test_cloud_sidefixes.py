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


# ------------------------------------------------------------------ 6. метки ролей портала: частичный сбой виден
def test_partial_roles_failure_is_visible(tmp_path, monkeypatch, capsys):
    """struct["roles"] = True, если метки получил хоть один сегмент: сбой портала на части листов не был виден."""
    import segments as S
    src = tmp_path / "in"
    src.mkdir()
    (src / "Doc.pdf").write_bytes(_pdf("Szoveg."))

    def fake(f, roles=False, **kw):
        segs = [{"id": 4 * pg + k, "text": f"Mondat {pg}.{k}", "page": pg, "loc": {"roles": True} if roles and pg == 0 else {}}
                for pg in range(3) for k in range(4)]
        segs.append({"id": 12, "text": "Mentés", "page": 1, "loc": {"shot": {"xref": 5}}})   # надпись скриншота — не в счёт
        return segs, {"ocr_pages": []}
    monkeypatch.setattr(S, "pdf_segments", fake)
    monkeypatch.setenv("HU_ROLES", "1")
    P.extract([src], tmp_path / "w")
    st = P.load_docs(tmp_path / "w")["Doc"]["struct"]
    assert st["mode"] == "text" and st["roles"] is True                          # как прежде
    assert st["roles_share"] == 0.333 and st["roles_pages"] == [1, 3]
    assert "роли строк порталом: да — у 33% сегментов, листов 1 из 3 (часть листов — правилами)" in capsys.readouterr().out
    assert P.roles_note(P.roles_seen([{"page": 0, "loc": {}}])) == "нет (сбой портала)"


# ------------------------------------------------------------------ 7. размеченный PDF: вложенные теги, нестандартный тип
def _tagged_pdf(path, content, tree):
    """Размеченный PDF из одного листа: content — поток листа (шрифт /helv), tree — [(тип, [MCID | поддерево])]."""
    import fitz
    d = fitz.open()
    pg = d.new_page(width=595, height=842)
    pg.insert_text((72, 72), "x", fontname="helv")             # шрифт /helv — в ресурсы листа
    pg.clean_contents()
    d.update_stream(pg.get_contents()[0], content.encode("latin-1"))
    page = pg.xref
    root, top = d.get_new_xref(), d.get_new_xref()

    def elem(t, kids, parent):
        x = d.get_new_xref()
        ks = [str(k) if isinstance(k, int) else f"{elem(k[0], k[1], x)} 0 R" for k in kids]
        d.update_object(x, f"<< /Type /StructElem /S /{t} /P {parent} 0 R /Pg {page} 0 R /K [{' '.join(ks)}] >>")
        return x
    kids = [elem(t, k, top) for t, k in tree]
    d.update_object(top, f"<< /Type /StructElem /S /Document /P {root} 0 R /K [{' '.join(f'{x} 0 R' for x in kids)}] >>")
    d.update_object(root, f"<< /Type /StructTreeRoot /K {top} 0 R >>")
    d.xref_set_key(d.pdf_catalog(), "StructTreeRoot", f"{root} 0 R")
    d.xref_set_key(d.pdf_catalog(), "MarkInfo", "<< /Marked true >>")
    d.save(path)


def _tj(x, y, t):
    return f"BT /helv 11 Tf {x} {y} Td ({t}) Tj ET "


def test_text_inside_nested_tag_without_mcid_stays_in_paragraph(tmp_path):
    """«/Span <</ActualText …>> BDC» без MCID внутри «/P <</MCID 0>>»: pdfplumber обнулял MCID — слово внутри
    пропадало молча, текст после него до конца абзаца уходил в артефакты. Артефакт внутри абзаца — по-прежнему не текст."""
    import docx_struct as D
    c = ("/P <</MCID 0>> BDC " + _tj(72, 700, "Elso mondat") +
         "/Span <</ActualText (kotojel)>> BDC " + _tj(150, 700, "kotojel") + "EMC " + _tj(200, 700, "utana.") +
         "/Artifact BMC " + _tj(500, 700, "Lapszam") + "EMC " + _tj(250, 700, "Vege.") + "EMC "
         "/P <</MCID 1>> BDC " + _tj(72, 650, "Masodik bekezdes.") + "EMC")
    _tagged_pdf(tmp_path / "t.pdf", c, [("P", [0]), ("P", [1])])
    blocks, _ = D.load(str(tmp_path / "t.pdf"))
    texts = [D.text_of(b["runs"]) for b in blocks if b["k"] == "p"]
    assert texts[-1] == "Masodik bekezdes."
    first = " ".join(texts[:-1])
    assert all(w in first for w in ("Elso mondat", "kotojel", "utana.", "Vege.")) and "Lapszam" not in first


def test_nonstandard_element_with_own_text_keeps_nested_table(tmp_path):
    """Элемент нестандартного типа (рамка текста) со своим MCID и таблицей внутри: весь текст поддерева сливался в
    один абзац — таблица шла строкой «Keret Bal Jobb»."""
    import docx_struct as D
    c = ("/P <</MCID 0>> BDC " + _tj(72, 700, "Keret felirata.") + "EMC "
         "/P <</MCID 1>> BDC " + _tj(72, 650, "Bal") + "EMC /P <</MCID 2>> BDC " + _tj(300, 650, "Jobb") + "EMC")
    cell = lambda m: ("TD", [("P", [m])])
    _tagged_pdf(tmp_path / "t.pdf", c, [("Textbox", [0, ("Table", [("TR", [cell(1), cell(2)])])])])
    blocks, _ = D.load(str(tmp_path / "t.pdf"))
    assert [b["k"] for b in blocks] == ["p", "table"]
    assert D.text_of(blocks[0]["runs"]).strip() == "Keret felirata."
    assert [D.text_of(c["blocks"][0]["runs"]).strip() for c in blocks[1]["rows"][0]] == ["Bal", "Jobb"]


# ------------------------------------------------------------------ попутное: метки ролей — каждому абзацу листа
def test_role_labels_reach_every_paragraph_of_page(tmp_path, monkeypatch):
    """С 6388eca (29.10) строки «метки портала → сегмент» стояли вне цикла по абзацам: метку получал только последний
    сегмент текстового листа (с ролью последнего абзаца), распознанные листы — ни один; заголовки и пункты по меткам
    в docx_layout не доходили."""
    import fitz
    import llm
    import roles as R
    import segments as S
    lines = ["1. Bevezetés", "Az első bekezdés szövege itt van, elég hosszú.", "A második bekezdés szövege is itt van."]
    d = fitz.open()
    pg = d.new_page()
    for i, t in enumerate(lines):
        pg.insert_text((72, 100 + 60 * i), t, fontsize=12)
    d.save(tmp_path / "r.pdf")
    monkeypatch.setattr(S, "OCR_DIR", tmp_path / "ocr")
    lab = lambda pages: {p: (["x"], {"x": 1}) for p in pages}
    monkeypatch.setattr(R, "doc_labels", lambda path, pages, workers=4: lab(pages))
    monkeypatch.setattr(R, "frag_labels", lambda frags, size, workers=4: lab(frags))

    def apply(pars, lines_, labels, join_lines):
        for p in pars:
            p["role"] = "heading" if p["text"].startswith("1.") else "body"
        return pars
    monkeypatch.setattr(R, "apply", apply)
    monkeypatch.setattr(llm, "chat", lambda messages, images=None, max_tokens=0, tag="":
                        {"text": "\n\n".join(lines), "finish": "stop"})      # чтение скана порталом
    for ocr in (False, True):                                                   # лист с текстовым слоем и скан
        segs, info = S.pdf_segments(tmp_path / "r.pdf", roles=True, shots=False, ocr=ocr, dpi_tess=150, dpi_portal=72)
        assert bool(info.get("ocr_pages")) == ocr and len(segs) == 3
        assert [s["loc"].get("roles") for s in segs] == [True] * 3
        assert [s["loc"].get("role") for s in segs] == ["heading", "body", "body"]


def test_old_doc_attachment_replaced_by_its_docx_translation(tmp_path):
    """Вложение .doc переводится через DOCX (<имя>_RU.docx), а сборка портфолио искала <имя>_RU.doc — такое вложение
    оставалось в переведённом портфолио непереведённым."""
    import io
    import fitz
    from docx import Document
    work, out = tmp_path / "work", tmp_path / "work" / "out"
    _portfolio(tmp_path / "30W.PDF", [("30W kezi.doc", b"\xd0\xcf\x11\xe0 old word"), ("30W lap.pdf", _pdf("Lap."))])
    P.unpack_portfolio(tmp_path / "30W.PDF", work / "unpacked" / "30W")
    out.mkdir(parents=True)
    dx = Document()
    dx.add_paragraph("Ручное управление.")
    dx.save(out / "30W kezi_RU.docx")
    (out / "30W lap_RU.pdf").write_bytes(_pdf("List."))
    P.rebuild_portfolios(work, out)
    d = fitz.open(out / "30W_RU.pdf")
    files = {P.attachment_name(d.embfile_info(i), ""): d.embfile_get(i) for i in range(d.embfile_count())}
    assert sorted(files) == ["30W kezi.docx", "30W lap.pdf"]
    assert Document(io.BytesIO(files["30W kezi.docx"])).paragraphs[0].text == "Ручное управление."
    assert "30W kezi.docx" in d[0].get_text().replace("\n", " ")                 # и в перечне на обложке


def test_replacing_later_attachment_keeps_earlier_untouched(tmp_path, monkeypatch):
    """embfile_del PyMuPDF 1.28 удаляет не ту пару массива имён (idx и idx+1 вместо 2·idx и 2·idx+1): замена вложения
    не первым по счёту отдавала предыдущему чужое содержимое — картинка или непереведённый документ перед переведённым
    пропадали из портфолио, а венгерский оригинал оказывался в нём дважды. Описание вложения теперь сохраняется."""
    import fitz
    monkeypatch.setenv("HU_ROLES", "0")
    png = fitz.Pixmap(fitz.csRGB, fitz.IRect(0, 0, 4, 4), False).tobytes("png")
    src, work = tmp_path / "in", tmp_path / "work"
    src.mkdir()
    d = fitz.open()
    d.new_page().insert_text((72, 72), "A legjobb eredmény érdekében…")
    d.embfile_add("k0", png, filename="30V kep.png", ufilename="30V kep.png", desc="séma")
    d.embfile_add("k1", _pdf("Nem forditott."), filename="30V regi.pdf", ufilename="30V regi.pdf")
    d.embfile_add("k2", _pdf("A szivattyu indul."), filename="30V feladat.pdf", ufilename="30V feladat.pdf", desc="utasítás")
    d.xref_set_key(d.pdf_catalog(), "Collection", "<< /Type /Collection >>")
    d.save(src / "30V.PDF")
    P.unpack_portfolio(src / "30V.PDF", work / "unpacked" / "30V")
    out = work / "out"
    out.mkdir(parents=True)
    (out / "30V feladat_RU.pdf").write_bytes(_pdf("Nasos zapuskaetsya."))          # «30V regi» — без перевода
    P.rebuild_portfolios(work, out)
    r = fitz.open(out / "30V_RU.pdf")
    got = [(P.attachment_name(r.embfile_info(i), ""), r.embfile_info(i).get("description")) for i in range(r.embfile_count())]
    was = fitz.open(src / "30V.PDF")
    assert got == [(P.attachment_name(was.embfile_info(i), ""), was.embfile_info(i).get("description")) for i in range(3)]
    assert got[0] == ("30V kep.png", "séma") and got[2] == ("30V feladat.pdf", "utasítás")
    assert r.embfile_get(0) == png
    assert fitz.open(stream=r.embfile_get(1), filetype="pdf")[0].get_text().strip() == "Nem forditott."
    assert fitz.open(stream=r.embfile_get(2), filetype="pdf")[0].get_text().strip() == "Nasos zapuskaetsya."
