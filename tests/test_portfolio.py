# -*- coding: utf-8 -*-
"""PDF-портфолио: видимая страница — заглушка Adobe, документы — вложения (30RDGT.PDF, 06.10)."""
import json

import fitz

import pipeline as P


def pdf_bytes(text):
    d = fitz.open()
    d.new_page().insert_text((72, 72), text)
    return d.tobytes()


def make_portfolio(path, attachments):
    d = fitz.open()
    d.new_page().insert_text((72, 72), "A legjobb eredmény érdekében a PDF-portfóliót érdemes az Acrobat-ban megnyitni.")
    for name, data in attachments:
        d.embfile_add(name, data, filename=name)
    cat = d.pdf_catalog()
    d.xref_set_key(cat, "Collection", "<< /Type /Collection >>")
    d.save(path)


def test_portfolio_translates_attachments_not_cover(tmp_path):
    src = tmp_path / "in"
    src.mkdir()
    make_portfolio(src / "30X.PDF", [("30X feladat.pdf", pdf_bytes("A szivattyu indul.")),
                                     ("modj.pdf", pdf_bytes("Modositas.")), ("kep.png", b"\x89PNG")])
    work = tmp_path / "work"
    P.extract([src], work)
    stems = sorted(f.stem for f in P.doc_files(work))
    assert stems == ["30X feladat", "30X modj"]               # заглушка пропущена, картинка — не документ
    d = json.loads((work / "json" / "30X feladat.json").read_text(encoding="utf-8"))
    assert "szivattyu" in d["segs"][0]["text"] and d["file"].endswith("30X feladat.pdf")


def test_rebuilt_portfolio_carries_translations(tmp_path):
    src = tmp_path / "in"
    src.mkdir()
    make_portfolio(src / "30X.PDF", [("30X feladat.pdf", pdf_bytes("A szivattyu indul."))])
    work = tmp_path / "work"
    P.extract([src], work)
    out = work / "out"
    out.mkdir()
    (out / "30X feladat_RU.pdf").write_bytes(pdf_bytes("Nasos zapuskaetsya."))     # как будто после сборки
    P.rebuild_portfolios(work, out)
    d = fitz.open(out / "30X_RU.pdf")
    assert d.embfile_count() == 1
    inner = fitz.open(stream=d.embfile_get(0), filetype="pdf")
    assert "Nasos" in inner[0].get_text()


def test_plain_pdf_is_not_portfolio(tmp_path):
    p = tmp_path / "a.pdf"
    p.write_bytes(pdf_bytes("Sima dokumentum."))
    assert P.unpack_portfolio(p, tmp_path / "u") is None
