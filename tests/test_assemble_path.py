# -*- coding: utf-8 -*-
"""Один путь сборки PDF-исходника (assemble_all): DOCX по сегментам задания → PDF из DOCX; запасной путь один —
pdf2docx, и громко (info["attention"], строка в лог и в лист «нарушения» вычитки)."""
import fitz
from openpyxl import load_workbook

import assemble as A


def _job(tmp_path, **kw):
    pdf = tmp_path / "doc.pdf"
    d = fitz.open()
    d.new_page().insert_text((72, 72), "Szöveg")
    d.save(str(pdf))
    return dict({"file": str(pdf), "segs": [], "res": {}, "info": {}}, **kw)


def _stubs(monkeypatch, calls, build=None):
    def p2d(pdf, out, progress=None):
        calls.append(("pdf2docx", None))
        return {"converter": "pdf2docx", "pages": 1}

    def struct_docx(src, out, segs, res, rules=None, progress=None, mode="struct", overlay=None, **k):
        calls.append((mode, overlay))
        if build:
            build()
        return {"converter": mode}, []
    monkeypatch.setattr(A, "build_pdf", lambda src, out, segs, res, progress=None: {"translated": 0})
    monkeypatch.setattr(A, "docx_from_pdf", p2d)
    monkeypatch.setattr(A, "build_struct_docx", struct_docx)
    monkeypatch.setattr(A, "pdf_from_docx", lambda docx, pdf: {"from": "DOCX (LibreOffice)", "pages": 1})


def test_job_modes_take_one_path_with_translated_overlay(tmp_path, monkeypatch):
    import docx_layout as DL
    calls = []
    _stubs(monkeypatch, calls)
    monkeypatch.setattr(DL, "suits", lambda src, ocr: True)
    out = tmp_path / "out"
    for mode in ("struct", "text", "layout"):            # распознанный лист — тоже d["struct"] (pipeline.extract)
        info = A.assemble_all(tmp_path / "doc.json", _job(tmp_path, struct={"mode": mode}), out)
        assert "attention" not in info and info["pdf"]["from"] == "DOCX (LibreOffice)"
    # старое задание без d["struct"]: распознанный лист — по сегментам листа, без переведённой подложки (как прежде)
    A.assemble_all(tmp_path / "doc.json", _job(tmp_path), out)
    pdf = out / "doc_RU.pdf"
    assert calls == [("struct", pdf), ("text", pdf), ("layout", pdf), ("layout", None)]


def test_scan_is_pdf2docx_quietly_and_failed_build_loudly(tmp_path, monkeypatch):
    import docx_layout as DL
    calls = []
    monkeypatch.setattr(DL, "suits", lambda src, ocr: False)
    _stubs(monkeypatch, calls)
    info = A.assemble_all(tmp_path / "doc.json", _job(tmp_path), tmp_path / "out")
    assert calls == [("pdf2docx", None)] and "attention" not in info and "fallback" not in info["docx"]

    def boom():
        raise RuntimeError("сломалось")
    calls.clear()
    _stubs(monkeypatch, calls, boom)
    info = A.assemble_all(tmp_path / "doc.json", _job(tmp_path, struct={"mode": "text"}), tmp_path / "out")
    assert calls == [("text", tmp_path / "out" / "doc_RU.pdf"), ("pdf2docx", None)]
    assert info["docx"]["fallback"] == "RuntimeError: сломалось"
    assert len(info["attention"]) == 1 and "сломалось" in info["attention"][0]
    wb = load_workbook(tmp_path / "out" / "doc_вычитка.xlsx")
    assert any("сломалось" in str(c.value) for ws in wb for row in ws.iter_rows() for c in row)
