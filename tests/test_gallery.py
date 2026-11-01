# -*- coding: utf-8 -*-
"""tools/gallery.py: список страниц, подбор страниц оригинала и сборок, сборка картинок на синтетических PDF (без данных проекта)."""
import json
import sys
from pathlib import Path

import pymupdf
import pytest

ROOT = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(ROOT / "tools"))
import gallery as GL  # noqa: E402
import gate_baseline as G  # noqa: E402


def make_pdf(path, texts):
    """PDF по странице на текст; текст — латиница (шрифт по умолчанию без кириллицы)."""
    doc = pymupdf.open()
    for t in texts:
        pg = doc.new_page()
        pg.insert_textbox(pymupdf.Rect(40, 40, 550, 800), t, fontsize=10)
    doc.save(str(path))
    return pymupdf.open(str(path))


CODES = ["30TK11S001 30TK12S002 30TK13S003 30TK14S004", "30YP21S101 30YP22S102 30YP23S103", "30RL31S201 30RL32S202 30RL33S203 30RL34S204"]


# ---------------------------------------------------------------- список страниц
def write_list(path, rows):
    import yaml
    path.write_text(yaml.safe_dump({"pages": rows}, allow_unicode=True), encoding="utf-8")
    return path


def test_load_pages_valid(tmp_path):
    p = write_list(tmp_path / "l.yaml", [dict(doc="3PR10", page=11, cls=1, look="поля"), dict(doc="X", page=2, cls=6, look="схема", orig_page=3)])
    rows = GL.load_pages(p)
    assert [(r["doc"], r["page"], r["cls"], r["orig_page"]) for r in rows] == [("3PR10", 11, 1, None), ("X", 2, 6, 3)]


@pytest.mark.parametrize("bad", [dict(doc="X", page=0, cls=1, look="a"), dict(doc="X", page=1, cls=9, look="a"),
                                 dict(doc="X", page=1, cls=1, look="  "), dict(page=1, cls=1, look="a"), dict(doc="X", page="1", cls=1, look="a")])
def test_load_pages_rejects_bad_rows(tmp_path, bad):
    with pytest.raises(SystemExit):
        GL.load_pages(write_list(tmp_path / "l.yaml", [bad]))


def test_shipped_list_is_complete_and_unique():
    rows = GL.load_pages()
    assert 35 <= len(rows) <= 50
    assert {r["cls"] for r in rows} == set(GL.CLASSES)                      # все семь классов
    assert all(sum(1 for r in rows if r["cls"] == c) >= 4 for c in GL.CLASSES)
    keys = [(r["doc"], r["page"]) for r in rows]
    assert len(keys) == len(set(keys))


def test_filter_rows_by_name_and_class():
    rows = [dict(doc="3PR10_X", cls=1), dict(doc="30HU375GT", cls=4), dict(doc="3SZ23", cls=4)]
    assert [r["doc"] for r in GL.filter_rows(rows, ["класс:4"])] == ["30HU375GT", "3SZ23"]
    assert [r["doc"] for r in GL.filter_rows(rows, ["3pr10"])] == ["3PR10_X"]
    assert GL.filter_rows(rows, []) == rows


# ---------------------------------------------------------------- подбор
def test_best_index_prefers_overlap_then_proportion():
    t = {"a", "b", "c"}
    assert GL.best_index(t, [{"x"}, {"a", "b", "c"}, {"a"}], 0, 0.25) == (1, 1.0, "по совпадению", True)
    i, j, how, ok = GL.best_index(t, [{"x"}, {"y"}, {"z"}, {"w"}], 2, 0.25)
    assert (i, j, how, ok) == (2, 0.0, "по пропорции", False)
    i, j, how, ok = GL.best_index({"a"}, [{"a", "q"}, {"a", "q"}, {"a", "q"}], 2, 0.25)    # ничья — ближе к подсказке
    assert i == 2 and how == "по совпадению"
    assert GL.best_index({"a", "b", "c", "d"}, [{"a", "x", "y", "z", "w"}], 0, 0.25)[3] is False   # J мал — ненадёжно


def test_orig_for_by_codes(tmp_path):
    orig = make_pdf(tmp_path / "o.pdf", CODES)
    base = make_pdf(tmp_path / "b.pdf", [CODES[2], CODES[0]])
    i, how, ok = GL.orig_for(base[0], 0, len(base), orig)
    assert (i, ok) == (2, True) and "по кодам" in how
    assert GL.orig_for(base[1], 1, len(base), orig)[0] == 0


def test_orig_for_single_page_pinned_and_proportion(tmp_path):
    one = make_pdf(tmp_path / "o1.pdf", ["nothing here"])
    base = make_pdf(tmp_path / "b.pdf", ["alpha bravo", "charlie delta"])
    assert GL.orig_for(base[0], 0, 2, one) == (0, "единственная страница оригинала", True)
    three = make_pdf(tmp_path / "o3.pdf", ["aaa", "bbb", "ccc"])
    assert GL.orig_for(base[0], 0, 2, three, pinned=3) == (2, "задана вручную", True)
    i, how, ok = GL.orig_for(base[1], 1, 2, three)                              # кодов нет вовсе — по пропорции
    assert (i, how, ok) == (2, "по пропорции", False)


def test_orig_for_sheet_number_needs_a_shared_code(tmp_path):
    orig = make_pdf(tmp_path / "o.pdf", ["Oldalszám: 1 text", "Oldalszám: 2 text 30TK11S001", "Oldalszám: 3 text"])
    base = make_pdf(tmp_path / "b.pdf", ["Oldalszám: 2 words", "Oldalszám: 2 30TK11S001 30XX99S999"])
    i, how, ok = GL.orig_for(base[0], 0, 2, orig)
    assert (i, how.startswith("по номеру листа 2"), ok) == (1, True, False)       # номер указал лист, общего кода нет — ненадёжно
    orig2 = make_pdf(tmp_path / "o2.pdf", ["Oldalszám: 1 text", "Oldalszám: 2 text 30TK11S001 30XX99S999 30YY88S888", "Oldalszám: 3"])
    base2 = make_pdf(tmp_path / "b2.pdf", ["Oldalszám: 2 30TK11S001", "Oldalszám: 3"])
    i, how, ok = GL.orig_for(base2[0], 0, 2, orig2)
    assert (i, ok) == (1, True)


def test_page_in_build_by_text(tmp_path):
    base = make_pdf(tmp_path / "b.pdf", ["alpha bravo charlie delta echo foxtrot 30TK11S001 30TK12S002 30TK13S003"])
    other = make_pdf(tmp_path / "x.pdf", ["zulu yankee xray whiskey victor uniform", "alpha bravo charlie delta echo foxtrot 30TK11S001 30TK12S002 30TK13S003", "tango sierra"])
    i, how, ok = GL.page_in_build(base[0], 0, 1, other)
    assert (i, ok) == (1, True) and "по тексту" in how
    assert GL.page_in_build(base[0], 0, 1, make_pdf(tmp_path / "s.pdf", ["x"]))[2] is True   # единственная страница


# ---------------------------------------------------------------- файлы
def test_resolve_stem():
    pdfs = {"3PR10_3PR10_HITELESITETT": 1, "30BLM1M2GT": 1, "30BLM1M2GT_CIML": 1, "30BLM1M2GT_MODJ": 1}
    assert GL.resolve_stem("3PR10", pdfs) == "3PR10_3PR10_HITELESITETT"
    assert GL.resolve_stem("30BLM1M2GT", pdfs) == "30BLM1M2GT"                   # точное имя важнее префикса
    with pytest.raises(SystemExit, match="несколько"):
        GL.resolve_stem("30BLM1M2GT_", pdfs)
    with pytest.raises(SystemExit, match="нет PDF"):
        GL.resolve_stem("NOPE", pdfs)


def test_pdfs_of_strips_language_suffix(tmp_path):
    (tmp_path / "out").mkdir()
    make_pdf(tmp_path / "out" / "A_RU.pdf", ["x"])
    make_pdf(tmp_path / "B_EN.pdf", ["x"])
    assert set(GL.pdfs_of(tmp_path)) == {"A", "B"}


def test_find_orig_from_json_file_then_orig_dir(tmp_path):
    root = tmp_path / "root"
    (root / "data" / "jobs" / "u" / "in").mkdir(parents=True)
    make_pdf(root / "data" / "jobs" / "u" / "in" / "A.PDF", ["x"])
    js = tmp_path / "A.json"
    js.write_text(json.dumps({"file": "/app/data/jobs/u/in/A.PDF"}), encoding="utf-8")
    assert GL.find_orig(root, "A", tmp_path / "none", {"A": js}) == root / "data" / "jobs" / "u" / "in" / "A.PDF"
    od = tmp_path / "orig"
    od.mkdir()
    make_pdf(od / "B.PDF", ["x"])
    assert GL.find_orig(root, "B", od, {}) == od / "B.PDF"
    assert GL.find_orig(root, "C", od, {}) is None


# ---------------------------------------------------------------- картинки
def test_wrap_breaks_long_text():
    from PIL import Image, ImageDraw
    d = ImageDraw.Draw(Image.new("RGB", (10, 10)))
    lines = GL.wrap(d, "слово " * 40, GL.font(13), 200)
    assert len(lines) > 3 and all(d.textlength(ln, font=GL.font(13)) <= 200 or " " not in ln for ln in lines)


def test_make_gallery_end_to_end(tmp_path, monkeypatch):
    texts = ["alpha bravo charlie delta echo " + CODES[k] for k in range(3)]
    base = tmp_path / "base"
    after = tmp_path / "after"
    base.mkdir()
    after.mkdir()
    make_pdf(base / "DOC1_RU.pdf", texts)
    make_pdf(after / "DOC1_RU.pdf", [texts[1], texts[2], texts[0], "extra page"])         # страницы переставлены и добавлена лишняя
    orig = tmp_path / "orig"
    orig.mkdir()
    make_pdf(orig / "DOC1.PDF", ["Oldalszám: %d %s" % (k + 1, CODES[k]) for k in range(3)])
    monkeypatch.setattr(G, "base_dir", lambda: base)
    rows = [dict(doc="DOC1", page=2, cls=3, look="шапка", orig_page=None), dict(doc="DOC1", page=1, cls=3, look="поля", orig_page=None),
            dict(doc="DOC1", page=3, cls=3, look="третья", orig_page=None), dict(doc="DOC1", page=1, cls=7, look="титул", orig_page=None)]
    out = tmp_path / "gal"
    (out / "pages").mkdir(parents=True)
    (out / "class3_headers_9.png").write_bytes(b"old")                                    # старый лист-сводка стирается
    said = []
    sheets, bad = GL.make_gallery([str(base), str(after)], out, rows, tmp_path / "noroot", orig, say=said.append)
    assert bad == 0
    names = sorted(s.name for s in sheets)
    assert names == ["class3_headers_1.png", "class7_portfolio_1.png"]                    # 3 пары класса 3 на одном листе, 1 пара класса 7
    assert not (out / "class3_headers_9.png").exists()
    assert len(list((out / "pages").glob("*.png"))) == 4
    from PIL import Image
    im = Image.open(out / "pages" / "01_DOC1_p2.png")
    assert im.width > 3 * GL.CELL_W and im.height > GL.CELL_H                              # оригинал | A | B
    index = (out / "index.md").read_text(encoding="utf-8")
    assert "Класс 3. шапки и номера листов" in index and "DOC1" in index and "pages/01_DOC1_p2.png" in index
    # у B страница найдена по тексту: базовая стр. 2 (texts[1]) лежит в after на стр. 1
    pg_b = GL.page_in_build(pymupdf.open(str(base / "DOC1_RU.pdf"))[1], 1, 3, pymupdf.open(str(after / "DOC1_RU.pdf")))
    assert pg_b[0] == 0


def test_make_gallery_marks_unreliable_pair(tmp_path, monkeypatch):
    base = tmp_path / "base"
    base.mkdir()
    make_pdf(base / "D_RU.pdf", ["zzz one", "zzz two"])
    orig = tmp_path / "orig"
    orig.mkdir()
    make_pdf(orig / "D.PDF", ["aaa", "bbb", "ccc"])
    monkeypatch.setattr(G, "base_dir", lambda: base)
    said = []
    _, bad = GL.make_gallery([str(base)], tmp_path / "g", [dict(doc="D", page=1, cls=1, look="x", orig_page=None)], tmp_path / "r", orig, say=said.append)
    assert bad == 1 and "НЕНАДЁЖНО" in said[0]
    assert "ненадёжно" in (tmp_path / "g" / "index.md").read_text(encoding="utf-8")


def test_main_rejects_three_builds(tmp_path):
    with pytest.raises(SystemExit, match="не больше двух"):
        GL.main(["a", "b", "c", "--out", str(tmp_path)])
