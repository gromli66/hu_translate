# Вёрстка без языка (22.10): лист-плакат и таблица-график — с переведённого PDF целиком; альбомные листы — альбомными
# разделами; стенд tools/layout_any.py (поддельный перевод любой длины).
import fitz
from docx import Document

import docx_build as B
import docx_layout as DL

from test_struct_docx import LAYOUT, PROFILE, p, png


def _seg(i, x0, y0, x1, text="címke"):
    return {"id": i, "page": 0, "text": text, "loc": {"bbox": [x0, y0, x1, y0 + 10]}}


def test_graphic_page_poster_and_fragmented_art():
    page = fitz.open().new_page(width=842, height=595)
    big = [(fitz.Rect(20, 20, 800, 560), {})]
    labels = [_seg(i, 100 + 40 * i, 200, 130 + 40 * i) for i in range(10)]
    assert DL._graphic_page(page, labels, big)                                  # плакат: схема с подписями
    pieces = [(fitz.Rect(20 + 90 * k, 20, 100 + 90 * k, 100), {}) for k in range(6)]
    assert DL._graphic_page(page, labels[:4], pieces)                           # картинка обрывками, текста мало
    prose = [_seg(i, 60, 100 + 14 * i, 780, "szöveg " * 20) for i in range(30)]
    shot = [(fitz.Rect(60, 300, 500, 560), {})]
    assert not DL._graphic_page(page, prose, shot)                              # абзацы и один снимок экрана


def test_chart_table_needs_shapes_inside_not_cell_fills():
    xs, ys = [40, 200, 300, 400, 500], [100, 120, 140, 160, 180]
    t = (xs, ys, [], [])
    bars = [{"rect": (210 + 20 * k, 103 + 20 * k, 290 + 20 * k, 112 + 20 * k), "solid": True} for k in range(3)]
    fills = [{"rect": (40, 100, 200, 120), "solid": True}, {"rect": (200, 120, 300, 140), "solid": True},
             {"rect": (300, 140, 400, 160), "solid": True}]
    assert DL._chart_table(t, bars) and not DL._chart_table(t, fills)


def test_landscape_pages_get_landscape_sections(tmp_path):
    blocks = [p("Első", 1, 100), p("Második", 2, 100), p("Harmadik", 3, 100)]
    mem = B.Memory([("Első", "Первый"), ("Második", "Второй"), ("Harmadik", "Третий")])
    L = dict(LAYOUT, first_is_body=True, page_sizes=[(595.3, 841.9), (841.9, 595.3), (595.3, 841.9)])
    out = tmp_path / "o.docx"
    B.Builder(mem, lambda img: png(), PROFILE).build(blocks, str(out), L)
    secs = Document(str(out)).sections
    assert [s.page_width > s.page_height for s in secs] == [False, True, False]


def test_word_export_recognizes_pdfmaker():
    """PDF из Word через Acrobat PDFMaker — экспорт Word (разметка по его соглашениям), как «Microsoft Word»."""
    assert B.word_export({"creator": "Acrobat PDFMaker 5.0 for Word", "producer": "Acrobat Distiller 5.0"})
    assert B.word_export({"producer": "Microsoft® Office Word 2007"})
    assert not B.word_export({"creator": "PDF24 Creator", "producer": "GPL Ghostscript"})


def test_builder_survives_control_chars_wide_logos_and_bad_margins(tmp_path):
    """Сборка не падает: управляющие знаки в тексте (NUL, FF), логотипы шапки шире полосы набора, подвал ниже листа
    (отрицательное поле) — №6, №27, №40, №48 партии 2 (22.10)."""
    blocks = [p("Első" + chr(12) + " sor" + chr(1), 2, 100)]
    mem = B.Memory([("Első sor", "Первая строка")])
    logo = {"xref": 1, "x0": 0, "x1": 595, "top": 10, "bottom": 60}
    L = dict(LAYOUT, first_is_body=True, head_imgs=[logo, dict(logo, x0=300, x1=595)], bottom=-200, footer=-5)
    out = tmp_path / "x.docx"
    B.Builder(mem, lambda img: png(), PROFILE).build(blocks, str(out), L)
    sec = Document(str(out)).sections[-1]
    assert sec.bottom_margin >= 0 and chr(12) not in "".join(par.text for par in Document(str(out)).paragraphs)
