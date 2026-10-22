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


def test_upright_copy_of_rotated_pages(tmp_path):
    """Лист с /Rotate 90 без разметки — копия, где слова идут строкой слева направо в координатах видимого листа (HAKSER:
    таблицы альбомных листов собирались узкими столбцами, E48)."""
    import segments as S
    src = tmp_path / "rot.pdf"
    d = fitz.open()
    pg = d.new_page(width=612, height=792)
    pg.insert_text((100, 100), "Mérés eredménye", fontsize=11, rotate=270)
    pg.set_rotation(90)
    d.save(str(src))
    up = S.upright(src, tmp_path / "up")
    assert up != src
    w = fitz.open(str(up))[0].get_text("words")
    x0, y0, x1, y1 = w[0][:4]
    assert x1 - x0 > y1 - y0                                    # слово лежит строкой, а не столбиком
    plain = tmp_path / "plain.pdf"
    fitz.open().new_page().insert_text((50, 50), "a")
    d2 = fitz.open(); d2.new_page().insert_text((50, 50), "szöveg"); d2.save(str(plain))
    assert S.upright(plain, tmp_path / "up") == plain           # без поворотов — исходный файл


def test_ocr_layer_only_invisible_text_over_scan():
    """Скан с невидимым слоем распознавания — не плакат; живой текст поверх фона-картинки (брошюра) — может быть."""
    import io
    from PIL import Image
    buf = io.BytesIO()
    Image.new("RGB", (60, 80), "lightblue").save(buf, "PNG")
    doc = fitz.open()
    for mode in (3, 0):
        pg = doc.new_page()
        pg.insert_image(pg.rect, stream=buf.getvalue())
        pg.insert_text((72, 100), "Szöveg a háttér felett", render_mode=mode)
    assert DL._ocr_layer(doc[0]) and not DL._ocr_layer(doc[1])


def test_nested_table_in_cell_keeps_its_text(tmp_path):
    """Таблица в ячейке таблицы (листовка Word) — вложенной таблицей, а не пропуском вместе с текстом."""
    inner = {"k": "table", "rows": [[{"blocks": [p("Belső szöveg")], "th": False}],
                                    [{"blocks": [p("Második sor")], "th": False}]]}
    outer = {"k": "table", "rows": [[{"blocks": [p("Külső"), inner, p("Utána")], "th": False},
                                     {"blocks": [p("Jobb")], "th": False}]]}
    mem = B.Memory([("Belső szöveg", "Внутренний текст"), ("Második sor", "Вторая строка"), ("Külső", "Внешний"),
                    ("Utána", "После"), ("Jobb", "Справа")])
    out = tmp_path / "n.docx"
    B.Builder(mem, lambda img: png(), PROFILE).build([outer], str(out), LAYOUT)
    doc = Document(str(out))
    cell = doc.tables[0].cell(0, 0)
    assert [c.text for r in cell.tables[0].rows for c in r.cells] == ["Внутренний текст", "Вторая строка"]
    texts = [q.text for q in cell.paragraphs if q.text]
    assert texts == ["Внешний", "После"] and doc.tables[0].cell(0, 1).text == "Справа"


def test_paragraph_styled_like_chapter_heading_starts_new_page():
    """E41: абзац с видом заголовков 1-го уровня (те — с нового листа), первым на листе на их высоте, — с нового листа;
    обычный абзац первым на листе и такой же жирный ниже по листу — нет."""
    import docx_struct as D

    def blk(k, text, page, top, size, bold, **kw):
        return dict({"k": k, "runs": [{"t": text, "size": size, "b": bold}],
                     "geom": {"page": page, "top": top, "bottom": top + 14, "x0": 70}}, **kw)
    blocks = [blk("h", "1. BEVEZETÉS", 2, 104, 14, True, level=1), blk("p", "Szöveg.", 2, 140, 12, False),
              blk("h", "2. FOGALMAK", 3, 104, 14, True, level=1), blk("p", "Szöveg.", 3, 140, 12, False),
              blk("p", "ELŐSZÓ", 4, 104, 14, True), blk("p", "Folytatás a lap tetején", 5, 104, 12, False),
              blk("p", "Kiemelt sor", 5, 300, 14, True)]
    tops = {2: 104, 3: 104, 4: 104, 5: 104}
    bots = {1: 760, 2: 760, 3: 760, 4: 760, 5: 760}
    D.mark_breaks(blocks, tops, bots)
    for b in blocks[:4]:
        b["break"] = b["k"] == "h"                    # заголовки 1-го уровня — с нового листа (как стилем у Word)
    D.mark_breaks(blocks, tops, bots)
    assert [bool(b.get("break")) for b in blocks[4:]] == [True, False, False]


def test_toc_first_on_page_after_short_page_starts_new_page(tmp_path):
    """E41: оглавление первым на листе после листа, занятого на четверть, — с нового листа (и в DOCX)."""
    import docx_struct as D
    body = p("Előszó szövege.", 5, 100)
    body["geom"]["bottom"] = 250
    toc = {"k": "toc", "info": {"page": 6, "levels": 2}, "geom": {"page": 6}}
    blocks = [body, toc]
    D.mark_breaks(blocks, {5: 100}, {1: 760, 2: 760, 3: 760, 4: 760, 5: 250})
    assert toc.get("break")
    out = tmp_path / "t.docx"
    B.Builder(B.Memory([("Előszó szövege.", "Текст предисловия.")]), lambda img: png(), PROFILE).build(
        blocks, str(out), LAYOUT)
    doc = Document(str(out))
    head = next(q for q in doc.paragraphs if q.style.name.startswith("TOC Title"))
    assert head.paragraph_format.page_break_before


def test_header_logo_is_not_a_screenshot():
    """E42: небольшая картинка у верха листа (логотип шапки) — не скриншот для перевода надписей; та же в теле — да."""
    import io
    from PIL import Image
    import shots as SH
    buf = io.BytesIO()
    Image.new("RGB", (240, 150), "navy").save(buf, "PNG")
    doc = fitz.open()
    pg = doc.new_page(width=842, height=595)
    pg.insert_image(fitz.Rect(61, 54, 169, 120), stream=buf.getvalue())
    pg2 = doc.new_page(width=842, height=595)
    pg2.insert_image(fitz.Rect(61, 300, 169, 366), stream=buf.getvalue())
    got = SH.find_shots(doc)
    assert [n for places in got.values() for n, _ in places] == [1]
