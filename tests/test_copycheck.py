# -*- coding: utf-8 -*-
"""Сверка копии скана тремя чтениями (src/copycheck.py): копия в меньшинстве — исправляется, согласия нет — выделение
и примечание; выделение доходит до перевода своей группой (segments.run_sig), 12.11."""
from docx import Document
from docx.oxml.ns import qn

import copycheck as CC
import segments as S


def make(tmp_path):
    doc = Document()
    p = doc.add_paragraph()
    p.add_run("Заявитель ").bold = True
    p.add_run("ООО «Авангард-ИСК», аттестат № РОСС RU.32248.04СЕ10.1.15, срок 01.02.2025.")
    f = tmp_path / "c.docx"; doc.save(f)
    return f


PLAIN = 'Заявитель ООО "Авангард-НСК", аттестат № РОСС RU.32248.04СЕЛО.1.15, срок 01.02.2025.'
SURYA = [{"t": "Заявитель ООО «Авангард-НСК», аттестат № РОСС RU.32248.04CEJI0.1.15, срок 01.02.2025.",
          "box": [10, 10, 900, 30], "c": 0.97}]


def test_minority_fixed_disagreement_marked(tmp_path):
    doc = Document(str(make(tmp_path)))
    fixed, marked = CC.check(doc, {0: PLAIN}, {0: SURYA})
    assert (fixed, marked) == (1, 1)
    p = doc.paragraphs[0]
    assert p.text == "Заявитель ООО «Авангард-НСК», аттестат № РОСС RU.32248.04СЕ10.1.15, срок 01.02.2025."
    hl = [r.text for r in p.runs if r.font.highlight_color is not None]
    assert hl == ["RU.32248.04СЕ10.1.15,"]
    notes = [c.text for c in doc.comments]
    assert any("было ««Авангард-ИСК»,»" in n for n in notes) and any("04CEJI0" in n and "04СЕЛО" in n for n in notes)
    assert p.runs[0].bold and p.runs[0].text == "Заявитель "                     # оформление прогонов не тронуто


def test_copy_in_majority_untouched(tmp_path):
    doc = Document(str(make(tmp_path)))
    plain = PLAIN.replace("Авангард-НСК", "Авангард-ИСК").replace("04СЕЛО", "04СЕ10")
    assert CC.check(doc, {0: plain}, {0: SURYA}) == (0, 0)


def test_highlight_survives_into_translation_groups(tmp_path):
    doc = Document(str(make(tmp_path)))
    CC.check(doc, {0: PLAIN}, {0: SURYA})
    out = tmp_path / "checked.docx"; doc.save(out)
    segs = S.docx_segments(out)
    assert len(segs[0]["loc"]["groups"]) == 4                                    # Заявитель · текст · выделенное · текст
    assert "<f3>RU.32248.04СЕ10.1.15,</f3>" in segs[0]["text"]                     # выделенное — своей группой


def test_placeholder_is_not_checked(tmp_path):
    doc = Document()
    doc.add_paragraph("Тел. +7 383 222-666-8, e-mail: info@avangardnsk.ru [логотип]")
    sur = [{"t": "Тел. +7 383 222-666-8, e-mail: info@avangardnsk.ru", "box": [0, 0, 500, 20], "c": 0.97}]
    plain = "Тел. +7 383 222-666-8, e-mail: info@avangardnsk.ru АВАНГАРД-НСК PLANET Networking"
    assert CC.check(doc, {0: plain}, {0: sur}) == (0, 0)


def test_stamp_lines_marked_by_ink(tmp_path):
    """Строка краски штампа (в цветной печати, найденной на скане) — "stamp"; чёрная строка внутри той же печати и
    вне её — нет."""
    import fitz
    import numpy as np
    img = np.full((1000, 600, 3), 255, np.uint8)
    blue = (40, 60, 200)
    img[470:474, 80:340] = blue                                       # рамка штампа
    img[586:590, 80:340] = blue
    img[470:590, 80:84] = blue
    img[470:590, 336:340] = blue
    img[500:520, 100:300] = blue                                      # синяя строка штампа
    img[540:560, 100:300] = 20                                        # чёрная строка внутри рамки штампа
    img[100:120, 100:300] = 20                                        # чёрная строка вне печати
    d = fitz.open(); p = d.new_page(width=432, height=720)          # 600×1000 px при 100 dpi
    p.insert_image(p.rect, pixmap=fitz.Pixmap(fitz.csRGB, 600, 1000, img.tobytes(), False))
    pdf = tmp_path / "s.pdf"; d.save(pdf)
    lines = [{"t": "КОПИЯ ВЕРНА", "box": [100, 500, 300, 520]}, {"t": "Текст", "box": [100, 540, 300, 560]},
             {"t": "Шапка", "box": [100, 100, 300, 120]}]
    assert CC.stamp_lines(pdf, {0: lines}, 100) == 1
    assert [bool(ln.get("stamp")) for ln in lines] == [True, False, False]
