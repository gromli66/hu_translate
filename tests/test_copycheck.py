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
