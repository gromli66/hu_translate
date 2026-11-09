# -*- coding: utf-8 -*-
"""Решатель «колонтитул листа» (zones.decide, E9b-Z) на сегментах без PDF: линии листов и рамки шапки — заданы."""
import zones as ZN

H = 842


def _seg(segs, page, text, y, x0=60, x1=300, **loc):
    segs.append({"id": len(segs), "page": page, "text": text,
                 "loc": dict({"bbox": [x0, y, x1, y + 10], "size": 10, "bold": False, "align": "left"}, **loc)})


def _rules(pages, lines=None):
    """(—, линии листов [(горизонтали, вертикали)], —, —, низ рамки шапки) — как docx_layout._page_rules."""
    return None, [(lines or {}).get(p, ([], [])) for p in range(pages)], None, None, [None] * pages


def test_series_header_fields_leave_body_from_first_sheet_of_series():
    """Серия листов со своей шапкой (3PR10 л. 20–34 «Azonosító: 3PR10/INFO», «verziószám: 0»): строки-поля шапки —
    не тело и на первом листе серии (оставались текстом в начале листа 20, у 1SZ01 «Oldalszám:1» — заголовком);
    бегущее название раздела (полужирное, по центру, не поле) — заголовком один раз, на первом листе."""
    segs = []
    for p in range(6):
        # шапка серии ниже шапки тела (3PR10: 45,8 pt против 34,4) — повтор по высоте её не берёт
        _seg(segs, p, *(("Azonosító: 3PR10", 40) if p < 4 else ("Azonosító: 3PR10/INFO", 52)))
        _seg(segs, p, f"Oldalszám: {p + 1}", 40, x0=400, x1=520)
        if p in (2, 3):
            _seg(segs, p, "1.15. Átfolyás a primerkörből", 80, x0=200, x1=400, bold=True, align="center")
        _seg(segs, p, f"Szöveg a lapon, sor {p}.", 300)
    content, why = ZN.decide(segs, H, 6, _rules(6))
    kept = {s["id"]: s for s in content}
    info = [s for s in segs if s["text"] == "Azonosító: 3PR10/INFO"]
    assert [why.get(s["id"]) for s in info] == ["шапка серии", "повтор с прошлого листа"]
    title = [s for s in segs if s["text"].startswith("1.15.")]
    assert title[0]["id"] in kept and kept[title[0]["id"]]["loc"].get("title") and title[1]["id"] not in kept
    assert all(s["id"] in kept for s in segs if s["text"].startswith("Szöveg"))


def test_repeat_from_previous_sheet_is_not_a_table_cell():
    """Строка вверху листа, чья подпись на прошлом листе стоит в ячейке таблицы с телом (таблица изменений), — тело:
    «Érintett fejezetek:» листа регистрации 30RL2126GT пропадала как «повтор с прошлого листа» (E9b-Z, Z4)."""
    segs = []
    _seg(segs, 0, "Verzió", 104, x1=190)
    _seg(segs, 0, "Érintett fejezetek:", 124, x0=210, x1=400)
    _seg(segs, 0, "Új TU kiadása, alkalmasság ellenőrzésre", 144, x0=210, x1=480)
    _seg(segs, 1, "Érintett fejezetek:", 134, x0=65, x1=180)
    _seg(segs, 1, "Ciklikusan felülvizsgált TU kiadása.", 300)
    xs, ys = [50, 200, 500], [100, 120, 140, 160]
    table = ([(y, 50, 500) for y in ys], [(x, 100, 160) for x in xs])
    content, why = ZN.decide(segs, H, 2, _rules(2, {0: table}))
    assert {s["id"] for s in content} == {s["id"] for s in segs}
    assert 1 in why.heads and 3 not in why.heads


def test_ocr_sheets_numbered_from_header_counter(tmp_path):
    """Путь А и для распознанного документа (E9b-Z, Z8): номер листа оригинала — счётчик шапки («Oldalszám: N»; первый
    лист PDF — «2», титула в файле нет), лист — свой раздел со своим номером, номер в шапке — поле PAGE на месте
    счётчика (метка), а не первое число перевода; лист, где OCR номера не прочёл, — между соседями."""
    from docx import Document
    from docx.oxml.ns import qn
    import docx_build as B
    import docx_struct as D
    from test_struct_docx import LAYOUT, PROFILE, ocr_doc, png
    src, segs = ocr_doc(tmp_path)
    for s in segs:
        if s["page"] == 4 and s["text"].startswith("Oldalszám"):
            s["text"] = "Oldalszám:"
    blocks, L = B.blocks_layout(src, segs, "layout")
    assert L["sheet_no"] == {p: p + 1 for p in range(1, 9)} and L["page_marked"]
    assert any(D.PAGE_MARK in c for r in L["head_grid"]["rows"] for c in r["cells"])
    out = tmp_path / "ocr.docx"
    B.Builder(B.Memory([(s["text"], s["text"]) for s in segs]), lambda img: png(), PROFILE).build(
        blocks, str(out), dict(LAYOUT, **L))
    doc = Document(str(out))
    starts = [(s._sectPr.find(qn("w:pgNumType")).get(qn("w:start")) if s._sectPr.find(qn("w:pgNumType")) is not None
               else None) for s in doc.sections]
    assert starts == [str(n) for n in range(2, 10)]
    # номер — полем на первой странице раздела-листа; его перелив — «N … (прод.)» текстом (путь А)
    assert all("PAGE" in sec.first_page_header._element.xml for sec in doc.sections)
