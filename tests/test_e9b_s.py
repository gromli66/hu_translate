"""E9b-S: путь А колонтитулов и у размеченного PDF (S) — номер листа оригинала по счётчику подписи, серии, метка."""
import fitz

import docx_struct as D


def _tagged_like(path, pages, amend=()):
    """PDF без разметки знаков (все знаки — вне дерева, как артефакты размеченного): шапка вверху, абзац тела, подвал
    «Módosítás sorszáma: [k.] N. oldal» внизу (ÜFK I: номер изменения заполнен на части листов). → блоки тела."""
    d = fitz.open()
    blocks = []
    for pno in range(pages):
        pg = d.new_page(width=595.3, height=841.9)
        pg.insert_text((70, 40), "Kezelesi utasitas", fontsize=9)
        pg.insert_text((70, 160), f"Torzs szoveg {pno}", fontsize=11)
        k = "2. " if pno + 1 in amend else ""
        pg.insert_text((70, 800), f"Modositas sorszama: {k}{pno + 1}. oldal", fontsize=9)
        blocks.append({"k": "p", "runs": [{"t": f"Torzs szoveg {pno}"}],
                       "geom": {"page": pno + 1, "top": 150, "bottom": 162, "x0": 70, "x1": 200}})
    d.save(str(path))
    return blocks


def test_tagged_pdf_gets_sheet_numbers_by_counter(tmp_path):
    """Размеченный PDF (без content) — путь А, как у текстового слоя: номер листа оригинала из счётчика подвала.
    Прежде у S путь А не строился (sheet_no не было): нумерация шла подряд по страницам перевода."""
    src = tmp_path / "s.pdf"
    blocks = _tagged_like(src, 7)
    L = D.layout(str(src), blocks)
    assert L.get("sheet_no") == {pn: pn for pn in range(1, 8)}


def test_counter_with_filled_field_before_it_takes_sheet_number(tmp_path):
    """Номер изменения «2.» перед «N. oldal» на листах 3–5 (ÜFK I л. 14–16, 34, 64, 65): номер листа — число, что
    продолжает ряд соседей, а не первое по подписи; метка поля PAGE — на нём, «2.» — текстом в подвале серии."""
    src = tmp_path / "amend.pdf"
    blocks = _tagged_like(src, 7, amend={3, 4, 5})
    L = D.layout(str(src), blocks)
    assert L.get("sheet_no") == {pn: pn for pn in range(1, 8)}
    v = [z for z in L["variants"] if z["from"] <= 3 <= z["to"]]
    assert v, "листы с номером изменения — своя серия"
    stamp = " ".join(c for r in v[0]["stamp"] for c in r["cells"])
    assert stamp.replace(" ", "") == "Modositassorszama:2." + D.PAGE_MARK + ".oldal"
