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
    """Номер изменения «2.» перед «N. oldal» на листах 4–6 из 12 (ÜFK I: л. 14–16, 34, 64, 65 из 157): номер листа —
    число, что продолжает ряд соседей, а не первое по подписи; метка поля PAGE — на нём, «2.» — текстом в подвале серии.
    Ряд, которого нет по первым числам (изменение на 3 листах из 7), подбором не создаётся — номеров нет, как прежде."""
    src = tmp_path / "amend.pdf"
    blocks = _tagged_like(src, 12, amend={4, 5, 6})
    L = D.layout(str(src), blocks)
    assert L.get("sheet_no") == {pn: pn for pn in range(1, 13)}
    v = [z for z in L["variants"] if z["from"] <= 4 <= z["to"]]
    assert v, "листы с номером изменения — своя серия"
    stamp = " ".join(c for r in v[0]["stamp"] for c in r["cells"])
    assert stamp.replace(" ", "") == "Modositassorszama:2." + D.PAGE_MARK + ".oldal"
    dense = tmp_path / "dense.pdf"
    assert not D.layout(str(dense), _tagged_like(dense, 7, amend={3, 4, 5})).get("sheet_no")


def test_foot_top_border_is_nearest_rule_not_body_table_bottom():
    """Верх подвала — по ближайшей линии над его строкой (и вплотную к ней), а не по любой линии ниже тела: низ таблицы
    тела в 12 pt над подвалом давал штампу лишнюю пустую строку (JPDS, 57 листов; E9b-S Z10)."""
    import docx_struct as D
    foot = {"text": "Д11", "top": 780.0, "bottom": 790.0, "x0": 70.0, "x1": 100.0}
    rule = lambda y: {"top": y, "bottom": y + 0.5, "x0": 70.0, "x1": 560.0}
    assert D._foot_top([foot], [rule(763.0), rule(775.0)], 842.0, 752.0) == 774
    assert D._foot_top([foot], [rule(763.0)], 842.0, 752.0) == 779          # линии рядом нет — по строке
    assert D._foot_top([foot], [rule(774.0), rule(775.5)], 842.0, 752.0) == 773   # двойная линия — целиком
