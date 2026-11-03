# E6 «Блок помнит свои сегменты» (sids.py): перевод блока из одного сегмента — по его id, составного — по тексту; сборка
# возвращает id, что взяла (used_ids); потеря = id с переводом, не взятый и не снятый с причиной (lost_ids).
import pytest

import assemble as A
import docx_build as B
import docx_layout as DL
import docx_struct as D
import sids as S

from test_assembly_sheets import _sheets_pdf
from test_struct_docx import LAYOUT, PROFILE, png
from test_text_forms import _segs

seg = lambda i, text, page=0, y=300.0, **loc: {"id": i, "page": page, "text": text,
                                               "loc": dict({"bbox": [60.0, y, 200.0, y + 10]}, **loc)}


def walk(bs):
    for b in bs:
        yield b
        if b["k"] == "list":
            for it in b["items"]:
                yield from walk(it)
        elif b["k"] == "table":
            for r in b["rows"]:
                for c in r:
                    yield from walk(c.get("blocks") or [])


def test_block_gets_its_own_translation_by_id_not_the_first_of_the_text():
    """Один венгерский текст в двух сегментах переведён по-разному (правка редактора, вариант портала): блок берёт
    перевод своего сегмента (res[id].ru), а не первый из сегментов с таким текстом."""
    mem = B.Memory([("Állapot ellenőrzése", "Проверка состояния", 1), ("Állapot ellenőrzése", "Контроль состояния", 2)])
    assert mem.get("Állapot ellenőrzése") == "Проверка состояния"           # вне блока — как прежде: первый побеждает
    prev = mem.at([2])
    assert mem.get("Állapot ellenőrzése") == "Контроль состояния" and mem.used == {2}
    mem.at(prev)
    assert mem.cur == ()


def test_composite_block_and_other_text_go_by_text_and_mark_all_ids_used():
    """Блок из двух сегментов (склеенные продолжения) и текст, не равный тексту сегмента (без метки пункта), — по тексту;
    перевод найден — все id блока взяты. Не найден — не взяты."""
    mem = B.Memory([("első sor", "первая строка", 1), ("második sor", "вторая строка", 2)])
    mem.at([1, 2])
    assert mem.get("első sor") == "первая строка" and mem.used == {1, 2}
    mem2 = B.Memory([("első sor", "первая строка", 1)])
    mem2.at([1, 5])
    assert mem2.get("ismeretlen szöveg") is None and mem2.used == set()


def test_tagged_segment_is_found_only_by_its_tagged_text():
    mem = B.Memory([("<f1>FIGYELEM!</f1><f2> áram alatt</f2>", "<f1>ВНИМАНИЕ!</f1><f2> под напряжением</f2>", 7)])
    mem.at([7])
    assert mem.get("<f1>FIGYELEM!</f1><f2> áram alatt</f2>") == "<f1>ВНИМАНИЕ!</f1><f2> под напряжением</f2>"
    assert mem.used == {7}


def test_value_without_words_takes_its_res_by_id():
    """«1,7 kW»: портал не нужен (needs_translation), но в DOCX — res[id].ru («1,7 кВт»), и id взят."""
    mem = B.Memory([("1,7 kW", "1,7 кВт", 3)])
    mem.at([3])
    assert mem.get("1,7 kW") == "1,7 кВт" and mem.used == {3}


def test_flag_off_memory_is_text_only(monkeypatch):
    """HU_U_SIDS=0: память не знает id — первый перевод текста побеждает везде, used пуст."""
    monkeypatch.setenv("HU_U_SIDS", "0")
    mem = B.Memory([("Állapot ellenőrzése", "Проверка состояния", 1), ("Állapot ellenőrzése", "Контроль состояния", 2)])
    prev = mem.at([2])
    mem.at(prev)
    assert mem.ids == {} and mem.get("Állapot ellenőrzése") == "Проверка состояния"
    assert S.of([seg(1, "x")]) == {} and S.add({"k": "p"}, [1]) == {"k": "p"}


def test_builder_writes_own_variant_per_block(tmp_path):
    """Сборка: два абзаца с одним текстом и разными переводами сегментов — каждый со своим переводом; id взяты;
    без sids в блоках (старое задание) — первый перевод в обоих."""
    from docx import Document
    para = lambda t, sids: dict({"k": "p", "runs": [{"t": t, "size": 10}], "seg": True,
                                 "geom": {"page": 1, "top": 100, "x0": 60, "x1": 200, "bottom": 110, "first_x0": 60}},
                                **({"sids": sids} if sids else {}))
    mem = B.Memory([("Állapot ellenőrzése", "Проверка состояния", 1), ("Állapot ellenőrzése", "Контроль состояния", 2)])
    used = B.Builder(mem, lambda img: png(), PROFILE).build(
        [para("Állapot ellenőrzése", [1]), para("Állapot ellenőrzése", [2])], str(tmp_path / "a.docx"), LAYOUT)
    assert used == {1, 2}
    assert [p.text for p in Document(str(tmp_path / "a.docx")).paragraphs if p.text] == [
        "Проверка состояния", "Контроль состояния"]
    mem = B.Memory([("Állapot ellenőrzése", "Проверка состояния", 1), ("Állapot ellenőrzése", "Контроль состояния", 2)])
    used = B.Builder(mem, lambda img: png(), PROFILE).build(
        [para("Állapot ellenőrzése", None), para("Állapot ellenőrzése", None)], str(tmp_path / "b.docx"), LAYOUT)
    assert used == set()
    assert [p.text for p in Document(str(tmp_path / "b.docx")).paragraphs if p.text] == [
        "Проверка состояния", "Проверка состояния"]


def test_every_segment_is_in_a_block_or_removed_with_a_reason(tmp_path):
    """docx_layout.load: каждый сегмент листа — в sids какого-то блока, в info[exempt] (колонтитул) или в sids_dropped
    склеенной таблицы (повтор шапки таблицы на следующем листе); других не бывает. Без флага ключей нет."""
    src = tmp_path / "a.pdf"
    segs = _sheets_pdf(src, pages=4)
    blocks, info = DL.load(str(src), segs, text=True)
    in_blocks = {i for b in walk(blocks) for i in b.get("sids") or ()}
    dropped = S.dropped(blocks)
    assert dropped and not dropped & in_blocks                          # повторы шапки таблицы — по три на листах 2–4
    assert {s["id"] for s in segs} == in_blocks | set(info["exempt"]) | dropped
    assert set(info["exempt"].values()) == {"колонтитул"}


def test_load_without_flag_leaves_blocks_as_before(tmp_path, monkeypatch):
    monkeypatch.setenv("HU_U_SIDS", "0")
    src = tmp_path / "a.pdf"
    segs = _sheets_pdf(src, pages=3)
    blocks, info = DL.load(str(src), segs, text=True)
    assert "exempt" not in info and not any("sids" in b or "sids_dropped" in b for b in walk(blocks))


def test_build_takes_every_body_segment_and_a_dropped_block_is_lost(tmp_path):
    """Сборка из блоков text-режима берёт все id тела; убранный из блоков абзац — потеря (lost_ids), колонтитул и повтор
    шапки таблицы — нет."""
    src = tmp_path / "a.pdf"
    segs = _sheets_pdf(src, pages=3)
    res = {str(s["id"]): {"ru": "ру " + s["text"]} for s in segs}
    blocks, info = DL.load(str(src), segs, text=True)
    L = D.layout(str(src), blocks, content=info.get("content_boxes"))
    lay = dict(LAYOUT, **{k: v for k, v in L.items() if k in ("head_grid", "variants", "page_marked", "page_tok",
                                                             "first_is_body", "sheet_no")})
    mem = B.job_memory(segs, res)
    used = B.Builder(mem, lambda img: png(), PROFILE).build(blocks, str(tmp_path / "a.docx"), lay)
    assert A.lost_ids(segs, res, used, info["exempt"], S.dropped(blocks)) == []
    # абзац пропал между разбором и сборкой (любая будущая заплатка, что выбрасывает блок): сборка этого не спросит
    import copy
    rest = copy.deepcopy(blocks)
    cell = next(c for b in walk(rest) if b["k"] == "table" for r in b["rows"] for c in r
                if any("A 1. lap" in B.text_of(x["runs"]) for x in c["blocks"] if x["k"] == "p" and x.get("runs")))
    victim = next(x for x in cell["blocks"] if x["k"] == "p" and x.get("runs") and "A 1. lap" in B.text_of(x["runs"]))
    cell["blocks"].remove(victim)
    used = B.Builder(B.job_memory(segs, res), lambda img: png(), PROFILE).build(rest, str(tmp_path / "b.docx"), lay)
    lost = A.lost_ids(segs, res, used, info["exempt"], S.dropped(rest))
    assert victim["sids"] and {s["id"] for s, _ in lost} == set(victim["sids"])


def test_lost_ids_rules():
    """Не считаются: без перевода, надпись на картинке, абзац разметки, лист-снимок, пометка о переносе; снятые с причиной
    (exempt, dropped); без словаря exempt — полосы колонтитула у края листа. Считается всё прочее, что не взято."""
    segs = [seg(1, "Egy szöveg"), seg(2, "Nincs fordítás"), seg(3, "Kép felirata", shot={"xref": 1}),
            seg(4, "Táblázat folytatása", struct_only=True), seg(5, "Lap képpel", page=3),
            seg(6, "Folytatás a következő oldalon", ), seg(7, "Fejléc", y=5.0), seg(8, "Ismétlődő fejléc"),
            seg(9, "Marad", y=400.0), seg(10, "Elveszett", y=420.0)]
    res = {str(i): {"ru": "ру"} for i in (1, 3, 4, 5, 7, 8, 9, 10)}
    # 1 взят; 3, 4 — не считаются; 5 — лист-снимок (page 3); 7 — колонтитул у края (нет словаря); 8 — dropped; 9 — exempt
    lost = A.lost_ids(segs, res, used={1}, exempt=None, dropped={8, 9}, skip_pages={3})
    assert [s["id"] for s, _ in lost] == [10]
    lost = A.lost_ids(segs, res, used={1}, exempt={7: "колонтитул", 9: "маркер пункта"}, dropped={8}, skip_pages={3})
    assert [s["id"] for s, _ in lost] == [10]
    lost = A.lost_ids(segs, res, used={1}, exempt={9: "маркер пункта"}, dropped={8}, skip_pages={3})
    assert [s["id"] for s, _ in lost] == [7, 10]                          # со словарём край листа сам не снимает


def test_lost_id_rows_do_not_repeat_what_the_docx_text_check_named():
    segs = [seg(1, "Egy"), seg(2, "Kettő")]
    lost = [(segs[0], "Первое слово абзаца"), (segs[1], "Второе слово абзаца")]
    rows, n = A.lost_id_rows(lost, {A._key24("Первое слово абзаца")})
    assert n == 1 and rows[0][1] == "потеря сегмента" and "№2" in rows[0][2]


def test_cut_row_tail_and_inline_merge_carry_ids():
    """Хвост абзаца, разрезанного листом (_cut_rows), и строчный фрагмент, склеенный merge_inline, отдают id абзацу,
    куда влились."""
    cell = lambda t, pg, sids: {"blocks": [{"k": "p", "runs": [{"t": t}], "sids": sids}], "geom": {"page": pg, "x0": 60.0},
                                "th": False}
    rows = D._cut_rows([[cell("a mondat eleje", 1, [1])], [cell("a folytatása", 2, [2])]])
    assert len(rows) == 1 and rows[0][0]["blocks"][0]["sids"] == [1, 2]
    geo = lambda top, pg=1: {"page": pg, "top": top, "bottom": top + 10, "x0": 60.0, "first_x0": 60.0, "size": 10}
    out = B.merge_inline([{"k": "p", "runs": [{"t": "lásd:"}], "geom": geo(100), "sids": [1]},
                          {"k": "p", "runs": [{"t": "5.6 fejezet"}], "geom": geo(100), "inline": True, "sids": [2]}])
    assert len(out) == 1 and out[0]["sids"] == [1, 2]


def test_attach_sids_to_tag_tree_paragraphs_by_text():
    """Размеченный PDF: абзацы дерева находят свои сегменты по тексту, по порядку; чего нет — без sids (поиск по тексту)."""
    segs = [seg(1, "Egy sor"), seg(2, "Egy  sor"), seg(3, "Más sor"), seg(4, "Képfelirat", shot={"xref": 1})]
    blocks = [{"k": "p", "runs": [{"t": "Egy sor"}]}, {"k": "p", "runs": [{"t": "Egy sor"}]},
              {"k": "list", "kind": "Decimal", "items": [[{"k": "p", "runs": [{"t": "Más sor"}]}]]},
              {"k": "p", "runs": [{"t": "Egy sor"}]}, {"k": "p", "runs": [{"t": "Képfelirat"}]}]
    B.attach_sids(blocks, segs)
    assert [b.get("sids") for b in blocks[:2]] == [[1], [2]] and blocks[2]["items"][0][0]["sids"] == [3]
    assert "sids" not in blocks[3] and "sids" not in blocks[4]
