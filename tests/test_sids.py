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

@pytest.fixture(autouse=True)
def _flag_on(monkeypatch):
    """Тесты этапа — с включённым флагом, что бы ни стояло в среде; выключенный проверяют свои тесты."""
    monkeypatch.setenv("HU_U_SIDS", "1")


seg = lambda i, text, page=0, y=300.0, **loc: {"id": i, "page": page, "text": text,
                                               "loc": dict({"bbox": [60.0, y, 200.0, y + 10]}, **loc)}


def attach(blocks, segs):
    """Память задачи (все сегменты переведены) раздаёт абзацам дерева разметки их id."""
    B.job_memory(segs, {str(s["id"]): {"ru": "ру"} for s in segs}).attach(blocks)


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
        [para("Állapot ellenőrzése", None), para("Állapot ellenőrzése", None)], str(tmp_path / "b.docx"),
        dict(LAYOUT, exempt_ids={}))
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
    rows = A.lost_id_rows(lost, {A._key24("Первое слово абзаца")})
    assert len(rows) == 1 and rows[0][1] == "потеря сегмента" and "№2" in rows[0][2]


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
    attach(blocks, segs)
    assert [b.get("sids") for b in blocks[:2]] == [[1], [2]] and blocks[2]["items"][0][0]["sids"] == [3]
    assert "sids" not in blocks[3] and "sids" not in blocks[4]


def test_same_source_text_gets_one_translation_per_document():
    """termvar.same_source: тот же исходный текст в разных сегментах — один перевод: утверждённый редактором, иначе самый
    частый, при равенстве — первый. Другой регистр или знак — другой текст; метки, копии и надписи на картинках не трогаются."""
    import termvar
    segs = [seg(1, "Állapot ellenőrzése"), seg(2, "Állapot ellenőrzése"), seg(3, "Állapot ellenőrzése"),
            seg(4, "Állapot ellenőrzése:"), seg(5, "Szint mérése"), seg(6, "Szint mérése"),
            seg(7, "Nyomás növelése"), seg(8, "Nyomás növelése"), seg(9, "Nyomás növelése", shot={"xref": 1}),
            seg(10, "<f1>Zárni</f1>"), seg(11, "<f1>Zárni</f1>")]
    res = {"1": {"ru": "Проверка состояния"}, "2": {"ru": "Контроль состояния"}, "3": {"ru": "Контроль состояния"},
           "4": {"ru": "Проверка состояния:"}, "5": {"ru": "Измерение уровня"}, "6": {"ru": "Замер уровня", "edited": True},
           "7": {"ru": "Повышение давления"}, "8": {"ru": "Рост давления"}, "9": {"ru": "Давление"},
           "10": {"ru": "<f1>Закрыть</f1>"}, "11": {"ru": "<f1>Закрывать</f1>"}}
    out, was = termvar.same_source(segs, res)
    ru = lambda i: out[str(i)]["ru"]
    assert [ru(1), ru(2), ru(3)] == ["Контроль состояния"] * 3 and ru(4) == "Проверка состояния:"
    assert [ru(5), ru(6)] == ["Замер уровня"] * 2                       # утверждённый редактором
    assert [ru(7), ru(8)] == ["Повышение давления"] * 2 and ru(9) == "Давление"   # равенство — первый, как было
    assert ru(10) == "<f1>Закрыть</f1>" and ru(11) == "<f1>Закрывать</f1>"
    assert was == {1: "Проверка состояния", 5: "Измерение уровня", 8: "Рост давления"}
    assert res["1"]["ru"] == "Проверка состояния" and out["1"]["unified"] == "Проверка состояния"   # res не изменён


def test_footnote_gets_its_own_ids_not_its_hosts():
    segs = [seg(1, "Fő szöveg"), seg(2, "Lábjegyzet szövege")]
    blocks = [{"k": "p", "runs": [{"t": "Fő szöveg"}], "notes": [{"runs": [{"t": "Lábjegyzet szövege"}], "pos": 1.0}]}]
    attach(blocks, segs)
    assert blocks[0]["sids"] == [1] and blocks[0]["notes"][0]["sids"] == [2]


def test_lost_ids_in_docx_by_tree_checks_only_tree_paragraphs():
    """DOCX по дереву разметки: сегмент страницы — не абзац дерева (строки склеены иначе), его «потерю» по id не судят;
    абзац разметки (struct_only) не взят — потеря."""
    segs = [seg(1, "Oldal sor"), seg(2, "Fa bekezdése", struct_only=True), seg(3, "Másik bekezdés", struct_only=True)]
    res = {"1": {"ru": "ру"}, "2": {"ru": "ру"}, "3": {"ru": "ру"}}
    assert [s["id"] for s, _ in A.lost_ids(segs, res, used={2}, tree=True)] == [3]
    assert [s["id"] for s, _ in A.lost_ids(segs, res, used={2}, tree=False)] == [1]    # (struct_only — не здесь)


def test_attach_claims_the_tree_segment_with_the_page_segment_of_the_same_text():
    """Абзац дерева, текст которого есть и сегментом страницы, и сегментом разметки (с метками), получает оба id:
    перевод ищется по тексту, и оба взяты."""
    segs = [seg(1, "Zárni a szelepet"), seg(2, "<f1>Zárni</f1><f2> a szelepet</f2>", struct_only=True)]
    blocks = [{"k": "p", "runs": [{"t": "Zárni", "b": True}, {"t": " a szelepet"}]}]
    attach(blocks, segs)
    assert blocks[0]["sids"] == [1, 2]


def test_attach_gives_every_tree_segment_of_the_text_to_the_paragraph():
    """Второй круг struct_fill: тот же абзац есть сегментом с метками и без них — оба сегмента разметки у абзаца
    (перевод тела берётся по тексту из любого), и каждому повтору абзаца."""
    segs = [seg(1, "Zárni a szelepet"), seg(2, "<f1>Zárni</f1><f2> a szelepet</f2>", struct_only=True),
            seg(3, "Zárni a szelepet", struct_only=True), seg(4, "Zárni a szelepet")]
    blocks = [{"k": "p", "runs": [{"t": "Zárni", "b": True}, {"t": " a szelepet"}]},
              {"k": "p", "runs": [{"t": "Zárni", "b": True}, {"t": " a szelepet"}]}]
    attach(blocks, segs)
    assert blocks[0]["sids"] == [1, 2, 3] and blocks[1]["sids"] == [4, 2, 3]


def test_part_found_in_a_tabbed_text_does_not_mark_the_block_used():
    """Текст с табуляцией ищется по частям; целого нет — ничего не взято (потерю видно)."""
    mem = B.Memory([("Kódszám neve", "Название кода", 1)])
    mem.at([5])
    assert mem("Kódszám neve	Ismeretlen érték szöveg") is None and mem.used == set()


def test_memory_context_is_restored_when_a_block_fails(tmp_path):
    """Исключение внутри блока не оставляет память в его контексте: следующий поиск (колонтитул) идёт вне блока."""
    mem = B.Memory([("Egy", "Один", 1)])
    bld = B.Builder(mem, lambda img: png(), PROFILE)
    bld._put_runs = lambda *a, **k: (_ for _ in ()).throw(RuntimeError("картинка не вставилась"))
    with pytest.raises(RuntimeError):
        bld.paragraph({"k": "p", "runs": [{"t": "Egy", "size": 10}], "sids": [1]})
    assert mem.cur == () and mem.used == set()


def test_flag_off_builder_ignores_sids_in_blocks(tmp_path, monkeypatch):
    """HU_U_SIDS=0: блоки с sids собираются так же, как без них (document.xml тот же), id не берутся."""
    import zipfile
    para = lambda t, sids: dict({"k": "p", "runs": [{"t": t, "size": 10}], "seg": True,
                                 "geom": {"page": 1, "top": 100, "x0": 60, "x1": 200, "bottom": 110, "first_x0": 60}},
                                **({"sids": sids} if sids else {}))
    pairs = [("Állapot ellenőrzése", "Проверка состояния", 1), ("Állapot ellenőrzése", "Контроль состояния", 2)]
    monkeypatch.setenv("HU_U_SIDS", "0")
    xml = []
    for name, ids in (("a", ([1], [2])), ("b", (None, None))):
        mem = B.Memory(pairs)
        used = B.Builder(mem, lambda img: png(), PROFILE).build(
            [para("Állapot ellenőrzése", ids[0]), para("Állapot ellenőrzése", ids[1])], str(tmp_path / f"{name}.docx"),
            LAYOUT)
        assert used == set()
        xml.append(zipfile.ZipFile(tmp_path / f"{name}.docx").read("word/document.xml"))
    assert xml[0] == xml[1]


def test_same_source_leaves_polished_and_shortened_translations_alone():
    import termvar
    segs = [seg(1, "Állapot ellenőrzése"), seg(2, "Állapot ellenőrzése"), seg(3, "Állapot ellenőrzése"),
            seg(4, "Típus:	X"), seg(5, "Típus: X")]
    res = {"1": {"ru": "Проверка состояния", "polish": {"of": "x"}}, "2": {"ru": "Контроль состояния"},
           "3": {"ru": "Контроль состояния", "ru_pdf": "Контроль", "ru_pdf_of": "Контроль состояния"},
           "4": {"ru": "Тип:	X"}, "5": {"ru": "Тип: X"}}
    out, was = termvar.same_source(segs, res)
    assert out["1"]["ru"] == "Проверка состояния" and out["3"]["ru"] == "Контроль состояния"   # защищены
    assert out["2"]["ru"] == "Проверка состояния" and was == {2: "Контроль состояния"}       # непривилегированный — к выбору
    assert out["4"]["ru"] == "Тип:	X" and out["5"]["ru"] == "Тип: X"                         # табуляция — не пробел


def test_same_source_keeps_first_use_abbreviation():
    """«СУЗ (SZBV)» при первом упоминании и «СУЗ» дальше — соглашение о первом упоминании, не разнобой: не унифицируется."""
    import termvar
    segs = [seg(1, "SZBV vizsgálata"), seg(2, "SZBV vizsgálata"), seg(3, "SZBV vizsgálata")]
    res = {"1": {"ru": "Испытание СУЗ (SZBV)"}, "2": {"ru": "Испытание СУЗ"}, "3": {"ru": "Испытание СУЗ"}}
    out, was = termvar.same_source(segs, res)
    assert was == {} and out["1"]["ru"] == "Испытание СУЗ (SZBV)"


def test_memory_marks_tree_segment_asked_outside_a_paragraph():
    """Подпись рисунка спрашивается у памяти вне абзаца (caption → self.tr(title)): сегмент разметки с этим текстом взят."""
    segs = [seg(1, "Rajz felirata"), seg(2, "Kapcsolási rajz", struct_only=True)]
    mem = B.job_memory(segs, {"1": {"ru": "Подпись рисунка"}, "2": {"ru": "Схема соединений"}})
    assert mem.get("Kapcsolási rajz") == "Схема соединений" and mem.used == {2}
    mem.at([1])
    mem.get("Rajz felirata")
    assert mem.used == {1, 2}


def test_missed_rows_name_texts_without_translation_once():
    rows = A.missed_rows(["Ismeretlen", "Ismeretlen", "  ", "Másik szöveg"])
    assert [r[1:] for r in rows] == [("нет перевода", "в DOCX остался текст оригинала", "Ismeretlen"),
                                     ("нет перевода", "в DOCX остался текст оригинала", "Másik szöveg")]


def test_same_source_prefers_clean_translation_and_keeps_abbreviation_by_position():
    """Тираж не размножает перевод с замечанием; первое упоминание аббревиатуры остаётся первым даже при разнобое слов."""
    import termvar
    segs = [seg(1, "Szelep nyitva"), seg(2, "Szelep nyitva"), seg(3, "SZBV leállt"), seg(4, "SZBV leállt"),
            seg(5, "SZBV leállt"), seg(6, "SZBV leállt")]
    res = {"1": {"ru": "Клапан открыт", "issues": ["число"]}, "2": {"ru": "Клапан открыт 5"},
           "3": {"ru": "СУЗ (SZBV) остановилась"}, "4": {"ru": "СУЗ прекратила работу"}, "5": {"ru": "СУЗ остановилась"},
           "6": {"ru": "СУЗ остановилась"}}
    out, was = termvar.same_source(segs, res)
    assert out["1"]["ru"] == "Клапан открыт 5" and out["2"]["ru"] == "Клапан открыт 5"
    assert [out[str(i)]["ru"] for i in (3, 4, 5, 6)] == ["СУЗ (SZBV) остановилась", "СУЗ остановилась",
                                                          "СУЗ остановилась", "СУЗ остановилась"]
