# -*- coding: utf-8 -*-
"""Три дефекта конвейера (глазами, 06.11.2026): struct_fill срабатывал один раз; код-тег уходил на перевод;
термин «риск высокий» не проверялся. Портала нет — подставной переводчик (_call)."""
import json

import pytest

import pipeline as P
import translate as T
from test_struct_docx import PAGES, RULES, RU, page_segs, tagged_pdf


def _project(tmp_path):
    g = tmp_path / "g.md"
    g.write_text("## Т\n\n| Венгерский | Русский | Комментарий |\n|---|---|---|\n| szelep | клапан | |\n", encoding="utf-8")
    rules = tmp_path / "conventions.json"
    rules.write_text(json.dumps({k: RULES[k] for k in ("latin", "cyrillic_to_latin", "warning")}), encoding="utf-8")
    return {"glossary": str(g), "rules": str(rules)}


def _job(tmp_path, struct):
    """Размеченный PDF; перевод есть только у сегментов страниц — абзацев разметки без пары нет в res."""
    pdf = tmp_path / "doc.pdf"
    segs = page_segs(tagged_pdf(pdf, PAGES))
    work = tmp_path / "work"
    (work / "json").mkdir(parents=True)
    res = {str(s["id"]): {"ru": RU[s["text"]], "issues": []} for s in segs}
    (work / "json" / "doc.json").write_text(json.dumps({"file": str(pdf), "segs": segs, "res": res, "info": {},
                                                        "docx": None, "struct": struct}), encoding="utf-8")
    return work


def _fake_portal(monkeypatch, asked):
    def call(self, units, *a, **k):
        asked.extend(u["text"] for u in units)
        return {i: RU[u["text"]] for i, u in enumerate(units)}, {}
    monkeypatch.setattr(T.Translator, "_call", call)


def _doc(work):
    return json.loads((work / "json" / "doc.json").read_text(encoding="utf-8"))


# ---------------------------------------------------------------- 1. struct_fill — добор при любой пересборке
def test_struct_fill_tops_up_json_that_was_filled_by_older_build(tmp_path, monkeypatch):
    """JSON задания 29.10: struct_fill прошёл до правки сборки (стоит "filled"), сборка теперь спрашивает новые
    тексты — их надо добрать, а не оставить «нет перевода» (30RVRXTSBIGT_TU, пп. 1.16–1.17)."""
    work, asked = _job(tmp_path, {"coverage": 1.0, "filled": 0}), []
    _fake_portal(monkeypatch, asked)
    P.struct_fill(work, _project(tmp_path))
    d = _doc(work)
    own = [s for s in d["segs"] if s["loc"].get("struct_only")]
    assert asked == ["<f1>FIGYELEM!</f1><f2> A teszt leállítása nem lehetséges.</f2>",
                     "<f1>Adjuk ki az alábbi parancsot:</f1><f2> sync</f2>"]
    assert len(own) == 2 and all(d["res"][str(s["id"])]["ru"] for s in own)
    assert d["struct"]["filled"] == 2


def test_struct_fill_does_not_ask_again_what_is_already_translated(tmp_path, monkeypatch):
    work, asked = _job(tmp_path, {"coverage": 1.0}), []
    _fake_portal(monkeypatch, asked)
    P.struct_fill(work, _project(tmp_path))
    first = _doc(work)
    assert len(asked) == 2
    P.struct_fill(work, _project(tmp_path))                  # вторая пересборка: добирать нечего
    P.struct_fill(work, _project(tmp_path))
    assert len(asked) == 2
    again = _doc(work)
    assert again["segs"] == first["segs"] and again["res"] == first["res"]
    assert again["struct"]["filled"] == 2                    # счёт добранного не обнуляется пустым кругом


def test_struct_fill_survives_portal_failure_and_retries_next_time(tmp_path, monkeypatch):
    """Пересборка без портала (нет токена): добор не получился — сборка идёт дальше, JSON не тронут, в следующий
    раз добор повторяется."""
    work, asked = _job(tmp_path, {"coverage": 1.0, "filled": 0}), []

    def down(self, units, *a, **k):
        raise RuntimeError("нет токена портала")
    monkeypatch.setattr(T.Translator, "_call", down)
    before = _doc(work)
    P.struct_fill(work, _project(tmp_path))                  # не падает
    assert _doc(work) == before
    _fake_portal(monkeypatch, asked)
    P.struct_fill(work, _project(tmp_path))
    assert len(asked) == 2 and _doc(work)["struct"]["filled"] == 2


def test_assemble_with_project_tops_up_first(tmp_path, monkeypatch):
    """Пересборка (правки редактора, CLI assemble) — с проектом: перед сборкой добирается недостающее."""
    monkeypatch.setenv("HU_U_SIDS", "1")
    work, asked = _job(tmp_path, {"coverage": 1.0, "filled": 0}), []
    _fake_portal(monkeypatch, asked)
    P.assemble(work, prj=_project(tmp_path))
    assert len(asked) == 2
    assert sum(1 for s in _doc(work)["segs"] if s["loc"].get("struct_only")) == 2


# ---------------------------------------------------------------- 2. код-тег по виду не переводится
TAGS_AS_IS = ["@.BSZG.TBE", "@.BSZG.BC", "@.MEAS.I5", "@.FIL.1_1", "@BSZG.TBE", "<f1>@.BSZG.TBE</f1>", "BSZG.TBE.X"]
NOT_TAGS = ["Szervezet:UBO", "KI.BE", "AUT.IND", "BSZG.TBE", "BSZG", "1.2.3", "É.K.", "pl.", "mm/s", "g/kg", "4.1.9.",
            "A.B", "NYIT.", "Oszt.vez", "AZONOSÍTÓ:", "FIGYELEM!", "ZÁRVA.NYITVA.X", "30YP15S101 zárva", "@ BSZG.TBE",
            "@.bszg.tbe", "@.BSZG.TBE érték", "FESZ.GJ."]       # AUT.IND, BSZG.TBE — две группы без «@»: сокращения


@pytest.mark.parametrize("text", TAGS_AS_IS)
def test_tag_shape_is_recognized(text):
    import conventions as C
    assert C.is_tag(text)


@pytest.mark.parametrize("text", NOT_TAGS)
def test_ordinary_text_is_not_a_tag(text):
    import conventions as C
    assert not C.is_tag(text)


def _tr(tmp_path):
    return T.Translator(glossary=_project(tmp_path)["glossary"])


def test_run_batch_does_not_send_tags_to_the_portal(tmp_path, monkeypatch):
    """01-02FMP0BIGT, id 1703: «@.BSZG.TBE» → «@.ИВС (BSZG).температура на входе в реактор» (портал расшифровал тег,
    проверка кодов пропустила: у тега нет цифр и «_»)."""
    asked = []

    def call(self, units, *a, **k):
        asked.extend(u["text"] for u in units)
        return {i: "@.ИВС (BSZG).температура на входе в реактор" for i, u in enumerate(units)}, {}
    monkeypatch.setattr(T.Translator, "_call", call)
    out = _tr(tmp_path).run_batch({"units": [{"key": 0, "text": "@.BSZG.TBE"}, {"key": 1, "text": "FIGYELEM!"}],
                                   "first_abbr": [], "context": ""})
    assert asked == ["FIGYELEM!"]
    assert out[0] == {"ru": "@.BSZG.TBE", "issues": [], "tries": 0, "copied": True}
    assert out[1]["ru"] == "@.ИВС (BSZG).температура на входе в реактор"      # прочее идёт как прежде


def test_translate_keeps_tags_as_is(tmp_path, monkeypatch):
    asked = []

    def call(self, units, *a, **k):
        asked.extend(u["text"] for u in units)
        return {i: "Идентификатор" for i, u in enumerate(units)}, {}
    monkeypatch.setattr(T.Translator, "_call", call)
    work = tmp_path / "work"
    (work / "json").mkdir(parents=True)
    segs = [{"id": i, "text": t, "src": t, "page": 1, "loc": {}} for i, t in
            enumerate(["@.BSZG.TBE", "Azonosító", "@.BSZG.BC", "<f1>@.BSZG.TBE</f1>"])]
    (work / "json" / "doc.json").write_text(json.dumps({"file": "x.pdf", "segs": segs, "res": {}, "info": {}}),
                                            encoding="utf-8")
    P.translate(work, _project(tmp_path))
    res = json.loads((work / "json" / "doc.json").read_text(encoding="utf-8"))["res"]
    assert asked == ["Azonosító"]
    assert [res[str(i)]["ru"] for i in (0, 2, 3)] == ["@.BSZG.TBE", "@.BSZG.BC", "<f1>@.BSZG.TBE</f1>"]
    assert all(res[str(i)].get("copied") for i in (0, 2, 3)) and res["1"]["ru"] == "Идентификатор"


def test_polish_restores_tag_translated_by_older_build():
    """JSON, переведённый до правки: в res уже стоит расшифрованный тег — шлифовка (код, без портала) возвращает его."""
    import conventions as C
    import polish as PL
    rules = {"latin": [], "cyrillic_to_latin": {}}
    hu, bad = "@.BSZG.TBE", "@.ИВС (BSZG).температура на входе в реактор"
    assert [v[0] for v in C.check(hu, bad, rules)] == ["Я6"] and C.check(hu, bad, rules)[0][2] is True
    assert C.check(hu, hu, rules) == []
    items = {"1": {"hu": hu, "ru": bad, "kind": "text", "page": 25, "id": 1703}}
    edits, left = PL.run(items, rules, chat=None, log=lambda *a: None)
    assert items["1"]["ru"] == hu and left == [] and edits[0]["who"] == "код"
