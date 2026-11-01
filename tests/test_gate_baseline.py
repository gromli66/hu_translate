# -*- coding: utf-8 -*-
"""tools/gate_baseline.py: корпус, команды Docker (без Docker), разбор сборки и сравнение «до/после» на синтетических папках."""
import json
import sys
from pathlib import Path

import pymupdf
import pytest

ROOT = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(ROOT / "tools"))
import gate_baseline as G  # noqa: E402


def make_pdf(path, pages):
    doc = pymupdf.open()
    for _ in range(pages):
        doc.new_page()
    doc.save(str(path))


def make_build(d, docs, orig=None):
    """docs: {stem: (число листов, [(проверка, уровень), ...])}; orig: {stem: листов оригинала} для _audit.txt."""
    d.mkdir(parents=True, exist_ok=True)
    audit = {}
    for stem, (pages, finds) in docs.items():
        audit[stem] = [dict(check=c, level=lv, msg="m", page=None, orig_page=None, frag="") for c, lv in finds]
        make_pdf(d / (stem + "_RU.pdf"), pages)
    (d / "_audit.json").write_text(json.dumps(audit, ensure_ascii=False), encoding="utf-8")
    if orig:
        (d / "_audit.txt").write_text("\n".join("══ %s  (перевод: %d л., оригинал: %d л.)\n  ✔ замечаний нет" % (s, docs[s][0], o)
                                               for s, o in orig.items()), encoding="utf-8")
    return d


# ---------------------------------------------------------------- корпус
def make_jobs(root):
    for uid, names in {"20261029-204130-aaaaaa": ["3PR10_X", "30BLM1M2GT"], "20261029-204133-bbbbbb": ["1SZ01_Y"],
                       "20261029-184418-cccccc": ["old"], "20261030-063659-dddddd": ["3PR10_X"]}.items():
        j = root / "data" / "jobs" / uid / "work" / "json"
        j.mkdir(parents=True)
        for n in names:
            (j / (n + ".json")).write_text("{}", encoding="utf-8")


def test_corpus_takes_only_night_jobs_sorted(tmp_path):
    make_jobs(tmp_path)
    assert [f.stem for f in G.corpus(tmp_path)] == ["1SZ01_Y", "30BLM1M2GT", "3PR10_X"]
    assert [f.stem for f in G.corpus(tmp_path, ["3pr"])] == ["3PR10_X"]


def test_corpus_rejects_duplicate_names(tmp_path):
    make_jobs(tmp_path)
    j = tmp_path / "data" / "jobs" / "20261029-204131-eeeeee" / "work" / "json"
    j.mkdir(parents=True)
    (j / "3PR10_X.json").write_text("{}", encoding="utf-8")
    with pytest.raises(SystemExit, match="3PR10_X"):
        G.corpus(tmp_path)


def test_container_path_maps_data_to_app_data(tmp_path):
    f = tmp_path / "data" / "jobs" / "u" / "work" / "json" / "a.json"
    assert G.container_path(tmp_path, f) == "/app/data/jobs/u/work/json/a.json"


# ---------------------------------------------------------------- метка, src
@pytest.mark.parametrize("label", ["base", "after_E3", "e0-1.2", "x" * 41])
def test_label_ok(label):
    assert G.check_label(label) == label


@pytest.mark.parametrize("label", ["", "a b", "../x", "до_E3", "a/b", "-x", "x" * 42])
def test_label_rejected(label):
    with pytest.raises(SystemExit):
        G.check_label(label)


def test_resolve_src_accepts_worktree_or_src(tmp_path):
    (tmp_path / "src").mkdir()
    (tmp_path / "src" / "assemble.py").write_text("", encoding="utf-8")
    assert G.resolve_src(tmp_path) == tmp_path / "src"
    assert G.resolve_src(tmp_path / "src") == tmp_path / "src"
    (tmp_path / "empty").mkdir()
    with pytest.raises(SystemExit, match="assemble.py"):
        G.resolve_src(tmp_path / "empty")


# ---------------------------------------------------------------- Docker (без Docker)
class FakeDocker:
    def __init__(self, running="true"):
        self.cmds, self.running = [], running

    def __call__(self, cmd, stream=None):
        self.cmds.append(cmd)
        if cmd[:2] == ["docker", "inspect"]:
            return 0, self.running + "\n"
        return 0, ""


def test_prepare_container_touches_only_tmp_label(tmp_path, monkeypatch):
    fake = FakeDocker()
    monkeypatch.setattr(G, "_run", fake)
    base = G.prepare_container("lab1", tmp_path / "src", ["/app/data/jobs/u/work/json/a.json"], with_projects=False)
    assert base == "/tmp/lab1"
    flat = [" ".join(c) for c in fake.cmds]
    assert not any(w in c for c in flat for w in (" restart", " stop ", " build", " kill", " rm hut", "compose"))
    for c in fake.cmds:
        if c[:3] == ["docker", "exec", "hut"] and c[3] in ("rm", "mkdir", "ln"):
            assert all(a.startswith("/tmp/lab1") or a.startswith("-") or a == "/app/projects" for a in c[4:]), c
        if c[:2] == ["docker", "cp"]:
            dest = c[3]
            assert dest.startswith("hut:/tmp/lab1/"), c


def test_prepare_container_refuses_stopped_container(monkeypatch):
    monkeypatch.setattr(G, "_run", FakeDocker(running="false"))
    with pytest.raises(SystemExit, match="не запущен"):
        G.prepare_container("lab1", Path("x"), [], False)


def test_run_in_container_collects_gate_lines(monkeypatch):
    rec = dict(stem="3PR10", sec=33.0, pdf_pages=31, err=None, pdf_from=G.LO_FROM, pdf_reason=None)
    lines = ["всего JSON: 1", "GATE\t" + json.dumps(rec, ensure_ascii=False), "DONE"]

    def fake(cmd, stream=None):
        for ln in lines:
            stream(ln)
        return 0, ""
    monkeypatch.setattr(G, "_run", fake)
    said = []
    code, records = G.run_in_container("/tmp/lab1", said.append)
    assert code == 0 and records == [rec]
    assert any("3PR10" in s and "LibreOffice" in s for s in said)


def test_pdf_kind_classes():
    assert G.pdf_kind(dict(err=None, pdf_from=G.LO_FROM))[0] == "lo"
    assert G.pdf_kind(dict(err=None, pdf_from=None))[0] == "place"
    k, label = G.pdf_kind(dict(err=None, pdf_from="вёрстка оригинала", pdf_reason="FileNotFoundError: soffice"))
    assert k == "other" and "soffice" in label
    assert G.pdf_kind(dict(err="KeyError: 3", pdf_from=None))[0] == "fail"


def test_summarize_run_flags_fallback_and_failures():
    ok = [dict(stem="a", err=None, pdf_from=G.LO_FROM), dict(stem="b", err=None, pdf_from=None)]
    lines, good = G.summarize_run(ok, 2)
    assert good and "через DOCX (LibreOffice) — 1" in lines[0] and "на месте — 1" in lines[0]
    bad = ok + [dict(stem="c", err=None, pdf_from="вёрстка оригинала", pdf_reason="нет soffice"),
                dict(stem="d", err="ValueError: x", pdf_from=None)]
    lines, good = G.summarize_run(bad, 4)
    assert not good
    assert any("c не через LibreOffice" in ln for ln in lines) and any("упала сборка d" in ln for ln in lines)
    _, good = G.summarize_run(ok, 3)
    assert not good                                       # собрано не всё


# ---------------------------------------------------------------- сравнение
def test_load_build_counts_levels_pages_and_orig(tmp_path):
    d = make_build(tmp_path / "b", {"X": (5, [("1 потеря содержимого", "red"), ("3 шапки", "yellow"), ("3 шапки", "red")]),
                                    "Y": (2, [])}, orig={"X": 4})
    docs = G.load_build(d)
    assert (docs["X"].red, docs["X"].yellow, docs["X"].pages, docs["X"].orig) == (2, 1, 5, 4)
    assert (docs["Y"].red, docs["Y"].yellow, docs["Y"].pages, docs["Y"].orig) == (0, 0, 2, None)
    assert docs["X"].by_check[("3 шапки", "red")] == 1
    assert G.load_build(d / "_audit.json").keys() == docs.keys()           # и по самому файлу


def test_load_build_without_audit_json_is_an_error(tmp_path):
    with pytest.raises(SystemExit, match="_audit.json"):
        G.load_build(tmp_path)


def test_judge_rules():
    D = G.Doc
    # меньше 🔴 при большем 🟡 — лучше (красное смягчилось)
    assert G.judge(D("a", 1, 0, 4, 4), D("a", 0, 1, 4, 4))[0] == "лучше"
    assert G.judge(D("a", 0, 1, 4, 4), D("a", 1, 0, 4, 4))[0] == "хуже"
    assert G.judge(D("a", 0, 1, 4, 4), D("a", 0, 2, 4, 4))[0] == "хуже"
    assert G.judge(D("a", 2, 0, 4, 4), D("a", 2, 0, 4, 4)) == ("так же", "")
    # лишние листы сверх оригинала: 6 при 4 (+2) → 7 при 4 (+3) хуже, даже когда 🔴 стало меньше
    v, why = G.judge(D("a", 3, 0, 6, 4), D("a", 1, 0, 7, 4))
    assert v == "хуже" and "листов +1" in why
    # без оригинала лишние листы — все листы
    assert G.judge(D("a", 0, 0, 3, None), D("a", 0, 0, 5, None))[0] == "хуже"
    # листов меньше оригинала — не «лишние»: 3 → 2 при оригинале 4 ничего не меняет
    assert G.judge(D("a", 0, 0, 3, 4), D("a", 0, 0, 2, 4))[0] == "так же"


def test_compare_and_render(tmp_path):
    a = make_build(tmp_path / "A", {"X": (5, [("3 шапки", "red")] * 2), "Y": (3, []), "Z": (4, [("1 потеря содержимого", "yellow")]),
                                    "OnlyA": (1, [])}, orig={"X": 4, "Y": 3, "Z": 4})
    b = make_build(tmp_path / "B", {"X": (4, [("3 шапки", "red")]), "Y": (3, [("3 шапки", "red")]), "Z": (4, [("1 потеря содержимого", "yellow")]),
                                    "New": (2, [])})
    rows = G.compare(G.load_build(a), G.load_build(b))
    verdict = {r[0]: r[3] for r in rows}
    assert verdict == {"New": "новый", "OnlyA": "хуже", "X": "лучше", "Y": "хуже", "Z": "так же"}
    text = G.render_compare(rows, "A", "B")
    lines = text.splitlines()
    first = [i for i, ln in enumerate(lines) if ln.startswith(("Y ", "OnlyA ", "X ", "New "))]
    assert [lines[i].split()[0] for i in first] == ["OnlyA", "Y", "New", "X"]       # хуже → новые → лучше
    assert "Z " not in text.replace("ИТОГО", "")                                      # неизменившийся скрыт
    assert "ХУЖЕ" in text and "лучше" in text
    assert "хуже 2, лучше 1, так же 1, новых 1" in text
    assert "2 → 1" in text and "✘ хуже документов: 2" in text
    assert "Z " in G.render_compare(rows, "A", "B", show_all=True)


def test_cmd_compare_exit_code_and_only(tmp_path, capsys):
    a = make_build(tmp_path / "A", {"X": (2, []), "Y": (2, [])})
    b = make_build(tmp_path / "B", {"X": (2, [("3 шапки", "red")]), "Y": (2, [])})
    assert G.main(["compare", str(a), str(b)]) == 1
    assert "ХУЖЕ" in capsys.readouterr().out
    assert G.main(["compare", str(a), str(b), "--only", "Y"]) == 0
    assert "✔ ничего не хуже" in capsys.readouterr().out


def test_locate_build_by_label(tmp_path, monkeypatch):
    monkeypatch.setattr(G, "gate_dir", lambda: tmp_path)
    d = make_build(tmp_path / "lab", {"X": (1, [])})
    assert G.locate_build("lab") == d
    with pytest.raises(SystemExit, match="нет сборки"):
        G.locate_build("nope")
