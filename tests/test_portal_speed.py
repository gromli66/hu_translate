# -*- coding: utf-8 -*-
"""Скорость пути перевода (10.11): независимые запросы к порталу — параллельно, выдача та же, что при запросах по
одному. Повторы упавших сегментов пачки, пачки добора структуры, сокращения под место; общий потолок запросов.
Портала нет — подставные _call / chat / urlopen."""
import json
import threading
import time

import pytest

import llm
import pipeline as P
import translate as T


class _Peak:
    """Сколько вызовов идёт одновременно (наибольшее)."""
    def __init__(self):
        self.now = self.peak = 0
        self.lock = threading.Lock()

    def __enter__(self):
        with self.lock:
            self.now += 1
            self.peak = max(self.peak, self.now)

    def __exit__(self, *a):
        with self.lock:
            self.now -= 1


def _glossary(tmp_path):
    g = tmp_path / "g.md"
    g.write_text("## Т\n\n| Венгерский | Русский | Комментарий |\n|---|---|---|\n| szelep | клапан | |\n", encoding="utf-8")
    return str(g)


# ---------------------------------------------------------------- повторы упавших сегментов пачки
def _run_batch(tmp_path, monkeypatch, workers):
    """Пачка из 8 сегментов с числами: первый ответ без чисел (не проходит проверку), повтор — верный у чётных,
    снова без числа у нечётных (две попытки, лучший результат)."""
    monkeypatch.setattr(T, "RETRY_WORKERS", workers)
    peak, calls = _Peak(), []

    def call(self, units, first_abbr, context, think, note=""):
        calls.append((tuple(u["text"] for u in units), note))
        if not note:
            return {i: "клапан" for i in range(len(units))}, {}
        with peak:
            time.sleep(0.05)
        n = units[0]["text"].split()[1]
        return {0: f"клапан {n}" if int(n) % 2 == 0 else "клапан закрыт"}, {}
    monkeypatch.setattr(T.Translator, "_call", call)
    tr = T.Translator(glossary=_glossary(tmp_path))
    units = [{"key": k, "text": f"szelep {10 + k}"} for k in range(8)]
    out = tr.run_batch({"units": units, "first_abbr": [], "context": "előző"})
    return out, sorted(calls), peak.peak


def test_retries_in_parallel_give_the_same_result(tmp_path, monkeypatch):
    seq, calls1, peak1 = _run_batch(tmp_path, monkeypatch, 1)
    par, calls6, peak6 = _run_batch(tmp_path, monkeypatch, 6)
    assert par == seq and calls6 == calls1          # те же запросы, тот же итог
    assert peak1 == 1 and peak6 > 1                 # прежде — по одному, теперь — одновременно
    assert seq[0]["ru"] == "клапан 10" and seq[0]["tries"] == 2 and not seq[0]["issues"]
    assert seq[1]["ru"] == "клапан" and seq[1]["tries"] == 3 and seq[1]["issues"]   # две неудачные попытки
    assert sum(1 for c in calls1 if c[1]) == 4 + 4 * 2                              # повторов: чётным 1, нечётным 2


# ---------------------------------------------------------------- добор структуры
def _struct_job(tmp_path):
    tmp_path.mkdir(parents=True)
    pdf = tmp_path / "doc.pdf"
    pdf.write_bytes(b"")
    segs = [{"id": 1, "text": "Szelep zárva.", "src": "Szelep zárva.", "page": 0, "loc": {}}]
    work = tmp_path / "work"
    (work / "json").mkdir(parents=True)
    (work / "json" / "doc.json").write_text(json.dumps({
        "file": str(pdf), "segs": segs, "res": {"1": {"ru": "Клапан закрыт.", "issues": []}}, "info": {},
        "struct": {"coverage": 1.0}}), encoding="utf-8")
    return work


def _struct_fill(tmp_path, monkeypatch, workers):
    import docx_build as B
    monkeypatch.setattr(P, "STRUCT_WORKERS", workers)
    texts = [f"{k}. szelep" for k in range(30)] + ["0. szelep"]          # 3 пачки; повтор текста — как в missing
    rounds = iter([[(t, k % 3, ()) for k, t in enumerate(texts)], []])
    monkeypatch.setattr(B, "missing", lambda *a, **k: next(rounds))
    peak = _Peak()

    def call(self, units, first_abbr, context, think, note=""):
        with peak:
            time.sleep(0.05)
        return {i: "клапан " + u["text"].split(".")[0] for i, u in enumerate(units)}, {}
    monkeypatch.setattr(T.Translator, "_call", call)
    work = _struct_job(tmp_path / str(workers))
    P.struct_fill(work, {"glossary": _glossary(tmp_path)})
    return json.loads((work / "json" / "doc.json").read_text(encoding="utf-8")), peak.peak


def test_struct_fill_in_parallel_gives_the_same_json(tmp_path, monkeypatch):
    seq, peak1 = _struct_fill(tmp_path, monkeypatch, 1)
    par, peak4 = _struct_fill(tmp_path, monkeypatch, 4)
    seq["file"] = par["file"] = ""
    assert par == seq
    assert peak1 == 1 and peak4 > 1
    assert seq["struct"]["filled"] == 31


# ---------------------------------------------------------------- сокращение под место
def test_fit_asks_in_parallel_and_gives_the_same_json(tmp_path, monkeypatch):
    import assemble as A
    long_ru = "Клапан подачи питательной воды закрыт и заблокирован оператором"

    def job(name):
        work = tmp_path / name / "work"
        (work / "json").mkdir(parents=True)
        segs = [{"id": k, "text": f"Szelep {k} zárva", "page": 0} for k in range(6)]
        res = {str(k): {"ru": f"{long_ru} {k}", "issues": []} for k in range(6)}
        res["5"].update(ru_pdf="старое", ru_pdf_of=res["5"]["ru"])                 # прежнее сокращение в силе
        (work / "json" / "doc.json").write_text(json.dumps({"file": "doc.pdf", "segs": segs, "res": res}),
                                                encoding="utf-8")
        return work
    monkeypatch.setattr(A, "prepare_translation", lambda d, stem, sink: (d,))
    monkeypatch.setattr(A, "build_pdf", lambda *a, **k: {"shrunk": [{"id": k, "scale": 0.5} for k in range(6)]})

    def run(workers):
        monkeypatch.setattr(P, "FIT_WORKERS", workers)
        peak, asked, memo = _Peak(), [], {}

        def chat(messages, max_tokens=8000, tag="", **k):          # кэш как у llm.chat: тот же запрос — без портала
            key = messages[0]["content"]
            if key not in memo:
                with peak:
                    time.sleep(0.05)
                asked.append(tag)
                k_ = tag[3:]
                memo[key] = {"text": f"Клапан ПВ закрыт {k_}" if k_ != "3" else long_ru + " 3 и ещё"}
            return memo[key]
        monkeypatch.setattr(llm, "chat", chat)
        work = job(f"w{workers}")
        P.fit(work, {})
        return json.loads((work / "json" / "doc.json").read_text(encoding="utf-8")), sorted(asked), peak.peak

    seq, asked1, peak1 = run(1)
    par, asked4, peak4 = run(4)
    assert par == seq and asked4 == asked1 == ["fit0", "fit1", "fit2", "fit3", "fit4"]   # каждый запрос — один раз
    assert peak1 == 1 and peak4 > 1
    assert seq["res"]["0"]["ru_pdf"] == "Клапан ПВ закрыт 0"
    assert "ru_pdf" not in seq["res"]["3"]                     # длиннее — не принято
    assert seq["res"]["5"]["ru_pdf"] == "старое"


# ---------------------------------------------------------------- общий потолок запросов к порталу
def test_portal_cap_limits_simultaneous_requests(tmp_path, monkeypatch):
    peak = _Peak()

    class Resp:
        def __init__(self, body):
            self.body = body

        def read(self):
            return self.body

    def urlopen(req, timeout=None):
        with peak:
            time.sleep(0.05)
        text = json.loads(req.data)["messages"][0]["content"]
        return Resp(json.dumps({"choices": [{"message": {"content": "ok " + text}, "finish_reason": "stop"}],
                                "usage": {}}).encode())
    monkeypatch.setattr(llm, "CACHE", tmp_path)
    monkeypatch.setattr(llm, "_TOK", "test")
    monkeypatch.setattr(llm.urllib.request, "urlopen", urlopen)
    monkeypatch.setattr(llm, "_par", threading.BoundedSemaphore(2))
    ts = [threading.Thread(target=llm.chat, args=([{"role": "user", "content": str(k)}],)) for k in range(8)]
    for t in ts:
        t.start()
    for t in ts:
        t.join()
    assert peak.peak == 2
    assert len(list(tmp_path.glob("*.json"))) == 8

