# -*- coding: utf-8 -*-
"""Кэш ответов портала (20.10): два процесса писали один ответ в один файл — ответы склеились, следующее чтение
падало с JSONDecodeError «Extra data». Теперь запись атомарная, испорченная запись — промах с перезапросом.
Портал не вызывается: urlopen подменён."""
import json
import os
import subprocess
import sys
import threading
import time
from pathlib import Path

import pytest

import llm

SRC = Path(__file__).resolve().parent.parent / "src"
MSGS = [{"role": "user", "content": "Fordítsd le: szivattyú"}]


class _Resp:
    def __init__(self, text):
        self.body = json.dumps({"choices": [{"message": {"content": text}, "finish_reason": "stop"}],
                                "usage": {"prompt_tokens": 5, "completion_tokens": 2}}).encode()

    def read(self):
        return self.body


@pytest.fixture
def portal(tmp_path, monkeypatch):
    """Подменённый портал: отвечает по очереди из answers, считает вызовы; кэш — во временной папке."""
    calls = []
    answers = ["насос"]

    def urlopen(req, timeout=None):
        calls.append(json.loads(req.data))
        return _Resp(answers[min(len(calls), len(answers)) - 1])

    monkeypatch.setattr(llm, "CACHE", tmp_path)
    monkeypatch.setattr(llm, "_TOK", "test")
    monkeypatch.setattr(llm.urllib.request, "urlopen", urlopen)
    return calls, answers


def _cache_file(tmp_path):
    files = list(tmp_path.glob("*.json"))
    assert len(files) == 1
    return files[0]


def test_answer_cached_and_reused(portal, tmp_path):
    calls, _ = portal
    assert llm.chat(MSGS)["text"] == "насос"
    assert llm.chat(MSGS)["text"] == "насос"
    assert len(calls) == 1
    assert json.loads(_cache_file(tmp_path).read_text(encoding="utf-8"))["text"] == "насос"
    assert not list(tmp_path.glob("*.tmp"))


@pytest.mark.parametrize("broken", [
    '{\n "text": "насос"\n}\n "text": "насос центробежный"\n}'.encode(),  # склейка двух ответов, как 20.10
    b"",                                                                  # пустой файл
    '{\n "tag": "", "text": "нас'.encode(),                               # обрыв записи
    '{\n "tag": "", "text": "нас'.encode()[:-1],                          # обрыв посреди буквы
    '{"tag": "", "text": ""}'.encode(),                                   # без текста ответа
])
def test_broken_cache_is_miss_and_rewritten(portal, tmp_path, capsys, broken):
    calls, answers = portal
    llm.chat(MSGS)
    f = _cache_file(tmp_path)
    f.write_bytes(broken)
    answers.append("насос центробежный")
    assert llm.chat(MSGS)["text"] == "насос центробежный"     # перезапрос, а не JSONDecodeError
    assert len(calls) == 2
    assert json.loads(f.read_text(encoding="utf-8"))["text"] == "насос центробежный"
    err = capsys.readouterr().err
    assert f"кэш {f.name} испорчен" in err and "перезапрос" in err
    assert llm.chat(MSGS)["text"] == "насос центробежный" and len(calls) == 2   # снова из кэша


def test_parallel_threads_same_request(portal, tmp_path, monkeypatch):
    # все потоки промахиваются и пишут один файл ответами разной длины
    calls, _ = portal
    lock = threading.Lock()

    def urlopen(req, timeout=None):
        with lock:
            calls.append(1); n = len(calls)
        time.sleep(0.05)
        return _Resp("насос " * (n * 300))

    monkeypatch.setattr(llm.urllib.request, "urlopen", urlopen)
    got = []
    ts = [threading.Thread(target=lambda: got.append(llm.chat(MSGS)["text"])) for _ in range(8)]
    for t in ts:
        t.start()
    for t in ts:
        t.join()
    assert len(got) == 8 and all(got)
    assert json.loads(_cache_file(tmp_path).read_text(encoding="utf-8"))["text"] in got
    assert not list(tmp_path.glob("*.tmp"))


WRITER = """
import sys
sys.path.insert(0, sys.argv[1])
from pathlib import Path
import llm
f, n = Path(sys.argv[2]), int(sys.argv[3])
for i in range(150):
    llm._write_cache(f, {"tag": f"w{n}", "text": "ы" * (n * 700 + i)})
"""


def test_parallel_processes_never_see_broken_file(tmp_path, capsys):
    # несколько процессов (как три сессии 20.10) пишут один файл; читатель всё время видит целую запись
    f = tmp_path / "k.json"
    env = dict(os.environ, HU_CACHE=str(tmp_path))
    ps = [subprocess.Popen([sys.executable, "-c", WRITER, str(SRC), str(f), str(n)], env=env,
                           stderr=subprocess.PIPE) for n in range(1, 5)]
    reads = 0
    while any(p.poll() is None for p in ps):
        d = llm._read_cache(f)
        if d is None:
            assert reads == 0, capsys.readouterr().err   # промах — только пока никто не записал
            continue
        assert set(d["text"]) == {"ы"}
        reads += 1
    errs = [p.stderr.read().decode("utf-8", "replace") for p in ps]
    assert all(p.returncode == 0 for p in ps), errs
    assert reads > 0
    assert "кэш" not in capsys.readouterr().err
    assert json.loads(f.read_text(encoding="utf-8"))["tag"].startswith("w")
    assert not list(tmp_path.glob("*.tmp"))
