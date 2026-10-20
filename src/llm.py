# -*- coding: utf-8 -*-
"""Клиент портала go.ai-rosatom.ru (Open WebUI, OpenAI-совместимый) для перевода.
Токен — ROSATOM_AI_TOKEN из окружения или .env в корне репозитория, в вывод не печатается.
Кэш на диск по sha1(модель+параметры+сообщения): повтор прогона стоит ноль."""
import sys, json, time, base64, hashlib, tempfile, threading, urllib.request, urllib.error
from pathlib import Path

import os

ROOT = Path(__file__).resolve().parent.parent


def _env(name, default=""):
    """Переменная окружения или строка NAME=... из .env в корне репозитория (.env — в .gitignore)."""
    v = os.environ.get(name, "").strip()
    if v:
        return v
    f = ROOT / ".env"
    if f.exists():
        for ln in f.read_text(encoding="utf-8", errors="replace").splitlines():
            if ln.strip().startswith(name + "="):
                return ln.split("=", 1)[1].strip().strip('"').strip("'")
    return default


BASE = _env("ROSATOM_AI_BASE", "https://go.ai-rosatom.ru").rstrip("/")
MODEL = _env("ROSATOM_MODEL", "privateLLM")


def _token():
    t = _env("ROSATOM_AI_TOKEN")
    if not t:
        raise RuntimeError("нет токена портала: задайте ROSATOM_AI_TOKEN в окружении или в .env в корне репозитория")
    return t


CACHE = Path(_env("HU_CACHE", str(ROOT / "cache"))); CACHE.mkdir(parents=True, exist_ok=True)
_TOK = None
_lock = threading.Lock()
STAT = {"calls": 0, "cached": 0, "sec": 0.0, "out_tok": 0, "in_tok": 0}


def _busy_retry(fn):
    """Windows: пока файл открыт другим или его подменяет os.replace другого процесса, открытие и replace
    дают PermissionError — повтор в пределах секунды, потом отказ наружу."""
    for _ in range(20):
        try:
            return fn()
        except PermissionError:
            time.sleep(0.05)
    return fn()


def _read_cache(f):
    """Запись кэша или None (промах). Испорченная или пустая запись — тоже промах с причиной в журнал:
    ответ перезапросится и перезапишет её (20.10 два ответа склеились в одном файле — сборка падала)."""
    try:
        b = _busy_retry(f.read_bytes)
    except FileNotFoundError:
        return None
    except PermissionError as e:
        print(f"llm: кэш {f.name} не прочитан ({e!r}) — перезапрос", file=sys.stderr, flush=True)
        return None
    try:
        d = json.loads(b.decode("utf-8"))
    except ValueError as e:  # UnicodeDecodeError (обрыв посреди буквы) и JSONDecodeError (пусто, склейка, обрыв)
        why = f"не JSON: {e}"
    else:
        if isinstance(d, dict) and isinstance(d.get("text"), str) and d["text"].strip():
            return d
        why = "нет текста ответа"
    print(f"llm: кэш {f.name} испорчен ({why}, {len(b)} байт) — перезапрос", file=sys.stderr, flush=True)
    return None


def _write_cache(f, out):
    """Атомарно: во временный файл в той же папке, затем os.replace — читатель видит старую запись
    или новую целиком. Не записалось — ответ всё равно возвращается, просто не кэшируется."""
    fd, tmp = tempfile.mkstemp(dir=f.parent, prefix=f.name + ".", suffix=".tmp")
    try:
        with os.fdopen(fd, "w", encoding="utf-8") as h:
            h.write(json.dumps(out, ensure_ascii=False, indent=1))
        _busy_retry(lambda: os.replace(tmp, f))
    except OSError as e:
        print(f"llm: кэш {f.name} не записан ({e!r}) — ответ не кэширован", file=sys.stderr, flush=True)
        Path(tmp).unlink(missing_ok=True)


def chat(messages, max_tokens=8000, think=False, temperature=0, tag="", images=None, sampling=None):
    """messages: список {role, content}. images: base64-JPEG/PNG (не больше двух) к последнему user.
    Возвращает dict: text, usage, sec, reasoning_len (или error)."""
    global _TOK
    msgs = json.loads(json.dumps(messages))
    if images:
        if len(images) > 2:
            raise ValueError("портал принимает не больше двух картинок")
        last = msgs[-1]
        last["content"] = [{"type": "text", "text": last["content"]}] + [
            {"type": "image_url", "image_url": {"url": u}} for u in images]
    # sampling (top_p, top_k…) — в ключ кэша только если задан: старые ответы остаются в силе
    key = hashlib.sha1(json.dumps([MODEL, max_tokens, think, temperature, msgs] + ([sampling] if sampling else []),
                                  ensure_ascii=False).encode()).hexdigest()
    f = CACHE / f"{key}.json"
    hit = _read_cache(f)
    if hit is not None:
        with _lock:
            STAT["cached"] += 1
        return hit
    with _lock:
        if _TOK is None:
            _TOK = _token()
    payload = {"model": MODEL, "stream": False, "temperature": temperature, "max_tokens": max_tokens,
               "messages": msgs}
    # рассуждение — только по заказу, и «без рассуждения» — явно: с 19.10 портал рассуждает по умолчанию, рассуждение
    # съедает лимит ответа (200 токенов на «переведи одну фразу») и ответ приходит пустым — перевод вставал
    payload["chat_template_kwargs"] = {"enable_thinking": bool(think)}
    if sampling:
        payload.update(sampling)
    data = json.dumps(payload).encode()
    last = None
    for att in range(5):
        req = urllib.request.Request(f"{BASE}/api/chat/completions", data=data, method="POST")
        req.add_header("Authorization", "Bearer " + _TOK)
        req.add_header("Content-Type", "application/json")
        t0 = time.time()
        try:
            raw = urllib.request.urlopen(req, timeout=900).read().decode()
            d = json.loads(raw)
            if not d.get("choices"):
                # портал под нагрузкой отвечает 200 без choices (01.10, ~12:10–12:25) — текст сбоя в журнал
                last = "нет choices: " + raw[:200]
                with _lock:
                    STAT["fail"] = STAT.get("fail", 0) + 1
                time.sleep(5 * (att + 1)); continue
            msg = d["choices"][0]["message"]
            txt = (msg.get("content") or "").strip()
            if not txt:
                last = "пустой content"; time.sleep(2 * (att + 1)); continue
            out = {"tag": tag, "text": txt, "usage": d.get("usage"), "sec": round(time.time() - t0, 1),
                   "finish": d["choices"][0].get("finish_reason"),
                   "reasoning_len": len(msg.get("reasoning_content") or msg.get("reasoning") or "")}
            _write_cache(f, out)
            with _lock:
                STAT["calls"] += 1; STAT["sec"] += out["sec"]
                u = out["usage"] or {}
                STAT["out_tok"] += u.get("completion_tokens") or 0
                STAT["in_tok"] += u.get("prompt_tokens") or 0
            return out
        except urllib.error.HTTPError as e:
            body = e.read().decode(errors="replace")[:300]
            if e.code in (401, 403):
                raise RuntimeError(f"HTTP {e.code}: {body}")
            last = f"HTTP {e.code}: {body}"
        except Exception as e:  # сеть/таймаут — повтор, причина уходит в итог
            last = repr(e)
        time.sleep(5 * (att + 1))
    return {"tag": tag, "text": "", "error": str(last)}


def img_b64(png_bytes, mime="image/png"):
    return f"data:{mime};base64," + base64.b64encode(png_bytes).decode()
