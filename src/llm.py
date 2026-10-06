# -*- coding: utf-8 -*-
"""Клиент портала go.ai-rosatom.ru (Open WebUI, OpenAI-совместимый) для перевода.
Токен — ROSATOM_AI_TOKEN из окружения или .env в корне репозитория, в вывод не печатается.
Кэш на диск по sha1(модель+параметры+сообщения): повтор прогона стоит ноль."""
import sys, json, time, base64, hashlib, threading, urllib.request, urllib.error
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
    if f.exists():
        with _lock:
            STAT["cached"] += 1
        return json.loads(f.read_text(encoding="utf-8"))
    with _lock:
        if _TOK is None:
            _TOK = _token()
    payload = {"model": MODEL, "stream": False, "temperature": temperature, "max_tokens": max_tokens,
               "messages": msgs}
    if think:
        payload["chat_template_kwargs"] = {"enable_thinking": True}
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
            f.write_text(json.dumps(out, ensure_ascii=False, indent=1), encoding="utf-8")
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
