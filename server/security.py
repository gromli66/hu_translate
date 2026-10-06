# -*- coding: utf-8 -*-
"""Пароли (scrypt из стандартной библиотеки) и шифрование токенов портала (Fernet).
Ключ — HUT_SECRET_KEY из окружения; если не задан, создаётся data/secret.key при первом запуске.
Потеря ключа = сохранённые токены не читаются, пользователи вписывают их заново."""
import os
import hmac
import hashlib

from cryptography.fernet import Fernet, InvalidToken

from server import config

_N, _R, _P = 2 ** 14, 8, 1


def hash_pw(pw):
    salt = os.urandom(16)
    return "scrypt$" + salt.hex() + "$" + hashlib.scrypt(pw.encode(), salt=salt, n=_N, r=_R, p=_P).hex()


def check_pw(pw, stored):
    try:
        _, salt, h = stored.split("$")
    except ValueError:
        return False
    return hmac.compare_digest(hashlib.scrypt(pw.encode(), salt=bytes.fromhex(salt), n=_N, r=_R, p=_P).hex(), h)


def _fernet():
    key = os.environ.get("HUT_SECRET_KEY", "").strip()
    if not key:
        f = config.DATA / "secret.key"
        if not f.exists():
            config.DATA.mkdir(parents=True, exist_ok=True)
            fd = os.open(f, os.O_WRONLY | os.O_CREAT | os.O_EXCL, 0o600)
            with os.fdopen(fd, "w") as out:
                out.write(Fernet.generate_key().decode())
        key = f.read_text().strip()
    return Fernet(key.encode())


def enc(token):
    return _fernet().encrypt(token.encode()).decode()


def dec(blob):
    """Расшифрованный токен или None (нет токена, другой ключ сервера)."""
    if not blob:
        return None
    try:
        return _fernet().decrypt(blob.encode()).decode()
    except InvalidToken:
        return None
