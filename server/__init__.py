# -*- coding: utf-8 -*-
"""Веб-сервис переводчика. Модули движка (src/) импортируются напрямую — путь добавляется здесь."""
import sys
from pathlib import Path

_SRC = str(Path(__file__).resolve().parent.parent / "src")
if _SRC not in sys.path:
    sys.path.insert(0, _SRC)
