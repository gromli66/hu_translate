# -*- coding: utf-8 -*-
"""Чтение листов Surya для сверки копии скана (src/copycheck.py). Запускается питоном окружения с Surya и моделями
(на этой машине — surya_env с видеокартой; в контейнер переводчика Surya пока не встроена, решение Максима 12.11).
usage: <python surya_env> surya_read.py <выход.json> <лист1.png> [<лист2.png> …]
Выход: [[{"t": текст строки, "box": [x0, y0, x1, y1] px, "poly": [[x, y] × 4] — повёрнутый четырёхугольник строки
(перекос листа и края строк для раскладки копии, src/copylayout.py), "c": уверенность}, …] по листам]."""
import json
import sys

from PIL import Image
from surya.detection import DetectionPredictor
from surya.foundation import FoundationPredictor
from surya.recognition import RecognitionPredictor

out, pngs = sys.argv[1], sys.argv[2:]
rec, det = RecognitionPredictor(FoundationPredictor()), DetectionPredictor()
res = []
for png in pngs:
    pr = rec([Image.open(png).convert("RGB")], det_predictor=det, math_mode=False)[0]
    res.append([{"t": ln.text, "box": [round(v, 1) for v in ln.bbox],
                 "poly": [[round(v, 1) for v in pt] for pt in ln.polygon], "c": round(ln.confidence or 0, 3)}
                for ln in pr.text_lines])
with open(out, "w", encoding="utf-8") as f:
    json.dump(res, f, ensure_ascii=False)
print(f"surya: листов {len(res)}, строк {sum(len(r) for r in res)}")
