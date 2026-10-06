# -*- coding: utf-8 -*-
"""Исполнитель одной задачи: отдельный процесс, который запускает очередь (server/jobs.py).
Глобальное состояние движка (токен, папка OCR, настройки проверяющего) живёт только в этом процессе.
Токен приходит в переменной окружения ROSATOM_AI_TOKEN, в командной строке и логе его нет.
usage: python -m server.runner <папка задачи>   (job.json, in/ → work/, progress.json)"""
import sys
import json
import time
from pathlib import Path

ROOT = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(ROOT / "src"))
sys.stdout.reconfigure(encoding="utf-8")
import pipeline as P

# доля шага в общей полосе прогресса, %
WEIGHTS = {True: {"извлечение": (0, 5), "перевод": (5, 50), "проверка": (50, 95), "сборка": (95, 100)},
           False: {"извлечение": (0, 5), "перевод": (5, 90), "сборка": (90, 100)}}


class Progress:
    """progress(шаг, сделано, всего) → progress.json (не чаще раза в 2 с). Время до готовности — по скорости шага."""

    def __init__(self, path, review):
        self.path, self.w = Path(path), WEIGHTS[review]
        self.stage, self.t_stage, self.last = None, 0.0, 0.0

    def __call__(self, stage, done, total):
        now = time.time()
        if stage != self.stage:
            self.stage, self.t_stage, self.last = stage, now, 0.0
        if now - self.last < 2 and done < total:
            return
        self.last = now
        a, b = self.w[stage]
        eta = (now - self.t_stage) / done * (total - done) if done and total else None
        self.write({"stage": stage, "done": done, "total": total,
                    "pct": round(a + (b - a) * (done / total if total else 1), 1),
                    "eta_s": round(eta) if eta is not None else None})

    def write(self, d):
        tmp = self.path.with_suffix(".tmp")
        tmp.write_text(json.dumps(dict(d, ts=time.time()), ensure_ascii=False), encoding="utf-8")
        tmp.replace(self.path)


def summary(work):
    """Строк перевода и сколько из них просят внимания: формальные замечания или подозрение смысловой проверки."""
    segs = issues = suspicious = 0
    for f in P.doc_files(work):
        for r in json.loads(f.read_text(encoding="utf-8"))["res"].values():
            if r.get("copied"):
                continue
            segs += 1
            issues += bool(r.get("issues"))
            suspicious += bool((r.get("review") or {}).get("serious"))
    return {"segments": segs, "issues": issues, "suspicious": suspicious}


def main(job_dir):
    jd = Path(job_dir)
    spec = json.loads((jd / "job.json").read_text(encoding="utf-8"))
    prj, work = P.load_project(spec["project"]), jd / "work"
    pr = Progress(jd / "progress.json", bool(spec["review"]))
    P.extract([jd / "in"], work, pr)
    P.translate(work, prj, progress=pr)
    if spec["review"]:
        P.review(work, prj, progress=pr)
    P.assemble(work, pr)
    pr.write({"stage": "готово", "pct": 100, "summary": summary(work)})


if __name__ == "__main__":
    main(sys.argv[1])
