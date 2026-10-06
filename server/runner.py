# -*- coding: utf-8 -*-
"""Исполнитель одной задачи: отдельный процесс, который запускает очередь (server/jobs.py).
Глобальное состояние движка (токен, папка OCR, настройки проверяющего) живёт только в этом процессе.
Токен приходит в переменной окружения ROSATOM_AI_TOKEN, в командной строке и логе его нет.
Режимы: translate — перевод с нуля (память переводов из tm.json, который готовит очередь из БД);
edits — разнос правок из edits_pending.json по похожим строкам и пересборка файлов.
usage: python -m server.runner <папка задачи> [translate|edits]   (job.json, in/ → work/, progress.json)"""
import sys
import json
import time
from pathlib import Path

ROOT = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(ROOT / "src"))
sys.stdout.reconfigure(encoding="utf-8")
import pipeline as P

# доля шага в общей полосе прогресса, %
W_REVIEW = {"извлечение": (0, 5), "перевод": (5, 50), "проверка": (50, 95), "сборка": (95, 100)}
W_FAST = {"извлечение": (0, 5), "перевод": (5, 90), "сборка": (90, 100)}
W_EDITS = {"правки": (0, 80), "сборка": (80, 100)}


class Progress:
    """progress(шаг, сделано, всего) → progress.json (не чаще раза в 2 с). Время до готовности — по скорости шага."""

    def __init__(self, path, weights):
        self.path, self.w = Path(path), weights
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


def main(job_dir, mode="translate"):
    jd = Path(job_dir)
    spec = json.loads((jd / "job.json").read_text(encoding="utf-8"))
    prj, work = P.load_project(spec["project"]), jd / "work"
    if mode == "edits":
        pend = jd / "edits_pending.json"
        pairs = [tuple(p) for p in json.loads(pend.read_text(encoding="utf-8"))] if pend.exists() else []
        pr = Progress(jd / "progress.json", W_EDITS)
        if pairs:
            P.propagate(work, prj, pairs, progress=pr)
        P.assemble(work, pr)
        pend.unlink(missing_ok=True)          # только после успешной пересборки: при сбое правки не теряются
        pr.write({"stage": "готово", "pct": 100, "summary": summary(work)})
        return
    prj["tm"] = str(jd / "tm.json")           # утверждённая память проекта + неподтверждённая память владельца
    pr = Progress(jd / "progress.json", W_REVIEW if spec["review"] else W_FAST)
    P.extract([jd / "in"], work, pr)
    P.translate(work, prj, progress=pr)
    if spec["review"]:
        P.review(work, prj, progress=pr)
    P.assemble(work, pr)
    pr.write({"stage": "готово", "pct": 100, "summary": summary(work)})


if __name__ == "__main__":
    main(*sys.argv[1:3])
