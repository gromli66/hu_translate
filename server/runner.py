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
W_REVIEW = {"извлечение": (0, 5), "перевод": (5, 50), "проверка": (50, 92), "сборка": (92, 96), "термины": (96, 100)}
W_FAST = {"извлечение": (0, 5), "перевод": (5, 85), "сборка": (85, 92), "термины": (92, 100)}
W_EDITS = {"правки": (0, 70), "сборка": (70, 85), "термины": (85, 100)}
TERMS_TOP = 100            # сколько частых сочетаний вне глоссария переводить в кандидаты за один перевод


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


def mine_terms(work, prj, out, pr):
    """Кандидаты в термины из текста комплекта: частые сочетания вне глоссариев → перевод порталом (только
    уверенные). Эксперт решает, что из этого станет термином. Сбой сбора не портит готовый перевод."""
    import termx
    import segments as S
    pr("термины", 0, 1)
    try:
        texts = sorted({s["text"] for f in P.doc_files(work) for s in json.loads(f.read_text(encoding="utf-8"))["segs"]
                        if S.needs_translation(s["text"])})
        # одиночные короткие слова (reaktor, turbina, csoport) портал переводит очевидно — эксперту это шум
        # (проба 06.10: 10 из 10 кандидатов на пробном комплекте были такими); оставляем сочетания и длинные
        # составные слова вроде tápszivattyú
        cands = [c for c in termx.candidates(texts, P.make_translator(prj).entries, TERMS_TOP * 3)
                 if len(c[0].split()) >= 2 or len(c[0]) >= 10][:TERMS_TOP]
        res = [{"hu": t["hu"], "ru": t["ru"], "count": t["count"], "example": t["example"]}
               for t in termx.translate_terms(cands, progress=lambda d, t: pr("термины", d, t)) if t["sure"]]
        out.write_text(json.dumps(res, ensure_ascii=False), encoding="utf-8")
        print(f"кандидатов в термины: {len(res)} из {len(cands)}", flush=True)
    except Exception as e:                       # noqa: BLE001 — кандидаты необязательны, перевод уже готов
        print(f"кандидаты в термины не собраны: {type(e).__name__}: {e}", flush=True)
    pr("термины", 1, 1)


def edit_terms(pend, out, pr):
    """Какие термины заменили редакторы: [венгерский, стало, было, автор] → кандидаты с автором правки."""
    import termx
    pr("термины", 0, 1)
    try:
        res = []
        for p in pend:
            if len(p) >= 4:
                res += [dict(t, author=p[3]) for t in termx.from_edits([(p[0], p[2], p[1])])]
        out.write_text(json.dumps(res, ensure_ascii=False), encoding="utf-8")
        print(f"терминов из правок: {len(res)}", flush=True)
    except Exception as e:                       # noqa: BLE001 — правки уже применены, кандидаты необязательны
        print(f"термины из правок не собраны: {type(e).__name__}: {e}", flush=True)
    pr("термины", 1, 1)


def main(job_dir, mode="translate"):
    jd = Path(job_dir)
    spec = json.loads((jd / "job.json").read_text(encoding="utf-8"))
    prj, work = P.load_project(spec["project"]), jd / "work"
    if mode == "edits":
        pend = jd / "edits_pending.json"
        items = json.loads(pend.read_text(encoding="utf-8")) if pend.exists() else []
        pr = Progress(jd / "progress.json", W_EDITS)
        if items:
            P.propagate(work, prj, [(p[0], p[1]) for p in items], progress=pr)
        P.assemble(work, pr)
        pend.unlink(missing_ok=True)          # только после успешной пересборки: при сбое правки не теряются
        edit_terms(items, jd / "edit_terms.json", pr)
        pr.write({"stage": "готово", "pct": 100, "summary": summary(work)})
        return
    prj["tm"] = str(jd / "tm.json")           # утверждённая память проекта + неподтверждённая память владельца
    pr = Progress(jd / "progress.json", W_REVIEW if spec["review"] else W_FAST)
    P.extract([jd / "in"], work, pr)
    P.translate(work, prj, progress=pr)
    if spec["review"]:
        P.review(work, prj, progress=pr)
    P.assemble(work, pr)
    mine_terms(work, prj, jd / "candidates.json", pr)
    pr.write({"stage": "готово", "pct": 100, "summary": summary(work)})


if __name__ == "__main__":
    main(*sys.argv[1:3])
