# -*- coding: utf-8 -*-
"""Словарь подписей известных форм (figures.FORMS) — из готовых переводов корпуса (E7, 7б).

    python -X utf8 tools/form_labels.py <job.json | папка с json> ... [--list файл-со-списком] [--out src/form_labels.json]
                                        [--min-docs 4] [--share 0.75]

Форма опознаётся по заголовку («NYILVÁNTARTÓ LAP», «OPERATÍV ÜZEMVITELI TESZTELÉSI UTASÍTÁS», «MÓDOSÍTÁS JEGYZÉK»); для
каждой формы берётся подпись (текст сегмента без «|» по краям — линий ячеек, прочитанных как знак), встретившаяся в
≥ min-docs разных документах, и её перевод — тот, что у ≥ share случаев. Подписи с цифрой, фамилией, числом, одним знаком
(«O», «X», «•») и точками-заполнителями в словарь не идут: это значения бланка, отметки и пиксели оригинала. Словарь —
данные: читайте файл глазами и правьте перед приёмкой."""
import argparse
import json
import re
import sys
from collections import Counter, defaultdict
from pathlib import Path

REPO = Path(__file__).resolve().parent.parent
# вычитано глазами: строки-описания изменений листа изменений (повторяются в документах, но это значения бланка, не подписи)
DROP = {"Armatúra utánhúzás megszüntetése", "Kiindulási állapot aktualizálása.", "Próba utáni állapot pontosítása.",
        'Érintett fejezetek: "'}
sys.path.insert(0, str(REPO / "src"))
import figures as FIG  # noqa: E402
import pagefit as PF  # noqa: E402


def collect(files):
    labels = defaultdict(lambda: defaultdict(lambda: defaultdict(Counter)))     # форма → подпись → документ → переводы
    for f in files:
        d = json.loads(Path(f).read_text(encoding="utf-8"))
        form = FIG.known_form(d["segs"])
        if not form:
            continue
        for s in d["segs"]:
            ru = (d["res"].get(str(s["id"])) or {}).get("ru")
            hu = FIG.label_key(s["text"])
            if ru and hu and FIG.is_label(hu):
                labels[form][hu][Path(f).stem][FIG.label_key(ru)] += 1
    return labels


def caps(t):
    """Подпись набрана заглавными (все буквы) — перевод должен быть набран так же, и наоборот (голосование смешивало
    «Nukleáris biztonságot érinti:» и «NUKLEÁRIS BIZTONSÁGOT ÉRINTI:» — подпись строчными получала заглавный перевод)."""
    letters = [c for c in t if c.isalpha()]
    return all(c.isupper() for c in letters) if len(letters) >= 4 else None         # «Ф. И. О.», «ÜBO» — не в счёт


def build(labels, min_docs=4, share=0.75):
    out = {}
    for form, texts in labels.items():
        rows = {}
        for hu, docs in texts.items():
            if len(docs) < min_docs:
                continue
            votes = Counter()
            for c in docs.values():
                votes.update(c)
            ru, n = votes.most_common(1)[0]
            if n >= share * sum(votes.values()) and ru and hu not in DROP and (None in (caps(hu), caps(ru)) or caps(hu) == caps(ru)):
                rows[hu] = ru
        out[form] = dict(sorted(rows.items()))
    return out


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("paths", nargs="*")
    ap.add_argument("--list", help="файл: по пути JSON на строку (имена с пробелами)")
    ap.add_argument("--out", default=str(REPO / "src" / "form_labels.json"))
    ap.add_argument("--min-docs", type=int, default=4)
    ap.add_argument("--share", type=float, default=0.75)
    a = ap.parse_args()
    files = []
    if a.list:
        a.paths += [ln.strip() for ln in Path(a.list).read_text(encoding="utf-8").splitlines() if ln.strip()]
    for p in map(Path, a.paths):
        files += sorted(p.glob("*.json")) if p.is_dir() else [p]
    out = build(collect(files), a.min_docs, a.share)
    Path(a.out).write_text(json.dumps(out, ensure_ascii=False, indent=1, sort_keys=True) + "\n", encoding="utf-8")
    print({k: len(v) for k, v in out.items()}, "→", a.out)


if __name__ == "__main__":
    main()
