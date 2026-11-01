# -*- coding: utf-8 -*-
"""Мерило «до/после» по корпусу (этап E0 единой модели вёрстки): собрать 123 документа ночных заданий кодом из указанной
папки и сравнить две сборки по воротам `tools/audit_out.py`.

КАК ЗАПУСКАТЬ (из корня репозитория или worktree; нужен Docker Desktop с запущенным контейнером `hut`, сам он не
перезапускается и образ не пересобирается — код копируется только в /tmp/<метка> контейнера):

  1. ДО этапа — собрать текущий код (или взять готовую базу `work/audit_2026-10-29/final_branch`, она же `база`):
         python -X utf8 tools/gate_baseline.py build before_E3
  2. Сделать этап (правки в src/ этого worktree), прогнать тесты: python -X utf8 -m pytest tests -q
  3. ПОСЛЕ этапа — собрать заново, под другой меткой:
         python -X utf8 tools/gate_baseline.py build after_E3
  4. Сравнить (таблица по документам: 🔴/🟡/листов «до → после», пометка «хуже», итог; код возврата 1, если что-то хуже):
         python -X utf8 tools/gate_baseline.py compare база after_E3        # изменившиеся документы
         python -X utf8 tools/gate_baseline.py compare before_E3 after_E3 --all # все 123
  5. Глазами — галерея страниц: python -X utf8 tools/gallery.py before_E3 after_E3 --out <папка>  (см. tools/gallery.py).

`build <метка> [--src <worktree|src>] [--out <папка>] [--only <часть имени> ...] [--clean] [--no-audit]`
  Корпус — JSON data/jobs/20261029-2041[3-9]*/work/json/*.json (123), оригиналы — `C:\\Users\\Maksim\\Desktop\\Новая папка`
  и вложения портфолио в data/jobs/*/work/unpacked. Результат — `work/gate/<метка>/`: `out/` (DOCX и PDF), `json/` (копии JSON
  корпуса), `_audit.json`/`_audit.txt` (audit_out с --fail-on never), `_run.json` (по документу: секунды, листов, откуда PDF).
  Каждый PDF проверяется: `info["pdf"]["from"]` должен быть "DOCX (LibreOffice)". У сканов и титулов `info` без `pdf` (PDF
  делается «на месте») — это норма и в отчёте считается отдельно; любое другое значение («вёрстка оригинала» — запасной путь без
  LibreOffice) или упавшая сборка — код возврата 1, и такая сборка к сравнению не годится.
  Среда сборки повторяет `final_branch` (/tmp/vf/final): код src, пустой кэш, без каталога projects/ (`--with-projects`
  добавляет ссылку на /app/projects контейнера — тогда словарь проекта влияет на опознание меток форм, как в рабочем образе).
  Дерево сборки — минут 30–40 на 123 документа: запускать в фоне.

`compare <A> <B> [--all]`: A и B — папка сборки (с `_audit.json`), сам `_audit.json`, метка из work/gate или `база`.
  Документ «хуже», если 🔴 больше, или 🔴 те же, а 🟡 больше, или лишних против оригинала листов больше; «лучше» — если
  ничего не хуже и что-то меньше. Переход 🔴 → 🟡 — не «хуже»."""
import argparse
import contextlib
import json
import re
import shutil
import subprocess
import sys
import tempfile
from collections import Counter
from pathlib import Path

REPO = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(REPO / "tools"))

CONTAINER = "hut"
JOBS_GLOB = "20261029-2041[3-9]*"             # ночные задания: 11 папок, 123 документа
EXPECTED_DOCS = 123
ORIG_DIR = Path(r"C:\Users\Maksim\Desktop\Новая папка")
LO_FROM = "DOCX (LibreOffice)"
RED, YELLOW = "🔴", "🟡"


def data_root():
    """Корень с data/jobs: сам репозиторий или — для worktree — основной checkout выше по дереву (data/ в git не лежит)."""
    for p in [REPO, *REPO.parents]:
        if (p / "data" / "jobs").is_dir():
            return p
    raise SystemExit("не найден data/jobs ни в %s, ни выше" % REPO)


def gate_dir():
    return data_root() / "work" / "gate"


def base_dir():
    return data_root() / "work" / "audit_2026-10-29" / "final_branch"


# ---------------------------------------------------------------- корпус
def corpus(root, only=()):
    """JSON ночных заданий (пути на хосте), отсортированные по имени документа. Имена должны быть уникальны."""
    files = sorted(Path(root).glob("data/jobs/%s/work/json/*.json" % JOBS_GLOB), key=lambda f: (f.stem, str(f)))
    if only:
        files = [f for f in files if any(o.lower() in f.stem.lower() for o in only)]
    dup = [s for s, n in Counter(f.stem for f in files).items() if n > 1]
    if dup:
        raise SystemExit("в корпусе повторяются имена документов: %s" % ", ".join(dup[:5]))
    return files


def container_path(root, f):
    """Путь JSON внутри контейнера: data/ хоста смонтирована как /app/data."""
    return "/app/data/" + Path(f).relative_to(Path(root) / "data").as_posix()


# ---------------------------------------------------------------- Docker
def _run(cmd, stream=None):
    """Запуск команды; stream(line) — строки вывода по мере появления. Возвращает (код, весь вывод)."""
    p = subprocess.Popen(cmd, stdout=subprocess.PIPE, stderr=subprocess.STDOUT, encoding="utf-8", errors="replace")
    out = []
    for line in p.stdout:
        out.append(line)
        if stream:
            stream(line.rstrip("\n"))
    return p.wait(), "".join(out)


def check_label(label):
    if not re.fullmatch(r"[\w][\w.\-]{0,40}", label or "", re.A):
        raise SystemExit("метка «%s» не годится: латиница/цифры/_.- до 41 знака (это имя папки /tmp/<метка> в контейнере)" % label)
    return label


def resolve_src(path):
    p = Path(path).resolve()
    src = p if (p / "assemble.py").exists() else p / "src"
    if not (src / "assemble.py").exists():
        raise SystemExit("в %s нет src/assemble.py" % p)
    return src


def docker(*args, stream=None):
    code, out = _run(["docker", *args], stream)
    if code:
        raise SystemExit("docker %s: код %d\n%s" % (" ".join(args[:3]), code, out.strip()[-600:]))
    return out


def prepare_container(label, src, corpus_lines, with_projects):
    """Только /tmp/<метка> контейнера: src, пустой кэш, список корпуса, сборщик. Контейнер должен уже работать."""
    code, out = _run(["docker", "inspect", "-f", "{{.State.Running}}", CONTAINER])
    if code or out.strip() != "true":
        raise SystemExit("контейнер %s не запущен (%s); запустите его сами — скрипт его не перезапускает" % (CONTAINER, out.strip()[:80]))
    base = "/tmp/" + label
    docker("exec", CONTAINER, "rm", "-rf", base)
    docker("exec", CONTAINER, "mkdir", "-p", base + "/cache")
    docker("cp", str(src), "%s:%s/src" % (CONTAINER, base))
    docker("exec", CONTAINER, "find", base + "/src", "-name", "__pycache__", "-prune", "-exec", "rm", "-rf", "{}", "+")
    if with_projects:
        docker("exec", CONTAINER, "ln", "-s", "/app/projects", base + "/projects")
    with tempfile.TemporaryDirectory() as td:
        lst = Path(td) / "corpus.txt"
        lst.write_text("\n".join(corpus_lines) + "\n", encoding="utf-8")
        docker("cp", str(lst), "%s:%s/corpus.txt" % (CONTAINER, base))
    docker("cp", str(REPO / "tools" / "gate_asm_all.py"), "%s:%s/gate_asm_all.py" % (CONTAINER, base))
    return base


def run_in_container(base, say=print):
    """Собрать корпус одним процессом Python в контейнере; возвращает записи GATE по документам."""
    records = []

    def on_line(line):
        if line.startswith("GATE\t"):
            rec = json.loads(line[5:])
            records.append(rec)
            say("  %-52s %5.0f с  листов %-3s  PDF: %s%s" % (rec["stem"][:52], rec["sec"], rec["pdf_pages"],
                                                              pdf_kind(rec)[1], "  ОШИБКА " + rec["err"] if rec["err"] else ""))
        elif line.strip():
            say("    " + line[:200])
    code, _ = _run(["docker", "exec", "-e", "PYTHONUNBUFFERED=1", "-w", base, CONTAINER, "python", base + "/gate_asm_all.py",
                    base + "/src", base + "/out", base + "/corpus.txt"], on_line)
    return code, records


def pdf_kind(rec):
    """(категория, подпись): lo — через DOCX и LibreOffice; place — PDF «на месте» (info без pdf); other — запасной путь; fail."""
    if rec.get("err"):
        return "fail", "—"
    f = rec.get("pdf_from")
    if f == LO_FROM:
        return "lo", "LibreOffice"
    if f is None:
        return "place", "на месте"
    return "other", "%s (%s)" % (f, (rec.get("pdf_reason") or "")[:80])


def summarize_run(records, expected):
    """Строки итога сборки и признак пригодности (False — есть упавшие, запасной путь PDF или не все документы)."""
    kinds = Counter(pdf_kind(r)[0] for r in records)
    lines = ["Собрано документов: %d из %d; через DOCX (LibreOffice) — %d, PDF на месте — %d" %
             (len(records), expected, kinds["lo"], kinds["place"])]
    ok = len(records) == expected
    if len(records) != expected:
        lines.append("⚠ собрано не всё: %d из %d" % (len(records), expected))
    for r in records:
        k, label = pdf_kind(r)
        if k == "fail":
            lines.append("⚠ упала сборка %s: %s" % (r["stem"], r["err"]))
        elif k == "other":
            lines.append("⚠ PDF %s не через LibreOffice: %s — сборка непригодна для сравнения" % (r["stem"], label))
    return lines, ok and not (kinds["fail"] or kinds["other"])


# ---------------------------------------------------------------- ворота
class Tee:
    def __init__(self, *streams):
        self.streams = streams

    def write(self, s):
        for st in self.streams:
            st.write(s)
        return len(s)

    def flush(self):
        for st in self.streams:
            st.flush()


def run_audit(out_dir, root, orig_dir):
    """tools/audit_out.py --fail-on never --json в том же процессе: у него REPO — корень с data/ (оригиналы портфолио)."""
    import audit_out as A
    A.REPO = Path(root)
    log = Path(out_dir) / "_audit.txt"
    with open(log, "w", encoding="utf-8") as f, contextlib.redirect_stdout(Tee(sys.stdout, f)):
        A.main([str(Path(out_dir) / "out"), "--orig", str(orig_dir), "--fail-on", "never", "--json", str(Path(out_dir) / "_audit.json")])
    return json.loads((Path(out_dir) / "_audit.json").read_text(encoding="utf-8"))


# ---------------------------------------------------------------- сравнение
HEAD = re.compile(r"^══ (.+?)\s+\(перевод: (\d+|—) л\., оригинал: (\d+|—) л\.", re.M)


class Doc:
    def __init__(self, stem, red=0, yellow=0, pages=None, orig=None, by_check=None):
        self.stem, self.red, self.yellow, self.pages, self.orig = stem, red, yellow, pages, orig
        self.by_check = by_check or Counter()          # (проверка, уровень) → число замечаний


def locate_build(arg):
    """Папка сборки по аргументу: `база`, метка из work/gate или путь к папке/`_audit.json`."""
    if arg in ("база", "base"):
        return base_dir()
    p = Path(arg)
    if p.is_file():
        return p.parent
    if p.is_dir():
        return p
    if (gate_dir() / arg).is_dir():
        return gate_dir() / arg
    raise SystemExit("нет сборки «%s»: ни пути, ни метки в %s" % (arg, gate_dir()))


def load_build(arg):
    """{имя документа: Doc} из `_audit.json` сборки; листов — из PDF рядом (рекурсивно), листов оригинала — из `_audit.txt`."""
    d = locate_build(arg)
    aj = d / "_audit.json"
    if not aj.exists():
        raise SystemExit("в %s нет _audit.json — сначала `build`, либо audit_out --json" % d)
    docs = {}
    for stem, fs in json.loads(aj.read_text(encoding="utf-8")).items():
        c = Counter((f["check"], f["level"]) for f in fs)
        docs[stem] = Doc(stem, sum(n for (_, lv), n in c.items() if lv == "red"), sum(n for (_, lv), n in c.items() if lv == "yellow"),
                         by_check=c)
    import pymupdf
    pymupdf.TOOLS.mupdf_display_errors(False)
    for p in d.rglob("*.pdf"):
        stem = re.sub(r"_(RU|EN|HU)$", "", p.stem)
        if stem in docs:
            try:
                with pymupdf.open(str(p)) as pdf:
                    docs[stem].pages = pdf.page_count
            except Exception:
                pass
    txt = d / "_audit.txt"
    if txt.exists():
        for stem, _, o in HEAD.findall(txt.read_text(encoding="utf-8")):
            if stem in docs and o.isdigit():
                docs[stem].orig = int(o)
    return docs


def excess(doc, ref):
    """Листов сверх оригинала (без оригинала — все листы): мерило «лишних листов»."""
    return None if doc.pages is None else max(0, doc.pages - (ref or 0))


def judge(a, b):
    """(вердикт, что именно) для пары Doc до/после: 'хуже' | 'лучше' | 'так же'."""
    ref = a.orig if a.orig is not None else b.orig
    worse, better = [], []
    if b.red != a.red:
        (worse if b.red > a.red else better).append("%s %+d" % (RED, b.red - a.red))
    if b.red == a.red and b.yellow != a.yellow:
        (worse if b.yellow > a.yellow else better).append("%s %+d" % (YELLOW, b.yellow - a.yellow))
    ea, eb = excess(a, ref), excess(b, ref)
    if ea is not None and eb is not None and ea != eb:
        (worse if eb > ea else better).append("листов %+d" % (eb - ea))
    if worse:
        return "хуже", ", ".join(worse)
    if better:
        return "лучше", ", ".join(better)
    return "так же", ""


def compare(A, B):
    """Строки сравнения: (имя, Doc до | None, Doc после | None, вердикт, пояснение); нет в одной из сборок — тоже «хуже»/«новый»."""
    rows = []
    for stem in sorted(set(A) | set(B)):
        a, b = A.get(stem), B.get(stem)
        if a is None:
            rows.append((stem, None, b, "новый", "нет в A"))
        elif b is None:
            rows.append((stem, a, None, "хуже", "нет в B"))
        else:
            rows.append((stem, a, b, *judge(a, b)))
    return rows


def _pair(x, y, fmt="%d"):
    return "%s → %s" % (fmt % x if x is not None else "—", fmt % y if y is not None else "—")


def render_compare(rows, name_a, name_b, show_all=False, width=46):
    out = ["Сравнение: A = %s  →  B = %s" % (name_a, name_b), ""]
    order = {"хуже": 0, "новый": 1, "лучше": 2, "так же": 3}
    shown = [r for r in rows if show_all or r[3] != "так же"]
    shown.sort(key=lambda r: (order[r[3]], r[0]))
    out.append("%-*s  %-10s %-10s %-12s  %s" % (width, "документ", RED + " до→после", YELLOW + " до→после", "листов (ориг.)", "вердикт"))
    for stem, a, b, v, why in shown:
        ref = (a or b).orig if (a or b) else None
        pg = _pair(a.pages if a else None, b.pages if b else None) + (" (%d)" % ref if ref else "")
        mark = {"хуже": "ХУЖЕ", "лучше": "лучше", "новый": "новый", "так же": ""}[v]
        out.append("%-*s  %-10s %-10s %-14s %s%s" % (width, stem[:width], _pair(a and a.red, b and b.red), _pair(a and a.yellow, b and b.yellow),
                                                    pg, mark, " — " + why if why else ""))
    if not shown:
        out.append("  (ни один документ не изменился)")
    cnt = Counter(r[3] for r in rows)
    ta = [r[1] for r in rows if r[1]]
    tb = [r[2] for r in rows if r[2]]
    s = lambda ds, k: sum(getattr(d, k) for d in ds)
    out += ["", "ИТОГО: документов A %d, B %d; хуже %d, лучше %d, так же %d%s" %
            (len(ta), len(tb), cnt["хуже"], cnt["лучше"], cnt["так же"], ", новых %d" % cnt["новый"] if cnt["новый"] else ""),
            "  %s %d → %d (%+d)   %s %d → %d (%+d)" % (RED, s(ta, "red"), s(tb, "red"), s(tb, "red") - s(ta, "red"),
                                                       YELLOW, s(ta, "yellow"), s(tb, "yellow"), s(tb, "yellow") - s(ta, "yellow"))]
    pa, pb = [d.pages for d in ta if d.pages is not None], [d.pages for d in tb if d.pages is not None]
    if pa and pb:
        out.append("  листов в переводах: %d → %d (%+d)" % (sum(pa), sum(pb), sum(pb) - sum(pa)))
    ca, cb = Counter(), Counter()
    for d in ta:
        ca.update(d.by_check)
    for d in tb:
        cb.update(d.by_check)
    diff = sorted(k for k in set(ca) | set(cb) if ca[k] != cb[k])
    if diff:
        out.append("  по проверкам:")
        for chk, lv in diff:
            out.append("    %s %-24s %d → %d" % (RED if lv == "red" else YELLOW, chk, ca[(chk, lv)], cb[(chk, lv)]))
    out.append("  " + ("✔ ничего не хуже" if not cnt["хуже"] else "✘ хуже документов: %d" % cnt["хуже"]))
    return "\n".join(out)


# ---------------------------------------------------------------- команды
def cmd_build(a):
    label = check_label(a.label)
    root = data_root()
    src = resolve_src(a.src or REPO)
    out_dir = Path(a.out) if a.out else gate_dir() / label
    files = corpus(root, a.only)
    if not files:
        raise SystemExit("корпус пуст: нет %s в %s" % (JOBS_GLOB, root / "data" / "jobs"))
    if not a.only and len(files) != EXPECTED_DOCS:
        print("⚠ в корпусе %d документов, ожидалось %d" % (len(files), EXPECTED_DOCS))
    print("Метка %s: код %s, корпус %d документов, результат %s" % (label, src, len(files), out_dir))
    ok = True
    if not a.audit_only:
        base = prepare_container(label, src, [container_path(root, f) for f in files], a.with_projects)
        code, records = run_in_container(base)
        if out_dir.exists() and (out_dir / "out").exists():
            shutil.rmtree(out_dir / "out")
        (out_dir / "json").mkdir(parents=True, exist_ok=True)
        for f in files:
            shutil.copy2(f, out_dir / "json" / f.name)
        (out_dir / "out").mkdir(parents=True, exist_ok=True)
        docker("cp", "%s:%s/out/." % (CONTAINER, base), str(out_dir / "out"))
        (out_dir / "_run.json").write_text(json.dumps(records, ensure_ascii=False, indent=1), encoding="utf-8")
        lines, ok = summarize_run(records, len(files))
        print("\n".join(lines))
        ok = ok and code == 0
        if a.clean:
            docker("exec", CONTAINER, "rm", "-rf", base)
    if not a.no_audit:
        res = run_audit(out_dir, root, a.orig)
        nr = sum(f["level"] == "red" for fs in res.values() for f in fs)
        ny = sum(f["level"] == "yellow" for fs in res.values() for f in fs)
        print("\nГОТОВО %s: %s %d  %s %d  (%s)" % (label, RED, nr, YELLOW, ny, out_dir))
        print("Сравнить с базой: python -X utf8 tools/gate_baseline.py compare база %s" % label)
    return 0 if ok else 1


def cmd_compare(a):
    A, B = load_build(a.a), load_build(a.b)
    if a.only:
        keep = lambda D: {k: v for k, v in D.items() if any(o.lower() in k.lower() for o in a.only)}
        A, B = keep(A), keep(B)
    rows = compare(A, B)
    print(render_compare(rows, a.a, a.b, a.all))
    return 1 if any(r[3] == "хуже" for r in rows) else 0


def main(argv=None):
    ap = argparse.ArgumentParser(description="Мерило «до/после» по корпусу (см. docstring модуля)")
    sub = ap.add_subparsers(dest="cmd", required=True)
    b = sub.add_parser("build", help="собрать корпус кодом папки и прогнать ворота")
    b.add_argument("label", help="метка сборки (имя папки work/gate/<метка> и /tmp/<метка> в контейнере)")
    b.add_argument("--src", help="worktree или папка src с кодом (по умолчанию — этот репозиторий)")
    b.add_argument("--out", help="куда класть результат (по умолчанию work/gate/<метка>)")
    b.add_argument("--orig", default=str(ORIG_DIR), help="папка с оригиналами")
    b.add_argument("--only", action="append", default=[], help="только документы с этой частью имени (пробная сборка)")
    b.add_argument("--with-projects", action="store_true", help="дать коду /app/projects контейнера (словарь проекта)")
    b.add_argument("--no-audit", action="store_true", help="только собрать, ворота не гонять")
    b.add_argument("--audit-only", action="store_true", help="не собирать: ворота по уже лежащей сборке")
    b.add_argument("--clean", action="store_true", help="удалить /tmp/<метка> в контейнере после копирования результата")
    c = sub.add_parser("compare", help="таблица до/после двух сборок")
    c.add_argument("a", help="сборка «до»: папка, _audit.json, метка или `база`")
    c.add_argument("b", help="сборка «после»")
    c.add_argument("--all", action="store_true", help="показать и неизменившиеся документы")
    c.add_argument("--only", action="append", default=[], help="только документы с этой частью имени (после пробной сборки --only)")
    a = ap.parse_args(argv)
    try:
        sys.stdout.reconfigure(encoding="utf-8")
    except Exception:
        pass
    return cmd_build(a) if a.cmd == "build" else cmd_compare(a)


if __name__ == "__main__":
    sys.exit(main())
