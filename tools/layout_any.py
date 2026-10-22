# -*- coding: utf-8 -*-
"""Стенд вёрстки, не привязанный к языку: документ собирается с «поддельным переводом» заданной длины — без портала,
без глоссария и без смысла. Текст каждого сегмента переводится в другое письмо (латиница ↔ кириллица) и растягивается
или сжимается по длине (×0,7 — перевод короче, ×1,0 — той же длины, ×1,4 — длиннее, как русский после венгерского).
Метки начертания <fN>, коды, числа и обозначения остаются. Так видно, что ломает сама вёрстка, а что — язык.

    python tools/layout_any.py build <случаи.tsv> <папка> [--factors 0.7,1.0,1.4] [--only метка,…]
    python tools/layout_any.py check <папка>          → <папка>/report.md, report.json, пары листов в pairs/

Случаи: TSV «метка ⇥ исходный файл ⇥ JSON задачи или -». JSON (рабочий файл конвейера: segs, info, struct) нужен
сканам и PDF без разметки — сегменты из него; без JSON размеченный PDF собирается по разметке, PDF с текстовым слоем —
по сегментам слоя (без OCR), DOCX — на месте. Печать DOCX — LibreOffice в контейнере hut (scripts/lo_print.sh).

Проверки (инварианты, одинаковые для любого языка):
  листы    — число листов сборки против оригинала: ×1,0 — не больше оригинала +10 % +1, ×0,7 — не больше оригинала +1,
             ×1,4 — не больше ×1,6 +1;
  потери   — слова оригинала (от 4 букв, не коды), которых нет в сборке в «переведённом» виде (×1,0 и ×1,4);
  остаток  — слова оригинала, оставшиеся непереведёнными (путь сборки мимо перевода: колонтитул, надпись, поле);
  мелкий   — доля текста кеглем меньше 60 % медианного кегля оригинала;
  картинки — растровых картинок в сборке меньше, чем в оригинале;
  пустые   — листы сборки без текста и без картинок;
  наложения — строки сборки, перекрытые другой строкой больше чем на четверть."""
import json
import os
import re
import subprocess
import sys
import time
import unicodedata
from collections import Counter
from pathlib import Path

ROOT = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(ROOT / "src"))
LO_PRINT = Path(os.environ.get("HU_LO_PRINT", "C:/project/hu_translate/work/check_2026-10-17/scripts/lo_print.sh"))
BASH = os.environ.get("HU_BASH", "C:/Program Files/Git/bin/bash.exe")

LAT = "abcdefghijklmnopqrstuvwxyz"
CYR = "абцдефгхийклмнопкрстуввхыз"          # латиница → кириллица (по звуку, без смысла)
TO_CYR = {a: c for a, c in zip(LAT, CYR)}
TO_LAT = {"а": "a", "б": "b", "в": "v", "г": "g", "д": "d", "е": "e", "ё": "e", "ж": "zh", "з": "z", "и": "i",
          "й": "j", "к": "k", "л": "l", "м": "m", "н": "n", "о": "o", "п": "p", "р": "r", "с": "s", "т": "t", "у": "u",
          "ф": "f", "х": "h", "ц": "c", "ч": "ch", "ш": "sh", "щ": "sch", "ъ": "", "ы": "y", "ь": "", "э": "e",
          "ю": "yu", "я": "ya"}
CODE = re.compile(r"\d|^[A-Z]{1,6}$|^[A-Z0-9]+[-_/.][A-Z0-9]|^[IVXLC]+\.?$|^[^\w]+$")
TAG = re.compile(r"(</?f\d+>|<su[bp]>|</su[bp]>)")


def _plain(ch):
    return unicodedata.normalize("NFKD", ch)[0]


def flip(word):
    """Слово — в другое письмо: латиница (с диакритикой) → кириллица, кириллица → латиница; регистр сохраняется."""
    out = []
    for ch in word:
        lo = ch.lower()
        p = _plain(lo)
        if p in TO_CYR:
            r = TO_CYR[p]
        elif lo in TO_LAT:
            r = TO_LAT[lo]
        else:
            r = ch
        out.append(r.upper() if ch.isupper() and len(r) == 1 else (r.capitalize() if ch.isupper() else r))
    return "".join(out)


def is_word(tok):
    core = tok.strip(".,;:!?()[]«»\"'„”“’-–—/")
    return len(core) >= 1 and not CODE.search(core) and any(ch.isalpha() for ch in core)


def fake_plain(text, factor):
    """Поддельный перевод куска без меток: слова — в другое письмо, длина — × factor (лишние слова — повтором
    соседних, недостающие — пропуском каждого n-го слова; коды, числа и знаки остаются)."""
    toks = re.split(r"(\s+)", text)
    words = [i for i, t in enumerate(toks) if t.strip() and is_word(t)]
    out = [flip(t) if i in set(words) else t for i, t in enumerate(toks)]
    if factor < 0.99 and len(words) >= 3:
        drop = round(len(words) * (1 - factor))
        step = len(words) / max(drop, 1)
        gone = {words[min(len(words) - 2, int(1 + k * step))] for k in range(drop)}
        out = [t for i, t in enumerate(out) if i not in gone]
        text2 = "".join(out)
        return re.sub(r"[ \t]{2,}", " ", text2)
    if factor > 1.01 and words:
        extra = round(len(words) * (factor - 1))
        if extra:
            ins = {}
            step = len(words) / extra
            for k in range(extra):
                i = words[min(len(words) - 1, int(k * step))]
                ins.setdefault(i, []).append(flip(toks[i]).strip(".,;:!?") or "ab")
            res = []
            for i, t in enumerate(out):
                res.append(t)
                for w in ins.get(i, []):
                    res.append(" " + w.lower())
            return "".join(res)
    return "".join(out)


def fake(text, factor):
    """Поддельный перевод с метками начертания и индексов: каждая часть между метками — fake_plain."""
    if not text:
        return text
    parts = TAG.split(text)
    return "".join(p if TAG.fullmatch(p) else fake_plain(p, factor) for p in parts)


def words_of(text):
    """Слова для сверки потерь: от 4 букв, без кодов; дефис переноса и мягкий перенос — склеены."""
    text = text.replace("\u00ad", "").replace("-\n", "")
    return [w.lower() for w in re.findall(r"[^\W\d_]{4,}", text)]


# ------------------------------------------------------------------------------------------------- сборка
def _no_portal():
    import llm

    def chat(*a, **k):
        raise RuntimeError("стенд layout_any: портал не вызывается")
    llm.chat = chat


def _fake_memory(factor):
    import docx_build as B

    class FakeMemory(B.Memory):
        def get(self, text):
            # «нужен ли перевод» конвейера смотрит на язык (русский текст — уже готов); стенду язык не важен:
            # переводится всё, где есть слово, кроме кодов и чисел
            plain = TAGS.sub("", text)
            if not any(is_word(t) for t in plain.split()):
                return text
            return fake(text, factor)
    TAGS = B.TAGS
    return FakeMemory([])


def lo_print(docx):
    """DOCX → PDF рядом (LibreOffice в контейнере hut). Печать идёт под уникальным именем: lo_print.sh кладёт файл в
    /tmp контейнера по имени, и два процесса стенда с одинаковыми «x1.0.docx» подменяли друг другу PDF."""
    import uuid
    docx = Path(docx)
    pdf = docx.with_suffix(".pdf")
    pdf.unlink(missing_ok=True)
    tmp = docx.with_name(f"lo_{uuid.uuid4().hex[:12]}.docx")
    tmp.write_bytes(docx.read_bytes())
    env = dict(os.environ, MSYS_NO_PATHCONV="1")
    try:
        subprocess.run([BASH, str(LO_PRINT).replace("\\", "/"), str(tmp).replace("\\", "/")], env=env,
                       capture_output=True, timeout=1800)
        if tmp.with_suffix(".pdf").exists():
            tmp.with_suffix(".pdf").replace(pdf)
    finally:
        tmp.unlink(missing_ok=True)
    return pdf if pdf.exists() else None


def build_case(label, src, jf, out, factors):
    import assemble as A
    import docx_build as B
    import segments as S
    src = Path(src)
    d = json.loads(Path(jf).read_text(encoding="utf-8")) if jf and jf != "-" else None
    segs = [s for s in (d or {}).get("segs", []) if not s["loc"].get("docx_only")] if d else None
    case = out / label
    case.mkdir(parents=True, exist_ok=True)
    info = {"label": label, "src": str(src), "json": jf, "runs": {}}
    if src.suffix.lower() in (".docx",):
        mode = "docx"
        segs = segs if segs is not None else S.docx_segments(src)
        orig_pdf = case / "orig.pdf"
        if not orig_pdf.exists():
            tmp = case / "orig.docx"
            tmp.write_bytes(src.read_bytes())
            lo_print(tmp)
        info["orig_pdf"] = str(orig_pdf)
    else:
        info["orig_pdf"] = str(src)
        st = (d or {}).get("struct")
        if d is not None:
            mode = (st or {}).get("mode") or ("struct" if st else None)
            if mode is None:
                if (d.get("info") or {}).get("ocr_pages"):
                    mode = "layout"
                else:
                    mode = "text"
        else:
            segs, sinfo = S.pdf_segments(src, ocr=False, shots=False)
            try:
                cov = B.coverage(str(src), segs)
            except Exception:                           # noqa: BLE001 — разметки нет
                cov = 0.0
            mode = "struct" if cov >= B.STRUCT_COVER else "text"
            if mode == "struct":
                segs = []
    info["mode"] = mode
    for f in factors:
        dst = case / f"x{f:.1f}.docx"
        t0 = time.time()
        try:
            mem = _fake_memory(f)
            B.job_memory = lambda segs_, res_, _m=mem: _m          # всё, что сборка спросит, — поддельным переводом
            res = {str(s["id"]): {"ru": fake(s["text"], f)} for s in (segs or [])}
            if mode == "docx":
                r = A.build_docx(src, dst, segs, res)
                import docx_fit
                try:
                    r["fit"] = docx_fit.fit(src, dst, lambda a, b: (lo_print(a), Path(a).with_suffix(".pdf").replace(b)))
                except Exception as e:                                  # noqa: BLE001 — подгонка не валит стенд
                    r["fit"] = f"{type(e).__name__}: {e}"[:200]
            else:
                r, _ = A.build_struct_docx(src, dst, segs or [], res, None, mode=mode)
            pdf = lo_print(dst)
            info["runs"][f"{f:.1f}"] = {"ok": pdf is not None, "sec": round(time.time() - t0), "build": str(r)[:300]}
        except Exception as e:                                          # noqa: BLE001 — случай падает, стенд идёт
            import traceback
            info["runs"][f"{f:.1f}"] = {"ok": False, "error": f"{type(e).__name__}: {e}"[:300],
                                        "trace": traceback.format_exc()[-1500:]}
        print(label, f, info["runs"][f"{f:.1f}"].get("ok"), info["runs"][f"{f:.1f}"].get("error", ""), flush=True)
    (case / "case.json").write_text(json.dumps(info, ensure_ascii=False, indent=1), encoding="utf-8")


def cmd_build(cases_tsv, out, factors, only=None):
    _no_portal()
    out = Path(out)
    out.mkdir(parents=True, exist_ok=True)
    for line in Path(cases_tsv).read_text(encoding="utf-8").splitlines():
        if not line.strip() or line.startswith("#"):
            continue
        label, src, jf = (line.split("\t") + ["-"])[:3]
        if only and label not in only:
            continue
        build_case(label, src, jf, out, factors)


# ------------------------------------------------------------------------------------------------- проверки
def _spans(page):
    out = []
    for b in page.get_text("dict")["blocks"]:
        for ln in b.get("lines", []):
            for sp in ln["spans"]:
                if sp["text"].strip():
                    out.append((sp["size"], sp["text"], ln["bbox"]))
    return out


def _overlaps(page):
    import fitz
    lines = []
    for b in page.get_text("dict")["blocks"]:
        for ln in b.get("lines", []):
            if "".join(s["text"] for s in ln["spans"]).strip():
                lines.append(fitz.Rect(ln["bbox"]))
    n = 0
    for i in range(len(lines)):
        for j in range(i + 1, len(lines)):
            a, b = lines[i], lines[j]
            inter = a & b
            if not inter.is_empty and inter.get_area() > 0.25 * min(a.get_area(), b.get_area()) > 0:
                n += 1
    return n


def _images(doc):
    n = 0
    for pg in doc:
        for im in pg.get_image_info():
            r = im["bbox"]
            if (r[2] - r[0]) * (r[3] - r[1]) > 400:
                n += 1
    return n


def check_case(case):
    import fitz
    info = json.loads((case / "case.json").read_text(encoding="utf-8"))
    o = fitz.open(info["orig_pdf"])
    otext = "\n".join(p.get_text() for p in o)
    owords = [w for w in words_of(otext)]
    o_sizes = [s for p in o for s, _, _ in _spans(p)]
    med = sorted(o_sizes)[len(o_sizes) // 2] if o_sizes else 10
    oimg = _images(o)
    expect = Counter(flip(w).lower() for w in owords)
    orig_lat = Counter(w for w in owords if re.search(r"[a-z]", _plain_word(w)))
    rows = {}
    for f, run in info["runs"].items():
        pdf = case / f"x{f}.pdf"
        if not run.get("ok") or not pdf.exists():
            rows[f] = {"error": run.get("error", "нет PDF")}
            continue
        b = fitz.open(pdf)
        btext = "\n".join(p.get_text() for p in b)
        bw = Counter(words_of(btext))
        lost = sum(1 for w in expect if w not in bw) / max(len(expect), 1) if float(f) >= 1.0 else None
        resid = sum(1 for w in orig_lat if w in bw and flip(w).lower() != w) / max(len(orig_lat), 1)
        sizes = [s for p in b for s, _, _ in _spans(p)]
        tiny = sum(1 for s in sizes if s < 0.6 * med) / max(len(sizes), 1)
        empty = [i + 1 for i, p in enumerate(b) if len(p.get_text().strip()) < 20 and not p.get_image_info()]
        ov = {i + 1: n for i, p in enumerate(b) if (n := _overlaps(p))}
        rows[f] = {"pages": len(b), "lost": lost, "resid": resid, "tiny": tiny, "imgs": _images(b), "empty": empty,
                   "overlaps": ov}
    return {"label": info["label"], "mode": info.get("mode"), "orig_pages": len(o), "orig_imgs": oimg, "runs": rows}


def _plain_word(w):
    return "".join(_plain(c) for c in w)


def flags(r):
    """Нарушения инвариантов по одному случаю → [(множитель, текст)]."""
    out = []
    o = r["orig_pages"]
    for f, x in r["runs"].items():
        if "error" in x:
            out.append((f, "сборка упала: " + x["error"][:120]))
            continue
        fv, p = float(f), x["pages"]
        lim = o + 1 if fv < 1 else (o * 1.1 + 1 if fv < 1.2 else o * 1.6 + 1)
        if p > lim:
            out.append((f, f"листов {p} при оригинале {o}"))
        if fv >= 1 and p < 0.75 * o - 1:
            out.append((f, f"листов {p} — меньше оригинала {o} (пропало?)"))
        if x["lost"] is not None and x["lost"] > 0.03:
            out.append((f, f"потеряно слов {x['lost']:.0%}"))
        if x["resid"] > 0.03:
            out.append((f, f"не переведено слов {x['resid']:.0%}"))
        if x["tiny"] > 0.02:
            out.append((f, f"мелкий текст {x['tiny']:.0%}"))
        if x["imgs"] < r["orig_imgs"]:
            out.append((f, f"картинок {x['imgs']} из {r['orig_imgs']}"))
        if x["empty"]:
            out.append((f, f"пустые листы {x['empty'][:8]}"))
        if x["overlaps"]:
            worst = sorted(x["overlaps"].items(), key=lambda kv: -kv[1])[:4]
            out.append((f, "наложения строк: " + ", ".join(f"л.{k}×{v}" for k, v in worst)))
    return out


def cmd_check(out):
    out = Path(out)
    res = []
    for case in sorted(p for p in out.iterdir() if (p / "case.json").exists()):
        try:
            res.append(check_case(case))
        except Exception as e:                                          # noqa: BLE001
            res.append({"label": case.name, "error": f"{type(e).__name__}: {e}"[:200], "runs": {}, "orig_pages": 0,
                        "orig_imgs": 0})
    (out / "report.json").write_text(json.dumps(res, ensure_ascii=False, indent=1), encoding="utf-8")
    lines = ["# Стенд вёрстки без языка — отчёт", "",
             "| случай | путь | листов: ориг. → ×0,7 / ×1,0 / ×1,4 | нарушения |", "|---|---|---|---|"]
    bad = 0
    for r in res:
        pages = " / ".join(str(r["runs"].get(f, {}).get("pages", "—")) for f in ("0.7", "1.0", "1.4"))
        fl = flags(r) if "error" not in r else [("", r["error"])]
        bad += bool(fl)
        lines.append(f"| {r['label']} | {r.get('mode', '')} | {r['orig_pages']} → {pages} | "
                     + ("; ".join(f"×{f}: {t}" for f, t in fl) or "—") + " |")
    lines.insert(2, f"Случаев: {len(res)}, с нарушениями: {bad}.\n")
    (out / "report.md").write_text("\n".join(lines) + "\n", encoding="utf-8")
    print("\n".join(lines))


if __name__ == "__main__":
    a = sys.argv[1:]
    if a and a[0] == "build":
        fs = [float(x) for x in (a[a.index("--factors") + 1] if "--factors" in a else "0.7,1.0,1.4").split(",")]
        only = set(a[a.index("--only") + 1].split(",")) if "--only" in a else None
        cmd_build(a[1], a[2], fs, only)
    elif a and a[0] == "check":
        cmd_check(a[1])
    else:
        print(__doc__)
