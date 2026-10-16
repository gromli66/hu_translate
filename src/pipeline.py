# -*- coding: utf-8 -*-
"""Конвейер перевода комплекта HU→RU на портале (схема, проверенная слепыми вычитками 01–06.10.2026).

Шаги (каждый можно запускать отдельно, состояние — в рабочей папке):
  extract   документы (DOCX/PDF) → сегменты с разметкой; PDF: текстовый слой или OCR (Tesseract hun + портал по картинке)
  translate перевод разделами ≤6000 знаков, контекст — свой перевод предыдущего раздела; общая память комплекта
            (одна фраза — один перевод) и шаблоны «то же до чисел»; проверки кодом, повтор с замечаниями; постобработка
            размеченный PDF (экспорт Word): абзацы разметки без пары в сегментах — своими сегментами (struct_fill)
  review    смысловая проверка агентом портала (насыщенный режим) + исправление + повторная проверка исправленного
  polish    шлифовка по языковым соглашениям комплекта и словарю документа: код, затем портал, не больше 3 проходов
  assemble  DOCX/PDF «на месте» (размеченный PDF — DOCX по разметке и профилю оформления оригинала) + таблица
            вычитки (G — правка редактора, H — смысловая проверка; листы «шлифовка» и «нарушения»)
  edits     правки редактора из колонки G → перевод, память переводов проекта, размножение на похожие фразы (RAG)

Рабочая папка: <work>/ocr (кэш OCR), <work>/json/<док>.json (состояние), <work>/out (результаты).
Ответы портала кэшируются (cache/), перезапуск после обрыва дешёвый."""
import re, json, time, difflib, threading, collections
from pathlib import Path
from concurrent.futures import ThreadPoolExecutor

import glossary as G
import segments as S
from translate import Translator, TAG, NUM, HU_CAPS_WORDS, LAT_ABBR, fix_caps, fix_units, gloss_misses, badness
from llm import STAT

SEC_CHARS, CTX_CHARS, STREAM_CHARS = 6000, 3000, 250000
cl = lambda t: TAG.sub("", t or "").strip()
# ключ памяти: схлопываются только пробелы; табуляции и переносы — часть разметки (иначе чужая раскладка)
norm = lambda t: re.sub(r" +", " ", t).strip(" ")
tmkey = lambda t: re.sub(r"\s+", " ", cl(t)).strip()           # ключ памяти переводов — как в apply_review.norm_hu


def log(msg):
    print(time.strftime("%H:%M:%S"), msg, flush=True)


def _quiet(stage, done, total):
    """Обратный вызов прогресса по умолчанию; веб-сервис передаёт свой: progress(шаг, сделано, всего)."""


# ------------------------------------------------------------------ проект и состояние
def load_project(path):
    """project.json: {"glossary": основной глоссарий заказчика (md), "terms": глоссарий комплекта (md, необяз.),
    "conventions": свод принятых соответствий (md, необяз.), "domain": первая фраза системного промпта (необяз.),
    "tm": память переводов (json, необяз.; создаётся правками редактора), "rules": языковые соглашения, машинно
    проверяемые (json, необяз.; src/conventions.py), "doc_terms": папка словарей документов <документ>_terms.json
    (необяз.; шлифовка)}. Пути — относительно project.json."""
    p = Path(path).resolve()
    prj = json.loads(p.read_text(encoding="utf-8"))
    for k in ("glossary", "terms", "conventions", "tm", "rules", "doc_terms"):
        if prj.get(k):
            prj[k] = str((p.parent / prj[k]).resolve())
    prj.setdefault("tm", str(p.parent / "tm.json"))
    if not prj.get("glossary"):
        raise ValueError(f"{p}: не задан glossary")
    return prj


def make_translator(prj, think=True):
    return Translator(think=think, extra=prj.get("terms"), glossary=prj["glossary"], domain=prj.get("domain"),
                      tm=prj.get("tm"))


def doc_files(work):
    return sorted(f for f in (Path(work) / "json").glob("*.json") if not f.name.endswith(".before_review.json"))


def load_docs(work):
    return {f.stem: json.loads(f.read_text(encoding="utf-8")) for f in doc_files(work)}


def save_doc(work, stem, d):
    (Path(work) / "json" / f"{stem}.json").write_text(json.dumps(d, ensure_ascii=False), encoding="utf-8")


# путь, имя файла, адрес — не текст: «/etc/uts/», «uts.ini», «uts_<éééé>.tar.gz» вносили «uts» в словарь документа,
# и сокращение UTS переводилось как венгерское слово прописными (UTS, сервис 15.10: 12 раскрытий)
# (слова через косую — «Aláírás/Dátum», «kg/óra» — текст)
CODE_TOKEN = re.compile(r"\S*(?:[_\\<>@]|://)\S*|(?<!\w)/\S+|\S+\.[a-z]{2,4}(?!\w)")


def set_vocab(tr, segs):
    low, caps = collections.Counter(), collections.Counter()
    for s in segs:
        t = CODE_TOKEN.sub(" ", TAG.sub(" ", s["text"]))
        low.update(re.findall(r"[a-záéíóöőúüű]{2,}", t))
        caps.update(w.lower() for w in LAT_ABBR.findall(t))
    # сокращение, раз-другой написанное строчными («az uts naplófájl» при 549 «UTS»), — не слово документа;
    # венгерское слово прописными (KEZELÉSI, NAGYOBB, LETT) встречается строчными чаще или с диакритикой
    words = {w for w, n in low.items() if not (n <= 2 and caps[w] >= 10 * n and re.fullmatch(r"[a-z]+", w))}
    tr.vocab = frozenset(words |
                         {w.lower() for e in tr.entries for w in re.findall(r"[A-Za-záéíóöőúüűÁÉÍÓÖŐÚÜŰ]{2,}", e.hu)
                          if not G._is_abbr(w)} | HU_CAPS_WORDS)


# ------------------------------------------------------------------ 1. извлечение
DOC_SUFFIXES = (".docx", ".pdf")


def unpack_portfolio(f, dest):
    """PDF-портфолио (в каталоге есть /Collection): видимая страница — заглушка Adobe «откройте в Acrobat X»,
    сами документы лежат вложениями. Вложенные DOCX/PDF распаковываются в dest и переводятся как отдельные
    документы; для обычного PDF — None. (30RDGT.PDF, 06.10: переводилась только заглушка из двух фраз.)"""
    import fitz
    d = fitz.open(f)
    if d.xref_get_key(d.pdf_catalog(), "Collection")[0] == "null" or not d.embfile_count():
        return None
    dest.mkdir(parents=True, exist_ok=True)
    out, emb = [], {}
    for i in range(d.embfile_count()):
        name = Path(d.embfile_info(i).get("filename") or f"{f.stem}_{i}").name
        if Path(name).suffix.lower() not in DOC_SUFFIXES:
            continue
        if not name.startswith(f.stem):                 # имя документа в комплекте должно быть уникальным
            name = f"{f.stem} {name}"
        p = dest / name
        p.write_bytes(d.embfile_get(i))
        out.append(p)
        emb[i] = name
    # опись: из чего собирать переведённое портфолио той же структуры (rebuild_portfolios)
    (dest / "_portfolio.json").write_text(json.dumps({"source": str(Path(f).resolve()), "map": emb}, ensure_ascii=False),
                                          encoding="utf-8")
    return out


def rebuild_portfolios(work, out):
    """Переведённое портфолио: копия исходного, где каждое вложение заменено своим переводом (PDF — PDF «на месте»,
    DOCX — DOCX). Заказчик получает один файл той же структуры, что прислал; переводы по отдельности тоже остаются."""
    import fitz
    for man in sorted((Path(work) / "unpacked").glob("*/_portfolio.json")):
        m = json.loads(man.read_text(encoding="utf-8"))
        d = fitz.open(m["source"])
        keys = d.embfile_names()                 # по ключам, не по номерам: удаление сдвигает номера
        done = 0
        for idx, name in m["map"].items():
            ru = Path(out) / f"{Path(name).stem}_RU{Path(name).suffix.lower()}"
            if ru.exists():
                key = keys[int(idx)]
                info = d.embfile_info(key)
                # embfile_upd в PyMuPDF 1.27 падает на bytes («no attribute m_internal») — заменяем удалением и добавлением
                d.embfile_del(key)
                d.embfile_add(key, ru.read_bytes(), filename=info.get("filename"), ufilename=info.get("ufilename"),
                              desc=info.get("desc"))
                done += 1
        target = Path(out) / f"{Path(m['source']).stem}_RU.pdf"
        d.save(target, garbage=3, deflate=True)
        log(f"сборка: {target.name} — портфолио, переведённых вложений {done} из {len(m['map'])}")


def extract(inputs, work, progress=_quiet):
    work = Path(work); (work / "json").mkdir(parents=True, exist_ok=True)
    S.OCR_DIR = work / "ocr"
    queue = []
    for x in inputs:
        x = Path(x)
        queue += sorted(p for p in (x.iterdir() if x.is_dir() else [x]) if p.suffix.lower() in DOC_SUFFIXES)
    files = []
    while queue:                                         # портфолио внутри портфолио тоже раскрывается
        f = queue.pop(0)
        inner = unpack_portfolio(f, work / "unpacked" / f.stem) if f.suffix.lower() == ".pdf" else None
        if inner is None:
            files.append(f)
        else:
            log(f"извлечение: {f.name} — PDF-портфолио, документов внутри: {len(inner)} "
                f"({', '.join(p.name for p in inner)}); страница-заглушка пропущена")
            queue = inner + queue
    for i, f in enumerate(files):
        progress("извлечение", i, len(files))
        jf = work / "json" / f"{f.stem}.json"
        if jf.exists():
            log(f"извлечение: {f.name} — уже есть"); continue
        t0 = time.time()
        if f.suffix.lower() == ".docx":
            segs, info = S.docx_segments(f), {}
        else:
            segs, info = S.pdf_segments(f, progress=lambda x: progress("извлечение", i + x, len(files)))
        dplan = struct = None
        ocr = info.get("ocr_pages") or []
        if f.suffix.lower() == ".pdf" and ocr:
            import fitz
            with fitz.open(f) as pdf_:
                few_ocr = len(ocr) < 0.5 * pdf_.page_count
        else:
            few_ocr = True
        if f.suffix.lower() == ".pdf" and few_ocr:
            # размеченный PDF (экспорт Word): DOCX — по дереву разметки и профилю оформления оригинала
            # (struct_fill, assemble.build_struct_docx); неполная разметка потеряла бы текст — тогда как раньше.
            # Отдельные распознанные листы (логотип на титуле руководства OAH ушёл в OCR) путь по разметке не
            # отменяют: прежде один такой лист уводил весь документ в перевод поверх PDF (живая проверка, 16.10)
            import docx_build
            try:
                cov = docx_build.coverage(f, segs)
            except Exception as e:                      # noqa: BLE001 — разметку не прочли: DOCX прежним путём
                cov = 0.0
                log(f"извлечение: {f.name} — разметка PDF не прочитана ({type(e).__name__}: {e})"[:300])
            if cov >= docx_build.STRUCT_COVER:
                struct = {"coverage": round(cov, 3)}
                log(f"извлечение: {f.name} — размеченный PDF, разметка покрывает {cov:.1%} текста: DOCX по разметке")
            elif cov:
                log(f"извлечение: {f.name} — разметка покрывает только {cov:.1%} текста: DOCX не по разметке")
        if f.suffix.lower() == ".pdf" and info.get("ocr_pages") and __import__("docx_layout").suits(f, info["ocr_pages"]):
            # распознанный PDF (текст листа — контурами, не скан): DOCX по сегментам листа (docx_layout); абзацы с
            # выделенными словами — переводом с метками начертания (struct_fill), как у размеченного
            struct = {"mode": "layout"}
        if f.suffix.lower() == ".pdf" and not info.get("ocr_pages") and not struct:
            # PDF с текстовым слоем: DOCX — из оригинала, его абзацы без пары в PDF переводятся своими сегментами
            import docx_src
            try:
                extra, dplan = docx_src.plan(f, segs, work / "docx" / f"{f.stem}.docx", len(segs))
            except Exception as e:                      # noqa: BLE001 — DOCX тогда прежним путём, перевод не страдает
                extra, dplan = None, f"{type(e).__name__}: {e}"[:200]
            if extra is None:
                log(f"извлечение: {f.name} — DOCX из оригинала не собран ({dplan}); DOCX будет из переведённого PDF")
                dplan = None
            else:
                segs += extra
                info["docx_own"] = len(extra)
        jf.write_text(json.dumps({"file": str(f.resolve()), "segs": segs, "res": {}, "info": info, "docx": dplan,
                                  "struct": struct}, ensure_ascii=False), encoding="utf-8")
        log(f"извлечение: {f.name} — сегментов {len(segs)}, OCR-страниц {len(info.get('ocr_pages', []))}, {time.time() - t0:.0f} с")
    progress("извлечение", len(files), len(files))


# ------------------------------------------------------------------ 2. перевод разделами
def translate(work, prj, workers=4, redo=False, progress=_quiet):
    """Переводит документы рабочей папки без перевода (redo — все заново). Уже переведённые документы комплекта
    засевают общую память: добавленный позже документ получает те же переводы тех же фраз."""
    docs_all = load_docs(work)
    docs = {k: d for k, d in docs_all.items() if redo or not d.get("res")}
    if not docs:
        log("перевод: всё уже переведено (заново — --redo)"); return
    tr = make_translator(prj)
    set_vocab(tr, [s for d in docs_all.values() for s in d["segs"]])
    tm = {tmkey(k): v for k, v in (tr.tm or {}).items()}
    memory, templ, lock = {}, {}, threading.Lock()
    for stem, d in docs_all.items():
        if stem not in docs:
            for s in d["segs"]:
                r = d["res"].get(str(s["id"])) or {}
                if r.get("ru") and not r.get("copied"):
                    memory.setdefault(norm(s["text"]), {"ru": r["ru"], "issues": r.get("issues", []), "tries": 0})

    def tkey(text):
        nums = NUM.findall(text)
        if len(text) <= 120 and 1 <= len(nums) <= 4 and not re.search(r"\d\s+[a-záéíóöőúüű]", text):
            return NUM.sub("§", norm(text))
        return None

    def from_memory(text):
        k = norm(text)
        if tmkey(text) in tm:
            return {"ru": tm[tmkey(text)], "issues": [], "tries": 0, "tm": True}
        with lock:
            if k in memory:
                return dict(memory[k], memory=True)
            t = tkey(text)
            if t and t in templ:
                src, ru = templ[t]
                if NUM.findall(ru) == NUM.findall(src):
                    it = iter(NUM.findall(text))
                    return {"ru": NUM.sub(lambda _: next(it), ru), "issues": [], "tries": 0, "templated": True}
        return None

    def remember(text, r):
        if not r["ru"]:                      # обрыв портала — пустой перевод не размножать по копиям
            return
        with lock:
            memory.setdefault(norm(text), {"ru": r["ru"], "issues": r["issues"], "tries": r.get("tries", 1)})
            t = tkey(text)
            if t and t not in templ and not r["issues"]:
                templ[t] = (text, r["ru"])

    streams = []
    for stem, d in docs.items():
        cur, size = [], 0
        for s in d["segs"]:
            cur.append(s); size += len(s["text"])
            if size >= STREAM_CHARS:
                streams.append((stem, cur)); cur, size = [], 0
        if cur:
            streams.append((stem, cur))
    streams.sort(key=lambda j: -sum(len(s["text"]) for s in j[1]))      # длинные первыми

    def sections(segs):
        secs, cur, size = [], [], 0
        for s in segs:
            if cur and size + len(s["text"]) > SEC_CHARS:
                secs.append(cur); cur, size = [], 0
            cur.append(s); size += len(s["text"])
        if cur:
            secs.append(cur)
        return secs

    streams = [(stem, sections(segs)) for stem, segs in streams]
    total = sum(len(secs) for _, secs in streams)
    done, cnt, t0 = collections.defaultdict(dict), {"sec": 0}, time.time()
    log(f"перевод: документов {len(docs)}, потоков {len(streams)}, разделов {total}, по {workers} одновременно")
    progress("перевод", 0, total)

    # одиночное латинское сокращение без записи в глоссарии («GUI», «CTM» в списке сокращений) — как есть: модель
    # раскрывала его в «графический пользовательский интерфейс», ячейка ужималась до 3 pt (UTS, 15.10); сокращения
    # глоссария (RVR → УСБ) переводятся
    # венгерское слово заглавными («HIBA», «VAGY», «NYITVA») — не сокращение: оно есть строчными в тексте
    # комплекта (tr.vocab) или в HU_CAPS_WORDS (красная команда 15.10: 90 сегментов «HIBA» остались бы без перевода)
    known = {e.hu.strip() for e in tr.entries}
    keep_abbr = lambda t: (bool(re.fullmatch(r"[A-Z][A-Z0-9]{1,6}", t.strip())) and t.strip() not in known
                           and t.strip().lower() not in tr.vocab and t.strip().lower() not in HU_CAPS_WORDS)

    def run_stream(job):
        stem, secs = job
        res, ctx = {}, ""
        for sec in secs:
            units, by_text = [], {}
            for s in sec:
                if not S.needs_translation(s["text"]) or keep_abbr(s["text"]):
                    res[s["id"]] = {"ru": s["text"], "issues": [], "tries": 0, "copied": True}; continue
                m = from_memory(s["text"])
                if m:
                    res[s["id"]] = m; continue
                if s["text"] not in by_text:
                    by_text[s["text"]] = len(units); units.append({"key": len(units), "text": s["text"]})
            if units:
                got = tr.run_batch({"units": units, "first_abbr": [], "context": ctx[-CTX_CHARS:]})
                for u in units:
                    remember(u["text"], got[u["key"]])
                for s in sec:
                    if s["id"] not in res:
                        res[s["id"]] = dict(got[by_text[s["text"]]])
            ctx = "\n".join(f"{cl(s['text'])} → {cl(res[s['id']]['ru'])}" for s in sec if not res[s["id"]].get("copied"))
            with lock:
                cnt["sec"] += 1
                progress("перевод", cnt["sec"], total)
                if cnt["sec"] % 20 == 0:
                    log(f"  разделов {cnt['sec']} за {time.time() - t0:.0f} с; запросов {STAT['calls']}")
        with lock:
            done[stem].update(res)

    with ThreadPoolExecutor(workers) as ex:
        list(ex.map(run_stream, streams))
    main = [e for e in tr.entries if id(e) in tr.main_ids]
    for stem, d in docs.items():
        out = {s["id"]: done[stem][s["id"]] for s in d["segs"]}
        # перевод из памяти переводов утверждён редактором — постобработка его не трогает
        tr.first_use(d["segs"], {k: v for k, v in out.items() if not v.get("tm")})
        for s in d["segs"]:
            r = out[s["id"]]
            if not r.get("copied") and not r.get("tm") and r["ru"]:
                r["ru"] = fix_units(fix_caps(s["text"], r["ru"]))
                r["issues"] = tr.full_check(s["text"], r["ru"])
                r["gloss_miss"] = gloss_misses(TAG.sub("", s["text"]), r["ru"], main, outer=tr.entries)
        d["res"] = {str(k): v for k, v in out.items()}
        d.pop("polish", None)                       # перевод новый — шлифовка заново
        save_doc(work, stem, d)
        log(f"перевод: {stem} — на вычитку {sum(1 for r in d['res'].values() if r.get('issues'))}")
    log(f"перевод готов за {time.time() - t0:.0f} с; {STAT}")


# ------------------------------------------------------------------ 2а. абзацы размеченного PDF без пары в сегментах
def struct_fill(work, prj, progress=_quiet):
    """Размеченный PDF (d["struct"], см. extract): тексты, которые сборка DOCX по разметке спросит у памяти, а среди
    сегментов страниц их нет (абзац со ссылкой Word из нескольких элементов дерева, «FIGYELEM!» жирным + текст, пункт
    с командой Consolas), — перевод порталом (сначала — память переводов проекта) своими сегментами (loc.struct_only,
    после текста своей страницы). Дальше они идут как все: проверка, шлифовка, таблица вычитки, правки редактора.
    Второй круг — то, что сборка спросит из-за первого: перевод с метками вернулся без меток — тот же абзац без них.
    Распознанный PDF (d["struct"]["mode"] == "layout", docx_layout) — так же: абзацы с полужирными и курсивными
    словами (по картинке листа) переводятся с метками начертания."""
    import docx_build as B
    files = doc_files(work)
    for i, f in enumerate(files):
        d = json.loads(f.read_text(encoding="utf-8"))
        st = d.get("struct")
        if not st or "filled" in st or not d.get("res"):
            continue
        progress("структура", i, len(files))
        t0, tr, added = time.time(), None, 0
        for _ in range(2):
            try:
                miss = B.missing(d["file"], d["segs"], d["res"], st.get("mode") or "struct")
            except Exception as e:                  # noqa: BLE001 — сборка по разметке тогда откатится на прежний путь
                st["error"] = f"{type(e).__name__}: {e}"[:200]
                log(f"структура: {f.stem} — пробная сборка DOCX по разметке не удалась ({st['error']})")
                break
            if not miss:
                break
            if tr is None:
                tr = make_translator(prj)
                set_vocab(tr, d["segs"])
                tm = {tmkey(k): v for k, v in (tr.tm or {}).items()}
                main = [e for e in tr.entries if id(e) in tr.main_ids]
            units = [t for t, _, _ in miss if tmkey(t) not in tm]
            got = {}
            for j in range(0, len(units), 12):
                chunk = units[j:j + 12]
                out = tr.run_batch({"units": [{"key": k, "text": t} for k, t in enumerate(chunk)], "first_abbr": [],
                                    "context": ""})
                got.update({t: out[k] for k, t in enumerate(chunk)})
                progress("структура", i + 0.9 * (j + len(chunk)) / len(units), len(files))
            for t, page, code in miss:
                if tmkey(t) in tm:
                    r = {"ru": tm[tmkey(t)], "issues": [], "tries": 0, "tm": True}
                else:
                    r = dict(got[t])
                    if r["ru"]:
                        # моноширинные части: команды — дословно, заполнитель <…> — в скобках (docx_build.code_part)
                        r["ru"] = B.keep_code(t, fix_units(fix_caps(t, r["ru"])), code)
                        r["issues"] = tr.full_check(t, r["ru"])
                        r["gloss_miss"] = gloss_misses(TAG.sub("", t), r["ru"], main, outer=tr.entries)
                sid = max(int(s["id"]) for s in d["segs"]) + 1
                pos = next((k for k, s in enumerate(d["segs"]) if page is not None and s["page"] is not None
                            and s["page"] > page), len(d["segs"]))
                d["segs"].insert(pos, {"id": sid, "text": t, "src": t, "page": page, "loc": {"struct_only": True}})
                d["res"][str(sid)] = r
                added += 1
        st["filled"] = added
        save_doc(work, f.stem, d)
        log(f"структура: {f.stem} — абзацев разметки без пары в сегментах {added}, {time.time() - t0:.0f} с")
    progress("структура", len(files), len(files))


# ------------------------------------------------------------------ 3. смысловая проверка + исправление
NOTE_FIX = ("Независимая проверка смысла нашла в прежнем переводе этого сегмента возможную ошибку: {reason}\n"
            "Прежний перевод: {ru}\nПереведи сегмент заново. Если замечание верно — исправь; если неверно — сохрани "
            "верный смысл. Остальное не меняй без необходимости: термины глоссария, коды, числа, регистр — как в оригинале.")


def review(work, prj, workers=6, compact=True, progress=_quiet):
    import review_portal as RP
    RP.configure(terms=prj.get("terms"), conventions=prj.get("conventions"), glossary=prj["glossary"])
    tr = make_translator(prj)
    main = [e for e in tr.entries if id(e) in tr.main_ids]
    t0, cnt, lock = time.time(), {"n": 0}, threading.Lock()
    plans = []
    for f in doc_files(work):
        d = json.loads(f.read_text(encoding="utf-8"))
        segs = [s for s in d["segs"] if (d["res"].get(str(s["id"])) or {}).get("ru")]
        ru = lambda s: cl(d["res"][str(s["id"])]["ru"])
        jobs, first = [], {}
        for i, s in enumerate(segs):
            r, hu = d["res"][str(s["id"])], cl(s["text"])
            if r.get("copied") or r.get("tm") or r.get("edited") or "review" in r \
                    or not re.search(r"[A-Za-zÁÉÍÓÖŐÚÜŰáéíóöőúüű]{3,}", hu):
                continue
            key = (hu, ru(s))
            first.setdefault(key, []).append(str(s["id"]))
            if len(first[key]) == 1:
                ctx = {"prev": [(cl(p["text"]), ru(p)) for p in segs[max(0, i - 2):i]],
                       "next": [(cl(p["text"]), ru(p)) for p in segs[i + 1:i + 2]], "other": ""}
                jobs.append((key, s, ctx))
        log(f"проверка: {f.stem} — к проверке {len(jobs)}")
        plans.append((f, d, segs, jobs, first))
    total = sum(len(p[3]) for p in plans)
    progress("проверка", 0, total)
    for f, d, segs, jobs, first in plans:
        set_vocab(tr, d["segs"])

        def run(job):
            key, s, ctx = job
            rv = RP.review(key[0], key[1], think=True, ctx=ctx, compact=compact)
            fx = None
            if rv["serious"]:
                # compact: исправленный перевод приходит в том же ответе проверяющего; нет его — отдельный перевод с причиной
                cand = rv.get("fix") if compact else ""
                if not cand:
                    got, _ = tr._call([{"key": 0, "text": s["text"]}], [], "", True, NOTE_FIX.format(reason=rv["reason"], ru=key[1]))
                    cand = got.get(0, "")
                if cand:
                    cand = fix_units(fix_caps(s["text"], cand))
                    rv2 = RP.review(key[0], cl(cand), think=True, ctx=ctx, compact=compact)
                    fx = {"ru": cand, "issues": tr.full_check(s["text"], cand), "rv2": rv2}
            with lock:
                cnt["n"] += 1
                progress("проверка", cnt["n"], total)
                if cnt["n"] % 200 == 0:
                    log(f"  проверено {cnt['n']} за {time.time() - t0:.0f} с; запросов {STAT['calls']}")
            return key, rv, fx

        with ThreadPoolExecutor(workers) as ex:
            results = list(ex.map(run, jobs))
        flagged = fixed = 0
        for key, rv, fx in results:
            flagged += rv["serious"]
            ok = bool(fx) and not fx["rv2"]["serious"]
            accepted = False
            for sid in first[key]:
                r = d["res"][sid]
                r["review"] = {"serious": rv["serious"], "reason": rv["reason"]}
                if fx:
                    accept = ok and badness(fx["issues"]) <= badness(r.get("issues") or [])
                    r["autofix"] = {"before": r["ru"], "reason": rv["reason"], "accepted": accept,
                                    "review2": {"serious": fx["rv2"]["serious"], "reason": fx["rv2"]["reason"]}}
                    if accept:
                        accepted = True
                        seg = next(x for x in segs if str(x["id"]) == sid)
                        r.update(ru=fx["ru"], issues=fx["issues"],
                                 gloss_miss=gloss_misses(TAG.sub("", seg["text"]), fx["ru"], main, outer=tr.entries),
                                 review={"serious": False, "reason": "исправлено автоматически: " + rv["reason"][:200]})
            fixed += accepted
        save_doc(work, f.stem, d)
        log(f"проверка: {f.stem} — отмечено {flagged}, исправлено {fixed}")
    log(f"проверка готова за {time.time() - t0:.0f} с; {STAT}")


# ------------------------------------------------------------------ 3а. шлифовка по языковым соглашениям
def polish(work, prj, progress=_quiet, chat=None):
    """Шлифовка переводов по языковым соглашениям комплекта (prj["rules"], conventions.json) и словарю документа
    (prj["doc_terms"]/<документ>_terms.json; нет словаря — только соглашения): правка кодом всего однозначного,
    порталом — только формулировок, не больше 3 проходов (polish.run). Одинаковая пара «оригинал — перевод» шлифуется
    один раз; правки редактора и память переводов проекта не трогаются. Журнал правок и оставшиеся нарушения —
    в d["polish"] (листы «шлифовка» и «нарушения» таблицы вычитки), у сегмента — r["polish"] (колонка H)."""
    import conventions as C
    import polish as PL
    if chat is None:
        from llm import chat
    files = doc_files(work)
    tr = None
    for i, f in enumerate(files):
        d = json.loads(f.read_text(encoding="utf-8"))
        if "polish" in d or not d.get("res"):
            continue
        progress("шлифовка", i, len(files))
        terms = Path(prj["doc_terms"]) / f"{f.stem}_terms.json" if prj.get("doc_terms") else None
        rules = C.load(prj.get("rules"), terms)
        if terms and terms.exists():
            log(f"шлифовка: {f.stem} — словарь документа {terms.name}, терминов {len(rules['terms'])}")
        items, group = {}, {}
        for s in d["segs"]:
            r = d["res"].get(str(s["id"])) or {}
            if not r.get("ru") or r.get("copied") or r.get("edited") or r.get("tm") or cl(r["ru"]) == cl(s["text"]):
                continue
            k = (s["text"], r["ru"])
            if k not in group:
                group[k] = []
                items[str(s["id"])] = {"hu": s["text"], "ru": r["ru"], "kind": "shot" if s["loc"].get("shot") else "text",
                                       "id": s["id"], "page": s["page"] + 1 if s["page"] is not None else None}
            group[k].append(str(s["id"]))
        before = {k: it["ru"] for k, it in items.items()}
        edits, remaining = PL.run(items, rules, chat=chat, log=log,
                                  progress=lambda x: progress("шлифовка", i + 0.95 * x, len(files)))
        if tr is None:
            tr = make_translator(prj)
        set_vocab(tr, d["segs"])
        main = [e for e in tr.entries if id(e) in tr.main_ids]
        segs = {str(s["id"]): s for s in d["segs"]}
        for k, it in items.items():
            if it["ru"] == before[k]:
                continue
            mine = [e for e in edits if e["key"] == k]
            note = {"before": before[k], "rules": ", ".join(dict.fromkeys(e["rule"] for e in mine)),
                    "who": ", ".join(dict.fromkeys(e["who"] for e in mine)), "of": it["ru"]}
            for sid in group[(it["hu"], before[k])]:
                r, hu = d["res"][sid], segs[sid]["text"]
                r.update(ru=it["ru"], polish=note, issues=tr.full_check(hu, it["ru"]),
                         gloss_miss=gloss_misses(TAG.sub("", hu), it["ru"], main, outer=tr.entries))
        d["polish"] = {"edits": edits, "remaining": remaining, "rules": rules}
        save_doc(work, f.stem, d)
        log(f"шлифовка: {f.stem} — правок {len(edits)} (код {sum(e['who'] == 'код' for e in edits)}, портал "
            f"{sum(e['who'] == 'портал' for e in edits)}), нарушений осталось {len(remaining)}")
    progress("шлифовка", len(files), len(files))


# ------------------------------------------------------------------ 3б. сокращение под место в PDF
FIT_PROMPT = ("Ниже — русский перевод фрагмента технической документации АЭС, который в PDF не помещается в своё место "
              "(ячейка таблицы или строка списка). Сократи его до {n} знаков или короче, сохранив смысл.\n"
              "- Термины, коды, обозначения, числа и латиница — без изменений, все до одного.\n"
              "- Можно убрать повторы и служебные слова, заменить оборот короче, использовать общепринятые сокращения "
              "(т. е., напр., см.).\n- Ничего не добавляй. Выведи только сокращённый текст.\n\n"
              "Оригинал (венгерский): {hu}\nПеревод: {ru}")


# короткая надпись (штамп, шапка таблицы): только сокращение тех же слов — портал менял слово на другое
# («Проектировщик:» → «Проект:», штамп UTS, 15.10)
FIT_PROMPT_SHORT = ("Надпись в тесной ячейке технического документа не помещается. Сократи её до {n} знаков или короче, "
                    "сокращая сами слова с точкой (Проектировщик → Проектир., Обязательный → Обяз.), не заменяя их другими "
                    "словами. Числа, коды и латиница — без изменений. Выведи только результат.\nНадпись: {ru}")


def _abbr_of(short, ru, hu):
    """Короткая надпись — те же слова или их сокращения с точкой (или исходное латинское сокращение)."""
    if short.strip() == hu.strip():
        return True
    words = [w.strip(":;,()«»").lower() for w in ru.split()]
    for w in short.split():
        w = w.strip(":;,()«»").lower()
        if not w or w in words:
            continue
        if not (w.endswith(".") and len(w) >= 4 and any(x.startswith(w[:-1]) for x in words)):
            return False
    return True


def fit(work, prj=None, limit=0.7):
    """Короткая формулировка для PDF там, где перевод ужимается мельче limit от исходного кегля (сборка «на месте»
    не может раздвинуть ячейку или строку списка). Портал сокращает, код проверяет: коды и числа те же, текст
    короче. Сокращённое — только в PDF (r["ru_pdf"]); в DOCX строка растёт, там полный перевод."""
    import tempfile
    from llm import chat
    from translate import CODE
    from assemble import build_pdf
    for f in doc_files(work):
        d = json.loads(f.read_text(encoding="utf-8"))
        if not d["file"].lower().endswith(".pdf"):
            continue
        with tempfile.TemporaryDirectory() as tmp:
            info = build_pdf(d["file"], Path(tmp) / "trial.pdf", d["segs"], d["res"])
        by_id = {str(s["id"]): s for s in d["segs"]}
        done = 0
        for it in info.get("shrunk", []):
            s, r = by_id[str(it["id"])], d["res"][str(it["id"])]
            ru = cl(r["ru"])
            # место под текст ~ квадрат кегля: до масштаба 0,75 — длина × (масштаб/0,75)², но не короче 60 %
            # (сокращать вдвое — терять смысл)
            n = max(8, int(len(ru) * max((it["scale"] / 0.75) ** 2, 0.6)))
            label = len(ru.split()) <= 3
            if r.get("ru_pdf") and r.get("ru_pdf_of") == r["ru"] and (not label or _abbr_of(r["ru_pdf"], ru, cl(s["text"]))):
                continue
            r.pop("ru_pdf", None); r.pop("ru_pdf_of", None)
            if n >= len(ru):
                continue
            prompt = (FIT_PROMPT_SHORT.format(n=n, ru=ru) if label
                      else FIT_PROMPT.format(n=n, hu=cl(s["text"]), ru=ru))
            out = chat([{"role": "user", "content": prompt}], max_tokens=2000, tag=f"fit{it['id']}")
            short = (out.get("text") or "").strip().strip("«»\"")
            codes = lambda t: sorted(CODE.findall(t))
            if short and len(short) < len(ru) and codes(short) == codes(ru) \
                    and (not label or _abbr_of(short, ru, cl(s["text"]))):
                r["ru_pdf"], r["ru_pdf_of"] = short, r["ru"]; done += 1
        save_doc(work, f.stem, d)
        log(f"сокращение: {f.stem} — ужато {len(info.get('shrunk', []))}, сокращено {done}")


# ------------------------------------------------------------------ 4. сборка
def assemble(work, progress=_quiet):
    from assemble import assemble_all
    files = doc_files(work)
    for i, f in enumerate(files):
        progress("сборка", i, len(files))
        d = json.loads(f.read_text(encoding="utf-8"))
        info = assemble_all(f, d, Path(work) / "out", progress=lambda x: progress("сборка", i + x, len(files)))
        log(f"сборка: {f.stem} — {str(info)[:80]}")
    rebuild_portfolios(work, Path(work) / "out")
    progress("сборка", len(files), len(files))


# ------------------------------------------------------------------ 5. правки редактора + RAG
W = re.compile(r"[a-záéíóöőúüű]{4,}")
NOTE_RAG = ("Образец, утверждённый редактором для похожей фразы этого же комплекта:\nОригинал: {hu}\nПеревод: {ru}\n"
            "Переведи сегмент по этому образцу: те же термины, формулировки и смысловые решения там, где фраза совпадает; "
            "то, что отличается (числа, коды, объекты, условия), — по своему оригиналу.")


def edits(work, prj, sim=0.8, workers=4):
    """Колонка G таблиц вычитки → перевод и память переводов; затем похожие фразы (≥ sim) по всему комплекту
    переводятся заново с утверждённой парой как образцом (замер 05.10: лучше 41 / хуже 1 из 92)."""
    import shutil
    import apply_review as AR
    out = Path(work) / "out"
    keep = out / "правки_редактора" / time.strftime("%Y%m%d_%H%M%S")
    pairs = []
    for f in doc_files(work):
        xf = out / f"{f.stem}_вычитка.xlsx"
        if xf.exists():
            # таблица пересобирается заново (колонка G пустеет) — файл редактора сохраняем как есть
            keep.mkdir(parents=True, exist_ok=True); shutil.copy(xf, keep / xf.name)
            pairs += AR.apply(f, xf, prj["tm"], out)
    if not pairs:
        log("правок редактора нет"); return
    changed = propagate(work, prj, pairs, sim, workers)
    from assemble import assemble_all
    for stem, d in changed.items():
        assemble_all(Path(work) / "json" / f"{stem}.json", d, out)


def propagate(work, prj, pairs, sim=0.8, workers=4, progress=_quiet):
    """Утверждённые пары (венгерский, русский) → та же фраза с другими числами получает правку подстановкой,
    похожие фразы (≥ sim) переводятся заново с парой как образцом. Документы сохраняются; возвращает
    {док: данные} изменённых — сборку делает вызывающий (CLI — только изменённые, веб — все)."""
    docs = load_docs(work)
    uniq = {}
    for stem, d in docs.items():
        for s in d["segs"]:
            r = d["res"].get(str(s["id"])) or {}
            if r.get("ru") and not r.get("copied") and not r.get("edited"):
                uniq.setdefault(tmkey(s["text"]), []).append((stem, str(s["id"]), s))
    texts = list(uniq); idx, by_num = collections.defaultdict(set), collections.defaultdict(list)
    for i, t in enumerate(texts):
        by_num[NUM.sub("§", t)].append(t)
        for w in set(W.findall(t.lower())):
            idx[w].add(i)
    tr = make_translator(prj)
    set_vocab(tr, [s for d in docs.values() for s in d["segs"]])
    jobs, seen, direct = [], set(), []
    for hu, ru in pairs:
        key = tmkey(hu); cand = collections.Counter()
        # та же фраза с другими числами (колонтитулы «Oldalszám: 95» — сотни вариантов): правка с подстановкой
        # чисел, без портала — иначе одна правка колонтитула стоила бы десятков запросов
        if NUM.findall(ru) == NUM.findall(key):
            for t in by_num[NUM.sub("§", key)]:
                if t != key and t not in seen:
                    it = iter(NUM.findall(t))
                    seen.add(t); direct.append(((hu, ru, uniq[t]), NUM.sub(lambda _: next(it), ru)))
        for w in set(W.findall(key.lower())):
            if len(idx[w]) < 400:
                for j in idx[w]:
                    cand[j] += 1
        for j, _ in cand.most_common(60):
            t = texts[j]
            if t != key and t not in seen and difflib.SequenceMatcher(None, key, t).ratio() >= sim:
                seen.add(t); jobs.append((hu, ru, uniq[t]))
    log(f"правок {len(pairs)}; та же фраза с другими числами {len(direct)}; похожих фраз к переводу по образцу {len(jobs)}")

    def run(job):
        hu, ru, places = job
        seg = places[0][2]
        got, _ = tr._call([{"key": 0, "text": seg["text"]}], [], "", True, NOTE_RAG.format(hu=cl(hu), ru=ru))
        return job, got.get(0, "")

    cnt, lock = {"n": 0}, threading.Lock()
    progress("правки", 0, len(jobs))

    def run_counted(job):
        out = run(job)
        with lock:
            cnt["n"] += 1
            progress("правки", cnt["n"], len(jobs))
        return out

    with ThreadPoolExecutor(workers) as ex:
        results = list(ex.map(run_counted, jobs))
    changed = set(); acc = 0
    # numbers=True — та же фраза с другими числами: решение редактора, повторённое механически (не «проверить»)
    for ((hu, ru, places), ru2), numbers in [(x, True) for x in direct] + [(x, False) for x in results]:
        if not ru2:
            continue
        stem0, sid0, seg = places[0]
        ru2 = fix_units(fix_caps(seg["text"], ru2))
        iss2 = tr.full_check(seg["text"], ru2)
        if badness(iss2) <= badness(docs[stem0]["res"][sid0].get("issues") or []):
            acc += 1
            for stem, sid, _ in places:
                r = docs[stem]["res"][sid]
                r.setdefault("rag", {"before": r["ru"], "example": cl(hu), "numbers": numbers})
                r.pop("review", None); r.pop("autofix", None)      # вердикт был о прежнем переводе
                r.update(ru=ru2, issues=iss2); changed.add(stem)
    for stem in changed:
        save_doc(work, stem, docs[stem])
    log(f"по образцу принято {acc} из {len(direct) + len(results)}; изменено документов {len(changed)}")
    progress("правки", len(jobs), len(jobs))
    return {stem: docs[stem] for stem in changed}


# ------------------------------------------------------------------ термины нового комплекта (портал)
def terms(work, prj, out_md, top=300):
    """Кандидаты в термины (частые сочетания вне глоссария) → перевод порталом → md для проверки специалистом."""
    import termx
    texts = {s["text"] for d in load_docs(work).values() for s in d["segs"] if S.needs_translation(s["text"])}
    _, entries = G.load(prj["glossary"])
    cands = termx.candidates(list(texts), entries, top)
    res = termx.translate_terms(cands)
    termx.write_md(res, out_md, "комплект")
    log(f"терминов {len(res)} (не уверен: {sum(not t['sure'] for t in res)}) → {out_md}")
