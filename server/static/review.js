// Вычитка: таблица Tabulator, фильтры-чипы, боковая панель (HTMX), правка в ячейке, горячие клавиши.
(function () {
  const el = document.getElementById("table");
  const uid = el.dataset.uid, doc = el.dataset.doc;
  const base = "/jobs/" + encodeURIComponent(uid) + "/review";
  const ATTENTION = ["error", "suspect", "auto"];
  let current = null;            // выбранная строка

  const table = new Tabulator(el, {
    ajaxURL: base + "/rows?doc=" + encodeURIComponent(doc),
    height: "70vh",
    layout: "fitColumns",
    index: "id",
    placeholder: "В этом фильтре строк нет",
    rowFormatter: (row) => {
      const e = row.getElement();
      e.className = e.className.replace(/\bk-\S+/g, "");
      e.classList.add("k-" + row.getData().kind);
    },
    columns: [
      {title: "№", field: "id", width: 64, headerSort: false},
      {title: "стр.", field: "page", width: 56, headerSort: false},
      {title: "Венгерский", field: "hu", formatter: "textarea", headerSort: false},
      {title: "Русский", field: "ru", formatter: "textarea", headerSort: false,
       editor: "textarea", editorParams: {shiftEnterSubmit: true, selectContents: false}},
    ],
  });

  function select(row) {
    if (current) current.getElement().classList.remove("is-current");
    current = row;
    row.getElement().classList.add("is-current");
    htmx.ajax("GET", base + "/side?doc=" + encodeURIComponent(doc) + "&sid=" + encodeURIComponent(row.getData().id),
              {target: "#side", swap: "innerHTML"});
  }

  table.on("rowClick", (e, row) => select(row));

  function setFilter(kind) {
    document.querySelectorAll(".chip").forEach((c) => c.setAttribute("aria-pressed", String(c.dataset.kind === kind)));
    if (!kind) table.clearFilter();
    else if (kind === "attention") table.setFilter("kind", "in", ATTENTION);
    else table.setFilter("kind", "=", kind);
  }
  document.querySelectorAll(".chip").forEach((c) => c.addEventListener("click", () => setFilter(c.dataset.kind)));
  // по умолчанию — строки, требующие внимания; если таких нет, все (пустая таблица на входе сбивает с толку)
  table.on("tableBuilt", () => setFilter(el.dataset.filter));

  function recount() {                       // счётчики на чипах — после каждой правки
    const n = {"": 0, attention: 0};
    table.getData().forEach((r) => {
      n[""] += 1; n[r.kind] = (n[r.kind] || 0) + 1;
      if (ATTENTION.includes(r.kind)) n.attention += 1;
    });
    document.querySelectorAll(".chip").forEach((c) => { c.textContent = c.dataset.label + " " + (n[c.dataset.kind] || 0); });
  }

  async function save(id, ru) {
    const body = new FormData();
    body.append("doc", doc); body.append("sid", id); body.append("ru", ru);
    const r = await fetch(base + "/edit", {method: "POST", body});
    const data = await r.json().catch(() => ({error: "Сервер не ответил. Попробуйте ещё раз."}));
    if (!r.ok || data.error) throw new Error(data.error || "Правка не сохранилась.");
    data.changed.forEach((c) => {
      const row = table.getRow(c.id);
      if (row) { row.update({ru: c.ru, kind: c.kind}); row.reformat(); }
    });
    recount();
    document.getElementById("pending-count").textContent = data.pending;
    document.getElementById("pending").hidden = !data.pending;
    const row = table.getRow(id);
    if (row && row === current) select(row);
  }

  table.on("cellEdited", (cell) => {
    save(cell.getRow().getData().id, cell.getValue()).catch((err) => { cell.restoreOldValue(); alert(err.message); });
  });

  window.hutRevert = (id, before) => {
    save(id, before).catch((err) => alert(err.message));
  };

  document.addEventListener("keydown", (e) => {
    if (e.target.tagName === "TEXTAREA" || e.target.tagName === "INPUT") return;
    if (e.altKey && (e.key === "ArrowDown" || e.key === "ArrowUp")) {
      const rows = table.getRows("active");
      if (!rows.length) return;
      const i = current ? rows.indexOf(current) : -1;
      const next = rows[Math.min(Math.max(i + (e.key === "ArrowDown" ? 1 : -1), 0), rows.length - 1)];
      table.scrollToRow(next, "center", false);
      select(next);
      e.preventDefault();
    } else if (e.key === "Enter" && current) {
      current.getCell("ru").edit();
      e.preventDefault();
    }
  });
})();
