/* ExamHub - ListView: Senfyar-style list. Table with sortable headers on wide screens,
   labelled cards + sort control on phones, "show more" paging. Pure rendering, no data access. */
(function () {
"use strict";
const esc = (s) => String(s ?? "").replace(/[&<>"']/g, (c) => ({ "&": "&amp;", "<": "&lt;", ">": "&gt;", '"': "&quot;", "'": "&#39;" }[c]));
const mq = window.matchMedia ? window.matchMedia("(min-width:700px)") : null;

function cmp(a, b, dir) {
  const ea = a === "" || a == null || a === Infinity, eb = b === "" || b == null || b === Infinity;
  if (ea && eb) return 0;
  if (ea) return 1;          // empty values always last, whatever the direction
  if (eb) return -1;
  const r = typeof a === "number" && typeof b === "number" ? a - b : String(a).localeCompare(String(b), "fa", { numeric: true });
  return dir === "desc" ? -r : r;
}

/* cfg: { columns:[{key,label,sort(row)->value,render(row)->html,primary?,badge?}], rows():array,
          rowAttrs(row)->string, defaultSort:{key,dir}, empty:string, pageSize? } */
window.ListView = function (el, cfg) {
  const st = { key: cfg.defaultSort.key, dir: cfg.defaultSort.dir, limit: cfg.pageSize || 50 };
  const cols = cfg.columns;
  const primary = cols.find((c) => c.primary) || cols[0];
  const badge = cols.find((c) => c.badge);

  function render() {
    const col = cols.find((c) => c.key === st.key) || cols[0];
    const rows = [...cfg.rows()].sort((a, b) => cmp(col.sort(a), col.sort(b), st.dir));
    const shown = rows.slice(0, st.limit);
    const wide = !!(mq && mq.matches);
    let html;
    if (!rows.length) {
      html = `<div class="empty">${cfg.empty || "موردی یافت نشد."}</div>`;
    } else if (wide) {
      html = `<div class="lv-wrap"><table class="lv-table"><thead><tr>${cols.map((c) =>
        `<th class="sortable" data-lv-sort="${c.key}">${c.label}${c.key === st.key ? (st.dir === "asc" ? " ▲" : " ▼") : ""}</th>`).join("")}</tr></thead><tbody>${
        shown.map((r) => `<tr class="lv-row" ${cfg.rowAttrs(r)}>${cols.map((c) => `<td>${c.render(r) || '<span class="muted">—</span>'}</td>`).join("")}</tr>`).join("")}</tbody></table></div>`;
    } else {
      const sortBar = `<div class="lv-sortbar"><label>مرتب‌سازی</label><select class="lv-sortsel">${cols.map((c) =>
        `<option value="${c.key}" ${c.key === st.key ? "selected" : ""}>${c.label}</option>`).join("")}</select>
        <button type="button" class="btn secondary sm lv-dir" aria-label="جهت مرتب‌سازی">${st.dir === "asc" ? "▲ صعودی" : "▼ نزولی"}</button></div>`;
      html = sortBar + `<div class="lv-cards">${shown.map((r) => {
        const fields = cols.filter((c) => c !== primary && c !== badge).map((c) => {
          const v = c.render(r);
          return v ? `<div class="lv-field"><span>${c.label}</span><b>${v}</b></div>` : "";
        }).join("");
        return `<div class="lv-card lv-row" ${cfg.rowAttrs(r)}>
          <div class="lv-card-head"><div class="lv-card-title">${primary.render(r)}</div>${badge ? badge.render(r) : ""}</div>
          ${fields ? `<div class="lv-fields">${fields}</div>` : ""}</div>`;
      }).join("")}</div>`;
    }
    if (rows.length > shown.length) {
      html += `<button type="button" class="btn secondary block lv-more">نمایش بیشتر (${(rows.length - shown.length).toLocaleString("fa-IR")} مورد باقی‌مانده)</button>`;
    }
    el.innerHTML = html;
    return rows.length;
  }

  function setSort(key) {
    if (st.key === key) st.dir = st.dir === "asc" ? "desc" : "asc"; else { st.key = key; st.dir = "asc"; }
    st.limit = cfg.pageSize || 50;
    render();
  }
  el.addEventListener("click", (e) => {
    const th = e.target.closest("[data-lv-sort]");
    if (th) return setSort(th.dataset.lvSort);
    if (e.target.closest(".lv-dir")) { st.dir = st.dir === "asc" ? "desc" : "asc"; return render(); }
    if (e.target.closest(".lv-more")) { st.limit += cfg.pageSize || 50; render(); }
  });
  el.addEventListener("change", (e) => {
    if (e.target.classList.contains("lv-sortsel")) { st.key = e.target.value; st.limit = cfg.pageSize || 50; render(); }
  });
  if (mq && mq.addEventListener) mq.addEventListener("change", render);

  return { render, resetPaging() { st.limit = cfg.pageSize || 50; } };
};
})();
