/* ExamHub - app.js
   Stage 3 build: Applicants, Exams, Teachers/Referrers, Dashboard, Jalali Calendar,
   Telegram message templates, Reports + Excel export, JSON Backup/Restore, PWA,
   and two-way sync with a Google Sheet (via a Google Apps Script web app).
   Not yet implemented (see Settings > About): passport OCR, Google Drive photos,
   encryption of sensitive fields, auto-lock with password.
*/

(() => {
"use strict";

const { toLatinDigits, Jalali } = window; // from js/jalali.js

/* ---------------------------------------------------------------------- */
/* Small utilities                                                        */
/* ---------------------------------------------------------------------- */
const uid = () => (crypto.randomUUID ? crypto.randomUUID() : "id-" + Date.now() + "-" + Math.random().toString(16).slice(2));
const $ = (sel, root = document) => root.querySelector(sel);
const $$ = (sel, root = document) => Array.from(root.querySelectorAll(sel));
const nowIso = () => new Date().toISOString();
function esc(s) { return String(s ?? "").replace(/[&<>"']/g, c => ({"&":"&amp;","<":"&lt;",">":"&gt;",'"':"&quot;","'":"&#39;"}[c])); }
function toast(msg) {
  const t = $("#toast"); t.textContent = msg; t.classList.add("show");
  clearTimeout(toast._t); toast._t = setTimeout(() => t.classList.remove("show"), 2200);
}
function fmtMoney(n) { n = Number(n) || 0; return n.toLocaleString("en-US"); }
function todayJalaliStr() { const j = Jalali.fromDate(new Date()); return Jalali.format(...j); }

// Keys of upcoming/done/cancelled/rescheduled are kept so existing data and sheet rows stay valid.
const STATUS_LABEL = {
  requested: "متقاضی آزمون",
  upcoming: "ثبت‌نام شده",
  rescheduled: "تغییر تاریخ",
  done: "آزمون داده",
  awaiting: "منتظر نتیجه",
  success: "موفق",
  suspended: "معلق",
  banned: "بن شده",
  cancelled: "کنسل شده",
  cancelled_by: "کنسل کرده",
};
// Older labels that may still sit in the Google Sheet.
const STATUS_ALIASES = { "پیش‌رو": "upcoming", "انجام‌شده": "done", "لغوشده": "cancelled" };
// kind drives dashboard counts and calendar colours: open = waiting for the exam day, past = exam taken,
// hold = not decided yet, bad = did not happen.
const STATUS_KIND = { requested: "hold", upcoming: "open", rescheduled: "open", done: "past", awaiting: "past", success: "past", suspended: "hold", banned: "bad", cancelled: "bad", cancelled_by: "bad" };
$("#examStatusFilter").insertAdjacentHTML("beforeend", Object.entries(STATUS_LABEL).map(([k, v]) => `<option value="${k}">${v}</option>`).join(""));
$("#repStatus").insertAdjacentHTML("beforeend", Object.entries(STATUS_LABEL).map(([k, v]) => `<option value="${k}">${v}</option>`).join(""));

/* ---------------------------------------------------------------------- */
/* IndexedDB layer                                                        */
/* ---------------------------------------------------------------------- */
const DB_NAME = "examhub-db";
const DB_VERSION = 1;
let dbPromise = null;

function openDb() {
  if (dbPromise) return dbPromise;
  dbPromise = new Promise((resolve, reject) => {
    const req = indexedDB.open(DB_NAME, DB_VERSION);
    req.onupgradeneeded = (e) => {
      const db = e.target.result;
      if (!db.objectStoreNames.contains("applicants")) {
        db.createObjectStore("applicants", { keyPath: "id" });
      }
      if (!db.objectStoreNames.contains("exams")) {
        const s = db.createObjectStore("exams", { keyPath: "id" });
        s.createIndex("applicantId", "applicantId");
      }
      if (!db.objectStoreNames.contains("teachers")) db.createObjectStore("teachers", { keyPath: "id" });
      if (!db.objectStoreNames.contains("referrers")) db.createObjectStore("referrers", { keyPath: "id" });
      if (!db.objectStoreNames.contains("examTypes")) db.createObjectStore("examTypes", { keyPath: "id" });
      if (!db.objectStoreNames.contains("meta")) db.createObjectStore("meta", { keyPath: "key" });
    };
    req.onsuccess = () => resolve(req.result);
    req.onerror = () => reject(req.error);
  });
  return dbPromise;
}
function tx(storeNames, mode = "readonly") {
  return openDb().then(db => db.transaction(storeNames, mode));
}
function storeAll(storeName) {
  return tx([storeName]).then(t => new Promise((resolve, reject) => {
    const req = t.objectStore(storeName).getAll();
    req.onsuccess = () => resolve(req.result);
    req.onerror = () => reject(req.error);
  }));
}
// Stores that are mirrored to the Google Sheet. Deleting one of their records
// writes a tombstone ({id, deleted:true}) so the deletion can reach other devices.
const SYNC_STORES = ["applicants", "exams", "teachers", "referrers", "examTypes"];

function storePutRaw(storeName, value) {
  return tx([storeName], "readwrite").then(t => new Promise((resolve, reject) => {
    const req = t.objectStore(storeName).put(value);
    req.onsuccess = () => resolve(value);
    req.onerror = () => reject(req.error);
  }));
}
function storePut(storeName, value) {
  if (!SYNC_STORES.includes(storeName)) return storePutRaw(storeName, value);
  value.updatedAt = nowIso();
  return storePutRaw(storeName, value).then(v => { onLocalChange(); return v; });
}
function storeDelete(storeName, id) {
  if (SYNC_STORES.includes(storeName)) {
    return storePutRaw(storeName, { id, deleted: true, updatedAt: nowIso() }).then(() => onLocalChange());
  }
  return tx([storeName], "readwrite").then(t => new Promise((resolve, reject) => {
    const req = t.objectStore(storeName).delete(id);
    req.onsuccess = () => resolve();
    req.onerror = () => reject(req.error);
  }));
}
function storeClear(storeName) {
  return tx([storeName], "readwrite").then(t => new Promise((resolve, reject) => {
    const req = t.objectStore(storeName).clear();
    req.onsuccess = () => resolve();
    req.onerror = () => reject(req.error);
  }));
}
function metaGet(key) {
  return tx(["meta"]).then(t => new Promise((resolve, reject) => {
    const req = t.objectStore("meta").get(key);
    req.onsuccess = () => resolve(req.result ? req.result.value : null);
    req.onerror = () => reject(req.error);
  }));
}
function metaSet(key, value) { return storePut("meta", { key, value }); }

/* ---------------------------------------------------------------------- */
/* In-memory cache (refreshed from IDB) for fast rendering                */
/* ---------------------------------------------------------------------- */
const DATA = { applicants: [], exams: [], teachers: [], referrers: [], examTypes: [] };

const isLive = (r) => !r.deleted;
const typeId = (name) => "etype:" + name.trim().toLowerCase();

async function refreshAll() {
  DATA.applicants = (await storeAll("applicants")).filter(isLive);
  DATA.exams = (await storeAll("exams")).filter(isLive);
  DATA.teachers = (await storeAll("teachers")).filter(isLive);
  DATA.referrers = (await storeAll("referrers")).filter(isLive);
  // exam types: one entry per name (older data may contain the same name twice)
  const seen = new Map();
  for (const t of (await storeAll("examTypes")).filter(isLive)) {
    const k = (t.name || "").trim().toLowerCase();
    if (k && !seen.has(k)) seen.set(k, t);
  }
  DATA.examTypes = [...seen.values()];
}

async function seedDefaultsIfEmpty() {
  if (DATA.examTypes.length === 0) {
    const defaults = ["IELTS", "TOEFL", "GRE", "Duolingo", "سایر"];
    for (const name of defaults) {
      await storePut("examTypes", { id: typeId(name), name });
    }
    await refreshAll();
  }
}

/* ---------------------------------------------------------------------- */
/* Modal helper                                                            */
/* ---------------------------------------------------------------------- */
function openModal(html) {
  $("#modalContent").innerHTML = html;
  $("#modalOverlay").classList.add("open");
}
function closeModal() {
  $("#modalOverlay").classList.remove("open");
  $("#modalContent").innerHTML = "";
}
$("#modalOverlay").addEventListener("click", (e) => { if (e.target.id === "modalOverlay") closeModal(); });

function confirmDialog(message, onYes) {
  openModal(`
    <div class="modal-head"><h3>تأیید</h3></div>
    <p>${esc(message)}</p>
    <div class="modal-actions">
      <button class="btn danger block" id="cfYes">تأیید</button>
      <button class="btn secondary block" id="cfNo">انصراف</button>
    </div>`);
  $("#cfYes").onclick = () => { closeModal(); onYes(); };
  $("#cfNo").onclick = closeModal;
}

/* ---------------------------------------------------------------------- */
/* Tabs / routing                                                          */
/* ---------------------------------------------------------------------- */
function showView(name) {
  $$(".view").forEach(v => v.classList.remove("active"));
  $(`#view-${name}`).classList.add("active");
  $$(".tab-btn").forEach(b => b.classList.toggle("active", b.dataset.view === name));
  render(name);
}
$$(".tab-btn").forEach(b => b.addEventListener("click", () => showView(b.dataset.view)));

function render(name) {
  if (name === "dashboard") renderDashboard();
  else if (name === "people") setPeopleTab(activePeopleTab);
  else if (name === "exams") renderExams();
  else if (name === "calendar") renderCalendar();
  else if (name === "reports") renderReportsSetup();
  else if (name === "messages") renderMessagesSetup();
  else if (name === "settings") renderSettings();
}

/* ---------------------------------------------------------------------- */
/* Theme                                                                   */
/* ---------------------------------------------------------------------- */
function applyTheme(t) { document.documentElement.setAttribute("data-theme", t); localStorage.setItem("examhub-theme", t); }
$("#btnTheme").addEventListener("click", () => {
  const cur = document.documentElement.getAttribute("data-theme") === "dark" ? "light" : "dark";
  applyTheme(cur);
});
applyTheme(localStorage.getItem("examhub-theme") || "light");

/* ---------------------------------------------------------------------- */
/* DASHBOARD                                                               */
/* ---------------------------------------------------------------------- */
function applicantName(a) { return a ? `${a.firstName || ""} ${a.lastName || ""}`.trim() : "—"; }
function getApplicant(id) { return DATA.applicants.find(a => a.id === id); }
function getTeacher(id) { return DATA.teachers.find(t => t.id === id); }
function getReferrer(id) { return DATA.referrers.find(r => r.id === id); }

function examSortKey(e) {
  // sort by jalali date+time ascending; exams without a date go last
  if (!e.examDateJ) return Infinity;
  const [jy, jm, jd] = e.examDateJ;
  const t = e.examTime || "00:00";
  return jy * 100000000 + jm * 1000000 + jd * 10000 + parseInt(t.replace(":", ""), 10);
}

// an exam is "ahead" if it is dated today or later (undated ones count as ahead)
function isAhead(e) {
  if (!e.examDateJ) return true;
  return jalaliValue(e.examDateJ) >= jalaliValue(Jalali.fromDate(new Date()));
}
const isOpen = (e) => STATUS_KIND[e.status] === "open";

function renderDashboard() {
  $("#statApplicants").textContent = DATA.applicants.length;
  $("#statUpcoming").textContent = DATA.exams.filter(e => isOpen(e) && isAhead(e)).length;
  const overdue = DATA.exams.filter(e => isOpen(e) && !isAhead(e)).length;
  const note = $("#overdueNote");
  note.style.display = overdue ? "block" : "none";
  note.textContent = overdue ? `⚠️ ${overdue} آزمون با تاریخ گذشته هنوز «ثبت‌نام شده» است — برای بررسی بزنید` : "";
  $("#statDone").textContent = DATA.exams.filter(e => STATUS_KIND[e.status] === "past").length;
  const totalAmount = DATA.exams.reduce((s, e) => s + (Number(e.amount) || 0), 0);
  $("#statAmount").textContent = fmtMoney(totalAmount);

  const todayJ = Jalali.fromDate(new Date());
  const todayStr = Jalali.format(...todayJ);
  const todayList = DATA.exams.filter(e => e.examDateJ && Jalali.format(...e.examDateJ) === todayStr);
  $("#todayExams").innerHTML = todayList.length ? renderExamRows(todayList) : `<div class="empty">آزمونی برای امروز ثبت نشده.</div>`;

  const upcoming = DATA.exams
    .filter(e => isOpen(e) && isAhead(e))
    .sort((a, b) => examSortKey(a) - examSortKey(b))
    .slice(0, 8);
  $("#upcomingExams").innerHTML = upcoming.length ? renderExamRows(upcoming) : `<div class="empty">آزمونی ثبت نشده.</div>`;

  const recent = [...DATA.applicants].sort((a, b) => (b.createdAt || "").localeCompare(a.createdAt || "")).slice(0, 6);
  $("#recentApplicants").innerHTML = recent.length ? recent.map(a => `
    <div class="list-item" data-open-applicant="${a.id}">
      <div class="main"><div class="title">${esc(applicantName(a))}</div><div class="sub">${esc(a.phone || "")}</div></div>
      <span>›</span>
    </div>`).join("") : `<div class="empty">متقاضی‌ای ثبت نشده.</div>`;
}

function renderExamRows(list) {
  return list.map(e => {
    const a = getApplicant(e.applicantId);
    return `<div class="list-item" data-open-exam="${e.id}">
      <div class="main">
        <div class="title">${esc(applicantName(a))} — ${esc(e.examType || "")}</div>
        <div class="sub">${e.examDateJ ? esc(Jalali.format(...e.examDateJ)) : "بدون تاریخ"} ${e.examTime ? "ساعت " + esc(e.examTime) : ""}</div>
      </div>
      <span class="chip ${e.status}">${STATUS_LABEL[e.status] || e.status}</span>
    </div>`;
  }).join("");
}

document.addEventListener("click", (e) => {
  const openA = e.target.closest("[data-open-applicant]");
  if (openA) { openApplicantModal(openA.dataset.openApplicant); return; }
  const openE = e.target.closest("[data-open-exam]");
  if (openE) { openExamModal(openE.dataset.openExam); return; }
});

/* ---------------------------------------------------------------------- */
/* APPLICANTS                                                              */
/* ---------------------------------------------------------------------- */
function renderApplicants() {
  const q = ($("#searchApplicant").value || "").trim().toLowerCase();
  let list = [...DATA.applicants];
  if (q) list = list.filter(a =>
    applicantName(a).toLowerCase().includes(q) ||
    (a.phone || "").includes(q) ||
    (a.passportNumber || "").toLowerCase().includes(q)
  );
  list.sort((a, b) => applicantName(a).localeCompare(applicantName(b), "fa"));
  $("#applicantCount").textContent = `${list.length.toLocaleString("fa-IR")} نفر`;
  $("#applicantsList").innerHTML = list.length ? list.map(a => {
    const examCount = DATA.exams.filter(x => x.applicantId === a.id).length;
    return `<div class="list-item" data-open-applicant="${a.id}">
      <div class="row-content">
        <div class="row-avatar">${esc((a.firstName || a.lastName || "؟").trim().charAt(0))}</div>
        <div class="main"><div class="title">${esc(applicantName(a))}</div>
          <div class="sub row-meta">${a.phone ? `<span>☎ ${esc(a.phone)}</span>` : ""}${a.passportNumber ? `<span>پاسپورت: ${esc(a.passportNumber)}</span>` : ""}</div>
        </div>
      </div>
      <span class="teacher-exam-count">${examCount.toLocaleString("fa-IR")} آزمون</span><span class="muted">‹</span>
    </div>`;
  }).join("") : `<div class="empty">متقاضی‌ای یافت نشد.</div>`;
}
$("#searchApplicant").addEventListener("input", renderApplicants);

// Teachers are shown in their own list view; their underlying records and sync schema stay unchanged.
function renderTeachers() {
  const q = ($("#searchTeacher").value || "").trim().toLowerCase();
  let list = [...DATA.teachers];
  if (q) list = list.filter(t => (t.name || "").toLowerCase().includes(q) || (t.phone || "").includes(q));
  list.sort((x,y) => (x.name || "").localeCompare(y.name || "", "fa"));
  $("#teacherCount").textContent = `${list.length.toLocaleString("fa-IR")} استاد`;
  $("#teachersList").innerHTML = list.length ? list.map(t => {
    const exams = DATA.exams.filter(e => e.teacherId === t.id);
    return `<div class="list-item">
      <div class="row-content">
        <div class="row-avatar">ا</div>
        <div class="main"><div class="title">${esc(t.name)}</div>
          <div class="sub row-meta">${t.phone ? `<span>☎ ${esc(t.phone)}</span>` : "<span>شماره تماسی ثبت نشده</span>"}</div>
        </div>
      </div>
      <div style="display:flex;align-items:center;gap:6px">
        <span class="teacher-exam-count">${exams.length.toLocaleString("fa-IR")} آزمون</span>
        <button class="btn ghost sm" data-edit-teacher="${t.id}">ویرایش</button>
      </div>
    </div>`;
  }).join("") : `<div class="empty">استادی ثبت نشده است. از دکمه افزودن استاد شروع کنید.</div>`;
}
$("#searchTeacher").addEventListener("input", renderTeachers);
$("#btnAddTeacherPage").addEventListener("click", () => openPersonForm("teachers"));

function renderReferrers() {
  const q = ($("#searchReferrer").value || "").trim().toLowerCase();
  let list = [...DATA.referrers];
  if (q) list = list.filter(r => (r.name || "").toLowerCase().includes(q) || (r.phone || "").includes(q));
  list.sort((a,b) => (a.name || "").localeCompare(b.name || "", "fa"));
  $("#referrerCount").textContent = `${list.length.toLocaleString("fa-IR")} معرف`;
  $("#referrersPeopleList").innerHTML = list.length ? list.map(r => {
    const exams = DATA.exams.filter(e => e.referrerId === r.id);
    return `<div class="list-item">
      <div class="row-content">
        <div class="row-avatar">م</div>
        <div class="main"><div class="title">${esc(r.name)}</div>
          <div class="sub row-meta">${r.phone ? `<span>☎ ${esc(r.phone)}</span>` : "<span>شماره تماسی ثبت نشده</span>"}</div>
        </div>
      </div>
      <div style="display:flex;align-items:center;gap:6px">
        <span class="teacher-exam-count">${exams.length.toLocaleString("fa-IR")} معرفی</span>
        <button class="btn ghost sm" data-edit-referrer="${r.id}">ویرایش</button>
      </div>
    </div>`;
  }).join("") : `<div class="empty">معرفی ثبت نشده است. از دکمه افزودن معرف شروع کنید.</div>`;
}
$("#searchReferrer").addEventListener("input", renderReferrers);
$("#btnAddReferrerPage").addEventListener("click", () => openPersonForm("referrers"));

let activePeopleTab = "applicants";
function setPeopleTab(name) {
  if (!["applicants", "teachers", "referrers"].includes(name)) name = "applicants";
  activePeopleTab = name;
  $$(".people-tab").forEach(b => {
    const active = b.dataset.peopleTab === name;
    b.classList.toggle("active", active);
    b.setAttribute("aria-selected", active ? "true" : "false");
  });
  $$(".people-panel").forEach(p => p.classList.toggle("active", p.id === `panel-${name}`));
  if (name === "applicants") renderApplicants();
  else if (name === "teachers") renderTeachers();
  else renderReferrers();
}
$$(".people-tab").forEach(b => b.addEventListener("click", () => setPeopleTab(b.dataset.peopleTab)));

function applicantFormHtml(a = {}) {
  return `
    <div class="modal-head"><h3>${a.id ? "ویرایش متقاضی" : "متقاضی جدید"}</h3>
      <button class="icon-btn" id="mClose">✕</button></div>
    <form id="applicantForm">
      <div class="row">
        <div><label>نام</label><input name="firstName" required value="${esc(a.firstName||"")}"></div>
        <div><label>نام خانوادگی</label><input name="lastName" value="${esc(a.lastName||"")}"></div>
      </div>
      <div class="row">
        <div><label>شماره تماس</label><input name="phone" value="${esc(a.phone||"")}"></div>
        <div><label>ایمیل</label><input name="email" type="email" value="${esc(a.email||"")}"></div>
      </div>
      <div class="row">
        <div><label>شماره پاسپورت</label><input name="passportNumber" value="${esc(a.passportNumber||"")}"></div>
        <div><label>تاریخ تولد (شمسی)</label><input name="birthDateJ" placeholder="1370/05/12" value="${esc(a.birthDateJ||"")}"></div>
      </div>
      <label>ملیت / کشور</label><input name="nationality" value="${esc(a.nationality||"")}">
      <label>توضیحات</label><textarea name="notes">${esc(a.notes||"")}</textarea>
      <div class="modal-actions">
        <button type="submit" class="btn block">ذخیره</button>
        ${a.id ? `<button type="button" class="btn danger block" id="mDelete">حذف</button>` : ""}
      </div>
    </form>`;
}

function openApplicantForm(id) {
  const a = id ? getApplicant(id) || {} : {};
  openModal(applicantFormHtml(a));
  $("#mClose").onclick = closeModal;
  $("#applicantForm").addEventListener("submit", async (ev) => {
    ev.preventDefault();
    const fd = new FormData(ev.target);
    const rec = {
      id: a.id || uid(),
      firstName: fd.get("firstName").trim(),
      lastName: (fd.get("lastName") || "").trim(),
      phone: fd.get("phone").trim(),
      email: fd.get("email").trim(),
      passportNumber: fd.get("passportNumber").trim(),
      birthDateJ: fd.get("birthDateJ").trim(),
      nationality: fd.get("nationality").trim(),
      notes: fd.get("notes").trim(),
      createdAt: a.createdAt || nowIso(),
    };
    await storePut("applicants", rec);
    await refreshAll();
    closeModal(); toast("ذخیره شد"); renderApplicants(); renderDashboard();
  });
  if (a.id) $("#mDelete").onclick = () => {
    const hasExams = DATA.exams.some(e => e.applicantId === a.id);
    confirmDialog(hasExams ? "این متقاضی دارای آزمون ثبت‌شده است. حذف پرونده او باعث قطع ارتباط با آزمون‌ها نمی‌شود اما توصیه نمی‌شود. حذف شود؟" : "پرونده این متقاضی حذف شود؟", async () => {
      await storeDelete("applicants", a.id);
      await refreshAll(); closeModal(); toast("حذف شد"); renderApplicants(); renderDashboard();
    });
  };
}

function openApplicantModal(id) {
  const a = getApplicant(id);
  if (!a) return;
  const exams = DATA.exams.filter(e => e.applicantId === id).sort((x,y) => examSortKey(y) - examSortKey(x));
  openModal(`
    <div class="modal-head"><h3>${esc(applicantName(a))}</h3><button class="icon-btn" id="mClose">✕</button></div>
    <p class="small muted">
      ${a.phone ? "📞 " + esc(a.phone) + "<br>" : ""}
      ${a.email ? "✉️ " + esc(a.email) + "<br>" : ""}
      ${a.passportNumber ? "🛂 " + esc(a.passportNumber) + "<br>" : ""}
      ${a.nationality ? "🌍 " + esc(a.nationality) + "<br>" : ""}
      ${a.notes ? "📝 " + esc(a.notes) : ""}
    </p>
    <div class="modal-actions">
      <button class="btn secondary block" id="mEdit">ویرایش</button>
      <button class="btn block" id="mAddExam">+ آزمون جدید</button>
    </div>
    <h3 style="margin-top:16px">سوابق آزمون</h3>
    ${exams.length ? renderExamRows(exams) : `<div class="empty">آزمونی ثبت نشده.</div>`}
  `);
  $("#mClose").onclick = closeModal;
  $("#mEdit").onclick = () => openApplicantForm(id);
  $("#mAddExam").onclick = () => openExamForm(null, id);
}

/* ---------------------------------------------------------------------- */
/* EXAMS                                                                   */
/* ---------------------------------------------------------------------- */
function renderExams() {
  const statusF = $("#examStatusFilter").value;
  let list = [...DATA.exams];
  if (statusF) list = list.filter(e => e.status === statusF);
  list.sort((a, b) => examSortKey(b) - examSortKey(a));
  $("#examsList").innerHTML = list.length ? renderExamRows(list) : `<div class="empty">آزمونی یافت نشد.</div>`;
}
$("#examStatusFilter").addEventListener("change", renderExams);

function applicantOptions(selectedId) {
  return DATA.applicants
    .sort((a,b)=>applicantName(a).localeCompare(applicantName(b),"fa"))
    .map(a => `<option value="${a.id}" ${a.id===selectedId?"selected":""}>${esc(applicantName(a))}</option>`).join("");
}
function teacherOptions(selectedId) {
  return `<option value="">—</option>` + DATA.teachers.map(t => `<option value="${t.id}" ${t.id===selectedId?"selected":""}>${esc(t.name)}</option>`).join("");
}
function referrerOptions(selectedId) {
  return `<option value="">—</option>` + DATA.referrers.map(r => `<option value="${r.id}" ${r.id===selectedId?"selected":""}>${esc(r.name)}</option>`).join("");
}
function examTypeOptions(selected) {
  const sel = (selected || "").trim().toLowerCase();
  let names = DATA.examTypes.map(t => t.name);
  // keep a value that is not in the list (e.g. old data) instead of silently changing it
  if (selected && !names.some(n => n.trim().toLowerCase() === sel)) names = [selected, ...names];
  return names.map(n => `<option value="${esc(n)}" ${n.trim().toLowerCase()===sel?"selected":""}>${esc(n)}</option>`).join("");
}

function examFormHtml(e = {}, fixedApplicantId) {
  const dateStr = e.examDateJ ? Jalali.format(...e.examDateJ) : "";
  return `
    <div class="modal-head"><h3>${e.id ? "ویرایش آزمون" : "آزمون جدید"}</h3><button class="icon-btn" id="mClose">✕</button></div>
    <form id="examForm">
      <label>متقاضی</label>
      ${fixedApplicantId
        ? `<input type="hidden" name="applicantId" value="${fixedApplicantId}"><input disabled value="${esc(applicantName(getApplicant(fixedApplicantId)))}">`
        : `<select name="applicantId" required><option value="">انتخاب کنید</option>${applicantOptions(e.applicantId)}</select>`}
      <div class="row">
        <div><label>نوع آزمون</label><select name="examType">${examTypeOptions(e.examType)}</select></div>
        <div><label>وضعیت</label><select name="status">
          ${Object.entries(STATUS_LABEL).map(([k,v])=>`<option value="${k}" ${e.status===k?"selected":""}>${v}</option>`).join("")}
        </select></div>
      </div>
      <div class="row">
        <div><label>تاریخ آزمون (شمسی)</label><input name="examDateJStr" placeholder="1403/08/20" value="${esc(dateStr)}"></div>
        <div><label>ساعت</label><input name="examTime" type="time" value="${esc(e.examTime||"")}"></div>
      </div>
      <div class="row">
        <div><label>استاد</label><select name="teacherId">${teacherOptions(e.teacherId)}</select></div>
        <div><label>معرف</label><select name="referrerId">${referrerOptions(e.referrerId)}</select></div>
      </div>
      <label>مبلغ دریافتی (تومان)</label><input name="amount" type="number" min="0" value="${e.amount||""}">
      <label>نتیجه / نمره</label><input name="result" value="${esc(e.result||"")}">
      <label>توضیحات</label><textarea name="notes">${esc(e.notes||"")}</textarea>
      <div class="modal-actions">
        <button type="submit" class="btn block">ذخیره</button>
        ${e.id ? `<button type="button" class="btn danger block" id="mDelete">حذف</button>` : ""}
      </div>
    </form>`;
}

function openExamForm(id, fixedApplicantId) {
  const e = id ? DATA.exams.find(x => x.id === id) || {} : {};
  openModal(examFormHtml(e, fixedApplicantId));
  $("#mClose").onclick = closeModal;
  $("#examForm").addEventListener("submit", async (ev) => {
    ev.preventDefault();
    const fd = new FormData(ev.target);
    const dateStrRaw = fd.get("examDateJStr").trim();
    const parsed = Jalali.parse(dateStrRaw);
    const prevDate = e.examDateJ ? Jalali.format(...e.examDateJ) : null;
    const rec = {
      id: e.id || uid(),
      applicantId: fixedApplicantId || fd.get("applicantId"),
      examType: fd.get("examType"),
      status: fd.get("status"),
      examDateJ: parsed || null,
      examTime: fd.get("examTime") || "",
      teacherId: fd.get("teacherId") || "",
      referrerId: fd.get("referrerId") || "",
      amount: Number(fd.get("amount")) || 0,
      result: fd.get("result").trim(),
      notes: fd.get("notes").trim(),
      history: e.history || [],
      createdAt: e.createdAt || nowIso(),
    };
    if (prevDate && parsed && Jalali.format(...parsed) !== prevDate) {
      rec.history.push({ at: nowIso(), previousDate: prevDate });
      if (rec.status === "upcoming") rec.status = "rescheduled";
    }
    if (!rec.applicantId) { toast("لطفاً متقاضی را انتخاب کنید"); return; }
    await storePut("exams", rec);
    await refreshAll();
    closeModal(); toast("ذخیره شد");
    renderExams(); renderDashboard();
  });
  if (e.id) $("#mDelete").onclick = () => {
    confirmDialog("این آزمون لغو شود یا کاملاً حذف شود؟ (پیشنهاد: به‌جای حذف، وضعیت را «لغوشده» کنید)", async () => {
      await storeDelete("exams", e.id);
      await refreshAll(); closeModal(); toast("حذف شد"); renderExams(); renderDashboard();
    });
  };
}

function openExamModal(id) {
  openExamForm(id);
}

/* ---------------------------------------------------------------------- */
/* CALENDAR                                                                */
/* ---------------------------------------------------------------------- */
let calJY, calJM, calSelectedDay;
function renderCalendar() {
  if (!calJY) {
    const t = Jalali.fromDate(new Date());
    calJY = t[0]; calJM = t[1]; calSelectedDay = t[2];
  }
  const todayJ = Jalali.fromDate(new Date());
  const daysInMonth = Jalali.jalMonthLength(calJY, calJM);
  if (!calSelectedDay) calSelectedDay = Math.min(todayJ[2], daysInMonth);
  else if (calSelectedDay > daysInMonth) calSelectedDay = daysInMonth;
  $("#calTitle").textContent = `${Jalali.MONTH_NAMES[calJM-1]} ${String(calJY).replace(/\d/g, d => "۰۱۲۳۴۵۶۷۸۹"[d])}`;
  const firstOfMonthDate = Jalali.toDate(calJY, calJM, 1);
  const leading = (firstOfMonthDate.getDay() + 1) % 7;
  const examsByDay = {};
  DATA.exams.forEach(e => {
    if (e.examDateJ && e.examDateJ[0] === calJY && e.examDateJ[1] === calJM) {
      const d = e.examDateJ[2];
      (examsByDay[d] = examsByDay[d] || []).push(e);
    }
  });
  const monthTotal = Object.values(examsByDay).reduce((sum, day) => sum + day.length, 0);
  $("#monthExamCount").textContent = `${monthTotal.toLocaleString("fa-IR")} آزمون`;
  const dows = ["ش","ی","د","س","چ","پ","ج"];
  let html = dows.map(d => `<div class="dow">${d}</div>`).join("");
  for (let i = 0; i < leading; i++) html += `<div class="day empty"></div>`;
  for (let d = 1; d <= daysInMonth; d++) {
    const isToday = todayJ[0]===calJY && todayJ[1]===calJM && todayJ[2]===d;
    const list = examsByDay[d] || [];
    const hasAny = list.length > 0;
    const kinds = list.map(e => STATUS_KIND[e.status] || "open");
    const colorClass = !hasAny ? "" : kinds.every(k => k === "bad") ? "bad" : kinds.includes("open") ? "open" : kinds.includes("hold") ? "hold" : "past";
    html += `<button type="button" class="day ${isToday?"today":""} ${hasAny?"has-exam":""} ${calSelectedDay===d?"selected":""}" data-day="${d}" aria-label="${d} ${list.length ? "، " + list.length + " آزمون" : ""}">
      <span>${d.toLocaleString("fa-IR")}</span>${hasAny ? `<span class="dot ${colorClass}" style="background:${{open:"#0f766e",past:"#22c55e",hold:"#f59e0b",bad:"#e11d48"}[colorClass]}"></span>` : ""}
    </button>`;
  }
  $("#calGrid").innerHTML = html;
  $$(".day[data-day]", $("#calGrid")).forEach(el => el.addEventListener("click", () => {
    calSelectedDay = Number(el.dataset.day);
    renderCalendar();
  }));
  const selected = examsByDay[calSelectedDay] || [];
  const dateLabel = `${calSelectedDay.toLocaleString("fa-IR")} ${Jalali.MONTH_NAMES[calJM-1]} ${String(calJY).replace(/\d/g, d => "۰۱۲۳۴۵۶۷۸۹"[d])}`;
  $("#calSelectedTitle").textContent = dateLabel;
  $("#calSelectedCount").textContent = `${selected.length.toLocaleString("fa-IR")} آزمون`;
  selected.sort((x,y) => examSortKey(x)-examSortKey(y));
  $("#calDayExams").innerHTML = selected.length ? renderExamRows(selected) : `<div class="empty">برای این روز آزمونی ثبت نشده است.</div>`;
}
$("#calPrev").addEventListener("click", () => { calJM--; if (calJM<1){calJM=12;calJY--;} renderCalendar(); });
$("#calNext").addEventListener("click", () => { calJM++; if (calJM>12){calJM=1;calJY++;} renderCalendar(); });

/* ---------------------------------------------------------------------- */
/* REPORTS                                                                 */
/* ---------------------------------------------------------------------- */
function renderReportsSetup() {
  $("#repType").innerHTML = `<option value="">همه</option>` + DATA.examTypes.map(t=>`<option value="${esc(t.name)}">${esc(t.name)}</option>`).join("");
  $("#repTeacher").innerHTML = `<option value="">همه</option>` + DATA.teachers.map(t=>`<option value="${t.id}">${esc(t.name)}</option>`).join("");
  $("#repReferrer").innerHTML = `<option value="">همه</option>` + DATA.referrers.map(r=>`<option value="${r.id}">${esc(r.name)}</option>`).join("");
}
function jalaliValue(j) { return j ? j[0]*10000+j[1]*100+j[2] : null; }
let lastReportRows = [];
function runReport() {
  const fromP = Jalali.parse($("#repFrom").value);
  const toP = Jalali.parse($("#repTo").value);
  const type = $("#repType").value, status = $("#repStatus").value, teacherId = $("#repTeacher").value, referrerId = $("#repReferrer").value;
  let list = DATA.exams.filter(e => {
    if (type && (e.examType || "").trim().toLowerCase() !== type.trim().toLowerCase()) return false;
    if (status && e.status !== status) return false;
    if (teacherId && e.teacherId !== teacherId) return false;
    if (referrerId && e.referrerId !== referrerId) return false;
    if (fromP && (!e.examDateJ || jalaliValue(e.examDateJ) < jalaliValue(fromP))) return false;
    if (toP && (!e.examDateJ || jalaliValue(e.examDateJ) > jalaliValue(toP))) return false;
    return true;
  }).sort((a,b)=>examSortKey(a)-examSortKey(b));
  lastReportRows = list;
  const totalAmount = list.reduce((s,e)=>s+(Number(e.amount)||0),0);
  $("#reportSummary").textContent = `تعداد: ${list.length} | مجموع دریافتی: ${fmtMoney(totalAmount)} تومان`;
  $("#reportTable thead").innerHTML = `<tr><th>متقاضی</th><th>نوع</th><th>تاریخ</th><th>وضعیت</th><th>استاد</th><th>معرف</th><th>مبلغ</th><th>نتیجه</th></tr>`;
  $("#reportTable tbody").innerHTML = list.map(e => `<tr>
    <td>${esc(applicantName(getApplicant(e.applicantId)))}</td>
    <td>${esc(e.examType||"")}</td>
    <td>${e.examDateJ?esc(Jalali.format(...e.examDateJ)):""}</td>
    <td>${STATUS_LABEL[e.status]||""}</td>
    <td>${esc(getTeacher(e.teacherId)?.name||"")}</td>
    <td>${esc(getReferrer(e.referrerId)?.name||"")}</td>
    <td>${fmtMoney(e.amount)}</td>
    <td>${esc(e.result||"")}</td>
  </tr>`).join("");
}
$("#btnRunReport").addEventListener("click", runReport);

function loadSheetJs() {
  return new Promise((resolve, reject) => {
    if (window.XLSX) return resolve();
    const s = document.createElement("script");
    s.src = "https://cdn.jsdelivr.net/npm/xlsx@0.18.5/dist/xlsx.full.min.js";
    s.onload = () => resolve();
    s.onerror = () => reject(new Error("offline"));
    document.head.appendChild(s);
  });
}
$("#btnExportExcel").addEventListener("click", async () => {
  if (!lastReportRows.length) { toast("ابتدا گزارش را نمایش دهید"); return; }
  const rows = lastReportRows.map(e => ({
    "متقاضی": applicantName(getApplicant(e.applicantId)),
    "نوع آزمون": e.examType || "",
    "تاریخ": e.examDateJ ? Jalali.format(...e.examDateJ) : "",
    "وضعیت": STATUS_LABEL[e.status] || "",
    "استاد": getTeacher(e.teacherId)?.name || "",
    "معرف": getReferrer(e.referrerId)?.name || "",
    "مبلغ": e.amount || 0,
    "نتیجه": e.result || "",
    "توضیحات": e.notes || "",
  }));
  try {
    await loadSheetJs();
    const ws = XLSX.utils.json_to_sheet(rows);
    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, ws, "گزارش");
    XLSX.writeFile(wb, `examhub-report-${Date.now()}.xlsx`);
  } catch {
    // Offline fallback: CSV with UTF-8 BOM so Persian text opens correctly in Excel
    const headers = Object.keys(rows[0]);
    const csv = "﻿" + [headers.join(",")].concat(
      rows.map(r => headers.map(h => `"${String(r[h]).replace(/"/g,'""')}"`).join(","))
    ).join("\r\n");
    const blob = new Blob([csv], { type: "text/csv;charset=utf-8" });
    const a = document.createElement("a"); a.href = URL.createObjectURL(blob);
    a.download = `examhub-report-${Date.now()}.csv`; a.click();
    toast("اتصال اینترنت نبود؛ خروجی CSV ساخته شد");
  }
});

/* ---------------------------------------------------------------------- */
/* TELEGRAM MESSAGE TEMPLATES                                             */
/* ---------------------------------------------------------------------- */
function renderMessagesSetup() {
  const list = [...DATA.exams].sort((a,b)=>examSortKey(b)-examSortKey(a));
  $("#msgExamSelect").innerHTML = `<option value="">— یک آزمون انتخاب کنید —</option>` + list.map(e => {
    const a = getApplicant(e.applicantId);
    const dateStr = e.examDateJ ? Jalali.format(...e.examDateJ) : "بدون تاریخ";
    return `<option value="${e.id}">${esc(applicantName(a))} — ${esc(e.examType||"")} — ${esc(dateStr)}</option>`;
  }).join("");
  $("#msgArea").style.display = "none";
}
$("#msgExamSelect").addEventListener("change", (ev) => {
  const id = ev.target.value;
  if (!id) { $("#msgArea").style.display = "none"; return; }
  const e = DATA.exams.find(x => x.id === id);
  const a = getApplicant(e.applicantId);
  const t = getTeacher(e.teacherId);
  const r = getReferrer(e.referrerId);
  const dateStr = e.examDateJ ? Jalali.format(...e.examDateJ) : "هنوز مشخص نشده";
  const timeStr = e.examTime || "";
  const resultLine = e.result ? `\nنتیجه: ${e.result}` : "";

  $("#msgApplicant").value =
`سلام ${applicantName(a)} عزیز
آزمون ${e.examType || ""} شما:
تاریخ: ${dateStr} ${timeStr ? "ساعت " + timeStr : ""}
وضعیت: ${STATUS_LABEL[e.status] || ""}${resultLine}
موفق باشید 🌟`;

  $("#msgTeacher").value = t ?
`سلام ${t.name} عزیز
متقاضی: ${applicantName(a)}
نوع آزمون: ${e.examType || ""}
تاریخ: ${dateStr} ${timeStr ? "ساعت " + timeStr : ""}
${e.notes ? "توضیحات: " + e.notes : ""}` : "استادی برای این آزمون انتخاب نشده است.";

  $("#msgReferrer").value = r ?
`سلام ${r.name} عزیز
متقاضی معرفی‌شده: ${applicantName(a)}
نوع آزمون: ${e.examType || ""}
تاریخ: ${dateStr}` : "معرفی برای این آزمون ثبت نشده است.";

  $("#msgArea").style.display = "block";
});
document.addEventListener("click", (e) => {
  const btn = e.target.closest("[data-copy]");
  if (!btn) return;
  const ta = $("#" + btn.dataset.copy);
  navigator.clipboard?.writeText(ta.value).then(() => toast("کپی شد")).catch(() => {
    ta.select(); document.execCommand("copy"); toast("کپی شد");
  });
});

/* ---------------------------------------------------------------------- */
/* SETTINGS: teachers / referrers / exam types / backup                   */
/* ---------------------------------------------------------------------- */
function renderSettings() {
  $("#examTypesList").innerHTML = DATA.examTypes.length ? DATA.examTypes.map(t => `
    <div class="list-item">
      <div class="main"><div class="title">${esc(t.name)}</div></div>
      <div><button class="btn ghost sm" data-del-examtype="${t.id}">حذف</button></div>
    </div>`).join("") : `<div class="empty">نوع آزمونی ثبت نشده.</div>`;

  renderSyncCard();
  metaGet("lastBackupAt").then(v => {
    $("#lastBackupInfo").textContent = v ? `آخرین پشتیبان: ${new Date(v).toLocaleString("fa-IR")}` : "هنوز پشتیبانی تهیه نشده.";
  });
}
document.addEventListener("click", (e) => {
  const et = e.target.closest("[data-edit-teacher]");
  if (et) return openPersonForm("teachers", et.dataset.editTeacher);
  const er = e.target.closest("[data-edit-referrer]");
  if (er) return openPersonForm("referrers", er.dataset.editReferrer);
  const dt = e.target.closest("[data-del-examtype]");
  if (dt) {
    const id = dt.dataset.delExamtype;
    const name = (DATA.examTypes.find(t => t.id === id)?.name || "").trim().toLowerCase();
    const inUse = DATA.exams.some(x => (x.examType || "").trim().toLowerCase() === name);
    confirmDialog(inUse ? "این نوع آزمون در حال استفاده است؛ حذف شود؟ (آزمون‌های موجود تغییر نمی‌کنند)" : "این نوع آزمون حذف شود؟", async () => {
      // remove every entry with this name (older data may hold duplicates)
      for (const t of await storeAll("examTypes")) {
        if (!t.deleted && (t.name || "").trim().toLowerCase() === name) await storeDelete("examTypes", t.id);
      }
      await refreshAll(); renderSettings();
    });
  }
});
function openPersonForm(store, id) {
  const label = store === "teachers" ? "استاد" : "معرف";
  const rec = id ? DATA[store].find(x => x.id === id) : {};
  openModal(`
    <div class="modal-head"><h3>${id ? "ویرایش " + label : label + " جدید"}</h3><button class="icon-btn" id="mClose">✕</button></div>
    <form id="personForm">
      <label>نام</label><input name="name" required value="${esc(rec?.name||"")}">
      <label>شماره تماس</label><input name="phone" value="${esc(rec?.phone||"")}">
      <label>توضیحات</label><textarea name="notes">${esc(rec?.notes||"")}</textarea>
      <div class="modal-actions">
        <button type="submit" class="btn block">ذخیره</button>
        ${id ? `<button type="button" class="btn danger block" id="mDelete">حذف</button>` : ""}
      </div>
    </form>`);
  $("#mClose").onclick = closeModal;
  $("#personForm").addEventListener("submit", async (ev) => {
    ev.preventDefault();
    const fd = new FormData(ev.target);
    await storePut(store, { id: rec?.id || uid(), name: fd.get("name").trim(), phone: fd.get("phone").trim(), notes: fd.get("notes").trim() });
    await refreshAll(); closeModal(); toast("ذخیره شد"); renderSettings(); renderTeachers(); renderReferrers();
  });
  if (id) $("#mDelete").onclick = () => confirmDialog(`این ${label} حذف شود؟`, async () => {
    await storeDelete(store, id); await refreshAll(); closeModal(); renderSettings(); renderTeachers(); renderReferrers();
  });
}

$("#btnAddExamType").addEventListener("click", () => {
  openModal(`
    <div class="modal-head"><h3>نوع آزمون جدید</h3><button class="icon-btn" id="mClose">✕</button></div>
    <form id="typeForm"><label>نام</label><input name="name" required>
    <div class="modal-actions"><button type="submit" class="btn block">ذخیره</button></div></form>`);
  $("#mClose").onclick = closeModal;
  $("#typeForm").addEventListener("submit", async (ev) => {
    ev.preventDefault();
    const name = new FormData(ev.target).get("name").trim();
    if (name) { await storePut("examTypes", { id: typeId(name), name }); await refreshAll(); closeModal(); renderSettings(); }
  });
});

/* Backup / Restore */
async function buildBackupObject() {
  return {
    appVersion: "examhub-stage2",
    exportedAt: nowIso(),
    applicants: (await storeAll("applicants")).filter(isLive),
    exams: (await storeAll("exams")).filter(isLive),
    teachers: (await storeAll("teachers")).filter(isLive),
    referrers: (await storeAll("referrers")).filter(isLive),
    examTypes: (await storeAll("examTypes")).filter(isLive),
  };
}
$("#btnBackup").addEventListener("click", async () => {
  const obj = await buildBackupObject();
  const blob = new Blob([JSON.stringify(obj, null, 2)], { type: "application/json" });
  const a = document.createElement("a");
  a.href = URL.createObjectURL(blob);
  a.download = `examhub-backup-${new Date().toISOString().slice(0,10)}.json`;
  a.click();
  await metaSet("lastBackupAt", Date.now());
  toast("پشتیبان ذخیره شد");
  renderSettings();
});
$("#btnRestore").addEventListener("click", () => $("#fileRestore").click());
$("#fileRestore").addEventListener("change", async (ev) => {
  const file = ev.target.files[0];
  if (!file) return;
  try {
    const text = await file.text();
    const obj = JSON.parse(text);
    if (!obj || !Array.isArray(obj.applicants)) throw new Error("فرمت فایل نامعتبر است");
    confirmDialog(`این فایل شامل ${obj.applicants.length} متقاضی و ${obj.exams?.length||0} آزمون است. اطلاعات فعلی با این فایل ادغام می‌شود (رکوردهای هم‌شناسه به‌روزرسانی می‌شوند). ادامه می‌دهید؟`, async () => {
      for (const key of ["applicants","exams","teachers","referrers","examTypes"]) {
        for (const rec of (obj[key] || [])) await storePut(key, rec);
      }
      await refreshAll();
      toast("بازیابی با موفقیت انجام شد");
      renderSettings(); renderDashboard();
    });
  } catch (err) {
    toast("خطا در خوانش فایل پشتیبان: " + err.message);
  } finally {
    ev.target.value = "";
  }
});

/* ---------------------------------------------------------------------- */
/* SYNC with Google Sheets (through a Google Apps Script web app)         */
/* The phone/PC keeps working offline on IndexedDB; the Sheet is the      */
/* shared database. Rule: the record with the newer updatedAt wins.       */
/* ---------------------------------------------------------------------- */
const SCHEMA = {
  applicants: { sheet: "Applicants", cols: ["id","firstName","lastName","phone","email","passportNumber","birthDateJ","nationality","notes","createdAt","updatedAt","deleted"] },
  exams: { sheet: "Exams", cols: ["id","applicantId","applicantName","examType","status","examDate","examTime","teacherId","teacherName","referrerId","referrerName","amount","result","notes","history","createdAt","updatedAt","deleted"] },
  teachers: { sheet: "Teachers", cols: ["id","name","phone","notes","updatedAt","deleted"] },
  referrers: { sheet: "Referrers", cols: ["id","name","phone","notes","updatedAt","deleted"] },
  examTypes: { sheet: "ExamTypes", cols: ["id","name","updatedAt","deleted"] },
};
const STATUS_KEY_BY_LABEL = Object.fromEntries(Object.entries(STATUS_LABEL).map(([k, v]) => [v, k]));
const EPOCH_ISO = "2000-01-01T00:00:00.000Z";

let syncCfg = null;          // { url, token } — kept in IndexedDB "meta", never in the code/repo
let syncing = false;
let syncTimer = null;
let syncInfo = { state: "off", msg: "", lastOkAt: 0 };

function toRow(store, r, ctx) {
  if (r.deleted) return { id: r.id, deleted: "TRUE", updatedAt: r.updatedAt || "" };
  const tail = { updatedAt: r.updatedAt || "", deleted: "" };
  if (store === "applicants") {
    return { id: r.id, firstName: r.firstName || "", lastName: r.lastName || "", phone: r.phone || "", email: r.email || "",
      passportNumber: r.passportNumber || "", birthDateJ: r.birthDateJ || "", nationality: r.nationality || "",
      notes: r.notes || "", createdAt: r.createdAt || "", ...tail };
  }
  if (store === "exams") {
    return { id: r.id, applicantId: r.applicantId || "", applicantName: applicantName(ctx.applicants.get(r.applicantId)),
      examType: r.examType || "", status: STATUS_LABEL[r.status] || r.status || "",
      examDate: r.examDateJ ? Jalali.format(...r.examDateJ) : "", examTime: r.examTime || "",
      teacherId: r.teacherId || "", teacherName: ctx.teachers.get(r.teacherId)?.name || "",
      referrerId: r.referrerId || "", referrerName: ctx.referrers.get(r.referrerId)?.name || "",
      amount: String(r.amount || 0), result: r.result || "", notes: r.notes || "",
      history: JSON.stringify(r.history || []), createdAt: r.createdAt || "", ...tail };
  }
  if (store === "examTypes") return { id: r.id, name: r.name || "", ...tail };
  return { id: r.id, name: r.name || "", phone: r.phone || "", notes: r.notes || "", ...tail };
}

function fromRow(store, row) {
  const id = String(row.id ?? "").trim();
  if (!id) return null;
  const updatedAt = String(row.updatedAt ?? "");
  if (/^(true|1|yes)$/i.test(String(row.deleted ?? "").trim())) return { id, deleted: true, updatedAt };
  const s = (k) => String(row[k] ?? "");
  if (store === "applicants") {
    return { id, firstName: s("firstName"), lastName: s("lastName"), phone: s("phone"), email: s("email"),
      passportNumber: s("passportNumber"), birthDateJ: s("birthDateJ"), nationality: s("nationality"),
      notes: s("notes"), createdAt: s("createdAt") || updatedAt, updatedAt };
  }
  if (store === "exams") {
    const st = s("status").trim();
    let history = [];
    try { const h = JSON.parse(s("history") || "[]"); if (Array.isArray(h)) history = h; } catch { /* keep [] */ }
    return { id, applicantId: s("applicantId"), examType: s("examType"),
      status: STATUS_KEY_BY_LABEL[st] || STATUS_ALIASES[st] || (STATUS_LABEL[st] ? st : "upcoming"),
      examDateJ: Jalali.parse(s("examDate")), examTime: toLatinDigits(s("examTime")).trim(),
      teacherId: s("teacherId"), referrerId: s("referrerId"),
      amount: Number(toLatinDigits(s("amount")).replace(/[,٬\s]/g, "")) || 0,
      result: s("result"), notes: s("notes"), history, createdAt: s("createdAt") || updatedAt, updatedAt };
  }
  if (store === "examTypes") return { id, name: s("name"), updatedAt };
  return { id, name: s("name"), phone: s("phone"), notes: s("notes"), updatedAt };
}

async function callScript(cfg, payload) {
  const ctl = new AbortController();
  const timer = setTimeout(() => ctl.abort(), 90000);
  try {
    // text/plain avoids a CORS preflight, which Apps Script web apps do not answer
    const res = await fetch(cfg.url, {
      method: "POST", redirect: "follow", signal: ctl.signal,
      headers: { "Content-Type": "text/plain;charset=utf-8" },
      body: JSON.stringify({ ...payload, token: cfg.token }),
    });
    const text = await res.text();
    try { return JSON.parse(text); }
    catch { throw new Error("پاسخ نامعتبر از Google؛ آدرس Web App و دسترسی «Anyone» را بررسی کنید"); }
  } catch (e) {
    if (e.name === "AbortError") throw new Error("زمان انتظار تمام شد");
    if (e instanceof TypeError) throw new Error("عدم اتصال به اینترنت یا دسترسی به Google");
    throw e;
  } finally { clearTimeout(timer); }
}

async function countPending() {
  const lastPushAt = (await metaGet("lastPushAt")) || "";
  let n = 0;
  for (const st of SYNC_STORES) for (const r of await storeAll(st)) if (!r.updatedAt || r.updatedAt > lastPushAt) n++;
  return n;
}

const SYNC_ICON = { off: "☁️", ok: "✅", pending: "⏳", syncing: "🔄", error: "⚠️" };
const SYNC_TEXT = { off: "همگام‌سازی غیرفعال است", ok: "همگام‌سازی شده", pending: "در انتظار همگام‌سازی", syncing: "در حال همگام‌سازی…", error: "خطای همگام‌سازی" };

function setSyncState(state, msg = "") {
  syncInfo.state = state; syncInfo.msg = msg;
  if (state === "ok") syncInfo.lastOkAt = Date.now();
  const b = $("#btnSync");
  b.textContent = SYNC_ICON[state];
  b.title = SYNC_TEXT[state] + (msg ? " — " + msg : "");
  b.style.opacity = state === "off" ? ".5" : "1";
  renderSyncStatusText();
}
async function renderSyncStatusText() {
  const el = $("#syncStatusText");
  if (!el) return;
  let t = SYNC_TEXT[syncInfo.state];
  if (syncInfo.msg) t += " — " + syncInfo.msg;
  if (syncCfg) {
    const last = await metaGet("lastSyncAt");
    const pending = await countPending();
    if (last) t += `\nآخرین همگام‌سازی موفق: ${new Date(last).toLocaleString("fa-IR")}`;
    t += `\nرکورد در انتظار ارسال: ${pending}`;
  }
  el.textContent = t;
  el.style.whiteSpace = "pre-line";
}

function onLocalChange() {
  if (!syncCfg) return;
  setSyncState("pending");
  clearTimeout(syncTimer);
  syncTimer = setTimeout(() => syncNow(), 2500);
}

async function syncNow() {
  if (!syncCfg || syncing) return;
  syncing = true;
  clearTimeout(syncTimer);
  setSyncState("syncing");
  const startedAt = nowIso();
  try {
    const lastPushAt = (await metaGet("lastPushAt")) || "";
    const since = (await metaGet("lastServerTime")) || 0;

    // load everything once; give old records without updatedAt a (very old) one so they get pushed
    const all = {};
    for (const st of SYNC_STORES) {
      all[st] = await storeAll(st);
      for (const r of all[st]) if (!r.updatedAt) { r.updatedAt = r.createdAt || EPOCH_ISO; await storePutRaw(st, r); }
    }
    const ctx = {
      applicants: new Map(all.applicants.filter(isLive).map(r => [r.id, r])),
      teachers: new Map(all.teachers.filter(isLive).map(r => [r.id, r])),
      referrers: new Map(all.referrers.filter(isLive).map(r => [r.id, r])),
    };
    const tables = {};
    for (const st of SYNC_STORES) {
      tables[SCHEMA[st].sheet] = {
        cols: SCHEMA[st].cols,
        rows: all[st].filter(r => r.updatedAt > lastPushAt).map(r => toRow(st, r, ctx)),
      };
    }

    const res = await callScript(syncCfg, { action: "sync", since, tables });
    if (!res.ok) throw new Error(res.error === "unauthorized" ? "توکن اشتباه است" : (res.error || "خطای ناشناخته"));

    // apply what the sheet sent back (only if newer than our copy)
    for (const st of SYNC_STORES) {
      const incoming = res.tables?.[SCHEMA[st].sheet]?.rows || [];
      const local = new Map(all[st].map(r => [r.id, r]));
      for (const row of incoming) {
        const rec = fromRow(st, row);
        if (!rec) continue;
        const mine = local.get(rec.id);
        if (!mine || (rec.updatedAt || "") > (mine.updatedAt || "")) await storePutRaw(st, rec);
      }
    }
    await metaSet("lastPushAt", startedAt);
    await metaSet("lastServerTime", res.serverTime);
    await metaSet("lastSyncAt", Date.now());
    await refreshAll();
    if (DATA.examTypes.length === 0) await seedDefaultsIfEmpty();
    syncing = false;
    setSyncState((await countPending()) > 0 ? "pending" : "ok");
    render($(".tab-btn.active").dataset.view);
  } catch (e) {
    syncing = false;
    setSyncState("error", e.message);
  }
}

async function setSyncConfig(cfg) {
  syncCfg = cfg;
  if (cfg) await metaSet("syncCfg", cfg);
  else await storeDelete("meta", "syncCfg");
}

function renderSyncCard() {
  const urlEl = $("#syncUrl"), tokEl = $("#syncToken");
  if (!urlEl) return;
  urlEl.value = syncCfg?.url || "";
  tokEl.value = "";
  tokEl.placeholder = syncCfg ? "•••••• (ذخیره شده؛ برای تغییر دوباره وارد کنید)" : "توکن مخفی";
  renderSyncStatusText();
}

$("#btnSyncSave").addEventListener("click", async () => {
  const url = $("#syncUrl").value.trim();
  const token = $("#syncToken").value.trim() || syncCfg?.token || "";
  if (!/^https:\/\/script\.google(usercontent)?\.com\//.test(url)) { toast("آدرس باید با https://script.google.com/ شروع شود"); return; }
  if (!token) { toast("توکن را وارد کنید"); return; }
  setSyncState("syncing", "تست اتصال…");
  try {
    const res = await callScript({ url, token }, { action: "ping" });
    if (!res.ok) throw new Error(res.error === "unauthorized" ? "توکن اشتباه است" : (res.error || "خطای ناشناخته"));
  } catch (e) {
    setSyncState(syncCfg ? "pending" : "off", e.message);
    toast("اتصال ناموفق: " + e.message);
    return;
  }
  await setSyncConfig({ url, token });
  toast("اتصال برقرار شد؛ در حال همگام‌سازی…");
  await syncNow();
});
$("#btnSyncNow").addEventListener("click", () => { if (!syncCfg) { toast("ابتدا آدرس و توکن را ذخیره کنید"); return; } syncNow(); });
$("#btnSyncOff").addEventListener("click", () => {
  if (!syncCfg) return;
  confirmDialog("اتصال به Google Sheets قطع شود؟ اطلاعات روی همین دستگاه و داخل شیت حذف نمی‌شوند.", async () => {
    await setSyncConfig(null);
    await metaSet("lastPushAt", ""); await metaSet("lastServerTime", 0);
    setSyncState("off"); renderSyncCard();
  });
});
$("#btnSync").addEventListener("click", () => {
  if (!syncCfg) { showView("settings"); toast("برای همگام‌سازی، آدرس و توکن را در تنظیمات وارد کنید"); return; }
  syncNow();
});
$("#overdueNote").addEventListener("click", () => { $("#examStatusFilter").value = "upcoming"; showView("exams"); });
window.addEventListener("online", () => syncNow());
document.addEventListener("visibilitychange", () => {
  if (document.visibilityState === "visible" && syncCfg && Date.now() - syncInfo.lastOkAt > 60000) syncNow();
});

/* ---------------------------------------------------------------------- */
/* FAB (context-aware "add" button)                                       */
/* ---------------------------------------------------------------------- */
$("#btnFab").addEventListener("click", () => {
  const active = $(".tab-btn.active")?.dataset.view;
  if (active === "people") {
    if (activePeopleTab === "teachers") openPersonForm("teachers");
    else if (activePeopleTab === "referrers") openPersonForm("referrers");
    else openApplicantForm(null);
  } else if (active === "exams") openExamForm(null);
  else openApplicantForm(null);
});

/* ---------------------------------------------------------------------- */
/* Boot                                                                    */
/* ---------------------------------------------------------------------- */
(async function boot() {
  await openDb();
  await refreshAll();
  syncCfg = await metaGet("syncCfg");
  // with sync on, a new device first receives the shared exam types instead of inventing its own
  if (!syncCfg) await seedDefaultsIfEmpty();
  renderDashboard();
  if (syncCfg) {
    setSyncState((await countPending()) > 0 ? "pending" : "ok");
    syncNow();
  } else {
    setSyncState("off");
  }
  if ("serviceWorker" in navigator) {
    navigator.serviceWorker.register("./sw.js").catch(() => {});
  }
})();

})();
