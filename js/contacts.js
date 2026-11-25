/* ExamHub - phone contacts: Contact Picker API (Chrome on Android) or a vCard (.vcf) file (iPhone and others).
   Contact data only fills a form / creates a person; it is not stored anywhere else. */
(function () {
"use strict";
const toLatin = (s) => (window.toLatinDigits ? window.toLatinDigits(String(s ?? "")) : String(s ?? ""));

function supported() {
  return typeof navigator !== "undefined" && "contacts" in navigator && "ContactsManager" in window;
}

// Numbers as 09xxxxxxxxx; mobiles first, then others. +98 / 0098 / 98 prefixes become 0.
function mobilesFrom(numbers) {
  const n = numbers.map((x) => toLatin(x).replace(/[^\d+]/g, ""))
    .map((x) => x.replace(/^(\+98|0098)/, "0").replace(/^98(9\d{9})$/, "0$1")).filter(Boolean);
  const uniq = [...new Set(n)];
  return [...uniq.filter((x) => /^09\d{9}$/.test(x)), ...uniq.filter((x) => !/^09\d{9}$/.test(x))];
}
// "Ali Mohammadi" -> first "Ali", last "Mohammadi"; one word -> first only.
function splitName(name) {
  const w = String(name || "").trim().split(/\s+/).filter(Boolean);
  return w.length < 2 ? { first: w[0] || "", last: "" } : { first: w[0], last: w.slice(1).join(" ") };
}
function fromPicker(c) {
  const name = (c.name || []).find((v) => v && v.trim())?.trim() || "";
  return { name, ...splitName(name), mobiles: mobilesFrom(c.tel || []) };
}
async function pickOne() {
  const [c] = (await navigator.contacts.select(["name", "tel"], { multiple: false })) || [];
  return c ? fromPicker(c) : null;
}
async function pickMany() {
  const list = (await navigator.contacts.select(["name", "tel"], { multiple: true })) || [];
  return list.map(fromPicker).filter((c) => c.name || c.mobiles.length);
}
function parseVCard(text) {
  const unfolded = String(text || "").replace(/\r\n/g, "\n").replace(/\n[ \t]/g, "");
  const card = unfolded.split(/BEGIN:VCARD/i)[1];
  if (!card) return null;
  let fn = "", structured = null; const tels = [];
  for (const line of card.split("\n")) {
    const i = line.indexOf(":"); if (i < 0) continue;
    const key = line.slice(0, i).split(";")[0].split(".").pop().toUpperCase();
    const val = line.slice(i + 1).trim();
    if (key === "FN" && !fn) fn = val;
    else if (key === "N" && !structured) { const [last = "", first = ""] = val.split(";"); structured = { first: first.trim(), last: last.trim() }; }
    else if (key === "TEL") tels.push(val);
  }
  const name = fn || [structured?.first, structured?.last].filter(Boolean).join(" ");
  const parts = structured && (structured.first || structured.last) ? structured : splitName(name);
  return name || tels.length ? { name, ...parts, mobiles: mobilesFrom(tels) } : null;
}
function parseVCards(text) {
  return String(text || "").split(/(?=BEGIN:VCARD)/i).map(parseVCard).filter(Boolean);
}
async function readVcfFile(file) { return parseVCards(await file.text()); }

window.PhoneContacts = { supported, pickOne, pickMany, parseVCards, readVcfFile, mobilesFrom, splitName };
})();
