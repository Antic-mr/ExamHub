/* ExamHub - Jalali calendar + digit helpers (pure functions, no DOM) */
(function () {
"use strict";
function toLatinDigits(s) {
  return String(s ?? "")
    .replace(/[۰-۹]/g, d => "۰۱۲۳۴۵۶۷۸۹".indexOf(d))
    .replace(/[٠-٩]/g, d => "٠١٢٣٤٥٦٧٨٩".indexOf(d));
}

/* ---------------------------------------------------------------------- */
/* Jalali (Persian) calendar conversion — standard astronomical algorithm */
/* ---------------------------------------------------------------------- */
const Jalali = (() => {
  // Standard, widely-verified Jalali<->Gregorian conversion algorithm
  // (the same one used by the jalaali-js library), based on Birashk's
  // 33-year leap-year break-point table. Pure arithmetic, no network needed.
  function div(a, b) { return ~~(a / b); }
  function mod(a, b) { return a - ~~(a / b) * b; }

  function g2d(gy, gm, gd) {
    // Fliegel & Van Flandern (1968) Gregorian -> Julian Day Number
    const a = Math.trunc((gm - 14) / 12);
    return Math.floor((1461 * (gy + 4800 + a)) / 4)
      + Math.floor((367 * (gm - 2 - 12 * a)) / 12)
      - Math.floor((3 * Math.floor((gy + 4900 + a) / 100)) / 4)
      + gd - 32075;
  }
  function d2g(jd) {
    // Fliegel & Van Flandern (1968) Julian Day Number -> Gregorian
    let L = jd + 68569;
    const N = Math.floor((4 * L) / 146097);
    L = L - Math.floor((146097 * N + 3) / 4);
    let I = Math.floor((4000 * (L + 1)) / 1461001);
    L = L - Math.floor((1461 * I) / 4) + 31;
    let J = Math.floor((80 * L) / 2447);
    const K = L - Math.floor((2447 * J) / 80);
    L = Math.floor(J / 11);
    J = J + 2 - 12 * L;
    I = 100 * (N - 49) + I + L;
    return [I, J, K];
  }
  function jalCal(jy) {
    const breaks = [-61,9,38,199,426,686,756,818,1111,1181,1210,1635,2060,2097,2192,2262,2324,2394,2456,3178];
    const bl = breaks.length;
    let gy = jy + 621, leapJ = -14, jp = breaks[0], jm, jump = 0, n, i;
    if (jy < jp || jy >= breaks[bl - 1]) throw new Error("invalid jalali year " + jy);
    for (i = 1; i < bl; i += 1) {
      jm = breaks[i]; jump = jm - jp;
      if (jy < jm) break;
      leapJ = leapJ + div(jump, 33) * 8 + div(mod(jump, 33), 4);
      jp = jm;
    }
    n = jy - jp;
    leapJ = leapJ + div(n, 33) * 8 + div(mod(n, 33) + 3, 4);
    if (mod(jump, 33) === 4 && jump - n === 4) leapJ += 1;
    const leapG = div(gy, 4) - div((div(gy, 100) + 1) * 3, 4) - 150;
    const march = 20 + leapJ - leapG;
    if (jump - n < 6) n = n - jump + div(jump + 4, 33) * 33;
    let leap = mod(mod(n + 1, 33) - 1, 4);
    if (leap === -1) leap = 4;
    return { leap, gy, march };
  }
  function j2d(jy, jm, jd) {
    const r = jalCal(jy);
    return g2d(r.gy, 3, r.march) + (jm - 1) * 31 - div(jm, 7) * (jm - 7) + jd - 1;
  }
  function d2j(jdn) {
    const gy = d2g(jdn)[0];
    let jy = gy - 621;
    const r = jalCal(jy);
    const jdn1f = g2d(r.gy, 3, r.march);
    let jd, jm, k = jdn - jdn1f;
    if (k >= 0) {
      if (k <= 185) { jm = 1 + div(k, 31); jd = mod(k, 31) + 1; return [jy, jm, jd]; }
      k -= 186;
    } else {
      jy -= 1; k += 179;
      if (r.leap === 1) k += 1;
    }
    jm = 7 + div(k, 30);
    jd = mod(k, 30) + 1;
    return [jy, jm, jd];
  }
  function isLeapJalaliYear(jy) { return jalCal(jy).leap === 0; }
  function jalMonthLength(jy, jm) {
    if (jm <= 6) return 31;
    if (jm <= 11) return 30;
    return isLeapJalaliYear(jy) ? 30 : 29;
  }
  function toJalali(gy, gm, gd) { return d2j(g2d(gy, gm, gd)); }
  function toGregorian(jy, jm, jd) { return d2g(j2d(jy, jm, jd)); }
  function fromDate(date) { return toJalali(date.getFullYear(), date.getMonth() + 1, date.getDate()); }
  function toDate(jy, jm, jd) {
    const [gy, gm, gd] = toGregorian(jy, jm, jd);
    return new Date(gy, gm - 1, gd);
  }
  const MONTH_NAMES = ["فروردین","اردیبهشت","خرداد","تیر","مرداد","شهریور","مهر","آبان","آذر","دی","بهمن","اسفند"];
  const WEEKDAYS = ["یک‌شنبه","دوشنبه","سه‌شنبه","چهارشنبه","پنجشنبه","جمعه","شنبه"];
  function format(jy, jm, jd) {
    return `${jy}/${String(jm).padStart(2,"0")}/${String(jd).padStart(2,"0")}`;
  }
  function parse(str) {
    if (!str) return null;
    const m = toLatinDigits(str).trim().match(/^(\d{3,4})[\/\-](\d{1,2})[\/\-](\d{1,2})$/);
    if (!m) return null;
    return [parseInt(m[1],10), parseInt(m[2],10), parseInt(m[3],10)];
  }
  return { toJalali, toGregorian, fromDate, toDate, format, parse, MONTH_NAMES, WEEKDAYS, isLeapJalaliYear, jalMonthLength };
})();
window.toLatinDigits = toLatinDigits;
window.Jalali = Jalali;
})();
