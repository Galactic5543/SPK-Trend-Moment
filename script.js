/* SPK Peramalan - Trend Moment (client-side)

   Alur: unggah CSV/Excel -> daftar kolom -> pencocokan kolom -> tabel x, y, xy, x²
         -> rumus Y' = a + bx -> indeks musim -> prediksi bulan/tahun pilihan.

   a = (Σx²·Σy − Σx·Σxy) / (n·Σx² − (Σx)²)
   b = (n·Σxy − Σx·Σy)   / (n·Σx² − (Σx)²)
   Indeks musim = rata-rata bulan tertentu / rata-rata seluruh periode
   Prediksi akhir = Y' × indeks musim */

const START_X = 0;        // indeks waktu mulai dari 0
const MAX_STEPS = 600;    // batas jarak target dari data terakhir
const MAX_FUTURE_LABELS = 36;

let chartInstance = null;
let table = null;         // {fileName, headers, rows}
let state = null;         // hasil perhitungan terakhir
let lastFocused = null;   // isian pencocokan yang terakhir difokuskan

const $ = (id) => document.getElementById(id);
const esc = (s) => String(s).replace(/[&<>"]/g, (c) => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;' }[c]));

/* ======================= Format angka ======================= */
function round2(v) {
  if (!isFinite(v) || v === 0 || Math.abs(v) >= 1e15 || Math.abs(v) < 1e-6) return v;
  v = Number(v.toPrecision(12));   // buang sisa pecahan biner
  const r = Number(Math.round(Number(v + 'e2')) + 'e-2');
  return Number.isNaN(r) ? v : r;
}
/* Memotong (bukan membulatkan) pada 2 desimal: -0,2867 -> -0,28 */
function trunc2(v) {
  if (!isFinite(v) || v === 0 || Math.abs(v) >= 1e15 || Math.abs(v) < 1e-6) return v;
  v = Number(v.toPrecision(12));
  const r = Number(Math.trunc(Number(v + 'e2')) + 'e-2');
  return Number.isNaN(r) ? v : r;
}
const nf2 = new Intl.NumberFormat('id-ID', { minimumFractionDigits: 2, maximumFractionDigits: 2 });
const nfAuto = new Intl.NumberFormat('id-ID', { minimumFractionDigits: 0, maximumFractionDigits: 2 });
const nf4 = new Intl.NumberFormat('id-ID', { minimumFractionDigits: 4, maximumFractionDigits: 4 });
const fmt = (v) => nf2.format(round2(v));       // selalu 2 desimal
const fmtN = (v) => nfAuto.format(round2(v));   // tanpa desimal bila bilangan bulat
const fmt4 = (v) => nf4.format(v);

/* ======================= Periode / bulan ======================= */
const MONTHS_ID = ['Januari', 'Februari', 'Maret', 'April', 'Mei', 'Juni', 'Juli', 'Agustus', 'September', 'Oktober', 'November', 'Desember'];
const MONTHS_EN = ['January', 'February', 'March', 'April', 'May', 'June', 'July', 'August', 'September', 'October', 'November', 'December'];
const MONTH_TOKENS = {
  jan: 1, januari: 1, january: 1, feb: 2, februari: 2, february: 2, mar: 3, maret: 3, march: 3,
  apr: 4, april: 4, mei: 5, may: 5, jun: 6, juni: 6, june: 6, jul: 7, juli: 7, july: 7,
  agu: 8, agt: 8, agus: 8, agustus: 8, aug: 8, august: 8, sep: 9, sept: 9, september: 9,
  okt: 10, oktober: 10, oct: 10, october: 10, nov: 11, november: 11, des: 12, desember: 12, dec: 12, december: 12
};
const EN_ONLY = new Set(['january', 'february', 'march', 'may', 'june', 'july', 'august', 'october', 'december', 'aug', 'oct', 'dec']);

/* Label periode -> {type, year, month}. Tanggal lengkap (dengan hari) sengaja tidak dianggap periode. */
function parsePeriod(label) {
  const s = String(label ?? '').trim();
  let m = s.match(/^(\d{4})[-\/](\d{1,2})$/);                           // 2025-06
  if (m && +m[2] >= 1 && +m[2] <= 12) return { type: 'ym', year: +m[1], month: +m[2] };
  m = s.match(/^(\d{1,2})[-\/](\d{4})$/);                               // 06-2025
  if (m && +m[1] >= 1 && +m[1] <= 12) return { type: 'ym', year: +m[2], month: +m[1] };
  m = s.match(/^([A-Za-z]+)\.?[\s\-\/]*(\d{4}|\d{2})?$/);               // Januari, Jan 2012, Mei-12
  if (m && MONTH_TOKENS[m[1].toLowerCase()]) {
    const tok = m[1].toLowerCase();
    const yr = m[2] ? (m[2].length === 2 ? 2000 + +m[2] : +m[2]) : null;
    return { type: 'name', month: MONTH_TOKENS[tok], year: yr, hasYear: !!m[2], shortYear: !!m[2] && m[2].length === 2, en: EN_ONLY.has(tok) };
  }
  if (/^\d{4}$/.test(s)) return { type: 'y', year: +s };
  return null;
}

/* Label periode setelah lastLabel */
function nextPeriods(lastLabel, count) {
  const p = parsePeriod(lastLabel);
  const out = [];
  if (!p) {
    for (let k = 1; k <= count; k++) out.push(`Periode +${k}`);
    return out;
  }
  if (p.type === 'y') {
    for (let k = 1; k <= count; k++) out.push(String(p.year + k));
    return out;
  }
  let y = p.year, mo = p.month;
  for (let k = 0; k < count; k++) {
    mo++;
    if (mo > 12) { mo = 1; if (y != null) y++; }
    if (p.type === 'ym') {
      out.push(`${y}-${String(mo).padStart(2, '0')}`);
    } else {
      const name = (p.en ? MONTHS_EN : MONTHS_ID)[mo - 1];
      const yy = p.hasYear ? ' ' + (p.shortYear ? String(y % 100).padStart(2, '0') : y) : '';
      out.push(name + yy);
    }
  }
  return out;
}

/* Tanggal lengkap -> {year, month}: yyyy-mm-dd, yyyy/mm/dd, dd-mm-yyyy, dd/mm/yyyy */
function parseDateYM(s) {
  s = String(s ?? '').trim();
  let y, mo, d;
  let m = s.match(/^(\d{4})[-\/](\d{1,2})[-\/](\d{1,2})/);
  if (m) { y = +m[1]; mo = +m[2]; d = +m[3]; }
  else {
    m = s.match(/^(\d{1,2})[-\/](\d{1,2})[-\/](\d{4})/);
    if (!m) return null;
    d = +m[1]; mo = +m[2]; y = +m[3];
  }
  if (mo < 1 || mo > 12 || d < 1 || d > 31) return null;
  return { year: y, month: mo };
}

/* Tanggal lengkap, "2025-06", atau "Januari 2012" -> {year, month} */
function toYM(v) {
  const d = parseDateYM(v);
  if (d) return d;
  const p = parsePeriod(v);
  if (p && p.month && p.year != null) return { year: p.year, month: p.month };
  return null;
}
/* Angka dari teks/sel; mengenali 1,234.56 / 1.234,56 / 12,5 */
function parseNum(v) {
  if (typeof v === 'number') return isFinite(v) ? v : NaN;
  let s = String(v ?? '').trim().replace(/\s/g, '');
  if (s === '') return NaN;
  const hasC = s.includes(','), hasD = s.includes('.');
  if (hasC && hasD) {
    if (s.lastIndexOf(',') > s.lastIndexOf('.')) s = s.replace(/\./g, '').replace(',', '.');
    else s = s.replace(/,/g, '');
  } else if (hasC) {
    if (/^-?\d{1,3}(,\d{3})+$/.test(s)) s = s.replace(/,/g, '');
    else s = s.replace(',', '.');
  } else if (hasD && /^-?\d{1,3}(\.\d{3}){2,}$/.test(s)) {
    s = s.replace(/\./g, '');
  }
  if (!/^-?\d*\.?\d+(e[-+]?\d+)?$/i.test(s)) return NaN;
  return parseFloat(s);
}

/* ======================= Membaca file ======================= */
function parseCSV(text) {
  text = text.replace(/^﻿/, '');
  const firstLine = text.split(/\r?\n/, 1)[0] || '';
  const delim = firstLine.split(';').length > firstLine.split(',').length ? ';' : ',';
  const rows = [];
  let row = [], cell = '', inQuote = false;
  for (let i = 0; i < text.length; i++) {
    const c = text[i];
    if (inQuote) {
      if (c === '"') { if (text[i + 1] === '"') { cell += '"'; i++; } else inQuote = false; }
      else cell += c;
    } else if (c === '"') {
      inQuote = true;
    } else if (c === delim) {
      row.push(cell); cell = '';
    } else if (c === '\n' || c === '\r') {
      if (c === '\r' && text[i + 1] === '\n') i++;
      row.push(cell); cell = ''; rows.push(row); row = [];
    } else {
      cell += c;
    }
  }
  if (cell !== '' || row.length) { row.push(cell); rows.push(row); }
  return rows;
}

function readText(file) {
  return new Promise((resolve, reject) => {
    const reader = new FileReader();
    reader.onerror = () => reject(new Error('Gagal membaca file.'));
    reader.onload = (e) => resolve(String(e.target.result));
    reader.readAsText(file);
  });
}

function readExcelRows(file) {
  return new Promise((resolve, reject) => {
    const reader = new FileReader();
    reader.onerror = () => reject(new Error('Gagal membaca file.'));
    reader.onload = (e) => {
      try {
        const wb = XLSX.read(new Uint8Array(e.target.result), { type: 'array', cellDates: true });
        const ws = wb.Sheets[wb.SheetNames[0]];
        const raw = XLSX.utils.sheet_to_json(ws, { header: 1, raw: true, defval: '' });
        resolve(raw.map(r => r.map(c => {
          if (c instanceof Date) {            // sel tanggal Excel -> yyyy-mm-dd
            const d = new Date(c.getTime() + 12 * 3600 * 1000);
            return `${d.getFullYear()}-${String(d.getMonth() + 1).padStart(2, '0')}-${String(d.getDate()).padStart(2, '0')}`;
          }
          return c;
        })));
      } catch (err) {
        reject(err);
      }
    };
    reader.readAsArrayBuffer(file);
  });
}

/* File -> {fileName, headers, rows}; baris pertama adalah nama kolom */
async function loadFile(file) {
  const isCsv = /\.csv$/i.test(file.name) || file.type === 'text/csv';
  const all = isCsv ? parseCSV(await readText(file)) : await readExcelRows(file);
  const rows = all.filter(r => r.some(v => String(v ?? '').trim() !== ''));
  if (rows.length < 2) throw new Error('File kosong atau tidak memiliki baris data.');
  const headers = rows[0].map((h, i) => String(h ?? '').trim() || `Kolom ${i + 1}`);
  return { fileName: file.name, headers, rows: rows.slice(1) };
}

/* ======================= Tebakan awal pencocokan kolom ======================= */
function columnStats(rows, j) {
  let dateOK = 0, num = 0, zeros = 0, filled = 0;
  rows.forEach(r => {
    const s = String(r[j] ?? '').trim();
    if (s === '') return;
    filled++;
    if (toYM(s)) dateOK++;
    else { const p = parsePeriod(s); if (p && p.type === 'name') dateOK++; }
    const v = parseNum(r[j]);
    if (!Number.isNaN(v)) { num++; if (v === 0) zeros++; }
  });
  return { dateOK, num, zeros, filled };
}

function detectColumns(t) {
  const n = Math.max(1, t.rows.length);
  const stats = t.headers.map((h, j) => columnStats(t.rows, j));
  const norm = (h) => h.toLowerCase().replace(/\s+/g, '');
  const find = (re) => t.headers.findIndex(h => re.test(norm(h)));
  const xIdx = find(/^x$/);
  const xyIdx = find(/^x[*×]?y$/);
  const x2Idx = find(/^(x2|x\^2|x²|x[*×]x|xkuadrat|x_kuadrat|xsquare|xsquared)$/);
  const reserved = new Set([xIdx, xyIdx, x2Idx].filter(i => i >= 0));

  let dateIdx = -1, best = -1;
  stats.forEach((s, j) => {
    if (reserved.has(j)) return;
    const ratio = s.dateOK / n;
    if (ratio < 0.5) return;
    const score = ratio + (/date|tanggal|tgl|periode|period|bulan|month|waktu|time/i.test(t.headers[j]) ? 1 : 0);
    if (score > best) { best = score; dateIdx = j; }
  });
  if (dateIdx < 0) {   // kolom teks pertama yang bukan angka
    dateIdx = stats.findIndex((s, j) => !reserved.has(j) && s.num / n < 0.5);
  }
  if (dateIdx < 0) dateIdx = 0;

  const cands = [];
  stats.forEach((s, j) => {
    if (j === dateIdx || reserved.has(j) || s.num / n < 0.5) return;
    if (/(^|[_\s])id$|^id([_\s]|$)|kode|^no$|nomor/i.test(t.headers[j])) return;
    cands.push(j);
  });
  const mostlyZero = (j) => stats[j].num > 0 && stats[j].zeros / stats[j].num > 0.5;
  const priority = [/revenue|sales|penjualan|omzet|pendapatan/i, /profit|laba|keuntungan/i, /nilai|value|total|jumlah|units|qty|^y$|\(y\)/i];
  let valIdx = -1;
  for (const re of priority) {
    const j = cands.find(c => re.test(t.headers[c]) && !mostlyZero(c));
    if (j !== undefined) { valIdx = j; break; }
  }
  if (valIdx < 0) { const j = cands.find(c => !mostlyZero(c)); valIdx = j !== undefined ? j : (cands[0] ?? -1); }

  return { dateIdx, valIdx };
}

/* ======================= Membentuk deret data ======================= */
const FULL_DATE = /^(\d{4})[-\/](\d{1,2})[-\/](\d{1,2})|^(\d{1,2})[-\/](\d{1,2})[-\/](\d{4})/;

function buildSeries(t, cfg) {
  const { dateIdx, valIdx } = cfg;
  const notes = [];
  const data = [];
  let skipped = 0;
  t.rows.forEach((r, i) => {
    const periode = String(r[dateIdx] ?? '').trim();
    const y = parseNum(r[valIdx]);
    if (periode === '' || Number.isNaN(y)) { skipped++; return; }
    data.push({ periode, y, line: i + 2 });
  });
  if (skipped) notes.push({ level: 'warning', text: `${skipped} baris dilewati karena periode kosong atau nilai y bukan angka.` });

  // Tanggal lengkap yang tiap bulannya hanya muncul sekali (mis. 1 Jan, 1 Feb) dijadikan periode bulanan
  if (data.length && data.every(d => FULL_DATE.test(d.periode))) {
    const ym = data.map(d => parseDateYM(d.periode));
    if (ym.every(Boolean) && new Set(ym.map(k => k.year * 100 + k.month)).size === data.length) {
      data.forEach((d, i) => { d.periode = `${ym[i].year}-${String(ym[i].month).padStart(2, '0')}`; });
      notes.push({ level: 'info', text: 'Kolom periode berisi tanggal, dan tiap bulan hanya muncul satu kali, sehingga tanggal diubah menjadi periode bulanan (yyyy-mm).' });
    }
  }
  const seen = new Set(), dup = new Set();
  data.forEach(d => { if (seen.has(d.periode)) dup.add(d.periode); seen.add(d.periode); });
  if (dup.size) {
    notes.push({ level: 'warning', text: `Ada periode yang muncul lebih dari satu kali (mis. ${esc([...dup].slice(0, 3).join(', '))}). Setiap baris tetap dihitung sebagai satu periode sendiri. `
      + 'Pastikan data sudah dirangkum menjadi satu baris per periode.' });
  }
  return { data, notes };
}

/* Jenis periode: 'bulan', 'tahun', atau 'indeks' (label bebas) */
function classify(data) {
  const ps = data.map(d => parsePeriod(d.periode));
  if (ps.every(p => p && (p.type === 'ym' || p.type === 'name'))) return 'bulan';
  if (ps.every(p => p && p.type === 'y')) return 'tahun';
  return 'indeks';
}

/* Tahun tiap periode bulanan. Bila label tidak memuat tahun, dipakai startYear dan bertambah setiap kali bulan "berputar". */
function resolveYears(data, startYear) {
  let cur = null, prev = null;
  return data.map(d => {
    const p = parsePeriod(d.periode);
    if (p.year != null) cur = p.year;
    else if (cur == null) cur = startYear != null && !Number.isNaN(startYear) ? startYear : null;
    else if (prev != null && p.month < prev) cur++;
    prev = p.month;
    return cur;
  });
}

/* Data tanpa tahun: periode terakhir dianggap bulan terakhir yang sudah lewat (tidak melebihi hari ini). */
function assumedLastYear(lastMonth) {
  const now = new Date();
  return now.getFullYear() - (lastMonth <= now.getMonth() + 1 ? 0 : 1);
}

/* ======================= Trend Moment ======================= */
function trendMoment(data, rounding) {
  const n = data.length;
  let sumY = 0, sumX = 0, sumXY = 0, sumX2 = 0;
  data.forEach((d, i) => {
    d.x = START_X + i;
    d.xy = d.x * d.y;
    d.x2 = d.x * d.x;
    sumY += d.y; sumX += d.x; sumXY += d.xy; sumX2 += d.x2;
  });
  const denom = n * sumX2 - sumX * sumX;
  if (denom === 0) throw new Error('Nilai x tidak bervariasi sehingga persamaan trend tidak dapat dihitung. Periksa kolom x yang dicocokkan.');
  const numA = sumY * sumX2 - sumX * sumXY;
  const numB = n * sumXY - sumX * sumY;
  const exactA = numA / denom, exactB = numB / denom;
  let a = exactA, b = exactB;
  if (rounding) { a = trunc2(exactA); b = trunc2(exactB); }
  return { n, sumY, sumX, sumXY, sumX2, denom, numA, numB, a, b, exactA, exactB, rataX: sumX / n, rataY: sumY / n };
}

function seasonalIndices(data, rataY, rounding) {
  const by = {};
  data.forEach(d => { (by[d.month] = by[d.month] || []).push(d.y); });
  const rows = [];
  for (let m = 1; m <= 12; m++) {
    const v = by[m];
    if (!v) { rows.push({ month: m, n: 0, avg: null, im: null }); continue; }
    const avg = v.reduce((s, q) => s + q, 0) / v.length;
    let im = rataY !== 0 ? avg / rataY : 1;
    if (rounding) im = round2(im);
    rows.push({ month: m, n: v.length, avg, im });
  }
  return { available: rataY !== 0, rows };
}

/* ======================= UI: langkah 1 (kolom) ======================= */
function showAlert(msg, type = 'danger') {
  $('alertBox').innerHTML = msg ? `<div class="alert alert-${type} mb-0" role="alert">${msg}</div>` : '';
}

function resetResults() {
  state = null;
  $('hasil').classList.add('d-none');
  document.body.classList.remove('has-result');
  $('catatanData').innerHTML = '';
}

function showColumns() {
  const t = table;
  $('kolomCount').textContent = `(${t.headers.length} kolom, ${t.rows.length} baris data)`;
  $('chips').innerHTML = t.headers.map((h, i) => `<button type="button" class="chip" data-i="${i}">${esc(h)}</button>`).join('');
  const head = `<thead><tr>${t.headers.map(h => `<th>${esc(h)}</th>`).join('')}</tr></thead>`;
  const body = t.rows.slice(0, 5).map(r => `<tr>${t.headers.map((h, j) => `<td>${esc(r[j] ?? '')}</td>`).join('')}</tr>`).join('');
  $('tabelPratinjau').innerHTML = head + `<tbody>${body}</tbody>`;
  $('listKolom').innerHTML = t.headers.map(h => `<option value="${esc(h)}"></option>`).join('');
  $('kolomInfo').classList.remove('d-none');
}

function fillMapping() {
  const det = detectColumns(table);
  const name = (i) => (i >= 0 ? table.headers[i] : '');
  $('mapPeriode').value = name(det.dateIdx);
  $('mapY').value = name(det.valIdx);
  document.querySelectorAll('.map-input').forEach(e => e.classList.remove('is-bad'));
  $('mapTahun').value = '';
  $('mapTahun').classList.remove('is-invalid');
  $('sec2').classList.remove('d-none');
}

function resolveCol(name) {
  const q = String(name || '').trim().toLowerCase();
  if (!q) return -1;
  let i = table.headers.findIndex(h => h.toLowerCase() === q);
  if (i < 0) i = table.headers.findIndex(h => h.toLowerCase().replace(/\s+/g, '') === q.replace(/\s+/g, ''));
  return i;
}

function readMapping() {
  document.querySelectorAll('.map-input').forEach(e => e.classList.remove('is-bad'));
  const available = () => `Kolom yang tersedia: ${table.headers.join(', ')}.`;

  const get = (id, label, required) => {
    const raw = $(id).value.trim();
    if (!raw) {
      if (required) { $(id).classList.add('is-bad'); throw new Error(`Isi nama kolom untuk ${label}. ${available()}`); }
      return -1;
    }
    const i = resolveCol(raw);
    if (i < 0) { $(id).classList.add('is-bad'); throw new Error(`Kolom "${esc(raw)}" (${label}) tidak ditemukan di file. ${available()}`); }
    return i;
  };

  const dateIdx = get('mapPeriode', 'Bulan (periode)', true);
  const valIdx = get('mapY', 'Jumlah (y)', true);

  const used = [['Bulan', dateIdx], ['Jumlah (y)', valIdx]].filter(([, i]) => i >= 0);
  const seen = {};
  used.forEach(([label, i]) => {
    if (seen[i] !== undefined) throw new Error(`Kolom "${esc(table.headers[i])}" dipakai untuk dua isian sekaligus (${seen[i]} dan ${label}). Setiap isian harus memakai kolom yang berbeda.`);
    seen[i] = label;
  });
  return { dateIdx, valIdx };
}

/* ======================= Hitung ======================= */
function onCalculate() {
  showAlert('');
  try {
    if (!table) throw new Error('Unggah file terlebih dahulu.');
    const cfg = readMapping();
    const { data, notes } = buildSeries(table, cfg);
    if (data.length < 2) throw new Error('Minimal diperlukan 2 periode data historis.');

    const rounding = $('roundAB').checked;
    const kind = classify(data);
    const startYear = parseInt($('mapTahun').value, 10);
    if ($('mapTahun').value.trim() !== '' && (Number.isNaN(startYear) || startYear < 1900 || startYear > 2300)) {
      $('mapTahun').classList.add('is-invalid');
      throw new Error('Tahun data pertama tidak valid. Isi dengan 4 angka, misalnya 2012.');
    }
    $('mapTahun').classList.remove('is-invalid');
    if (kind === 'bulan' && !Number.isNaN(startYear)) {
      // Periode tanpa tahun diberi tahun (bertambah tiap bulan berputar), label yang sudah bertahun dibiarkan
      const yrs = resolveYears(data, startYear);
      data.forEach((d, i) => {
        const p = parsePeriod(d.periode);
        if (p.year == null) d.periode = `${MONTHS_ID[p.month - 1]} ${yrs[i]}`;
      });
    }
    const res = trendMoment(data, rounding);
    data.forEach(d => { d.yp = res.a + res.b * d.x; });

    if (kind === 'bulan') {
      data.forEach(d => { d.month = parsePeriod(d.periode).month; });
    }
    const valid = data.filter(d => d.y !== 0);
    const mape = valid.length ? valid.reduce((s, d) => s + Math.abs((d.y - d.yp) / d.y), 0) / valid.length * 100 : NaN;
    const seasonal = kind === 'bulan' ? seasonalIndices(data, res.rataY, rounding) : { available: false, rows: [] };

    state = {
      data, res, kind, rounding, mape, seasonal,
      names: { periode: table.headers[cfg.dateIdx], y: cfg.valIdx >= 0 ? table.headers[cfg.valIdx] : 'Jumlah baris' },
      fmtAB: rounding ? fmt : fmt4
    };

    $('catatanData').innerHTML = notes.map(n => `<div class="alert alert-${n.level} py-2 mb-2 small">${n.text}</div>`).join('');
    renderTable();
    renderRumus();
    renderMusim();
    setupTarget();
    $('hasil').classList.remove('d-none');
    document.body.classList.add('has-result');
    runTarget();
    $('hasil').scrollIntoView({ behavior: 'smooth', block: 'start' });
  } catch (err) {
    resetResults();
    showAlert(err.message || 'Terjadi kesalahan saat memproses data.');
    $('alertBox').scrollIntoView({ behavior: 'smooth', block: 'center' });
  }
}

/* ======================= Langkah 3: tabel hasil perhitungan ======================= */
function renderTable() {
  const { data, res, names } = state;
  const rows = data.map(d => `
    <tr><td>${esc(d.periode)}</td><td>${fmtN(d.x)}</td><td>${fmtN(d.y)}</td><td>${fmtN(d.xy)}</td><td>${fmtN(d.x2)}</td></tr>`).join('');
  $('tabelHasil').innerHTML = `
    <thead><tr>
      <th>${esc(names.periode)}</th><th>x</th><th>${esc(names.y)}<br>(y)</th><th>xy</th><th>x&sup2;</th>
    </tr></thead>
    <tbody>${rows}</tbody>
    <tfoot>
      <tr><th>&Sigma;</th><th>${fmtN(res.sumX)}</th><th>${fmtN(res.sumY)}</th><th>${fmtN(res.sumXY)}</th><th>${fmtN(res.sumX2)}</th></tr>
      <tr><th>rata-rata</th><th>${fmtN(res.rataX)}</th><th>${fmtN(res.rataY)}</th><th></th><th></th></tr>
    </tfoot>`;

  const max = data.reduce((m, d) => d.y > m.y ? d : m, data[0]);
  const min = data.reduce((m, d) => d.y < m.y ? d : m, data[0]);
  $('kesimpulanTabel').innerHTML = `
    <div class="alert alert-primary mb-0">
      <h3 class="h6 mb-2">Kesimpulan Tabel</h3>
      <ul class="mb-0">
        <li>Terdapat <strong>${res.n}</strong> periode data. Total y (&Sigma;y) = <strong>${fmtN(res.sumY)}</strong> dan rata-rata per periode = <strong>${fmtN(res.rataY)}</strong>.</li>
        <li>Nilai tertinggi pada <strong>${esc(max.periode)}</strong> (${fmtN(max.y)}) dan terendah pada <strong>${esc(min.periode)}</strong> (${fmtN(min.y)}).</li>
        <li>Total yang dipakai untuk mencari a dan b: &Sigma;x = ${fmtN(res.sumX)}, &Sigma;y = ${fmtN(res.sumY)}, &Sigma;xy = ${fmtN(res.sumXY)}, &Sigma;x&sup2; = ${fmtN(res.sumX2)}.</li>
      </ul>
    </div>`;
}

/* ======================= Langkah 4: rumus Y' = a + bx ======================= */
function renderRumus() {
  const { res, mape, rounding, fmtAB, data } = state;
  const { n, sumY, sumX, sumXY, sumX2, denom, numA, numB, a, b, rataY } = res;
  const naik = b >= 0;
  const persenB = rataY !== 0 ? (b / rataY) * 100 : 0;
  let arah;
  if (Math.abs(persenB) < 0.5) arah = (naik ? 'naik' : 'turun') + ' sangat landai (nyaris datar)';
  else if (Math.abs(persenB) < 2) arah = (naik ? 'naik' : 'turun') + ' secara moderat';
  else arah = (naik ? 'naik' : 'turun') + ' cukup tajam';

  const calcA = [
    `a = (Σx²·Σy − Σx·Σxy) / (n·Σx² − (Σx)²)`,
    `  = ((${fmtN(sumX2)} × ${fmtN(sumY)}) − (${fmtN(sumX)} × ${fmtN(sumXY)})) / ((${n} × ${fmtN(sumX2)}) − (${fmtN(sumX)})²)`,
    `  = (${fmtN(sumX2 * sumY)} − ${fmtN(sumX * sumXY)}) / (${fmtN(n * sumX2)} − ${fmtN(sumX * sumX)})`,
    `  = ${fmtN(numA)} / ${fmtN(denom)}`,
    `  = ${fmt4(res.exactA)}` + (rounding ? `  →  a = ${fmt(a)}` : '')
  ].join('\n');
  const calcB = [
    `b = (n·Σxy − Σx·Σy) / (n·Σx² − (Σx)²)`,
    `  = ((${n} × ${fmtN(sumXY)}) − (${fmtN(sumX)} × ${fmtN(sumY)})) / ((${n} × ${fmtN(sumX2)}) − (${fmtN(sumX)})²)`,
    `  = (${fmtN(n * sumXY)} − ${fmtN(sumX * sumY)}) / (${fmtN(n * sumX2)} − ${fmtN(sumX * sumX)})`,
    `  = ${fmtN(numB)} / ${fmtN(denom)}`,
    `  = ${fmt4(res.exactB)}` + (rounding ? `  →  b = ${fmt(b)}` : '')
  ].join('\n');

  let mapeText = '';
  if (!Number.isNaN(mape)) {
    const level = mape < 10 ? 'sangat akurat' : mape < 20 ? 'baik' : mape < 50 ? 'cukup (masih wajar)' : 'kurang akurat';
    const above = data.filter(d => d.y > d.yp).length, below = data.filter(d => d.y < d.yp).length;
    mapeText = `<p>Tingkat kesalahan garis tren (MAPE) sebesar <strong>${fmt(mape)}%</strong>, tergolong <strong>${level}</strong>.
      Nilai aktual berada di atas garis tren pada ${above} periode dan di bawahnya pada ${below} periode.</p>`;
  }

  $('langkahRumus').innerHTML = `
    <p>Persamaan trend: <code>Y' = a + bx</code>, dengan Y' = nilai trend (peramalan), a = bilangan konstan,
    b = slope (koefisien kecondongan garis trend), dan x = indeks waktu. Nilai a dan b dicari dari total pada tabel di langkah 3:</p>
    <pre class="calc">${calcA}</pre>
    <pre class="calc">${calcB}</pre>
    ${rounding ? '<p class="small text-muted">Nilai a dan b diambil 2 desimal (dipotong, bukan dibulatkan) sebelum dipakai meramal. Matikan opsi 2 desimal di langkah 2 untuk memakai nilai tanpa pemotongan.</p>' : ''}
    <p class="mb-2">Persamaan yang diperoleh:</p>
    <div class="eq-result mb-3">Y' = ${fmtAB(a)} ${naik ? '+' : '−'} ${fmtAB(Math.abs(b))} x</div>
    <p class="mb-2">Dengan persamaan tersebut, nilai Y' tiap periode dihitung dengan memasukkan nilai x:</p>
    <div class="table-responsive mb-3">
      <table class="tbl-spk" style="max-width:640px">
        <thead><tr><th>${esc(state.names.periode)}</th><th>x</th><th>Perhitungan Y' = a + bx</th><th>Y'</th><th>y (aktual)</th></tr></thead>
        <tbody>${data.map(d => `<tr><td>${esc(d.periode)}</td><td>${fmtN(d.x)}</td><td>${fmtAB(a)} ${b >= 0 ? '+' : '−'} (${fmtAB(Math.abs(b))} × ${fmtN(d.x)})</td><td>${fmt(d.yp)}</td><td>${fmtN(d.y)}</td></tr>`).join('')}</tbody>
      </table>
    </div>
    <p>Setiap periode, nilai pada garis tren ${naik ? 'naik' : 'turun'} sekitar ${fmtAB(Math.abs(b))}
    (${Math.abs(persenB).toFixed(2).replace('.', ',')}% dari rata-rata data), sehingga tren data ${arah}.</p>
    ${mapeText}`;
}

/* ======================= Langkah 5: indeks musim ======================= */
function renderMusim() {
  const { seasonal, kind, res, data } = state;
  if (kind !== 'bulan') {
    $('langkahMusim').innerHTML = `<p class="mb-0">Indeks musim tidak dihitung karena kolom periode tidak berisi bulan yang dapat dikenali
      (contoh yang dikenali: "Januari", "Agustus 2013", "2025-06"). Prediksi akhir sama dengan prediksi trend (Y').</p>`;
    return;
  }
  const rows = seasonal.rows.map(r => `
    <tr><td>${MONTHS_ID[r.month - 1]}</td><td>${r.n || '-'}</td><td>${r.avg == null ? '-' : fmt(r.avg)}</td><td>${r.im == null ? '-' : fmt(r.im)}</td></tr>`).join('');
  const missing = seasonal.rows.filter(r => !r.n).length;
  $('langkahMusim').innerHTML = `
    <p>Data yang dipengaruhi faktor musiman dikoreksi dengan <strong>indeks musim</strong>:
    rata-rata nilai pada bulan tertentu dibagi rata-rata seluruh periode (<strong>${fmt(res.rataY)}</strong>).
    Bila satu bulan muncul pada beberapa tahun, dipakai rata-rata bulan tersebut.</p>
    <pre class="calc">indeks musim = rata-rata bulan tertentu / rata-rata seluruh periode
prediksi akhir = Y' × indeks musim</pre>
    <div class="table-responsive">
      <table class="tbl-spk" style="max-width:560px">
        <thead><tr><th>Bulan</th><th>Jumlah data</th><th>Rata-rata bulan</th><th>Indeks musim</th></tr></thead>
        <tbody>${rows}</tbody>
      </table>
    </div>
    ${data.length < 12 || missing ? `<p class="small text-muted mt-2 mb-0">Data baru mencakup ${data.length} periode${missing ? ` dan ${missing} bulan belum punya data` : ''}.
      Bulan tanpa data memakai indeks 1,00 (tanpa koreksi musim). Indeks musim paling andal bila data mencakup minimal 2 tahun.</p>` : ''}`;
}

/* ======================= Langkah 6: prediksi bulan terpilih ======================= */
function setupTarget() {
  const { data, kind } = state;
  const last = data[data.length - 1];
  const sel = $('tgtMonth');
  if (sel.options.length <= 1) MONTHS_ID.forEach((m, i) => { const o = document.createElement('option'); o.value = String(i + 1); o.textContent = m; sel.appendChild(o); });

  $('boxTgtMonth').classList.toggle('d-none', kind !== 'bulan');
  $('boxTgtYear').classList.toggle('d-none', kind === 'indeks');
  $('boxTgtX').classList.toggle('d-none', kind !== 'indeks');
  const years = kind === 'bulan' ? resolveYears(data, NaN) : [];
  const yearless = kind === 'bulan' && years[years.length - 1] == null;
  $('boxTgtStart').classList.add('d-none');
  $('targetAlert').innerHTML = '';

  if (kind === 'bulan') {
    const nm = last.month % 12 + 1;
    $('tgtMonth').value = String(nm);
    state.anchorL = null;
    const ly = years[years.length - 1];
    $('tgtYear').value = yearless ? '' : String(ly + (last.month === 12 ? 1 : 0));
    $('targetPetunjuk').textContent = yearless
      ? `Pilih bulan dan tahun yang ingin diprediksi, lalu klik "Hitung Prediksi". Data tidak memuat tahun, sehingga bulan dan tahun pertama yang Anda pilih dianggap sebagai periode pertama setelah data terakhir (${last.periode}).`
      : `Pilih bulan dan tahun yang ingin diprediksi (data terakhir: ${last.periode}), lalu klik "Hitung Prediksi".`;
  } else if (kind === 'tahun') {
    $('tgtYear').value = String(parsePeriod(last.periode).year + 1);
    $('targetPetunjuk').textContent = `Isi tahun yang ingin diprediksi (data terakhir: ${last.periode}), lalu klik "Hitung Prediksi".`;
  } else {
    $('tgtX').value = String(last.x + 1);
    $('targetPetunjuk').textContent = `Isi nomor periode (x) yang ingin diprediksi. Data terakhir berada pada x = ${fmtN(last.x)} (${last.periode}).`;
  }
}

function planTarget() {
  const { data, kind } = state;
  const last = data[data.length - 1];
  let steps, label, month = null;

  if (kind === 'bulan') {
    const years = resolveYears(data, NaN);
    month = parseInt($('tgtMonth').value, 10);
    const yr = parseInt($('tgtYear').value, 10);
    if (!month) throw new Error('Pilih bulan target.');
    if (Number.isNaN(yr)) throw new Error('Isi tahun target.');
    let ly = years[years.length - 1];
    if (ly == null) {
      if (state.anchorL == null) {
        const first = ((month - last.month + 11) % 12) + 1;
        state.anchorL = (yr * 12 + month - first - last.month) / 12;
      }
      ly = state.anchorL;
    }
    steps = (yr * 12 + month) - (ly * 12 + last.month);
    label = `${MONTHS_ID[month - 1]} ${yr}`;
  } else if (kind === 'tahun') {
    const yr = parseInt($('tgtYear').value, 10);
    if (Number.isNaN(yr)) throw new Error('Isi tahun target.');
    steps = yr - parsePeriod(last.periode).year;
    label = String(yr);
  } else {
    const x = parseInt($('tgtX').value, 10);
    if (Number.isNaN(x)) throw new Error('Isi nomor periode (x) target.');
    steps = x - last.x;
    label = `Periode ke-${x}`;
  }
  if (steps > MAX_STEPS) throw new Error(`Target terlalu jauh dari data terakhir (maksimal ${MAX_STEPS} periode).`);
  if (last.x + steps < data[0].x) throw new Error(`Target berada sebelum data pertama (${data[0].periode}). Pilih periode yang sama dengan atau setelah data pertama.`);
  return { steps, label, month, x: last.x + steps };
}

function runTarget() {
  if (!state) return;
  $('targetAlert').innerHTML = '';
  let plan;
  try {
    plan = planTarget();
  } catch (err) {
    $('targetAlert').innerHTML = `<div class="alert alert-${err.soft ? 'warning' : 'danger'} mb-0" role="alert">${esc(err.message)}</div>`;
    $('targetHasil').innerHTML = '';
    return;
  }
  const { data, res, seasonal, kind, fmtAB, mape } = state;
  const last = data[data.length - 1];
  const yp = res.a + res.b * plan.x;

  let im = 1, imMissing = false, avgMonth = null;
  if (kind === 'bulan' && seasonal.available) {
    const row = seasonal.rows[plan.month - 1];
    if (row.n) { im = row.im; avgMonth = row.avg; } else imMissing = true;
  }
  const useSeason = kind === 'bulan' && seasonal.available;
  const final = yp * im;
  const inHistory = plan.steps <= 0;
  const actual = inHistory ? data[data.length - 1 + plan.steps].y : null;

  const calc = [
    `${plan.label}  (x = ${fmtN(plan.x)})`,
    `Y' = a + bx`,
    `   = ${fmtAB(res.a)} + (${fmtAB(res.b)} × ${fmtN(plan.x)})`,
    `   = ${fmt(yp)}`
  ];
  if (useSeason) {
    if (imMissing) {
      calc.push('', `Bulan ${MONTHS_ID[plan.month - 1]} tidak ada di data historis → indeks musim = 1,00`, `Prediksi = ${fmt(yp)} × 1,00 = ${fmt(final)}`);
    } else {
      calc.push('', `indeks musim ${MONTHS_ID[plan.month - 1]} = rata-rata bulan / rata-rata seluruh periode`,
        `                  = ${fmt(avgMonth)} / ${fmt(res.rataY)}`, `                  = ${fmt(im)}`,
        `Prediksi = Y' × indeks musim = ${fmt(yp)} × ${fmt(im)} = ${fmt(final)}`);
    }
  }

  const naik = res.b >= 0;
  const n = data.length;
  const satuan = kind === 'bulan' ? 'bulan' : 'periode';
  let p1 = `Berdasarkan ${n} ${satuan} data yang diunggah, nilai cenderung <strong>${naik ? 'naik' : 'turun'}</strong> sekitar ${fmtAB(Math.abs(res.b))} setiap ${satuan}. `
    + `Rumus garis tren yang diperoleh adalah Y' = ${fmtAB(res.a)} ${naik ? '+' : '−'} ${fmtAB(Math.abs(res.b))}x. `;
  if (inHistory) {
    p1 += `<strong>${esc(plan.label)}</strong> sudah ada di data, dengan nilai sebenarnya ${fmtN(actual)}. Garis tren memberi angka ${fmt(yp)}`
      + (useSeason ? `, lalu dikoreksi indeks musim menjadi <strong>${fmt(final)}</strong>.` : '.');
  } else {
    p1 += `Untuk <strong>${esc(plan.label)}</strong> (${plan.steps} ${satuan} setelah data terakhir), garis tren memberi angka ${fmt(yp)}`
      + (useSeason ? `, lalu dikoreksi indeks musim sehingga hasil ramalan akhirnya adalah <strong>${fmt(final)}</strong>.` : `, sehingga hasil ramalannya adalah <strong>${fmt(final)}</strong>.`);
  }
  let p2 = '';
  if (useSeason && !imMissing) {
    const persen = Math.round(im * 100);
    const arah = im < 0.95 ? `biasanya lebih rendah dari rata-rata bulan lain (sekitar ${persen}% dari rata-rata)`
      : im > 1.05 ? `biasanya lebih tinggi dari rata-rata bulan lain (sekitar ${persen}% dari rata-rata)`
      : 'biasanya hampir sama dengan rata-rata bulan lain';
    p2 = `Koreksi dilakukan karena ${MONTHS_ID[plan.month - 1]} ${arah}, yaitu indeks musim ${fmt(im)}. `;
  } else if (useSeason && imMissing) {
    p2 = `Bulan ${MONTHS_ID[plan.month - 1]} belum ada di data, jadi tidak ada koreksi musim (indeks 1,00). `;
  }
  if (!Number.isNaN(mape)) {
    const tingkat = mape < 10 ? 'sangat kecil, sehingga garis tren cocok dengan data'
      : mape < 20 ? 'kecil, sehingga garis tren cukup cocok dengan data'
      : mape < 50 ? 'sedang, sehingga garis tren hanya cukup mewakili data'
      : 'cukup besar, artinya data naik-turun tajam dan tidak mengikuti garis tren dengan dekat';
    p2 += `Rata-rata selisih antara garis tren dan data sebenarnya (MAPE) adalah ${fmt(mape)}%, tergolong ${tingkat}. `
      + (mape >= 20 ? 'Karena itu, anggap hasil ini sebagai perkiraan kasar, bukan angka pasti.' : 'Hasil ini cukup layak dijadikan acuan.');
  }
  if (kind === 'bulan' && n < 12) p2 += ' Data kurang dari 12 bulan, sehingga indeks musim hanya perkiraan kasar.';
  const kesimpulanHtml = `<p class="mb-2">${p1}</p>` + (p2 ? `<p class="mb-0">${p2.trim()}</p>` : '');

  $('targetHasil').innerHTML = `
    <div class="card target-card">
      <div class="card-body">
        <div class="row g-3 align-items-center">
          <div class="col-md-4">
            <div class="small text-muted">Prediksi ${esc(plan.label)}</div>
            <div class="target-value">${fmt(final)}</div>
            <div class="small text-muted mt-1">${useSeason ? `Y' ${fmt(yp)} × indeks musim ${fmt(im)}${imMissing ? '*' : ''}` : `Y' (tanpa koreksi musiman)`}</div>
          </div>
          <div class="col-md-8"><pre class="calc mb-0">${calc.join('\n')}</pre></div>
        </div>
      </div>
    </div>
    <div class="alert alert-primary mt-3 mb-0">
      <h3 class="h6 mb-2">Kesimpulan</h3>
      ${kesimpulanHtml}
    </div>`;

  renderChart(plan, yp, final, useSeason);
}

/* ======================= Langkah 7: grafik ======================= */
function renderChart(plan, yp, final, useSeason) {
  const { data, res } = state;
  const last = data[data.length - 1];
  const labels = data.map(d => d.periode);
  const aktual = data.map(d => d.y);
  const trend = data.map(d => d.yp);
  const hasil = data.map(() => null);

  if (plan.steps <= 0) {
    hasil[data.length - 1 + plan.steps] = final;
  } else {
    const future = plan.steps <= MAX_FUTURE_LABELS
      ? nextPeriods(last.periode, plan.steps)
      : [plan.label];
    future.forEach((lb, i) => {
      const x = future.length === plan.steps ? last.x + i + 1 : plan.x;
      labels.push(lb);
      aktual.push(null);
      trend.push(res.a + res.b * x);
      hasil.push(null);
    });
    labels[labels.length - 1] = plan.label;
    hasil[hasil.length - 1] = final;
  }

  const kanan = plan.steps > 0 ? `${plan.steps} ${state.kind === 'bulan' ? 'bulan' : 'periode'} ke depan` : 'di dalam data';
  const turun = final < yp - 1e-9, naik2 = final > yp + 1e-9;
  $('penjelasanGrafik').innerHTML = `
    <h3 class="h6 mb-2">Cara membaca grafik</h3>
    <ul class="mb-2">
      <li><strong style="color:#0d6efd">Garis biru</strong>: data sebenarnya. Naik-turunnya garis menunjukkan nilai tiap ${state.kind === 'bulan' ? 'bulan' : 'periode'} yang berubah-ubah.</li>
      <li><strong style="color:#e8590c">Garis oranye putus-putus</strong>: arah umum (tren) dari seluruh data, dibuat lurus supaya mudah terlihat apakah nilai cenderung ${res.b >= 0 ? 'naik' : 'turun'}. Garis ini diteruskan ${kanan} sampai ${esc(plan.label)}.</li>
      <li><strong style="color:#2f9e44">Titik hijau (belah ketupat)</strong>: hasil ramalan akhir untuk ${esc(plan.label)}, yaitu ${fmt(final)}.</li>
    </ul>
    <p class="mb-0">${useSeason && (turun || naik2)
      ? `Titik hijau berada ${turun ? 'di bawah' : 'di atas'} garis oranye karena angka tren (${fmt(yp)}) masih dikoreksi dengan kebiasaan musim bulan tersebut. `
      : 'Titik hijau berada pada garis oranye karena tidak ada koreksi musim. '}
      Semakin dekat garis biru ke garis oranye, semakin baik garis tren mewakili data; bila garis biru naik-turun jauh dari garis oranye, ramalan sebaiknya dianggap perkiraan kasar.</p>`;

  if (chartInstance) chartInstance.destroy();
  chartInstance = new Chart($('chartTrend'), {
    type: 'line',
    data: {
      labels,
      datasets: [
        { label: 'Data Aktual', data: aktual, borderColor: '#0d6efd', backgroundColor: '#0d6efd', borderWidth: 2.5, pointRadius: 4, tension: 0.15, spanGaps: false },
        { label: "Prediksi Trend (Y')", data: trend, borderColor: '#e8590c', backgroundColor: '#e8590c', borderWidth: 2.5, borderDash: [7, 4], pointRadius: 3, tension: 0 },
        { label: useSeason ? "Prediksi Akhir (Y' × Indeks Musim)" : 'Prediksi Akhir', data: hasil, borderColor: '#2f9e44', backgroundColor: '#2f9e44', pointStyle: 'rectRot', pointRadius: 9, showLine: false }
      ]
    },
    options: {
      responsive: true,
      maintainAspectRatio: false,
      interaction: { mode: 'index', intersect: false },
      plugins: {
        legend: { position: 'top' },
        tooltip: { callbacks: { label: (c) => c.raw == null ? '' : `${c.dataset.label}: ${fmt(c.raw)}` } }
      },
      scales: {
        y: { title: { display: true, text: 'Nilai' }, ticks: { callback: (v) => nf2.format(v) } },
        x: { title: { display: true, text: 'Periode' } }
      }
    }
  });
}

/* ======================= Inisialisasi ======================= */
async function onFileChange() {
  showAlert('');
  resetResults();
  table = null;
  $('kolomInfo').classList.add('d-none');
  $('sec2').classList.add('d-none');
  const file = $('excelFile').files[0];
  if (!file) return;
  try {
    table = await loadFile(file);
    showColumns();
    fillMapping();
  } catch (err) {
    table = null;
    showAlert(esc(err.message || 'Gagal membaca file.'));
  }
}

$('excelFile').addEventListener('change', onFileChange);
$('btnHitung').addEventListener('click', onCalculate);
$('btnTarget').addEventListener('click', runTarget);
['tgtMonth', 'tgtYear', 'tgtX'].forEach(id => $(id).addEventListener('change', runTarget));
$('roundAB').addEventListener('change', () => { if (state) onCalculate(); });
document.querySelectorAll('.map-input').forEach(el => {
  el.addEventListener('focus', () => { lastFocused = el; });
  el.addEventListener('input', () => el.classList.remove('is-bad'));
});
$('chips').addEventListener('click', (e) => {
  const b = e.target.closest('.chip');
  if (!b) return;
  const name = table.headers[+b.dataset.i];
  const target = lastFocused && !lastFocused.disabled ? lastFocused : ($('mapPeriode').value ? $('mapY') : $('mapPeriode'));
  target.value = name;
  target.classList.remove('is-bad');
  target.dispatchEvent(new Event('change'));
  target.focus();
});