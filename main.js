/*
Pocket QSO
Copyright (C) 2026 Mike Poppelaars

This program is free software: you can redistribute it and/or modify
it under the terms of the GNU General Public License as published by
the Free Software Foundation, version 3.

This program is distributed in the hope that it will be useful,
but WITHOUT ANY WARRANTY; without even the implied warranty of
MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE. See the
GNU General Public License for more details.

You should have received a copy of the GNU General Public License
along with this program. If not, see <https://www.gnu.org/licenses/>.
SPDX-License-Identifier: GPL-3.0-only
*/
let qsos = [];
let seenSerials = new Set();
const LOCAL_KEY = "hamlog.qsos.v1";
const OUTBOX_KEY = "hamlog.outbox.v1";
const LAST_APPLIED_SERIAL_KEY = "hamlog.last_serial.v1";
let editingId = null;
let pendingDeleteId = null;
let pendingDeleteTimer = null;
const DEFAULT_SEND_UPDATE_INTERVAL = 10000;
const DEFAULT_SEND_UPDATE_MAX_SIZE = 128000;
const OUTBOX_FLUSH_INTERVAL_MS = 12000;
const OUTBOX_RETRY_MIN_DELAY_MS = 250;
const MAX_UPDATE_INFO_BYTES = 128;
let lastSendUpdateAt = 0;
const AUTO_DT_REFRESH_MS = 30000;
let dtManualOverride = false;
let lastAutoDtValue = "";
let outbox = [];
let outboxFlushInProgress = false;
let outboxRetryTimer = null;
let lastAppliedSerial = 0;

function canSendUpdate() {
  return !!window.webxdc?.sendUpdate;
}

function getSendUpdateIntervalMs() {
  const configured = Number(window.webxdc?.sendUpdateInterval);
  if (Number.isFinite(configured) && configured >= 0) return configured;
  return DEFAULT_SEND_UPDATE_INTERVAL;
}

function getSendUpdateMaxSize() {
  const configured = Number(window.webxdc?.sendUpdateMaxSize);
  if (Number.isFinite(configured) && configured > 0) return configured;
  return DEFAULT_SEND_UPDATE_MAX_SIZE;
}

function truncateUtf8(str, maxBytes) {
  if (!str || maxBytes <= 0) return "";
  const source = String(str);

  if (typeof TextEncoder !== "undefined") {
    const encoder = new TextEncoder();
    if (encoder.encode(source).length <= maxBytes) return source;

    let out = "";
    let bytes = 0;
    for (const ch of source) {
      const chBytes = encoder.encode(ch).length;
      if (bytes + chBytes > maxBytes) break;
      out += ch;
      bytes += chBytes;
    }
    return out;
  }

  let out = "";
  let bytes = 0;
  for (const ch of source) {
    const cp = ch.codePointAt(0);
    const chBytes = cp <= 0x7F ? 1 : (cp <= 0x7FF ? 2 : (cp <= 0xFFFF ? 3 : 4));
    if (bytes + chBytes > maxBytes) break;
    out += ch;
    bytes += chBytes;
  }
  return out;
}

function utf8ByteLength(str) {
  if (str == null) return 0;
  const source = String(str);

  if (typeof TextEncoder !== "undefined") {
    return new TextEncoder().encode(source).length;
  }

  let bytes = 0;
  for (const ch of source) {
    const cp = ch.codePointAt(0);
    bytes += cp <= 0x7F ? 1 : (cp <= 0x7FF ? 2 : (cp <= 0xFFFF ? 3 : 4));
  }
  return bytes;
}

function updateSizeBytes(update) {
  return utf8ByteLength(JSON.stringify(update));
}

function sanitizeUpdateInfo(info) {
  if (info == null) return "";
  const normalized = String(info).replace(/[\r\n]+/g, " ").trim();
  return truncateUtf8(normalized, MAX_UPDATE_INFO_BYTES);
}

function makeUpdate(payload, info) {
  const update = { payload };
  const safeInfo = sanitizeUpdateInfo(info);
  if (safeInfo) update.info = safeInfo;
  return update;
}

function setSyncStatus(text) {
  const el = document.getElementById("syncStatus");
  if (el) el.textContent = `Sync: ${text || "checking"}`;
}

function updatePendingStatus() {
  const el = document.getElementById("pendingStatus");
  if (el) el.textContent = `Pending updates: ${outbox.length}`;
}

function setStorageWarning(message) {
  const el = document.getElementById("storageStatus");
  if (!el) return;
  if (!message) {
    el.textContent = "";
    el.classList.add("hidden");
    return;
  }
  el.textContent = message;
  el.classList.remove("hidden");
}

function readStoredArray(key, label) {
  try {
    const raw = localStorage.getItem(key);
    if (!raw) return [];
    const parsed = JSON.parse(raw);
    if (Array.isArray(parsed)) return parsed;
    setStorageWarning(`Storage warning: ${label} was not valid and was reset.`);
    return [];
  } catch (_) {
    setStorageWarning(`Storage warning: could not read ${label}.`);
    return [];
  }
}

function writeStoredArray(key, value, label) {
  try {
    localStorage.setItem(key, JSON.stringify(value));
    setStorageWarning("");
    return true;
  } catch (_) {
    setStorageWarning(`Storage warning: could not save ${label}.`);
    return false;
  }
}

function readStoredNumber(key, fallback = 0) {
  try {
    const raw = localStorage.getItem(key);
    if (raw == null) return fallback;
    const parsed = Number(raw);
    return Number.isFinite(parsed) ? parsed : fallback;
  } catch (_) {
    return fallback;
  }
}

function writeStoredNumber(key, value) {
  try {
    localStorage.setItem(key, String(value));
    return true;
  } catch (_) {
    return false;
  }
}

function isPlainObject(value) {
  return !!value && typeof value === "object" && !Array.isArray(value);
}

function toText(value) {
  return value == null ? "" : String(value);
}

function normalizeQso(raw) {
  if (!isPlainObject(raw)) return null;
  const id = toText(raw.id).trim();
  const callsign = toText(raw.callsign).trim().toUpperCase();
  const dt = toText(raw.dt).trim();
  if (!id || !callsign || !dt) return null;

  const parsedTs = Number(raw.ts);
  const ts = Number.isFinite(parsedTs) ? parsedTs : (tsFromDT(dt) || Date.now());
  return {
    id,
    callsign,
    dt,
    band: toText(raw.band).trim(),
    freq: toText(raw.freq).trim(),
    mode: toText(raw.mode).trim().toUpperCase(),
    setup: toText(raw.setup).trim(),
    myGrid: toText(raw.myGrid).trim().toUpperCase(),
    theirGrid: toText(raw.theirGrid).trim().toUpperCase(),
    rstS: toText(raw.rstS).trim(),
    rstR: toText(raw.rstR).trim(),
    notes: toText(raw.notes).trim(),
    ts
  };
}

function normalizeUpdatePayload(payload) {
  if (!isPlainObject(payload)) return null;
  const type = toText(payload.type).trim();
  if (type === "add_qso" || type === "edit_qso") {
    const qso = normalizeQso(payload.qso);
    return qso ? { type, qso } : null;
  }
  if (type === "bulk_add") {
    if (!Array.isArray(payload.qsos)) return null;
    const qsos = payload.qsos.map(normalizeQso).filter(Boolean);
    return qsos.length ? { type, qsos } : null;
  }
  if (type === "delete_qso") {
    const id = toText(payload.id).trim();
    return id ? { type, id } : null;
  }
  return null;
}

function normalizeOutboxUpdates(items) {
  if (!Array.isArray(items)) return [];
  const normalized = [];
  for (const item of items) {
    if (!isPlainObject(item)) continue;
    const payload = normalizeUpdatePayload(item.payload);
    if (!payload) continue;
    normalized.push(makeUpdate(payload, item.info));
  }
  return normalized;
}

function saveOutbox() {
  return writeStoredArray(OUTBOX_KEY, outbox, "pending updates");
}

function loadOutbox() {
  outbox = normalizeOutboxUpdates(readStoredArray(OUTBOX_KEY, "pending updates"));
}

function loadLastAppliedSerial() {
  const saved = readStoredNumber(LAST_APPLIED_SERIAL_KEY, 0);
  lastAppliedSerial = Number.isFinite(saved) && saved > 0 ? Math.floor(saved) : 0;
}

function saveLastAppliedSerial(serial) {
  if (!Number.isFinite(serial) || serial <= lastAppliedSerial) return;
  lastAppliedSerial = Math.floor(serial);
  writeStoredNumber(LAST_APPLIED_SERIAL_KEY, lastAppliedSerial);
}

function enqueueUpdate(payload, info) {
  const normalizedPayload = normalizeUpdatePayload(payload);
  if (!normalizedPayload) return false;
  outbox.push(makeUpdate(normalizedPayload, info));
  saveOutbox();
  updatePendingStatus();
  flushOutbox();
  return true;
}

async function sendWebxdcUpdate(update) {
  if (!window.webxdc?.sendUpdate) throw new Error("sendUpdate unavailable");
  // Spec: second argument is deprecated; pass empty string for compatibility.
  const maybePromise = window.webxdc.sendUpdate(update, "");
  if (maybePromise && typeof maybePromise.then === "function") {
    await maybePromise;
  }
}

function clearOutboxRetryTimer() {
  if (!outboxRetryTimer) return;
  clearTimeout(outboxRetryTimer);
  outboxRetryTimer = null;
}

function scheduleOutboxRetry(waitMs) {
  if (!Number.isFinite(waitMs) || waitMs <= 0) {
    clearOutboxRetryTimer();
    return;
  }
  const delay = Math.max(OUTBOX_RETRY_MIN_DELAY_MS, Math.ceil(waitMs));
  clearOutboxRetryTimer();
  outboxRetryTimer = setTimeout(() => {
    outboxRetryTimer = null;
    flushOutbox();
  }, delay);
}

async function flushOutbox() {
  updatePendingStatus();
  if (!canSendUpdate()) {
    clearOutboxRetryTimer();
    setSyncStatus("local mode");
    return;
  }
  if (!outbox.length) {
    clearOutboxRetryTimer();
    setSyncStatus("up to date");
    return;
  }
  if (outboxFlushInProgress) return;

  const waitMs = Math.max(0, (lastSendUpdateAt + getSendUpdateIntervalMs()) - Date.now());
  if (waitMs > 0) {
    setSyncStatus("waiting interval");
    scheduleOutboxRetry(waitMs);
    return;
  }
  clearOutboxRetryTimer();

  const nextUpdate = outbox[0];
  const maxUpdateSize = getSendUpdateMaxSize();
  const nextUpdateSize = updateSizeBytes(nextUpdate);
  if (nextUpdateSize > maxUpdateSize) {
    outbox.shift();
    saveOutbox();
    updatePendingStatus();
    setSyncStatus("oversized update skipped");
    setStatus(`Skipped one oversized sync update (${nextUpdateSize}/${maxUpdateSize} bytes).`);
    console.warn("Skipped oversized update", {
      size: nextUpdateSize,
      max: maxUpdateSize,
      payloadType: nextUpdate?.payload?.type || "unknown"
    });
    if (outbox.length) flushOutbox();
    return;
  }

  outboxFlushInProgress = true;
  setSyncStatus("sending");
  try {
    await sendWebxdcUpdate(nextUpdate);
    outbox.shift();
    saveOutbox();
    lastSendUpdateAt = Date.now();
    setSyncStatus(outbox.length ? "pending" : "up to date");
  } catch (_) {
    // Back off according to sendUpdateInterval to avoid tight retry loops.
    lastSendUpdateAt = Date.now();
    setSyncStatus("retrying");
  } finally {
    outboxFlushInProgress = false;
    updatePendingStatus();
  }

  if (outbox.length) flushOutbox();
}

function escapeHtml(s) {
  return (s ?? "").toString().replace(/[&<>"']/g, c => ({
    "&":"&amp;","<":"&lt;",">":"&gt;", "\"":"&quot;","'":"&#39;"
  }[c]));
}

function pad2(n) {
  return String(n).padStart(2, "0");
}

function parseLocalDT(dt) {
  if (!dt) return null;
  const m = dt.match(/^(\d{4})-(\d{2})-(\d{2})T(\d{2}):(\d{2})/);
  if (!m) return null;
  return new Date(Number(m[1]), Number(m[2]) - 1, Number(m[3]), Number(m[4]), Number(m[5]));
}

function parseFlexibleDT(dt) {
  if (!dt) return null;
  const s = dt.trim();
  if (/[zZ]|[+-]\d{2}:?\d{2}$/.test(s)) {
    const t = Date.parse(s);
    if (!Number.isNaN(t)) return new Date(t);
  }
  return parseLocalDT(s.replace(" ", "T"));
}

function toLocalInputValue(date) {
  return `${date.getFullYear()}-${pad2(date.getMonth() + 1)}-${pad2(date.getDate())}T${pad2(date.getHours())}:${pad2(date.getMinutes())}`;
}

function tsFromLocalDT(dt) {
  const d = parseLocalDT(dt);
  return d ? d.getTime() : null;
}

function tsFromDT(dt) {
  const d = parseFlexibleDT(dt);
  return d ? d.getTime() : null;
}

function formatDisplayDT(dt) {
  return dt ? dt.replace("T", " ") : "";
}

function getDtElement() {
  return document.getElementById("dt");
}

function setDtNow() {
  const dtEl = getDtElement();
  if (!dtEl) return "";
  const nowValue = toLocalInputValue(new Date());
  dtEl.value = nowValue;
  lastAutoDtValue = nowValue;
  return nowValue;
}

function refreshDtIfAuto() {
  if (editingId || dtManualOverride) return;
  const dtEl = getDtElement();
  if (!dtEl) return;
  const nowValue = toLocalInputValue(new Date());
  if (dtEl.value !== nowValue) {
    dtEl.value = nowValue;
    lastAutoDtValue = nowValue;
  }
}

function handleDtInput() {
  const dtEl = getDtElement();
  if (!dtEl || editingId) return;
  dtManualOverride = dtEl.value !== lastAutoDtValue;
}

function formatMHzValue(mhz) {
  const s = mhz.toFixed(6);
  return s.replace(/\.?0+$/, "");
}

function parseFreqMHz(freq) {
  if (!freq) return null;
  const s = freq.toString().trim().toLowerCase();
  if (!s) return null;
  // Avoid interpreting band labels like "20m" as MHz.
  if (/[a-z]/.test(s) && !s.includes("hz")) return null;

  const num = parseFloat(s.replace(",", "."));
  if (Number.isNaN(num)) return null;

  if (s.includes("ghz")) return num * 1000;
  if (s.includes("khz")) return num / 1000;
  if (s.includes("hz") && !s.includes("khz") && !s.includes("mhz") && !s.includes("ghz")) return num / 1e6;
  return num; // default MHz
}

function formatFreqDisplay(freq) {
  if (!freq) return "";
  const mhz = parseFreqMHz(freq);
  if (mhz == null) return freq.toString().trim();
  return `${formatMHzValue(mhz)} MHz`;
}

function makeId(ts) {
  return globalThis.crypto?.randomUUID?.() ?? `${ts}-${Math.random()}`;
}

function qsoKey(qso) {
  if (!qso?.callsign || !qso?.dt) return "";
  return `${qso.callsign}|${qso.dt}|${qso.band || ""}|${qso.freq || ""}|${qso.mode || ""}|${qso.myGrid || ""}|${qso.theirGrid || ""}`.toUpperCase();
}

function hasDuplicate(qso) {
  const key = qsoKey(qso);
  return qsos.some(x => x.id === qso.id || (key && qsoKey(x) === key));
}

function insertQso(qso) {
  if (!qso) return false;
  if (hasDuplicate(qso)) return false;
  qsos.push(qso);
  return true;
}

function qsoEquals(a, b) {
  if (!a || !b) return false;
  return (
    a.id === b.id &&
    a.callsign === b.callsign &&
    a.dt === b.dt &&
    (a.band || "") === (b.band || "") &&
    (a.freq || "") === (b.freq || "") &&
    (a.mode || "") === (b.mode || "") &&
    (a.setup || "") === (b.setup || "") &&
    (a.myGrid || "") === (b.myGrid || "") &&
    (a.theirGrid || "") === (b.theirGrid || "") &&
    (a.rstS || "") === (b.rstS || "") &&
    (a.rstR || "") === (b.rstR || "") &&
    (a.notes || "") === (b.notes || "") &&
    Number(a.ts || 0) === Number(b.ts || 0)
  );
}

function upsertQso(qso) {
  if (!qso) return false;
  const idx = qsos.findIndex(x => x.id === qso.id);
  if (idx >= 0) {
    if (qsoEquals(qsos[idx], qso)) return false;
    qsos[idx] = qso;
    return true;
  }
  return insertQso(qso);
}

function removeQsoById(id) {
  const before = qsos.length;
  qsos = qsos.filter(x => x.id !== id);
  return qsos.length !== before;
}

function sortQsos() {
  qsos.sort((a, b) => (b.ts || 0) - (a.ts || 0));
}

function saveLocalQsos() {
  return writeStoredArray(LOCAL_KEY, qsos, "QSO log");
}

function loadLocalQsos() {
  const data = readStoredArray(LOCAL_KEY, "QSO log");
  qsos = data.map(normalizeQso).filter(Boolean);
  sortQsos();
}

function applyUpdate(update) {
  const p = normalizeUpdatePayload(update?.payload);
  if (!p) return false;
  let changed = false;

  if (p.type === "add_qso" && p.qso) {
    // Avoid duplicates by id
    if (insertQso(p.qso)) changed = true;
  }

  if (p.type === "edit_qso" && p.qso) {
    if (upsertQso(p.qso)) changed = true;
  }

  if (p.type === "bulk_add" && Array.isArray(p.qsos)) {
    for (const qso of p.qsos) {
      if (insertQso(qso)) changed = true;
    }
  }

  if (p.type === "delete_qso" && p.id) {
    if (removeQsoById(p.id)) changed = true;
  }

  if (changed) {
    sortQsos();
  }
  return changed;
}

function render() {
  const list = document.getElementById("list");
  const q = (document.getElementById("search").value || "").toLowerCase();

  const filtered = qsos.filter(x => {
    const hay = `${x.callsign} ${x.band} ${x.freq} ${x.mode} ${x.myGrid} ${x.theirGrid} ${x.setup} ${x.notes}`.toLowerCase();
    return hay.includes(q);
  });

  if (!filtered.length) {
    list.innerHTML = `<div class="empty">No QSOs yet.</div>`;
    return;
  }

  list.innerHTML = filtered.map(x => {
    const bandVal = x.band ? escapeHtml(x.band) : "";
    const freqVal = x.freq ? escapeHtml(formatFreqDisplay(x.freq)) : "";
    const bandFreq = bandVal && freqVal ? `${bandVal} • ${freqVal}` : (bandVal || freqVal || "-");
    const grids = [x.myGrid, x.theirGrid].filter(Boolean).map(escapeHtml).join(" • ");
    const setup = x.setup ? `Setup: ${escapeHtml(x.setup)}` : "";
    return `
    <div class="item">
      <b>${escapeHtml(x.callsign)}</b>
      <div class="meta">
        ${escapeHtml(formatDisplayDT(x.dt))} • ${bandFreq} • ${escapeHtml(x.mode || "-")}
        • RST ${escapeHtml(x.rstS || "-")}/${escapeHtml(x.rstR || "-")}
        ${grids ? ` • ${grids}` : ""}
        ${setup ? ` • ${setup}` : ""}
      </div>
      ${x.notes ? `<div class="notes">${escapeHtml(x.notes)}</div>` : ""}
      <div class="actions">
        <button type="button" data-action="edit" data-id="${escapeHtml(x.id)}">Edit</button>
        <button type="button" class="${x.id === pendingDeleteId ? "danger" : "ghost"}" data-action="delete" data-id="${escapeHtml(x.id)}">${x.id === pendingDeleteId ? "Confirm" : "Delete"}</button>
      </div>
    </div>
  `;
  }).join("");
}

function addQso(qso, descriptionOverride) {
  const added = insertQso(qso);
  if (added) {
    sortQsos();
    saveLocalQsos();
  }
  render();

  if (added) {
    enqueueUpdate(
      { type: "add_qso", qso },
      descriptionOverride || `QSO ${qso.callsign} ${qso.band || ""} ${qso.mode || ""}`.trim()
    );
  }
  return added;
}

function addQsosBatch(qsoList, descriptionOverride) {
  if (!qsoList.length) return 0;
  const insertedQsos = [];
  for (const qso of qsoList) {
    if (insertQso(qso)) insertedQsos.push(qso);
  }
  const changed = insertedQsos.length > 0;
  if (changed) {
    sortQsos();
    saveLocalQsos();
  }
  render();
  if (!changed) return 0;

  const maxSize = getSendUpdateMaxSize();
  const chunks = [];
  const singleUpdates = [];
  let current = [];
  for (const qso of insertedQsos) {
    const candidate = current.concat(qso);
    const candidateInfo = descriptionOverride || `Imported ${candidate.length} QSO(s)`;
    const candidateUpdate = makeUpdate({ type: "bulk_add", qsos: candidate }, candidateInfo);
    if (updateSizeBytes(candidateUpdate) <= maxSize) {
      current = candidate;
      continue;
    }

    if (current.length) {
      chunks.push(current);
      current = [];
    }

    const singleInfo = descriptionOverride || "Imported 1 QSO";
    const singleBulkUpdate = makeUpdate({ type: "bulk_add", qsos: [qso] }, singleInfo);
    if (updateSizeBytes(singleBulkUpdate) <= maxSize) {
      current = [qso];
    } else {
      singleUpdates.push(qso);
    }
  }
  if (current.length) chunks.push(current);

  for (const chunk of chunks) {
    enqueueUpdate(
      { type: "bulk_add", qsos: chunk },
      descriptionOverride || `Imported ${chunk.length} QSO(s)`
    );
  }

  for (const qso of singleUpdates) {
    enqueueUpdate(
      { type: "add_qso", qso },
      descriptionOverride || `Imported ${qso.callsign || "QSO"}`
    );
  }
  return insertedQsos.length;
}

function editQso(qso, descriptionOverride) {
  const changed = upsertQso(qso);
  if (changed) {
    sortQsos();
    saveLocalQsos();
  }
  render();
  if (changed) {
    enqueueUpdate(
      { type: "edit_qso", qso },
      descriptionOverride || `Edited QSO ${qso.callsign}`
    );
  }
  return changed;
}

function deleteQso(id, descriptionOverride) {
  if (!id) return;
  const changed = removeQsoById(id);
  if (changed) {
    sortQsos();
    saveLocalQsos();
  }
  render();
  if (changed) {
    enqueueUpdate(
      { type: "delete_qso", id },
      descriptionOverride || "Deleted QSO"
    );
  }
  return changed;
}

function downloadText(filename, text, mime = "text/plain") {
  const blob = new Blob([text], { type: mime });
  const url = URL.createObjectURL(blob);
  const a = document.createElement("a");
  a.href = url;
  a.download = filename;
  document.body.appendChild(a);
  a.click();
  a.remove();
  URL.revokeObjectURL(url);
}

function sendFileToChat(filename, text, mime, label) {
  if (!window.webxdc?.sendToChat) return false;
  window.webxdc.sendToChat({
    text: label || `Exported ${filename}`,
    file: { name: filename, plainText: text }
  }).catch(() => {
    // Fallback to download if sending fails
    downloadText(filename, text, mime);
  });
  return true;
}

function csvEscape(value) {
  const s = (value ?? "").toString();
  if (/[",\n\r]/.test(s)) return `"${s.replace(/"/g, '""')}"`;
  return s;
}

function exportCsv() {
  const header = ["callsign","dt","band","freq","mode","setup","myGrid","theirGrid","rstS","rstR","notes","id","ts"];
  const rows = qsos
    .slice()
    .sort((a,b) => (a.ts||0)-(b.ts||0)) // oldest first for export
    .map(q => header.map(k => csvEscape(q[k])).join(","));
  const out = [header.join(","), ...rows].join("\n");
  if (!sendFileToChat("hamlog.csv", out, "text/csv", "Ham Log CSV export")) {
    downloadText("hamlog.csv", out, "text/csv");
  }
  setStatus("Exported CSV.");
}

function parseCsv(text) {
  const lines = text.replace(/\r\n/g, "\n").replace(/\r/g, "\n").split("\n").filter(l => l.trim().length);
  if (!lines.length) return [];

  const parseLine = (line) => {
    const out = [];
    let cur = "";
    let inQuotes = false;
    for (let i=0; i<line.length; i++) {
      const ch = line[i];
      if (inQuotes) {
        if (ch === '"' && line[i+1] === '"') { cur += '"'; i++; }
        else if (ch === '"') inQuotes = false;
        else cur += ch;
      } else {
        if (ch === '"') inQuotes = true;
        else if (ch === ",") { out.push(cur); cur = ""; }
        else cur += ch;
      }
    }
    out.push(cur);
    return out;
  };

  const header = parseLine(lines[0]).map((h, idx) => (idx === 0 ? h.replace(/^\uFEFF/, "") : h).trim());
  const records = [];
  for (let i=1; i<lines.length; i++) {
    const cols = parseLine(lines[i]);
    const obj = {};
    for (let j=0; j<header.length; j++) obj[header[j]] = cols[j] ?? "";
    records.push(obj);
  }
  return records;
}

function importCsv(text) {
  const rows = parseCsv(text);
  const qsoList = [];

  for (const r of rows) {
    const callsign = (r.callsign || "").trim().toUpperCase();
    const dt = (r.dt || "").trim();
    if (!callsign || !dt) continue;

    const parsedTs = tsFromDT(dt);
    const ts = Number(r.ts) || parsedTs || Date.now();
    const qso = {
      id: (r.id && r.id.trim()) ? r.id.trim() : makeId(ts),
      callsign,
      dt,
      band: (r.band || "").trim(),
      freq: (r.freq || "").trim(),
      mode: (r.mode || "").trim().toUpperCase(),
      setup: (r.setup || "").trim(),
      myGrid: (r.myGrid || "").trim().toUpperCase(),
      theirGrid: (r.theirGrid || "").trim().toUpperCase(),
      rstS: (r.rstS || "").trim(),
      rstR: (r.rstR || "").trim(),
      notes: (r.notes || "").trim(),
      ts
    };
    qsoList.push(qso);
  }

  const inserted = addQsosBatch(qsoList);
  return inserted;
}

function setStatus(msg) {
  const el = document.getElementById("importStatus");
  if (el) el.textContent = msg || "";
}

function init() {
  // Default datetime = now (local time)
  const dt = getDtElement();
  setDtNow();
  dt.addEventListener("input", handleDtInput);
  setInterval(refreshDtIfAuto, AUTO_DT_REFRESH_MS);

  document.getElementById("qsoForm").addEventListener("submit", (e) => {
    e.preventDefault();
    const callsign = document.getElementById("callsign").value.trim().toUpperCase();
    const dtLocal = document.getElementById("dt").value;
    if (!callsign || !dtLocal) return;

    const band = document.getElementById("band").value.trim();
    const freq = document.getElementById("freq").value.trim();
    const mode = document.getElementById("mode").value.trim().toUpperCase();
    const setup = document.getElementById("setup").value.trim();
    const myGrid = document.getElementById("myGrid").value.trim().toUpperCase();
    const theirGrid = document.getElementById("theirGrid").value.trim().toUpperCase();
    const rstS = document.getElementById("rstS").value.trim();
    const rstR = document.getElementById("rstR").value.trim();
    const notes = document.getElementById("notes").value.trim();

    const ts = tsFromLocalDT(dtLocal) || Date.now();
    const qso = {
      id: editingId || makeId(ts),
      callsign,
      dt: dtLocal,
      band,
      freq,
      mode,
      setup,
      myGrid,
      theirGrid,
      rstS,
      rstR,
      notes,
      ts
    };
    if (editingId) editQso(qso);
    else addQso(qso);
    // Keep commonly reused fields after submission.
    const keep = { band, freq, mode, setup, myGrid };
    setEditing(null);
    document.getElementById("band").value = keep.band;
    document.getElementById("freq").value = keep.freq;
    document.getElementById("mode").value = keep.mode;
    document.getElementById("setup").value = keep.setup;
    document.getElementById("myGrid").value = keep.myGrid;
  });

  document.getElementById("search").addEventListener("input", render);

  document.getElementById("exportCsvBtn").addEventListener("click", exportCsv);

  document.getElementById("importFile").addEventListener("change", async (e) => {
    const file = e.target.files?.[0];
    if (!file) return;

    setStatus(`Importing ${file.name}…`);
    const text = await file.text();
    const name = file.name.toLowerCase();

    try {
      if (!name.endsWith(".csv")) {
        setStatus("Only CSV files are supported.");
        return;
      }
      const n = importCsv(text);
      setStatus(`Imported ${n} QSO(s).`);
    } catch (err) {
      console.error(err);
      setStatus(`Import failed: ${err?.message || err}`);
    } finally {
      // reset so importing same file again triggers change
      e.target.value = "";
    }
  });

  document.getElementById("importTextBtn").addEventListener("click", () => {
    const el = document.getElementById("importText");
    const text = el.value.trim();
    if (!text) return;
    try {
      const n = importCsv(text);
      setStatus(`Imported ${n} QSO(s) from pasted text.`);
      el.value = "";
    } catch (err) {
      console.error(err);
      setStatus(`Import failed: ${err?.message || err}`);
    }
  });

  document.getElementById("cancelEdit").addEventListener("click", () => {
    setEditing(null);
  });

  document.getElementById("list").addEventListener("click", (e) => {
    const btn = e.target.closest("button[data-action]");
    if (!btn) return;
    const action = btn.getAttribute("data-action");
    const id = btn.getAttribute("data-id");
    if (!id) return;

    if (action === "delete") {
      if (pendingDeleteId === id) {
        pendingDeleteId = null;
        if (pendingDeleteTimer) clearTimeout(pendingDeleteTimer);
        deleteQso(id);
      } else {
        pendingDeleteId = id;
        if (pendingDeleteTimer) clearTimeout(pendingDeleteTimer);
        pendingDeleteTimer = setTimeout(() => {
          pendingDeleteId = null;
          render();
        }, 4000);
        render();
      }
      return;
    }
    if (action === "edit") {
      const qso = qsos.find(x => x.id === id);
      if (!qso) return;
      setEditing(qso);
    }
  });

  loadLocalQsos();
  loadOutbox();
  loadLastAppliedSerial();
  updatePendingStatus();
  setSyncStatus("checking");

  if (window.webxdc?.setUpdateListener) {
    try {
      const startSerial = Number.isFinite(lastAppliedSerial) && lastAppliedSerial > 0 ? lastAppliedSerial : 0;
      const ready = window.webxdc.setUpdateListener((update) => {
        const serial = Number(update?.serial);
        if (Number.isFinite(serial) && serial <= lastAppliedSerial) return;
        if (Number.isFinite(serial) && seenSerials.has(serial)) return;
        if (Number.isFinite(serial) && serial > 0) seenSerials.add(serial);
        const changed = applyUpdate(update);
        if (changed) saveLocalQsos();
        if (Number.isFinite(serial) && serial > 0) saveLastAppliedSerial(serial);
        render();
      }, startSerial);
      if (ready && typeof ready.catch === "function") {
        ready.catch(() => setSyncStatus("listener failed"));
      }
    } catch (_) {
      setSyncStatus("listener failed");
    }
  } else {
    setSyncStatus("local mode");
  }

  setInterval(flushOutbox, OUTBOX_FLUSH_INTERVAL_MS);
  document.addEventListener("visibilitychange", () => {
    if (document.visibilityState === "visible") flushOutbox();
  });
  flushOutbox();
  render();
}

init();

function setEditing(qso) {
  const form = document.getElementById("qsoForm");
  const saveBtn = document.getElementById("saveBtn");
  const cancelBtn = document.getElementById("cancelEdit");
  if (!qso) {
    editingId = null;
    dtManualOverride = false;
    pendingDeleteId = null;
    if (pendingDeleteTimer) clearTimeout(pendingDeleteTimer);
    form.reset();
    setDtNow();
    saveBtn.textContent = "Save QSO";
    cancelBtn.classList.add("hidden");
    return;
  }
  editingId = qso.id;
  dtManualOverride = true;
  document.getElementById("callsign").value = qso.callsign || "";
  const dtValue = qso.dt || toLocalInputValue(new Date());
  document.getElementById("dt").value = dtValue;
  lastAutoDtValue = dtValue;
  document.getElementById("band").value = qso.band || "";
  document.getElementById("freq").value = qso.freq || "";
  document.getElementById("mode").value = qso.mode || "";
  document.getElementById("setup").value = qso.setup || "";
  document.getElementById("myGrid").value = qso.myGrid || "";
  document.getElementById("theirGrid").value = qso.theirGrid || "";
  document.getElementById("rstS").value = qso.rstS || "";
  document.getElementById("rstR").value = qso.rstR || "";
  document.getElementById("notes").value = qso.notes || "";
  saveBtn.textContent = "Update QSO";
  cancelBtn.classList.remove("hidden");
}
