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
let editingId = null;
let pendingDeleteId = null;
let pendingDeleteTimer = null;

function isWebxdc() {
  return typeof window.webxdc !== "undefined";
}

function canSendUpdate() {
  return !!window.webxdc?.sendUpdate;
}

function sendUpdateCompat(payload, info) {
  const update = { payload };
  if (info) update.info = info;
  // Spec: second argument is deprecated; pass empty string for compatibility.
  if (window.webxdc?.sendUpdate) return window.webxdc.sendUpdate(update, "");
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

function upsertQso(qso) {
  if (!qso) return false;
  const idx = qsos.findIndex(x => x.id === qso.id);
  if (idx >= 0) {
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
  try {
    localStorage.setItem(LOCAL_KEY, JSON.stringify(qsos));
  } catch (_) {}
}

function loadLocalQsos() {
  try {
    const raw = localStorage.getItem(LOCAL_KEY);
    if (!raw) return;
    const data = JSON.parse(raw);
    if (Array.isArray(data)) {
      qsos = data.filter(Boolean);
      sortQsos();
    }
  } catch (_) {}
}

function applyUpdate(update) {
  const p = update?.payload;
  if (!p?.type) return;
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
  if (!canSendUpdate()) {
    if (insertQso(qso)) {
      sortQsos();
      saveLocalQsos();
    }
    render();
    return;
  }

  sendUpdateCompat(
    { type: "add_qso", qso },
    descriptionOverride || `QSO ${qso.callsign} ${qso.band || ""} ${qso.mode || ""}`.trim()
  );
}

function addQsosBatch(qsoList, descriptionOverride) {
  if (!qsoList.length) return;
  if (!canSendUpdate()) {
    let changed = false;
    for (const qso of qsoList) {
      if (insertQso(qso)) changed = true;
    }
    if (changed) {
      sortQsos();
      saveLocalQsos();
      render();
    }
    return;
  }
  const maxSize = window.webxdc?.sendUpdateMaxSize || 128000;
  const chunks = [];
  let current = [];
  for (const qso of qsoList) {
    current.push(qso);
    const size = JSON.stringify({ payload: { type: "bulk_add", qsos: current } }).length;
    if (size > maxSize && current.length > 1) {
      current.pop();
      chunks.push(current);
      current = [qso];
    }
  }
  if (current.length) chunks.push(current);

  for (const chunk of chunks) {
    sendUpdateCompat(
      { type: "bulk_add", qsos: chunk },
      descriptionOverride || `Imported ${chunk.length} QSO(s)`
    );
  }
}

function editQso(qso, descriptionOverride) {
  if (!canSendUpdate()) {
    if (upsertQso(qso)) {
      sortQsos();
      saveLocalQsos();
      render();
    }
    return;
  }
  sendUpdateCompat(
    { type: "edit_qso", qso },
    descriptionOverride || `Edited QSO ${qso.callsign}`
  );
}

function deleteQso(id, descriptionOverride) {
  if (!id) return;
  if (!canSendUpdate()) {
    if (removeQsoById(id)) {
      sortQsos();
      saveLocalQsos();
      render();
    }
    return;
  }
  sendUpdateCompat(
    { type: "delete_qso", id },
    descriptionOverride || "Deleted QSO"
  );
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

// ---------- CSV ----------
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

// Basic CSV parser (handles quoted fields)
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

  const header = parseLine(lines[0]).map(h => h.trim());
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
      id: (r.id && r.id.trim()) ? r.id.trim() : (crypto.randomUUID?.() ?? `${ts}-${Math.random()}`),
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

  addQsosBatch(qsoList, `Imported ${qsoList.length} QSO(s)`);
  return qsoList.length;
}

// ---------- UI wiring ----------
function setStatus(msg) {
  const el = document.getElementById("importStatus");
  if (el) el.textContent = msg || "";
}

function init() {
  // Default datetime = now (local time)
  const dt = document.getElementById("dt");
  dt.value = toLocalInputValue(new Date());

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
      id: editingId || (crypto.randomUUID?.() ?? `${ts}-${Math.random()}`),
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

  if (window.webxdc?.setUpdateListener) {
    window.webxdc.setUpdateListener((update) => {
      if (update.serial && seenSerials.has(update.serial)) return;
      if (update.serial) seenSerials.add(update.serial);
      applyUpdate(update);
      render();
    }, 0);
  } else {
    loadLocalQsos();
  }

  render();
}

init();

function setEditing(qso) {
  const form = document.getElementById("qsoForm");
  const saveBtn = document.getElementById("saveBtn");
  const cancelBtn = document.getElementById("cancelEdit");
  if (!qso) {
    editingId = null;
    pendingDeleteId = null;
    if (pendingDeleteTimer) clearTimeout(pendingDeleteTimer);
    form.reset();
    document.getElementById("dt").value = toLocalInputValue(new Date());
    saveBtn.textContent = "Save QSO";
    cancelBtn.classList.add("hidden");
    return;
  }
  editingId = qso.id;
  document.getElementById("callsign").value = qso.callsign || "";
  document.getElementById("dt").value = qso.dt || toLocalInputValue(new Date());
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
