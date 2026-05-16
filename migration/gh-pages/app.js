const storageKeys = {
  apiUrl: "selfStudy.apiUrl",
  writeToken: "selfStudy.writeToken",
};

const el = {
  apiUrl: document.getElementById("apiUrl"),
  writeToken: document.getElementById("writeToken"),
  saveConfigBtn: document.getElementById("saveConfigBtn"),
  healthBtn: document.getElementById("healthBtn"),
  status: document.getElementById("status"),
  logForm: document.getElementById("logForm"),
  recordDate: document.getElementById("recordDate"),
  duration: document.getElementById("duration"),
  studentId: document.getElementById("studentId"),
  studentName: document.getElementById("studentName"),
  subject: document.getElementById("subject"),
  memo: document.getElementById("memo"),
  fromDate: document.getElementById("fromDate"),
  toDate: document.getElementById("toDate"),
  filterStudentId: document.getElementById("filterStudentId"),
  loadBtn: document.getElementById("loadBtn"),
  logTableBody: document.querySelector("#logTable tbody"),
};

function todayIso() {
  return new Date().toISOString().slice(0, 10);
}

function currentMonthRange() {
  const now = new Date();
  const y = now.getFullYear();
  const m = now.getMonth();
  const first = new Date(y, m, 1).toISOString().slice(0, 10);
  const last = new Date(y, m + 1, 0).toISOString().slice(0, 10);
  return { first, last };
}

function setStatus(message, kind) {
  el.status.textContent = message;
  el.status.className = "status";
  if (kind === "ok") el.status.classList.add("ok");
  if (kind === "err") el.status.classList.add("err");
}

function getApiUrl() {
  return el.apiUrl.value.trim();
}

function getToken() {
  return el.writeToken.value.trim();
}

function assertApiUrl() {
  const apiUrl = getApiUrl();
  if (!apiUrl) {
    throw new Error("Set Apps Script Web App URL first.");
  }
  return apiUrl;
}

async function apiGet(params) {
  const apiUrl = assertApiUrl();
  const qs = new URLSearchParams(params);
  const res = await fetch(`${apiUrl}?${qs.toString()}`);
  if (!res.ok) {
    throw new Error(`HTTP ${res.status}`);
  }
  return res.json();
}

async function apiPost(action, payload) {
  const apiUrl = assertApiUrl();
  const body = new URLSearchParams();
  body.set("action", action);
  body.set("payload", JSON.stringify(payload));

  const res = await fetch(apiUrl, {
    method: "POST",
    body,
  });
  if (!res.ok) {
    throw new Error(`HTTP ${res.status}`);
  }
  return res.json();
}

function loadConfigFromStorage() {
  const savedUrl = localStorage.getItem(storageKeys.apiUrl);
  const savedToken = localStorage.getItem(storageKeys.writeToken);
  if (savedUrl) el.apiUrl.value = savedUrl;
  if (savedToken) el.writeToken.value = savedToken;
}

function saveConfigToStorage() {
  localStorage.setItem(storageKeys.apiUrl, getApiUrl());
  localStorage.setItem(storageKeys.writeToken, getToken());
  setStatus("Config saved locally.", "ok");
}

function renderRows(rows) {
  el.logTableBody.innerHTML = "";
  if (!rows.length) {
    const tr = document.createElement("tr");
    tr.innerHTML = '<td colspan="7">No rows</td>';
    el.logTableBody.appendChild(tr);
    return;
  }

  rows.forEach((row) => {
    const tr = document.createElement("tr");
    const student = [row.studentId || "-", row.studentName || "-"].join(" / ");
    const actor = row.actorEmail || row.actorName || "-";
    tr.innerHTML = `
      <td>${escapeHtml(row.createdAt || "-")}</td>
      <td>${escapeHtml(row.recordDate || "-")}</td>
      <td>${escapeHtml(actor)}</td>
      <td>${escapeHtml(student)}</td>
      <td>${escapeHtml(row.subject || "-")}</td>
      <td>${escapeHtml(String(row.duration || "-"))}</td>
      <td>${escapeHtml(row.memo || "-")}</td>
    `;
    el.logTableBody.appendChild(tr);
  });
}

function escapeHtml(value) {
  return value
    .replaceAll("&", "&amp;")
    .replaceAll("<", "&lt;")
    .replaceAll(">", "&gt;")
    .replaceAll('"', "&quot;")
    .replaceAll("'", "&#39;");
}

async function checkHealth() {
  try {
    setStatus("Checking health...");
    const data = await apiGet({ action: "health" });
    if (!data.ok) throw new Error(data.error || "health check failed");
    setStatus(`Connected. timezone=${data.timezone}`, "ok");
  } catch (err) {
    setStatus(err.message, "err");
  }
}

async function submitLog(event) {
  event.preventDefault();
  try {
    setStatus("Saving log...");
    const payload = {
      token: getToken(),
      date: el.recordDate.value,
      duration: Number(el.duration.value),
      studentId: el.studentId.value.trim(),
      studentName: el.studentName.value.trim(),
      subject: el.subject.value.trim(),
      memo: el.memo.value.trim(),
      actorName: "github-pages-client",
      clientTimestamp: new Date().toISOString(),
      clientTag: location.origin || "github-pages",
    };

    const data = await apiPost("appendLog", payload);
    if (!data.ok) throw new Error(data.error || "append failed");
    setStatus(
      `Saved to ${data.monthKey} / ${data.sheetName}. totalRows=${data.totalRows}`,
      "ok",
    );
    el.subject.value = "";
    el.duration.value = "";
    el.memo.value = "";
  } catch (err) {
    setStatus(err.message, "err");
  }
}

async function loadLogs() {
  try {
    setStatus("Loading logs...");
    const params = {
      action: "getLogs",
      from: el.fromDate.value,
      to: el.toDate.value,
      studentId: el.filterStudentId.value.trim(),
      token: getToken(),
    };
    const data = await apiGet(params);
    if (!data.ok) throw new Error(data.error || "load failed");
    renderRows(data.rows || []);
    setStatus(`Loaded ${data.rows.length} rows.`, "ok");
  } catch (err) {
    setStatus(err.message, "err");
  }
}

function initDefaultDates() {
  el.recordDate.value = todayIso();
  const { first, last } = currentMonthRange();
  el.fromDate.value = first;
  el.toDate.value = last;
}

function bindEvents() {
  el.saveConfigBtn.addEventListener("click", saveConfigToStorage);
  el.healthBtn.addEventListener("click", checkHealth);
  el.logForm.addEventListener("submit", submitLog);
  el.loadBtn.addEventListener("click", loadLogs);
}

function init() {
  loadConfigFromStorage();
  initDefaultDates();
  bindEvents();
  setStatus("Ready");
}

init();
