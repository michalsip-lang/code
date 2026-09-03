const DB_NAME = "akcni-plan-local";
const DB_VERSION = 2;
const STORE = "tasks";
const LS_KEY = "akcni-plan-local-tasks";
const RECOVERY_KEY = "akcni-plan-recovery-attempted";
const SYNC_URL_KEY = "akcni-plan-sync-url";
const SYNC_ANON_KEY = "akcni-plan-sync-anon-key";
const SYNC_TOKEN_KEY = "akcni-plan-sync-access-token";
const SYNC_PENDING_ACTION_KEY = "akcni-plan-sync-pending-action";
const DEFAULT_SYNC_URL = "https://vpjgpcnvpwarvcxfoteo.supabase.co";
const DEFAULT_SYNC_ANON_KEY = "eyJhbGciOiJIUzI1NiIsInR5cCI6IkpXVCJ9.eyJpc3MiOiJzdXBhYmFzZSIsInJlZiI6InZwamdwY252cHdhcnZjeGZvdGVvIiwicm9sZSI6ImFub24iLCJpYXQiOjE3ODg0NDM0MDksImV4cCI6MjEwNDAxOTQwOX0.5bgXCFJZ-gfFfjb7Ua2dmpKU8KMGnyFFtNY3dTUAJPs";

const AREA_ORDER = ["Svp", "Sdp", "Bozp", "Po", "Jine"];
const STATUS_ORDER = ["Todo", "InProgress", "Done", "Blocked"];

const AREA_LABEL = {
  Svp: "SVP",
  Sdp: "SDP",
  Bozp: "BOZP",
  Po: "PO",
  Jine: "Jine"
};

const STATUS_LABEL = {
  Todo: "K vyrizeni",
  InProgress: "Rozpracovano",
  Done: "Hotovo",
  Blocked: "Blokovano"
};

const FILTER_LABEL = {
  overdue: "Po terminu",
  today: "Na dnesek",
  week: "Tento tyden",
  completed: "Dokoncene",
  inprogress: "Rozpracovane"
};

const brand = {
  blue: "#292982",
  grey: "#808184",
  red: "#e01b37",
  blueSoft: "rgba(41, 41, 130, 0.16)",
  greySoft: "rgba(128, 129, 132, 0.20)",
  redSoft: "rgba(224, 27, 55, 0.18)"
};

let db;
let tasks = [];
let activeFilter = null;
let editTaskId = null;
let charts = [];
let storageMode = "indexeddb";
let syncConfig = { url: DEFAULT_SYNC_URL, anonKey: DEFAULT_SYNC_ANON_KEY };
let authState = { accessToken: "", userId: "", email: "" };

init().catch(async (error) => {
  console.error(error);
  const recovered = await tryClientRecovery(error);
  if (!recovered) {
    alert("Aplikaci se nepodarilo inicializovat. Na iPadu zkuste vypnout Soukrome prohlizeni nebo povolit data webu pro Safari.");
  }
});

async function init() {
  ensureDomContract();
  setupNavigation();
  setupCreateForm();
  setupAutoForm();
  setupSyncPanel();
  setupServiceWorker();

  try {
    db = await openDb();
    storageMode = "indexeddb";
  } catch (error) {
    console.warn("IndexedDB neni dostupna, prepinam na localStorage", error);
    storageMode = "localstorage";
  }

  tasks = await loadTasks();

  renderAll();
}

function ensureDomContract() {
  const requiredIds = [
    "task-form", "auto-form", "kpi-grid", "area-picker", "area-panels", "recommendations", "top-priority-body", "heatmap",
    "sync-form", "supabase-url", "supabase-key", "sync-push", "sync-pull", "sync-status",
    "auth-email", "auth-password", "auth-signup", "auth-login", "auth-github", "auth-logout", "auth-status"
  ];
  requiredIds.forEach((id) => {
    if (!document.getElementById(id)) {
      throw new Error(`Missing required element: ${id}`);
    }
  });
}

async function tryClientRecovery(error) {
  try {
    if (sessionStorage.getItem(RECOVERY_KEY) === "1") {
      return false;
    }

    const message = String(error?.message || error || "").toLowerCase();
    const shouldRecover = message.includes("missing required element")
      || message.includes("indexeddb")
      || message.includes("quota")
      || message.includes("invalidstateerror")
      || message.includes("notfounderror");

    if (!shouldRecover) {
      return false;
    }

    sessionStorage.setItem(RECOVERY_KEY, "1");

    if ("serviceWorker" in navigator) {
      const registrations = await navigator.serviceWorker.getRegistrations();
      await Promise.all(registrations.map((registration) => registration.unregister()));
    }

    if ("caches" in window) {
      const keys = await caches.keys();
      await Promise.all(keys.map((key) => caches.delete(key)));
    }

    const url = new URL(window.location.href);
    url.searchParams.set("v", "5");
    url.searchParams.set("t", String(Date.now()));
    window.location.replace(url.toString());
    return true;
  } catch {
    return false;
  }
}

function openDb() {
  return new Promise((resolve, reject) => {
    const request = indexedDB.open(DB_NAME, DB_VERSION);

    request.onupgradeneeded = () => {
      const nextDb = request.result;
      let store;
      if (!nextDb.objectStoreNames.contains(STORE)) {
        store = nextDb.createObjectStore(STORE, { keyPath: "id" });
      } else {
        store = request.transaction.objectStore(STORE);
      }

      if (!store.indexNames.contains("status")) {
        store.createIndex("status", "status", { unique: false });
      }
      if (!store.indexNames.contains("createdAt")) {
        store.createIndex("createdAt", "createdAt", { unique: false });
      }
    };

    request.onsuccess = () => resolve(request.result);
    request.onerror = () => reject(request.error);
  });
}

function tx(mode = "readonly") {
  return db.transaction(STORE, mode).objectStore(STORE);
}

function loadTasks() {
  if (storageMode === "localstorage") {
    return Promise.resolve(loadTasksFromLocalStorage());
  }

  return new Promise((resolve, reject) => {
    const request = tx().getAll();
    request.onsuccess = () => {
      const rows = (request.result || []).map(normalizeTask);
      rows.forEach((task) => {
        task.priorityScore = calculatePriority(task);
      });
      resolve(rows);
    };
    request.onerror = () => reject(request.error);
  });
}

function persistTask(task) {
  if (storageMode === "localstorage") {
    saveTasksToLocalStorage(tasks);
    return Promise.resolve();
  }

  return new Promise((resolve, reject) => {
    const request = tx("readwrite").put(task);
    request.onsuccess = () => resolve();
    request.onerror = () => reject(request.error);
  });
}

function deleteTaskById(id) {
  if (storageMode === "localstorage") {
    tasks = tasks.filter((row) => row.id !== id);
    saveTasksToLocalStorage(tasks);
    return Promise.resolve();
  }

  return new Promise((resolve, reject) => {
    const request = tx("readwrite").delete(id);
    request.onsuccess = () => resolve();
    request.onerror = () => reject(request.error);
  });
}

function clearStore() {
  if (storageMode === "localstorage") {
    try {
      localStorage.removeItem(LS_KEY);
    } catch {
    }
    return Promise.resolve();
  }

  return new Promise((resolve, reject) => {
    const request = tx("readwrite").clear();
    request.onsuccess = () => resolve();
    request.onerror = () => reject(request.error);
  });
}

function loadTasksFromLocalStorage() {
  try {
    const raw = localStorage.getItem(LS_KEY);
    if (!raw) {
      return [];
    }

    const parsed = JSON.parse(raw);
    if (!Array.isArray(parsed)) {
      return [];
    }

    return parsed.map(normalizeTask);
  } catch {
    return [];
  }
}

function setupSyncPanel() {
  syncConfig = loadSyncConfig();
  authState = loadAuthState();
  hydrateAuthFromUrlHash();

  const form = document.getElementById("sync-form");
  const urlInput = document.getElementById("supabase-url");
  const keyInput = document.getElementById("supabase-key");
  const emailInput = document.getElementById("auth-email");
  const passwordInput = document.getElementById("auth-password");
  const signUpButton = document.getElementById("auth-signup");
  const loginButton = document.getElementById("auth-login");
  const githubButton = document.getElementById("auth-github");
  const logoutButton = document.getElementById("auth-logout");
  const pushButton = document.getElementById("sync-push");
  const pullButton = document.getElementById("sync-pull");

  urlInput.value = syncConfig.url;
  keyInput.value = syncConfig.anonKey;
  urlInput.readOnly = true;
  keyInput.readOnly = true;
  emailInput.value = authState.email || "";

  form.addEventListener("submit", (event) => {
    event.preventDefault();
    syncConfig = {
      url: String(urlInput.value || "").trim().replace(/\/$/, ""),
      anonKey: String(keyInput.value || "").trim()
    };

    saveSyncConfig(syncConfig);
    updateSyncStatus(syncConfigReady() ? "Nastaveni sync ulozeno." : "Dopln Supabase URL a Anon key.", !syncConfigReady());
    refreshAuthStatus();
  });

  signUpButton.addEventListener("click", async () => {
    await runAuthAction(async () => {
      const email = String(emailInput.value || "").trim();
      const password = String(passwordInput.value || "").trim();
      await signUpSupabase(email, password);
      updateSyncStatus("Ucet vytvoren. Ted klikni Prihlasit.");
    });
  });

  loginButton.addEventListener("click", async () => {
    await runAuthAction(async () => {
      const email = String(emailInput.value || "").trim();
      const password = String(passwordInput.value || "").trim();
      await loginSupabase(email, password);
      refreshAuthStatus();
      updateSyncStatus("Prihlaseni uspesne.");
    });
  });

  githubButton.addEventListener("click", async () => {
    await runAuthAction(async () => {
      startGithubOAuth();
    });
  });

  logoutButton.addEventListener("click", async () => {
    await runAuthAction(async () => {
      clearAuthState();
      refreshAuthStatus();
      updateSyncStatus("Odhlaseno.");
    });
  });

  pushButton.addEventListener("click", async () => {
    await runSyncAction(pushToCloud, "push");
  });

  pullButton.addEventListener("click", async () => {
    await runSyncAction(pullFromCloud, "pull");
  });

  refreshAuthStatus();
  updateSyncStatus(syncConfigReady() ? "Cloud sync pripraven." : "Cloud sync neni nastaven.");
  resumePendingSyncAction();
}

function startGithubOAuth() {
  const redirectTo = getOAuthRedirectUrl();
  const authUrl = `${syncConfig.url}/auth/v1/authorize?provider=github&redirect_to=${encodeURIComponent(redirectTo)}`;
  window.location.assign(authUrl);
}

function getOAuthRedirectUrl() {
  return `${window.location.origin}${window.location.pathname}`;
}

function hydrateAuthFromUrlHash() {
  if (!window.location.hash) {
    return;
  }

  const hash = new URLSearchParams(window.location.hash.slice(1));
  const accessToken = hash.get("access_token");
  if (!accessToken) {
    return;
  }

  saveAuthToken(accessToken);
  const cleanUrl = `${window.location.origin}${window.location.pathname}${window.location.search}`;
  window.history.replaceState({}, document.title, cleanUrl);
}

function loadSyncConfig() {
  try {
    const url = localStorage.getItem(SYNC_URL_KEY) || DEFAULT_SYNC_URL;
    const anonKey = localStorage.getItem(SYNC_ANON_KEY) || DEFAULT_SYNC_ANON_KEY;
    if (!localStorage.getItem(SYNC_URL_KEY) || !localStorage.getItem(SYNC_ANON_KEY)) {
      saveSyncConfig({ url, anonKey });
    }
    return {
      url,
      anonKey
    };
  } catch {
    return { url: DEFAULT_SYNC_URL, anonKey: DEFAULT_SYNC_ANON_KEY };
  }
}

function saveSyncConfig(config) {
  try {
    localStorage.setItem(SYNC_URL_KEY, config.url || "");
    localStorage.setItem(SYNC_ANON_KEY, config.anonKey || "");
  } catch (error) {
    console.warn("Nepodarilo se ulozit sync konfiguraci", error);
  }
}

function syncConfigReady() {
  return Boolean(syncConfig.url && syncConfig.anonKey);
}

function loadAuthState() {
  try {
    const token = localStorage.getItem(SYNC_TOKEN_KEY) || "";
    if (!token) {
      return { accessToken: "", userId: "", email: "" };
    }

    const payload = parseJwt(token);
    return {
      accessToken: token,
      userId: payload?.sub || "",
      email: payload?.email || ""
    };
  } catch {
    return { accessToken: "", userId: "", email: "" };
  }
}

function saveAuthToken(token) {
  localStorage.setItem(SYNC_TOKEN_KEY, token);
  const payload = parseJwt(token);
  authState = {
    accessToken: token,
    userId: payload?.sub || "",
    email: payload?.email || ""
  };
}

function clearAuthState() {
  localStorage.removeItem(SYNC_TOKEN_KEY);
  authState = { accessToken: "", userId: "", email: "" };
}

function parseJwt(token) {
  try {
    const parts = token.split(".");
    if (parts.length < 2) {
      return null;
    }
    const base64 = parts[1].replace(/-/g, "+").replace(/_/g, "/");
    const padded = base64.padEnd(base64.length + ((4 - (base64.length % 4)) % 4), "=");
    return JSON.parse(atob(padded));
  } catch {
    return null;
  }
}

function refreshAuthStatus() {
  const host = document.getElementById("auth-status");
  if (!syncConfigReady()) {
    host.textContent = "Cloud neni nastaven.";
    return;
  }

  if (authState.userId) {
    host.textContent = `Pripojeno: ANO (${authState.email || authState.userId})`;
  } else {
    host.textContent = "Pripojeno: NE (prihlasi se az pri Nacist/Nahrat).";
  }
}

async function runAuthAction(action) {
  if (!syncConfigReady()) {
    updateSyncStatus("Cloud neni nastaven.", true);
    return;
  }

  try {
    await action();
  } catch (error) {
    console.error(error);
    updateSyncStatus(`Auth selhal: ${String(error.message || error)}`, true);
  }
}

async function signUpSupabase(email, password) {
  if (!email || !password) {
    throw new Error("Vypln e-mail a heslo.");
  }

  const response = await fetch(`${syncConfig.url}/auth/v1/signup`, {
    method: "POST",
    headers: {
      "apikey": syncConfig.anonKey,
      "Content-Type": "application/json"
    },
    body: JSON.stringify({ email, password })
  });

  if (!response.ok) {
    const body = await safeJson(response);
    throw new Error(body?.msg || body?.error_description || `Signup selhal (${response.status})`);
  }
}

async function loginSupabase(email, password) {
  if (!email || !password) {
    throw new Error("Vypln e-mail a heslo.");
  }

  const response = await fetch(`${syncConfig.url}/auth/v1/token?grant_type=password`, {
    method: "POST",
    headers: {
      "apikey": syncConfig.anonKey,
      "Content-Type": "application/json"
    },
    body: JSON.stringify({ email, password })
  });

  const body = await safeJson(response);
  if (!response.ok || !body?.access_token) {
    throw new Error(body?.msg || body?.error_description || `Login selhal (${response.status})`);
  }

  saveAuthToken(body.access_token);
}

async function safeJson(response) {
  try {
    return await response.json();
  } catch {
    return null;
  }
}

function updateSyncStatus(text, isError = false) {
  const host = document.getElementById("sync-status");
  host.textContent = text;
  host.style.color = isError ? "#a61f2c" : "";
}

async function runSyncAction(action, actionName = "") {
  if (!syncConfigReady()) {
    updateSyncStatus("Cloud neni nastaven.", true);
    return;
  }

  if (!authState.accessToken || !authState.userId) {
    if (actionName) {
      sessionStorage.setItem(SYNC_PENDING_ACTION_KEY, actionName);
    }
    updateSyncStatus("Pro cloud sync je vyzadovano prihlaseni. Presmerovavam na GitHub...");
    startGithubOAuth();
    return;
  }

  try {
    updateSyncStatus("Probiha synchronizace...");
    await action();
    renderAll();
  } catch (error) {
    console.error(error);
    updateSyncStatus(`Sync selhal: ${String(error.message || error)}`, true);
  }
}

function resumePendingSyncAction() {
  const pending = sessionStorage.getItem(SYNC_PENDING_ACTION_KEY);
  if (!pending || !authState.accessToken || !authState.userId) {
    return;
  }

  sessionStorage.removeItem(SYNC_PENDING_ACTION_KEY);
  if (pending === "push") {
    runSyncAction(pushToCloud, "push");
    return;
  }
  if (pending === "pull") {
    runSyncAction(pullFromCloud, "pull");
  }
}

function syncHeaders() {
  return {
    "apikey": syncConfig.anonKey,
    "Authorization": `Bearer ${authState.accessToken}`,
    "Content-Type": "application/json",
    "Prefer": "return=minimal"
  };
}

async function pushToCloud() {
  const baseUrl = `${syncConfig.url}/rest/v1/tasks_sync`;
  const profile = encodeURIComponent(authState.userId);

  const deleteResponse = await fetch(`${baseUrl}?profile_id=eq.${profile}`, {
    method: "DELETE",
    headers: syncHeaders()
  });

  if (!deleteResponse.ok) {
    throw new Error(`Smazani cloud dat selhalo (${deleteResponse.status})`);
  }

  const payload = tasks.map((task) => ({
    profile_id: authState.userId,
    task_id: task.id,
    updated_at: task.updatedAt || new Date().toISOString(),
    task
  }));

  if (payload.length > 0) {
    const insertResponse = await fetch(baseUrl, {
      method: "POST",
      headers: syncHeaders(),
      body: JSON.stringify(payload)
    });

    if (!insertResponse.ok) {
      throw new Error(`Nahrani cloud dat selhalo (${insertResponse.status})`);
    }
  }

  updateSyncStatus(`Nahrano do cloudu: ${tasks.length} ukolu.`);
}

async function pullFromCloud() {
  const baseUrl = `${syncConfig.url}/rest/v1/tasks_sync`;
  const profile = encodeURIComponent(authState.userId);
  const selectResponse = await fetch(`${baseUrl}?select=task,updated_at&profile_id=eq.${profile}&order=updated_at.desc`, {
    method: "GET",
    headers: syncHeaders()
  });

  if (!selectResponse.ok) {
    throw new Error(`Nacteni cloud dat selhalo (${selectResponse.status})`);
  }

  const rows = await selectResponse.json();
  const incoming = Array.isArray(rows) ? rows.map((row) => normalizeTask(row.task || {})) : [];
  incoming.forEach((task) => {
    task.priorityScore = calculatePriority(task);
  });

  tasks = incoming;
  await persistAllTasksLocally(tasks);
  updateSyncStatus(`Nacteno z cloudu: ${tasks.length} ukolu.`);
}

async function persistAllTasksLocally(list) {
  if (storageMode === "localstorage") {
    saveTasksToLocalStorage(list);
    return;
  }

  await clearStore();
  for (const task of list) {
    await persistTask(task);
  }
}

function saveTasksToLocalStorage(list) {
  try {
    localStorage.setItem(LS_KEY, JSON.stringify(list));
  } catch (error) {
    console.warn("Nepodarilo se ulozit data do localStorage", error);
  }
}

function newId() {
  if (typeof crypto !== "undefined" && typeof crypto.randomUUID === "function") {
    return crypto.randomUUID();
  }
  return `id-${Date.now()}-${Math.random().toString(16).slice(2)}`;
}

function normalizeTask(task) {
  const nowIso = new Date().toISOString();
  return {
    id: task.id || newId(),
    title: String(task.title || "").trim(),
    description: String(task.description || "").trim(),
    dueDate: task.dueDate || "",
    estimatedHours: toNumber(task.estimatedHours, 2),
    actualHours: toNumber(task.actualHours, 0),
    importance: clamp(Math.round(toNumber(task.importance, 3)), 1, 5),
    area: AREA_ORDER.includes(task.area) ? task.area : "Jine",
    status: STATUS_ORDER.includes(task.status) ? task.status : "Todo",
    tags: Array.isArray(task.tags) ? task.tags.map((t) => String(t).trim()).filter(Boolean) : [],
    dependencyIds: Array.isArray(task.dependencyIds) ? task.dependencyIds : [],
    createdAt: task.createdAt || nowIso,
    updatedAt: task.updatedAt || task.createdAt || nowIso,
    completedAt: task.completedAt || null,
    priorityScore: 0
  };
}

function setupNavigation() {
  const links = document.querySelectorAll(".nav-link");
  const shortcutButtons = document.querySelectorAll("[data-nav-target]");

  links.forEach((button) => {
    button.addEventListener("click", () => showView(button.dataset.nav));
  });

  shortcutButtons.forEach((button) => {
    button.addEventListener("click", () => showView(button.dataset.navTarget));
  });
}

function showView(view) {
  document.querySelectorAll(".view").forEach((section) => section.classList.remove("is-active"));
  document.querySelectorAll(".nav-link").forEach((button) => button.classList.remove("is-active"));

  document.getElementById(`view-${view}`)?.classList.add("is-active");
  document.querySelector(`.nav-link[data-nav='${view}']`)?.classList.add("is-active");
}

function setupCreateForm() {
  const form = document.getElementById("task-form");

  form.innerHTML = `
    <div class="col-12"><label>Nazev</label><input name="title" required maxlength="180" /></div>
    <div class="col-12"><label>Popis</label><textarea name="description" rows="3"></textarea></div>
    <div class="col-3"><label>Termin</label><input name="dueDate" type="date" /></div>
    <div class="col-3"><label>Odhad pracnosti (h)</label><input name="estimatedHours" type="number" step="0.25" min="0" value="2" /></div>
    <div class="col-3"><label>Skutecna pracnost (h)</label><input name="actualHours" type="number" step="0.25" min="0" value="0" /></div>
    <div class="col-3"><label>Dulezitost (1-5)</label><input name="importance" type="number" min="1" max="5" value="3" /></div>
    <div class="col-4"><label>Oblast</label>
      <select name="area">
        <option value="Svp">SVP</option>
        <option value="Sdp">SDP</option>
        <option value="Bozp">BOZP</option>
        <option value="Po">PO</option>
        <option value="Jine" selected>Jine</option>
      </select>
    </div>
    <div class="col-4"><label>Stav</label>
      <select name="status">
        <option value="Todo">K vyrizeni</option>
        <option value="InProgress">Rozpracovano</option>
        <option value="Done">Hotovo</option>
        <option value="Blocked">Blokovano</option>
      </select>
    </div>
    <div class="col-4"><label>Stitky (CSV)</label><input name="tagsCsv" placeholder="napr. Finance, Report" /></div>
    <div class="col-12"><label>Zavislosti</label><select name="dependencyIds" multiple size="6" id="dependency-select"></select></div>
    <div class="col-12"><button class="btn btn-primary" type="submit">Ulozit ukol</button> <button class="btn btn-outline" type="button" id="cancel-edit">Zrusit upravy</button></div>
  `;

  form.addEventListener("submit", onCreateOrEditSubmit);
  document.getElementById("cancel-edit").addEventListener("click", resetCreateForm);
  refreshDependencyOptions();
}

function setupAutoForm() {
  const form = document.getElementById("auto-form");
  form.addEventListener("submit", async (event) => {
    event.preventDefault();
    const value = document.getElementById("auto-input").value;
    const lines = value
      .split("\n")
      .map((line) => line.trim())
      .filter(Boolean);

    if (lines.length === 0) {
      return;
    }

    const existingTitles = new Set(tasks.map((task) => task.title.toLowerCase()));
    for (const line of lines) {
      if (existingTitles.has(line.toLowerCase())) {
        continue;
      }

      const now = new Date();
      const task = normalizeTask({
        id: newId(),
        title: line,
        description: "Automaticky vygenerovany ukol z textoveho vstupu.",
        dueDate: toDateInput(addDays(now, 2)),
        estimatedHours: 1,
        actualHours: 0,
        importance: 3,
        area: inferArea(line),
        status: "Todo",
        tags: [],
        dependencyIds: [],
        createdAt: now.toISOString(),
        updatedAt: now.toISOString(),
        completedAt: null
      });
      task.priorityScore = calculatePriority(task);
      tasks.push(task);
      await persistTask(task);
      existingTitles.add(line.toLowerCase());
    }

    form.reset();
    renderAll();
    showView("tasks");
  });
}

async function onCreateOrEditSubmit(event) {
  event.preventDefault();
  const form = event.currentTarget;
  const data = new FormData(form);

  const dependencyIds = Array.from(form.querySelector("#dependency-select").selectedOptions)
    .map((opt) => opt.value)
    .filter((id) => id && id !== editTaskId);

  const now = new Date();
  const status = String(data.get("status") || "Todo");
  const prev = editTaskId ? tasks.find((task) => task.id === editTaskId) : null;

  const task = normalizeTask({
    id: editTaskId || newId(),
    title: data.get("title"),
    description: data.get("description"),
    dueDate: data.get("dueDate"),
    estimatedHours: data.get("estimatedHours"),
    actualHours: data.get("actualHours"),
    importance: data.get("importance"),
    area: data.get("area"),
    status,
    tags: parseCsv(String(data.get("tagsCsv") || "")),
    dependencyIds,
    createdAt: prev?.createdAt || now.toISOString(),
    updatedAt: now.toISOString(),
    completedAt: status === "Done" ? (prev?.completedAt || now.toISOString()) : null
  });

  task.priorityScore = calculatePriority(task);

  if (editTaskId) {
    const index = tasks.findIndex((t) => t.id === editTaskId);
    if (index >= 0) {
      tasks[index] = task;
    }
  } else {
    tasks.push(task);
  }

  await persistTask(task);
  resetCreateForm();
  renderAll();
  showView("tasks");
}

function resetCreateForm() {
  editTaskId = null;
  document.getElementById("task-form").reset();
  document.querySelector("#task-form button[type='submit']").textContent = "Ulozit ukol";
}

function refreshDependencyOptions() {
  const select = document.getElementById("dependency-select");
  if (!select) {
    return;
  }

  const currentSelected = new Set(Array.from(select.selectedOptions).map((o) => o.value));
  select.innerHTML = "";

  tasks
    .sort((a, b) => a.title.localeCompare(b.title, "cs"))
    .forEach((task) => {
      if (task.id === editTaskId) {
        return;
      }
      const option = document.createElement("option");
      option.value = task.id;
      option.textContent = task.title;
      option.selected = currentSelected.has(task.id);
      select.appendChild(option);
    });
}

function renderAll() {
  tasks.forEach((task) => {
    task.priorityScore = calculatePriority(task);
  });

  renderDashboard();
  renderTasks();
  refreshDependencyOptions();
}

function renderDashboard() {
  const kpi = getKpi();
  const kpiGrid = document.getElementById("kpi-grid");
  kpiGrid.innerHTML = "";

  const items = [
    { key: "overdue", label: "Po terminu", value: kpi.overdue, cls: "kpi-red" },
    { key: "today", label: "Na dnesek", value: kpi.today, cls: "kpi-grey" },
    { key: "week", label: "Tento tyden", value: kpi.week, cls: "kpi-blue" },
    { key: "completed", label: "Dokoncene", value: kpi.completed, cls: "kpi-blue" },
    { key: "inprogress", label: "Rozpracovane", value: kpi.inprogress, cls: "kpi-grey" },
    { key: null, label: "KPI plneni", value: `${kpi.completionRate.toFixed(1)}%`, cls: "kpi-light" }
  ];

  items.forEach((item) => {
    const button = document.createElement("button");
    button.className = `kpi ${item.cls}`;
    button.innerHTML = `<div class="kpi-label">${item.label}</div><div class="kpi-value">${item.value}</div>`;
    button.addEventListener("click", () => {
      activeFilter = item.key;
      showView("tasks");
      renderTasks();
    });
    kpiGrid.appendChild(button);
  });

  renderRecommendations();
  renderTopPriority();
  renderCharts();
}

function renderRecommendations() {
  const host = document.getElementById("recommendations");
  const open = tasks.filter((task) => task.status !== "Done");
  const overdue = open.filter((task) => isOverdue(task)).length;
  const blocked = open.filter((task) => task.status === "Blocked").length;
  const next3Days = open.filter((task) => withinDays(task, 3)).reduce((sum, task) => sum + task.estimatedHours, 0);

  const tips = [];
  if (overdue > 0) {
    tips.push(`Mas ${overdue} ukolu po terminu. Zacni dnes jejich shortlistem.`);
  }
  if (blocked > 0) {
    tips.push(`Mas ${blocked} blokovanych ukolu. Over zavislosti a dalsi krok.`);
  }
  if (next3Days > 12) {
    tips.push("Kapacita dalsich 3 dni je vysoka. Presun cast ukolu nebo sniz scope.");
  }
  if (tips.length === 0) {
    tips.push("Plan vypada stabilne. Drz fokus na top 3 ukolech podle skore.");
  }

  host.innerHTML = tips.map((tip) => `<li>${escapeHtml(tip)}</li>`).join("");
}

function renderTopPriority() {
  const body = document.getElementById("top-priority-body");
  const list = tasks
    .slice()
    .sort((a, b) => b.priorityScore - a.priorityScore)
    .slice(0, 8);

  body.innerHTML = list
    .map((task) => `
      <tr>
        <td>${escapeHtml(task.title)}</td>
        <td>${task.dueDate || "-"}</td>
        <td><span class="badge badge-blue">${task.priorityScore}</span></td>
        <td>${STATUS_LABEL[task.status]}</td>
      </tr>
    `)
    .join("") || '<tr><td colspan="4">Zatim bez ukolu.</td></tr>';
}

function renderTasks() {
  const filtered = applyFilter(tasks, activeFilter);

  const filterBadge = document.getElementById("active-filter");
  if (activeFilter) {
    filterBadge.classList.remove("hidden");
    filterBadge.innerHTML = `Filtr: ${FILTER_LABEL[activeFilter] || activeFilter} <button class="btn btn-outline btn-sm" id="clear-filter">Zrusit filtr</button>`;
    document.getElementById("clear-filter").addEventListener("click", () => {
      activeFilter = null;
      renderTasks();
    });
  } else {
    filterBadge.classList.add("hidden");
    filterBadge.textContent = "";
  }

  const grouped = new Map(AREA_ORDER.map((area) => [area, []]));
  filtered.forEach((task) => grouped.get(task.area).push(task));

  renderAreaPicker(grouped);
  renderAreaPanels(grouped);
}

function renderAreaPicker(grouped) {
  const host = document.getElementById("area-picker");
  host.innerHTML = "";

  let firstWithData = AREA_ORDER.find((area) => grouped.get(area).length > 0) || AREA_ORDER[0];
  let selectedArea = host.dataset.selectedArea || firstWithData;
  if (!AREA_ORDER.includes(selectedArea)) {
    selectedArea = firstWithData;
  }

  AREA_ORDER.forEach((area) => {
    const button = document.createElement("button");
    button.className = `area-picker area-${area.toLowerCase()}${selectedArea === area ? " is-active" : ""}`;
    button.innerHTML = `<span class="area-title">${AREA_LABEL[area]}</span><span class="area-count">${grouped.get(area).length} ukolu</span>`;
    button.addEventListener("click", () => {
      host.dataset.selectedArea = area;
      renderAreaPicker(grouped);
      renderAreaPanels(grouped);
    });
    host.appendChild(button);
  });

  host.dataset.selectedArea = selectedArea;
}

function renderAreaPanels(grouped) {
  const host = document.getElementById("area-panels");
  const selectedArea = document.getElementById("area-picker").dataset.selectedArea || AREA_ORDER[0];
  host.innerHTML = "";

  AREA_ORDER.forEach((area) => {
    const areaTasks = grouped.get(area);
    const panel = document.createElement("article");
    panel.className = `card area-card area-${area.toLowerCase()}${selectedArea === area ? "" : " hidden"}`;

    const rows = areaTasks.length === 0
      ? '<tr><td colspan="7">Zatim bez ukolu.</td></tr>'
      : areaTasks
        .slice()
        .sort((a, b) => b.priorityScore - a.priorityScore)
        .map((task) => {
          const tagText = task.tags.join(", ");
          const doneDisabled = task.status === "Done" ? "disabled" : "";
          return `
            <tr>
              <td>
                <div><strong>${escapeHtml(task.title)}</strong> <span class="area-chip area-${task.area.toLowerCase()}">${AREA_LABEL[task.area]}</span></div>
                <div class="subtitle">${escapeHtml(task.description || "")}</div>
              </td>
              <td><span class="badge badge-blue">${task.priorityScore}</span></td>
              <td>${task.dueDate || "-"}</td>
              <td>${task.actualHours} / ${task.estimatedHours} h</td>
              <td><span class="badge ${statusBadgeClass(task.status)}">${STATUS_LABEL[task.status]}</span></td>
              <td>${escapeHtml(tagText)}</td>
              <td>
                <button class="btn btn-outline btn-sm" data-action="edit" data-id="${task.id}">Upravit</button>
                <button class="btn btn-outline btn-sm" data-action="done" data-id="${task.id}" ${doneDisabled}>Hotovo</button>
                <button class="btn btn-danger btn-sm" data-action="delete" data-id="${task.id}">Smazat</button>
              </td>
            </tr>
          `;
        }).join("");

    panel.innerHTML = `
      <div class="top-row">
        <h2 class="panel-title">${AREA_LABEL[area]}</h2>
        <span class="badge badge-grey">${areaTasks.length} ukolu</span>
      </div>
      <div class="table-wrap">
        <table class="table">
          <thead><tr><th>Ukol</th><th>Priorita</th><th>Termin</th><th>Pracnost</th><th>Stav</th><th>Stitky</th><th>Akce</th></tr></thead>
          <tbody>${rows}</tbody>
        </table>
      </div>
    `;

    panel.querySelectorAll("button[data-action]").forEach((button) => {
      button.addEventListener("click", () => handleTaskAction(button.dataset.action, button.dataset.id));
    });

    host.appendChild(panel);
  });
}

async function handleTaskAction(action, id) {
  const task = tasks.find((row) => row.id === id);
  if (!task) {
    return;
  }

  if (action === "edit") {
    editTask(task);
    return;
  }

  if (action === "done" && task.status !== "Done") {
    task.status = "Done";
    task.updatedAt = new Date().toISOString();
    task.completedAt = new Date().toISOString();
    task.actualHours = Math.max(task.actualHours, task.estimatedHours);
    task.priorityScore = calculatePriority(task);
    await persistTask(task);
    renderAll();
    return;
  }

  if (action === "delete") {
    const allow = await askConfirm();
    if (!allow) {
      return;
    }
    tasks = tasks.filter((row) => row.id !== id);
    await deleteTaskById(id);
    renderAll();
  }
}

function editTask(task) {
  editTaskId = task.id;
  showView("create");

  const form = document.getElementById("task-form");
  form.title.value = task.title;
  form.description.value = task.description;
  form.dueDate.value = task.dueDate;
  form.estimatedHours.value = task.estimatedHours;
  form.actualHours.value = task.actualHours;
  form.importance.value = task.importance;
  form.area.value = task.area;
  form.status.value = task.status;
  form.tagsCsv.value = task.tags.join(", ");

  refreshDependencyOptions();
  Array.from(form.dependencyIds.options).forEach((opt) => {
    opt.selected = task.dependencyIds.includes(opt.value);
  });

  form.querySelector("button[type='submit']").textContent = "Ulozit zmeny";
}

function calculatePriority(task) {
  if (task.status === "Done") {
    return 0;
  }

  const now = new Date();
  const due = task.dueDate ? new Date(`${task.dueDate}T00:00:00`) : addDays(now, 7);
  const diffDays = Math.floor((due.getTime() - now.getTime()) / 86400000);

  const dueUrgency = diffDays <= 0 ? 100 : clamp(100 - diffDays * 10, 20, 100);
  const effortScore = clamp((task.estimatedHours / 8) * 100, 10, 100);
  const importanceScore = clamp((task.importance / 5) * 100, 20, 100);
  const overdueScore = diffDays < 0 ? clamp(Math.abs(diffDays) * 12, 0, 100) : 0;
  const depsScore = clamp(task.dependencyIds.length * 20, 0, 100);

  const weighted = dueUrgency * 0.35
    + effortScore * 0.15
    + importanceScore * 0.30
    + overdueScore * 0.10
    + depsScore * 0.10;

  return Math.round(clamp(weighted, 0, 100));
}

function getKpi() {
  const now = new Date();
  const today = stripTime(now);
  const weekEnd = addDays(today, 7);

  const open = tasks.filter((task) => task.status !== "Done");
  const done = tasks.filter((task) => task.status === "Done");

  const overdue = open.filter((task) => task.dueDate && new Date(`${task.dueDate}T00:00:00`) < today).length;
  const todayCount = open.filter((task) => task.dueDate && sameDate(new Date(`${task.dueDate}T00:00:00`), today)).length;
  const week = open.filter((task) => {
    if (!task.dueDate) {
      return false;
    }
    const due = new Date(`${task.dueDate}T00:00:00`);
    return due >= today && due <= weekEnd;
  }).length;

  return {
    overdue,
    today: todayCount,
    week,
    completed: done.length,
    inprogress: tasks.filter((task) => task.status === "InProgress").length,
    completionRate: tasks.length ? (done.length / tasks.length) * 100 : 0
  };
}

function applyFilter(list, filter) {
  const now = new Date();
  const today = stripTime(now);
  const weekEnd = addDays(today, 7);

  if (!filter) {
    return list;
  }

  if (filter === "overdue") {
    return list.filter((task) => task.status !== "Done" && task.dueDate && new Date(`${task.dueDate}T00:00:00`) < today);
  }
  if (filter === "today") {
    return list.filter((task) => task.status !== "Done" && task.dueDate && sameDate(new Date(`${task.dueDate}T00:00:00`), today));
  }
  if (filter === "week") {
    return list.filter((task) => task.status !== "Done" && task.dueDate && new Date(`${task.dueDate}T00:00:00`) >= today && new Date(`${task.dueDate}T00:00:00`) <= weekEnd);
  }
  if (filter === "completed") {
    return list.filter((task) => task.status === "Done");
  }
  if (filter === "inprogress") {
    return list.filter((task) => task.status === "InProgress");
  }

  return list;
}

function renderCharts() {
  if (typeof Chart === "undefined") {
    return;
  }

  charts.forEach((chart) => chart.destroy());
  charts = [];

  const { daily, weekly, monthly, states, heatmap } = buildSeries();

  charts.push(new Chart(document.getElementById("dailyChart"), {
    type: "line",
    data: {
      labels: Object.keys(daily),
      datasets: [{ label: "Dokoncene ukoly / den", data: Object.values(daily), borderColor: brand.blue, backgroundColor: brand.blueSoft, tension: 0.3, fill: true }]
    }
  }));

  charts.push(new Chart(document.getElementById("weeklyChart"), {
    type: "bar",
    data: {
      labels: Object.keys(weekly),
      datasets: [{ label: "Dokoncene ukoly / tyden", data: Object.values(weekly), backgroundColor: brand.red }]
    }
  }));

  charts.push(new Chart(document.getElementById("monthlyChart"), {
    type: "line",
    data: {
      labels: Object.keys(monthly),
      datasets: [{ label: "Dokoncene ukoly / mesic", data: Object.values(monthly), borderColor: brand.grey, backgroundColor: brand.greySoft, tension: 0.3 }]
    }
  }));

  charts.push(new Chart(document.getElementById("statusChart"), {
    type: "doughnut",
    data: {
      labels: states.map((s) => STATUS_LABEL[s.status]),
      datasets: [{ data: states.map((s) => s.count), backgroundColor: [brand.blue, brand.grey, brand.red, "#b8bbc4"] }]
    }
  }));

  renderHeatmap(heatmap);
}

function buildSeries() {
  const done = tasks.filter((task) => task.status === "Done" && task.completedAt);

  const daily = countByRange(done, 14, "day");
  const weekly = countByRange(done, 10, "week");
  const monthly = countByRange(done, 8, "month");

  const states = STATUS_ORDER.map((status) => ({
    status,
    count: tasks.filter((task) => task.status === status).length
  }));

  const heatmap = {};
  for (let i = 119; i >= 0; i--) {
    const day = stripTime(addDays(new Date(), -i));
    const key = toDateInput(day);
    heatmap[key] = 0;
  }

  done.forEach((task) => {
    const key = toDateInput(stripTime(new Date(task.completedAt)));
    if (Object.prototype.hasOwnProperty.call(heatmap, key)) {
      heatmap[key] += 1;
    }
  });

  return { daily, weekly, monthly, states, heatmap };
}

function renderHeatmap(heatmap) {
  const host = document.getElementById("heatmap");
  host.innerHTML = "";

  Object.keys(heatmap).sort().forEach((date) => {
    const val = heatmap[date];
    const level = Math.min(4, val);
    const div = document.createElement("div");
    div.className = `heat-cell heat-${level}`;
    div.title = `${date}: ${val}`;
    host.appendChild(div);
  });
}

function countByRange(doneTasks, units, mode) {
  const out = {};

  for (let i = units - 1; i >= 0; i--) {
    let date;
    let key;

    if (mode === "day") {
      date = stripTime(addDays(new Date(), -i));
      key = toDateInput(date);
    } else if (mode === "week") {
      date = stripTime(addDays(new Date(), -(i * 7)));
      key = `T${isoWeek(date)}-${date.getFullYear()}`;
    } else {
      date = new Date();
      date.setMonth(date.getMonth() - i, 1);
      key = `${date.getFullYear()}-${String(date.getMonth() + 1).padStart(2, "0")}`;
    }

    out[key] = 0;
  }

  doneTasks.forEach((task) => {
    const dt = new Date(task.completedAt);
    let key;
    if (mode === "day") {
      key = toDateInput(stripTime(dt));
    } else if (mode === "week") {
      key = `T${isoWeek(dt)}-${dt.getFullYear()}`;
    } else {
      key = `${dt.getFullYear()}-${String(dt.getMonth() + 1).padStart(2, "0")}`;
    }

    if (Object.prototype.hasOwnProperty.call(out, key)) {
      out[key] += 1;
    }
  });

  return out;
}

function setupServiceWorker() {
  if (!("serviceWorker" in navigator)) {
    return;
  }
  navigator.serviceWorker.register("./service-worker.js?v=5").then((registration) => {
    registration.update();
  }).catch((error) => {
    console.error("Registrace service workeru selhala", error);
  });
}

function inferArea(text) {
  const source = String(text || "").toLowerCase();
  if (source.includes("svp") || source.includes("inspekce")) {
    return "Svp";
  }
  if (source.includes("sdp") || source.includes("distribuc")) {
    return "Sdp";
  }
  if (source.includes("bozp") || source.includes("bezpecnost")) {
    return "Bozp";
  }
  if (source.includes("pozar") || source.includes("hasic") || source.includes(" po ")) {
    return "Po";
  }
  return "Jine";
}

function statusBadgeClass(status) {
  if (status === "Todo") return "status-todo";
  if (status === "InProgress") return "status-inprogress";
  if (status === "Done") return "status-done";
  return "status-blocked";
}

function isOverdue(task) {
  if (!task.dueDate || task.status === "Done") {
    return false;
  }
  return new Date(`${task.dueDate}T00:00:00`) < stripTime(new Date());
}

function withinDays(task, days) {
  if (!task.dueDate || task.status === "Done") {
    return false;
  }
  const due = new Date(`${task.dueDate}T00:00:00`);
  const now = stripTime(new Date());
  const end = addDays(now, days);
  return due >= now && due <= end;
}

function toNumber(value, fallback) {
  const num = Number(value);
  return Number.isFinite(num) ? num : fallback;
}

function clamp(value, min, max) {
  return Math.min(max, Math.max(min, value));
}

function addDays(date, days) {
  const out = new Date(date);
  out.setDate(out.getDate() + days);
  return out;
}

function stripTime(date) {
  return new Date(date.getFullYear(), date.getMonth(), date.getDate());
}

function toDateInput(date) {
  return `${date.getFullYear()}-${String(date.getMonth() + 1).padStart(2, "0")}-${String(date.getDate()).padStart(2, "0")}`;
}

function sameDate(a, b) {
  return a.getFullYear() === b.getFullYear() && a.getMonth() === b.getMonth() && a.getDate() === b.getDate();
}

function parseCsv(csv) {
  return csv
    .split(",")
    .map((part) => part.trim())
    .filter(Boolean)
    .filter((value, index, arr) => arr.findIndex((row) => row.toLowerCase() === value.toLowerCase()) === index);
}

function isoWeek(date) {
  const target = new Date(Date.UTC(date.getFullYear(), date.getMonth(), date.getDate()));
  const day = target.getUTCDay() || 7;
  target.setUTCDate(target.getUTCDate() + 4 - day);
  const yearStart = new Date(Date.UTC(target.getUTCFullYear(), 0, 1));
  return Math.ceil((((target - yearStart) / 86400000) + 1) / 7);
}

function askConfirm() {
  const dialog = document.getElementById("confirm-dialog");
  if (typeof dialog.showModal !== "function") {
    return Promise.resolve(window.confirm("Smazat ukol?"));
  }

  dialog.showModal();
  return new Promise((resolve) => {
    dialog.addEventListener("close", () => resolve(dialog.returnValue === "confirm"), { once: true });
  });
}

function escapeHtml(value) {
  return String(value)
    .replaceAll("&", "&amp;")
    .replaceAll("<", "&lt;")
    .replaceAll(">", "&gt;")
    .replaceAll('"', "&quot;")
    .replaceAll("'", "&#39;");
}

window.AkcniPlanPwa = {
  clearAll: async () => {
    await clearStore();
    tasks = [];
    renderAll();
  }
};
