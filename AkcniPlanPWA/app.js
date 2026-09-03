const DB_NAME = "akcni-plan-local";
const DB_VERSION = 1;
const STORE = "tasks";

const form = document.getElementById("task-form");
const taskList = document.getElementById("task-list");
const empty = document.getElementById("empty");
const filterStatus = document.getElementById("filter-status");
const clearAllButton = document.getElementById("clear-all");
const confirmDialog = document.getElementById("confirm-dialog");

let db;

init().catch((error) => {
  console.error(error);
  alert("Aplikaci se nepodarilo inicializovat.");
});

async function init() {
  db = await openDb();

  form.addEventListener("submit", onSubmit);
  filterStatus.addEventListener("change", renderTasks);
  clearAllButton.addEventListener("click", onClearAll);

  if ("serviceWorker" in navigator) {
    navigator.serviceWorker.register("./service-worker.js").catch((error) => {
      console.error("Registrace service workeru selhala", error);
    });
  }

  await renderTasks();
}

function openDb() {
  return new Promise((resolve, reject) => {
    const request = indexedDB.open(DB_NAME, DB_VERSION);

    request.onupgradeneeded = () => {
      const nextDb = request.result;
      const store = nextDb.createObjectStore(STORE, { keyPath: "id" });
      store.createIndex("status", "status", { unique: false });
      store.createIndex("createdAt", "createdAt", { unique: false });
    };

    request.onsuccess = () => resolve(request.result);
    request.onerror = () => reject(request.error);
  });
}

function tx(mode = "readonly") {
  return db.transaction(STORE, mode).objectStore(STORE);
}

function getAllTasks() {
  return new Promise((resolve, reject) => {
    const request = tx().getAll();
    request.onsuccess = () => resolve(request.result || []);
    request.onerror = () => reject(request.error);
  });
}

function saveTask(task) {
  return new Promise((resolve, reject) => {
    const request = tx("readwrite").put(task);
    request.onsuccess = () => resolve();
    request.onerror = () => reject(request.error);
  });
}

function deleteTask(id) {
  return new Promise((resolve, reject) => {
    const request = tx("readwrite").delete(id);
    request.onsuccess = () => resolve();
    request.onerror = () => reject(request.error);
  });
}

function clearTasks() {
  return new Promise((resolve, reject) => {
    const request = tx("readwrite").clear();
    request.onsuccess = () => resolve();
    request.onerror = () => reject(request.error);
  });
}

async function onSubmit(event) {
  event.preventDefault();

  const data = new FormData(form);
  const title = String(data.get("title") || "").trim();
  if (!title) {
    return;
  }

  const task = {
    id: crypto.randomUUID(),
    title,
    description: String(data.get("description") || "").trim(),
    status: String(data.get("status") || "Todo"),
    dueDate: String(data.get("dueDate") || ""),
    createdAt: new Date().toISOString()
  };

  await saveTask(task);
  form.reset();
  await renderTasks();
}

async function onClearAll() {
  const result = await askConfirm();
  if (!result) {
    return;
  }

  await clearTasks();
  await renderTasks();
}

function askConfirm() {
  if (typeof confirmDialog.showModal !== "function") {
    return Promise.resolve(window.confirm("Opravdu smazat vsechny ukoly?"));
  }

  confirmDialog.showModal();
  return new Promise((resolve) => {
    confirmDialog.addEventListener(
      "close",
      () => resolve(confirmDialog.returnValue === "confirm"),
      { once: true }
    );
  });
}

async function renderTasks() {
  const tasks = await getAllTasks();
  const filter = filterStatus.value;
  const filtered = tasks
    .filter((task) => filter === "All" || task.status === filter)
    .sort((a, b) => {
      const byDue = (a.dueDate || "9999-12-31").localeCompare(b.dueDate || "9999-12-31");
      if (byDue !== 0) {
        return byDue;
      }
      return b.createdAt.localeCompare(a.createdAt);
    });

  taskList.innerHTML = "";

  filtered.forEach((task) => {
    const item = document.createElement("li");
    item.className = "task";

    const dueLabel = task.dueDate ? `Termin: ${formatDate(task.dueDate)}` : "Termin: bez data";
    const desc = task.description ? escapeHtml(task.description) : "Bez popisu";

    item.innerHTML = `
      <div class="task-head">
        <div>
          <h3 class="task-title">${escapeHtml(task.title)}</h3>
          <p class="task-meta">${dueLabel}</p>
        </div>
        <span class="badge ${task.status}">${statusLabel(task.status)}</span>
      </div>
      <p class="task-meta">${desc}</p>
      <div class="task-actions">
        <button type="button" class="ghost" data-action="toggle" data-id="${task.id}">Prepnout stav</button>
        <button type="button" class="danger" data-action="delete" data-id="${task.id}">Smazat</button>
      </div>
    `;

    taskList.appendChild(item);
  });

  taskList.querySelectorAll("button").forEach((button) => {
    button.addEventListener("click", async () => {
      const id = button.getAttribute("data-id");
      const action = button.getAttribute("data-action");
      if (!id || !action) {
        return;
      }

      if (action === "delete") {
        await deleteTask(id);
      }

      if (action === "toggle") {
        const current = tasks.find((task) => task.id === id);
        if (current) {
          current.status = nextStatus(current.status);
          await saveTask(current);
        }
      }

      await renderTasks();
    });
  });

  empty.style.display = filtered.length === 0 ? "block" : "none";
}

function nextStatus(status) {
  if (status === "Todo") {
    return "InProgress";
  }
  if (status === "InProgress") {
    return "Done";
  }
  return "Todo";
}

function statusLabel(status) {
  if (status === "Todo") {
    return "K vyreseni";
  }
  if (status === "InProgress") {
    return "Rozpracovano";
  }
  return "Hotovo";
}

function formatDate(value) {
  const date = new Date(`${value}T00:00:00`);
  if (Number.isNaN(date.getTime())) {
    return value;
  }
  return new Intl.DateTimeFormat("cs-CZ").format(date);
}

function escapeHtml(value) {
  return value
    .replaceAll("&", "&amp;")
    .replaceAll("<", "&lt;")
    .replaceAll(">", "&gt;")
    .replaceAll('"', "&quot;")
    .replaceAll("'", "&#39;");
}
