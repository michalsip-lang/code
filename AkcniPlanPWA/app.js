const DB_NAME = "akcni-plan-local";
const DB_VERSION = 2;
const STORE = "tasks";

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

init().catch((error) => {
  console.error(error);
  alert("Aplikaci se nepodarilo inicializovat.");
});

async function init() {
  db = await openDb();
  tasks = await loadTasks();

  setupNavigation();
  setupCreateForm();
  setupAutoForm();
  setupServiceWorker();

  renderAll();
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
  return new Promise((resolve, reject) => {
    const request = tx("readwrite").put(task);
    request.onsuccess = () => resolve();
    request.onerror = () => reject(request.error);
  });
}

function deleteTaskById(id) {
  return new Promise((resolve, reject) => {
    const request = tx("readwrite").delete(id);
    request.onsuccess = () => resolve();
    request.onerror = () => reject(request.error);
  });
}

function clearStore() {
  return new Promise((resolve, reject) => {
    const request = tx("readwrite").clear();
    request.onsuccess = () => resolve();
    request.onerror = () => reject(request.error);
  });
}

function normalizeTask(task) {
  return {
    id: task.id || crypto.randomUUID(),
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
    createdAt: task.createdAt || new Date().toISOString(),
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
        id: crypto.randomUUID(),
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
    id: editTaskId || crypto.randomUUID(),
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
  navigator.serviceWorker.register("./service-worker.js").catch((error) => {
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
