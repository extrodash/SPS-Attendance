import { initializeApp } from "https://www.gstatic.com/firebasejs/10.8.0/firebase-app.js";
import {
  collection,
  doc,
  enableIndexedDbPersistence,
  getDocs,
  getFirestore,
  onSnapshot,
  query,
  setDoc,
  where,
} from "https://www.gstatic.com/firebasejs/10.8.0/firebase-firestore.js";
import {
  getAuth,
  onAuthStateChanged,
  signInAnonymously,
} from "https://www.gstatic.com/firebasejs/10.8.0/firebase-auth.js";

const firebaseConfig = {
  apiKey: "AIzaSyArUCqRKjBkmORPQcAWH1JMaUXj6YPzhK8",
  authDomain: "sps-attendance.firebaseapp.com",
  projectId: "sps-attendance",
  storageBucket: "sps-attendance.firebasestorage.app",
  messagingSenderId: "239458002074",
  appId: "1:239458002074:web:4e3b10e87702e42e34ab23",
};

const elements = {
  datePicker: document.getElementById("date-picker"),
  saveBtn: document.getElementById("save-btn"),
  saveStatus: document.getElementById("save-status"),
  peopleList: document.getElementById("people-list"),
  editPeopleBtn: document.getElementById("edit-people-btn"),
  peopleDialog: document.getElementById("people-dialog"),
  peopleTextarea: document.getElementById("people-textarea"),
  savePeopleBtn: document.getElementById("save-people"),
  configWarning: document.getElementById("config-warning"),
  prevMonth: document.getElementById("prev-month"),
  nextMonth: document.getElementById("next-month"),
  monthLabel: document.getElementById("month-label"),
  calendarGrid: document.getElementById("calendar-grid"),
  trendsList: document.getElementById("trends-list"),
  exportBtn: document.getElementById("export-btn"),
  importFile: document.getElementById("import-file"),
};

const LOCAL_KEY = "attendance-app-data-v1";
const DEFAULT_PEOPLE = [];
const LATE_REGEX = /\blate\b/i;

const state = {
  currentDate: todayISO(),
  people: [],
  attendance: {},
  monthCursor: new Date(),
  monthAttendance: {},
  trendsAttendance: [],
  dirty: false,
  firebaseReady: false,
};

let db = null;
let auth = null;
let unsubscribeAttendance = null;
let unsubscribePeople = null;
let unsubscribeMonth = null;

init();

async function init() {
  elements.datePicker.value = state.currentDate;
  elements.datePicker.addEventListener("change", onDateChange);
  elements.saveBtn.addEventListener("click", onSave);
  elements.editPeopleBtn.addEventListener("click", openPeopleDialog);
  elements.savePeopleBtn.addEventListener("click", onSavePeople);
  elements.prevMonth.addEventListener("click", () => shiftMonth(-1));
  elements.nextMonth.addEventListener("click", () => shiftMonth(1));
  elements.exportBtn.addEventListener("click", exportBackup);
  elements.importFile.addEventListener("change", importBackup);

  if ("serviceWorker" in navigator) {
    navigator.serviceWorker.register("sw.js");
  }

  const hasConfig = isFirebaseConfigured();
  if (!hasConfig) {
    elements.configWarning.hidden = false;
    hydrateFromLocal();
    renderAll();
    return;
  }

  const app = initializeApp(firebaseConfig);
  db = getFirestore(app);
  auth = getAuth(app);

  try {
    await enableIndexedDbPersistence(db);
  } catch (error) {
    if (error.code !== "failed-precondition") {
      console.warn("Offline persistence error", error);
    }
  }

  signInAnonymously(auth).catch((error) => {
    console.warn("Auth error", error);
  });

  onAuthStateChanged(auth, (user) => {
    if (!user) {
      return;
    }
    state.firebaseReady = true;
    hydrateFromLocal();
    listenPeople();
    listenAttendance(state.currentDate);
    loadMonth(state.monthCursor);
    loadTrends();
  });
}

function isFirebaseConfigured() {
  return !Object.values(firebaseConfig).some((value) => {
    if (!value) {
      return true;
    }
    return String(value).includes("YOUR_");
  });
}

function todayISO() {
  const now = new Date();
  return now.toISOString().slice(0, 10);
}

function onDateChange(event) {
  state.currentDate = event.target.value;
  state.dirty = false;
  setSaveStatus("");
  if (state.firebaseReady) {
    listenAttendance(state.currentDate);
  } else {
    hydrateAttendanceFromLocal();
    renderPeopleList();
  }
}

function openPeopleDialog() {
  elements.peopleTextarea.value = state.people.map((person) => person.name).join("\n");
  elements.peopleDialog.showModal();
}

async function onSavePeople(event) {
  event.preventDefault();
  const lines = elements.peopleTextarea.value
    .split("\n")
    .map((line) => line.trim())
    .filter(Boolean);

  state.people = buildPeopleList(lines);
  await savePeople();
  renderPeopleList();
  markDirty();
  elements.peopleDialog.close();
}

function buildPeopleList(names) {
  if (names.length === 0) {
    return [];
  }

  const existing = new Map(
    state.people.map((person) => [person.name.toLowerCase(), person])
  );
  const usedIds = new Set();
  const list = [];

  names.forEach((name) => {
    const key = name.toLowerCase();
    const match = existing.get(key);
    if (match) {
      list.push({ id: match.id, name: match.name });
      usedIds.add(match.id);
      return;
    }

    let id = slugify(name);
    let suffix = 2;
    while (usedIds.has(id)) {
      id = `${slugify(name)}-${suffix}`;
      suffix += 1;
    }
    usedIds.add(id);
    list.push({ id, name });
  });

  return list;
}

function slugify(text) {
  return text
    .toLowerCase()
    .replace(/[^a-z0-9]+/g, "-")
    .replace(/(^-|-$)/g, "");
}

function renderAll() {
  renderPeopleList();
  renderCalendar();
  renderTrends();
}

function renderPeopleList() {
  elements.peopleList.innerHTML = "";

  if (!state.people.length) {
    const empty = document.createElement("div");
    empty.className = "hint";
    empty.textContent = "No people yet. Click 'Edit list' to add names.";
    elements.peopleList.appendChild(empty);
    return;
  }

  state.people.forEach((person) => {
    const row = document.createElement("div");
    row.className = "person-row";
    row.dataset.personId = person.id;

    const name = document.createElement("div");
    name.className = "person-name";
    name.textContent = person.name;

    const statusGroup = document.createElement("div");
    statusGroup.className = "status-group";

    const hereBtn = createStatusButton("✓", "here");
    const notBtn = createStatusButton("✗", "not");
    const clearBtn = document.createElement("button");
    clearBtn.className = "clear-btn";
    clearBtn.type = "button";
    clearBtn.textContent = "–";

    statusGroup.append(hereBtn, notBtn, clearBtn);

    const note = document.createElement("input");
    note.className = "note-input";
    note.placeholder = "Note";
    note.value = getEntry(person.id).note || "";

    const subGroup = document.createElement("div");
    subGroup.className = "sub-group";

    const amBtn = createSubButton("AM", "am");
    const pmBtn = createSubButton("PM", "pm");

    subGroup.append(amBtn, pmBtn);

    hereBtn.addEventListener("click", () => toggleStatus(person.id, "here"));
    notBtn.addEventListener("click", () => toggleStatus(person.id, "not"));
    clearBtn.addEventListener("click", () => clearStatus(person.id));
    amBtn.addEventListener("click", () => cycleMeeting(person.id, "am"));
    pmBtn.addEventListener("click", () => cycleMeeting(person.id, "pm"));
    note.addEventListener("input", (event) => {
      setNote(person.id, event.target.value);
    });

    row.append(name, statusGroup, note, subGroup);
    elements.peopleList.appendChild(row);

    applyRowState(person.id, row);
  });
}

function createStatusButton(label, status) {
  const button = document.createElement("button");
  button.className = `status-btn ${status}`;
  button.type = "button";
  button.textContent = label;
  return button;
}

function createSubButton(label, field) {
  const button = document.createElement("button");
  button.className = "sub-btn";
  button.type = "button";
  button.textContent = label;
  button.dataset.field = field;
  return button;
}

function applyRowState(personId, row) {
  const entry = getEntry(personId);
  const hereBtn = row.querySelector(".status-btn.here");
  const notBtn = row.querySelector(".status-btn.not");
  hereBtn.classList.toggle("active", entry.status === "here");
  notBtn.classList.toggle("active", entry.status === "not");

  const amBtn = row.querySelector(".sub-btn[data-field='am']");
  const pmBtn = row.querySelector(".sub-btn[data-field='pm']");

  setMeetingButtonState(amBtn, entry.am);
  setMeetingButtonState(pmBtn, entry.pm);
}

function setMeetingButtonState(button, value) {
  button.classList.toggle("active", Boolean(value));
  button.classList.toggle("here", value === "here");
  button.classList.toggle("not", value === "not");
}

function getEntry(personId) {
  if (!state.attendance[personId]) {
    state.attendance[personId] = { status: null, note: "", am: null, pm: null };
  }
  return state.attendance[personId];
}

function toggleStatus(personId, status) {
  const entry = getEntry(personId);
  entry.status = entry.status === status ? null : status;
  markDirty();
  refreshRow(personId);
}

function clearStatus(personId) {
  const entry = getEntry(personId);
  entry.status = null;
  entry.note = "";
  entry.am = null;
  entry.pm = null;
  markDirty();
  refreshRow(personId);
}

function cycleMeeting(personId, field) {
  const entry = getEntry(personId);
  const current = entry[field];
  if (!current) {
    entry[field] = "here";
  } else if (current === "here") {
    entry[field] = "not";
  } else {
    entry[field] = null;
  }
  markDirty();
  refreshRow(personId);
}

function setNote(personId, value) {
  const entry = getEntry(personId);
  entry.note = value;
  markDirty();
}

function refreshRow(personId) {
  const row = elements.peopleList.querySelector(
    `.person-row[data-person-id='${personId}']`
  );
  if (!row) {
    return;
  }
  applyRowState(personId, row);
}

function markDirty() {
  if (!state.dirty) {
    state.dirty = true;
    setSaveStatus("Unsaved changes");
  }
}

function setSaveStatus(message) {
  elements.saveStatus.textContent = message;
}

async function onSave() {
  const payload = buildAttendancePayload();

  if (state.firebaseReady) {
    await saveAttendance(state.currentDate, payload);
  }

  saveLocalBackup(payload, state.currentDate);
  state.dirty = false;
  setSaveStatus(`Saved ${new Date().toLocaleTimeString()}`);

  await loadMonth(state.monthCursor);
  await loadTrends();
}

function buildAttendancePayload() {
  const peopleData = {};
  state.people.forEach((person) => {
    const entry = getEntry(person.id);
    peopleData[person.id] = {
      name: person.name,
      status: entry.status || null,
      note: entry.note || "",
      am: entry.am || null,
      pm: entry.pm || null,
    };
  });

  return {
    date: state.currentDate,
    updatedAt: new Date().toISOString(),
    peopleOrder: state.people.map((person) => person.id),
    people: peopleData,
  };
}

async function saveAttendance(date, payload) {
  const ref = doc(db, "attendance", date);
  await setDoc(ref, payload, { merge: true });
}

async function savePeople() {
  if (state.firebaseReady) {
    await setDoc(
      doc(db, "config", "people"),
      {
        updatedAt: new Date().toISOString(),
        people: state.people,
      },
      { merge: true }
    );
  }

  const local = loadLocalStore();
  local.people = state.people;
  saveLocalStore(local);
}

function listenPeople() {
  if (unsubscribePeople) {
    unsubscribePeople();
  }

  const ref = doc(db, "config", "people");
  unsubscribePeople = onSnapshot(ref, (snapshot) => {
    const data = snapshot.data();
    if (!data || !data.people) {
      if (!state.people.length && DEFAULT_PEOPLE.length) {
        state.people = buildPeopleList(DEFAULT_PEOPLE);
      }
      renderPeopleList();
      return;
    }

    state.people = data.people;
    renderPeopleList();
  });
}

function listenAttendance(date) {
  if (unsubscribeAttendance) {
    unsubscribeAttendance();
  }

  const ref = doc(db, "attendance", date);
  unsubscribeAttendance = onSnapshot(ref, (snapshot) => {
    if (state.dirty) {
      return;
    }
    const data = snapshot.data();
    if (!data) {
      state.attendance = {};
      renderPeopleList();
      setSaveStatus("");
      return;
    }

    state.attendance = mapAttendanceEntries(data.people || {});
    renderPeopleList();
    setSaveStatus("Loaded");
  });
}

function mapAttendanceEntries(peopleData) {
  const mapped = {};
  Object.entries(peopleData).forEach(([id, entry]) => {
    mapped[id] = {
      status: entry.status || null,
      note: entry.note || "",
      am: entry.am || null,
      pm: entry.pm || null,
    };
  });
  return mapped;
}

function hydrateFromLocal() {
  const local = loadLocalStore();
  state.people = local.people || DEFAULT_PEOPLE.map((name) => ({
    id: slugify(name),
    name,
  }));
  state.attendance = local.attendance?.[state.currentDate] || {};
  state.monthAttendance = local.attendance || {};
  renderAll();
  loadTrends();
}

function hydrateAttendanceFromLocal() {
  const local = loadLocalStore();
  state.attendance = local.attendance?.[state.currentDate] || {};
}

function loadLocalStore() {
  try {
    return JSON.parse(localStorage.getItem(LOCAL_KEY)) || {
      people: [],
      attendance: {},
    };
  } catch (error) {
    console.warn("Local store error", error);
    return { people: [], attendance: {} };
  }
}

function saveLocalStore(data) {
  localStorage.setItem(LOCAL_KEY, JSON.stringify(data));
}

function saveLocalBackup(payload, date) {
  const local = loadLocalStore();
  local.people = state.people;
  local.attendance = local.attendance || {};
  local.attendance[date] = payload;
  saveLocalStore(local);
}

function shiftMonth(delta) {
  state.monthCursor = new Date(
    state.monthCursor.getFullYear(),
    state.monthCursor.getMonth() + delta,
    1
  );
  loadMonth(state.monthCursor);
}

async function loadMonth(date) {
  const { start, end } = monthRange(date);
  elements.monthLabel.textContent = date.toLocaleString("en-US", {
    month: "long",
    year: "numeric",
  });

  if (!state.firebaseReady) {
    const local = loadLocalStore();
    const attendance = local.attendance || {};
    const filtered = {};
    Object.keys(attendance).forEach((key) => {
      if (key >= start && key <= end) {
        filtered[key] = attendance[key];
      }
    });
    state.monthAttendance = filtered;
    renderCalendar();
    return;
  }

  if (unsubscribeMonth) {
    unsubscribeMonth();
  }

  const q = query(
    collection(db, "attendance"),
    where("date", ">=", start),
    where("date", "<=", end)
  );

  unsubscribeMonth = onSnapshot(q, (snapshot) => {
    const map = {};
    snapshot.forEach((docSnap) => {
      map[docSnap.id] = docSnap.data();
    });
    state.monthAttendance = map;
    renderCalendar();
  });
}

function monthRange(date) {
  const year = date.getFullYear();
  const month = date.getMonth();
  const start = new Date(year, month, 1);
  const end = new Date(year, month + 1, 0);
  return {
    start: start.toISOString().slice(0, 10),
    end: end.toISOString().slice(0, 10),
  };
}

function renderCalendar() {
  elements.calendarGrid.innerHTML = "";
  elements.monthLabel.textContent = state.monthCursor.toLocaleString("en-US", {
    month: "long",
    year: "numeric",
  });

  const cursor = new Date(state.monthCursor.getFullYear(), state.monthCursor.getMonth(), 1);
  const month = cursor.getMonth();
  const startDay = cursor.getDay();

  for (let i = 0; i < startDay; i += 1) {
    elements.calendarGrid.appendChild(createCalendarCell());
  }

  while (cursor.getMonth() === month) {
    const dateKey = cursor.toISOString().slice(0, 10);
    const data = state.monthAttendance[dateKey];
    const cell = createCalendarCell({ date: cursor.getDate(), dateKey, data });
    elements.calendarGrid.appendChild(cell);
    cursor.setDate(cursor.getDate() + 1);
  }
}

function createCalendarCell(payload) {
  const cell = document.createElement("div");
  cell.className = "calendar-cell";

  if (!payload) {
    return cell;
  }

  const { date, dateKey, data } = payload;
  const day = document.createElement("div");
  day.className = "day-num";
  day.textContent = date;

  const meta = document.createElement("div");
  meta.className = "day-meta";

  if (data?.people) {
    const { hereCount, total } = countAttendance(data.people);
    meta.textContent = `${hereCount}/${total} here`;
    if (total > 0) {
      const ratio = hereCount / total;
      if (ratio >= 0.75) {
        cell.classList.add("busy");
      } else if (ratio <= 0.4) {
        cell.classList.add("slow");
      }
      cell.title = `Here: ${hereCount} / ${total}`;
    }
  }

  cell.addEventListener("click", () => {
    state.currentDate = dateKey;
    elements.datePicker.value = dateKey;
    if (state.firebaseReady) {
      listenAttendance(dateKey);
    } else {
      hydrateAttendanceFromLocal();
      renderPeopleList();
    }
  });

  cell.append(day, meta);
  return cell;
}

function countAttendance(peopleData) {
  const entries = Object.values(peopleData || {});
  let hereCount = 0;
  let total = 0;
  entries.forEach((entry) => {
    if (!entry) {
      return;
    }
    if (entry.status) {
      total += 1;
      if (entry.status === "here") {
        hereCount += 1;
      }
    }
  });
  return { hereCount, total };
}

async function loadTrends() {
  const end = todayISO();
  const start = addDays(end, -90);

  if (!state.firebaseReady) {
    const local = loadLocalStore();
    const attendance = local.attendance || {};
    state.trendsAttendance = Object.values(attendance).filter(
      (entry) => entry?.date >= start && entry?.date <= end
    );
    renderTrends();
    return;
  }

  const q = query(
    collection(db, "attendance"),
    where("date", ">=", start),
    where("date", "<=", end)
  );

  const snapshot = await getDocs(q);
  state.trendsAttendance = snapshot.docs.map((docSnap) => docSnap.data());
  renderTrends();
}

function addDays(dateString, delta) {
  const date = new Date(dateString);
  date.setDate(date.getDate() + delta);
  return date.toISOString().slice(0, 10);
}

function renderTrends() {
  elements.trendsList.innerHTML = "";

  if (!state.trendsAttendance.length) {
    const empty = document.createElement("div");
    empty.className = "hint";
    empty.textContent = "No attendance data yet.";
    elements.trendsList.appendChild(empty);
    return;
  }

  const stats = new Map();

  state.trendsAttendance.forEach((entry) => {
    if (!entry?.people) {
      return;
    }
    Object.entries(entry.people).forEach(([id, personEntry]) => {
      const name = personEntry.name || id;
      const record = stats.get(id) || {
        id,
        name,
        here: 0,
        not: 0,
        total: 0,
        late: 0,
      };

      if (personEntry.status) {
        record.total += 1;
        if (personEntry.status === "here") {
          record.here += 1;
        } else if (personEntry.status === "not") {
          record.not += 1;
        }
      }

      if (personEntry.note && LATE_REGEX.test(personEntry.note)) {
        record.late += 1;
      }

      stats.set(id, record);
    });
  });

  const sorted = Array.from(stats.values()).sort((a, b) => {
    if (b.total !== a.total) {
      return b.total - a.total;
    }
    return a.name.localeCompare(b.name);
  });

  sorted.forEach((record) => {
    const row = document.createElement("div");
    row.className = "trend-row";

    const name = document.createElement("div");
    name.className = "name";
    name.textContent = record.name;

    const detail = document.createElement("div");
    const percent = record.total
      ? Math.round((record.here / record.total) * 100)
      : 0;
    detail.textContent = `Here ${record.here}/${record.total} (${percent}%), Not ${record.not}, Late ${record.late}`;

    row.append(name, detail);
    elements.trendsList.appendChild(row);
  });
}

function exportBackup() {
  const local = loadLocalStore();
  const data = {
    exportedAt: new Date().toISOString(),
    people: local.people || state.people,
    attendance: local.attendance || {},
  };

  const blob = new Blob([JSON.stringify(data, null, 2)], {
    type: "application/json",
  });
  const url = URL.createObjectURL(blob);
  const link = document.createElement("a");
  link.href = url;
  link.download = `attendance-backup-${todayISO()}.json`;
  document.body.appendChild(link);
  link.click();
  link.remove();
  URL.revokeObjectURL(url);
}

async function importBackup(event) {
  const file = event.target.files?.[0];
  if (!file) {
    return;
  }

  try {
    const text = await file.text();
    const data = JSON.parse(text);
    if (!data || !data.people || !data.attendance) {
      throw new Error("Invalid backup format");
    }

    saveLocalStore({
      people: data.people,
      attendance: data.attendance,
    });

    state.people = data.people;
    state.attendance = data.attendance[state.currentDate] || {};
    state.monthAttendance = data.attendance || {};
    renderAll();
  } catch (error) {
    alert("Could not import that file.");
    console.warn(error);
  } finally {
    event.target.value = "";
  }
}
