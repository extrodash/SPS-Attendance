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
  teamFilter: document.getElementById("team-filter"),
  attendanceList: document.getElementById("attendance-list"),
  attendanceHint: document.getElementById("attendance-hint"),
  directoryList: document.getElementById("directory-list"),
  directoryPanel: document.getElementById("directory-panel"),
  toggleDirectory: document.getElementById("toggle-directory"),
  peopleInput: document.getElementById("person-input"),
  teamInput: document.getElementById("team-input"),
  addPersonBtn: document.getElementById("add-person-btn"),
  peopleMessage: document.getElementById("people-message"),
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
const TEAM_FILTER_ALL = "__all__";
const TEAM_FILTER_UNASSIGNED = "__unassigned__";
const AUTO_SAVE_DELAY_MS = 20000;
const WEEKDAYS = [
  { key: "mon", label: "M", full: "Monday", index: 1 },
  { key: "tue", label: "T", full: "Tuesday", index: 2 },
  { key: "wed", label: "W", full: "Wednesday", index: 3 },
  { key: "thu", label: "T", full: "Thursday", index: 4 },
  { key: "fri", label: "F", full: "Friday", index: 5 },
];
const DEFAULT_AVAILABILITY = WEEKDAYS.map((day) => day.key);

const state = {
  currentDate: todayISO(),
  people: [],
  attendance: {},
  monthCursor: new Date(),
  monthAttendance: {},
  trendsAttendance: [],
  dirty: false,
  firebaseReady: false,
  selectedTeam: TEAM_FILTER_ALL,
  directoryOpen: false,
};

let db = null;
let auth = null;
let unsubscribeAttendance = null;
let unsubscribePeople = null;
let unsubscribeMonth = null;
let peopleMessageTimer = null;
let autoSaveTimer = null;
let saveInProgress = false;
const removeConfirmTimers = new WeakMap();
const REMOVE_CONFIRM_MS = 3500;

init();

async function init() {
  elements.datePicker.value = state.currentDate;
  elements.datePicker.addEventListener("change", onDateChange);
  elements.saveBtn.addEventListener("click", onSave);
  elements.addPersonBtn.addEventListener("click", addPersonFromInput);
  elements.teamFilter.addEventListener("change", onTeamFilterChange);
  if (elements.toggleDirectory) {
    elements.toggleDirectory.addEventListener("click", () => {
      setDirectoryOpen(!state.directoryOpen);
    });
  }
  elements.peopleInput.addEventListener("keydown", (event) => {
    if (event.key === "Enter") {
      addPersonFromInput(event);
    }
  });
  elements.prevMonth.addEventListener("click", () => shiftMonth(-1));
  elements.nextMonth.addEventListener("click", () => shiftMonth(1));
  elements.exportBtn.addEventListener("click", exportBackup);
  elements.importFile.addEventListener("change", importBackup);
  setDirectoryOpen(state.directoryOpen);

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
  clearAutoSaveTimer();
  if (state.firebaseReady) {
    listenAttendance(state.currentDate);
  } else {
    hydrateAttendanceFromLocal();
    renderAttendanceList();
  }
}

function normalizeName(value) {
  return value.replace(/\s+/g, " ").trim();
}

function normalizeTeam(value) {
  if (!value) {
    return "";
  }
  return value.replace(/\s+/g, " ").trim();
}

function setPeopleMessage(message, tone = "muted") {
  if (!elements.peopleMessage) {
    return;
  }
  elements.peopleMessage.textContent = message;
  elements.peopleMessage.dataset.tone = tone;

  if (peopleMessageTimer) {
    clearTimeout(peopleMessageTimer);
    peopleMessageTimer = null;
  }

  if (message) {
    peopleMessageTimer = setTimeout(() => {
      elements.peopleMessage.textContent = "";
      elements.peopleMessage.dataset.tone = "muted";
    }, 3500);
  }
}

function setDirectoryOpen(open) {
  state.directoryOpen = Boolean(open);
  if (elements.directoryPanel) {
    elements.directoryPanel.hidden = !state.directoryOpen;
  }
  if (elements.toggleDirectory) {
    elements.toggleDirectory.textContent = state.directoryOpen
      ? "Close directory"
      : "Edit profiles";
  }
}

async function addPersonFromInput(event) {
  if (event) {
    event.preventDefault();
  }

  const raw = elements.peopleInput.value || "";
  const cleaned = normalizeName(raw);
  const teamInput = elements.teamInput?.value || "";
  const fallbackTeam =
    state.selectedTeam !== TEAM_FILTER_ALL &&
    state.selectedTeam !== TEAM_FILTER_UNASSIGNED
      ? state.selectedTeam
      : "";
  const cleanedTeam = normalizeTeam(teamInput || fallbackTeam);

  if (!cleaned) {
    setPeopleMessage("Enter a name to add.", "warn");
    return;
  }

  const exists = state.people.some(
    (person) => person.name.toLowerCase() === cleaned.toLowerCase()
  );
  if (exists) {
    setPeopleMessage("That name already exists.", "warn");
    elements.peopleInput.select();
    return;
  }

  state.people = [...state.people, createPerson(cleaned, cleanedTeam)];
  elements.peopleInput.value = "";
  if (elements.teamInput && elements.teamInput.value !== cleanedTeam) {
    elements.teamInput.value = cleanedTeam;
  }
  await savePeople();
  renderDirectoryList();
  renderAttendanceList();
  setPeopleMessage(
    cleanedTeam ? `Added ${cleaned} to ${cleanedTeam}.` : `Added ${cleaned}.`,
    "success"
  );
}

function slugify(text) {
  return text
    .toLowerCase()
    .replace(/[^a-z0-9]+/g, "-")
    .replace(/(^-|-$)/g, "");
}

function sanitizePeopleList(people) {
  if (!Array.isArray(people)) {
    return [];
  }

  const usedIds = new Set();
  const list = [];

  people.forEach((raw) => {
    if (!raw) {
      return;
    }

    const name = normalizeName(typeof raw === "string" ? raw : raw.name || "");
    if (!name) {
      return;
    }

    let id =
      typeof raw === "string"
        ? slugify(name)
        : raw.id || slugify(name);

    if (!id) {
      id = `person-${list.length + 1}`;
    }

    let candidate = id;
    let suffix = 2;
    while (usedIds.has(candidate)) {
      candidate = `${id}-${suffix}`;
      suffix += 1;
    }
    usedIds.add(candidate);

    const team =
      typeof raw === "string" ? "" : normalizeTeam(raw.team || "");
    const availability =
      typeof raw === "string"
        ? [...DEFAULT_AVAILABILITY]
        : normalizeAvailability(raw.availability);

    list.push({
      id: candidate,
      name,
      team,
      availability,
    });
  });

  return list;
}

function normalizeAvailability(value) {
  if (value === undefined || value === null) {
    return [...DEFAULT_AVAILABILITY];
  }

  if (!Array.isArray(value)) {
    return [...DEFAULT_AVAILABILITY];
  }

  const valid = new Set(WEEKDAYS.map((day) => day.key));
  const cleaned = [];
  value.forEach((entry) => {
    if (!entry) {
      return;
    }
    const key = String(entry).toLowerCase();
    if (!valid.has(key)) {
      return;
    }
    if (!cleaned.includes(key)) {
      cleaned.push(key);
    }
  });

  return cleaned;
}

function createPerson(name, team) {
  const usedIds = new Set(state.people.map((person) => person.id));
  const baseId = slugify(name);
  let candidate = baseId || `person-${Date.now()}`;
  let suffix = 2;

  while (usedIds.has(candidate)) {
    candidate = baseId ? `${baseId}-${suffix}` : `person-${suffix}`;
    suffix += 1;
  }

  return {
    id: candidate,
    name,
    team: normalizeTeam(team),
    availability: [...DEFAULT_AVAILABILITY],
  };
}

function getWeekdayKey(dateISO) {
  const date = new Date(`${dateISO}T00:00:00`);
  const day = date.getDay();
  const match = WEEKDAYS.find((entry) => entry.index === day);
  return match ? match.key : null;
}

function getWeekdayLabel(dateISO) {
  const date = new Date(`${dateISO}T00:00:00`);
  return date.toLocaleDateString("en-US", { weekday: "long" });
}

function renderAll() {
  renderDirectoryList();
  renderAttendanceList();
  renderCalendar();
  renderTrends();
}

function getTeamOptions() {
  const teams = new Map();
  state.people.forEach((person) => {
    const team = normalizeTeam(person.team);
    if (!team) {
      return;
    }
    const key = team.toLowerCase();
    if (!teams.has(key)) {
      teams.set(key, team);
    }
  });

  return Array.from(teams.values()).sort((a, b) => a.localeCompare(b));
}

function renderTeamFilter() {
  if (!elements.teamFilter) {
    return;
  }

  const teams = getTeamOptions();
  elements.teamFilter.innerHTML = "";

  const allOption = new Option("All teams", TEAM_FILTER_ALL);
  elements.teamFilter.add(allOption);

  teams.forEach((team) => {
    elements.teamFilter.add(new Option(team, team));
  });

  elements.teamFilter.add(new Option("Unassigned", TEAM_FILTER_UNASSIGNED));

  if (state.selectedTeam === TEAM_FILTER_ALL || state.selectedTeam === TEAM_FILTER_UNASSIGNED) {
    elements.teamFilter.value = state.selectedTeam;
    return;
  }

  const match = teams.find(
    (team) => team.toLowerCase() === state.selectedTeam.toLowerCase()
  );
  if (match) {
    state.selectedTeam = match;
    elements.teamFilter.value = match;
  } else {
    state.selectedTeam = TEAM_FILTER_ALL;
    elements.teamFilter.value = TEAM_FILTER_ALL;
  }
}

function syncTeamInputWithFilter() {
  if (!elements.teamInput) {
    return;
  }
  if (elements.teamInput.value) {
    return;
  }
  if (
    state.selectedTeam !== TEAM_FILTER_ALL &&
    state.selectedTeam !== TEAM_FILTER_UNASSIGNED
  ) {
    elements.teamInput.value = state.selectedTeam;
  }
}

function onTeamFilterChange(event) {
  state.selectedTeam = event.target.value;
  syncTeamInputWithFilter();
  renderAttendanceList();
}

function filterPeopleByTeam(people) {
  if (state.selectedTeam === TEAM_FILTER_ALL) {
    return people;
  }

  if (state.selectedTeam === TEAM_FILTER_UNASSIGNED) {
    return people.filter((person) => !normalizeTeam(person.team));
  }

  const target = state.selectedTeam.toLowerCase();
  return people.filter(
    (person) => normalizeTeam(person.team).toLowerCase() === target
  );
}

function renderDirectoryList() {
  if (!elements.directoryList) {
    return;
  }

  elements.directoryList.innerHTML = "";

  if (!state.people.length) {
    const empty = document.createElement("div");
    empty.className = "hint";
    empty.textContent = "No people yet. Add someone above to get started.";
    elements.directoryList.appendChild(empty);
    return;
  }

  state.people.forEach((person) => {
    const row = document.createElement("div");
    row.className = "person-row directory";
    row.dataset.personId = person.id;

    const identity = document.createElement("div");
    identity.className = "person-identity";

    const nameInput = document.createElement("input");
    nameInput.className = "name-input";
    nameInput.placeholder = "Name";
    nameInput.value = person.name;
    nameInput.setAttribute("aria-label", `Name for ${person.name}`);

    const teamInput = document.createElement("input");
    teamInput.className = "team-input";
    teamInput.placeholder = "Team";
    teamInput.value = person.team || "";
    teamInput.setAttribute("aria-label", `Team for ${person.name}`);

    const availabilityGroup = document.createElement("div");
    availabilityGroup.className = "availability-group";
    const availability = normalizeAvailability(person.availability);

    WEEKDAYS.forEach((day) => {
      const button = document.createElement("button");
      button.className = "avail-btn";
      button.type = "button";
      button.textContent = day.label;
      button.dataset.day = day.key;
      button.setAttribute("aria-label", `${person.name} available ${day.full}`);
      button.classList.toggle("active", availability.includes(day.key));
      button.addEventListener("click", () => {
        toggleAvailability(person.id, day.key, button);
      });
      availabilityGroup.appendChild(button);
    });

    const removeBtn = document.createElement("button");
    removeBtn.className = "remove-btn";
    removeBtn.type = "button";
    removeBtn.textContent = "Remove";
    removeBtn.setAttribute("aria-label", `Remove ${person.name}`);

    teamInput.addEventListener("change", (event) => {
      updatePersonTeam(person.id, event.target.value);
    });
    teamInput.addEventListener("keydown", (event) => {
      if (event.key === "Enter") {
        event.preventDefault();
        teamInput.blur();
      }
    });
    nameInput.addEventListener("change", (event) => {
      updatePersonName(person.id, event.target.value, nameInput);
    });
    nameInput.addEventListener("keydown", (event) => {
      if (event.key === "Enter") {
        event.preventDefault();
        nameInput.blur();
      }
    });
    removeBtn.addEventListener("click", () => {
      requestRemovePerson(person.id, person.name, removeBtn);
    });

    identity.append(nameInput, teamInput);
    row.append(identity, availabilityGroup, removeBtn);
    elements.directoryList.appendChild(row);
  });
}

function renderAttendanceList() {
  if (!elements.attendanceList) {
    return;
  }

  renderTeamFilter();
  elements.attendanceList.innerHTML = "";

  const weekdayKey = getWeekdayKey(state.currentDate);

  if (elements.attendanceHint) {
    const weekday = getWeekdayLabel(state.currentDate);
    if (weekdayKey) {
      elements.attendanceHint.textContent = weekday;
    } else {
      elements.attendanceHint.textContent = `${weekday} (no schedule)`;
    }
  }

  if (!state.people.length) {
    const empty = document.createElement("div");
    empty.className = "hint";
    empty.textContent = "No people yet. Add someone in the directory to get started.";
    elements.attendanceList.appendChild(empty);
    return;
  }

  const visiblePeople = filterPeopleByTeam(state.people).filter((person) => {
    if (!weekdayKey) {
      return false;
    }
    const availability = normalizeAvailability(person.availability);
    return availability.includes(weekdayKey);
  });

  if (!visiblePeople.length) {
    const empty = document.createElement("div");
    empty.className = "hint";
    if (!weekdayKey) {
      empty.textContent = "No one scheduled for weekends.";
    } else if (state.selectedTeam === TEAM_FILTER_UNASSIGNED) {
      empty.textContent = "No unassigned people scheduled today.";
    } else if (state.selectedTeam === TEAM_FILTER_ALL) {
      empty.textContent = "No one scheduled today.";
    } else {
      empty.textContent = `No one in ${state.selectedTeam} scheduled today.`;
    }
    elements.attendanceList.appendChild(empty);
    return;
  }

  visiblePeople.forEach((person) => {
    const row = document.createElement("div");
    row.className = "person-row attendance";
    row.dataset.personId = person.id;

    const name = document.createElement("div");
    name.className = "person-name";
    name.textContent = person.name;

    const statusGroup = document.createElement("div");
    statusGroup.className = "status-group";

    const hereBtn = createStatusButton("✓", "here");
    const tardyBtn = createStatusButton("T", "tardy");
    const notBtn = createStatusButton("✗", "not");
    statusGroup.append(hereBtn, tardyBtn, notBtn);

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
    tardyBtn.addEventListener("click", () => toggleStatus(person.id, "tardy"));
    notBtn.addEventListener("click", () => toggleStatus(person.id, "not"));
    amBtn.addEventListener("click", () => cycleMeeting(person.id, "am"));
    pmBtn.addEventListener("click", () => cycleMeeting(person.id, "pm"));
    note.addEventListener("input", (event) => {
      setNote(person.id, event.target.value);
    });

    row.append(name, statusGroup, note, subGroup);
    elements.attendanceList.appendChild(row);

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

function requestRemovePerson(personId, name, button) {
  if (button.dataset.confirming === "true") {
    setRemoveButtonConfirming(button, false);
    removePerson(personId);
    return;
  }

  setRemoveButtonConfirming(button, true);
  setPeopleMessage(`Click confirm to remove ${name}.`, "warn");
}

function setRemoveButtonConfirming(button, confirming) {
  const timer = removeConfirmTimers.get(button);
  if (timer) {
    clearTimeout(timer);
    removeConfirmTimers.delete(button);
  }

  if (!confirming) {
    button.dataset.confirming = "false";
    button.classList.remove("confirming");
    button.textContent = "Remove";
    return;
  }

  button.dataset.confirming = "true";
  button.classList.add("confirming");
  button.textContent = "Confirm";
  const nextTimer = setTimeout(() => {
    setRemoveButtonConfirming(button, false);
  }, REMOVE_CONFIRM_MS);
  removeConfirmTimers.set(button, nextTimer);
}

async function removePerson(personId) {
  const index = state.people.findIndex((person) => person.id === personId);
  if (index === -1) {
    return;
  }

  const [removed] = state.people.splice(index, 1);
  delete state.attendance[personId];
  await savePeople();
  renderDirectoryList();
  renderAttendanceList();
  markDirty();
  setPeopleMessage(`${removed.name} removed.`, "muted");
}

async function updatePersonTeam(personId, value) {
  const person = state.people.find((entry) => entry.id === personId);
  if (!person) {
    return;
  }

  const nextTeam = normalizeTeam(value);
  if (person.team === nextTeam) {
    return;
  }

  person.team = nextTeam;
  await savePeople();
  renderDirectoryList();
  renderAttendanceList();
  setPeopleMessage(
    nextTeam ? `${person.name} set to ${nextTeam}.` : `${person.name} unassigned.`,
    "muted"
  );
}

async function updatePersonName(personId, value, input) {
  const person = state.people.find((entry) => entry.id === personId);
  if (!person) {
    return;
  }

  const cleaned = normalizeName(value || "");
  const currentName = person.name;

  if (!cleaned) {
    if (input) {
      input.value = currentName;
    }
    setPeopleMessage("Name can't be empty.", "warn");
    return;
  }

  const duplicate = state.people.some(
    (entry) =>
      entry.id !== personId && entry.name.toLowerCase() === cleaned.toLowerCase()
  );

  if (duplicate) {
    if (input) {
      input.value = currentName;
    }
    setPeopleMessage("That name already exists.", "warn");
    return;
  }

  if (cleaned === currentName) {
    if (input) {
      input.value = currentName;
    }
    return;
  }

  person.name = cleaned;
  await savePeople();
  renderDirectoryList();
  renderAttendanceList();
  setPeopleMessage(`Renamed to ${cleaned}.`, "success");
}

async function toggleAvailability(personId, dayKey, button) {
  const person = state.people.find((entry) => entry.id === personId);
  if (!person) {
    return;
  }

  const availability = normalizeAvailability(person.availability);
  const index = availability.indexOf(dayKey);
  if (index === -1) {
    availability.push(dayKey);
  } else {
    availability.splice(index, 1);
  }

  person.availability = availability;
  if (button) {
    button.classList.toggle("active", availability.includes(dayKey));
  }

  await savePeople();
  renderAttendanceList();
}

function applyRowState(personId, row) {
  const entry = getEntry(personId);
  const hereBtn = row.querySelector(".status-btn.here");
  const tardyBtn = row.querySelector(".status-btn.tardy");
  const notBtn = row.querySelector(".status-btn.not");
  hereBtn.classList.toggle("active", entry.status === "here");
  if (tardyBtn) {
    tardyBtn.classList.toggle("active", entry.status === "tardy");
  }
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
  if (!elements.attendanceList) {
    return;
  }
  const row = elements.attendanceList.querySelector(
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
  scheduleAutoSave();
}

function setSaveStatus(message) {
  elements.saveStatus.textContent = message;
}

function clearAutoSaveTimer() {
  if (!autoSaveTimer) {
    return;
  }
  clearTimeout(autoSaveTimer);
  autoSaveTimer = null;
}

function scheduleAutoSave() {
  clearAutoSaveTimer();
  autoSaveTimer = setTimeout(async () => {
    autoSaveTimer = null;
    if (!state.dirty || saveInProgress) {
      return;
    }
    await onSave();
  }, AUTO_SAVE_DELAY_MS);
}

async function onSave() {
  if (saveInProgress) {
    return;
  }
  saveInProgress = true;
  const payload = buildAttendancePayload();

  try {
    if (state.firebaseReady) {
      await saveAttendance(state.currentDate, payload);
    }

    saveLocalBackup(payload, state.currentDate);
    state.dirty = false;
    setSaveStatus(`Saved ${new Date().toLocaleTimeString()}`);
    clearAutoSaveTimer();

    await loadMonth(state.monthCursor);
    await loadTrends();
  } finally {
    saveInProgress = false;
  }
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
        state.people = sanitizePeopleList(DEFAULT_PEOPLE);
      }
      renderDirectoryList();
      renderAttendanceList();
      return;
    }

    state.people = sanitizePeopleList(data.people);
    renderDirectoryList();
    renderAttendanceList();
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
      renderAttendanceList();
      setSaveStatus("");
      return;
    }

    state.attendance = mapAttendanceEntries(data.people || {});
    renderAttendanceList();
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
  const basePeople =
    local.people && local.people.length ? local.people : DEFAULT_PEOPLE;
  state.people = sanitizePeopleList(basePeople);
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
      renderAttendanceList();
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
      if (entry.status === "here" || entry.status === "tardy") {
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

  if (!state.people.length) {
    const empty = document.createElement("div");
    empty.className = "hint";
    empty.textContent = "No profiles in the directory yet.";
    elements.trendsList.appendChild(empty);
    return;
  }

  const stats = new Map();

  state.people.forEach((person) => {
    stats.set(person.id, {
      id: person.id,
      name: person.name,
      here: 0,
      not: 0,
      total: 0,
      tardy: 0,
    });
  });

  state.trendsAttendance.forEach((entry) => {
    if (!entry?.people) {
      return;
    }
    Object.entries(entry.people).forEach(([id, personEntry]) => {
      const record = stats.get(id);
      if (!record) {
        return;
      }

      if (personEntry.status) {
        record.total += 1;
        if (personEntry.status === "here" || personEntry.status === "tardy") {
          record.here += 1;
        } else if (personEntry.status === "not") {
          record.not += 1;
        }
        if (personEntry.status === "tardy") {
          record.tardy += 1;
        }
      }

      if (
        personEntry.note &&
        LATE_REGEX.test(personEntry.note) &&
        personEntry.status !== "tardy"
      ) {
        record.tardy += 1;
      }
    });
  });

  const sorted = Array.from(stats.values()).sort((a, b) =>
    a.name.localeCompare(b.name)
  );

  const hasAnyData = sorted.some(
    (record) => record.total > 0 || record.tardy > 0
  );

  if (!hasAnyData) {
    const empty = document.createElement("div");
    empty.className = "hint";
    empty.textContent = "No attendance data yet for current profiles.";
    elements.trendsList.appendChild(empty);
    return;
  }

  const header = document.createElement("div");
  header.className = "trend-row trend-header";
  header.innerHTML = `
    <div>Person</div>
    <div class="metric">Here %</div>
    <div class="metric">Tardy</div>
    <div class="metric">Not</div>
    <div class="metric">Total</div>
  `;
  elements.trendsList.appendChild(header);

  sorted.forEach((record) => {
    if (record.total === 0 && record.tardy === 0) {
      return;
    }
    const row = document.createElement("div");
    row.className = "trend-row";

    const name = document.createElement("div");
    name.className = "name";
    name.textContent = record.name;

    const percent = record.total
      ? Math.round((record.here / record.total) * 100)
      : 0;
    const here = document.createElement("div");
    here.className = "metric";
    here.textContent = `${percent}%`;

    const tardy = document.createElement("div");
    tardy.className = "metric";
    tardy.textContent = `${record.tardy}`;

    const not = document.createElement("div");
    not.className = "metric";
    not.textContent = `${record.not}`;

    const total = document.createElement("div");
    total.className = "metric";
    total.textContent = `${record.total}`;

    row.append(name, here, tardy, not, total);
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

    const people = sanitizePeopleList(data.people);

    saveLocalStore({
      people,
      attendance: data.attendance,
    });

    state.people = people;
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
