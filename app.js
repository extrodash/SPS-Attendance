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
import {
  addDays,
  countAttendance,
  formatDateKey,
  getLocalAttendanceForDate,
  mapAttendanceEntries,
  parseDateKey,
  summarizeEntry,
  todayISO,
} from "./attendance-core.js";

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
  directoryTeamFilter: document.getElementById("directory-team-filter"),
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
  trendsRange: document.getElementById("trends-range"),
  teamFilterAllBtn: document.getElementById("team-filter-all"),
};

const LOCAL_KEY = "attendance-app-data-v1";
const DEFAULT_PEOPLE = [];
const TEAM_FILTER_ALL = "__all__";
const TEAM_FILTER_NONE = "__none__";
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
const TREND_RANGES = {
  week: { days: 7, label: "Past week" },
  month: { days: 30, label: "Past month" },
  quarter: { days: 90, label: "Past 3 months" },
  all: { days: null, label: "All time" },
};

const state = {
  currentDate: todayISO(),
  people: [],
  attendance: {},
  monthCursor: new Date(),
  monthAttendance: {},
  trendsAttendance: [],
  dirty: false,
  firebaseReady: false,
  selectedTeam: TEAM_FILTER_NONE,
  teamFilterAll: false,
  directoryTeamFilter: TEAM_FILTER_ALL,
  directoryOpen: false,
  trendsRange: "quarter",
  trendsExpandedTeams: new Set(),
};

let db = null;
let auth = null;
let unsubscribeAttendance = null;
let unsubscribePeople = null;
let unsubscribeMonth = null;
let peopleMessageTimer = null;
let autoSaveTimer = null;
let saveInProgress = false;
let trendsLoadSeq = 0;
const removeConfirmTimers = new WeakMap();
const REMOVE_CONFIRM_MS = 3500;
const meetingStatusPrompted = new Set();

init();

async function init() {
  elements.datePicker.value = state.currentDate;
  elements.datePicker.addEventListener("change", onDateChange);
  elements.saveBtn.addEventListener("click", onSave);
  elements.addPersonBtn.addEventListener("click", addPersonFromInput);
  elements.teamFilter.addEventListener("change", onTeamFilterChange);
  if (elements.teamFilterAllBtn) {
    elements.teamFilterAllBtn.addEventListener("click", toggleTeamFilterAll);
  }
  if (elements.directoryTeamFilter) {
    elements.directoryTeamFilter.addEventListener(
      "change",
      onDirectoryTeamFilterChange
    );
  }
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
  if (elements.trendsRange) {
    elements.trendsRange.value = state.trendsRange;
    elements.trendsRange.addEventListener("change", onTrendsRangeChange);
  }
  setDirectoryOpen(state.directoryOpen);

  if ("serviceWorker" in navigator) {
    navigator.serviceWorker.register("sw.js");
  }

  const hasConfig = isFirebaseConfigured();
  if (!hasConfig) {
    elements.configWarning.hidden = false;
    hydrateFromLocal();
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

async function onDateChange(event) {
  const nextDate = event?.target?.value;
  const switched = await switchDate(nextDate);
  if (!switched && elements.datePicker) {
    elements.datePicker.value = state.currentDate;
  }
}

async function switchDate(nextDate) {
  if (!parseDateKey(nextDate)) {
    return false;
  }

  if (nextDate === state.currentDate) {
    if (elements.datePicker) {
      elements.datePicker.value = nextDate;
    }
    return true;
  }

  if (state.dirty) {
    setSaveStatus("Saving...");
    const saved = await onSave();
    if (!saved) {
      setSaveStatus("Could not save current day.");
      return false;
    }
  }

  state.currentDate = nextDate;
  state.dirty = false;
  clearAutoSaveTimer();
  meetingStatusPrompted.clear();
  setSaveStatus("");
  if (elements.datePicker) {
    elements.datePicker.value = nextDate;
  }

  if (state.firebaseReady) {
    listenAttendance(nextDate);
  } else {
    hydrateAttendanceFromLocal();
    renderAttendanceList();
  }

  return true;
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
    !state.teamFilterAll && state.selectedTeam !== TEAM_FILTER_NONE
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
  renderTeamScopedViews();
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
  const date = parseDateKey(dateISO);
  if (!date) {
    return null;
  }
  const day = date.getDay();
  const match = WEEKDAYS.find((entry) => entry.index === day);
  return match ? match.key : null;
}

function getWeekdayLabel(dateISO) {
  const date = parseDateKey(dateISO);
  if (!date) {
    return "";
  }
  return date.toLocaleDateString("en-US", { weekday: "long" });
}

function renderAll() {
  renderDirectoryList();
  renderAttendanceList();
  renderCalendar();
  renderTrends();
}

function renderTeamScopedViews() {
  renderAttendanceList();
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

  const noneOption = new Option("Not Selected", TEAM_FILTER_NONE);
  elements.teamFilter.add(noneOption);

  teams.forEach((team) => {
    elements.teamFilter.add(new Option(team, team));
  });

  const match = teams.find(
    (team) => team.toLowerCase() === state.selectedTeam.toLowerCase()
  );
  if (match) {
    state.selectedTeam = match;
    elements.teamFilter.value = match;
  } else {
    state.selectedTeam = TEAM_FILTER_NONE;
    elements.teamFilter.value = TEAM_FILTER_NONE;
  }

  updateTeamFilterUI();
}

function syncTeamInputWithFilter() {
  if (!elements.teamInput) {
    return;
  }
  if (elements.teamInput.value) {
    return;
  }
  if (
    !state.teamFilterAll &&
    state.selectedTeam !== TEAM_FILTER_NONE &&
    state.selectedTeam !== TEAM_FILTER_UNASSIGNED
  ) {
    elements.teamInput.value = state.selectedTeam;
  }
}

function onTeamFilterChange(event) {
  state.selectedTeam = event.target.value;
  state.teamFilterAll = false;
  updateTeamFilterUI();
  syncTeamInputWithFilter();
  renderTeamScopedViews();
}

function onDirectoryTeamFilterChange(event) {
  state.directoryTeamFilter = event.target.value;
  renderDirectoryList();
}

function onTrendsRangeChange(event) {
  const nextRange = event?.target?.value;
  if (!TREND_RANGES[nextRange]) {
    return;
  }
  state.trendsRange = nextRange;
  loadTrends();
}

function toggleTeamFilterAll() {
  state.teamFilterAll = !state.teamFilterAll;
  updateTeamFilterUI();
  renderTeamScopedViews();
}

function updateTeamFilterUI() {
  if (elements.teamFilterAllBtn) {
    elements.teamFilterAllBtn.classList.toggle("active", state.teamFilterAll);
    elements.teamFilterAllBtn.setAttribute(
      "aria-pressed",
      state.teamFilterAll ? "true" : "false"
    );
  }
}

function filterPeopleByTeam(people, selectedTeam = state.selectedTeam) {
  if (selectedTeam === TEAM_FILTER_ALL) {
    return people;
  }

  if (selectedTeam === TEAM_FILTER_NONE) {
    return [];
  }

  if (selectedTeam === TEAM_FILTER_UNASSIGNED) {
    return people.filter((person) => !normalizeTeam(person.team));
  }

  const target = selectedTeam.toLowerCase();
  return people.filter(
    (person) => normalizeTeam(person.team).toLowerCase() === target
  );
}

function renderDirectoryFilter() {
  if (!elements.directoryTeamFilter) {
    return;
  }

  const teams = getTeamOptions();
  elements.directoryTeamFilter.innerHTML = "";

  const allOption = new Option("All teams", TEAM_FILTER_ALL);
  elements.directoryTeamFilter.add(allOption);

  teams.forEach((team) => {
    elements.directoryTeamFilter.add(new Option(team, team));
  });

  elements.directoryTeamFilter.add(
    new Option("Unassigned", TEAM_FILTER_UNASSIGNED)
  );

  if (
    state.directoryTeamFilter === TEAM_FILTER_ALL ||
    state.directoryTeamFilter === TEAM_FILTER_UNASSIGNED
  ) {
    elements.directoryTeamFilter.value = state.directoryTeamFilter;
    return;
  }

  const match = teams.find(
    (team) => team.toLowerCase() === state.directoryTeamFilter.toLowerCase()
  );

  if (match) {
    state.directoryTeamFilter = match;
    elements.directoryTeamFilter.value = match;
  } else {
    state.directoryTeamFilter = TEAM_FILTER_ALL;
    elements.directoryTeamFilter.value = TEAM_FILTER_ALL;
  }
}

function renderDirectoryList() {
  if (!elements.directoryList) {
    return;
  }

  renderDirectoryFilter();
  elements.directoryList.innerHTML = "";

  if (!state.people.length) {
    const empty = document.createElement("div");
    empty.className = "hint";
    empty.textContent = "No people yet. Add someone above to get started.";
    elements.directoryList.appendChild(empty);
    return;
  }

  const visiblePeople = filterPeopleByTeam(
    state.people,
    state.directoryTeamFilter
  );

  if (!visiblePeople.length) {
    const empty = document.createElement("div");
    empty.className = "hint";
    if (state.directoryTeamFilter === TEAM_FILTER_UNASSIGNED) {
      empty.textContent = "No unassigned people.";
    } else {
      empty.textContent = `No one in ${state.directoryTeamFilter}.`;
    }
    elements.directoryList.appendChild(empty);
    return;
  }

  visiblePeople.forEach((person) => {
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

  if (!state.teamFilterAll && state.selectedTeam === TEAM_FILTER_NONE) {
    const empty = document.createElement("div");
    empty.className = "hint";
    empty.textContent = "Select a team to view attendance.";
    elements.attendanceList.appendChild(empty);
    return;
  }

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

  const teamPeople = state.teamFilterAll
    ? state.people
    : filterPeopleByTeam(state.people, state.selectedTeam);

  const visiblePeople = teamPeople.filter((person) => {
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
    } else if (state.teamFilterAll) {
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
  renderTeamScopedViews();
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
  renderTeamScopedViews();
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
  renderTeamScopedViews();
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
  if (!entry.status) {
    const promptKey = `${state.currentDate}:${personId}`;
    if (!meetingStatusPrompted.has(promptKey)) {
      meetingStatusPrompted.add(promptKey);
      window.alert("Mark general attendance (Here/Tardy/Not) before AM/PM.");
    }
  }
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
    return false;
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
    return true;
  } catch (error) {
    console.warn("Save error", error);
    setSaveStatus("Save failed. Please retry.");
    return false;
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
      renderTeamScopedViews();
      return;
    }

    state.people = sanitizePeopleList(data.people);
    renderDirectoryList();
    renderTeamScopedViews();
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

function hydrateFromLocal() {
  const local = loadLocalStore();
  const basePeople =
    local.people && local.people.length ? local.people : DEFAULT_PEOPLE;
  const attendanceStore = local.attendance || {};
  state.people = sanitizePeopleList(basePeople);
  state.attendance = getLocalAttendanceForDate(
    attendanceStore,
    state.currentDate
  );
  state.monthAttendance = attendanceStore;
  renderAll();
  loadTrends();
}

function hydrateAttendanceFromLocal() {
  const local = loadLocalStore();
  state.attendance = getLocalAttendanceForDate(
    local.attendance || {},
    state.currentDate
  );
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
        const entry = attendance[key];
        filtered[key] = entry?.people
          ? entry
          : {
              date: key,
              people: entry && typeof entry === "object" ? entry : {},
            };
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
    start: formatDateKey(start),
    end: formatDateKey(end),
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
    const dateKey = formatDateKey(cursor);
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
    switchDate(dateKey);
  });

  cell.append(day, meta);
  return cell;
}

async function loadTrends() {
  const loadSeq = ++trendsLoadSeq;
  const { start, end } = resolveTrendsRange();

  if (!state.firebaseReady) {
    const local = loadLocalStore();
    const attendance = local.attendance || {};
    const records = Object.entries(attendance)
      .filter(([dateKey]) => !start || (dateKey >= start && dateKey <= end))
      .map(([dateKey, entry]) => normalizeTrendEntry(dateKey, entry));
    if (loadSeq !== trendsLoadSeq) {
      return;
    }
    state.trendsAttendance = records;
    renderTrends();
    return;
  }

  try {
    let snapshot;
    if (!start) {
      snapshot = await getDocs(collection(db, "attendance"));
    } else {
      const q = query(
        collection(db, "attendance"),
        where("date", ">=", start),
        where("date", "<=", end)
      );
      snapshot = await getDocs(q);
    }

    if (loadSeq !== trendsLoadSeq) {
      return;
    }

    state.trendsAttendance = snapshot.docs
      .map((docSnap) => normalizeTrendEntry(docSnap.id, docSnap.data()))
      .filter((entry) => !start || (entry.date >= start && entry.date <= end));
    renderTrends();
  } catch (error) {
    if (loadSeq !== trendsLoadSeq) {
      return;
    }
    console.warn("Trend load error", error);
    state.trendsAttendance = [];
    renderTrends();
  }
}

function resolveTrendsRange() {
  const rangeKey = TREND_RANGES[state.trendsRange] ? state.trendsRange : "quarter";
  if (rangeKey !== state.trendsRange) {
    state.trendsRange = rangeKey;
    if (elements.trendsRange) {
      elements.trendsRange.value = rangeKey;
    }
  }

  const end = todayISO();
  const config = TREND_RANGES[rangeKey];
  if (config.days === null) {
    return { start: null, end };
  }

  return {
    start: addDays(end, -(config.days - 1)),
    end,
  };
}

function normalizeTrendEntry(dateKey, entry) {
  const source = entry && typeof entry === "object" ? entry : {};
  if (source.people && typeof source.people === "object") {
    return {
      date: source.date || dateKey,
      people: source.people,
    };
  }

  return {
    date: dateKey,
    people: source,
  };
}

function buildTrendTeams() {
  const map = new Map();
  state.people.forEach((person) => {
    const teamName = normalizeTeam(person.team);
    if (!teamName) {
      return;
    }
    const teamKey = teamName.toLowerCase();
    if (!map.has(teamKey)) {
      map.set(teamKey, { key: teamKey, name: teamName, members: [] });
    }
    map.get(teamKey).members.push(person);
  });

  const teams = Array.from(map.values()).sort((a, b) =>
    a.name.localeCompare(b.name)
  );
  teams.forEach((team) => {
    team.members.sort((a, b) => a.name.localeCompare(b.name));
  });
  return teams;
}

function buildTrendPersonStats(teams) {
  const stats = new Map();
  teams.forEach((team) => {
    team.members.forEach((person) => {
      stats.set(person.id, {
        id: person.id,
        name: person.name,
        here: 0,
        not: 0,
        total: 0,
        tardy: 0,
      });
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
      const summary = summarizeEntry(personEntry);
      record.total += summary.total;
      record.here += summary.here;
      record.not += summary.not;
      record.tardy += summary.tardy;
    });
  });

  return stats;
}

function teamTotals(team, stats) {
  return team.members.reduce(
    (totals, person) => {
      const record = stats.get(person.id);
      if (!record) {
        return totals;
      }
      totals.here += record.here;
      totals.not += record.not;
      totals.total += record.total;
      totals.tardy += record.tardy;
      return totals;
    },
    { here: 0, not: 0, total: 0, tardy: 0 }
  );
}

function formatTrendPercent(here, total) {
  if (!total) {
    return "--";
  }
  return `${Math.round((here / total) * 100)}%`;
}

function createTrendDetailTable(team, stats) {
  const wrapper = document.createElement("div");
  wrapper.className = "trend-table-wrap";

  const table = document.createElement("table");
  table.className = "trend-table";
  table.innerHTML = `
    <colgroup>
      <col />
      <col />
      <col />
      <col />
      <col />
    </colgroup>
  `;

  const thead = document.createElement("thead");
  thead.innerHTML = `
    <tr>
      <th scope="col">Person</th>
      <th scope="col">Here %</th>
      <th scope="col">Tardy</th>
      <th scope="col">Not</th>
      <th scope="col">Total</th>
    </tr>
  `;

  const tbody = document.createElement("tbody");
  team.members.forEach((person) => {
    const record = stats.get(person.id) || {
      name: person.name,
      here: 0,
      not: 0,
      total: 0,
      tardy: 0,
    };

    const row = document.createElement("tr");
    if (!record.total && !record.tardy) {
      row.className = "trend-row-empty";
    }

    const name = document.createElement("th");
    name.scope = "row";
    name.textContent = record.name;

    const herePercent = document.createElement("td");
    herePercent.className = "num";
    herePercent.textContent = formatTrendPercent(record.here, record.total);

    const tardy = document.createElement("td");
    tardy.className = "num";
    tardy.textContent = `${record.tardy}`;

    const not = document.createElement("td");
    not.className = "num";
    not.textContent = `${record.not}`;

    const total = document.createElement("td");
    total.className = "num";
    total.textContent = `${record.total}`;

    row.append(name, herePercent, tardy, not, total);
    tbody.appendChild(row);
  });

  table.append(thead, tbody);
  wrapper.appendChild(table);
  return wrapper;
}

function createTrendTeamCard(team, stats) {
  const totals = teamTotals(team, stats);
  const percent = formatTrendPercent(totals.here, totals.total);
  const isOpen = state.trendsExpandedTeams.has(team.key);
  const detailId = `trend-team-${team.key.replace(/[^a-z0-9_-]/g, "-")}`;
  const rangeLabel = TREND_RANGES[state.trendsRange]?.label || "Past 3 months";

  const card = document.createElement("div");
  card.className = "trend-team-card";
  card.dataset.open = isOpen ? "true" : "false";

  const header = document.createElement("div");
  header.className = "trend-team-header";

  const identity = document.createElement("div");
  identity.className = "trend-team-identity";

  const name = document.createElement("div");
  name.className = "trend-team-name";
  name.textContent = team.name;

  const memberCount = document.createElement("div");
  memberCount.className = "trend-team-meta";
  memberCount.textContent = `${team.members.length} member${team.members.length === 1 ? "" : "s"}`;

  identity.append(name, memberCount);

  const actions = document.createElement("div");
  actions.className = "trend-team-actions";

  const kpi = document.createElement("div");
  kpi.className = "trend-team-kpi";

  const percentText = document.createElement("div");
  percentText.className = "trend-team-percent";
  percentText.textContent = percent;

  const countText = document.createElement("div");
  countText.className = "trend-team-count";
  countText.textContent = totals.total
    ? `${totals.here}/${totals.total} here`
    : "No check-ins";

  const periodText = document.createElement("div");
  periodText.className = "trend-team-period";
  periodText.textContent = rangeLabel;

  kpi.append(percentText, countText, periodText);

  const toggle = document.createElement("button");
  toggle.className = "trend-team-toggle";
  toggle.type = "button";
  toggle.setAttribute("aria-expanded", isOpen ? "true" : "false");
  toggle.setAttribute("aria-controls", detailId);
  toggle.setAttribute("aria-label", `Toggle ${team.name} details`);

  const caret = document.createElement("span");
  caret.className = "trend-caret";
  caret.textContent = ">";

  const label = document.createElement("span");
  label.className = "trend-toggle-label";
  label.textContent = isOpen ? "Hide details" : "Show details";

  toggle.append(caret, label);
  toggle.addEventListener("click", () => {
    if (state.trendsExpandedTeams.has(team.key)) {
      state.trendsExpandedTeams.delete(team.key);
    } else {
      state.trendsExpandedTeams.add(team.key);
    }
    renderTrends();
  });

  actions.append(kpi, toggle);
  header.append(identity, actions);
  card.appendChild(header);

  const bar = document.createElement("div");
  bar.className = "trend-team-bar";
  const fill = document.createElement("div");
  fill.className = "trend-team-fill";
  fill.style.width = totals.total ? `${Math.max(4, (totals.here / totals.total) * 100)}%` : "0%";
  bar.appendChild(fill);
  card.appendChild(bar);

  const detail = document.createElement("div");
  detail.id = detailId;
  detail.className = "trend-team-detail";
  detail.hidden = !isOpen;

  if (isOpen) {
    detail.appendChild(createTrendDetailTable(team, stats));
  }

  card.appendChild(detail);
  return card;
}

function renderTrends() {
  elements.trendsList.innerHTML = "";

  const teams = buildTrendTeams();

  if (!teams.length) {
    const empty = document.createElement("div");
    empty.className = "trend-team-placeholder";
    empty.textContent = "No teams in the directory yet.";
    elements.trendsList.appendChild(empty);
    return;
  }

  const validKeys = new Set(teams.map((team) => team.key));
  state.trendsExpandedTeams = new Set(
    [...state.trendsExpandedTeams].filter((key) => validKeys.has(key))
  );

  const stats = buildTrendPersonStats(teams);
  teams.forEach((team) => {
    elements.trendsList.appendChild(createTrendTeamCard(team, stats));
  });
}
