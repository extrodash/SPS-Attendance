export function formatDateKey(date) {
  const year = date.getFullYear();
  const month = String(date.getMonth() + 1).padStart(2, "0");
  const day = String(date.getDate()).padStart(2, "0");
  return `${year}-${month}-${day}`;
}

export function parseDateKey(dateKey) {
  if (typeof dateKey !== "string") {
    return null;
  }

  const match = /^(\d{4})-(\d{2})-(\d{2})$/.exec(dateKey);
  if (!match) {
    return null;
  }

  const year = Number(match[1]);
  const monthIndex = Number(match[2]) - 1;
  const day = Number(match[3]);
  const date = new Date(year, monthIndex, day);

  if (
    date.getFullYear() !== year ||
    date.getMonth() !== monthIndex ||
    date.getDate() !== day
  ) {
    return null;
  }

  return date;
}

export function todayISO(now = new Date()) {
  return formatDateKey(now);
}

export function addDays(dateString, delta) {
  const date = parseDateKey(dateString);
  if (!date) {
    return todayISO();
  }
  date.setDate(date.getDate() + delta);
  return formatDateKey(date);
}

export function mapAttendanceEntries(peopleData) {
  const mapped = {};
  Object.entries(peopleData || {}).forEach(([id, entry]) => {
    if (!entry || typeof entry !== "object") {
      return;
    }
    mapped[id] = {
      status: entry.status || null,
      note: entry.note || "",
      am: entry.am || null,
      pm: entry.pm || null,
    };
  });
  return mapped;
}

export function getLocalAttendanceForDate(attendanceStore, date) {
  const record = attendanceStore?.[date];
  if (!record || typeof record !== "object") {
    return {};
  }

  if (record.people && typeof record.people === "object") {
    return mapAttendanceEntries(record.people);
  }

  return mapAttendanceEntries(record);
}

export function summarizeEntry(entry) {
  const summary = { here: 0, not: 0, total: 0, tardy: 0 };
  if (!entry || !entry.status) {
    return summary;
  }

  if (entry.status === "tardy") {
    summary.tardy = 1;
  }

  const hasAm = entry.am === "here" || entry.am === "not";
  const hasPm = entry.pm === "here" || entry.pm === "not";

  if (!hasAm && !hasPm) {
    summary.total = 1;
    if (entry.status === "here" || entry.status === "tardy") {
      summary.here = 1;
    } else if (entry.status === "not") {
      summary.not = 1;
    }
    return summary;
  }

  const generalHere = entry.status === "here" || entry.status === "tardy";
  const generalNot = entry.status === "not";
  const amValue = hasAm ? entry.am : generalHere ? "here" : generalNot ? "not" : null;
  const pmValue = hasPm ? entry.pm : generalHere ? "here" : generalNot ? "not" : null;

  [amValue, pmValue].forEach((value) => {
    if (!value) {
      return;
    }
    summary.total += 1;
    if (value === "here") {
      summary.here += 1;
    } else if (value === "not") {
      summary.not += 1;
    }
  });

  return summary;
}

export function countAttendance(peopleData) {
  const entries = Object.values(peopleData || {});
  let hereCount = 0;
  let notCount = 0;
  let total = 0;
  entries.forEach((entry) => {
    const summary = summarizeEntry(entry);
    hereCount += summary.here;
    notCount += summary.not;
    total += summary.total;
  });
  return { hereCount, notCount, total };
}
