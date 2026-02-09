import test from "node:test";
import assert from "node:assert/strict";
import {
  addDays,
  countAttendance,
  formatDateKey,
  getLocalAttendanceForDate,
  mapAttendanceEntries,
  parseDateKey,
  summarizeEntry,
} from "../attendance-core.js";

test("date keys round-trip and reject invalid inputs", () => {
  const date = parseDateKey("2026-02-09");
  assert.ok(date);
  assert.equal(formatDateKey(date), "2026-02-09");

  assert.equal(parseDateKey("2026-02-30"), null);
  assert.equal(parseDateKey("2026-2-9"), null);
  assert.equal(parseDateKey("invalid"), null);
});

test("addDays stays calendar-correct across boundaries", () => {
  assert.equal(addDays("2026-02-28", 1), "2026-03-01");
  assert.equal(addDays("2026-01-01", -1), "2025-12-31");
});

test("mapAttendanceEntries sanitizes malformed records", () => {
  const mapped = mapAttendanceEntries({
    valid: { status: "here", note: "ok", am: "here", pm: null },
    badNumber: 4,
    badNull: null,
  });

  assert.deepEqual(mapped, {
    valid: { status: "here", note: "ok", am: "here", pm: null },
  });
});

test("getLocalAttendanceForDate supports payload and legacy shapes", () => {
  const payloadShape = {
    "2026-02-09": {
      date: "2026-02-09",
      people: {
        a: { status: "here", note: "", am: null, pm: null },
      },
    },
  };
  const legacyShape = {
    "2026-02-10": {
      b: { status: "not", note: "sick", am: null, pm: null },
    },
  };

  assert.deepEqual(getLocalAttendanceForDate(payloadShape, "2026-02-09"), {
    a: { status: "here", note: "", am: null, pm: null },
  });
  assert.deepEqual(getLocalAttendanceForDate(legacyShape, "2026-02-10"), {
    b: { status: "not", note: "sick", am: null, pm: null },
  });
});

test("summarizeEntry + countAttendance preserve attendance math", () => {
  const tardySingle = summarizeEntry({ status: "tardy", note: "", am: null, pm: null });
  assert.deepEqual(tardySingle, { here: 1, not: 0, total: 1, tardy: 1 });

  const splitDay = summarizeEntry({ status: "here", note: "", am: "here", pm: "not" });
  assert.deepEqual(splitDay, { here: 1, not: 1, total: 2, tardy: 0 });

  const totals = countAttendance({
    a: { status: "here", note: "", am: null, pm: null },
    b: { status: "not", note: "", am: "not", pm: "not" },
    c: { status: "tardy", note: "", am: "here", pm: null },
  });
  assert.deepEqual(totals, { hereCount: 3, notCount: 2, total: 5 });
});
