export const MORNING_START = 8 * 60;
export const MORNING_END = 12 * 60;
export const AFTERNOON_START = 13 * 60 + 30;
export const AFTERNOON_END = 17 * 60 + 30;
const PUNCH_FROM = 6 * 60;
const PUNCH_BEFORE = 19 * 60;

export type CalendarDate = {
  year: number;
  month: number;
  day: number;
};

export type PunchDay = CalendarDate & {
  weekday: number;
  punches: string[];
};

export type ParsedEmployee = {
  name: string;
  days: PunchDay[];
};

export type ParsedAttendance = {
  periodLabel: string;
  start: CalendarDate;
  end: CalendarDate;
  employees: ParsedEmployee[];
};

export type SummaryRow = {
  name: string;
  morning: string;
  afternoon: string;
  fullDay: string;
  punch: string;
  total: number | string;
};

export const punchNoteSession = "quên chấm";

export type ReviewNote = {
  name: string;
  dateLabel: string;
  session: string;
  detail: string;
};

export type AttendanceSummary = {
  title: string;
  periodLabel: string;
  rows: SummaryRow[];
  notes: ReviewNote[];
};

export type ExcludedSpan = {
  start: CalendarDate;
  end: CalendarDate;
};

type Sheet = {
  name: string;
  rows: unknown[][];
};

function cellText(value: unknown): string {
  if (value == null) return "";
  return String(value).replace(/\u00a0/g, " ").trim();
}

function fold(value: string): string {
  return value.normalize("NFD").replace(/\p{M}/gu, "").toLowerCase();
}

export function formatMinutes(total: number): string {
  const minutes = ((total % (24 * 60)) + 24 * 60) % (24 * 60);
  const hour = Math.floor(minutes / 60);
  const minute = minutes % 60;
  return `${String(hour).padStart(2, "0")}:${String(minute).padStart(2, "0")}`;
}

export function formatDate(date: CalendarDate): string {
  return `${String(date.day).padStart(2, "0")}/${String(date.month).padStart(2, "0")}`;
}

export function formatTotal(total: number): string {
  if (Number.isInteger(total)) return String(total);
  return total.toLocaleString("vi-VN", { maximumFractionDigits: 1 });
}

function inclusiveDays(startDay: number, startMonth: number, endDay: number, endMonth: number): number | null {
  if (startMonth < 1 || startMonth > 12 || endMonth < 1 || endMonth > 12) return null;
  if (startDay < 1 || startDay > 31 || endDay < 1 || endDay > 31) return null;
  const year = 2024;
  const start = new Date(Date.UTC(year, startMonth - 1, startDay));
  if (start.getUTCDate() !== startDay || start.getUTCMonth() !== startMonth - 1) return null;
  let end = new Date(Date.UTC(year, endMonth - 1, endDay));
  if (end.getUTCDate() !== endDay || end.getUTCMonth() !== endMonth - 1) {
    return null;
  }
  if (end.getTime() < start.getTime()) {
    end = new Date(Date.UTC(year + 1, endMonth - 1, endDay));
  }
  return Math.round((end.getTime() - start.getTime()) / 86400000) + 1;
}

function dateTokenKey(month: number, day: number): string {
  return `${month}-${day}`;
}

export function listedDateKeys(text: string): Set<string> {
  const seen = new Set<string>();
  for (const part of text.split(/[,;\n]+/)) {
    const token = part.trim();
    if (!token) continue;
    const range = token.match(/^(\d{1,2})\/(\d{1,2})\s*-\s*(\d{1,2})\/(\d{1,2})$/);
    if (range) {
      const span = inclusiveDays(Number(range[1]), Number(range[2]), Number(range[3]), Number(range[4]));
      if (span == null) continue;
      const cursor = new Date(Date.UTC(2024, Number(range[2]) - 1, Number(range[1])));
      for (let index = 0; index < span; index += 1) {
        const day = new Date(cursor.getTime() + index * 86400000);
        seen.add(dateTokenKey(day.getUTCMonth() + 1, day.getUTCDate()));
      }
      continue;
    }
    const single = token.match(/^(\d{1,2})\/(\d{1,2})$/);
    if (!single) continue;
    const day = Number(single[1]);
    const month = Number(single[2]);
    if (inclusiveDays(day, month, day, month) == null) continue;
    seen.add(dateTokenKey(month, day));
  }
  return seen;
}

export function countListedDays(text: string): number {
  return listedDateKeys(text).size;
}

export function notesMatchingRows(
  notes: ReviewNote[],
  sourceRows: Array<{ name: string }>,
  editedRows: Array<{ name: string; morning: string; afternoon: string; fullDay: string; punch?: string }>,
): ReviewNote[] {
  return notes.flatMap((note) => {
    const index = sourceRows.findIndex((row) => row.name === note.name);
    const edited = editedRows[index];
    if (!edited) return [];
    const label = note.dateLabel.match(/^(\d{1,2})\/(\d{1,2})$/);
    if (!label) return [];
    if (note.session === punchNoteSession) {
      const listed = (edited.punch ?? "")
        .split(/[;\n]+/)
        .some((part) => part.trim().startsWith(note.dateLabel));
      return listed ? [{ ...note, name: edited.name }] : [];
    }
    const text =
      note.session === "sáng" ? edited.morning : note.session === "chiều" ? edited.afternoon : edited.fullDay;
    const key = dateTokenKey(Number(label[2]), Number(label[1]));
    if (!listedDateKeys(text).has(key)) return [];
    return [{ ...note, name: edited.name }];
  });
}

export function leaveTotal(morning: string, afternoon: string, fullDay: string): number {
  return countListedDays(morning) * 0.5 + countListedDays(afternoon) * 0.5 + countListedDays(fullDay);
}

function dateKey(date: CalendarDate): number {
  return date.year * 10000 + date.month * 100 + date.day;
}

function isRealDate(date: CalendarDate): boolean {
  if (!Number.isInteger(date.year) || !Number.isInteger(date.month) || !Number.isInteger(date.day)) return false;
  if (date.year < 2000 || date.year > 2100) return false;
  const utc = new Date(Date.UTC(date.year, date.month - 1, date.day));
  return utc.getUTCFullYear() === date.year && utc.getUTCMonth() === date.month - 1 && utc.getUTCDate() === date.day;
}

export function toIsoDate(date: CalendarDate): string {
  return `${date.year}-${String(date.month).padStart(2, "0")}-${String(date.day).padStart(2, "0")}`;
}

export function parseIsoDate(value: string): CalendarDate | null {
  const match = value.match(/^(\d{4})-(\d{2})-(\d{2})$/);
  if (!match) return null;
  const date = { year: Number(match[1]), month: Number(match[2]), day: Number(match[3]) };
  return isRealDate(date) ? date : null;
}

export function normalizeExcludedSpan(start: CalendarDate, end: CalendarDate): ExcludedSpan | null {
  if (!isRealDate(start) || !isRealDate(end)) return null;
  if (dateKey(end) < dateKey(start)) return { start: { ...end }, end: { ...start } };
  return { start: { ...start }, end: { ...end } };
}

export function isExcludedDate(date: CalendarDate, spans: ExcludedSpan[]): boolean {
  const key = dateKey(date);
  return spans.some((span) => key >= dateKey(span.start) && key <= dateKey(span.end));
}

function weekdayOf(date: CalendarDate): number {
  return new Date(Date.UTC(date.year, date.month - 1, date.day)).getUTCDay();
}

function addDays(date: CalendarDate, days: number): CalendarDate {
  const next = new Date(Date.UTC(date.year, date.month - 1, date.day + days));
  return {
    year: next.getUTCFullYear(),
    month: next.getUTCMonth() + 1,
    day: next.getUTCDate(),
  };
}

function excelTime(value: number): string | null {
  if (!Number.isFinite(value) || value < 0 || value >= 1) return null;
  const minutes = Math.round(value * 24 * 60);
  return formatMinutes(minutes);
}

function acceptPunch(time: string): boolean {
  const minutes = toMinutes(time);
  return minutes >= PUNCH_FROM && minutes < PUNCH_BEFORE;
}

function parsePunches(value: unknown): string[] {
  if (typeof value === "number") {
    const time = excelTime(value);
    return time ? [time] : [];
  }
  return cellText(value)
    .split(/[\s,;]+/)
    .map((part) => part.trim())
    .filter((part) => /^\d{1,2}:\d{2}$/.test(part))
    .map((part) => {
      const [hour, minute] = part.split(":");
      return `${hour.padStart(2, "0")}:${minute}`;
    });
}

function toMinutes(hhmm: string): number {
  const [hour, minute] = hhmm.split(":").map(Number);
  return hour * 60 + minute;
}

function findPeriod(rows: unknown[][]): { label: string; start: CalendarDate; end: CalendarDate } | null {
  for (const row of rows) {
    for (const cell of row) {
      const text = cellText(cell);
      const match = text.match(
        /(\d{2})-(\d{2})-(\d{4})\s*~\s*(\d{2})-(\d{2})-(\d{4})/,
      );
      if (!match) continue;
      return {
        label: `${match[1]}/${match[2]}/${match[3]} – ${match[4]}/${match[5]}/${match[6]}`,
        start: { day: Number(match[1]), month: Number(match[2]), year: Number(match[3]) },
        end: { day: Number(match[4]), month: Number(match[5]), year: Number(match[6]) },
      };
    }
  }
  return null;
}

function inPeriod(date: CalendarDate, start: CalendarDate, end: CalendarDate): boolean {
  const key = dateKey(date);
  return key >= dateKey(start) && key <= dateKey(end);
}

function resolveDate(day: number, start: CalendarDate, end: CalendarDate): CalendarDate | null {
  let cursor = { ...start };
  const last = dateKey(end);
  while (dateKey(cursor) <= last) {
    if (cursor.day === day) return cursor;
    cursor = addDays(cursor, 1);
    if (dateKey(cursor) > last + 40) break;
  }
  return null;
}

function isWeekdayRow(row: unknown[]): boolean {
  const labels = row.filter((cell) => /^(T[2-7]|CN)$/.test(cellText(cell)));
  return labels.length >= 7;
}

function parseSheet(rows: unknown[][]): ParsedAttendance | null {
  const period = findPeriod(rows);
  if (!period) return null;

  const employees: ParsedEmployee[] = [];
  for (let index = 0; index < rows.length; index += 1) {
    const row = rows[index] ?? [];
    const nameIndex = row.findIndex((cell) => cellText(cell) === "Họ tên");
    if (nameIndex < 0) continue;

    let name = "";
    for (let column = nameIndex + 1; column < row.length; column += 1) {
      const text = cellText(row[column]);
      if (text) {
        name = text;
        break;
      }
    }
    if (!name) continue;

    const dayRow = rows[index + 1] ?? [];
    const weekdayRow = rows[index + 2] ?? [];
    const punchRow = rows[index + 3] ?? [];
    const dayCount = dayRow.filter((cell) => {
      const day = typeof cell === "number" ? cell : Number(cellText(cell));
      return Number.isInteger(day) && day >= 1 && day <= 31;
    }).length;
    if (dayCount < 7 || !isWeekdayRow(weekdayRow)) continue;

    const days: PunchDay[] = [];
    for (let column = 0; column < dayRow.length; column += 1) {
      const raw = dayRow[column];
      const day = typeof raw === "number" ? raw : Number(cellText(raw));
      if (!Number.isInteger(day) || day < 1 || day > 31) continue;
      const date = resolveDate(day, period.start, period.end);
      if (!date || !inPeriod(date, period.start, period.end)) continue;
      days.push({
        ...date,
        weekday: weekdayOf(date),
        punches: parsePunches(punchRow[column]),
      });
    }

    if (days.length > 0) employees.push({ name, days });
  }

  if (employees.length === 0) return null;
  return {
    periodLabel: period.label,
    start: period.start,
    end: period.end,
    employees,
  };
}

export function parseWorkbook(sheets: Sheet[]): ParsedAttendance {
  const preferred = sheets.find((sheet) => fold(sheet.name).includes("ban ghi"));
  const ordered = preferred ? [preferred, ...sheets.filter((sheet) => sheet !== preferred)] : sheets;
  for (const sheet of ordered) {
    const parsed = parseSheet(sheet.rows);
    if (parsed) return parsed;
  }
  throw new Error(
    "Không đọc được bảng chấm công. Hãy tải file xuất từ máy chấm công, gồm sheet Bản ghi chấm công.",
  );
}

type AbsenceKind = "morning" | "afternoon" | "full";

function sessionLabel(kind: AbsenceKind): string {
  if (kind === "morning") return "sáng";
  if (kind === "afternoon") return "chiều";
  return "nguyên ngày";
}

function sortedPunches(day: PunchDay): string[] {
  return [...day.punches].sort((left, right) => toMinutes(left) - toMinutes(right));
}

function describeLeave(day: PunchDay, kind: AbsenceKind, grace: number): string {
  const punches = sortedPunches(day);
  const valid = punches.filter(acceptPunch);
  const invalid = punches.filter((time) => !acceptPunch(time));
  if (punches.length === 0) return "không chấm công";
  if (valid.length <= 1 && invalid.length > 0) return `thời gian chấm không hợp lệ (${invalid.join(", ")})`;
  if (valid.length === 1) return `chỉ chấm 1 lần lúc ${valid[0]}`;

  const arrived = valid[0];
  const left = valid[valid.length - 1];
  const first = toMinutes(arrived);
  const last = toMinutes(left);
  const lateStart = first > MORNING_START + grace;
  const leftMorningEarly = last < MORNING_END - grace;
  const leftAfternoonEarly = last < AFTERNOON_END - grace;

  if (kind === "morning") {
    if (lateStart && leftMorningEarly) return `vào lúc ${arrived}, về lúc ${left}`;
    if (leftMorningEarly) return `chấm công về lúc ${left}`;
    if (lateStart) return `chấm công vào lúc ${arrived}`;
  }

  if (kind === "afternoon" && leftAfternoonEarly) return `chấm công về lúc ${left}`;
  return `vào lúc ${arrived}, về lúc ${left}`;
}

type DayAssessment =
  | { kind: "leave"; absence: AbsenceKind }
  | { kind: "punch"; detail: string };

function latePhrase(minutes: number, deadline: number): string {
  return `đến muộn ${minutes} phút so với ${formatMinutes(deadline)}`;
}

function earlyPhrase(minutes: number, deadline: number): string {
  return `về sớm ${minutes} phút so với ${formatMinutes(deadline)}`;
}

function invalidPhrase(times: string[]): string {
  return `chấm không hợp lệ (${times.join(", ")})`;
}

function assessDay(day: PunchDay, grace: number): DayAssessment | null {
  if (day.weekday === 0) return null;
  const punches = sortedPunches(day);
  const valid = punches.filter(acceptPunch);
  const invalid = punches.filter((time) => !acceptPunch(time));
  const saturday = day.weekday === 6;
  const inBy = MORNING_START + grace;
  const outFrom = (saturday ? MORNING_END : AFTERNOON_END) - grace;

  if (valid.length === 0) {
    if (invalid.length > 0) return { kind: "punch", detail: invalidPhrase(invalid) };
    return { kind: "leave", absence: saturday ? "morning" : "full" };
  }

  if (valid.length === 1) {
    const minutes = toMinutes(valid[0]);
    const forgot: string[] = [];
    if (minutes > inBy) forgot.push("quên chấm vào");
    if (minutes < outFrom) forgot.push("quên chấm ra");
    if (forgot.length === 0) forgot.push(minutes <= inBy ? "quên chấm ra" : "quên chấm vào");
    const parts = [`${forgot.join(", ")} (chấm lúc ${valid[0]})`];
    if (invalid.length > 0) parts.push(invalidPhrase(invalid));
    return { kind: "punch", detail: parts.join(", ") };
  }

  const first = toMinutes(valid[0]);
  const last = toMinutes(valid[valid.length - 1]);
  const morningAttended =
    first < MORNING_END && last >= MORNING_START && (saturday || last >= MORNING_END - grace || last >= AFTERNOON_START);
  const afternoonAttended = !saturday && first < AFTERNOON_END && last >= AFTERNOON_START;

  if (saturday) {
    if (!morningAttended) return { kind: "leave", absence: "morning" };
  } else if (!morningAttended && !afternoonAttended) {
    return { kind: "leave", absence: "full" };
  } else if (!morningAttended) {
    return { kind: "leave", absence: "morning" };
  } else if (!afternoonAttended) {
    return { kind: "leave", absence: "afternoon" };
  }

  const parts: string[] = [];
  if (first > inBy) parts.push(latePhrase(first - inBy, inBy));
  if (last < outFrom) parts.push(earlyPhrase(outFrom - last, outFrom));
  if (invalid.length > 0) parts.push(invalidPhrase(invalid));
  if (parts.length === 0) return null;
  return { kind: "punch", detail: parts.join(", ") };
}

function groupDates(dates: CalendarDate[]): string {
  if (dates.length === 0) return "";
  const sorted = [...dates].sort((left, right) => dateKey(left) - dateKey(right));
  const parts: string[] = [];
  let start = sorted[0];
  let previous = sorted[0];

  const push = () => {
    parts.push(dateKey(start) === dateKey(previous) ? formatDate(start) : `${formatDate(start)}-${formatDate(previous)}`);
  };

  for (const date of sorted.slice(1)) {
    const next = addDays(previous, 1);
    if (dateKey(next) === dateKey(date)) {
      previous = date;
      continue;
    }
    push();
    start = date;
    previous = date;
  }
  push();
  return parts.join(", ");
}

export function summarize(
  parsed: ParsedAttendance,
  graceMinutes: number,
  excluded: ExcludedSpan[] = [],
): AttendanceSummary {
  const grace = Number.isFinite(graceMinutes) ? Math.max(0, Math.round(graceMinutes)) : 0;
  const rows: SummaryRow[] = [];
  const notes: ReviewNote[] = [];

  for (const employee of parsed.employees) {
    const morning: CalendarDate[] = [];
    const afternoon: CalendarDate[] = [];
    const full: CalendarDate[] = [];
    const punchNotes: Array<{ date: CalendarDate; detail: string }> = [];

    for (const day of employee.days) {
      if (isExcludedDate(day, excluded)) continue;
      const assessed = assessDay(day, grace);
      if (!assessed) continue;
      if (assessed.kind === "leave") {
        if (assessed.absence === "morning") morning.push(day);
        if (assessed.absence === "afternoon") afternoon.push(day);
        if (assessed.absence === "full") full.push(day);
        notes.push({
          name: employee.name,
          dateLabel: formatDate(day),
          session: sessionLabel(assessed.absence),
          detail: describeLeave(day, assessed.absence, grace),
        });
        continue;
      }
      punchNotes.push({ date: day, detail: assessed.detail });
      notes.push({
        name: employee.name,
        dateLabel: formatDate(day),
        session: punchNoteSession,
        detail: assessed.detail,
      });
    }

    const total = morning.length * 0.5 + afternoon.length * 0.5 + full.length;
    rows.push({
      name: employee.name,
      morning: groupDates(morning),
      afternoon: groupDates(afternoon),
      fullDay: groupDates(full),
      punch: [...punchNotes]
        .sort((left, right) => dateKey(left.date) - dateKey(right.date))
        .map((item) => `${formatDate(item.date)} ${item.detail}`)
        .join("\n"),
      total,
    });
  }

  const month = parsed.start.month;
  return {
    title: `BẢNG CHẤM CÔNG T${month}`,
    periodLabel: parsed.periodLabel,
    rows,
    notes,
  };
}

export function graceWindow(graceMinutes: number): {
  morningIn: string;
  morningOut: string;
  afternoonIn: string;
  afternoonOut: string;
} {
  const grace = Number.isFinite(graceMinutes) ? Math.max(0, Math.round(graceMinutes)) : 0;
  return {
    morningIn: formatMinutes(MORNING_START + grace),
    morningOut: formatMinutes(MORNING_END - grace),
    afternoonIn: formatMinutes(AFTERNOON_START + grace),
    afternoonOut: formatMinutes(AFTERNOON_END - grace),
  };
}
