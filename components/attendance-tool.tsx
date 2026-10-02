"use client";

import { useMemo, useState, useSyncExternalStore } from "react";
import { DateField } from "@/components/date-field";
import type { ExportedAttendance } from "@/lib/export-summary";
import {
  formatDate,
  formatTotal,
  graceWindow,
  isExcludedDate,
  leaveTotal,
  normalizeExcludedSpan,
  notesMatchingRows,
  parseIsoDate,
  parseWorkbook,
  punchNoteSession,
  summarize,
  toIsoDate,
  type AttendanceSummary,
  type CalendarDate,
  type ExcludedSpan,
  type ParsedAttendance,
  type ParsedEmployee,
  type PunchDay,
  type ReviewNote,
} from "@/lib/attendance";

const graceKey = "amber-tool.grace-minutes";
const excludedKey = "amber-tool.excluded-spans";
const defaultGrace = 15;
const emptyExcluded: ExcludedSpan[] = [];

function clampGrace(value: number): number {
  if (!Number.isFinite(value)) return defaultGrace;
  return Math.min(180, Math.max(0, Math.round(value)));
}

function readGrace(): number {
  try {
    const raw = localStorage.getItem(graceKey);
    if (raw == null) return defaultGrace;
    return clampGrace(Number(raw));
  } catch {
    return defaultGrace;
  }
}

const graceListeners = new Set<() => void>();
let graceValue = defaultGrace;
let graceHydrated = false;

function hydrateGrace() {
  if (graceHydrated) return;
  graceHydrated = true;
  graceValue = readGrace();
}

function subscribeGrace(listener: () => void) {
  graceListeners.add(listener);
  return () => {
    graceListeners.delete(listener);
  };
}

function graceSnapshot() {
  hydrateGrace();
  return graceValue;
}

function setGraceValue(value: number) {
  graceValue = Math.max(0, value);
  graceHydrated = true;
  try {
    localStorage.setItem(graceKey, String(graceValue));
  } catch {
    // Keep the in-page value when storage is unavailable.
  }
  graceListeners.forEach((listener) => listener());
}

const excludedListeners = new Set<() => void>();
let excludedValue: ExcludedSpan[] = emptyExcluded;
let excludedHydrated = false;

function readExcluded(): ExcludedSpan[] {
  try {
    const raw = localStorage.getItem(excludedKey);
    if (!raw) return emptyExcluded;
    const parsed: unknown = JSON.parse(raw);
    if (!Array.isArray(parsed)) return emptyExcluded;
    const spans = parsed.flatMap((item) => {
      if (!item || typeof item !== "object") return [];
      const record = item as { start?: CalendarDate; end?: CalendarDate };
      if (!record.start || !record.end) return [];
      const span = normalizeExcludedSpan(record.start, record.end);
      return span ? [span] : [];
    });
    return spans.length > 0 ? sortSpans(spans) : emptyExcluded;
  } catch {
    return emptyExcluded;
  }
}

function spanKey(span: ExcludedSpan): string {
  return `${toIsoDate(span.start)}:${toIsoDate(span.end)}`;
}

function sortSpans(spans: ExcludedSpan[]): ExcludedSpan[] {
  return [...spans].sort((left, right) => spanKey(left).localeCompare(spanKey(right)));
}

function hydrateExcluded() {
  if (excludedHydrated) return;
  excludedHydrated = true;
  excludedValue = readExcluded();
}

function subscribeExcluded(listener: () => void) {
  excludedListeners.add(listener);
  return () => {
    excludedListeners.delete(listener);
  };
}

function excludedSnapshot() {
  hydrateExcluded();
  return excludedValue;
}

function commitExcluded(next: ExcludedSpan[]) {
  excludedValue = next.length > 0 ? sortSpans(next) : emptyExcluded;
  excludedHydrated = true;
  try {
    localStorage.setItem(excludedKey, JSON.stringify(excludedValue));
  } catch {
    // Keep the in-page value when storage is unavailable.
  }
  excludedListeners.forEach((listener) => listener());
}

function addExcludedSpan(start: CalendarDate, end: CalendarDate): "added" | "duplicate" | "invalid" {
  const span = normalizeExcludedSpan(start, end);
  if (!span) return "invalid";
  hydrateExcluded();
  const key = spanKey(span);
  const exists = excludedValue.some((item) => spanKey(item) === key);
  if (exists) return "duplicate";
  commitExcluded([...excludedValue, span]);
  return "added";
}

function removeExcludedSpan(index: number) {
  hydrateExcluded();
  commitExcluded(excludedValue.filter((_, itemIndex) => itemIndex !== index));
}

function spanLabel(span: ExcludedSpan): string {
  const start = `${formatDate(span.start)}/${span.start.year}`;
  const end = `${formatDate(span.end)}/${span.end.year}`;
  return start === end ? start : `${start} – ${end}`;
}

function spanTouchesPeriod(span: ExcludedSpan, start: CalendarDate, end: CalendarDate): boolean {
  return toIsoDate(span.end) >= toIsoDate(start) && toIsoDate(span.start) <= toIsoDate(end);
}

export function AttendanceTool() {
  const [parsed, setParsed] = useState<ParsedAttendance | null>(null);
  const [fileName, setFileName] = useState("");
  const grace = useSyncExternalStore(subscribeGrace, graceSnapshot, () => defaultGrace);
  const excluded = useSyncExternalStore(subscribeExcluded, excludedSnapshot, () => emptyExcluded);
  const [error, setError] = useState("");
  const [reading, setReading] = useState(false);
  const [exporting, setExporting] = useState(false);

  const summary = useMemo(
    () => (parsed ? summarize(parsed, grace, excluded) : null),
    [parsed, grace, excluded],
  );
  const window = graceWindow(grace);

  async function onFile(file: File | undefined) {
    if (!file) return;
    setReading(true);
    setError("");
    try {
      const XLSX = await import("xlsx");
      const workbook = XLSX.read(await file.arrayBuffer(), { type: "array" });
      const sheets = workbook.SheetNames.map((name) => ({
        name,
        rows: XLSX.utils.sheet_to_json(workbook.Sheets[name], {
          header: 1,
          raw: true,
          defval: null,
        }) as unknown[][],
      }));
      setParsed(parseWorkbook(sheets));
      setFileName(file.name);
    } catch (cause) {
      setParsed(null);
      setFileName("");
      setError(cause instanceof Error ? cause.message : "Không đọc được file chấm công.");
    } finally {
      setReading(false);
    }
  }

  async function onExport(current: AttendanceSummary & { attendance: ExportedAttendance[] }) {
    setExporting(true);
    setError("");
    try {
      const response = await fetch("/api/attendance-export", {
        method: "POST",
        headers: { "Content-Type": "application/json" },
        body: JSON.stringify(current),
      });
      if (!response.ok) throw new Error("Không xuất được file Excel.");
      const blob = await response.blob();
      const matched = response.headers.get("Content-Disposition")?.match(/filename="([^"]+)"/);
      const url = URL.createObjectURL(blob);
      const link = document.createElement("a");
      link.href = url;
      link.download = matched?.[1] || "Bang-cham-cong.xlsx";
      document.body.appendChild(link);
      link.click();
      link.remove();
      URL.revokeObjectURL(url);
    } catch (cause) {
      setError(cause instanceof Error ? cause.message : "Không xuất được file Excel.");
    } finally {
      setExporting(false);
    }
  }

  return (
    <main className="mx-auto flex w-full max-w-6xl flex-col gap-6 px-4 py-8 sm:px-6">
      <header className="flex flex-col gap-2">
        <p className="text-sm font-medium tracking-wide text-amber-700">Amber Tool</p>
        <h1 className="text-3xl font-semibold tracking-tight text-slate-900">Tổng hợp ngày nghỉ</h1>
        <p className="max-w-3xl text-base leading-7 text-slate-600">
          Tải file chấm công gốc. Bảng bên dưới gom các ngày hoặc khoảng ngày nghỉ của từng nhân viên.
          Thứ 7 chỉ tính buổi sáng, Chủ nhật không tính. Ngày lễ hoặc khoảng ngày loại trừ không có chấm công vẫn tính là đi làm.
        </p>
      </header>

      <section className="grid gap-4 rounded-2xl border border-slate-200 bg-white p-5 shadow-sm md:grid-cols-[1fr_220px]">
        <label className="flex cursor-pointer flex-col items-center justify-center gap-2 rounded-xl border border-dashed border-slate-300 bg-slate-50 px-4 py-8 text-center hover:border-amber-500 hover:bg-amber-50">
          <span className="text-sm font-medium text-slate-900">
            {reading ? "Đang đọc file..." : "Chọn file .xls hoặc .xlsx"}
          </span>
          <span className="text-sm text-slate-500">{fileName || "File xuất từ máy chấm công"}</span>
          <input
            className="sr-only"
            type="file"
            accept=".xls,.xlsx,application/vnd.ms-excel,application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            onChange={(event) => {
              const file = event.target.files?.[0];
              void onFile(file);
              event.target.value = "";
            }}
          />
        </label>

        <label className="flex flex-col gap-2">
          <span className="text-sm font-medium text-slate-900">Xông xênh (phút)</span>
          <input
            className="h-11 rounded-lg border border-slate-300 px-3 text-lg text-slate-900 outline-none focus:border-amber-500"
            type="number"
            min={0}
            max={180}
            value={grace}
            onChange={(event) => setGraceValue(Math.max(0, Number(event.target.value) || 0))}
          />
          <span className="text-xs leading-5 text-slate-500">
            Đến {window.morningIn} vẫn kịp giờ sáng, tan sáng từ {window.morningOut}. Đến {window.afternoonIn} kịp giờ chiều, tan chiều từ {window.afternoonOut}. Chỉ tính giờ vào từ 06:00 và giờ về trước 19:00.
          </span>
        </label>
      </section>

      <ExcludedDays period={parsed ? { start: parsed.start, end: parsed.end } : null} spans={excluded} />

      {error ? (
        <p className="rounded-xl border border-red-200 bg-red-50 px-4 py-3 text-sm text-red-700">{error}</p>
      ) : null}

      {summary && parsed ? (
        <Result
          summary={summary}
          employees={parsed.employees}
          excluded={excluded}
          exporting={exporting}
          onExport={(edited) => void onExport(edited)}
        />
      ) : null}
    </main>
  );
}

type DraftRow = {
  name: string;
  morning: string;
  afternoon: string;
  fullDay: string;
  punch: string;
  total: string;
};

const columns: Array<{ key: keyof DraftRow; label: string; align: "left" | "center"; bold: boolean; width: string }> = [
  { key: "name", label: "Họ và tên", align: "left", bold: true, width: "w-[15%]" },
  { key: "morning", label: "Sáng", align: "left", bold: false, width: "w-[14%]" },
  { key: "afternoon", label: "Chiều", align: "left", bold: false, width: "w-[14%]" },
  { key: "fullDay", label: "Nguyên ngày", align: "left", bold: false, width: "w-[18%]" },
  { key: "punch", label: "Quên / Không chấm công", align: "left", bold: false, width: "w-[29%]" },
  { key: "total", label: "Tổng ngày nghỉ", align: "center", bold: true, width: "w-[10%]" },
];

function toDraft(summary: AttendanceSummary): DraftRow[] {
  return summary.rows.map((row) => ({
    name: row.name,
    morning: row.morning,
    afternoon: row.afternoon,
    fullDay: row.fullDay,
    punch: row.punch,
    total: typeof row.total === "number" ? formatTotal(row.total) : row.total,
  }));
}

const weekdays = ["CN", "T2", "T3", "T4", "T5", "T6", "T7"];

function ExcludedDays({
  period,
  spans,
}: {
  period: { start: CalendarDate; end: CalendarDate } | null;
  spans: ExcludedSpan[];
}) {
  const [single, setSingle] = useState("");
  const [from, setFrom] = useState("");
  const [to, setTo] = useState("");
  const [formError, setFormError] = useState("");
  const min = period ? toIsoDate(period.start) : undefined;
  const max = period ? toIsoDate(period.end) : undefined;

  function withinPeriod(value: string): boolean {
    if (!min || !max) return true;
    return value >= min && value <= max;
  }

  function addSingle() {
    const date = parseIsoDate(single);
    if (!date) {
      setFormError("Chọn một ngày hợp lệ.");
      return;
    }
    if (!withinPeriod(single)) {
      setFormError("Ngày này nằm ngoài kỳ chấm công.");
      return;
    }
    const result = addExcludedSpan(date, date);
    if (result === "duplicate") {
      setFormError("Ngày này đã có trong danh sách loại trừ.");
      return;
    }
    setSingle("");
    setFormError("");
  }

  function addRange() {
    const start = parseIsoDate(from);
    const end = parseIsoDate(to);
    if (!start || !end) {
      setFormError("Chọn đủ ngày bắt đầu và kết thúc.");
      return;
    }
    if (!withinPeriod(from) || !withinPeriod(to)) {
      setFormError("Khoảng ngày nằm ngoài kỳ chấm công.");
      return;
    }
    const result = addExcludedSpan(start, end);
    if (result === "duplicate") {
      setFormError("Khoảng ngày này đã có trong danh sách loại trừ.");
      return;
    }
    setFrom("");
    setTo("");
    setFormError("");
  }

  return (
    <section className="rounded-2xl border border-slate-200 bg-white p-5 shadow-sm">
      <h2 className="text-base font-semibold text-slate-900">Ngày loại trừ</h2>
      <p className="mt-1 max-w-3xl text-sm leading-6 text-slate-500">
        Thêm ngày lễ hoặc khoảng ngày trong tuần không có chấm công. Những ngày này vẫn tính là đi làm, không đưa vào ngày nghỉ.
      </p>
      <div className="mt-4 grid gap-4 lg:grid-cols-2">
        <div className="flex flex-col gap-2 rounded-xl border border-slate-200 bg-slate-50 p-4">
          <DateField label="Một ngày" value={single} onChange={setSingle} min={min} max={max} />
          <button
            type="button"
            className="h-11 rounded-lg border border-slate-300 bg-white px-4 text-sm font-medium text-slate-900 hover:bg-slate-100"
            onClick={addSingle}
          >
            Thêm ngày
          </button>
        </div>
        <div className="flex flex-col gap-2 rounded-xl border border-slate-200 bg-slate-50 p-4">
          <div className="grid gap-2 sm:grid-cols-2">
            <DateField label="Từ ngày" value={from} onChange={setFrom} min={min} max={max} />
            <DateField label="Đến ngày" value={to} onChange={setTo} min={min} max={max} />
          </div>
          <button
            type="button"
            className="h-11 rounded-lg border border-slate-300 bg-white px-4 text-sm font-medium text-slate-900 hover:bg-slate-100"
            onClick={addRange}
          >
            Thêm khoảng ngày
          </button>
        </div>
      </div>
      {formError ? <p className="mt-3 text-sm text-red-700">{formError}</p> : null}
      {spans.length === 0 ? (
        <p className="mt-3 text-sm text-slate-500">Chưa có ngày loại trừ.</p>
      ) : (
        <ul className="mt-3 flex flex-wrap gap-2">
          {spans.map((span, index) => {
            const outside = period ? !spanTouchesPeriod(span, period.start, period.end) : false;
            const label = spanLabel(span);
            return (
              <li
                key={spanKey(span)}
                className={`inline-flex items-center gap-2 rounded-full border px-3 py-1 text-sm ${
                  outside
                    ? "border-slate-200 bg-slate-100 text-slate-500"
                    : "border-emerald-200 bg-emerald-50 text-emerald-950"
                }`}
              >
                <span>
                  {label}
                  {outside ? " · ngoài kỳ này" : ""}
                </span>
                <button
                  type="button"
                  aria-label={`Bỏ ${label}`}
                  className="text-base leading-none text-current hover:opacity-70"
                  onClick={() => {
                    removeExcludedSpan(index);
                    setFormError("");
                  }}
                >
                  ×
                </button>
              </li>
            );
          })}
        </ul>
      )}
    </section>
  );
}

function Result({
  summary,
  employees,
  excluded,
  exporting,
  onExport,
}: {
  summary: AttendanceSummary;
  employees: ParsedEmployee[];
  excluded: ExcludedSpan[];
  exporting: boolean;
  onExport: (summary: AttendanceSummary & { attendance: ExportedAttendance[] }) => void;
}) {
  const [rows, setRows] = useState<DraftRow[]>(() => toDraft(summary));
  const [openRows, setOpenRows] = useState<Record<number, boolean>>({});
  const [renderedSummary, setRenderedSummary] = useState(summary);

  if (summary !== renderedSummary) {
    setRenderedSummary(summary);
    setRows(toDraft(summary));
    setOpenRows({});
  }

  const notesByRow = useMemo(
    () =>
      summary.rows.map((sourceRow) =>
        notesMatchingRows(
          summary.notes.filter((note) => note.name === sourceRow.name),
          summary.rows,
          rows,
        ),
      ),
    [summary, rows],
  );
  const notes = notesByRow.flat();

  function updateRow(index: number, key: Exclude<keyof DraftRow, "total">, value: string) {
    setRows((current) =>
      current.map((row, rowIndex) => {
        if (rowIndex !== index) return row;
        const next = { ...row, [key]: value };
        next.total = formatTotal(leaveTotal(next.morning, next.afternoon, next.fullDay));
        return next;
      }),
    );
  }

  return (
    <section className="flex flex-col gap-4">
      <div className="flex flex-wrap items-end justify-between gap-3">
        <div>
          <h2 className="text-center text-xl font-bold tracking-wide text-slate-900 sm:text-left">{summary.title}</h2>
          <p className="text-sm text-slate-500">{summary.periodLabel}</p>
          <p className="mt-1 text-sm text-slate-500">
            Bấm vào ô để sửa. Tổng ngày nghỉ tính lại theo sáng và chiều 0,5 ngày, nguyên ngày 1 ngày. Cột Quên / Không chấm công không cộng vào tổng này. Mũi tên cạnh tên mở giờ chấm công và ghi chú.
          </p>
        </div>
        <button
          type="button"
          className="h-11 rounded-lg bg-slate-900 px-4 text-sm font-medium text-white hover:bg-slate-700 disabled:opacity-60"
          disabled={exporting}
          onClick={() =>
            onExport({
              ...summary,
              rows,
              notes,
              attendance: employees.map((employee, index) => ({
                name: rows[index]?.name || employee.name,
                days: employee.days.map((day) => ({
                  day: day.day,
                  weekday: day.weekday,
                  label: formatDate(day),
                  punches: [...day.punches].sort(),
                  excused: isExcludedDate(day, excluded),
                })),
              })),
            })
          }
        >
          {exporting ? "Đang xuất..." : "Xuất file .xlsx"}
        </button>
      </div>

      <div className="overflow-x-auto rounded-2xl border border-slate-200 bg-white shadow-sm">
        <table className="w-full min-w-275 table-fixed border-collapse text-sm">
          <thead>
            <tr className="bg-slate-100 text-slate-900">
              <th className="w-10 border border-slate-200 px-2 py-3" aria-label="Mở chi tiết" />
              {columns.map((column) => (
                <th key={column.key} className={`border border-slate-200 px-3 py-3 text-center font-semibold ${column.width}`}>
                  {column.label}
                </th>
              ))}
            </tr>
          </thead>
          <tbody>
            {rows.map((row, index) => {
              const open = Boolean(openRows[index]);
              return (
                <EmployeeRows
                  key={summary.rows[index]?.name ?? index}
                  row={row}
                  index={index}
                  open={open}
                  days={employees[index]?.days ?? []}
                  notes={notesByRow[index] ?? []}
                  excluded={excluded}
                  onToggle={() => setOpenRows((current) => ({ ...current, [index]: !current[index] }))}
                  onChange={(key, value) => updateRow(index, key, value)}
                />
              );
            })}
          </tbody>
        </table>
      </div>
    </section>
  );
}

function EmployeeRows({
  row,
  index,
  open,
  days,
  notes,
  excluded,
  onToggle,
  onChange,
}: {
  row: DraftRow;
  index: number;
  open: boolean;
  days: PunchDay[];
  notes: ReviewNote[];
  excluded: ExcludedSpan[];
  onToggle: () => void;
  onChange: (key: Exclude<keyof DraftRow, "total">, value: string) => void;
}) {
  const noteByDate = new Map(notes.map((note) => [note.dateLabel, note]));

  return (
    <>
      <tr>
        <td className="border border-slate-200 p-1 text-center align-middle">
          <button
            type="button"
            aria-expanded={open}
            aria-label={`${open ? "Thu gọn" : "Mở"} chấm công của ${row.name || `dòng ${index + 1}`}`}
            className="inline-flex size-8 items-center justify-center rounded-md text-slate-600 hover:bg-slate-100"
            onClick={onToggle}
          >
            {open ? "▾" : "▸"}
          </button>
        </td>
        {columns.map((column) => {
          if (column.key === "total") {
            return (
              <td key={column.key} className="border border-slate-200 p-1 align-top">
                <div
                  aria-label={`${column.label} của dòng ${index + 1}`}
                  className="min-h-8 px-2 py-2 text-center text-base font-semibold text-slate-900"
                >
                  {row.total}
                </div>
              </td>
            );
          }

          const key = column.key;
          const value = key === "punch" ? row[key].replace(/; /g, "\n") : row[key];
          return (
            <td key={key} className="border border-slate-200 p-1 align-top">
              <textarea
                aria-label={`${column.label} của dòng ${index + 1}`}
                value={value}
                rows={key === "punch" ? Math.max(1, value.split("\n").length) : 1}
                onChange={(event) => onChange(key, event.target.value)}
                className={`field-sizing-content min-h-8 w-full resize-none bg-transparent px-2 py-2 text-sm text-slate-900 outline-none focus:bg-amber-50 ${
                  column.bold ? "font-semibold" : ""
                }`}
              />
            </td>
          );
        })}
      </tr>
      {open ? (
        <tr>
          <td colSpan={columns.length + 1} className="border border-slate-200 bg-slate-50 p-0">
            <div className="px-4 py-3">
            <p className="mb-2 text-sm font-semibold text-slate-800">Chấm công trong tháng — {row.name}</p>
            <div className="overflow-x-auto rounded-lg border border-slate-200 bg-white">
              <table className="w-max border-collapse text-center text-xs text-slate-900">
                <thead>
                  <tr>
                    {days.map((day) => (
                      <th key={`day-${day.day}`} className="min-w-14 border border-slate-200 bg-slate-100 px-1 py-1 font-semibold">
                        {day.day}
                      </th>
                    ))}
                  </tr>
                  <tr>
                    {days.map((day) => (
                      <th key={`week-${day.day}`} className="border border-slate-200 px-1 py-1 font-normal text-slate-500">
                        {weekdays[day.weekday]}
                      </th>
                    ))}
                  </tr>
                </thead>
                <tbody>
                  <tr>
                    {days.map((day) => {
                      const note = noteByDate.get(formatDate(day));
                      const excused = isExcludedDate(day, excluded);
                      const punchIssue = note?.session === punchNoteSession;
                      const punches = [...day.punches].sort();
                      return (
                        <td
                          key={`punch-${day.day}`}
                          className={`border border-slate-200 px-1 py-1 align-top leading-4 ${excused ? "bg-emerald-50" : ""} ${
                            !excused && punchIssue ? "bg-sky-50" : ""
                          } ${!excused && note && !punchIssue ? "bg-amber-50" : ""}`}
                        >
                          {punches.map((punch, punchIndex) => (
                            <div key={`${punch}-${punchIndex}`}>{punch}</div>
                          ))}
                          {excused ? <div className="font-medium text-emerald-800">đi làm</div> : null}
                        </td>
                      );
                    })}
                  </tr>
                </tbody>
              </table>
            </div>
            <div className="mt-3 grid gap-4 lg:grid-cols-2">
              <div>
                <p className="text-sm font-semibold text-amber-950">Ghi chú ngày nghỉ</p>
                {notes.filter((note) => note.session !== punchNoteSession).length === 0 ? (
                  <p className="mt-1 text-sm text-slate-500">Không có ngày nghỉ.</p>
                ) : (
                  <ul className="mt-2 flex flex-col gap-1 text-sm text-amber-950">
                    {notes
                      .filter((note) => note.session !== punchNoteSession)
                      .map((note) => (
                        <li key={`${note.dateLabel}-${note.session}`}>
                          {note.session} {note.dateLabel} — {note.detail}
                        </li>
                      ))}
                  </ul>
                )}
              </div>
              <div>
                <p className="text-sm font-semibold text-sky-950">Quên / Không chấm công</p>
                {notes.filter((note) => note.session === punchNoteSession).length === 0 ? (
                  <p className="mt-1 text-sm text-slate-500">Không có ngày quên chấm hoặc lệch giờ.</p>
                ) : (
                  <ul className="mt-2 flex flex-col gap-1 text-sm text-sky-950">
                    {notes
                      .filter((note) => note.session === punchNoteSession)
                      .map((note) => (
                        <li key={`${note.dateLabel}-${note.session}`}>
                          {note.dateLabel} — {note.detail}
                        </li>
                      ))}
                  </ul>
                )}
              </div>
            </div>
            </div>
          </td>
        </tr>
      ) : null}
    </>
  );
}
