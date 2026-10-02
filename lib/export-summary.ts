import type { AttendanceSummary } from "@/lib/attendance";

const NAVY = "FF1F4E79";
const HEADER = "FFE8EEF7";
const NOTE = "FFFFF7E8";
const LINE = "FFBFCBDC";
const EXCUSED = "FFD1FAE5";
const PUNCH = "FFE0F2FE";
const LEAVE = "FFFEF3C7";
const WEEKEND = "FFF1F5F9";
const MUTED = "FF64748B";

const weekdays = ["CN", "T2", "T3", "T4", "T5", "T6", "T7"];
const summaryLabels = ["HỌ VÀ TÊN", "SÁNG", "CHIỀU", "NGUYÊN NGÀY", "QUÊN / KHÔNG CHẤM CÔNG", "TỔNG NGÀY NGHỈ"];
const groupWeights = [3, 3, 3, 4, 6, 2];

export type ExportedDay = {
  day: number;
  weekday: number;
  label: string;
  punches: string[];
  excused: boolean;
};

export type ExportedAttendance = {
  name: string;
  days: ExportedDay[];
};

type WorkbookLike = {
  xlsx: { writeBuffer: () => Promise<ArrayBuffer> };
};

type WorksheetLike = {
  mergeCells: (row: number, col: number, endRow: number, endCol: number) => void;
  getCell: (row: number, col: number) => CellLike;
  getRow: (row: number) => { height: number };
  getColumn: (col: number) => { width: number };
};

type CellLike = {
  value: string | number | null;
  font: object;
  fill: object;
  border: object;
  alignment: object;
  numFmt: string;
};

type RangeStyle = {
  bold?: boolean;
  size?: number;
  color?: string;
  fill?: string;
  horizontal?: "left" | "center";
  vertical?: "top" | "middle";
  wrap?: boolean;
};

const border = {
  top: { style: "thin" as const, color: { argb: LINE } },
  left: { style: "thin" as const, color: { argb: LINE } },
  bottom: { style: "thin" as const, color: { argb: LINE } },
  right: { style: "thin" as const, color: { argb: LINE } },
};

function totalCellValue(total: number | string): number | string {
  if (typeof total === "number") return total;
  const trimmed = total.trim();
  const numeric = trimmed.replace(",", ".");
  if (/^-?\d+(?:\.\d+)?$/.test(numeric)) {
    const value = Number(numeric);
    return Number.isFinite(value) ? value : trimmed;
  }
  return trimmed;
}

async function loadExcel() {
  const excelModule = await import("exceljs");
  return excelModule.default ?? excelModule;
}

function columnGroups(columnCount: number): Array<{ start: number; end: number }> {
  const sum = groupWeights.reduce((total, weight) => total + weight, 0);
  const sizes = groupWeights.map((weight) => Math.max(1, Math.floor((weight / sum) * columnCount)));
  let used = sizes.reduce((total, size) => total + size, 0);
  while (used < columnCount) {
    sizes[4] += 1;
    used += 1;
  }
  while (used > columnCount) {
    const index = sizes.findIndex((size) => size > 1);
    if (index < 0) break;
    sizes[index] -= 1;
    used -= 1;
  }

  let cursor = 1;
  return sizes.map((size) => {
    const start = cursor;
    cursor += size;
    return { start, end: cursor - 1 };
  });
}

function paint(sheet: WorksheetLike, row: number, start: number, end: number, value: string | number, style: RangeStyle) {
  for (let column = start; column <= end; column += 1) {
    const cell = sheet.getCell(row, column);
    if (column === start) cell.value = value;
    cell.font = {
      name: "Arial",
      bold: Boolean(style.bold),
      size: style.size ?? 11,
      color: { argb: style.color ?? "FF0F172A" },
    };
    cell.border = border;
    cell.alignment = {
      horizontal: style.horizontal ?? "left",
      vertical: style.vertical ?? "middle",
      wrapText: style.wrap ?? true,
    };
    if (style.fill) {
      cell.fill = { type: "pattern", pattern: "solid", fgColor: { argb: style.fill } };
    }
  }
  if (end > start) sheet.mergeCells(row, start, row, end);
}

function sectionTitle(sheet: WorksheetLike, row: number, columnCount: number, text: string, color: string) {
  if (columnCount > 1) sheet.mergeCells(row, 1, row, columnCount);
  const cell = sheet.getCell(row, 1);
  cell.value = text;
  cell.font = { name: "Arial", bold: true, size: 13, color: { argb: color } };
  cell.alignment = { horizontal: "left", vertical: "middle" };
  sheet.getRow(row).height = 22;
}

function punchText(day: ExportedDay): string {
  const punches = day.punches.filter(Boolean);
  if (day.excused) return punches.length > 0 ? `${punches.join("\n")}\nđi làm` : "đi làm";
  return punches.join("\n");
}

function dayFill(day: ExportedDay, noteSession: string | undefined): string | undefined {
  if (day.excused) return EXCUSED;
  if (noteSession === "quên chấm") return PUNCH;
  if (noteSession) return LEAVE;
  if (day.weekday === 0 || day.weekday === 6) return WEEKEND;
  return undefined;
}

function calendarDays(attendance: ExportedAttendance[]): ExportedDay[] {
  return attendance.reduce<ExportedDay[]>((longest, employee) => (employee.days.length > longest.length ? employee.days : longest), []);
}

function writeAttendance(
  sheet: WorksheetLike,
  startRow: number,
  attendance: ExportedAttendance[],
  notes: AttendanceSummary["notes"],
  columnCount: number,
): number {
  const columns = calendarDays(attendance);
  if (columns.length === 0) return startRow;

  let row = startRow;
  const noteByName = new Map<string, Map<string, string>>();
  for (const note of notes) {
    const byDate = noteByName.get(note.name) ?? new Map<string, string>();
    if (!byDate.has(note.dateLabel)) byDate.set(note.dateLabel, note.session);
    noteByName.set(note.name, byDate);
  }

  for (const employee of attendance) {
    const byLabel = new Map(employee.days.map((day) => [day.label, day]));
    for (let column = 1; column <= columnCount; column += 1) {
      const cell = sheet.getCell(row, column);
      if (column === 1) {
        cell.value = employee.name;
        cell.font = { name: "Arial", bold: true, size: 12, color: { argb: NAVY } };
        cell.alignment = { horizontal: "left", vertical: "middle" };
      }
      cell.fill = { type: "pattern", pattern: "solid", fgColor: { argb: HEADER } };
      cell.border = border;
    }
    if (columnCount > 1) sheet.mergeCells(row, 1, row, columnCount);
    sheet.getRow(row).height = 22;
    row += 1;

    const sessions = noteByName.get(employee.name);
    let punchLines = 1;
    columns.forEach((columnDay, index) => {
      const day = byLabel.get(columnDay.label) ?? { ...columnDay, punches: [], excused: false };
      const column = index + 1;
      const text = punchText(day);
      punchLines = Math.max(punchLines, text ? text.split("\n").length : 1);
      const fill = dayFill(day, sessions?.get(day.label));
      paint(sheet, row, column, column, day.day, {
        bold: true,
        size: 10,
        color: NAVY,
        fill: fill ?? HEADER,
        horizontal: "center",
      });
      paint(sheet, row + 1, column, column, weekdays[day.weekday] ?? "", {
        size: 9,
        color: MUTED,
        fill,
        horizontal: "center",
      });
      paint(sheet, row + 2, column, column, text, {
        size: 9,
        fill,
        horizontal: "center",
        vertical: "top",
      });
    });
    sheet.getRow(row + 2).height = Math.max(32, punchLines * 14);
    row += 4;
  }

  return row;
}

export async function summaryToBuffer(summary: AttendanceSummary, attendance: ExportedAttendance[] = []): Promise<ArrayBuffer> {
  const ExcelJS = await loadExcel();
  const workbook = new ExcelJS.Workbook();
  workbook.creator = "Amber Tool";
  const sheet = workbook.addWorksheet("Bang cham cong") as WorksheetLike;

  const dayCount = attendance.reduce((max, employee) => Math.max(max, employee.days.length), 0);
  const columnCount = Math.max(6, dayCount);
  const groups = columnGroups(columnCount);
  const classicWidths = dayCount < 6;

  sheet.mergeCells(1, 1, 1, columnCount);
  const title = sheet.getCell(1, 1);
  title.value = summary.title;
  title.font = { name: "Arial", bold: true, size: 16, color: { argb: NAVY } };
  title.alignment = { horizontal: "center", vertical: "middle" };
  sheet.getRow(1).height = 28;

  sheet.mergeCells(2, 1, 2, columnCount);
  const period = sheet.getCell(2, 1);
  period.value = summary.periodLabel;
  period.font = { name: "Arial", size: 11, color: { argb: "FF475569" } };
  period.alignment = { horizontal: "center" };

  sectionTitle(sheet, 4, columnCount, "Chấm công", NAVY);
  let row = writeAttendance(sheet, 5, attendance, summary.notes, columnCount);

  sectionTitle(sheet, row, columnCount, "Bảng tổng hợp", NAVY);
  row += 1;
  const headerRow = row;
  summaryLabels.forEach((label, index) => {
    const group = groups[index];
    paint(sheet, headerRow, group.start, group.end, label, {
      bold: true,
      color: NAVY,
      fill: HEADER,
      horizontal: "center",
    });
  });
  sheet.getRow(headerRow).height = 32;
  row += 1;

  summary.rows.forEach((item) => {
    const total = totalCellValue(item.total);
    const punch = String(item.punch ?? "").replace(/; /g, "\n");
    const values: Array<string | number> = [item.name, item.morning, item.afternoon, item.fullDay, punch, total];
    const lineCount = punch ? punch.split("\n").length : 1;
    sheet.getRow(row).height = Math.max(22, lineCount * 18);
    values.forEach((value, index) => {
      const group = groups[index];
      paint(sheet, row, group.start, group.end, value, {
        bold: index === 0 || index === 5,
        horizontal: index === 5 ? "center" : "left",
        vertical: index === 4 ? "top" : "middle",
      });
      if (index === 5 && typeof value === "number") {
        sheet.getCell(row, group.start).numFmt = Number.isInteger(value) ? "0" : "0.0";
      }
    });
    row += 1;
  });

  row += 1;
  sectionTitle(sheet, row, columnCount, "Ghi chú — ngày nghỉ và quên chấm công", "FF9A3412");
  row += 1;

  if (summary.notes.length === 0) {
    if (columnCount > 1) sheet.mergeCells(row, 1, row, columnCount);
    const empty = sheet.getCell(row, 1);
    empty.value = "Không có ngày nghỉ.";
    empty.font = { name: "Arial", size: 11, italic: true, color: { argb: MUTED } };
  } else {
    const noteGroups = [
      groups[0],
      groups[1],
      groups[2],
      { start: groups[3].start, end: groups[5].end },
    ];
    ["HỌ VÀ TÊN", "NGÀY", "BUỔI", "CHI TIẾT"].forEach((label, index) => {
      const group = noteGroups[index];
      paint(sheet, row, group.start, group.end, label, {
        bold: true,
        color: "FF9A3412",
        fill: NOTE,
        horizontal: "center",
      });
    });
    row += 1;
    summary.notes.forEach((note) => {
      const values = [note.name, note.dateLabel, note.session, note.detail];
      const detailLines = Math.max(1, Math.ceil(note.detail.length / 48));
      sheet.getRow(row).height = Math.max(20, detailLines * 16);
      values.forEach((value, index) => {
        const group = noteGroups[index];
        paint(sheet, row, group.start, group.end, value, {
          vertical: index === 3 ? "top" : "middle",
        });
      });
      row += 1;
    });
  }

  const columnWidth = classicWidths ? 0 : columnCount >= 24 ? 9 : columnCount >= 12 ? 12 : 16;
  const classic = [24, 24, 24, 32, 46, 16];
  for (let column = 1; column <= columnCount; column += 1) {
    sheet.getColumn(column).width = classicWidths ? (classic[column - 1] ?? 16) : columnWidth;
  }

  const written = await (workbook as WorkbookLike).xlsx.writeBuffer();
  return written;
}

export function summaryFileName(summary: AttendanceSummary): string {
  const year = summary.periodLabel.match(/(\d{4})\s*$/)?.[1] ?? "";
  const month = summary.title.match(/T(\d+)/)?.[1] ?? "";
  return `Bang-cham-cong-T${month}-${year}.xlsx`;
}
