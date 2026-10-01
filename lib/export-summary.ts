import type { AttendanceSummary } from "@/lib/attendance";

const NAVY = "FF1F4E79";
const HEADER = "FFE8EEF7";
const NOTE = "FFFFF7E8";
const LINE = "FFBFCBDC";

type WorkbookLike = {
  xlsx: { writeBuffer: () => Promise<ArrayBuffer> };
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

export async function summaryToBuffer(summary: AttendanceSummary, includeNotes = false): Promise<ArrayBuffer> {
  const ExcelJS = await loadExcel();
  const workbook = new ExcelJS.Workbook();
  workbook.creator = "Amber Tool";
  const sheet = workbook.addWorksheet("Tong hop nghi", {
    views: [{ state: "frozen", ySplit: 3 }],
  });

  const border = {
    top: { style: "thin" as const, color: { argb: LINE } },
    left: { style: "thin" as const, color: { argb: LINE } },
    bottom: { style: "thin" as const, color: { argb: LINE } },
    right: { style: "thin" as const, color: { argb: LINE } },
  };

  sheet.mergeCells("A1:E1");
  const title = sheet.getCell("A1");
  title.value = summary.title;
  title.font = { name: "Arial", bold: true, size: 16, color: { argb: NAVY } };
  title.alignment = { horizontal: "center", vertical: "middle" };
  sheet.getRow(1).height = 28;

  sheet.mergeCells("A2:E2");
  const period = sheet.getCell("A2");
  period.value = summary.periodLabel;
  period.font = { name: "Arial", size: 11, color: { argb: "FF475569" } };
  period.alignment = { horizontal: "center" };

  const header = sheet.getRow(3);
  ["HỌ VÀ TÊN", "SÁNG", "CHIỀU", "NGUYÊN NGÀY", "TỔNG NGÀY NGHỈ"].forEach((label, index) => {
    const cell = header.getCell(index + 1);
    cell.value = label;
    cell.font = { name: "Arial", bold: true, size: 11, color: { argb: NAVY } };
    cell.fill = { type: "pattern", pattern: "solid", fgColor: { argb: HEADER } };
    cell.alignment = { horizontal: "center", vertical: "middle", wrapText: true };
    cell.border = border;
  });
  header.height = 22;

  summary.rows.forEach((row, index) => {
    const excelRow = sheet.getRow(4 + index);
    const total = totalCellValue(row.total);
    const values: Array<string | number> = [row.name, row.morning, row.afternoon, row.fullDay, total];
    values.forEach((value, column) => {
      const cell = excelRow.getCell(column + 1);
      cell.value = value;
      cell.font = { name: "Arial", size: 11, bold: column === 4 };
      cell.border = border;
      cell.alignment = {
        horizontal: column === 4 ? "center" : "left",
        vertical: "middle",
        wrapText: true,
      };
      if (column === 4 && typeof value === "number") {
        cell.numFmt = Number.isInteger(value) ? "0" : "0.0";
      }
    });
  });

  sheet.getColumn(1).width = 24;
  sheet.getColumn(2).width = 28;
  sheet.getColumn(3).width = 28;
  sheet.getColumn(4).width = 36;
  sheet.getColumn(5).width = 18;

  if (includeNotes) {
    const noteTitleRow = summary.rows.length + 5;
    sheet.mergeCells(noteTitleRow, 1, noteTitleRow, 5);
    const noteTitle = sheet.getCell(noteTitleRow, 1);
    noteTitle.value = "Ghi chú — ngày nghỉ, có chấm công hoặc không";
    noteTitle.font = { name: "Arial", bold: true, size: 12, color: { argb: "FF9A3412" } };

    if (summary.notes.length === 0) {
      sheet.mergeCells(noteTitleRow + 1, 1, noteTitleRow + 1, 5);
      const empty = sheet.getCell(noteTitleRow + 1, 1);
      empty.value = "Không có ngày nghỉ.";
      empty.font = { name: "Arial", size: 11, italic: true, color: { argb: "FF64748B" } };
    } else {
      const noteHeader = sheet.getRow(noteTitleRow + 1);
      ["HỌ VÀ TÊN", "NGÀY", "BUỔI", "CHI TIẾT"].forEach((label, index) => {
        const cell = noteHeader.getCell(index + 1);
        cell.value = label;
        cell.font = { name: "Arial", bold: true, size: 11, color: { argb: "FF9A3412" } };
        cell.fill = { type: "pattern", pattern: "solid", fgColor: { argb: NOTE } };
        cell.border = border;
      });
      summary.notes.forEach((note, index) => {
        const noteRow = sheet.getRow(noteTitleRow + 2 + index);
        [note.name, note.dateLabel, note.session, note.detail].forEach((value, column) => {
          const cell = noteRow.getCell(column + 1);
          cell.value = value;
          cell.font = { name: "Arial", size: 11 };
          cell.border = border;
          cell.alignment = { vertical: "middle" };
        });
      });
    }
  }

  const written = await (workbook as WorkbookLike).xlsx.writeBuffer();
  return written;
}

export function summaryFileName(summary: AttendanceSummary): string {
  const year = summary.periodLabel.match(/(\d{4})\s*$/)?.[1] ?? "";
  const month = summary.title.match(/T(\d+)/)?.[1] ?? "";
  return `Bang-cham-cong-T${month}-${year}.xlsx`;
}

