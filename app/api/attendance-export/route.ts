import { summaryFileName, summaryToBuffer, type ExportedAttendance, type ExportedDay } from "@/lib/export-summary";
import type { AttendanceSummary } from "@/lib/attendance";

function isDay(value: unknown): value is ExportedDay {
  if (!value || typeof value !== "object") return false;
  const day = value as ExportedDay;
  return (
    Number.isInteger(day.day) &&
    Number.isInteger(day.weekday) &&
    typeof day.label === "string" &&
    typeof day.excused === "boolean" &&
    Array.isArray(day.punches) &&
    day.punches.every((punch) => typeof punch === "string")
  );
}

function isAttendance(value: unknown): value is ExportedAttendance[] {
  return (
    Array.isArray(value) &&
    value.every((item) => {
      if (!item || typeof item !== "object") return false;
      const block = item as ExportedAttendance;
      return typeof block.name === "string" && Array.isArray(block.days) && block.days.every(isDay);
    })
  );
}

function isSummary(value: unknown): value is AttendanceSummary & { attendance: ExportedAttendance[] } {
  if (!value || typeof value !== "object") return false;
  const summary = value as AttendanceSummary & { attendance?: unknown };
  return (
    typeof summary.title === "string" &&
    typeof summary.periodLabel === "string" &&
    Array.isArray(summary.rows) &&
    Array.isArray(summary.notes) &&
    isAttendance(summary.attendance)
  );
}

export async function POST(request: Request) {
  const body: unknown = await request.json();
  if (!isSummary(body)) {
    return Response.json({ error: "Dữ liệu không hợp lệ." }, { status: 400 });
  }

  const buffer = await summaryToBuffer(body, body.attendance);
  return new Response(buffer, {
    headers: {
      "Content-Type": "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
      "Content-Disposition": `attachment; filename="${summaryFileName(body)}"`,
    },
  });
}
