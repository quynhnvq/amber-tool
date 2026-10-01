import { summaryFileName, summaryToBuffer } from "@/lib/export-summary";
import type { AttendanceSummary } from "@/lib/attendance";

function isSummary(value: unknown): value is AttendanceSummary {
  if (!value || typeof value !== "object") return false;
  const summary = value as AttendanceSummary;
  return typeof summary.title === "string" && typeof summary.periodLabel === "string" && Array.isArray(summary.rows) && Array.isArray(summary.notes);
}

export async function POST(request: Request) {
  const body: unknown = await request.json();
  if (!isSummary(body)) {
    return Response.json({ error: "Dữ liệu không hợp lệ." }, { status: 400 });
  }

  const includeNotes = Boolean((body as AttendanceSummary & { includeNotes?: boolean }).includeNotes);
  const buffer = await summaryToBuffer(body, includeNotes);
  return new Response(buffer, {
    headers: {
      "Content-Type": "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
      "Content-Disposition": `attachment; filename="${summaryFileName(body)}"`,
    },
  });
}
