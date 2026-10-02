"use client";

import { format } from "date-fns";
import { useEffect, useId, useLayoutEffect, useRef, useState } from "react";
import { DayPicker, type Matcher } from "react-day-picker";
import { vi } from "react-day-picker/locale";
import "react-day-picker/style.css";
import { formatDate, parseIsoDate } from "@/lib/attendance";

type DateFieldProps = {
  label: string;
  value: string;
  onChange: (value: string) => void;
  min?: string;
  max?: string;
};

function isoToDate(value: string | undefined): Date | undefined {
  const parsed = value ? parseIsoDate(value) : null;
  if (!parsed) return undefined;
  return new Date(parsed.year, parsed.month - 1, parsed.day);
}

function dateToIso(date: Date): string {
  const month = String(date.getMonth() + 1).padStart(2, "0");
  const day = String(date.getDate()).padStart(2, "0");
  return `${date.getFullYear()}-${month}-${day}`;
}

function fieldLabel(value: string): string {
  const parsed = parseIsoDate(value);
  return parsed ? `${formatDate(parsed)}/${parsed.year}` : "Chọn ngày";
}

export function DateField({ label, value, onChange, min, max }: DateFieldProps) {
  const [open, setOpen] = useState(false);
  const rootRef = useRef<HTMLDivElement>(null);
  const popoverRef = useRef<HTMLDivElement>(null);
  const labelId = useId();
  const selected = isoToDate(value);
  const minDate = isoToDate(min);
  const maxDate = isoToDate(max);
  const now = new Date();
  const startMonth = minDate ? new Date(minDate.getFullYear(), minDate.getMonth(), 1) : new Date(now.getFullYear() - 5, 0, 1);
  const endMonth = maxDate ? new Date(maxDate.getFullYear(), maxDate.getMonth(), 1) : new Date(now.getFullYear() + 1, 11, 1);
  const disabled: Matcher[] = [];
  if (minDate) disabled.push({ before: minDate });
  if (maxDate) disabled.push({ after: maxDate });

  useLayoutEffect(() => {
    if (!open) return;
    const anchor = rootRef.current;
    const popover = popoverRef.current;
    if (!anchor || !popover) return;

    function place() {
      if (!anchor || !popover) return;
      const rect = anchor.getBoundingClientRect();
      const height = popover.offsetHeight;
      const width = popover.offsetWidth;
      const margin = 8;
      let top = rect.bottom + margin;
      if (top + height > window.innerHeight - margin) top = rect.top - height - margin;
      if (top < margin) top = Math.max(margin, window.innerHeight - height - margin);
      let left = rect.left;
      if (left + width > window.innerWidth - margin) left = window.innerWidth - width - margin;
      if (left < margin) left = margin;
      popover.style.top = `${top}px`;
      popover.style.left = `${left}px`;
    }

    place();
    window.addEventListener("resize", place);
    window.addEventListener("scroll", place, true);
    return () => {
      window.removeEventListener("resize", place);
      window.removeEventListener("scroll", place, true);
    };
  }, [open]);

  useEffect(() => {
    if (!open) return;
    function onPointerDown(event: MouseEvent) {
      if (!rootRef.current?.contains(event.target as Node)) setOpen(false);
    }
    function onKeyDown(event: KeyboardEvent) {
      if (event.key === "Escape") setOpen(false);
    }
    document.addEventListener("mousedown", onPointerDown);
    document.addEventListener("keydown", onKeyDown);
    return () => {
      document.removeEventListener("mousedown", onPointerDown);
      document.removeEventListener("keydown", onKeyDown);
    };
  }, [open]);

  return (
    <div className="flex flex-col gap-2 text-sm font-medium text-slate-900">
      <span id={labelId}>{label}</span>
      <div className="relative" ref={rootRef}>
        <button
          type="button"
          aria-labelledby={labelId}
          aria-expanded={open}
          aria-haspopup="dialog"
          className="flex h-11 w-full items-center justify-between rounded-lg border border-slate-300 bg-white px-3 text-left text-base font-normal outline-none focus:border-amber-500"
          onClick={() => setOpen((current) => !current)}
        >
          <span className={selected ? "text-slate-900" : "text-slate-400"}>{fieldLabel(value)}</span>
          <CalendarIcon />
        </button>
        {open ? (
          <div
            ref={popoverRef}
            className="fixed z-30 rounded-xl border border-slate-200 bg-white p-3 shadow-lg"
          >
            <DayPicker
              mode="single"
              locale={vi}
              captionLayout="dropdown"
              navLayout="around"
              showOutsideDays
              weekStartsOn={1}
              className="amber-calendar"
              startMonth={startMonth}
              endMonth={endMonth}
              defaultMonth={selected ?? minDate ?? now}
              selected={selected}
              disabled={disabled}
              formatters={{
                formatWeekdayName: (date) => format(date, "EEEEE", { locale: vi }),
              }}
              onSelect={(date) => {
                if (!date) return;
                onChange(dateToIso(date));
                setOpen(false);
              }}
            />
          </div>
        ) : null}
      </div>
    </div>
  );
}

function CalendarIcon() {
  return (
    <svg aria-hidden="true" viewBox="0 0 20 20" className="size-5 text-slate-500">
      <path
        fill="currentColor"
        d="M6 2.5a.75.75 0 0 1 .75.75V4h6.5v-.75a.75.75 0 0 1 1.5 0V4h.75A2.25 2.25 0 0 1 17.75 6.25v9.5A2.25 2.25 0 0 1 15.5 18h-11A2.25 2.25 0 0 1 2.25 15.75v-9.5A2.25 2.25 0 0 1 4.5 4h.75v-.75A.75.75 0 0 1 6 2.5Zm9.5 5.25h-11v8c0 .414.336.75.75.75h9.5a.75.75 0 0 0 .75-.75v-8Z"
      />
    </svg>
  );
}
