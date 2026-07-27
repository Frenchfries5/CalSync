"use client";

import { useMemo } from "react";
import type { AppliedRow, Occurrence } from "@/lib/types";

/**
 * Week grid for the matched meetings. Occurrences come from Graph's
 * `calendarView` (which expands recurrence properly, including exceptions and
 * cancellations) and are joined back to the preview rows by sourceEventId, so a
 * click here toggles exactly the same selection as the list checkbox.
 */

const DAY_ORDER = [1, 2, 3, 4, 5, 6, 0];
const PIXELS_PER_MINUTE = 1.15;

export function mondayOf(date: Date): Date {
  const result = new Date(date.getFullYear(), date.getMonth(), date.getDate());
  const day = result.getDay();
  result.setDate(result.getDate() + (day === 0 ? -6 : 1 - day));
  result.setHours(0, 0, 0, 0);
  return result;
}

export function addDays(date: Date, days: number): Date {
  const result = new Date(date);
  result.setDate(result.getDate() + days);
  return result;
}

/** Graph returns naive UTC strings when no timezone preference is sent. */
function parseGraphDate(iso: string): Date | null {
  if (!iso) return null;
  const date = new Date(/[zZ]|[+-]\d\d:\d\d$/.test(iso) ? iso : `${iso}Z`);
  return Number.isNaN(date.getTime()) ? null : date;
}

function formatHour(hour: number): string {
  const period = hour >= 12 ? "PM" : "AM";
  const display = hour % 12 === 0 ? 12 : hour % 12;
  return `${display} ${period}`;
}

function formatTime(minutes: number): string {
  const hour = Math.floor(minutes / 60);
  const minute = ((minutes % 60) + 60) % 60;
  const period = hour >= 12 ? "PM" : "AM";
  const display = hour % 12 === 0 ? 12 : hour % 12;
  return `${display}:${String(minute).padStart(2, "0")} ${period}`;
}

export function formatWeekLabel(weekStart: Date): string {
  const end = addDays(weekStart, 6);
  return `${weekStart.toLocaleDateString(undefined, {
    month: "short",
    day: "numeric",
  })}–${end.toLocaleDateString(undefined, { month: "short", day: "numeric", year: "numeric" })}`;
}

type Placed = {
  row: AppliedRow;
  weekday: number;
  startMin: number;
  endMin: number;
  lane: number;
  laneCount: number;
};

/** Greedy earliest-available-lane packing so overlapping events sit side by side. */
function layoutDay(events: Omit<Placed, "lane" | "laneCount">[]): Placed[] {
  const sorted = [...events].sort((a, b) => a.startMin - b.startMin || a.endMin - b.endMin);
  const laneEnds: number[] = [];
  const placed: Omit<Placed, "laneCount">[] = [];
  for (const event of sorted) {
    let lane = laneEnds.findIndex((end) => end <= event.startMin);
    if (lane === -1) {
      lane = laneEnds.length;
      laneEnds.push(event.endMin);
    } else {
      laneEnds[lane] = event.endMin;
    }
    placed.push({ ...event, lane });
  }
  const laneCount = laneEnds.length || 1;
  return placed.map((event) => ({ ...event, laneCount }));
}

function visualClass(row: AppliedRow, isSelected: boolean): string {
  if (row.status === "error") return "err";
  if (row.status === "added" || row.status === "forwarded") return "done";
  if (row.status === "already-attendee" || row.status === "skipped") return "neutral";
  return isSelected ? "selected" : "";
}

export default function WeekCalendar({
  rows,
  occurrences,
  weekStart,
  selected,
  onToggle,
  loading,
  error,
}: {
  rows: AppliedRow[];
  occurrences: Occurrence[];
  weekStart: Date;
  selected: Set<string>;
  onToggle: (key: string) => void;
  loading: boolean;
  error: string;
}) {
  const byEventId = useMemo(() => {
    const map = new Map<string, AppliedRow>();
    for (const row of rows) map.set(row.sourceEventId, row);
    return map;
  }, [rows]);

  const { timed, allDay } = useMemo(() => {
    const timedEvents: Omit<Placed, "lane" | "laneCount">[] = [];
    const allDayEvents: { row: AppliedRow; weekday: number }[] = [];
    for (const occurrence of occurrences) {
      const row = byEventId.get(occurrence.sourceEventId);
      if (!row) continue;
      const start = parseGraphDate(occurrence.startIso);
      const end = parseGraphDate(occurrence.endIso);
      if (!start || !end) continue;
      if (occurrence.isAllDay) {
        allDayEvents.push({ row, weekday: start.getDay() });
        continue;
      }
      const startMin = start.getHours() * 60 + start.getMinutes();
      let endMin = end.getHours() * 60 + end.getMinutes();
      if (endMin <= startMin) endMin = startMin + 30;
      timedEvents.push({ row, weekday: start.getDay(), startMin, endMin });
    }
    return { timed: timedEvents, allDay: allDayEvents };
  }, [occurrences, byEventId]);

  if (loading) return <p className="muted small">Loading week…</p>;
  if (error) return <p className="muted small">Could not load this week: {error}</p>;
  if (!rows.length) return <p className="muted small">Run a preview first.</p>;
  if (!timed.length && !allDay.length) {
    return <p className="muted small">No matched meetings occur this week. Try another week.</p>;
  }

  const used = new Set<number>();
  timed.forEach((event) => used.add(event.weekday));
  allDay.forEach((event) => used.add(event.weekday));
  const days = DAY_ORDER.filter((day) => (day >= 1 && day <= 5) || used.has(day));

  const minStart = timed.length ? Math.min(...timed.map((e) => e.startMin)) : 8 * 60;
  const maxEnd = timed.length ? Math.max(...timed.map((e) => e.endMin)) : 17 * 60;
  let gridStart = Math.max(6 * 60, Math.min(minStart - 30, 8 * 60));
  let gridEnd = Math.min(21 * 60, Math.max(maxEnd + 30, 17 * 60));
  if (gridEnd - gridStart < 8 * 60) {
    const mid = (gridStart + gridEnd) / 2;
    gridStart = Math.max(6 * 60, mid - 4 * 60);
    gridEnd = Math.min(21 * 60, mid + 4 * 60);
  }
  const heightPx = Math.round((gridEnd - gridStart) * PIXELS_PER_MINUTE);

  const hourLabels: number[] = [];
  for (let hour = Math.ceil(gridStart / 60); hour <= Math.floor(gridEnd / 60); hour++) {
    hourLabels.push(hour);
  }

  return (
    <div className="calendar">
      <div className="cal-gutter">
        <div className="cal-gutter-head" />
        <div className="cal-gutter-body" style={{ height: heightPx }}>
          {hourLabels.map((hour) => (
            <span
              key={hour}
              className="cal-gutter-label"
              style={{ top: Math.round((hour * 60 - gridStart) * PIXELS_PER_MINUTE) }}
            >
              {formatHour(hour)}
            </span>
          ))}
        </div>
      </div>
      <div className="cal-days">
        {days.map((day) => {
          const date = addDays(weekStart, day === 0 ? 6 : day - 1);
          const dayTimed = layoutDay(timed.filter((event) => event.weekday === day));
          const dayAllDay = allDay.filter((event) => event.weekday === day);
          return (
            <div className="cal-day" key={day}>
              <div className="cal-day-head">
                {date.toLocaleDateString(undefined, { weekday: "short", day: "numeric" })}
              </div>
              <div className="cal-day-allday">
                {dayAllDay.map((event, index) => (
                  <span
                    key={`${event.row.key}-${index}`}
                    className={`cal-allday-chip tag ${visualClass(
                      event.row,
                      selected.has(event.row.key),
                    )}`}
                    onClick={() => onToggle(event.row.key)}
                  >
                    {event.row.subject}
                  </span>
                ))}
              </div>
              <div className="cal-day-body" style={{ height: heightPx }}>
                {dayTimed.map((event, index) => {
                  const top = Math.round((event.startMin - gridStart) * PIXELS_PER_MINUTE);
                  const height = Math.max(
                    16,
                    Math.round((event.endMin - event.startMin) * PIXELS_PER_MINUTE) - 2,
                  );
                  const width = 100 / event.laneCount;
                  return (
                    <div
                      key={`${event.row.key}-${index}`}
                      className={`cal-event ${visualClass(event.row, selected.has(event.row.key))}`}
                      style={{
                        top,
                        height,
                        left: `${event.lane * width}%`,
                        width: `calc(${width}% - 3px)`,
                      }}
                      title={event.row.subject}
                      onClick={() => onToggle(event.row.key)}
                    >
                      <span className="cal-event-title">{event.row.subject}</span>
                      <span className="cal-event-time">
                        {formatTime(event.startMin)}–{formatTime(event.endMin)}
                      </span>
                    </div>
                  );
                })}
              </div>
            </div>
          );
        })}
      </div>
    </div>
  );
}
