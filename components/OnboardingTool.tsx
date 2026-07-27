"use client";

import { useCallback, useEffect, useState } from "react";
import WeekCalendar, { addDays, formatWeekLabel, mondayOf } from "./WeekCalendar";
import type { AppliedRow, MeetingRow, Occurrence, ResolvedReference } from "@/lib/types";

/**
 * Vercel functions give us real headroom (see `maxDuration` on the routes), so
 * unlike the Val Town version this isn't a workaround for a 60s kill. It's kept
 * so the operator sees progress on a long run, and so one failing batch doesn't
 * lose the results of the batches that already succeeded.
 */
const APPLY_BATCH_SIZE = 10;

type Phase = "idle" | "previewing" | "previewed" | "applying" | "done";

/**
 * Single place that talks to our API. Bounces to /login on 401 and never
 * assumes the body is JSON — a proxy or platform error page would otherwise
 * surface to the user as an "Unexpected token '<'" parse error.
 */
async function postJson<T>(url: string, body: unknown): Promise<T> {
  const response = await fetch(url, {
    method: "POST",
    headers: { "Content-Type": "application/json" },
    body: JSON.stringify(body),
  });
  if (response.status === 401) {
    window.location.href = "/login";
    throw new Error("Session expired — redirecting to sign in.");
  }
  const text = await response.text();
  let data: unknown = null;
  try {
    data = text ? JSON.parse(text) : null;
  } catch {
    throw new Error(`Unexpected ${response.status} response from the server.`);
  }
  if (!response.ok) {
    throw new Error((data as { error?: string })?.error || `Request failed (${response.status})`);
  }
  return data as T;
}

function statusTag(row: AppliedRow): { className: string; label: string } {
  switch (row.status) {
    case "added":
      return { className: "done", label: "added" };
    case "forwarded":
      return { className: "done", label: "forwarded" };
    case "already-attendee":
      return { className: "neutral", label: "already on it" };
    case "skipped":
      return { className: "neutral", label: "skipped" };
    case "error":
      return { className: "err", label: "error" };
    default:
      return row.method === "direct"
        ? { className: "direct", label: "will add" }
        : { className: "neutral", label: "will forward" };
  }
}

export default function OnboardingTool({ presets }: { presets: string[] }) {
  const [newHireEmail, setNewHireEmail] = useState("");
  const [sourceMailbox, setSourceMailbox] = useState(presets[0] || "");
  const [referenceEmails, setReferenceEmails] = useState("");
  const [calendar, setCalendar] = useState("");
  const [windowDays, setWindowDays] = useState(90);
  const [minGroupMembers, setMinGroupMembers] = useState(2);
  const [references, setReferences] = useState<ResolvedReference[]>([]);
  const [attendeeType, setAttendeeType] = useState<"required" | "optional">("required");

  const [phase, setPhase] = useState<Phase>("idle");
  const [message, setMessage] = useState("");
  const [rows, setRows] = useState<AppliedRow[]>([]);
  const [selected, setSelected] = useState<Set<string>>(new Set());

  // Frozen at preview time so editing the form doesn't desync what we apply to.
  const [appliedContext, setAppliedContext] = useState({ newHireEmail: "", sourceMailbox: "", calendar: "" });

  const [view, setView] = useState<"list" | "calendar">("list");
  const [weekStart, setWeekStart] = useState(() => mondayOf(new Date()));
  const [occurrences, setOccurrences] = useState<Occurrence[]>([]);
  const [calLoading, setCalLoading] = useState(false);
  const [calError, setCalError] = useState("");

  const selectable = rows.filter((row) => row.status === "ready");
  const isBusy = phase === "previewing" || phase === "applying";

  const fetchWeek = useCallback(async () => {
    if (!appliedContext.sourceMailbox || rows.length === 0) {
      setOccurrences([]);
      return;
    }
    setCalLoading(true);
    setCalError("");
    try {
      const data = await postJson<{ occurrences: Occurrence[] }>("/api/calendar-week", {
        sourceMailbox: appliedContext.sourceMailbox,
        calendar: appliedContext.calendar,
        weekStart: weekStart.toISOString(),
        weekEnd: addDays(weekStart, 7).toISOString(),
      });
      setOccurrences(data.occurrences || []);
    } catch (error) {
      setOccurrences([]);
      setCalError(error instanceof Error ? error.message : String(error));
    } finally {
      setCalLoading(false);
    }
  }, [appliedContext.sourceMailbox, appliedContext.calendar, weekStart, rows.length]);

  useEffect(() => {
    if (view === "calendar") void fetchWeek();
  }, [view, fetchWeek]);

  async function runPreview() {
    if (!newHireEmail.trim()) {
      setMessage("Enter the new hire's email.");
      return;
    }
    if (!sourceMailbox.trim()) {
      setMessage("Enter the mailbox whose calendar you want to mirror.");
      return;
    }
    setPhase("previewing");
    setMessage("Scanning calendars…");
    setRows([]);
    setOccurrences([]);
    try {
      const data = await postJson<{
        rows: MeetingRow[];
        references: ResolvedReference[];
        sourceKind: "mailbox" | "group";
        sourceLabel: string;
        calendarsScanned: number;
        occurrencesScanned: number;
      }>("/api/preview", {
        newHireEmail: newHireEmail.trim(),
        sourceMailbox: sourceMailbox.trim(),
        referenceEmails: referenceEmails
          .split(",")
          .map((entry) => entry.trim())
          .filter(Boolean),
        calendar: calendar.trim(),
        windowDays,
        minGroupMembers,
      });

      const results: AppliedRow[] = data.rows || [];
      setReferences(data.references || []);
      setRows(results);
      setSelected(new Set(results.filter((row) => row.status === "ready").map((row) => row.key)));
      setAppliedContext({
        newHireEmail: newHireEmail.trim(),
        sourceMailbox: sourceMailbox.trim(),
        calendar: calendar.trim(),
      });

      const ready = results.filter((row) => row.status === "ready");
      const direct = ready.filter((row) => row.method === "direct").length;
      const sourceDesc =
        data.sourceKind === "group"
          ? `the Microsoft 365 group calendar for ${data.sourceLabel}`
          : `${data.calendarsScanned} calendar(s) in ${data.sourceLabel}`;
      setMessage(
        `Scanned ${sourceDesc}, ${data.occurrencesScanned} occurrence(s) over ${windowDays} days → ${results.length} meeting(s). ` +
          `${direct} can be added as a real attendee, ${ready.length - direct} can only be forwarded.`,
      );
      setPhase("previewed");
    } catch (error) {
      setMessage(`Failed: ${error instanceof Error ? error.message : String(error)}`);
      setPhase("idle");
    }
  }

  async function applySelected() {
    const items = rows.filter((row) => row.status === "ready" && selected.has(row.key));
    if (items.length === 0) return;

    setPhase("applying");
    const outcomes = new Map<string, AppliedRow>();

    for (let start = 0; start < items.length; start += APPLY_BATCH_SIZE) {
      const batch = items.slice(start, start + APPLY_BATCH_SIZE);
      setMessage(
        `Processing ${Math.min(start + batch.length, items.length)} of ${items.length} meeting(s)…`,
      );
      try {
        const data = await postJson<{ results: AppliedRow[] }>("/api/apply", {
          newHireEmail: appliedContext.newHireEmail,
          sourceMailbox: appliedContext.sourceMailbox,
          attendeeType,
          items: batch.map((row) => ({
            key: row.key,
            sourceEventId: row.sourceEventId,
            sourceCalendarId: row.sourceCalendarId,
            subject: row.subject,
          })),
        });
        for (const result of data.results) outcomes.set(result.key, result);
      } catch (error) {
        const detail = error instanceof Error ? error.message : String(error);
        for (const row of batch) outcomes.set(row.key, { ...row, status: "error", detail });
      }
    }

    const updated = rows.map<AppliedRow>((row) => {
      const outcome = outcomes.get(row.key);
      if (outcome) return outcome;
      if (row.status === "ready") return { ...row, status: "skipped" };
      return row;
    });
    setRows(updated);

    const added = updated.filter((row) => row.status === "added").length;
    const forwarded = updated.filter((row) => row.status === "forwarded").length;
    const failed = updated.filter((row) => row.status === "error").length;
    setMessage(
      `Added ${added} as a real attendee, forwarded ${forwarded}${failed ? `, ${failed} failed` : ""}. ` +
        `Run a new preview to retry failures or onboard someone else.`,
    );
    setPhase("done");
  }

  function toggle(key: string) {
    const row = rows.find((candidate) => candidate.key === key);
    if (!row || row.status !== "ready") return;
    setSelected((current) => {
      const next = new Set(current);
      if (next.has(key)) next.delete(key);
      else next.add(key);
      return next;
    });
  }

  function reset() {
    setRows([]);
    setReferences([]);
    setSelected(new Set());
    setOccurrences([]);
    setCalError("");
    setView("list");
    setPhase("idle");
    setMessage("Cleared. Run a new preview when ready.");
  }

  const forwardOnly = selectable.filter((row) => row.method === "forward").length;
  const selectedCount = selectable.filter((row) => selected.has(row.key)).length;

  return (
    <section className="panel">
      <div className="panel-intro">
        <h2>Mirror someone&rsquo;s meetings onto a new hire</h2>
        <p>
          Reads the chosen mailbox&rsquo;s calendars directly and, for every meeting whose organizer
          is inside your tenant, edits the <em>organizer&rsquo;s own copy</em> so the new hire becomes
          a genuine attendee on the real meeting — not a forwarded invite.
        </p>
        <p className="note">
          Adding an attendee through Graph makes Exchange send a meeting update to{" "}
          <strong>everyone</strong> already on the series — there is no &ldquo;notify added attendees
          only&rdquo; option in the API. Expect the team to get an updated invite per meeting.
        </p>
      </div>

      <div className="step">
        <div className="step-rail">
          <span className="step-node">1</span>
          <span className="step-line" />
        </div>
        <div className="step-body">
          <p className="eyebrow">Step 1</p>
          <p className="step-title">Who&rsquo;s joining</p>
          <div className="field">
            <label htmlFor="newhire">New hire email</label>
            <input
              id="newhire"
              type="email"
              autoComplete="off"
              placeholder="alice@coverdash.com"
              value={newHireEmail}
              onChange={(event) => setNewHireEmail(event.target.value)}
            />
          </div>
        </div>
      </div>

      <div className="step">
        <div className="step-rail">
          <span className="step-node">2</span>
          <span className="step-line" />
        </div>
        <div className="step-body">
          <p className="eyebrow">Step 2</p>
          <p className="step-title">What to mirror</p>

          <div className="field">
            <label htmlFor="source">Mailbox to mirror</label>
            <input
              id="source"
              type="email"
              list="preset-sources"
              placeholder="veronica.r@coverdash.com"
              value={sourceMailbox}
              onChange={(event) => setSourceMailbox(event.target.value)}
            />
            <datalist id="preset-sources">
              {presets.map((preset) => (
                <option key={preset} value={preset} />
              ))}
            </datalist>
            <p className="hint">
              Whose calendar we read. Because this runs on an application token, it does not have to
              be your mailbox and you do not have to be on the meetings. A <strong>person</strong>,
              a shared or room mailbox, or a <strong>Microsoft 365 group</strong> all work.
              A <strong>distribution list will not</strong> — a DL is a routing rule with no mailbox
              behind it, so there is no calendar to read. For a DL, mirror someone who is on the
              meetings and put the DL in the field below instead.
            </p>
          </div>

          <div className="field">
            <label htmlFor="references">Only meetings involving (optional)</label>
            <input
              id="references"
              type="text"
              placeholder="growth@coverdash.com, veronica.r@coverdash.com"
              value={referenceEmails}
              onChange={(event) => setReferenceEmails(event.target.value)}
            />
            <p className="hint">
              Comma-separated. Narrows the scan to meetings involving one of these — useful for
              picking out one team&rsquo;s meetings. Leave blank for every meeting on that calendar.
              Group addresses (Microsoft 365 groups and distribution lists alike) are resolved to
              their members, so a meeting still matches when the organizer expanded the group into
              individuals before sending.
            </p>
          </div>

          {referenceEmails.trim() !== "" && (
            <div className="field">
              <label htmlFor="minmembers">Members needed to match a group</label>
              <input
                id="minmembers"
                type="number"
                min={1}
                max={20}
                value={minGroupMembers}
                onChange={(event) => setMinGroupMembers(Number(event.target.value) || 2)}
              />
              <p className="hint">
                Only applies when the group address itself isn&rsquo;t on the meeting and we&rsquo;re
                matching members instead — the group address being present always matches outright.
                The mirrored person doesn&rsquo;t count toward this, so <code>2</code> keeps 1:1s with
                a single teammate out while letting real team meetings through. Drop to{" "}
                <code>1</code> to match any meeting touching the group.
              </p>
            </div>
          )}

          <div className="field">
            <label htmlFor="calendar">Calendar name filter (optional)</label>
            <input
              id="calendar"
              type="text"
              placeholder="Leave blank to scan every calendar in the mailbox"
              value={calendar}
              onChange={(event) => setCalendar(event.target.value)}
            />
          </div>

          <div className="field">
            <label htmlFor="window">Look ahead (days)</label>
            <input
              id="window"
              type="number"
              min={7}
              max={365}
              value={windowDays}
              onChange={(event) => setWindowDays(Number(event.target.value) || 90)}
            />
            <p className="hint">
              A recurring series is picked up as long as it has an occurrence in this window, however
              long ago it started.
            </p>
          </div>

          <div className="field">
            <label htmlFor="attendeetype">Add as</label>
            <select
              id="attendeetype"
              value={attendeeType}
              onChange={(event) => setAttendeeType(event.target.value as "required" | "optional")}
            >
              <option value="required">Required</option>
              <option value="optional">Optional</option>
            </select>
            <p className="hint">
              Applies to direct adds. Meetings that fall back to forwarding always land as optional —
              that&rsquo;s Graph&rsquo;s behaviour, not a setting.
            </p>
          </div>

          <button className="button" onClick={runPreview} disabled={isBusy}>
            {phase === "previewing" ? "Scanning…" : "Preview"}
          </button>
        </div>
      </div>

      <div className="step">
        <div className="step-rail">
          <span className="step-node">3</span>
        </div>
        <div className="step-body">
          <p className="eyebrow">Step 3</p>
          <p className="step-title">Review &amp; confirm</p>

          <p className="status">{message}</p>

          {references.length > 0 && (
            <ul className="results" style={{ marginBottom: 16 }}>
              {references.map((reference) => (
                <li className="result" key={reference.input}>
                  <div className="pickrow">
                    <span className="subject">{reference.input}</span>
                    <span className={`tag ${reference.kind === "group" ? "direct" : "neutral"}`}>
                      {reference.kind === "group" ? reference.groupType : "plain address"}
                    </span>
                  </div>
                  <span className="muted meta">
                    {reference.kind === "group"
                      ? `${reference.displayName || "group"} · ${reference.memberAddresses.length} member(s) · ${reference.addresses.length} address(es)`
                      : "not found in the directory — matched as a literal address"}
                    {reference.warning ? ` · ${reference.warning}` : ""}
                  </span>
                </li>
              ))}
            </ul>
          )}

          {rows.length > 0 && (
            <>
              <div className="viewtoggle">
                <button
                  type="button"
                  className={`viewtoggle-btn ${view === "list" ? "active" : ""}`}
                  onClick={() => setView("list")}
                >
                  List
                </button>
                <button
                  type="button"
                  className={`viewtoggle-btn ${view === "calendar" ? "active" : ""}`}
                  onClick={() => setView("calendar")}
                >
                  Calendar
                </button>
              </div>
              <p className="hint" style={{ marginBottom: 12 }}>
                In calendar view, click a meeting to toggle it — same as the checkboxes in list view.
              </p>
            </>
          )}

          {view === "list" ? (
            <ul className="results">
              {rows.map((row) => {
                const tag = statusTag(row);
                const isSelectable = row.status === "ready";
                const isPicked = selected.has(row.key);
                return (
                  <li
                    key={row.key}
                    className={`result ${isSelectable && !isPicked ? "unpicked" : ""}`}
                  >
                    <label className="pickrow">
                      {isSelectable && (
                        <input
                          type="checkbox"
                          checked={isPicked}
                          onChange={() => toggle(row.key)}
                        />
                      )}
                      <span className="subject">{row.subject}</span>
                      <span className={`tag ${tag.className}`}>{tag.label}</span>
                    </label>
                    <span className="muted meta">
                      organizer: {row.organizerEmail || "?"} · {row.isRecurring ? "recurring" : "one-off"} ·
                      {row.matchedVia?.length ? ` via ${row.matchedVia.join(", ")} · ` : " "}
                      method: {row.method === "direct" ? "direct add to organizer's copy" : "forward"}
                      {row.methodReason ? ` (${row.methodReason})` : ""}
                      {row.detail ? ` · ${row.detail}` : ""}
                    </span>
                  </li>
                );
              })}
              {rows.length === 0 && phase === "previewed" && (
                <li className="result muted">No matching meetings found.</li>
              )}
            </ul>
          ) : (
            <>
              <div className="calnav">
                <button
                  type="button"
                  className="button secondary calnav-btn"
                  onClick={() => setWeekStart((current) => addDays(current, -7))}
                >
                  ←
                </button>
                <span className="calnav-label">{formatWeekLabel(weekStart)}</span>
                <button
                  type="button"
                  className="button secondary calnav-btn"
                  onClick={() => setWeekStart((current) => addDays(current, 7))}
                >
                  →
                </button>
                <button
                  type="button"
                  className="button secondary calnav-btn"
                  onClick={() => setWeekStart(mondayOf(new Date()))}
                >
                  Today
                </button>
              </div>
              <WeekCalendar
                rows={rows}
                occurrences={occurrences}
                weekStart={weekStart}
                selected={selected}
                onToggle={toggle}
                loading={calLoading}
                error={calError}
              />
            </>
          )}

          {phase === "previewed" && selectable.length > 0 && (
            <>
              {forwardOnly > 0 && (
                <div className="callout">
                  {forwardOnly} meeting(s) can only be forwarded because their organizer&rsquo;s
                  mailbox isn&rsquo;t reachable by this app. The new hire will get a{" "}
                  <code>Fwd:</code> invite they have to accept, and will land as optional. If any of
                  those organizers are internal, widen the Exchange Application Access Policy.
                </div>
              )}
              <div className="confirmbar">
                <label className="checkbox">
                  <input
                    type="checkbox"
                    checked={selectedCount === selectable.length && selectable.length > 0}
                    onChange={(event) =>
                      setSelected(
                        event.target.checked
                          ? new Set(selectable.map((row) => row.key))
                          : new Set(),
                      )
                    }
                  />
                  Select all
                </label>
                <button className="button go" onClick={applySelected} disabled={isBusy || selectedCount === 0}>
                  {selectedCount ? `Add ${selectedCount} selected` : "Add selected"}
                </button>
                <button className="button secondary" onClick={reset} disabled={isBusy}>
                  Cancel
                </button>
              </div>
            </>
          )}

          {phase === "done" && (
            <div className="confirmbar">
              <button className="button secondary" onClick={reset}>
                Start over
              </button>
            </div>
          )}
        </div>
      </div>
    </section>
  );
}
