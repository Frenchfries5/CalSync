import { GraphError, graphAll, graphGet, graphPatch, graphPost, odataString } from "./graph";
import { classify, findGroup, resolveReferences } from "./groups";
import type {
  AppliedRow,
  GraphAttendee,
  GraphCalendar,
  GraphEvent,
  MeetingRow,
  Occurrence,
  ResolvedReference,
} from "./types";

/**
 * How far ahead we look for meetings. `calendarView` expands recurrence, so a
 * series that started two years ago still shows up as long as it has an
 * occurrence inside this window — which also means we never page through years
 * of dead history the way the original version did.
 */
export const DEFAULT_WINDOW_DAYS = 90;

const EVENT_SELECT =
  "id,iCalUId,subject,type,seriesMasterId,organizer,attendees,start,end,isAllDay,isCancelled";

function lower(value: string | undefined | null): string {
  return (value || "").trim().toLowerCase();
}

function attendeeSet(event: Pick<GraphEvent, "attendees">): Set<string> {
  const set = new Set<string>();
  for (const attendee of event.attendees || []) {
    const address = lower(attendee.emailAddress?.address);
    if (address) set.add(address);
  }
  return set;
}

/** Runs `worker` over `items` with bounded concurrency, preserving order. */
async function mapLimit<T, R>(
  items: T[],
  limit: number,
  worker: (item: T, index: number) => Promise<R>,
): Promise<R[]> {
  const results = new Array<R>(items.length);
  let cursor = 0;
  const runners = Array.from({ length: Math.min(limit, items.length) }, async () => {
    while (cursor < items.length) {
      const index = cursor++;
      results[index] = await worker(items[index], index);
    }
  });
  await Promise.all(runners);
  return results;
}

export async function listCalendars(
  mailbox: string,
  nameFilter?: string,
): Promise<GraphCalendar[]> {
  const calendars = await graphAll<GraphCalendar>(
    `/users/${encodeURIComponent(mailbox)}/calendars?$select=id,name&$top=50`,
  );
  const filter = lower(nameFilter);
  if (!filter) return calendars;

  const matches = calendars.filter((calendar) => lower(calendar.name).includes(filter));
  if (matches.length === 0) {
    const available = calendars.map((c) => c.name || "(unnamed)").join(", ");
    throw new Error(
      `No calendar in ${mailbox} has a name containing "${nameFilter}". Available: ${available}`,
    );
  }
  return matches;
}

/**
 * Where we read meetings from. A user mailbox can hold several calendars; a
 * Microsoft 365 group has exactly one, reached through a different Graph path
 * (`/groups/{id}` rather than `/users/{upn}`), which is why this is a union
 * rather than just a string.
 */
export type CalendarSource =
  | { kind: "mailbox"; address: string; calendars: GraphCalendar[] }
  | { kind: "group"; address: string; groupId: string; displayName: string };

const sourceCache = new Map<string, { value: CalendarSource; expiresAt: number }>();
const SOURCE_CACHE_TTL_MS = 5 * 60 * 1000;

/**
 * Works out what the operator typed into "mailbox to mirror".
 *
 * The important case to get right is the one that can't work: a classic
 * distribution list is a routing rule with no mailbox behind it, so it has no
 * calendar to read and never will. That deserves an explanation, not a 404 —
 * it's an easy and reasonable thing to try.
 */
export async function resolveSource(
  address: string,
  calendarNameFilter?: string,
): Promise<CalendarSource> {
  const key = `${lower(address)}::${lower(calendarNameFilter)}`;
  const cached = sourceCache.get(key);
  if (cached && cached.expiresAt > Date.now()) return cached.value;

  let source: CalendarSource;
  try {
    await graphGet<{ id: string }>(`/users/${encodeURIComponent(address)}?$select=id`);
    source = {
      kind: "mailbox",
      address,
      calendars: await listCalendars(address, calendarNameFilter),
    };
  } catch (error) {
    if (!(error instanceof GraphError) || (error.status !== 404 && error.status !== 400)) {
      // 403 in particular means the mailbox exists but the Application Access
      // Policy excludes it — saying "not found" there would send you hunting
      // for a typo that isn't the problem.
      throw error;
    }

    const group = await findGroup(lower(address));
    if (!group) {
      throw new Error(
        `${address} isn't a mailbox or a group this app can see. Check the address, and that the Exchange Application Access Policy covers it.`,
      );
    }

    const groupType = classify(group);
    if (groupType !== "microsoft365") {
      const label =
        groupType === "distribution" ? "a distribution list" : "a security group";
      throw new Error(
        `${address} is ${label}, which has no calendar of its own — only a Microsoft 365 group does. ` +
          `Mirror a person who is on the meetings instead, and put ${address} in "Only meetings involving" to narrow the scan to that team's meetings.`,
      );
    }

    source = {
      kind: "group",
      address,
      groupId: group.id,
      displayName: group.displayName || address,
    };
  }

  sourceCache.set(key, { value: source, expiresAt: Date.now() + SOURCE_CACHE_TTL_MS });
  return source;
}

/** Calendar ids to iterate for a source. Groups have a single implicit one. */
function sourceCalendarIds(source: CalendarSource): string[] {
  return source.kind === "group" ? [""] : source.calendars.map((c) => c.id).filter(Boolean);
}

function calendarViewPath(source: CalendarSource, calendarId: string): string {
  if (source.kind === "group") {
    return `/groups/${encodeURIComponent(source.groupId)}/calendarView`;
  }
  return `/users/${encodeURIComponent(source.address)}/calendars/${encodeURIComponent(calendarId)}/calendarView`;
}

/** Path to a single event on the source, for re-reading and forwarding. */
function sourceEventPath(source: CalendarSource, calendarId: string, eventId: string): string {
  if (source.kind === "group") {
    return `/groups/${encodeURIComponent(source.groupId)}/events/${encodeURIComponent(eventId)}`;
  }
  return `/users/${encodeURIComponent(source.address)}/calendars/${encodeURIComponent(
    calendarId,
  )}/events/${encodeURIComponent(eventId)}`;
}

async function calendarView(
  source: CalendarSource,
  calendarId: string,
  startIso: string,
  endIso: string,
): Promise<GraphEvent[]> {
  const params = new URLSearchParams({
    startDateTime: startIso,
    endDateTime: endIso,
    $select: EVENT_SELECT,
    $top: "100",
    $orderby: "start/dateTime",
  });
  return graphAll<GraphEvent>(`${calendarViewPath(source, calendarId)}?${params}`);
}

/**
 * Outcome of looking for the organizer's copy. The two failure modes have very
 * different fixes — widen the access policy vs. the meeting genuinely isn't
 * findable in their default calendar — so they must not be collapsed into one
 * "can't write to this mailbox" message.
 */
type OrganizerLookup =
  | { status: "found"; event: GraphEvent; patchPath: string }
  | { status: "unreachable"; reason: string }
  | { status: "not-found"; reason: string };

// Organizer calendar lists are reused heavily across a scan (one organizer
// typically owns several of the matched meetings) and barely change.
const organizerCalendarCache = new Map<string, { value: GraphCalendar[]; expiresAt: number }>();

async function organizerCalendars(email: string): Promise<GraphCalendar[]> {
  const key = lower(email);
  const cached = organizerCalendarCache.get(key);
  if (cached && cached.expiresAt > Date.now()) return cached.value;
  const calendars = await graphAll<GraphCalendar>(
    `/users/${encodeURIComponent(email)}/calendars?$select=id,name&$top=50`,
  );
  organizerCalendarCache.set(key, { value: calendars, expiresAt: Date.now() + 5 * 60 * 1000 });
  return calendars;
}

function pickMatch(events: GraphEvent[]): GraphEvent | undefined {
  // A recurring meeting returns its series master here, which is exactly what
  // we want to patch — editing the master covers all future occurrences.
  return events.find((e) => e.type === "seriesMaster") || events[0];
}

/**
 * Finds the ORGANIZER's own copy of a meeting.
 *
 * Exchange gives every mailbox its own event id for the same meeting, so the id
 * we read off the source calendar is useless against the organizer's mailbox.
 * `iCalUId` is the cross-mailbox identity, so we look their copy up by it.
 *
 * The UID passed in MUST come from a series master, not from an occurrence —
 * see masterICalUId(). Occurrences carry a per-instance UID that matches nothing
 * in the organizer's mailbox.
 */
async function findOrganizerCopy(
  organizerEmail: string,
  iCalUId: string,
): Promise<OrganizerLookup> {
  if (!organizerEmail) return { status: "not-found", reason: "the event has no organizer" };
  if (!iCalUId) {
    return { status: "not-found", reason: "the event has no iCalUId to match on" };
  }
  const query = `$filter=iCalUId eq ${odataString(iCalUId)}&$select=${EVENT_SELECT}&$top=5`;
  const userBase = `/users/${encodeURIComponent(organizerEmail)}`;

  try {
    // Fast path: /users/{id}/events covers only the DEFAULT calendar, but that's
    // where the overwhelming majority of meetings live, so try it in one call
    // before fanning out.
    const body = await graphGet<{ value?: GraphEvent[] }>(`${userBase}/events?${query}`);
    const match = pickMatch(body.value || []);
    if (match) {
      return { status: "found", event: match, patchPath: `${userBase}/events/${encodeURIComponent(match.id)}` };
    }

    // Slow path: the default calendar isn't the only one they can organise
    // from. Search the rest before concluding we can't reach the meeting.
    const calendars = await organizerCalendars(organizerEmail);
    for (const calendar of calendars) {
      if (!calendar.id) continue;
      const scoped = await graphGet<{ value?: GraphEvent[] }>(
        `${userBase}/calendars/${encodeURIComponent(calendar.id)}/events?${query}`,
      );
      const scopedMatch = pickMatch(scoped.value || []);
      if (scopedMatch) {
        return {
          status: "found",
          event: scopedMatch,
          patchPath: `${userBase}/calendars/${encodeURIComponent(
            calendar.id,
          )}/events/${encodeURIComponent(scopedMatch.id)}`,
        };
      }
    }

    return {
      status: "not-found",
      reason: `their mailbox is readable, but none of their ${calendars.length} calendar(s) hold an event matching this meeting's UID`,
    };
  } catch (error) {
    if (error instanceof GraphError && error.status === 403) {
      return {
        status: "unreachable",
        reason: `Graph returned 403 for their mailbox — the Exchange Application Access Policy doesn't cover ${organizerEmail}`,
      };
    }
    if (error instanceof GraphError && error.status === 404) {
      return {
        status: "unreachable",
        reason: `no mailbox found for ${organizerEmail} in this tenant (external organiser)`,
      };
    }
    throw error;
  }
}

/**
 * Locates the copy of a meeting we should patch.
 *
 * When the organizer is the very mailbox we're mirroring, the event already in
 * hand *is* the organizer's copy — no lookup needed, and looking anyway would
 * mean searching a mailbox for something we're holding. That case is common
 * (you mirror someone who runs several of their own team's meetings) and it
 * used to fail outright whenever they organised from a secondary calendar,
 * because the fallback search couldn't see past the default one.
 */
async function locateOrganizerCopy(
  source: CalendarSource,
  calendarId: string,
  event: GraphEvent,
): Promise<OrganizerLookup> {
  const organizerEmail = event.organizer?.emailAddress?.address || "";
  if (organizerEmail && lower(organizerEmail) === lower(source.address)) {
    return {
      status: "found",
      event,
      patchPath: sourceEventPath(source, calendarId, event.id),
    };
  }
  return findOrganizerCopy(organizerEmail, event.iCalUId || "");
}

/**
 * The UID to match the organizer's copy against.
 *
 * calendarView hands back OCCURRENCES, and an occurrence's iCalUId encodes the
 * instance — it is not the series master's UID and matches nothing in anyone
 * else's mailbox. Fetching the master costs one call per series and is the
 * difference between recurring meetings being added properly and silently
 * falling back to a forward.
 */
async function masterEvent(
  source: CalendarSource,
  calendarId: string,
  occurrence: GraphEvent,
): Promise<GraphEvent> {
  if (!occurrence.seriesMasterId) return occurrence;
  try {
    return await graphGet<GraphEvent>(
      `${sourceEventPath(source, calendarId, occurrence.seriesMasterId)}?$select=${EVENT_SELECT}`,
    );
  } catch {
    return occurrence; // better a degraded match than a failed row
  }
}

/**
 * How many of a group's members must be on a meeting for it to count as "that
 * group's meeting" when the group address itself isn't on the attendee list.
 *
 * Needed because member matching is much broader than address matching: with a
 * threshold of 1, a 1:1 between the mirrored person and a single teammate would
 * match the whole team's filter. The mirrored person is excluded from the count,
 * so a 1:1 contributes one member and a real team meeting contributes several.
 */
export const DEFAULT_MIN_GROUP_MEMBERS = 2;

export type ScanOptions = {
  /** Mailbox whose calendar we read (the person being mirrored). */
  sourceMailbox: string;
  /** Optional extra filter: only meetings involving one of these addresses. */
  referenceEmails: string[];
  /** The new hire — used only to mark meetings they're already on. */
  newHireEmail: string;
  calendarNameFilter?: string;
  windowDays?: number;
  minGroupMembers?: number;
};

export type ScanResult = {
  rows: MeetingRow[];
  calendarsScanned: number;
  occurrencesScanned: number;
  windowStart: string;
  windowEnd: string;
  references: ResolvedReference[];
  sourceKind: CalendarSource["kind"];
  sourceLabel: string;
};

/**
 * Decides whether one event matches one resolved reference, and says how.
 * Returns null for no match.
 */
export function matchReference(
  reference: ResolvedReference,
  attendees: Set<string>,
  organizer: string,
  sourceMailbox: string,
  minGroupMembers: number,
): string | null {
  const onEvent = (address: string) => attendees.has(address) || address === organizer;

  // The group's own address (or an alias) being present is unambiguous.
  if (reference.addresses.some(onEvent)) {
    return reference.kind === "group" ? `${reference.input} (group invited)` : reference.input;
  }

  if (reference.kind !== "group" || reference.memberAddresses.length === 0) return null;

  const hits = reference.memberAddresses.filter(
    (address) => address !== sourceMailbox && onEvent(address),
  );
  if (hits.length >= minGroupMembers) {
    // Count excludes the mirrored person, so this can read lower than the number
    // of group members visibly on the meeting. That's the number that mattered.
    return `${reference.input} (${hits.length} member${hits.length === 1 ? "" : "s"} matched)`;
  }
  return null;
}

export async function scanMailbox(options: ScanOptions): Promise<ScanResult> {
  const windowDays = options.windowDays ?? DEFAULT_WINDOW_DAYS;
  const windowStart = new Date();
  const windowEnd = new Date(windowStart.getTime() + windowDays * 86_400_000);
  const startIso = windowStart.toISOString();
  const endIso = windowEnd.toISOString();

  const source = await resolveSource(options.sourceMailbox, options.calendarNameFilter);
  const calendarIds = sourceCalendarIds(source);
  const newHire = lower(options.newHireEmail);
  const sourceMailbox = lower(options.sourceMailbox);
  const minGroupMembers = Math.max(1, options.minGroupMembers ?? DEFAULT_MIN_GROUP_MEMBERS);

  // Group addresses often don't survive onto the event (Outlook lets the
  // organizer expand a group into individuals before sending), so each
  // reference is resolved to its members up front and matched on either.
  const references = await resolveReferences(
    options.referenceEmails.map((entry) => entry.trim()).filter(Boolean),
  );

  // Dedupe occurrences down to one row per series (or per one-off), keeping the
  // earliest occurrence as the display sample.
  const byEvent = new Map<
    string,
    { event: GraphEvent; calendarId: string; matchedVia: string[] }
  >();
  let occurrencesScanned = 0;

  for (const calendarId of calendarIds) {
    const events = await calendarView(source, calendarId, startIso, endIso);
    for (const event of events) {
      occurrencesScanned++;
      if (event.isCancelled) continue;

      let matchedVia: string[] = [];
      if (references.length > 0) {
        const attendees = attendeeSet(event);
        const organizer = lower(event.organizer?.emailAddress?.address);
        matchedVia = references
          .map((reference) =>
            matchReference(reference, attendees, organizer, sourceMailbox, minGroupMembers),
          )
          .filter((label): label is string => label !== null);
        if (matchedVia.length === 0) continue;
      }

      const key = `${calendarId}::${event.seriesMasterId || event.id}`;
      if (!byEvent.has(key)) {
        byEvent.set(key, { event, calendarId, matchedVia });
      }
    }
  }

  const candidates = [...byEvent.entries()];

  const rows = await mapLimit<(typeof candidates)[number], MeetingRow>(
    candidates,
    6,
    async ([key, { event, calendarId, matchedVia }]) => {
      const organizerEmail = event.organizer?.emailAddress?.address || "";
      const organizerName = event.organizer?.emailAddress?.name || organizerEmail;
      const isRecurring = Boolean(event.seriesMasterId) || event.type === "seriesMaster";

      const base: MeetingRow = {
        key,
        sourceEventId: event.seriesMasterId || event.id,
        sourceCalendarId: calendarId,
        subject: event.subject || "(no subject)",
        organizerEmail,
        organizerName,
        isRecurring,
        method: "forward",
        status: "ready",
        matchedVia,
        sampleStart: event.start?.dateTime,
        sampleEnd: event.end?.dateTime,
        isAllDay: Boolean(event.isAllDay),
      };

      try {
        // Match on the series master's UID, never the occurrence's — see
        // masterEvent(). This is also the copy whose attendee list tells us
        // whether the new hire is already on the whole series.
        const master = await masterEvent(source, calendarId, event);
        const lookup = await locateOrganizerCopy(source, calendarId, master);

        if (lookup.status !== "found") {
          return {
            ...base,
            method: "forward",
            methodReason: lookup.reason,
            status: attendeeSet(master).has(newHire) ? "already-attendee" : "ready",
          };
        }

        return {
          ...base,
          method: "direct",
          status: attendeeSet(lookup.event).has(newHire) ? "already-attendee" : "ready",
        };
      } catch (error) {
        return {
          ...base,
          status: "error",
          detail: error instanceof Error ? error.message : String(error),
        };
      }
    },
  );

  rows.sort((a, b) => (a.sampleStart || "").localeCompare(b.sampleStart || ""));

  return {
    rows,
    references,
    sourceKind: source.kind,
    sourceLabel: source.kind === "group" ? source.displayName : source.address,
    calendarsScanned: calendarIds.length,
    occurrencesScanned,
    windowStart: startIso,
    windowEnd: endIso,
  };
}

/**
 * Applies one meeting. Everything is re-read from Graph here — the client only
 * hands back ids, never attendee data, so a tampered payload can't rewrite a
 * meeting's attendee list.
 */
export async function applyOne(
  row: Pick<MeetingRow, "key" | "sourceEventId" | "sourceCalendarId" | "subject">,
  sourceMailbox: string,
  newHireEmail: string,
  attendeeType: "required" | "optional",
): Promise<AppliedRow> {
  const calendarSource = await resolveSource(sourceMailbox);
  const eventPath = sourceEventPath(calendarSource, row.sourceCalendarId, row.sourceEventId);

  const source = await graphGet<GraphEvent>(`${eventPath}?$select=${EVENT_SELECT}`);
  const organizerEmail = source.organizer?.emailAddress?.address || "";
  const organizerName = source.organizer?.emailAddress?.name || organizerEmail;
  const newHire = newHireEmail.trim();

  const base: AppliedRow = {
    key: row.key,
    sourceEventId: row.sourceEventId,
    sourceCalendarId: row.sourceCalendarId,
    subject: source.subject || row.subject,
    organizerEmail,
    organizerName,
    isRecurring: source.type === "seriesMaster",
    method: "forward",
    status: "skipped",
    sampleStart: source.start?.dateTime,
    sampleEnd: source.end?.dateTime,
    isAllDay: Boolean(source.isAllDay),
  };

  // sourceEventId is already the series master's id, so source.iCalUId is the
  // master UID here — no occurrence lookup needed on this path.
  const lookup = await locateOrganizerCopy(calendarSource, row.sourceCalendarId, source);

  if (lookup.status === "found") {
    const organizerCopy = lookup.event;
    if (attendeeSet(organizerCopy).has(lower(newHire))) {
      return { ...base, method: "direct", status: "already-attendee" };
    }
    // PATCHing `attendees` REPLACES the whole list, so the existing attendees
    // have to be resent or they're dropped from the meeting. We send only
    // address + type; Exchange keeps each unchanged attendee's RSVP because it
    // matches them by address.
    const attendees: GraphAttendee[] = [
      ...(organizerCopy.attendees || []).map((attendee) => ({
        emailAddress: attendee.emailAddress,
        type: attendee.type || "required",
      })),
      { emailAddress: { address: newHire, name: newHire }, type: attendeeType },
    ];
    // patchPath came from wherever the copy was actually found — their default
    // calendar, a secondary one, or the source itself when they organise it.
    await graphPatch(lookup.patchPath, { attendees });
    return { ...base, method: "direct", status: "added" };
  }

  if (attendeeSet(source).has(lower(newHire))) {
    return { ...base, method: "forward", status: "already-attendee" };
  }
  await graphPost(`${eventPath}/forward`, {
    Comment: "Adding you to this meeting as part of onboarding.",
    ToRecipients: [{ EmailAddress: { Address: newHire, Name: newHire } }],
  });
  return {
    ...base,
    method: "forward",
    methodReason: lookup.reason,
    status: "forwarded",
  };
}

/** Backs the week grid in the UI. Timing only — the client joins on sourceEventId. */
export async function listWeekOccurrences(
  sourceMailbox: string,
  calendarNameFilter: string | undefined,
  startIso: string,
  endIso: string,
): Promise<Occurrence[]> {
  const source = await resolveSource(sourceMailbox, calendarNameFilter);
  const occurrences: Occurrence[] = [];
  for (const calendarId of sourceCalendarIds(source)) {
    const events = await calendarView(source, calendarId, startIso, endIso);
    for (const event of events) {
      if (event.isCancelled) continue;
      occurrences.push({
        sourceEventId: event.seriesMasterId || event.id,
        startIso: event.start?.dateTime || "",
        endIso: event.end?.dateTime || "",
        isAllDay: Boolean(event.isAllDay),
      });
    }
  }
  return occurrences;
}
