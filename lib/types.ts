export type GraphAttendee = {
  emailAddress?: { address?: string; name?: string };
  type?: "required" | "optional" | "resource";
  status?: { response?: string; time?: string };
  [key: string]: unknown;
};

export type GraphCalendar = {
  id: string;
  name?: string;
};

export type GraphEvent = {
  id: string;
  iCalUId?: string;
  subject?: string;
  type?: "singleInstance" | "occurrence" | "exception" | "seriesMaster";
  seriesMasterId?: string;
  organizer?: { emailAddress?: { address?: string; name?: string } };
  attendees?: GraphAttendee[];
  start?: { dateTime?: string; timeZone?: string };
  end?: { dateTime?: string; timeZone?: string };
  isAllDay?: boolean;
  isCancelled?: boolean;
};

/**
 * How a given meeting will be (or was) acted on.
 *  - "direct"  : we can reach the organizer's mailbox, so we PATCH their copy
 *                and the new hire becomes a genuine attendee on the real
 *                meeting. This is the outcome the tool exists for.
 *  - "forward" : the organizer is outside the tenant (or their mailbox isn't
 *                covered by the Application Access Policy), so the best we can
 *                do is forward from the source mailbox — same weaker result the
 *                old delegated version had for every non-self-organized meeting.
 */
export type ApplyMethod = "direct" | "forward";

/**
 * A reference address after directory lookup. `addresses` are the group's own
 * addresses (primary + aliases) — one of those on an attendee list is a
 * definitive match. `memberAddresses` are the people inside it, used to catch
 * meetings where the organizer expanded the group before sending.
 */
export type ResolvedReference = {
  input: string;
  kind: "group" | "address";
  displayName?: string;
  groupType?: "microsoft365" | "distribution" | "mail-enabled-security" | "security";
  addresses: string[];
  memberAddresses: string[];
  warning?: string;
};

/**
 * A row's lifecycle in one union. "ready" is the only selectable state; the
 * rest are terminal, either straight out of the preview or after applying.
 */
export type RowStatus =
  | "ready"
  | "already-attendee"
  | "added"
  | "forwarded"
  | "skipped"
  | "error";

/**
 * One matched meeting, as shown in the preview list. Deliberately carries no
 * attendee data — the apply step re-reads the event from Graph rather than
 * trusting anything that round-tripped through the browser.
 */
export type MeetingRow = {
  /** Stable row key. Also the id the calendar view cross-references by. */
  key: string;
  /** Event id in the SOURCE mailbox (series master id for recurring). */
  sourceEventId: string;
  sourceCalendarId: string;
  subject: string;
  organizerEmail: string;
  organizerName: string;
  isRecurring: boolean;
  method: ApplyMethod;
  /** Why the method is "forward" rather than "direct". */
  methodReason?: string;
  /** Which reference(s) this meeting matched, and how. Empty when unfiltered. */
  matchedVia?: string[];
  status: RowStatus;
  detail?: string;
  /** One representative occurrence, for display only. */
  sampleStart?: string;
  sampleEnd?: string;
  isAllDay: boolean;
};

/** Same shape; the alias marks rows that have been through the apply step. */
export type AppliedRow = MeetingRow;

export type Occurrence = {
  /** Matches MeetingRow.sourceEventId so the client can join them. */
  sourceEventId: string;
  startIso: string;
  endIso: string;
  isAllDay: boolean;
};
