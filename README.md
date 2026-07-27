# Onboarding → meetings

Adds a new hire to the meetings someone else is already on — recurring series and
upcoming one-offs — so you don't have to forward a dozen invites by hand.

Next.js on Vercel. This is a port of the Val Town val of the same purpose, with
one substantive behaviour change: **it uses application (app-only) Graph
permissions instead of a delegated token.**

## Why that change matters

The Val Town version signed you in and acted as you (`/me`). That meant:

- it could only see meetings already on *your* calendars, and
- for any meeting you didn't organize, it could only **forward** the invite —
  the new hire got a `Fwd:` email they had to accept, landed as *optional*, and
  never became a proper attendee on the organizer's copy.

Since team standups are usually organized by the team lead rather than by whoever
runs onboarding, that fallback was hitting the main case the tool exists for.

With an app-only token this version:

1. reads the mirrored person's calendars directly (no "were you invited too?" gap), then
2. looks up **the organizer's own copy** of each meeting and patches its attendee
   list, so the new hire becomes a genuine required attendee on the real meeting.

Exchange gives every mailbox a different event id for the same meeting, so step 2
matches on `iCalUId`, which is stable across mailboxes. See
`findOrganizerCopy()` in `lib/calendar.ts`.

Forwarding still exists, but only as the fallback for meetings whose organizer
this app can't reach — external organizers, or mailboxes outside the access
policy. The preview labels which path each meeting will take before anything is sent.

## How meetings are found

Two separate things, worth not confusing:

**The source calendar** is what finds meetings. The tool reads
`/users/{sourceMailbox}/calendars` and expands each with `calendarView`. Everything
on that person's calendar is a candidate, regardless of how they were invited. This
is the robust path — mirror a person, not a group.

**The reference filter** is optional narrowing on top. It matches an address against
the event's attendee list and organizer. Because a group address often doesn't
survive onto the event — Outlook lets the organizer expand a group into individual
recipients before sending, and the group address then disappears from `attendees`
entirely — each reference is resolved through Graph's directory first:

- Both **Microsoft 365 groups and classic distribution lists** resolve, since both
  appear in `/groups` (a DL is `mailEnabled && !securityEnabled` with no `Unified`
  group type). Mail-enabled security groups too.
- `transitiveMembers` flattens nesting, so a DL containing a DL resolves to people.
- Aliases are covered via `proxyAddresses`, so a group invited under a secondary
  address still matches.
- An address that isn't a group falls through to plain literal matching — that's
  the normal case for mirroring a colleague.

A meeting matches if the group's **own address** is on it (unambiguous), or if at
least `minGroupMembers` of its members are, excluding the mirrored person. That
threshold defaults to 2 because member matching is broad: at 1, a 1:1 between the
mirrored person and any single teammate would match the whole team's filter.

Note that a **distribution list has no calendar of its own** — it's a routing rule.
Only a Microsoft 365 group has a mailbox, and reading a group's calendar needs
`/groups/{id}/events` rather than `/users/{x}/calendars`, which this tool doesn't
do. So a group address works as a *filter* but not as the *mailbox to mirror*.
Shared and room mailboxes do work as a source, since those are real mailboxes.

## Known limitations

- **Adding an attendee notifies everyone.** Graph has no equivalent of Outlook's
  "send update to added attendees only", so Exchange mails a meeting update to the
  whole series. Onboarding one person across a dozen recurring meetings means the
  team gets a dozen update emails. There is no workaround at the API level.
- **External organizers still fall back to forwarding**, with all the old caveats
  (optional-only, arrives as a forwarded mail, organizer may have forwarding disabled).
- The scan window defaults to 90 days ahead. A series is picked up if it has an
  occurrence in that window, no matter when it started — but a quarterly meeting
  needs a wider window.

## Setup

### 1. Graph app registration (the calendar writes)

Create an app registration and grant **Application** permissions, then click
*Grant admin consent*:

- `Calendars.ReadWrite`
- `User.Read.All`
- `Group.Read.All` — only for the optional reference filter (see below). If it's
  missing, the filter degrades to literal address matching and says so in the
  preview instead of silently returning nothing.

Create a client secret. Fill in `GRAPH_TENANT_ID`, `GRAPH_CLIENT_ID`,
`GRAPH_CLIENT_SECRET`.

> `Calendars.ReadWrite` as an application permission grants access to **every
> mailbox in the tenant**. Scope it down with an Exchange Application Access
> Policy so the app can only touch the mailboxes it needs — typically a mail-
> enabled security group containing the people you mirror plus the organizers of
> their meetings:
>
> ```powershell
> New-ApplicationAccessPolicy -AppId <GRAPH_CLIENT_ID> `
>   -PolicyScopeGroupId onboarding-scope@coverdash.com `
>   -AccessRight RestrictAccess `
>   -Description "Onboarding meeting adder"
> ```
>
> Any organizer outside that group degrades to the forward fallback rather than
> failing — that's expected, and the preview says so.

### 2. Sign-in app registration (who may use the tool)

Separate concern from the Graph token: this only proves who is sitting at the
keyboard. A standard Auth.js Entra ID setup — redirect URI
`https://<your-app>.vercel.app/api/auth/callback/microsoft-entra-id`.

Set `AUTH_MICROSOFT_ENTRA_ID_ID`, `AUTH_MICROSOFT_ENTRA_ID_SECRET`,
`AUTH_MICROSOFT_ENTRA_ID_ISSUER`, and `AUTH_SECRET` (`openssl rand -base64 32`).

`ALLOWED_USERS` is a comma-separated allowlist. It **fails closed** — empty means
nobody gets in. This is a real security boundary here, because the app token can
write to other people's calendars.

### 3. Deploy

```bash
vercel --prod
```

Copy `.env.example` to `.env.local` for local dev (`npm run dev`).

## How a run works

1. **Preview** — scans the mirrored mailbox's calendars over the window, dedupes
   occurrences down to one row per series, and resolves each organizer's copy to
   decide direct-add vs forward. Sends nothing.
2. **Select** — uncheck anything the new hire shouldn't be on. List or week-grid view.
3. **Apply** — processes the checked rows one at a time. The server re-reads every
   event from Graph at this point; the browser only ever hands back ids, so a
   tampered payload can't rewrite a meeting's attendee list.

## Layout

| Path | What's there |
| --- | --- |
| `lib/graph.ts` | App-only token minting + Graph fetch/pagination helpers |
| `lib/calendar.ts` | Scan, organizer-copy resolution, apply, week occurrences |
| `lib/config.ts` | Allowlist and preset sources |
| `app/api/preview` | Step 1 — dry run |
| `app/api/apply` | Step 3 — the only route that writes |
| `app/api/calendar-week` | Week grid data |
| `components/OnboardingTool.tsx` | The three-step UI |
| `components/WeekCalendar.tsx` | Week grid |
