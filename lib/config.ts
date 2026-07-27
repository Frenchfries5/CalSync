/**
 * Who is allowed to sign in and use the tool.
 *
 * This matters more here than it did in the Val Town version. That one ran on
 * *your* delegated token, so it could only ever do what you could already do by
 * hand. This one holds an application token that can write to other people's
 * calendars, so the app's own allowlist is a real security boundary — keep it
 * tight, and keep the Exchange Application Access Policy tight too.
 */
export function allowedUsers(): string[] {
  return (process.env.ALLOWED_USERS || "")
    .split(",")
    .map((entry) => entry.trim().toLowerCase())
    .filter(Boolean);
}

export function isAllowedUser(email: string | null | undefined): boolean {
  const list = allowedUsers();
  const normalized = (email || "").trim().toLowerCase();
  if (!normalized) return false;
  // Fail closed: an empty allowlist lets nobody in rather than everybody.
  return list.includes(normalized);
}

/**
 * Prefill for "whose calendar to mirror". Set ONBOARDING_PRESETS to a
 * comma-separated list of addresses, e.g.
 *   ONBOARDING_PRESETS=veronica.r@coverdash.com,growth@coverdash.com
 */
export function presetSources(): string[] {
  return (process.env.ONBOARDING_PRESETS || "")
    .split(",")
    .map((entry) => entry.trim())
    .filter(Boolean);
}
