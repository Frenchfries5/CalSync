import { GraphError, graphAll, graphGet, odataString } from "./graph";
import type { ResolvedReference } from "./types";

/**
 * Turns a "reference" address typed into the UI into the set of addresses that
 * should count as a match on a meeting's attendee list.
 *
 * Why this exists: inviting a group to a meeting doesn't reliably leave the
 * group's address on the event. Outlook lets the organizer expand a group into
 * its individual members before sending, and plenty of people do. When that
 * happens the group address is gone from `attendees` entirely and a literal
 * string match finds nothing — with no error to explain the empty result.
 *
 * So we resolve the address against Graph's directory and match on the members
 * too. Both Microsoft 365 groups and classic distribution lists show up in
 * `/groups` (a DL is mailEnabled + !securityEnabled + no "Unified" groupType),
 * so the same lookup covers both.
 *
 * Requires the application permission `Group.Read.All`.
 */

export type GraphGroup = {
  id: string;
  displayName?: string;
  mail?: string;
  groupTypes?: string[];
  mailEnabled?: boolean;
  securityEnabled?: boolean;
  proxyAddresses?: string[];
};

type DirectoryMember = {
  "@odata.type"?: string;
  mail?: string;
  userPrincipalName?: string;
};

// Directory membership barely changes within a session and a scan resolves the
// same handful of addresses repeatedly, so a short cache saves real round-trips.
const CACHE_TTL_MS = 5 * 60 * 1000;
const cache = new Map<string, { value: ResolvedReference; expiresAt: number }>();

export function classify(group: GraphGroup): ResolvedReference["groupType"] {
  if (group.groupTypes?.includes("Unified")) return "microsoft365";
  if (group.mailEnabled && !group.securityEnabled) return "distribution";
  if (group.mailEnabled && group.securityEnabled) return "mail-enabled-security";
  return "security";
}

/** "SMTP:growth@coverdash.com" / "smtp:growth@..." → the address. */
function smtpAddresses(proxyAddresses: string[] | undefined): string[] {
  return (proxyAddresses || [])
    .filter((entry) => entry.toLowerCase().startsWith("smtp:"))
    .map((entry) => entry.slice(5).trim().toLowerCase())
    .filter(Boolean);
}

export async function findGroup(address: string): Promise<GraphGroup | null> {
  const select = "id,displayName,mail,groupTypes,mailEnabled,securityEnabled,proxyAddresses";
  const byMail = await graphGet<{ value?: GraphGroup[] }>(
    `/groups?$filter=mail eq ${odataString(address)}&$select=${select}&$top=2`,
  );
  if (byMail.value?.length) return byMail.value[0];

  // A group invited under an alias won't match on `mail`, which holds only the
  // primary address. proxyAddresses carries every alias, but isn't filterable —
  // so fall back to matching the local part and checking aliases client-side.
  const localPart = address.split("@")[0];
  if (!localPart) return null;
  const byNickname = await graphGet<{ value?: GraphGroup[] }>(
    `/groups?$filter=mailNickname eq ${odataString(localPart)}&$select=${select}&$top=10`,
  );
  return (
    byNickname.value?.find((group) => smtpAddresses(group.proxyAddresses).includes(address)) || null
  );
}

async function memberAddresses(groupId: string): Promise<string[]> {
  // transitiveMembers flattens nested groups, so a DL containing a DL resolves
  // to the actual people rather than to another group object.
  const members = await graphAll<DirectoryMember>(
    `/groups/${encodeURIComponent(groupId)}/transitiveMembers?$select=id,mail,userPrincipalName&$top=999`,
  );
  const addresses = new Set<string>();
  for (const member of members) {
    const mail = member.mail?.trim().toLowerCase();
    const upn = member.userPrincipalName?.trim().toLowerCase();
    if (mail) addresses.add(mail);
    if (upn) addresses.add(upn);
  }
  return [...addresses];
}

export async function resolveReference(input: string): Promise<ResolvedReference> {
  const address = input.trim().toLowerCase();
  const cached = cache.get(address);
  if (cached && cached.expiresAt > Date.now()) return cached.value;

  const plain: ResolvedReference = {
    input,
    kind: "address",
    addresses: [address],
    memberAddresses: [],
  };

  let resolved: ResolvedReference = plain;
  try {
    const group = await findGroup(address);
    if (group) {
      const members = await memberAddresses(group.id);
      const own = new Set<string>([address, ...smtpAddresses(group.proxyAddresses)]);
      if (group.mail) own.add(group.mail.toLowerCase());
      resolved = {
        input,
        kind: "group",
        displayName: group.displayName,
        groupType: classify(group),
        addresses: [...own],
        memberAddresses: members,
      };
    }
  } catch (error) {
    // A directory lookup failure must not sink the whole scan — fall back to
    // the literal address, but say so, because a silent downgrade here looks
    // identical to "this group is on no meetings".
    if (error instanceof GraphError) {
      resolved = {
        ...plain,
        warning:
          error.status === 403
            ? "Could not read the directory (is Group.Read.All granted?) — matched on the literal address only."
            : `Directory lookup failed (${error.status}) — matched on the literal address only.`,
      };
    } else {
      throw error;
    }
  }

  cache.set(address, { value: resolved, expiresAt: Date.now() + CACHE_TTL_MS });
  return resolved;
}

export function resolveReferences(inputs: string[]): Promise<ResolvedReference[]> {
  return Promise.all(inputs.map((input) => resolveReference(input)));
}
