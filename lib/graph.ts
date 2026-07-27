/**
 * Microsoft Graph client — APPLICATION (app-only) permissions.
 *
 * This is the key difference from the original Val Town version, which used a
 * delegated token and could therefore only ever touch `/me`. With an app-only
 * token we address mailboxes directly as `/users/{upn}/...`, which is what lets
 * us edit *the organizer's own copy* of a meeting instead of forwarding an
 * invite from the side.
 *
 * Required Azure app registration (Application permissions, admin consented):
 *   - Calendars.ReadWrite   (read/write any mailbox's calendars)
 *   - User.Read.All         (resolve/validate mailbox addresses)
 *
 * IMPORTANT: `Calendars.ReadWrite` as an application permission grants access
 * to EVERY mailbox in the tenant. Scope it down with an Exchange Application
 * Access Policy so this app can only reach the mailboxes it needs. See README.
 */

const GRAPH = "https://graph.microsoft.com/v1.0";

type TokenResponse = {
  access_token?: string;
  expires_in?: number;
  error?: string;
  error_description?: string;
};

// Module-level cache. Serverless instances are short-lived and each holds its
// own copy, which is fine — the token is valid for ~60min and re-minting it is
// a single cheap call.
let cachedToken: { value: string; expiresAt: number } | null = null;

function requireEnv(name: string): string {
  const value = process.env[name]?.trim();
  if (!value) {
    throw new Error(`Missing required environment variable ${name}`);
  }
  return value;
}

export async function appToken(): Promise<string> {
  if (cachedToken && cachedToken.expiresAt > Date.now() + 60_000) {
    return cachedToken.value;
  }
  const tenantId = requireEnv("GRAPH_TENANT_ID");
  const response = await fetch(
    `https://login.microsoftonline.com/${encodeURIComponent(tenantId)}/oauth2/v2.0/token`,
    {
      method: "POST",
      headers: { "Content-Type": "application/x-www-form-urlencoded" },
      body: new URLSearchParams({
        client_id: requireEnv("GRAPH_CLIENT_ID"),
        client_secret: requireEnv("GRAPH_CLIENT_SECRET"),
        grant_type: "client_credentials",
        scope: "https://graph.microsoft.com/.default",
      }),
    },
  );
  const body = (await response.json()) as TokenResponse;
  if (!response.ok || !body.access_token) {
    throw new Error(
      `Could not get a Graph app token: ${body.error_description || body.error || response.status}`,
    );
  }
  cachedToken = {
    value: body.access_token,
    expiresAt: Date.now() + Math.max((body.expires_in || 3600) - 120, 60) * 1000,
  };
  return cachedToken.value;
}

export class GraphError extends Error {
  constructor(
    message: string,
    readonly status: number,
    readonly code?: string,
  ) {
    super(message);
  }
}

async function graphRequest(
  method: string,
  pathOrUrl: string,
  body?: unknown,
  extraHeaders: Record<string, string> = {},
): Promise<unknown> {
  const url = pathOrUrl.startsWith("https://") ? pathOrUrl : `${GRAPH}${pathOrUrl}`;
  const response = await fetch(url, {
    method,
    headers: {
      Authorization: `Bearer ${await appToken()}`,
      Accept: "application/json",
      ...(body === undefined ? {} : { "Content-Type": "application/json" }),
      ...extraHeaders,
    },
    ...(body === undefined ? {} : { body: JSON.stringify(body) }),
  });

  const text = await response.text();
  let parsed: unknown = text;
  try {
    parsed = text ? JSON.parse(text) : {};
  } catch {
    // Non-JSON body (rare); fall through with the raw text.
  }

  if (!response.ok) {
    const err = (parsed as { error?: { code?: string; message?: string } })?.error;
    throw new GraphError(
      err?.message || text || `Graph ${method} failed`,
      response.status,
      err?.code,
    );
  }
  return parsed;
}

export function graphGet<T>(path: string): Promise<T> {
  return graphRequest("GET", path) as Promise<T>;
}

export function graphPatch<T>(path: string, body: unknown): Promise<T> {
  return graphRequest("PATCH", path, body) as Promise<T>;
}

export function graphPost<T>(path: string, body: unknown): Promise<T> {
  return graphRequest("POST", path, body) as Promise<T>;
}

/** Follows @odata.nextLink and returns every page's `value` concatenated. */
export async function graphAll<T>(path: string): Promise<T[]> {
  const values: T[] = [];
  let next: string | undefined = path;
  // Hard page cap so a pathological mailbox can't spin the function forever.
  for (let page = 0; next && page < 40; page++) {
    const body: { value?: T[]; "@odata.nextLink"?: string } = await graphGet(next);
    values.push(...(body.value || []));
    next = body["@odata.nextLink"];
  }
  return values;
}

/** Escapes a value for use inside an OData string literal. */
export function odataString(value: string): string {
  return `'${value.replace(/'/g, "''")}'`;
}
