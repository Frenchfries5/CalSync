import type { NextAuthConfig } from "next-auth";
import MicrosoftEntraID from "next-auth/providers/microsoft-entra-id";

/**
 * Edge-safe config shared with the proxy. Sign-in here proves *who is using the
 * tool*; it is deliberately unrelated to the Graph token, which is app-only and
 * minted server-side (see lib/graph.ts).
 */

/**
 * The provider falls back to `https://login.microsoftonline.com/common/v2.0`
 * when no issuer is given, which fails on a single-tenant app registration
 * (AADSTS50194) and would otherwise mean "any Microsoft account may sign in" —
 * a silent widening we never want. GRAPH_TENANT_ID already holds the tenant, so
 * derive from it rather than depending on a second variable being set correctly.
 */
export function entraIssuer(): string | undefined {
  const explicit = process.env.AUTH_MICROSOFT_ENTRA_ID_ISSUER?.trim();
  if (explicit) return explicit;
  const tenantId = process.env.GRAPH_TENANT_ID?.trim();
  return tenantId ? `https://login.microsoftonline.com/${tenantId}/v2.0` : undefined;
}

export default {
  providers: [
    MicrosoftEntraID({
      clientId: process.env.AUTH_MICROSOFT_ENTRA_ID_ID,
      clientSecret: process.env.AUTH_MICROSOFT_ENTRA_ID_SECRET,
      issuer: entraIssuer(),
    }),
  ],
  pages: { signIn: "/login" },
  session: { strategy: "jwt" },
  callbacks: {
    authorized({ auth, request: { nextUrl } }) {
      const { pathname } = nextUrl;
      if (pathname === "/login" || pathname.startsWith("/api/auth")) return true;
      return Boolean(auth?.user);
    },
  },
} satisfies NextAuthConfig;
