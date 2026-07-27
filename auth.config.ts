import type { NextAuthConfig } from "next-auth";
import MicrosoftEntraID from "next-auth/providers/microsoft-entra-id";

/**
 * Edge-safe config shared with the middleware. Sign-in here proves *who is
 * using the tool*; it is deliberately unrelated to the Graph token, which is
 * app-only and minted server-side (see lib/graph.ts).
 */
export default {
  providers: [
    MicrosoftEntraID({
      clientId: process.env.AUTH_MICROSOFT_ENTRA_ID_ID,
      clientSecret: process.env.AUTH_MICROSOFT_ENTRA_ID_SECRET,
      issuer: process.env.AUTH_MICROSOFT_ENTRA_ID_ISSUER,
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
