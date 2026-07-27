import NextAuth from "next-auth";
import authConfig from "./auth.config";
import { isAllowedUser } from "./lib/config";

export const { handlers, auth, signIn, signOut } = NextAuth({
  ...authConfig,
  callbacks: {
    ...authConfig.callbacks,
    signIn({ user }) {
      return isAllowedUser(user.email);
    },
  },
});

/**
 * Throws a JSON 401 unless the caller is signed in AND on the allowlist.
 * Route handlers catch `instanceof Response` and return it as-is, so the client
 * always gets parseable JSON — never an HTML login redirect.
 */
export async function requireOperator(): Promise<string> {
  const session = await auth();
  const email = session?.user?.email;
  if (!isAllowedUser(email)) {
    throw Response.json({ error: "Not signed in" }, { status: 401 });
  }
  return email!;
}
