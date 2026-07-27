import NextAuth from "next-auth";
import authConfig from "./auth.config";

export default NextAuth(authConfig).auth;

export const config = {
  // Pages only. /api is deliberately excluded: the proxy answers an
  // unauthenticated request with an HTML redirect to /login, which a fetch()
  // can't do anything useful with. The API routes gate themselves via
  // requireOperator() and return a real JSON 401 instead.
  matcher: ["/((?!api/|_next/static|_next/image|favicon.ico|.*\\.[\\w]+$).*)"],
};
