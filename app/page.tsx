import { redirect } from "next/navigation";
import { auth, signOut } from "@/auth";
import { isAllowedUser, presetSources } from "@/lib/config";
import OnboardingTool from "@/components/OnboardingTool";

const REQUIRED_ENV = [
  "GRAPH_TENANT_ID",
  "GRAPH_CLIENT_ID",
  "GRAPH_CLIENT_SECRET",
  "AUTH_MICROSOFT_ENTRA_ID_ID",
  "AUTH_MICROSOFT_ENTRA_ID_SECRET",
  "AUTH_MICROSOFT_ENTRA_ID_ISSUER",
  "AUTH_SECRET",
  "ALLOWED_USERS",
];

export default async function HomePage() {
  const session = await auth();
  const email = session?.user?.email;
  if (!email) redirect("/login");

  if (!isAllowedUser(email)) {
    return (
      <section className="panel error">
        <div className="panel-content">
          <h2>Not authorized</h2>
          <p className="muted small" style={{ marginTop: 8 }}>
            {email} is not in <code>ALLOWED_USERS</code>.
          </p>
        </div>
      </section>
    );
  }

  const missing = REQUIRED_ENV.filter((key) => !process.env[key]?.trim());

  return (
    <>
      <div className="topbar">
        <div>
          <h1>
            Onboarding <span className="arrow">→</span> meetings
          </h1>
          <p className="muted">Signed in as {session!.user!.name || email}</p>
        </div>
        <form
          action={async () => {
            "use server";
            await signOut({ redirectTo: "/login" });
          }}
        >
          <button className="button secondary" type="submit">
            Log out
          </button>
        </form>
      </div>

      {missing.length > 0 && (
        <section className="panel warning">
          <div className="panel-content">
            <h2>Setup needed</h2>
            <p className="muted small" style={{ marginTop: 8 }}>
              These environment variables are not set:
            </p>
            <ul>
              {missing.map((key) => (
                <li key={key}>
                  <code>{key}</code>
                </li>
              ))}
            </ul>
          </div>
        </section>
      )}

      <OnboardingTool presets={presetSources()} />
    </>
  );
}
