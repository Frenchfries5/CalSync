import { redirect } from "next/navigation";
import { AlertTriangle, LogOut } from "lucide-react";
import { auth, signOut } from "@/auth";
import { isAllowedUser, presetSources } from "@/lib/config";
import AppBar from "@/components/AppBar";
import OnboardingTool from "@/components/OnboardingTool";

// AUTH_MICROSOFT_ENTRA_ID_ISSUER is absent on purpose: entraIssuer() derives it
// from GRAPH_TENANT_ID when it isn't set explicitly, so requiring it here would
// report a problem that doesn't exist.
const REQUIRED_ENV = [
  "GRAPH_TENANT_ID",
  "GRAPH_CLIENT_ID",
  "GRAPH_CLIENT_SECRET",
  "AUTH_MICROSOFT_ENTRA_ID_ID",
  "AUTH_MICROSOFT_ENTRA_ID_SECRET",
  "AUTH_SECRET",
  "ALLOWED_USERS",
];

export default async function HomePage() {
  const session = await auth();
  const email = session?.user?.email;
  if (!email) redirect("/login");

  if (!isAllowedUser(email)) {
    return (
      <>
        <AppBar />
        <main>
          <section className="panel error">
            <div className="panel-content">
              <h2>Not authorized</h2>
              <p className="muted small" style={{ marginTop: 8 }}>
                {email} is not in <code>ALLOWED_USERS</code>.
              </p>
            </div>
          </section>
        </main>
      </>
    );
  }

  const missing = REQUIRED_ENV.filter((key) => !process.env[key]?.trim());

  return (
    <>
      <AppBar>
        <div className="appbar-user">
          <span className="muted small">{session!.user!.name || email}</span>
          <form
            action={async () => {
              "use server";
              await signOut({ redirectTo: "/login" });
            }}
          >
            <button className="button secondary small" type="submit">
              <LogOut size={15} strokeWidth={1.5} />
              Log out
            </button>
          </form>
        </div>
      </AppBar>

      <main>
        <h1>
          Add a new hire to the <span className="hl">right</span> meetings
        </h1>
        <p className="muted" style={{ margin: 0, fontSize: 15 }}>
          Mirror an existing schedule onto someone joining, without forwarding invites by hand.
        </p>

        {missing.length > 0 && (
          <section className="panel warning">
            <div className="panel-content">
              <h2 style={{ display: "flex", alignItems: "center", gap: 8 }}>
                <AlertTriangle size={18} strokeWidth={1.5} />
                Setup needed
              </h2>
              <p className="small" style={{ marginTop: 8 }}>
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
      </main>
    </>
  );
}
