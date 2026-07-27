import { signIn } from "@/auth";

export default function LoginPage() {
  return (
    <>
      <div className="topbar">
        <div>
          <h1>
            Onboarding <span className="arrow">→</span> meetings
          </h1>
          <p className="muted">
            Sign in with Microsoft to add a new hire to a colleague&rsquo;s recurring meetings.
          </p>
        </div>
      </div>
      <section className="panel">
        <div className="panel-content">
          <h2>Sign in to get started</h2>
          <p className="muted small" style={{ marginTop: 8 }}>
            Access is limited to the allowlist in <code>ALLOWED_USERS</code>.
          </p>
          <form
            action={async () => {
              "use server";
              await signIn("microsoft-entra-id", { redirectTo: "/" });
            }}
          >
            <button className="button" type="submit" style={{ marginTop: 16 }}>
              Sign in with Microsoft
            </button>
          </form>
        </div>
      </section>
    </>
  );
}
