import { signIn } from "@/auth";
import { entraIssuer } from "@/auth.config";

export default function LoginPage() {
  // Without an issuer the provider silently targets /common, which fails on a
  // single-tenant app registration with AADSTS50194 *after* the redirect to
  // Microsoft — far from the actual cause. Say so before the round trip.
  const issuer = entraIssuer();

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
      {!issuer && (
        <section className="panel warning">
          <div className="panel-content">
            <h2>Setup needed</h2>
            <p className="muted small" style={{ marginTop: 8 }}>
              Neither <code>AUTH_MICROSOFT_ENTRA_ID_ISSUER</code> nor{" "}
              <code>GRAPH_TENANT_ID</code> is set, so sign-in would be sent to the{" "}
              <code>/common</code> endpoint and rejected with <code>AADSTS50194</code>. Set{" "}
              <code>GRAPH_TENANT_ID</code> to your Directory (tenant) ID and redeploy — Vercel does
              not inject new variables into an existing deployment.
            </p>
          </div>
        </section>
      )}
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
