/* eslint-disable @next/next/no-img-element */
import { AlertTriangle } from "lucide-react";
import { signIn } from "@/auth";
import { entraIssuer } from "@/auth.config";

export default function LoginPage() {
  // Without an issuer the provider silently targets /common, which fails on a
  // single-tenant app registration with AADSTS50194 *after* the redirect to
  // Microsoft — far from the actual cause. Say so before the round trip.
  const issuer = entraIssuer();

  return (
    <main>
      <div className="signin-wrap">
        <div className="signin-mark">
          <img src="/logo.svg" alt="Coverdash" width={198} height={36} />
        </div>

        {!issuer && (
          <section className="panel warning">
            <div className="panel-content">
              <h2 style={{ display: "flex", alignItems: "center", gap: 8 }}>
                <AlertTriangle size={18} strokeWidth={1.5} />
                Setup needed
              </h2>
              <p className="small" style={{ marginTop: 8, lineHeight: 1.6 }}>
                Neither <code>AUTH_MICROSOFT_ENTRA_ID_ISSUER</code> nor{" "}
                <code>GRAPH_TENANT_ID</code> is set, so sign-in would be sent to the{" "}
                <code>/common</code> endpoint and rejected with <code>AADSTS50194</code>. Set{" "}
                <code>GRAPH_TENANT_ID</code> to your Directory (tenant) ID and redeploy — Vercel
                does not inject new variables into an existing deployment.
              </p>
            </div>
          </section>
        )}

        <section className="panel">
          <div className="panel-content" style={{ textAlign: "center", padding: "32px 28px" }}>
            <h2>Sign in to CalSync</h2>
            <p className="muted small" style={{ margin: "10px 0 24px", lineHeight: 1.6 }}>
              Add a new hire to a colleague&rsquo;s recurring meetings. Access is limited to the
              allowlist in <code>ALLOWED_USERS</code>.
            </p>
            <form
              action={async () => {
                "use server";
                await signIn("microsoft-entra-id", { redirectTo: "/" });
              }}
            >
              <button className="button" type="submit" style={{ width: "100%" }}>
                Continue with Microsoft
              </button>
            </form>
          </div>
        </section>
      </div>
    </main>
  );
}
