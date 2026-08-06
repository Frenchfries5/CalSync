/* eslint-disable @next/next/no-img-element */
import { ExternalLink } from "lucide-react";

/**
 * Coverdash app chrome. The wordmark is the design system's logo.svg, used
 * as-is rather than recreated — the brand mark is not something to approximate.
 *
 * The sister-app link lives here rather than in the page so it shows on every
 * surface that renders the bar, including the not-authorized state. The two
 * tools are halves of the same job: the scheduler builds a new hire's
 * onboarding week from a template, this one adds them to the recurring
 * meetings that already exist.
 */
export default function AppBar({ children }: { children?: React.ReactNode }) {
  return (
    <header className="appbar">
      <div className="appbar-inner">
        <div className="appbar-brand">
          <img src="/logo.svg" alt="Coverdash" width={154} height={28} />
          <span className="appbar-divider" />
          <span className="appbar-title">CalSync</span>
        </div>
        <div className="appbar-right">
          <a
            className="appbar-link"
            href="https://calendarhelper.vercel.app/"
            target="_blank"
            rel="noopener noreferrer"
          >
            <span className="wide-only">Onboarding scheduler</span>
            <span className="narrow-only">Scheduler</span>
            <ExternalLink size={13} strokeWidth={1.5} />
          </a>
          {children}
        </div>
      </div>
    </header>
  );
}
