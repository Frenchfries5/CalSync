/* eslint-disable @next/next/no-img-element */

/**
 * Coverdash app chrome. The wordmark is the design system's logo.svg, used
 * as-is rather than recreated — the brand mark is not something to approximate.
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
        {children}
      </div>
    </header>
  );
}
