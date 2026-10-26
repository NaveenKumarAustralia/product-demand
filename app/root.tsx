import { useEffect } from "react";
import { Links, Meta, Outlet, Scripts, ScrollRestoration, useRouteError, isRouteErrorResponse } from "react-router";

// App-wide: stop the mouse wheel from changing a focused <input type="number">
// (quantities, prices, etc.) — a native browser behaviour that kept nudging
// numbers by accident. The wheel still scrolls the nearest scroll container (or
// the page), and the field keeps focus, so nothing is saved by accident.
function useNoWheelNumberChange() {
  useEffect(() => {
    const block = (e: WheelEvent) => {
      const el = e.currentTarget as HTMLElement;
      e.preventDefault();
      let node = el.parentElement;
      while (node) {
        const oy = getComputedStyle(node).overflowY;
        if ((oy === "auto" || oy === "scroll") && node.scrollHeight > node.clientHeight) break;
        node = node.parentElement;
      }
      if (node) { node.scrollTop += e.deltaY; node.scrollLeft += e.deltaX; }
      else window.scrollBy(e.deltaX, e.deltaY);
    };
    const onFocusIn = (e: FocusEvent) => {
      const el = e.target as HTMLElement | null;
      if (el instanceof HTMLInputElement && el.type === "number") el.addEventListener("wheel", block, { passive: false });
    };
    const onFocusOut = (e: FocusEvent) => {
      const el = e.target as HTMLElement | null;
      if (el instanceof HTMLInputElement && el.type === "number") el.removeEventListener("wheel", block);
    };
    document.addEventListener("focusin", onFocusIn);
    document.addEventListener("focusout", onFocusOut);
    return () => { document.removeEventListener("focusin", onFocusIn); document.removeEventListener("focusout", onFocusOut); };
  }, []);
}

export function ErrorBoundary() {
  const error = useRouteError();
  const message = isRouteErrorResponse(error)
    ? error.data
    : error instanceof Error
    ? error.message
    : "An unexpected error occurred.";
  return (
    <html lang="en">
      <head>
        <meta charSet="utf-8" />
        <meta name="viewport" content="width=device-width,initial-scale=1" />
        <title>Error</title>
      </head>
      <body style={{ fontFamily: "sans-serif", padding: "2rem", color: "#333" }}>
        <h1 style={{ marginBottom: "0.5rem" }}>Something went wrong</h1>
        <p style={{ color: "#666", marginBottom: "1.5rem" }}>{String(message)}</p>
        <a href="/portal" style={{ color: "#0070f3" }}>← Back to portal</a>
      </body>
    </html>
  );
}

export default function App() {
  useNoWheelNumberChange();
  return (
    <html lang="en">
      <head>
        <meta charSet="utf-8" />
        <meta name="viewport" content="width=device-width,initial-scale=1" />
        <link rel="preconnect" href="https://cdn.shopify.com/" />
        <link
          rel="stylesheet"
          href="https://cdn.shopify.com/static/fonts/inter/v4/styles.css"
        />
        <Meta />
        <Links />
      </head>
      <body>
        <Outlet />
        <ScrollRestoration />
        <Scripts />
      </body>
    </html>
  );
}
