import type { ActionFunctionArgs, LoaderFunctionArgs } from "react-router";
import { isConfigured, list, search, allowedRoots, temporaryLink, rename } from "../dropbox.server";

// Rename a Dropbox file in place.  POST { path, newName }
export const action = async ({ request }: ActionFunctionArgs) => {
  if (!isConfigured()) return Response.json({ error: "Dropbox not connected" }, { status: 200 });
  const form = await request.formData();
  const path = String(form.get("path") ?? "");
  const newName = String(form.get("newName") ?? "");
  if (!path || !newName.trim()) return Response.json({ error: "Path and new name required" }, { status: 200 });
  try {
    const result = await rename(path, newName);
    return Response.json({ ok: true, ...result });
  } catch (e) {
    return Response.json({ error: (e as Error).message || "Rename failed" }, { status: 200 });
  }
};

// Browse / search the shared Dropbox folders (scoped to the allowed roots).
//   ?op=config              → { configured, roots }
//   ?op=browse&path=<path>  → { path, entries }
//   ?op=search&q=<query>    → { entries }   (recursive across allowed roots)
export const loader = async ({ request }: LoaderFunctionArgs) => {
  const url = new URL(request.url);
  const op = url.searchParams.get("op") ?? "browse";

  if (!isConfigured()) {
    return Response.json({ configured: false, error: "Dropbox isn't connected. Add the DROPBOX_* env vars." }, { status: 200 });
  }

  try {
    if (op === "config") {
      return Response.json({ configured: true, roots: allowedRoots() });
    }
    if (op === "search") {
      const q = url.searchParams.get("q") ?? "";
      const entries = q.trim().length >= 2 ? await search(q, { limit: 120 }) : [];
      return Response.json({ configured: true, entries });
    }
    if (op === "link") {
      const path = url.searchParams.get("path") ?? "";
      if (!path) return Response.json({ error: "Missing path" }, { status: 400 });
      return Response.json({ link: await temporaryLink(path) });
    }
    // default: browse
    const path = url.searchParams.get("path") ?? "";
    const result = await list(path);
    return Response.json({ configured: true, ...result });
  } catch (e) {
    return Response.json({ configured: true, error: (e as Error).message || "Dropbox error", entries: [] }, { status: 200 });
  }
};
