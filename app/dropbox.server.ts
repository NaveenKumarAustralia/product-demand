// ─── Dropbox media source ─────────────────────────────────────────────────────
//
// Read access (+ rename) to a couple of shared Dropbox folders, used as an image
// source for the production portal. The account is a Dropbox BUSINESS TEAM, so
// every file call needs team headers (Select-User + Path-Root). We mint
// short-lived access tokens from a long-lived refresh token and cache them, and
// hard-scope browsing to ALLOWED_ROOTS so the portal only ever sees those
// folders — never the rest of the company Dropbox.
//
// Ported from the Ecommerce-Dashboard growth engine. Configure via Railway:
//   DROPBOX_APP_KEY, DROPBOX_APP_SECRET, DROPBOX_REFRESH_TOKEN
//   DROPBOX_TEAM_MEMBER_EMAIL (optional), DROPBOX_ALLOWED_FOLDERS (optional csv)

import { tmpdir } from "node:os";
import { join } from "node:path";
import { mkdirSync, existsSync, readFileSync, writeFileSync } from "node:fs";
import { createHash } from "node:crypto";

const TOKEN_URL = "https://api.dropbox.com/oauth2/token";
const RPC_BASE = "https://api.dropboxapi.com/2";
const CONTENT_BASE = "https://content.dropboxapi.com/2";

// On-disk thumbnail cache so a browsed folder hits Dropbox once per image.
const THUMB_CACHE = join(tmpdir(), "pd-dbx-thumbs");
try { mkdirSync(THUMB_CACHE, { recursive: true }); } catch { /* best effort */ }

const APP_KEY = process.env.DROPBOX_APP_KEY;
const APP_SECRET = process.env.DROPBOX_APP_SECRET;
const REFRESH_TOKEN = process.env.DROPBOX_REFRESH_TOKEN;
const MEMBER_EMAIL = process.env.DROPBOX_TEAM_MEMBER_EMAIL || "info@karmaeast.com.au";

// The only folders the portal may browse. Overridable via env (comma-list).
const ALLOWED_ROOTS = (process.env.DROPBOX_ALLOWED_FOLDERS ||
  "/**COLLECTIONS*,/Karma East/Karma East Branding")
  .split(",").map((s) => s.trim()).filter(Boolean);

export const isConfigured = () => Boolean(APP_KEY && APP_SECRET && REFRESH_TOKEN);

export type DropboxEntry =
  | { type: "folder"; name: string; path: string }
  | { type: "file"; name: string; path: string; id?: string; size?: number; kind: DropboxKind; rev?: string; modified?: string };
export type DropboxKind = "image" | "video" | "other";

// ─── Access-token cache (minted from the refresh token) ───────────────────────
let tokenCache: { value: string | null; expiresAt: number } = { value: null, expiresAt: 0 };
async function accessToken(): Promise<string> {
  if (tokenCache.value && Date.now() < tokenCache.expiresAt - 60_000) return tokenCache.value;
  const res = await fetch(TOKEN_URL, {
    method: "POST",
    headers: { "Content-Type": "application/x-www-form-urlencoded" },
    body: new URLSearchParams({
      grant_type: "refresh_token", refresh_token: REFRESH_TOKEN ?? "",
      client_id: APP_KEY ?? "", client_secret: APP_SECRET ?? "",
    }),
  });
  if (!res.ok) throw new Error(`Dropbox token refresh failed: ${res.status} ${await res.text()}`);
  const j = await res.json() as { access_token: string; expires_in?: number };
  tokenCache = { value: j.access_token, expiresAt: Date.now() + (j.expires_in || 14400) * 1000 };
  return j.access_token;
}

// ─── Team context (member id + shared root namespace), discovered once ────────
let ctx: { team: boolean; member: string | null; rootNs: string | null } | null = null;
async function context() {
  if (ctx) return ctx;
  const token = await accessToken();
  const probe = await fetch(`${RPC_BASE}/users/get_current_account`, {
    method: "POST", headers: { Authorization: `Bearer ${token}` },
  });
  if (probe.ok) {
    const j = await probe.json() as { root_info?: { root_namespace_id?: string } };
    ctx = { team: false, member: null, rootNs: j.root_info?.root_namespace_id || null };
    return ctx;
  }
  const body = await probe.text();
  if (!/Business team/i.test(body)) throw new Error(`Dropbox auth failed: ${body.slice(0, 200)}`);
  const member = await findMember(token, MEMBER_EMAIL);
  const acct = await rpc("/users/get_current_account", null, { token, member }) as { root_info?: { root_namespace_id?: string } };
  ctx = { team: true, member, rootNs: acct.root_info?.root_namespace_id || null };
  return ctx;
}

async function findMember(token: string, email: string): Promise<string> {
  const res = await fetch(`${RPC_BASE}/team/members/list_v2`, {
    method: "POST",
    headers: { Authorization: `Bearer ${token}`, "Content-Type": "application/json" },
    body: JSON.stringify({ limit: 100 }),
  });
  if (!res.ok) throw new Error(`Dropbox team lookup failed: ${await res.text()}`);
  const j = await res.json() as { members?: Array<{ profile?: { email?: string; team_member_id?: string; status?: { ".tag"?: string } } }> };
  const members = j.members || [];
  const hit = members.find((m) => (m.profile?.email || "").toLowerCase() === email.toLowerCase())
    || members.find((m) => m.profile?.status?.[".tag"] === "active");
  if (!hit?.profile?.team_member_id) throw new Error("No Dropbox team member found to act as.");
  return hit.profile.team_member_id;
}

async function teamHeaders(): Promise<Record<string, string>> {
  const c = await context();
  const h: Record<string, string> = {};
  if (c.team && c.member) h["Dropbox-API-Select-User"] = c.member;
  if (c.rootNs) h["Dropbox-API-Path-Root"] = JSON.stringify({ ".tag": "root", root: String(c.rootNs) });
  return h;
}

async function rpc(path: string, arg: unknown, opts: { token?: string; member?: string } = {}): Promise<any> {
  const tok = opts.token || (await accessToken());
  const headers: Record<string, string> = { Authorization: `Bearer ${tok}` };
  if (arg != null) headers["Content-Type"] = "application/json";
  if (opts.member) {
    headers["Dropbox-API-Select-User"] = opts.member;
    if (ctx?.rootNs) headers["Dropbox-API-Path-Root"] = JSON.stringify({ ".tag": "root", root: String(ctx.rootNs) });
  } else {
    Object.assign(headers, await teamHeaders());
  }
  const res = await fetch(`${RPC_BASE}${path}`, {
    method: "POST", headers, body: arg == null ? null : JSON.stringify(arg),
  });
  if (!res.ok) throw new Error(`Dropbox ${path} → ${res.status}: ${(await res.text()).slice(0, 300)}`);
  return res.json();
}

// ─── Path safety: only inside the allowed roots ───────────────────────────────
const norm = (p: string) => ("/" + String(p || "").replace(/^\/+/, "")).replace(/\/+$/, "") || "/";
function isAllowed(path: string): boolean {
  const p = norm(path);
  return ALLOWED_ROOTS.some((root) => { const r = norm(root); return p === r || p.startsWith(r + "/"); });
}
function assertAllowed(path: string) {
  if (!isAllowed(path)) throw Object.assign(new Error("That folder is outside the allowed Dropbox folders."), { forbidden: true });
}

const IMAGE_EXT = new Set(["jpg", "jpeg", "png", "gif", "webp", "heic", "tif", "tiff", "bmp"]);
const VIDEO_EXT = new Set(["mp4", "mov", "m4v", "webm", "avi", "mkv"]);
const kindOf = (name: string): DropboxKind => {
  const ext = (name.split(".").pop() || "").toLowerCase();
  return IMAGE_EXT.has(ext) ? "image" : VIDEO_EXT.has(ext) ? "video" : "other";
};
export const fileKind = (name: string) => kindOf(name);

export async function list(path = ""): Promise<{ path: string; entries: DropboxEntry[] }> {
  const p = norm(path);
  if (p === "/" || !path) {
    return { path: "", entries: ALLOWED_ROOTS.map((r) => ({ type: "folder" as const, name: r.replace(/^\//, ""), path: norm(r) })) };
  }
  assertAllowed(p);
  const out: DropboxEntry[] = [];
  let cursor: string | null = null;
  do {
    const j: any = cursor
      ? await rpc("/files/list_folder/continue", { cursor })
      : await rpc("/files/list_folder", { path: p, include_media_info: false, limit: 500 });
    for (const e of j.entries || []) {
      if (e[".tag"] === "folder") out.push({ type: "folder", name: e.name, path: e.path_display });
      else if (e[".tag"] === "file") out.push({ type: "file", name: e.name, path: e.path_display, id: e.id, size: e.size, kind: kindOf(e.name), rev: e.rev, modified: e.client_modified });
    }
    cursor = j.has_more ? j.cursor : null;
  } while (cursor);
  out.sort((a, b) => (a.type === b.type ? a.name.localeCompare(b.name) : a.type === "folder" ? -1 : 1));
  return { path: p, entries: out };
}

export async function temporaryLink(path: string): Promise<string> {
  assertAllowed(path);
  const j = await rpc("/files/get_temporary_link", { path: norm(path) });
  return j.link;
}

export async function thumbnail(path: string, size = "w256h256", rev = ""): Promise<Buffer> {
  assertAllowed(path);
  const key = createHash("sha1").update(`${size}:${rev}:${norm(path)}`).digest("hex") + ".jpg";
  const cacheFile = join(THUMB_CACHE, key);
  if (existsSync(cacheFile)) {
    try { return readFileSync(cacheFile); } catch { /* refetch */ }
  }
  const token = await accessToken();
  const arg = { resource: { ".tag": "path", path: norm(path) }, format: "jpeg", size, mode: "strict" };
  const res = await fetch(`${CONTENT_BASE}/files/get_thumbnail_v2`, {
    method: "POST",
    headers: { Authorization: `Bearer ${token}`, "Dropbox-API-Arg": JSON.stringify(arg), ...(await teamHeaders()) },
  });
  if (!res.ok) throw new Error(`Dropbox thumbnail → ${res.status}`);
  const buf = Buffer.from(await res.arrayBuffer());
  try { writeFileSync(cacheFile, buf); } catch { /* best effort */ }
  return buf;
}

export async function download(path: string): Promise<{ bytes: Buffer; name: string }> {
  assertAllowed(path);
  const token = await accessToken();
  const res = await fetch(`${CONTENT_BASE}/files/download`, {
    method: "POST",
    headers: { Authorization: `Bearer ${token}`, "Dropbox-API-Arg": JSON.stringify({ path: norm(path) }), ...(await teamHeaders()) },
  });
  if (!res.ok) throw new Error(`Dropbox download → ${res.status}`);
  return { bytes: Buffer.from(await res.arrayBuffer()), name: norm(path).split("/").pop() || "file" };
}

export async function rename(fromPath: string, newName: string): Promise<{ name: string; path: string; id: string }> {
  const from = norm(fromPath);
  assertAllowed(from);
  const clean = String(newName || "").trim().replace(/[/\\]/g, "");
  if (!clean) throw new Error("New name required.");
  const dir = from.slice(0, from.lastIndexOf("/"));
  const to = norm(`${dir}/${clean}`);
  assertAllowed(to);
  const j = await rpc("/files/move_v2", { from_path: from, to_path: to, autorename: false });
  const e = j.metadata;
  return { name: e.name, path: e.path_display, id: e.id };
}

// Recursive image/video search across the allowed folders only — NOT the whole
// Dropbox. Runs Dropbox search within each allowed root.
export async function search(query: string, { limit = 80 }: { limit?: number } = {}): Promise<DropboxEntry[]> {
  const q = String(query || "").trim();
  if (!q) return [];
  const out: DropboxEntry[] = [];
  const seen = new Set<string>();
  for (const root of ALLOWED_ROOTS) {
    let cursor: string | null = null, guard = 0;
    do {
      const body = cursor
        ? { cursor }
        : { query: q, options: { path: norm(root), max_results: 100, file_status: "active", filename_only: true } };
      const j: any = await rpc(cursor ? "/files/search/continue_v2" : "/files/search_v2", body);
      for (const m of j.matches || []) {
        const md = m.metadata?.metadata;
        if (!md || md[".tag"] !== "file") continue;
        const kind = kindOf(md.name);
        if (kind !== "image" && kind !== "video") continue;
        if (!isAllowed(md.path_display) || seen.has(md.path_display)) continue;
        seen.add(md.path_display);
        out.push({ type: "file", name: md.name, path: md.path_display, id: md.id, kind, rev: md.rev });
      }
      cursor = j.has_more ? j.cursor : null;
      guard++;
    } while (cursor && guard < 8 && out.length < limit);
    if (out.length >= limit) break;
  }
  return out.slice(0, limit);
}

export async function sharedLink(path: string): Promise<string> {
  const p = norm(path);
  assertAllowed(p);
  let url: string | null = null;
  try {
    const created = await rpc("/sharing/create_shared_link_with_settings", { path: p });
    url = created.url;
  } catch {
    try {
      const listed = await rpc("/sharing/list_shared_links", { path: p, direct_only: true });
      url = (listed.links || [])[0]?.url || null;
    } catch { /* fall through */ }
  }
  if (!url) throw new Error("Could not create a Dropbox link for that file.");
  if (/[?&]dl=0/.test(url)) return url.replace(/([?&])dl=0/, "$1raw=1");
  if (/[?&]raw=1/.test(url)) return url;
  return url + (url.includes("?") ? "&" : "?") + "raw=1";
}

export const allowedRoots = () => ALLOWED_ROOTS.slice();
