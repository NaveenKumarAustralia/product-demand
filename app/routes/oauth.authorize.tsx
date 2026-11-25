import type { ActionFunctionArgs, LoaderFunctionArgs } from "react-router";
import { clientIp, rateLimited } from "../mcp/http.server";
import { allowedRedirectHosts, issueAuthorizationCode, OAuthError, validateAuthorizeRequest, type AuthorizeParams } from "../mcp/oauth-core";
import { portalUserDirectory, prismaOAuthStore } from "../mcp/oauth-store.server";

/**
 * The page ChatGPT opens when Koku clicks "Connect". He signs in with his
 * Production Portal admin login and approves. Rendered as plain HTML (no app
 * shell) and never framed, so it can't be click-jacked.
 */

const PAGE_HEADERS = {
  "Content-Type": "text/html; charset=utf-8",
  "Cache-Control": "no-store",
  "X-Frame-Options": "DENY",
  // form-action must also allow the redirect back to ChatGPT after approval.
  "Content-Security-Policy": `default-src 'none'; style-src 'unsafe-inline'; form-action 'self' ${allowedRedirectHosts().map((h) => `https://${h}`).join(" ")}; frame-ancestors 'none'; base-uri 'none'`,
  "Referrer-Policy": "no-referrer",
};

const esc = (s: string) => s.replace(/[&<>"']/g, (c) => ({ "&": "&amp;", "<": "&lt;", ">": "&gt;", '"': "&quot;", "'": "&#39;" })[c]!);

function page(body: string, status = 200) {
  return new Response(
    `<!doctype html><html lang="en"><head><meta charset="utf-8"><meta name="viewport" content="width=device-width,initial-scale=1"><title>Connect ChatGPT · Karma East Portal</title>
<style>body{font-family:system-ui,-apple-system,sans-serif;background:#f6f5f2;margin:0;display:grid;place-items:center;min-height:100vh;color:#1f2937}
main{background:#fff;max-width:420px;width:92%;padding:28px;border-radius:14px;box-shadow:0 2px 14px rgba(0,0,0,.08)}
h1{font-size:20px;margin:0 0 6px}p{line-height:1.5;margin:8px 0}ul{padding-left:18px;line-height:1.6}
label{display:block;font-size:14px;margin-top:14px}input{width:100%;box-sizing:border-box;padding:10px;border:1px solid #d1d5db;border-radius:8px;font-size:15px;margin-top:4px}
.row{display:flex;gap:10px;margin-top:20px}button{flex:1;padding:11px;border-radius:8px;border:0;font-size:15px;cursor:pointer}
.ok{background:#111827;color:#fff}.no{background:#e5e7eb}.err{background:#fee2e2;color:#991b1b;padding:10px;border-radius:8px;font-size:14px}.small{font-size:13px;color:#6b7280}</style></head>
<body><main>${body}</main></body></html>`,
    { status, headers: PAGE_HEADERS },
  );
}

function hiddenFields(p: AuthorizeParams) {
  const fields: Record<string, string> = {
    response_type: "code",
    client_id: p.clientId,
    redirect_uri: p.redirectUri,
    code_challenge: p.codeChallenge,
    code_challenge_method: "S256",
    scope: p.scope,
    ...(p.state ? { state: p.state } : {}),
    ...(p.resource ? { resource: p.resource } : {}),
  };
  return Object.entries(fields).map(([k, v]) => `<input type="hidden" name="${esc(k)}" value="${esc(v)}">`).join("");
}

function consentForm(clientName: string, p: AuthorizeParams, error?: string) {
  const write = p.scope.includes("portal:write");
  return page(`<h1>Connect ${esc(clientName)} to the Production Portal</h1>
<p>${esc(clientName)} is asking for:</p>
<ul><li><b>Read</b> orders, stock, sales, fabric, shipments and collections</li>
${write ? "<li><b>Prepare changes</b> to production orders (create, amend sizes, status, notes). Every change is previewed and needs your approval in chat before it is applied, and is logged.</li>" : ""}</ul>
<p class="small">The connection stays signed in until you revoke it. Sign in with your Production Portal admin account to approve.</p>
${error ? `<p class="err">${esc(error)}</p>` : ""}
<form method="post" autocomplete="on">${hiddenFields(p)}
<label>Portal username<input name="username" autocomplete="username" required></label>
<label>Password<input name="password" type="password" autocomplete="current-password" required></label>
<div class="row"><button class="no" name="decision" value="deny" formnovalidate>Cancel</button><button class="ok" name="decision" value="approve">Approve</button></div>
</form>`);
}

function errorPage(err: unknown) {
  const msg = err instanceof OAuthError ? err.description : "Something went wrong";
  return page(`<h1>Can't connect</h1><p class="err">${esc(msg)}</p><p class="small">Close this window and try connecting again from ChatGPT.</p>`, 400);
}

export const loader = async ({ request }: LoaderFunctionArgs) => {
  try {
    const { client, params } = await validateAuthorizeRequest(prismaOAuthStore, new URL(request.url).searchParams);
    return consentForm(client.name, params);
  } catch (err) {
    return errorPage(err);
  }
};

export const action = async ({ request }: ActionFunctionArgs) => {
  const form = await request.formData();
  const q = new URLSearchParams();
  for (const [k, v] of form.entries()) if (typeof v === "string" && k !== "password") q.set(k, v);

  let validated: Awaited<ReturnType<typeof validateAuthorizeRequest>>;
  try {
    validated = await validateAuthorizeRequest(prismaOAuthStore, q);
  } catch (err) {
    return errorPage(err);
  }
  const { client, params } = validated;

  if (form.get("decision") !== "approve") {
    const url = new URL(params.redirectUri);
    url.searchParams.set("error", "access_denied");
    if (params.state) url.searchParams.set("state", params.state);
    return new Response(null, { status: 303, headers: { Location: url.toString() } });
  }

  const ip = clientIp(request);
  if (rateLimited(`login:${ip}`, 10, 15 * 60 * 1000)) {
    return consentForm(client.name, params, "Too many attempts. Wait 15 minutes and try again.");
  }
  const user = await portalUserDirectory.verifyLogin(String(form.get("username") ?? ""), String(form.get("password") ?? ""));
  if (!user) return consentForm(client.name, params, "Wrong username or password.");
  if (!user.isAdmin) return consentForm(client.name, params, "Only a portal admin can connect ChatGPT.");

  try {
    const location = await issueAuthorizationCode(prismaOAuthStore, params, user);
    console.info(`MCP: ${user.name} approved client ${client.id} (${client.name}) from ${ip}`);
    return new Response(null, { status: 303, headers: { Location: location } });
  } catch (err) {
    return errorPage(err);
  }
};
