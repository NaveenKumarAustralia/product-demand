import type { LoaderFunctionArgs } from "react-router";
import prisma from "../db.server";
import { requirePreorderPortalUser } from "../preorder/preorder-portal-auth.server";

// Admin one-off: copy the bullet list (and any numbered list) from a SOURCE
// product's description into a TARGET product that was duplicated before the
// "copy description verbatim" fix (so its bullets were dropped). Preserves the
// target's existing wording — it only APPENDS the missing list(s). Writes both
// the Shopify product and the linked Collections row so they match.
//
//   Dry run (see what it WOULD do):
//     /api/collection-fix-description?product=Brixton Dress Indigo Thread&from=Brixton Dress Joy
//   Apply:
//     /api/collection-fix-description?product=Brixton Dress Indigo Thread&from=Brixton Dress Joy&apply=1
//   Full verbatim replace instead of append:  &mode=replace
const API_VERSION = "2025-10";
const numericId = (v: string) => String(v ?? "").split("/").pop()?.replace(/[^0-9]/g, "") ?? "";
const listBlocks = (html: string): string[] => (html.match(/<(ul|ol)[\s\S]*?<\/\1>/gi) ?? []);

export const loader = async ({ request }: LoaderFunctionArgs) => {
  const actor = await requirePreorderPortalUser(request);
  if (actor.admin !== true) return Response.json({ ok: false, error: "Admin only." }, { status: 403 });

  const url = new URL(request.url);
  const targetTitle = (url.searchParams.get("product") ?? "").trim();
  const sourceTitle = (url.searchParams.get("from") ?? "").trim();
  const apply = url.searchParams.get("apply") === "1";
  const mode = (url.searchParams.get("mode") ?? "append").toLowerCase(); // append | replace
  if (!targetTitle || !sourceTitle) return Response.json({ ok: false, error: "Pass ?product=<target title>&from=<source title> (&apply=1 to write, &mode=replace for full copy)." }, { status: 400 });

  const session = await prisma.session.findFirst({ where: { isOnline: false, accessToken: { not: "" } }, orderBy: { expires: "desc" }, select: { shop: true, accessToken: true } });
  if (!session?.shop || !session.accessToken) return Response.json({ ok: false, error: "No offline Shopify session." }, { status: 500 });
  const { shop, accessToken } = session;
  const gql = async (query: string, variables: Record<string, unknown>) => {
    const res = await fetch(`https://${shop}/admin/api/${API_VERSION}/graphql.json`, {
      method: "POST", headers: { "Content-Type": "application/json", "X-Shopify-Access-Token": accessToken },
      body: JSON.stringify({ query, variables }),
    });
    return res.json() as Promise<any>;
  };
  const findProduct = async (title: string) => {
    const j = await gql(`query FindP($q: String!) { products(first: 5, query: $q) { nodes { id title descriptionHtml } } }`, { q: `title:${title}` }).catch(() => null);
    const nodes: any[] = j?.data?.products?.nodes ?? [];
    return nodes.find((n) => String(n.title ?? "").trim().toLowerCase() === title.toLowerCase()) ?? nodes[0] ?? null;
  };

  const source = await findProduct(sourceTitle);
  const target = await findProduct(targetTitle);
  if (!source) return Response.json({ ok: false, error: `Source product "${sourceTitle}" not found.` }, { status: 404 });
  if (!target) return Response.json({ ok: false, error: `Target product "${targetTitle}" not found.` }, { status: 404 });

  const sourceHtml = String(source.descriptionHtml ?? "");
  const targetHtml = String(target.descriptionHtml ?? "");
  const sourceLists = listBlocks(sourceHtml);

  let newHtml: string;
  if (mode === "replace") {
    newHtml = sourceHtml; // full verbatim copy (you then re-word manually)
  } else {
    // Append only the lists the target is missing (preserves target wording).
    if (listBlocks(targetHtml).length) {
      return Response.json({ ok: true, changed: false, note: "Target already has a bullet/numbered list — nothing appended. Use &mode=replace to overwrite fully.", targetTitle, sourceLists }, { headers: { "Cache-Control": "no-store" } });
    }
    if (!sourceLists.length) {
      return Response.json({ ok: false, error: `Source "${sourceTitle}" has no <ul>/<ol> list in its description — nothing to copy.`, sourceHtmlPreview: sourceHtml.slice(0, 400) }, { status: 400 });
    }
    newHtml = `${targetHtml}\n${sourceLists.join("\n")}`;
  }

  if (!apply) {
    return Response.json({ ok: true, dryRun: true, targetTitle, sourceTitle, mode, listsToCopy: sourceLists, before: targetHtml, after: newHtml, note: "Add &apply=1 to write this to Shopify + the collection row." }, { headers: { "Cache-Control": "no-store" } });
  }

  // 1) Write to Shopify.
  const up = await gql(`mutation FixDesc($input: ProductInput!) { productUpdate(input: $input) { product { id } userErrors { field message } } }`, { input: { id: target.id, descriptionHtml: newHtml } });
  const errs = up?.data?.productUpdate?.userErrors ?? [];
  if (errs.length) return Response.json({ ok: false, error: errs.map((e: any) => e.message).join("; ") }, { status: 500 });

  // 2) Mirror into the linked Collections row so the portal popup matches.
  const targetNum = numericId(String(target.id));
  let rowUpdated = false;
  const collections = await prisma.collection.findMany({ select: { id: true, rows: true } });
  for (const c of collections) {
    const rows = Array.isArray(c.rows) ? (c.rows as Array<Record<string, unknown>>) : [];
    let changed = false;
    const next = rows.map((r) => {
      if (numericId(String(r?.["__shopifyProductId"] ?? "")) === targetNum) { changed = true; return { ...r, description: newHtml }; }
      return r;
    });
    if (changed) { await prisma.collection.update({ where: { id: c.id }, data: { rows: next as unknown as object, updatedAt: new Date() } }); rowUpdated = true; }
  }

  return Response.json({ ok: true, applied: true, targetTitle, mode, listsCopied: sourceLists.length, shopifyUpdated: true, rowUpdated, after: newHtml }, { headers: { "Cache-Control": "no-store" } });
};
