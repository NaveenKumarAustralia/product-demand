import fs from "node:fs";

// Deterministic, fail-closed patch for the oversized legacy portal route.
const path = "app/routes/portal._index.tsx";
let source = fs.readFileSync(path, "utf8");

function replaceOnce(label, before, after) {
  const count = source.split(before).length - 1;
  if (count === 0) {
    if (source.includes(after)) {
      console.log(`[portal patch] ${label}: already applied`);
      return;
    }
    throw new Error(`[portal patch] ${label}: target not found`);
  }
  if (count !== 1) throw new Error(`[portal patch] ${label}: expected 1 target, found ${count}`);
  source = source.replace(before, after);
  console.log(`[portal patch] ${label}: applied`);
}

function replaceExact(label, before, after, expectedCount) {
  const count = source.split(before).length - 1;
  if (count === 0) {
    const appliedCount = source.split(after).length - 1;
    if (appliedCount === expectedCount) {
      console.log(`[portal patch] ${label}: already applied (${appliedCount})`);
      return;
    }
    throw new Error(`[portal patch] ${label}: target not found; applied count=${appliedCount}`);
  }
  if (count !== expectedCount) {
    throw new Error(`[portal patch] ${label}: expected ${expectedCount} targets, found ${count}`);
  }
  source = source.split(before).join(after);
  console.log(`[portal patch] ${label}: applied (${count})`);
}

replaceOnce(
  "imports",
  'import { VisionBoardV2Panel } from "../portal-vision-board";\nimport { unauthenticated } from "../shopify.server";',
  'import { VisionBoardV2Panel } from "../portal-vision-board";\nimport { PreordersDashboard } from "../portal-preorders";\nimport { loadPreorderDashboardData } from "../preorder/preorder-dashboard.server";\nimport { unauthenticated } from "../shopify.server";',
);

replaceOnce(
  "preorder nav item",
  '  { id: "reorder", label: "Reorder Planner", href: "/portal?page=reorder" },\n  { id: "fabric", label: "Fabric in stock", href: "/portal?page=fabric" },',
  '  { id: "reorder", label: "Reorder Planner", href: "/portal?page=reorder" },\n  { id: "preorders", label: "Pre-orders", href: "/portal?page=preorders" },\n  { id: "fabric", label: "Fabric in stock", href: "/portal?page=fabric" },',
);

replaceOnce(
  "default nav order",
  'const DEFAULT_NAV_ORDER: NavItemId[] = ["restock", "jj-restock", "jj-new-products", "reorder", "fabric", "packing", "productinfo", "samples", "visionboard", "collections", "dropbox"];',
  'const DEFAULT_NAV_ORDER: NavItemId[] = ["restock", "jj-restock", "jj-new-products", "reorder", "preorders", "fabric", "packing", "productinfo", "samples", "visionboard", "collections", "dropbox"];',
);

replaceOnce(
  "loader preorder access and data",
  '  const currentUser = getCurrentPortalUser(request, usersWithSeed);\n  // JJ-only supplier lockdown: a non-admin user whose ONLY granted page is',
  '  const currentUser = getCurrentPortalUser(request, usersWithSeed);\n  // Pre-orders uses the portal\'s existing per-page permission model. The menu\n  // is hidden client-side for users without pageAccess.preorders, and this\n  // server-side redirect also blocks direct URL access. Only the superadmin\n  // bypasses the explicit page grant.\n  const canViewPreorders = Boolean(currentUser && (currentUser.role === "superadmin" || currentUser.pageAccess?.preorders));\n  if (page === "preorders" && currentUser && !canViewPreorders) {\n    const fallbackId = DEFAULT_NAV_ORDER.find((id) => Boolean(currentUser.pageAccess?.[id])) ?? "restock";\n    const fallbackHref = fallbackId === "restock" ? "/portal" : `/portal?page=${fallbackId}`;\n    return new Response(null, { status: 302, headers: { Location: fallbackHref } });\n  }\n  const preorderDashboard = page === "preorders" && canViewPreorders\n    ? await loadPreorderDashboardData()\n    : null;\n  // JJ-only supplier lockdown: a non-admin user whose ONLY granted page is',
);

replaceExact(
  "loader return + client destructure preorder data",
  '    activityLogs,\n    navOrder,\n    fabricSheets,',
  '    activityLogs,\n    navOrder,\n    preorderDashboard,\n    fabricSheets,',
  2,
);

replaceOnce(
  "active page title",
  '    : page === "reorder" ? "Reorder Planner"\n    : page === "usa-stock" ? "USA Stock"',
  '    : page === "reorder" ? "Reorder Planner"\n    : page === "preorders" ? "Pre-orders"\n    : page === "usa-stock" ? "USA Stock"',
);

replaceOnce(
  "preorder page renderer",
  '        ) : page === "dropbox" ? (\n          <DropboxPanel />\n        ) : page === "reorder" ? (\n          <ReorderPlannerPage search={reorderSearch} />',
  '        ) : page === "dropbox" ? (\n          <DropboxPanel />\n        ) : page === "preorders" && preorderDashboard ? (\n          <PreordersDashboard data={preorderDashboard} />\n        ) : page === "reorder" ? (\n          <ReorderPlannerPage search={reorderSearch} />',
);

fs.writeFileSync(path, source);
console.log("[portal patch] complete");
