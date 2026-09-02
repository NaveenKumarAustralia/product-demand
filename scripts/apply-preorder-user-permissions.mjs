import fs from 'node:fs';

const path = 'app/routes/portal._index.tsx';
let text = fs.readFileSync(path, 'utf8');

function replaceOnce(from, to, label) {
  const count = text.split(from).length - 1;
  if (count !== 1) throw new Error(`${label}: expected exactly one match, found ${count}`);
  text = text.replace(from, to);
}

replaceOnce(
`  pageAccess: Record<string, boolean>;\n  // When true (or for any admin), this account sees the per-product status`,
`  pageAccess: Record<string, boolean>;\n  preorderAccess?: {\n    manage: boolean;\n    eta: boolean;\n    safetyBuffer: boolean;\n    notifications: boolean;\n    reports: boolean;\n  };\n  // When true (or for any admin), this account sees the per-product status`,
'PortalUser preorderAccess type');

replaceOnce(
`        pageAccess: (user.pageAccess && typeof user.pageAccess === "object" && !Array.isArray(user.pageAccess))\n          ? user.pageAccess as Record<string, boolean>\n          : {},\n      };`,
`        pageAccess: (user.pageAccess && typeof user.pageAccess === "object" && !Array.isArray(user.pageAccess))\n          ? user.pageAccess as Record<string, boolean>\n          : {},\n        preorderAccess: (() => {\n          const raw = user.preorderAccess && typeof user.preorderAccess === "object" && !Array.isArray(user.preorderAccess)\n            ? user.preorderAccess as Record<string, unknown>\n            : {};\n          return {\n            manage: Boolean(raw.manage),\n            eta: Boolean(raw.eta),\n            safetyBuffer: Boolean(raw.safetyBuffer),\n            notifications: Boolean(raw.notifications),\n            reports: Boolean(raw.reports),\n          };\n        })(),\n      };`,
'normalize portal user preorderAccess');

replaceOnce(
`    await savePortalUsers([...users, { id: crypto.randomUUID(), name, username: name.toLowerCase(), passwordHash, role, admin: isAdmin, canLoadInventory: isAdmin, active: true, pageAccess: {} }]);`,
`    await savePortalUsers([...users, { id: crypto.randomUUID(), name, username: name.toLowerCase(), passwordHash, role, admin: isAdmin, canLoadInventory: isAdmin, active: true, pageAccess: {}, preorderAccess: { manage: false, eta: false, safetyBuffer: false, notifications: false, reports: false } }]);`,
'new portal user defaults');

replaceOnce(
`    if (form.has("pageAccess")) {\n      try { updated.pageAccess = JSON.parse(String(form.get("pageAccess"))); } catch { /* keep existing */ }\n    }\n    if (form.has("canLoadInventory"))`,
`    if (form.has("pageAccess")) {\n      try { updated.pageAccess = JSON.parse(String(form.get("pageAccess"))); } catch { /* keep existing */ }\n    }\n    if (form.has("preorderAccess")) {\n      try {\n        const raw = JSON.parse(String(form.get("preorderAccess"))) as Record<string, unknown>;\n        updated.preorderAccess = {\n          manage: Boolean(raw.manage),\n          eta: Boolean(raw.eta),\n          safetyBuffer: Boolean(raw.safetyBuffer),\n          notifications: Boolean(raw.notifications),\n          reports: Boolean(raw.reports),\n        };\n      } catch { /* keep existing */ }\n    }\n    if (form.has("canLoadInventory"))`,
'update portal user preorderAccess');

replaceOnce(
`    pageAccess: {},\n  };`,
`    pageAccess: {},\n    preorderAccess: { manage: true, eta: true, safetyBuffer: true, notifications: true, reports: true },\n  };`,
'superadmin preorder defaults');

replaceOnce(
`  const [pageAccess, setPageAccess] = useState<Record<string, boolean>>(user.pageAccess ?? {});\n  const [canLoadInventory, setCanLoadInventory] = useState(user.canLoadInventory);`,
`  const [pageAccess, setPageAccess] = useState<Record<string, boolean>>(user.pageAccess ?? {});\n  const [preorderAccess, setPreorderAccess] = useState(() => ({\n    manage: Boolean(user.preorderAccess?.manage),\n    eta: Boolean(user.preorderAccess?.eta),\n    safetyBuffer: Boolean(user.preorderAccess?.safetyBuffer),\n    notifications: Boolean(user.preorderAccess?.notifications),\n    reports: Boolean(user.preorderAccess?.reports),\n  }));\n  const [canLoadInventory, setCanLoadInventory] = useState(user.canLoadInventory);`,
'user edit preorder state');

replaceOnce(
`      {showPageAccess && (\n        <div>\n          <div style={{ fontSize: 12, fontWeight: 700, color: "#374151", marginBottom: 8 }}>Page access</div>\n          <div style={{ display: "grid", gridTemplateColumns: "repeat(auto-fill, minmax(160px, 1fr))", gap: 6 }}>\n            {navItems.map((item) => (\n              <label key={item.id} style={{ display: "flex", alignItems: "center", gap: 7, fontSize: 13, cursor: "pointer", padding: "4px 0" }}>\n                <input\n                  type="checkbox"\n                  checked={Boolean(pageAccess[item.id])}\n                  onChange={(e) => { setJustSaved(false); setPageAccess((p) => ({ ...p, [item.id]: e.target.checked })); }}\n                />\n                {item.label}\n              </label>\n            ))}\n          </div>\n        </div>\n      )}\n\n      <div style={{ display: "flex", alignItems: "center", gap: 12 }}>`,
`      {showPageAccess && (\n        <div>\n          <div style={{ fontSize: 12, fontWeight: 700, color: "#374151", marginBottom: 8 }}>Page access</div>\n          <div style={{ display: "grid", gridTemplateColumns: "repeat(auto-fill, minmax(160px, 1fr))", gap: 6 }}>\n            {navItems.map((item) => (\n              <label key={item.id} style={{ display: "flex", alignItems: "center", gap: 7, fontSize: 13, cursor: "pointer", padding: "4px 0" }}>\n                <input\n                  type="checkbox"\n                  checked={Boolean(pageAccess[item.id])}\n                  onChange={(e) => { setJustSaved(false); setPageAccess((p) => ({ ...p, [item.id]: e.target.checked })); }}\n                />\n                {item.label}\n              </label>\n            ))}\n          </div>\n        </div>\n      )}\n\n      {showPageAccess && (\n        <div style={{ borderTop: "1px solid #e5e7eb", paddingTop: 10 }}>\n          <div style={{ fontSize: 12, fontWeight: 700, color: "#374151", marginBottom: 3 }}>Pre-order permissions</div>\n          <div style={{ fontSize: 11, color: "#9ca3af", marginBottom: 8 }}>These controls apply inside the Pre-orders page. Page access above still controls whether the user can open Pre-orders.</div>\n          <div style={{ display: "grid", gridTemplateColumns: "repeat(auto-fill, minmax(210px, 1fr))", gap: 6 }}>\n            {[\n              ["manage", "Enable / pause pre-orders"],\n              ["eta", "Change expected dispatch date"],\n              ["safetyBuffer", "Change safety buffer"],\n              ["notifications", "Approve customer notifications"],\n              ["reports", "View pre-order reports"],\n            ].map(([key, label]) => (\n              <label key={key} style={{ display: "flex", alignItems: "center", gap: 7, fontSize: 13, cursor: "pointer", padding: "4px 0" }}>\n                <input\n                  type="checkbox"\n                  checked={Boolean(preorderAccess[key as keyof typeof preorderAccess])}\n                  onChange={(e) => { setJustSaved(false); setPreorderAccess((p) => ({ ...p, [key]: e.target.checked })); }}\n                />\n                {label}\n              </label>\n            ))}\n          </div>\n        </div>\n      )}\n\n      <div style={{ display: "flex", alignItems: "center", gap: 12 }}>`,
'user edit preorder permission UI');

replaceOnce(
`            const fields: Record<string, string> = { intent: "update_portal_user", userId: user.id, name, role, canLoadInventory: canLoadInventory ? "on" : "off", canSeeProductStatus: canSeeProductStatus ? "on" : "off", pageAccess: JSON.stringify(pageAccess) };`,
`            const fields: Record<string, string> = { intent: "update_portal_user", userId: user.id, name, role, canLoadInventory: canLoadInventory ? "on" : "off", canSeeProductStatus: canSeeProductStatus ? "on" : "off", pageAccess: JSON.stringify(pageAccess), preorderAccess: JSON.stringify(preorderAccess) };`,
'submit preorderAccess');

fs.writeFileSync(path, text);
console.log('Applied portal preorder permissions into Settings → Users');
