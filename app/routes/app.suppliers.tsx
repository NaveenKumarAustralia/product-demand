import type { ActionFunctionArgs, LoaderFunctionArgs } from "react-router";
import { Form, useActionData, useLoaderData, useNavigation } from "react-router";
import { useState } from "react";
import { authenticate } from "../shopify.server";
import prisma from "../db.server";
import { hashPassword } from "../portal.session.server";

export const loader = async ({ request }: LoaderFunctionArgs) => {
  const { session } = await authenticate.admin(request);
  const suppliers = await prisma.supplierAccount.findMany({
    where: { shop: session.shop },
    orderBy: { name: "asc" },
    select: { id: true, name: true, email: true, active: true, createdAt: true },
  });
  return { suppliers };
};

export const action = async ({ request }: ActionFunctionArgs) => {
  const { session } = await authenticate.admin(request);
  const shop = session.shop;
  const form = await request.formData();
  const intent = form.get("intent");

  if (intent === "create") {
    const name = String(form.get("name") ?? "").trim();
    const email = String(form.get("email") ?? "").trim().toLowerCase();
    const password = String(form.get("password") ?? "");
    if (!name || !email || !password) return { error: "All fields are required" };
    if (password.length < 8) return { error: "Password must be at least 8 characters" };
    const existing = await prisma.supplierAccount.findUnique({ where: { shop_email: { shop, email } } });
    if (existing) return { error: "A supplier with this email already exists" };
    await prisma.supplierAccount.create({ data: { shop, name, email, passwordHash: hashPassword(password) } });
    return { success: `Account created for ${name}` };
  }

  if (intent === "toggle") {
    const id = Number(form.get("id"));
    const acc = await prisma.supplierAccount.findFirst({ where: { id, shop } });
    if (acc) await prisma.supplierAccount.update({ where: { id }, data: { active: !acc.active } });
  }

  if (intent === "delete") {
    const id = Number(form.get("id"));
    await prisma.supplierAccount.deleteMany({ where: { id, shop } });
  }

  if (intent === "reset_password") {
    const id = Number(form.get("id"));
    const password = String(form.get("new_password") ?? "");
    if (password.length < 8) return { error: "Password must be at least 8 characters" };
    await prisma.supplierAccount.updateMany({ where: { id, shop }, data: { passwordHash: hashPassword(password) } });
    return { success: "Password updated" };
  }

  return null;
};

const PORTAL_URL = "https://product-demand-production.up.railway.app/portal";

export default function SuppliersPage() {
  const { suppliers } = useLoaderData<typeof loader>();
  const actionData = useActionData<typeof action>();
  const navigation = useNavigation();
  const submitting = navigation.state === "submitting";

  const [name, setName] = useState("");
  const [email, setEmail] = useState("");
  const [password, setPassword] = useState("");

  return (
    <div style={s.page}>
      <h1 style={s.pageTitle}>Supplier Accounts</h1>
      <p style={s.pageSubtitle}>
        Portal URL: <a href={PORTAL_URL} target="_blank" rel="noreferrer" style={s.link}>{PORTAL_URL}</a>
      </p>

      {actionData?.success && <div style={s.successBanner}>{actionData.success}</div>}
      {actionData?.error && <div style={s.errorBanner}>{actionData.error}</div>}

      <div style={s.card}>
        <h2 style={s.cardTitle}>Create Supplier Account</h2>
        <p style={s.hint}>Supplier name must match exactly what you type when placing orders.</p>
        <Form method="post" style={s.form}>
          <input type="hidden" name="intent" value="create" />
          <div style={s.row}>
            <div style={s.field}>
              <label style={s.label}>Supplier name</label>
              <input name="name" value={name} onChange={(e) => setName(e.target.value)} required style={s.input} />
            </div>
            <div style={s.field}>
              <label style={s.label}>Email</label>
              <input name="email" type="email" value={email} onChange={(e) => setEmail(e.target.value)} required style={s.input} />
            </div>
            <div style={s.field}>
              <label style={s.label}>Password (min 8 chars)</label>
              <input name="password" type="password" value={password} onChange={(e) => setPassword(e.target.value)} required style={s.input} />
            </div>
          </div>
          <button type="submit" disabled={submitting} style={s.btn}>Create Account</button>
        </Form>
      </div>

      <div style={s.card}>
        <h2 style={s.cardTitle}>Accounts</h2>
        {suppliers.length === 0 ? (
          <p style={{ color: "#6b7280", fontSize: 14 }}>No supplier accounts yet.</p>
        ) : (
          <table style={s.table}>
            <thead>
              <tr>
                <th style={s.th}>Name</th>
                <th style={s.th}>Email</th>
                <th style={s.th}>Status</th>
                <th style={s.th}>Actions</th>
              </tr>
            </thead>
            <tbody>
              {suppliers.map((sup) => (
                <SupplierRow key={sup.id} sup={sup} submitting={submitting} />
              ))}
            </tbody>
          </table>
        )}
      </div>
    </div>
  );
}

function SupplierRow({ sup, submitting }: { sup: { id: number; name: string; email: string; active: boolean }; submitting: boolean }) {
  const [showReset, setShowReset] = useState(false);
  const [newPw, setNewPw] = useState("");

  return (
    <tr>
      <td style={s.td}>{sup.name}</td>
      <td style={s.td}>{sup.email}</td>
      <td style={s.td}>
        <span style={{ ...s.badge, background: sup.active ? "#d1fae5" : "#f3f4f6", color: sup.active ? "#065f46" : "#6b7280" }}>
          {sup.active ? "Active" : "Disabled"}
        </span>
      </td>
      <td style={{ ...s.td, display: "flex", gap: 8, flexWrap: "wrap", alignItems: "center" }}>
        <Form method="post" style={{ display: "inline" }}>
          <input type="hidden" name="intent" value="toggle" />
          <input type="hidden" name="id" value={sup.id} />
          <button type="submit" style={s.linkBtn}>{sup.active ? "Disable" : "Enable"}</button>
        </Form>

        {!showReset ? (
          <button style={s.linkBtn} onClick={() => setShowReset(true)}>Reset password</button>
        ) : (
          <Form method="post" style={{ display: "inline-flex", gap: 6, alignItems: "center" }}>
            <input type="hidden" name="intent" value="reset_password" />
            <input type="hidden" name="id" value={sup.id} />
            <input type="password" name="new_password" placeholder="New password" value={newPw} onChange={(e) => setNewPw(e.target.value)} style={{ ...s.input, width: 140, padding: "4px 8px" }} />
            <button type="submit" disabled={submitting} style={s.linkBtn}>Save</button>
            <button type="button" style={s.linkBtn} onClick={() => setShowReset(false)}>Cancel</button>
          </Form>
        )}

        <Form method="post" style={{ display: "inline" }}>
          <input type="hidden" name="intent" value="delete" />
          <input type="hidden" name="id" value={sup.id} />
          <button type="submit" style={{ ...s.linkBtn, color: "#b91c1c" }}>Delete</button>
        </Form>
      </td>
    </tr>
  );
}

const s: Record<string, React.CSSProperties> = {
  page: { padding: "32px 24px", maxWidth: 900, margin: "0 auto", fontFamily: "-apple-system, BlinkMacSystemFont, 'Segoe UI', sans-serif" },
  pageTitle: { fontSize: 24, fontWeight: 700, margin: "0 0 4px", color: "#111827" },
  pageSubtitle: { fontSize: 14, color: "#6b7280", margin: "0 0 24px" },
  link: { color: "#008060" },
  successBanner: { background: "#d1fae5", color: "#065f46", border: "1px solid #a7f3d0", borderRadius: 8, padding: "12px 16px", marginBottom: 20, fontSize: 14 },
  errorBanner: { background: "#fef2f2", color: "#b91c1c", border: "1px solid #fecaca", borderRadius: 8, padding: "12px 16px", marginBottom: 20, fontSize: 14 },
  card: { background: "#fff", borderRadius: 12, padding: 24, boxShadow: "0 1px 4px rgba(0,0,0,0.06)", marginBottom: 24 },
  cardTitle: { fontSize: 17, fontWeight: 700, margin: "0 0 8px", color: "#111827" },
  hint: { fontSize: 13, color: "#6b7280", margin: "0 0 16px" },
  form: { display: "flex", flexDirection: "column", gap: 16 },
  row: { display: "flex", gap: 16, flexWrap: "wrap" },
  field: { display: "flex", flexDirection: "column", gap: 6, flex: 1, minWidth: 180 },
  label: { fontSize: 13, fontWeight: 500, color: "#374151" },
  input: { padding: "9px 12px", borderRadius: 8, border: "1px solid #d1d5db", fontSize: 14, outline: "none" },
  btn: { alignSelf: "flex-start", padding: "10px 20px", background: "#008060", color: "#fff", border: "none", borderRadius: 8, fontSize: 14, fontWeight: 600, cursor: "pointer" },
  table: { width: "100%", borderCollapse: "collapse" },
  th: { textAlign: "left", fontSize: 12, fontWeight: 600, color: "#6b7280", textTransform: "uppercase", letterSpacing: "0.05em", padding: "8px 12px", borderBottom: "1px solid #f3f4f6" },
  td: { padding: "12px 12px", fontSize: 14, color: "#374151", borderBottom: "1px solid #f9fafb" },
  badge: { padding: "3px 10px", borderRadius: 20, fontSize: 12, fontWeight: 600 },
  linkBtn: { background: "none", border: "none", color: "#008060", fontSize: 13, cursor: "pointer", padding: 0, textDecoration: "underline" },
};
