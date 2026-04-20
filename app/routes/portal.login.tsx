import type { ActionFunctionArgs, LoaderFunctionArgs } from "react-router";
import { Form, redirect, useActionData, useNavigation } from "react-router";
import prisma from "../db.server";
import { getSession, commitSession, verifyPassword } from "../portal.session.server";

export const loader = async ({ request }: LoaderFunctionArgs) => {
  const session = await getSession(request.headers.get("Cookie"));
  if (session.get("supplierId")) throw redirect("/portal");
  return null;
};

export const action = async ({ request }: ActionFunctionArgs) => {
  const form = await request.formData();
  const email = String(form.get("email") ?? "").trim().toLowerCase();
  const password = String(form.get("password") ?? "");

  if (!email || !password) {
    return { error: "Email and password are required" };
  }

  const account = await prisma.supplierAccount.findFirst({
    where: { email, active: true },
  });

  if (!account || !verifyPassword(password, account.passwordHash)) {
    return { error: "Invalid email or password" };
  }

  const session = await getSession(request.headers.get("Cookie"));
  session.set("supplierId", account.id);
  session.set("supplierName", account.name);
  session.set("supplierShop", account.shop);

  throw redirect("/portal", {
    headers: { "Set-Cookie": await commitSession(session) },
  });
};

export default function PortalLogin() {
  const actionData = useActionData<typeof action>();
  const navigation = useNavigation();
  const submitting = navigation.state === "submitting";

  return (
    <div style={styles.page}>
      <div style={styles.card}>
        <h1 style={styles.title}>Supplier Portal</h1>
        <p style={styles.subtitle}>Sign in to view and manage your orders</p>

        <Form method="post" style={styles.form}>
          <div style={styles.field}>
            <label style={styles.label}>Email</label>
            <input
              name="email"
              type="email"
              autoComplete="email"
              required
              style={styles.input}
            />
          </div>
          <div style={styles.field}>
            <label style={styles.label}>Password</label>
            <input
              name="password"
              type="password"
              autoComplete="current-password"
              required
              style={styles.input}
            />
          </div>

          {actionData?.error && (
            <div style={styles.error}>{actionData.error}</div>
          )}

          <button type="submit" disabled={submitting} style={styles.button}>
            {submitting ? "Signing in…" : "Sign in"}
          </button>
        </Form>
      </div>
    </div>
  );
}

const styles: Record<string, React.CSSProperties> = {
  page: {
    minHeight: "100vh",
    background: "#f4f6f8",
    display: "flex",
    alignItems: "center",
    justifyContent: "center",
    fontFamily: "-apple-system, BlinkMacSystemFont, 'Segoe UI', sans-serif",
  },
  card: {
    background: "#fff",
    borderRadius: 12,
    padding: "40px 48px",
    width: "100%",
    maxWidth: 400,
    boxShadow: "0 2px 16px rgba(0,0,0,0.08)",
  },
  title: { margin: 0, fontSize: 24, fontWeight: 700, color: "#1a1a1a" },
  subtitle: { margin: "8px 0 28px", fontSize: 14, color: "#6b7280" },
  form: { display: "flex", flexDirection: "column", gap: 16 },
  field: { display: "flex", flexDirection: "column", gap: 6 },
  label: { fontSize: 14, fontWeight: 500, color: "#374151" },
  input: {
    padding: "10px 12px",
    borderRadius: 8,
    border: "1px solid #d1d5db",
    fontSize: 15,
    outline: "none",
  },
  error: {
    background: "#fef2f2",
    border: "1px solid #fecaca",
    borderRadius: 8,
    padding: "10px 14px",
    fontSize: 14,
    color: "#b91c1c",
  },
  button: {
    marginTop: 8,
    padding: "12px",
    background: "#008060",
    color: "#fff",
    border: "none",
    borderRadius: 8,
    fontSize: 15,
    fontWeight: 600,
    cursor: "pointer",
  },
};
