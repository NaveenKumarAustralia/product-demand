import type { ActionFunctionArgs, LoaderFunctionArgs } from "react-router";
import { Form, useActionData, useLoaderData, useNavigation } from "react-router";
import {
  Page,
  Layout,
  Card,
  DataTable,
  Button,
  TextField,
  FormLayout,
  Banner,
  Badge,
  EmptyState,
  BlockStack,
  InlineStack,
  Text,
} from "@shopify/polaris";
import { useState } from "react";
import { authenticate } from "../shopify.server";
import prisma from "../db.server";
import { hashPassword } from "../portal.session.server";

export const loader = async ({ request }: LoaderFunctionArgs) => {
  const { session } = await authenticate.admin(request);
  const shop = session.shop;

  const suppliers = await prisma.supplierAccount.findMany({
    where: { shop },
    orderBy: { name: "asc" },
    select: { id: true, name: true, email: true, active: true, createdAt: true },
  });

  return { shop, suppliers };
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

    if (!name || !email || !password) {
      return { error: "All fields are required", intent: "create" };
    }
    if (password.length < 8) {
      return { error: "Password must be at least 8 characters", intent: "create" };
    }

    const existing = await prisma.supplierAccount.findUnique({
      where: { shop_email: { shop, email } },
    });
    if (existing) {
      return { error: "A supplier with this email already exists", intent: "create" };
    }

    await prisma.supplierAccount.create({
      data: { shop, name, email, passwordHash: hashPassword(password) },
    });

    return { success: `Supplier account created for ${name}`, intent: "create" };
  }

  if (intent === "toggle") {
    const id = Number(form.get("id"));
    const account = await prisma.supplierAccount.findFirst({ where: { id, shop } });
    if (account) {
      await prisma.supplierAccount.update({
        where: { id },
        data: { active: !account.active },
      });
    }
    return null;
  }

  if (intent === "delete") {
    const id = Number(form.get("id"));
    await prisma.supplierAccount.deleteMany({ where: { id, shop } });
    return null;
  }

  if (intent === "reset_password") {
    const id = Number(form.get("id"));
    const password = String(form.get("new_password") ?? "");
    if (password.length < 8) {
      return { error: "Password must be at least 8 characters", intent: "reset" };
    }
    await prisma.supplierAccount.updateMany({
      where: { id, shop },
      data: { passwordHash: hashPassword(password) },
    });
    return { success: "Password updated", intent: "reset" };
  }

  return null;
};

export default function SuppliersPage() {
  const { suppliers } = useLoaderData<typeof loader>();
  const actionData = useActionData<typeof action>();
  const navigation = useNavigation();
  const submitting = navigation.state === "submitting";

  const [name, setName] = useState("");
  const [email, setEmail] = useState("");
  const [password, setPassword] = useState("");

  const portalUrl = "https://product-demand-production.up.railway.app/portal";

  return (
    <Page
      title="Supplier Accounts"
      subtitle={`Portal login: ${portalUrl}`}
    >
      <Layout>
        <Layout.Section>
          {actionData?.success && (
            <Banner tone="success" onDismiss={() => {}}>
              {actionData.success}
            </Banner>
          )}
        </Layout.Section>

        <Layout.Section>
          <Card>
            <BlockStack gap="400">
              <Text variant="headingMd" as="h2">Create Supplier Account</Text>
              <Text tone="subdued" as="p">
                Share the portal URL and login credentials with your supplier.
              </Text>
              <Form method="post">
                <input type="hidden" name="intent" value="create" />
                <FormLayout>
                  <FormLayout.Group>
                    <TextField
                      label="Supplier name"
                      name="name"
                      value={name}
                      onChange={setName}
                      autoComplete="off"
                      helpText="Must match the supplier name used when placing orders"
                    />
                    <TextField
                      label="Email"
                      name="email"
                      type="email"
                      value={email}
                      onChange={setEmail}
                      autoComplete="off"
                    />
                  </FormLayout.Group>
                  <TextField
                    label="Password"
                    name="password"
                    type="password"
                    value={password}
                    onChange={setPassword}
                    autoComplete="new-password"
                    helpText="At least 8 characters"
                  />
                  {actionData?.intent === "create" && actionData?.error && (
                    <Banner tone="critical">{actionData.error}</Banner>
                  )}
                  <Button submit variant="primary" loading={submitting}>
                    Create Account
                  </Button>
                </FormLayout>
              </Form>
            </BlockStack>
          </Card>
        </Layout.Section>

        <Layout.Section>
          <Card>
            <BlockStack gap="400">
              <Text variant="headingMd" as="h2">Supplier Accounts</Text>

              {suppliers.length === 0 ? (
                <EmptyState
                  heading="No supplier accounts yet"
                  image=""
                >
                  <p>Create an account above to give a supplier access to the portal.</p>
                </EmptyState>
              ) : (
                <DataTable
                  columnContentTypes={["text", "text", "text", "text"]}
                  headings={["Name", "Email", "Status", "Actions"]}
                  rows={suppliers.map((s) => [
                    s.name,
                    s.email,
                    <Badge tone={s.active ? "success" : "enabled"}>
                      {s.active ? "Active" : "Disabled"}
                    </Badge>,
                    <InlineStack gap="200">
                      <Form method="post" style={{ display: "inline" }}>
                        <input type="hidden" name="intent" value="toggle" />
                        <input type="hidden" name="id" value={s.id} />
                        <Button submit size="slim" variant="plain">
                          {s.active ? "Disable" : "Enable"}
                        </Button>
                      </Form>
                      <ResetPasswordForm id={s.id} submitting={submitting} />
                      <Form method="post" style={{ display: "inline" }}>
                        <input type="hidden" name="intent" value="delete" />
                        <input type="hidden" name="id" value={s.id} />
                        <Button submit size="slim" variant="plain" tone="critical">
                          Delete
                        </Button>
                      </Form>
                    </InlineStack>,
                  ])}
                />
              )}
            </BlockStack>
          </Card>
        </Layout.Section>
      </Layout>
    </Page>
  );
}

function ResetPasswordForm({ id, submitting }: { id: number; submitting: boolean }) {
  const [show, setShow] = useState(false);
  const [pw, setPw] = useState("");

  if (!show) {
    return (
      <Button size="slim" variant="plain" onClick={() => setShow(true)}>
        Reset password
      </Button>
    );
  }

  return (
    <Form method="post" style={{ display: "inline-flex", gap: 8, alignItems: "center" }}>
      <input type="hidden" name="intent" value="reset_password" />
      <input type="hidden" name="id" value={id} />
      <input
        name="new_password"
        type="password"
        placeholder="New password"
        value={pw}
        onChange={(e) => setPw(e.target.value)}
        style={{ padding: "4px 8px", borderRadius: 6, border: "1px solid #ccc", fontSize: 13 }}
      />
      <Button submit size="slim" loading={submitting}>Save</Button>
      <Button size="slim" variant="plain" onClick={() => setShow(false)}>Cancel</Button>
    </Form>
  );
}
