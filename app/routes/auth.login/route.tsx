import { AppProvider } from "@shopify/shopify-app-react-router/react";
import { useState } from "react";
import type { ActionFunctionArgs, LoaderFunctionArgs } from "react-router";
import { Form, useActionData, useLoaderData } from "react-router";

import { login } from "../../shopify.server";
import { loginErrorMessage } from "./error.server";

export const loader = async ({ request }: LoaderFunctionArgs) => {
  const errors = loginErrorMessage(await login(request));

  return { errors };
};

export const action = async ({ request }: ActionFunctionArgs) => {
  const errors = loginErrorMessage(await login(request));

  return {
    errors,
  };
};

export default function Auth() {
  const loaderData = useLoaderData<typeof loader>();
  const actionData = useActionData<typeof action>();
  const [shop, setShop] = useState("");
  const { errors } = actionData || loaderData;

  // If rendered inside the Shopify admin iframe, accounts.shopify.com blocks
  // OAuth redirects via X-Frame-Options. Escape to the top-level window so the
  // OAuth flow can complete without being sandboxed.
  function handleSubmit(e: React.FormEvent<HTMLFormElement>) {
    if (typeof window !== "undefined" && window.top !== window.self) {
      e.preventDefault();
      const form = e.currentTarget;
      const data = new FormData(form);
      const shopValue = data.get("shop") as string;
      const params = new URLSearchParams({ shop: shopValue });
      window.top!.location.href = `/auth/login?${params.toString()}`;
    }
  }

  return (
    <AppProvider embedded={false}>
      <s-page>
        <Form method="post" onSubmit={handleSubmit}>
        <s-section heading="Log in">
          <s-text-field
            name="shop"
            label="Shop domain"
            details="example.myshopify.com"
            value={shop}
            onChange={(e) => setShop(e.currentTarget.value)}
            autocomplete="on"
            error={errors.shop}
          ></s-text-field>
          <s-button type="submit">Log in</s-button>
        </s-section>
        </Form>
      </s-page>
    </AppProvider>
  );
}
