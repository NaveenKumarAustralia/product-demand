# ChatGPT connector (MCP) for the Production Portal

The portal exposes a remote MCP server so ChatGPT can read production data and
prepare changes. Every change needs Koku's explicit approval in chat.

- MCP URL: `https://product-demand-production.up.railway.app/mcp`
- Auth: OAuth 2.1 with dynamic client registration and PKCE. Approving needs a
  Production Portal **admin** login. The connection stays signed in (refresh
  tokens rotate and last 365 days from last use) until revoked.

## Connect in ChatGPT
1. ChatGPT → Settings → Apps & Connectors → Advanced settings → turn on **Developer mode**.
2. Settings → Apps & Connectors → **Create** (or "Add custom connector").
3. Name: `Karma East Production Portal`. MCP server URL: the URL above.
   Authentication: **OAuth**. Leave client id/secret empty (ChatGPT registers itself).
4. Click Create → a portal page opens → sign in with your portal admin login → **Approve**.
5. In a chat, enable the connector from the tools menu.

## Tools
Read: `get_portal_options`, `search_products`, `list_production_orders`,
`get_production_order`, `get_product_production_view`, `reorder_planner_overview`,
`list_shipments`, `get_shipment`, `get_fabric_availability`,
`list_new_collections`, `get_new_collection`, `recent_activity`, `list_pending_changes`.

Write (always two steps): `preview_create_production_order`,
`preview_amend_order_quantities`, `preview_update_order`, `preview_add_order_note`
→ Koku approves → `confirm_change`. `cancel_change` discards a preview.

## Safeguards
- Previews change nothing. A change id applies once only; re-sending it does nothing.
- Creating an order is refused if the product already has an open order on order
  or in production, unless explicitly allowed (prevents duplicates). The check
  runs again at confirm time.
- At confirm time the order is re-read; if it changed since the preview, the
  change is refused and must be previewed again. Previews expire after 30 minutes.
- Writes run through the portal's own action handler as the portal user
  **ChatGPT** (admin role, no password, can't log in to the UI), so the normal
  validation, fabric consumption, cut log, pre-order alerts and activity log apply.
- `McpAuditLog` records every preview, applied change, refusal and failure with
  before/after values and the approving admin.
- Tokens and codes are stored as SHA-256 hashes only. Re-using an old refresh
  token revokes the connection. Deactivating or demoting the approving admin
  cuts access immediately.

## Revoke access
Remove the connector in ChatGPT, and/or deactivate the approving admin. To kill
all ChatGPT connections at once:
`UPDATE "McpOAuthToken" SET "revokedAt" = now() WHERE "revokedAt" IS NULL;`

## Settings (Railway variables)
- `MCP_PUBLIC_URL` (optional): public base URL, e.g. `https://product-demand-production.up.railway.app`.
- `MCP_ALLOWED_REDIRECT_HOSTS` (optional): default `chatgpt.com,chat.openai.com`.
- `PORTAL_BACKGROUND_JOBS=off`: disables pre-order jobs and webhook registration.
  They also never run outside the Railway `production` environment (test forks).

## Not covered yet (phase 2)
Fabric sheet edits, packing lists, collections edits and other portal
operations. Production lead times are not stored by the portal (the reorder
planner keeps them in the browser), so they're not readable.

## Capability matrix

Production commit `08c241a` contains PR #21. Tool names in the connected
ChatGPT client must be checked separately after each server upgrade.

| Module | Supported operations | Verification and gaps |
| --- | --- | --- |
| Product Information and measurements | Search products and read recorded style information and fabric overrides | PR #21 deployed. Product Information tool was absent from the connected client. Dedicated measurement and size-chart queries remain open. |
| Fabric and capacity | Read availability, recorded order consumption and selected-fabric capacity | Existing read tools connected. Full multi-fabric requirements, purchasing, allocation edits and stock writes remain open. |
| Production and reorder | List orders, read detail and planner, preview create, edit, quantity changes and cancellation | Existing tools connected. |
| Notes | Preview append, replace or clear notes and factory notes | Append connected. Replace and clear were deployed in PR #21 but absent from the connected client. |
| Deletion and restoration | Preview eligible uncut order deletion | Phase 2 carried forward. Restore is not registered. The portal UI's immediate undo uses a partial snapshot and creates a new ID. |
| Shipments and packing | Read shipment lists and detail | Existing tools connected. Packing-list edits and shipment writes remain open. Shipment quantities can overlap with ordered quantities. |
| Collections | List and read planning collections | Existing tools connected. Collection and collection order edits remain open. |
| Samples and Vision Board | List samples with latest iteration and list both versions of boards with item counts | Implemented in the integration branch. COPY and LIVE validation pending. Full details and writes remain open. |
| Sales, inventory and demand | Read planner aggregates | Separate reports and inventory writes remain open. |
| Costing, tracking and activity | Read recorded sample cost, order and shipment status, and recent activity | Full costing, production tracking and reporting workflows remain open. |
| Growth Engine | Read-only policy | No advertising writes or automation changes are authorized. |

The route audit also surfaced preorder operations, packing inventory loads,
collection media and tasks, product variant codes, supplier configuration,
photo shoots, and portal threads. These need operation-specific review before
registration. Some write paths call Shopify or notification services.

## Test and release

Generate Prisma, then run `npm run typecheck`, `npm run test:mcp`,
`npm run test:preorder` and `npm run build`. Before write tests in Railway
`mcp-test`, verify the application's database reference points to its own
Postgres service and that any external store and notification integrations
are isolated or disabled. A distinct Railway environment alone is insufficient.
Use synthetic records to test OAuth, discovery, preview, approval, stale
previews, duplicate confirmation and persisted results.

Deploy the verified branch to COPY, then merge the tested change to `main`.
Check the Railway deployment commit and status, the live application's health
and the authenticated MCP tool list. Test reads via the existing LIVE
connector. If it reports an older tool list, refresh it through ChatGPT's
custom connector settings, then verify that the newly registered tools can
be called. For rollback, select the last successful Railway deployment and
check compatibility with any database migrations before restoring it.

See [Railway workspace inventory](RAILWAY-INVENTORY.md).
