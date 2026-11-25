/**
 * Minimal MCP server over Streamable HTTP (stateless, JSON responses).
 * Pure: the tool list and runner are injected, so it is unit-tested directly.
 *
 * Supports: initialize, notifications/*, ping, tools/list, tools/call.
 */

export const SUPPORTED_PROTOCOL_VERSIONS = ["2025-06-18", "2025-03-26", "2024-11-05"];
export const SERVER_INFO = { name: "karma-east-production-portal", title: "Karma East Production Portal", version: "1.0.0" };

export const SERVER_INSTRUCTIONS = [
  "You are connected to Karma East's Production Portal (production and reorder orders, stock, fabric, shipments).",
  "Read freely. Before any change: gather the facts with read tools (e.g. get_product_production_view), then call a preview_* tool.",
  "A preview never changes anything. Show Koku the preview summary and ask him to approve it explicitly.",
  "Call confirm_change only after Koku clearly says yes to that specific preview in this conversation. Never confirm on your own initiative, never batch-confirm, and never treat earlier approvals as covering new changes.",
  "If a preview is blocked as a possible duplicate, tell Koku and suggest amending the existing order instead.",
].join(" ");

export type McpToolSpec = {
  name: string;
  title: string;
  description: string;
  inputSchema: Record<string, unknown>;
  annotations: Record<string, unknown>;
};
export type McpToolOutcome = { structured: unknown; text?: string; isError?: boolean };

export type JsonRpcRequest = { jsonrpc?: string; id?: string | number | null; method?: string; params?: Record<string, unknown> };
type JsonRpcResponse = { jsonrpc: "2.0"; id: string | number | null; result?: unknown; error?: { code: number; message: string } };

export async function handleJsonRpc(
  message: JsonRpcRequest,
  deps: { tools: McpToolSpec[]; call: (name: string, args: Record<string, unknown>) => Promise<McpToolOutcome> },
): Promise<JsonRpcResponse | null> {
  const id = message?.id ?? null;
  const isNotification = message?.id === undefined || message?.id === null;
  const fail = (code: number, msg: string): JsonRpcResponse => ({ jsonrpc: "2.0", id, error: { code, message: msg } });

  if (!message || message.jsonrpc !== "2.0" || typeof message.method !== "string") {
    return isNotification ? null : fail(-32600, "Invalid request");
  }
  const method = message.method;
  if (method.startsWith("notifications/")) return null;

  switch (method) {
    case "initialize": {
      const asked = String(message.params?.protocolVersion ?? "");
      const protocolVersion = SUPPORTED_PROTOCOL_VERSIONS.includes(asked) ? asked : SUPPORTED_PROTOCOL_VERSIONS[0];
      return {
        jsonrpc: "2.0",
        id,
        result: { protocolVersion, capabilities: { tools: { listChanged: false } }, serverInfo: SERVER_INFO, instructions: SERVER_INSTRUCTIONS },
      };
    }
    case "ping":
      return { jsonrpc: "2.0", id, result: {} };
    case "tools/list":
      return { jsonrpc: "2.0", id, result: { tools: deps.tools } };
    case "tools/call": {
      const name = String(message.params?.name ?? "");
      const args = (message.params?.arguments ?? {}) as Record<string, unknown>;
      if (!deps.tools.some((t) => t.name === name)) return fail(-32602, `Unknown tool: ${name}`);
      if (typeof args !== "object" || Array.isArray(args)) return fail(-32602, "arguments must be an object");
      const out = await deps.call(name, args);
      const text = out.text ?? JSON.stringify(out.structured, null, 2);
      const structured = out.structured && typeof out.structured === "object" && !Array.isArray(out.structured) ? out.structured : { result: out.structured };
      return { jsonrpc: "2.0", id, result: { content: [{ type: "text", text }], structuredContent: structured, isError: Boolean(out.isError) } };
    }
    default:
      return isNotification ? null : fail(-32601, `Method not found: ${method}`);
  }
}
