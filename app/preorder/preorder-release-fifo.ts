// Pure FIFO release picker — no DB/Shopify imports so it's unit-testable in
// isolation (the server module that uses it pulls in prisma, which the node test
// runner can't resolve). Given the still-waiting orders (each with the
// (batch,variant,market) groups its lines draw on) and a per-group budget of spare
// units, pick which whole orders to release — OLDEST reservation first. An order is
// only released when EVERY group it needs has enough budget (never half-ship a
// multi-line / multi-batch order). Budgets are consumed as orders are picked, so
// one returned/spare unit releases exactly one waiting order.
export type ReleasableOrder = {
  orderId: string;
  oldestReservedMs: number;
  rowIds: number[];
  needByGroup: Record<string, number>;
  readyTags: string[];
};

export function selectFifoReleasableOrders(orders: ReleasableOrder[], budgetByGroup: Map<string, number>): ReleasableOrder[] {
  const budget = new Map(budgetByGroup);
  const out: ReleasableOrder[] = [];
  const sorted = [...orders].sort((a, b) => a.oldestReservedMs - b.oldestReservedMs || a.orderId.localeCompare(b.orderId));
  for (const order of sorted) {
    const groups = Object.entries(order.needByGroup);
    if (!groups.every(([key, need]) => (budget.get(key) ?? 0) >= need)) continue;
    for (const [key, need] of groups) budget.set(key, (budget.get(key) ?? 0) - need);
    out.push(order);
  }
  return out;
}
