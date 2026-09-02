import fs from 'node:fs';
const path = 'app/portal-preorders.tsx';
let text = fs.readFileSync(path, 'utf8');
function replaceOnce(from, to, label) {
  const count = text.split(from).length - 1;
  if (count !== 1) throw new Error(`${label}: expected one match, found ${count}`);
  text = text.replace(from, to);
}
replaceOnce(
`      <div style={s.controlActions}>
        <button
          type="button"
          style={s.secondaryButton}
          disabled={busy}
          onClick={() => onManage({
            operation: "update-settings",
            shipDate,
            safetyBufferPercent: bufferPercent,
            safetyBufferQty: bufferQty,
          }, "Preorder batch settings saved.")}
        >
          {busy ? "Saving…" : "Save settings"}
        </button>
        <button
          type="button"
          style={{ ...s.primaryButton, ...(batch.enabled ? s.pauseButton : {}) }}
          disabled={busy || (!batch.enabled && !canActivate)}
          title={!batch.enabled && !canActivate ? "Batch must be On Production and assigned to AUS or USA first." : undefined}
          onClick={() => onManage({ operation: "set-enabled", enabled: !batch.enabled }, batch.enabled ? "Preorder paused." : "Preorder enabled.")}
        >
          {busy ? "Working…" : batch.enabled ? "Pause preorder" : "Enable preorder"}
        </button>
      </div>`,
`      <div style={s.controlActions}>
        <button
          type="button"
          style={s.secondaryButton}
          disabled={busy}
          onClick={() => onManage({
            operation: "update-settings",
            shipDate,
            safetyBufferPercent: bufferPercent,
            safetyBufferQty: bufferQty,
          }, "Preorder batch settings saved.")}
        >
          {busy ? "Saving…" : "Save settings"}
        </button>
        <button
          type="button"
          style={{ ...s.primaryButton, ...(batch.enabled ? s.pauseButton : {}) }}
          disabled={busy || (!batch.enabled && !canActivate)}
          title={!batch.enabled && !canActivate ? "Batch must be On Production and assigned to AUS or USA first." : undefined}
          onClick={() => onManage({ operation: "set-enabled", enabled: !batch.enabled }, batch.enabled ? "Preorder paused and removed from Shopify." : "Preorder enabled internally. Activate it on Shopify when ready.")}
        >
          {busy ? "Working…" : batch.enabled ? "Pause preorder" : "Enable preorder"}
        </button>
      </div>
      <div style={{ ...s.controlActions, marginTop: 10, paddingTop: 10, borderTop: "1px solid #e2e8f0" }}>
        <div style={{ fontSize: 13, fontWeight: 700, color: batch.shopifySellingPlanActive ? "#166534" : "#64748b" }}>
          Shopify: {batch.shopifySellingPlanActive ? "Live preorder selling plan" : "Not active"}
        </div>
        {batch.shopifySellingPlanActive ? (
          <button
            type="button"
            style={s.pauseButton}
            disabled={busy}
            onClick={() => onManage({ operation: "remove-shopify" }, "Shopify preorder selling plan removed.")}
          >
            {busy ? "Working…" : "Remove from Shopify"}
          </button>
        ) : (
          <button
            type="button"
            style={s.primaryButton}
            disabled={busy || !batch.enabled || !batch.eligible || batch.totalAvailable <= 0}
            title={!batch.enabled ? "Enable this preorder batch first." : batch.totalAvailable <= 0 ? "No preorder capacity is currently available." : undefined}
            onClick={() => onManage({ operation: "activate-shopify" }, "Shopify preorder selling plan activated.")}
          >
            {busy ? "Working…" : "Activate on Shopify"}
          </button>
        )}
      </div>`,
  'Shopify activation controls',
);
fs.writeFileSync(path, text);
console.log('Wired preorder Shopify activation controls');
