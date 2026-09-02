import fs from 'node:fs';
const path = 'app/portal-preorders.tsx';
let text = fs.readFileSync(path, 'utf8');
function replaceOnce(from, to, label) {
  const count = text.split(from).length - 1;
  if (count !== 1) throw new Error(`${label}: expected one match, found ${count}`);
  text = text.replace(from, to);
}
replaceOnce(
  'import { PreorderWebhookStatusPanel } from "./preorder/preorder-webhook-status-panel";\n',
  'import { PreorderWebhookStatusPanel } from "./preorder/preorder-webhook-status-panel";\nimport { PreorderShopifyReadinessPanel } from "./preorder/preorder-shopify-readiness-panel";\n',
  'readiness import',
);
replaceOnce(
  '        <>\n          <PreorderWebhookStatusPanel />\n          <PreorderSettingsPanel configuration={data.configuration} />\n        </>',
  '        <>\n          <PreorderShopifyReadinessPanel />\n          <PreorderWebhookStatusPanel />\n          <PreorderSettingsPanel configuration={data.configuration} />\n        </>',
  'settings readiness panel',
);
fs.writeFileSync(path, text);
console.log('Wired Shopify preorder readiness panel');
