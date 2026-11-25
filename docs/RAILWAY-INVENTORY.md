# Railway workspace inventory

Verified 25 November 2026 through Railway project, service and deployment reads. Deployment status reflects Railway and does not establish application health.

| Project | Environment | Service | Source | Deployment |
| --- | --- | --- | --- | --- |
| ashas-first-game | production | web | Source not reported by Railway | SUCCESS |
| ashas-first-game | production | Postgres | ghcr.io/railwayapp-templates/postgres-ssl:18 | SUCCESS |
| Workshop Portal | production | Workshop-portal- | NaveenKumarAustralia/Workshop-portal- (main) | SUCCESS |
| bundle builder | production | Postgres | ghcr.io/railwayapp-templates/postgres-ssl:18 | SUCCESS |
| bundle builder | production | bundle-builder-app | NaveenKumarAustralia/bundle-builder-app (main) | SUCCESS |
| PIckPackV1 | production | Postgres | ghcr.io/railwayapp-templates/postgres-ssl:18 | SUCCESS |
| PIckPackV1 | production | PickPackV1 | NaveenKumarAustralia/PickPackV1 (main) | SUCCESS |
| Collection sort V5 | production | Collection rearange  | NaveenKumarAustralia/collection-sort-v5 (main) | SUCCESS |
| Collection sort V5 | production | Postgres | ghcr.io/railwayapp-templates/postgres-ssl:18 | SUCCESS |
| Product demand | mcp-test | Postgres | ghcr.io/railwayapp-templates/postgres-ssl:18 | SUCCESS |
| Product demand | mcp-test | product-demand | NaveenKumarAustralia/product-demand (chatgpt/mcp-portal-upgrade) | SUCCESS |
| Product demand | production | Postgres | ghcr.io/railwayapp-templates/postgres-ssl:18 | SUCCESS |
| Product demand | production | product-demand | NaveenKumarAustralia/product-demand (main) | SUCCESS |
| celebrated-enchantment | production | Postgres | ghcr.io/railwayapp-templates/postgres-ssl:18 | SUCCESS |
| Inventory load restock | production | Postgres | ghcr.io/railwayapp-templates/postgres-ssl:18 | SUCCESS |
| Inventory load restock | production | inventory-load-existing-restock | NaveenKumarAustralia/inventory-load-existing-restock (main) | SUCCESS |
| Aisling Equation | production | Postgres | ghcr.io/railwayapp-templates/postgres-ssl:18 | SUCCESS |
| Aisling Equation | production | Aisling equation project | NaveenKumarAustralia/aisling-equation (main) | SUCCESS |
| Ecommerce Dashboard | production | Ecommerce-Dashboard | NaveenKumarAustralia/Ecommerce-Dashboard (main) | SUCCESS |
| Help Desk | production | Postgres | ghcr.io/railwayapp-templates/postgres-ssl:18 | SUCCESS |
| Help Desk | production | Help-desk | NaveenKumarAustralia/Help-desk (main) | SUCCESS |
| Size Guide Recommender | production | Postgres | ghcr.io/railwayapp-templates/postgres-ssl:18 | SUCCESS |
| Size Guide Recommender | production | size-guide-recommender | NaveenKumarAustralia/size-guide-recommender (main) | SUCCESS |

The Railway connection exposes deployment listings, logs, metrics, configuration and redeploy operations. GitHub source inspection was verified for Product demand. This inventory does not exercise writes, rollback, or database connections on unrelated projects. Database and volume mount details are available through Railway service inspection. Keep variable values and credentials out of reports.
