# Fabric Planner Notes

## Context

The app currently has:

- Existing Products Restock
- Fabric in Stock
- Packing Lists

Fabric in Stock was originally imported from an Excel/Google Sheet, then unlinked so future fabric stock changes can be managed manually inside the app.

The next major improvement being explored is a better way to plan garment production, fabric usage, and costing without relying on Google Sheets.

## Problem

The business has around 300 garment styles and keeps adding more. Each garment style can be made in different fabrics/prints, and the same style can use a different number of meters depending on the fabric.

For example:

- 1200 meters of a fabric in stock
- A style may use 3 meters per garment, allowing 400 pieces
- The same style in another fabric might use 2.5 meters per garment

Each garment also has extra costs:

- fabric cost per meter
- stitching cost
- button cost
- zip cost
- lining/trim meters
- other trim costs
- optional wastage/buffer percentage, for example 5% added to fabric cost

The current spreadsheet method is slow and easy to lose track of.

## Preferred Direction

Start from garment styles instead of starting from fabric.

Suggested flow:

1. Garment category tiles
2. Garment style tiles inside a category
3. Style detail/costing matrix
4. Link planned quantities to Existing Products Restock and reserve/deduct fabric

## Page 1: Garment Category Tiles

Possible categories:

- Short Sleeve Dresses
- Long Sleeve Dresses
- Short Sleeve Tops
- Long Sleeve Tops
- Mid Length Skirts
- Long Skirts
- Pants
- Jackets

Each tile could eventually show:

- number of styles inside
- open planned orders
- total fabric meters reserved
- missing costing info warnings

## Page 2: Garment Style Tiles

Clicking a category opens style tiles.

Example styles:

- Sundress
- Mabel Dress
- Vivien Dress
- Tilda Dress
- Frankie Dress

Each style tile could show:

- default meter range
- number of fabrics/prints costed for this style
- average making cost
- last order quantity
- warning if any fabric plan exceeds available stock

## Page 3: Style Detail / Costing Matrix

For one garment style, show all fabrics/prints that can be used for that style.

Example columns:

- Fabric picture
- Fabric/print name
- Fabric meters in stock
- Fabric cost per meter
- Meters per garment
- Lining/trim meters
- Stitching cost
- Zip/button size/type
- Zip/button cost
- Planned quantity
- Fabric required
- Can make
- Reserved fabric
- Remaining fabric
- Fabric cost
- Total garment cost
- Notes

Example calculation:

- Fabric: Keepsake
- In stock: 1200m
- Meters per garment: 3.3m
- Planned quantity: 100
- Fabric required: 330m
- Remaining: 870m
- Fabric cost per garment: `3.3 * fabric cost per meter`
- If using 5% fabric buffer: `fabric cost * 1.05`
- Total garment cost: `fabric cost + stitching + zip/button + trims`

## Key Actions

### Add Fabric To Style

From the style detail page, choose an existing fabric from Fabric in Stock and enter:

- meters per garment
- lining/trim meters
- stitching cost
- zip/button type
- zip/button cost
- notes

### Create Restock Order

From the style detail page, enter planned quantities for one or more fabrics and create/update restock orders.

The app should:

- calculate fabric required
- show remaining fabric before confirming
- warn if planned quantity exceeds available fabric
- create or update records on Existing Products Restock
- reserve fabric against the order

## Fabric Reservation Idea

Avoid permanently deducting fabric just because someone is typing a draft quantity.

Use reservations:

- Draft/planned order reserves fabric temporarily
- Confirmed/placed order commits the fabric usage
- Cancelled order releases fabric back

Fabric availability should show:

- total fabric in stock
- fabric reserved for planned/open orders
- fabric still available
- fabric required by the current draft plan

## Data Objects To Consider

### Garment Category

- id
- name
- sort order

### Garment Style

- id
- category id
- name
- SKU or style code
- default stitching cost
- default trim costs
- notes

### Style Fabric Costing

- id
- style id
- fabric id or fabric sheet/row reference
- meters per garment
- lining/trim meters
- stitching cost override
- zip/button type
- zip/button cost
- buffer percentage
- notes

### Fabric Reservation

- id
- restock order id
- style fabric costing id
- quantity planned
- meters reserved
- status: draft, planned, confirmed, cancelled, completed

## Open Questions

- Should fabric stock be deducted at order creation, confirmation, or when production starts?
- Should Existing Products Restock become the source of confirmed production quantities, or should the new planner create draft plans first?
- How should fabric rows be uniquely identified now that Fabric in Stock is manual data?
- Should product/style data be imported from Shopify products, manually created, or both?
- Do different sizes of the same garment need different fabric meters, or is one average meters-per-piece enough for now?

## Resume Prompt

To continue this later, ask:

`Read docs/fabric-planner-notes.md and continue designing the garment style tile / fabric costing planner.`
