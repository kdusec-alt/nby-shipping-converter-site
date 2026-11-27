# NBY Shipping Data Converter

Live: https://kdusec-alt.github.io/nby-shipping-converter-site/

Browser-only processing; no source files leave the browser. Runtime dependencies are vendored. Do not commit operational files, templates containing customer data, or generated reports to this public repository.

## Daily workflow

1. Import the FC, SAP ZRSD1019, MB52, UPS, SF Baozun and AWB mapping files.
2. Choose the FC worksheet explicitly and set separate dates for Orders Update, Data Shipping and the Shipping Status period. The reconciliation reference date is separate from these output filters.
3. Optional: load existing XLSX workbooks as separate Orders Update, Data Shipping and Shipping Status templates.
4. Run, inspect exceptions and quantity references, then download individual workbooks or ZIP. Missing templates produce explicitly named DRAFT workbooks. Unsupported template layouts block only the affected formal workbook.
5. ZIP includes processing audit metadata and selected report dates. The last 30 audit entries are kept in this browser localStorage; source rows are not persisted there.

## Confirmed and pending rules

- Orders Update and Data Shipping each filter FC by their own Manifest Scan Date. Shipping Status takes a start/end date and updates only dates present in the selected FC source.
- FC Customer Reference joins SAP. Conflicting documents/items remain unresolved. MB52 matches document plus available material/item and retains storage location, batch and stock states.
- Shipping Status currently counts distinct nonblank Work Order Nbr by Manifest date and Product Name. It is explicitly a work-order count, not a certified shoe-pair quantity. Actually Shipment business grain must be confirmed against the official workbook; Expect Shipment is not calculated.
- UPS and SF Baozun reconciliation remain separate. Tracking events are not shoe-pair quantities. UPS pair-quantity validation is pending until a quantity field and unit are confirmed.
- SF Qty is shown against SAP Confirmed Quantity as a reference, and alongside MB52 Unrestricted / Blocked Stock with an arithmetic delta. This is not an approved shipment reconciliation: MB52 is a stock snapshot, not daily shipped quantity. The Exception Report labels these rows as quantity references, not passes.
- AWB comes only from OrderNumber / ConsolidationShipmentId. Missing/conflicting mappings stay blank; the app never generates AWB values.
- XLSX templates are edited inside the original ZIP package to preserve sheets/styles and supported formulas. Unsupported layouts block formal output for review. Shipping Status requires each updated date and product to already exist in its Summary and Detail shipment layout.
- Legacy XLS can be read as input, but format-preserving templates must be XLSX.

## Golden Sample status

The 2026-11-20 processed sample previously reproduced 364 FC work-order / SAP / MB52 matches, 366 UPS events with 364 distinct tracking numbers, and one SF detail row. The FC base was extracted from a processed Data Shipping sample, and the SF file was partial. These checks do not certify the full outputs. Independent raw files and manual golden workbooks are still required for row-by-row acceptance.

## Development

Edit this repository directly and deploy through its existing GitHub Pages flow. Preserve the existing hosting destination. Do not migrate to a separate host without request. Private samples belong outside the public checkout.
