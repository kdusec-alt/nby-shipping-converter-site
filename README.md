# NBY Shipping Data Converter

Live: https://kdusec-alt.github.io/nby-shipping-converter-site/

Browser-only processing; no source files leave the browser. Runtime dependencies are vendored. Do not commit operational files, templates containing customer data, or generated reports to this public repository.

## Daily workflow

1. Select Manifest Scan Date, then import FC / ZRSD1019 / MB52 / UPS / SF / AWB files together.
2. Choose the FC worksheet explicitly. Multiple SAP or carrier batches are rejected; remove redundant files first. Reimporting the same filename replaces that file.
3. Optional: load existing XLSX workbooks as separate Orders Update, Data Shipping and Shipping Status templates.
4. Run, inspect exceptions, then download individual workbooks or ZIP. Missing templates produce explicitly named DRAFT workbooks. Unsupported template layouts block only the affected formal workbook.
5. ZIP includes processing audit metadata. The last 30 audit entries are kept in this browser localStorage; source rows are not persisted there.

## Rules

- Filter by Manifest date, count distinct nonblank work orders. This is a work-order count, not an independently confirmed shoe-pair quantity. Conflicting FC rows are flagged and not silently resolved.
- FC Customer Reference joins SAP. Conflicting documents/items remain unresolved. MB52 matches document plus available material/item and retains storage location, batch and stock states. Stock is never subtracted from actual shipment.
- UPS tracking events are retained. Record counts do not represent shoe-pair quantities; UPS quantity validation remains pending until a quantity field/unit is confirmed.
- SF uses work order, carrier, tracking and ship group; repeated work-order quantities are flagged for grain review. Total footer rows are excluded.
- AWB comes only from OrderNumber / ConsolidationShipmentId. Missing/conflicting mappings stay blank.
- Expect Shipment is not calculated. An uploaded Shipping Status template retains its existing Expect cells and historical dates.
- XLSX templates are edited by changing worksheet XML inside the original ZIP; styles, other sheets and package parts are retained. Supported tabular bodies may resize and row formulas shift. Merged data areas, Excel tables and shared/array formulas are blocked for review. Shipping Status currently requires the target date and product entries to exist in its Summary and Detail shipment layout.
- Legacy XLS can be read as input, but format-preserving templates must be XLSX. This does not claim lossless XLS conversion.

## Verification (2026-11-27)

8 synthetic reconciliation regression tests: `node --test tests/engine.test.cjs`.
The 11/20 processed sample reproduced 364 FC work-order / SAP / MB52 matches; 366 UPS events with 364 distinct tracking numbers; one SF detail row. The supplied comparison workbook matched 3,701 checked work-order fields and 364 tracking summaries. This is not independent end-to-end golden certification: the FC base was extracted from the processed Data Shipping sample, SF is partial, and Shipping Status/Expect rules remain pending.

## Development

Edit this repository directly and deploy through its existing GitHub Pages flow. Preserve the existing hosting destination. Do not migrate to a separate host without request. Private samples belong outside the public checkout.
