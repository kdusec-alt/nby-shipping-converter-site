# Project continuation

- Continue development in this repository and deploy to its existing GitHub Pages site. Do not create another host/project unless explicitly requested.
- Never commit operational Excel/CSV files, customer information, generated reports, credentials, or private templates to this public repository.
- Preserve original XLSX workbook structure and formatting through template editing. Label generic exports DRAFT; never imply they reproduce the manual design.
- Match FC to SAP by Work Order / Customer Reference, then MB52 by document plus available material/item. Preserve stock location/batch/status distinctions.
- UPS and SF reconciliation are independent. Tracking events are not shoe-pair quantities. MB52 stock does not reduce Actually Shipment.
- Use Manifest date explicitly. Do not hardcode sample dates, auto-pick the source with the best match rate, fabricate AWBs, or guess Expect Shipment.
- Flag ambiguous mappings and unresolved quantity grain. Do not report Golden Sample certification from a processed output used as input.
- Run npm test for engine/template changes. Private sample checks may run with `node tests/private-sample.cjs /absolute/private/sample/directory`.
