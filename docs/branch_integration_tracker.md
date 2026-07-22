# Branch Integration Tracker

Last updated: 2026-07-21
Integration branch: `codex/integration-step2-category-profiles`

## Active Branches

| Branch | Worktree | Purpose | Status |
|---|---|---|---|
| `codex/step3-workbook-cleanup` | `C:\Users\Sunco\Sunco Lighting\Product - Manny Tools\Ideation Development` | Current stable working base for Step 3 workbook cleanup and action vocabulary | Clean base, not `main` |
| `codex/category-profile-baseline` | `C:\Users\Sunco\Projects\ideation-category-profile-baseline` | Generate baseline category feature and attribute decision profiles from category intelligence data | Committed locally at `471a25b` |
| `codex/step2-sku-normalizer` | `C:\Users\Sunco\Projects\ideation-step2-sku-normalizer` | Rework Step 2 into a SKU concept normalizer while preserving Step 3 workbook contract | Committed locally at `11196cd` |
| `codex/integration-step2-category-profiles` | `C:\Users\Sunco\Projects\ideation-step2-category-integration` | Merge and validate both branches together before any main update | Active integration branch |

## Merge Status

- `codex/category-profile-baseline` merged cleanly into integration.
- `codex/step2-sku-normalizer` merged cleanly into integration.
- Integration-only patch applied after smoke review: generated profile match terms no longer include broad raw terms such as `voltage`, `wattage`, `beam angle`, or bare numbers that can cross-match unrelated attributes.

## Validation Completed

- `python -m compileall backend\app product_demand_ideation\src backend\maintenance_scripts\smoke_step2_sku_normalizer.py`
- `python -m json.tool backend\app\opportunity_engine\category_attribute_profiles.json`
- `python -m json.tool backend\app\prd_research_tool\config\category_feature_signal_profiles.json`
- `python -m json.tool backend\app\prd_research_tool\config\category_attribute_decision_profiles.json`
- `python "0 - Refresh Backend Data.py" --skip-line-review --skip-stackline`
- `python backend\maintenance_scripts\smoke_step2_sku_normalizer.py`
- Direct Step 3 prepare/finalize from the synthetic Panels Step 2 workbook.
- Real non-panel Step 2 smoke from existing Vaportights Step 1B workbook.
- Direct Vaportights generated-profile selection check.

## Generated Smoke Artifacts

- Step 2 smoke workbook:
  `C:\Users\Sunco\Projects\ideation-step2-category-integration\outputs\Ideations\PRD Ideation Workbooks\panels\panels_prd_ideations_2026-07-21_193231.xlsx`
- Step 3 smoke report:
  `C:\Users\Sunco\Projects\ideation-step2-category-integration\outputs\Research\Reports\panels_2026-07-21_193312_completed_rows.xlsx`
- Vaportights Step 2 smoke workbook:
  `C:\Users\Sunco\Projects\ideation-step2-category-integration\outputs\Ideations\PRD Ideation Workbooks\vaportights\vaportights_prd_ideations_2026-07-21_193631.xlsx`

## Current Findings

- Category intelligence rebuild generates baseline profiles for all 54 active categories.
- Step 3 profile loaders now see 58 feature-signal profiles: 4 manual and 54 generated.
- Step 3 attribute decision loader now sees 56 attribute-decision profiles: 2 manual and 54 generated.
- Synthetic Step 2 smoke produced three Panels rows in `NPD, NPD, Revision` order.
- Step 3 smoke prepared a Panels session, autofilled six Stackline-backed Amazon/Home Depot artifacts, finalized locally, and published a combined workbook.
- Banned artifact wording was not found in the smoke Step 3 workbook: `schema_references`, `templates/Competitors.md`, `Bulb-dependent`, `chandelier`, `validate against Step 1 listing and supplier file`, `TBD from supplier`, `pack-count target`, `_target_pack_count`.
- Vaportights generated profile selection now resolves to `generated_vaportights` and `generated_vaportights_attributes`, with clean labels such as `Wet rated`, `Hardwired`, `IP65`, `5000K`, `120-277V`, `Dimmable`, `Ceiling mount`, `Frosted finish`, `Polycarbonate`, `40W`, `DLC`, `FCC`, `ETL`.
- Real Vaportights Step 2 smoke generated 20 rows from `vaportights_true_gaps_2026-06-29_113653_product_demand_step1b.xlsx`.
- Real Vaportights Step 2 smoke did not contain banned/category-bleed terms: `chandelier`, `Bulb-dependent`, `schema_references`, `templates/Competitors.md`, `pack-count target`, `_target_pack_count`, `validate against Step 1 listing and supplier file`, `TBD from supplier`.
- Vaportights Step 3 prepare could not be completed in this sandbox because the `Redshift` ODBC DSN is not configured. The failure occurred before session creation while enforcing the fresh Redshift Stackline cache requirement.

## Remaining Before Main

- Commit the integration-only broad-term fix and this tracker.
- Push the integration branch for review.
- Do not update `main` until Manny reviews the integrated outputs.
- Validate Vaportights Step 3 prepare again on the user machine or deployment environment where the Redshift DSN/connector is available.
