# Deployment Smoke Tests And Team Rollout

## Current Deployment Goal

Package the Ideation Development tool so Product Managers can run the same Step 0-3 process from the shared project folder without Codex manually steering every run.

The release should preserve the visible workflow:

1. Refresh backend data when needed.
2. Generate category ideations / gap workbook.
3. Generate PRD ideation workbook.
4. Prepare and finalize ideation research workbook.

## Smoke Test Status

Last smoke test date: 2026-07-01

Category: Vaportights

Step 2 input:

`outputs/Ideations/Gap Workbooks/vaportights/vaportights_true_gaps_2026-06-29_113653.xlsx`

Step 2 output:

`outputs/Ideations/PRD Ideation Workbooks/vaportights/vaportights_prd_ideations_2026-07-01_120557.xlsx`

Step 2 validation result:

- Excel COM open/save validation passed.
- 20 ideation rows generated.
- `Target MSRP`, `Target Vendor Cost`, `Target Margin % (Shopify)`, and `Target Margin % (Amazon)` stayed numeric.
- `Cost Type` stayed `Landed`.
- No scan hits for old validation/vendor wording or category bleed terms.

Step 3 smoke input:

`backend/cache/smoke_tests/vaportights_step2_smoke_row4_2026-07-01_121350.xlsx`

Step 3 smoke session:

`backend/research_sessions/vaportights_2026-07-01_121403`

Step 3 smoke report:

`outputs/Research/Reports/vaportights_2026-07-01_121403_completed_rows.xlsx`

Step 3 validation result:

- One-row session prepared successfully.
- Three raw artifacts were written for Amazon, brick-and-mortar, and brand-site collection.
- Raw artifact validation passed: 3 valid, 0 invalid.
- Manifest update marked collection stages complete.
- Finalization produced one combined Excel report with one Summary sheet and one ideation sheet.
- Workbook text scan found no noisy search-planning strings such as `target from Step 1 evidenceW`, `selectableK`, supplier-file wording, chandelier bleed, or bulb-dependent bleed.

## Issue Found And Fixed During Smoke Test

Step 3 task generation was structurally working, but the generated Codex collection task had noisy search phrases such as:

- `75W target from Step 1 evidenceW`
- `3CCT selectableK`
- full competitor product snippets being treated as brand names

Fix:

- Cleaned Step 3 search-planning terms before unit formatting.
- Removed placeholder mounting/CCT text from query terms.
- Extracted real brand names from Step 1 competitor snippets before building brand-site watchlists.

Changed file:

`backend/app/prd_research_tool/tools/competitive_research_engine.py`

## Deployment Readiness Gates

Before sharing with the full Product team:

1. Run one clean end-to-end category from Step 0 through Step 3.
2. Run one category where Step 3 has Stackline/local autofill.
3. Run one category where Step 3 requires Codex web collection.
4. Confirm the scripts open from the shared folder and do not require manual repo navigation.
5. Confirm users understand that Step 3 has two phases: prepare collection, then finalize after collection is complete.
6. Keep all backend sessions, logs, caches, and smoke files out of user-facing output folders.
7. Provide a one-page `Start Here` guide and a short troubleshooting section.
8. Decide who owns backend refreshes and how often they should run.

## Rollout Packaging Notes

Recommended deployment model:

- Keep the shared folder as the user-facing launch point.
- Keep only the numbered scripts, README, templates, and output folders visible at the top level.
- Keep raw sessions, logs, caches, smoke workbooks, and helper internals under `backend`.
- Use one run folder/session per category request in the backend.
- Publish only final Excel workbooks into the user-facing output folders.

Recommended release communication:

- Tell PMs that Step 1 supports line review and opportunity selection.
- Tell PMs that Step 2 turns selected opportunities into structured SKU-level ideations.
- Tell PMs that Step 3 is a research report to prepare for PRD/RFQ decisions, not a PO Ready form.
- Tell PMs that PO approval evidence is stronger after RFQ, samples/EVT, vendor PO form, pack-size review, and RiskCalculatorV3.

## Next Work

- Strengthen Step 2 after Step 1 changes so PRD ideation rows carry enough clean, SKU-specific detail.
- Run a full Step 3 category with all rows, not only a smoke subset.
- Build the team-facing deployment folder/package and test it from the same shared path a PM would use.
- Push the current fixes once smoke-test outputs and deployment docs are reviewed.
