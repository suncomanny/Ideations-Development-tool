# Step 2 Improvement Tracker

Last updated: 2026-07-22
Branch: `codex/step2-sku-normalizer`

## Core Purpose

Step 2 is the PM-curated opportunity-to-SKU translation layer.

Step 1 finds and ranks product opportunities. PMs review that workbook and delete the rows they do not want to move forward. Step 2 should then convert the remaining usable rows into clean SKU-level ideation rows that can support PRD drafting, RFQ preparation, and Step 3 research.

Step 2 should not behave like another demand-ranking gate. PM row deletion is the gate.

## Scope

In scope:
- Read the PM-curated Step 1 workbook.
- Preserve all remaining usable opportunity rows unless a row is structurally unusable.
- Convert each remaining opportunity into one clean candidate SKU ideation row.
- Extract concise proposed requirements from Step 1 evidence.
- Keep source links and audit evidence available for PM verification.
- Keep opportunity vocabulary consistent across Step 1, Step 2, and Step 3.
- Keep internal uncertainty out of PRD-facing spec cells.
- Produce a workbook Step 3 can consume without extra interpretation.

Out of scope:
- Pack-size and pack-count recommendations. These are handled by the independent pack-size workflow from the other team.
- Final vendor quote values.
- Final MOQ, first order quantity, or PO-ready approval quantities.
- Final engineering validation, sample approval, or EVT decisions.
- Re-ranking opportunities after the PM has already curated the Step 1 workbook.

## Controlled PM-Facing Product Action Vocabulary

Use the same product-action labels across all three steps:
- `NPD` - New Product Development; bring in as a new SKU/product and move toward PRD/RFQ.
- `Revision` - rolling change, feature add, product update, or merchandising/listing correction on an existing SKU/family.
- `Concept Review` - enough signal to review more closely before deciding whether the path is NPD, Revision, or no action.
- `Hold` - do not prioritize from this run; keep the evidence for reference.

Keep gap reasons separate from product actions. Gap reasons can include feature gap, existing Sunco coverage, partial coverage, strategic outlier, or incumbent optimization, but they should not replace the product-action label.

Sort product lists by product action first: `NPD`, then `Revision`, then `Concept Review`, then `Hold`. Revision rows must also identify the existing SKU/family to revise and the specific feature, spec, price, listing, or merchandising change being recommended.

## Improvement Checklist

### 1. Row Intake Contract

Status: Implemented - validated with Panels and Wraparounds smoke tests

Goal:
Step 2 should process the PM-curated workbook as the source of truth.

Actions:
- Replace strict `Priority = High` and `Confidence = High` intake with a broader "remaining usable rows" rule.
- Keep priority and confidence as context fields, not as a hard inclusion gate.
- Skip only rows that are blank, missing a recommendation, missing usable evidence, or marked as non-actionable by a controlled skip value.
- Add audit counts for rows read, rows converted, and rows skipped.

Acceptance criteria:
- A `Medium / Directional` strategic outlier left in the Step 1 workbook flows into Step 2.
- PM-deleted rows do not flow into Step 2.
- Step 2 audit explains why any remaining row was skipped.

Implementation notes:
- Step 2 now converts remaining usable rows from `Recommendations` and `Amazon Recommendations` without requiring `Priority = High` and `Confidence = High`.
- Priority and confidence are preserved as context in Source Mapping instead of being used as hard filters.
- Explicit PM skip markers are supported through `[skip]`, `[do not convert]`, `[do not move forward]`, `[hold]`, or exact skip values in Priority, Confidence, or PM action fields.
- The Run Audit now reports rows read, rows converted, rows skipped, skipped-row detail, and rows remaining after dedupe.
- Panels validation confirmed `Strategic outlier / High-output watchlist` rows now flow into Step 2 when left in the Step 1 workbook.

### 2. Requirement Field Cleanup

Status: Partially implemented - validated with Panels smoke test

Goal:
PRD-facing fields should read like proposed product requirements, not internal validation notes.

Actions:
- Keep fields such as voltage, lumens, wattage, CCT, dimming, mounting, certifications, MSRP, and target vendor cost concise.
- Move uncertainty, caveats, URL checks, and source rationale into research notes or audit fields.
- Remove phrases such as `validate against supplier file`, `TBD from supplier`, `before PRD lock`, and other process language from spec cells.

Acceptance criteria:
- Step 2 spec cells can be copied into a PRD/RFQ draft without cleanup.
- Any remaining uncertainty is contained in notes/audit fields only.

### 3. SKU-Level Candidate Normalization

Status: First pass implemented - validated with synthetic Panels smoke test

Goal:
Each Step 2 row should represent one candidate SKU concept at the level Step 3 will research.

Actions:
- Split only when Step 1 evidence clearly names distinct SKU-defining options such as size, form factor, lumen/output tier, light count, finish, mounting type, or control type.
- Do not split on pack counts or pack sizes.
- Do not force extra rows to hit a minimum count.
- Preserve the parent opportunity relationship in notes when a split occurs.

Acceptance criteria:
- One Step 2 row maps cleanly to one Step 3 report sheet.
- Pack-count permutations are not generated by Step 2.
- Splits are explainable from Step 1 evidence.

Implementation notes:
- Pack-size and pack-count splitting was removed from Step 2.
- Under Cabinet no longer generates separate rows for pack-count mentions.
- SKU candidate generation is now centralized in `_step2_candidate_rows`.
- Default behavior remains one PM-kept Step 1 row to one Step 2 SKU concept row.
- Step 2 now splits only when the selected category profile allows the SKU-defining attribute and Step 1 evidence or PM action explicitly names multiple options.
- Panels profile rules now allow explicit size, output-tier, wattage, mounting, and control-type splits without enabling decorative/chandelier assumptions.
- Panels validation confirmed wattage and lumen requirement cells now output clean values such as `210W` and `33,600lm` instead of internal wording like `target from Step 1 evidence`.

### 4. Category-Safe Attribute Extraction

Status: First pass implemented - validated with Panels/Wraparounds historical smoke tests and synthetic Panels SKU-normalizer smoke test

Goal:
Step 2 should extract specs using category-aware logic without category bleed.

Actions:
- Use Step 1 evidence first.
- Use selected category profile/defaults second.
- Use generic fallback only when category-specific data is absent.
- Prevent decorative/chandelier/bulb language from leaking into panels, wraps, vaportights, strips, or other integrated LED fixture categories.
- Preserve luminaire logic: lumens are the user-facing brightness target, wattage is the installer/load target.

Acceptance criteria:
- Panels do not receive chandelier/bulb-dependent assumptions.
- Integrated LED fixture categories do not get socket/bulb guidance unless Step 1 explicitly supports it.
- Category profiles improve the row without overriding direct evidence.

Implementation notes:
- Attribute mode is now read from the selected category profile where defined.
- Form-factor inference honors `form_factor_terms` from the selected category profile before applying generic integrated LED fallbacks.
- Decorative fixture enrichment no longer writes `Bulb-dependent` into PRD-facing spec cells.
- Step 2 no longer falls back to the old `schema_references` ideation template when the maintained template is missing.

### 5. Source Link and Evidence Preservation

Status: Implemented - validated with Panels smoke test

Goal:
PMs should be able to verify every major claim by clicking through to source evidence.

Actions:
- Preserve competitor PDP/review links from Step 1.
- Keep URL status results visible in notes/audit fields.
- Avoid attaching unrelated links to fields that are not supported by that source.
- Keep evidence language concise enough for Step 3 to parse.

Acceptance criteria:
- Every Step 2 row has a traceable source link when Step 1 provided one.
- Dead or blocked links are flagged, not silently treated as verified.
- Links are placed in relevant source/audit fields, not random spec fields.

Implementation notes:
- Source Mapping now includes dedicated `Source URL` and `URL status` columns.
- Long evidence text is clipped in Source Mapping so the sheet is easier to scan while full row evidence remains in the Ideations research notes.
- URL status continues to be summarized in Run Audit.

### 6. Step 2 Audit Sheet

Status: Partially implemented - validated with Panels smoke test

Goal:
The workbook should explain what Step 2 did without requiring code review.

Actions:
- Add or improve audit rows for source workbook, category, owner, run timestamp, rows read, rows converted, rows skipped, URL status summary, category profile used, and SKU split decisions.
- Include a clear note that pack-size/pack-count recommendations are intentionally excluded and handled by the separate pack-size workflow.

Acceptance criteria:
- A PM or reviewer can understand why Step 2 included each row.
- Any skipped rows are auditable.
- The audit does not include stale Claude-era instructions or legacy data-source assumptions.

Implementation notes:
- The Run Audit now documents the PM-curated row intake rule.
- The Run Audit now states that pack-size and pack-count recommendations are intentionally excluded from Step 2 and handled by the separate pack-size workflow.
- Step 2 now filters pack-count-only Amazon cues and removes `multi-pack` from generated ideation/action language while preserving raw source listing titles for verification.
- Additional audit cleanup may still be needed after the next workbook review.
- Source Mapping headers are now normalized on each run so URL support columns are always present.

### 7. Sunco Reference SKU Integrity

Status: Implemented - validated with Panels smoke test

Goal:
Column D should contain a real Sunco reference SKU or family reference, not a text fragment parsed from Step 1 evidence.

Actions:
- Parse Step 1 active-catalog coverage case-insensitively so SKUs with lowercase size tokens such as `2x4` are not truncated.
- Canonicalize parsed coverage SKUs against the `Existing SKU Line Review` family references when the line-review sheet is available.
- Fall back to the best-selling adjacent line-review family when coverage cannot be verified.

Acceptance criteria:
- Column D does not contain malformed tail fragments such as `4-50W-0K-10PK`.
- Non-TBD column D values resolve to `Existing SKU Line Review` references when that sheet is present.

## Latest Validation Results

Validated on 2026-07-14:
- Compile check passed: `python -m compileall backend\app product_demand_ideation\src`
- Panels Step 2 output: `outputs\Ideations\PRD Ideation Workbooks\panels\panels_prd_ideations_2026-07-13_182150.xlsx`
- Panels result: 18 recommendation rows read / 18 converted; 10 Amazon rows read / 10 converted; 0 skipped; 24 Step 2 candidate rows after dedupe and SKU splitting.
- Panels strategic outlier check: `PLT-80112` / `33,600lm` row flowed into Step 2 as `Strategic outlier / High-output watchlist`.
- Wraparounds Step 2 output: `outputs\Ideations\PRD Ideation Workbooks\wraparounds\wraparounds_prd_ideations_2026-07-13_182345.xlsx`
- Wraparounds result: 16 recommendation rows read / 16 converted; 6 Amazon rows read / 6 converted; 0 skipped; 20 Step 2 candidate rows.
- Step 3 prepare check passed from the new Panels Step 2 workbook: `backend\research_sessions\panels_2026-07-13_182456`
- Banned wording scan passed for Panels and Wraparounds for:
  - `validate against Step 1 listing and supplier file`
  - `TBD from supplier`
  - `Bulb-dependent`
  - `chandelier`
  - `pack-count target`
  - `_target_pack_count`

Validated on 2026-07-21:
- Compile check passed: `python -m compileall backend\app product_demand_ideation\src`
- Panels Step 2 output: `outputs\Ideations\PRD Ideation Workbooks\panels\panels_prd_ideations_2026-07-21_160538.xlsx`
- Panels result: 18 recommendation rows read / 18 converted; 10 Amazon rows read / 10 converted; 0 skipped; 24 Step 2 candidate rows.
- Strategic outlier check: `PLT-80112` / `33,600lm` row remained in Step 2 as `Strategic outlier / High-output watchlist`.
- Requirement cell polish check passed: strategic outlier row has `210W`, `33,600lm`, `4000K/5000K`, numeric MSRP, and numeric target vendor cost.
- Source Mapping check passed: headers include `Source URL` and `URL status`; the PLT-80112 row includes `https://www.1000bulbs.com/product/232521/PLT-80112.html` and `verified live URL (200)`.
- Step 3 prepare check passed from the new Panels Step 2 workbook: `backend\research_sessions\panels_2026-07-21_160756`
- Banned wording scan passed for:
  - `target from Step 1 evidence`
  - `validate against Step 1 listing and supplier file`
  - `TBD from supplier`
  - `Bulb-dependent`
  - `chandelier`
  - `pack economics`
  - `pack-count target`
  - `_target_pack_count`

Validated on 2026-07-21 after Sunco Reference SKU fix:
- Compile check passed: `python -m compileall backend\app product_demand_ideation\src`
- Panels Step 2 output: `outputs\Ideations\PRD Ideation Workbooks\panels\panels_prd_ideations_2026-07-21_163911.xlsx`
- Column D validation passed: 24 filled rows; 0 values start with malformed numeric SKU tails.
- Every non-TBD column D value resolved to an `Existing SKU Line Review` family reference.
- Example corrected references:
  - `PN_SM2x2-40W-0K-1PK` maps to `PN_SM2X2-40W-0K`
  - `PN_SM2x4-50W-0K-10PK` maps to `PN_SM2X4-50W-0K`
  - `ULI-HG-GRUL-BBW2X4-YYA-4PK` maps to `ULI-HG-GRUL-BBW2X4-YYA`

Validated on 2026-07-22 for Step 2 SKU normalizer first pass:
- Compile check passed: `python -m compileall backend\app product_demand_ideation\src backend\maintenance_scripts\smoke_step2_sku_normalizer.py`
- Synthetic Panels Step 1 smoke workbook: `backend\cache\smoke_tests\step2_sku_normalizer\panels_step1_smoke_2026-07-21_192236.xlsx`
- Synthetic Panels Step 2 output: `outputs\Ideations\PRD Ideation Workbooks\panels\panels_prd_ideations_2026-07-21_192236.xlsx`
- Synthetic Panels result: one explicit two-size/two-output NPD opportunity became two SKU concept rows, and one Revision opportunity stayed one Revision row.
- Smoke assertions passed for action sort order, 2x2/3,200lm and 2x4/5,000lm row targets, Revision SKU/change notes, and banned wording scan for `schema_references`, `templates/Competitors.md`, `Bulb-dependent`, `chandelier`, `pack-count target`, and `_target_pack_count`.
- Excel COM validation could not complete in the sandbox: `A specified logon session does not exist. It may already have been terminated.`

## Validation Plan

Run after each implementation pass:
- `python -m compileall backend\app product_demand_ideation\src`
- Step 2 from a fresh Panels Step 1 workbook.
- Step 2 from a category with a strategic outlier or medium-confidence row intentionally left in the workbook.
- Step 2 from one non-panel category to confirm no category bleed.
- Scan generated workbook for banned wording:
  - `validate against Step 1 listing and supplier file`
  - `supplier file`
  - `TBD from supplier`
  - `Bulb-dependent`
  - `chandelier` when the active category is not decorative
- Confirm Step 3 can prepare a session from the generated Step 2 workbook.

## Current Open Decisions

- Controlled skip values implemented: `[skip]`, `[do not convert]`, `[do not move forward]`, `[hold]`, or exact skip values in Priority, Confidence, or PM action fields.
- `Strategic outlier / High-output watchlist` rows flow into Step 2 automatically whenever left in the workbook.
- Decide whether Step 2 should add a lightweight "PM Review Notes" column if the template can support it without breaking Step 3.
