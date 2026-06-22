# Sunco Product Opportunity Engine

User-facing folder name: **Ideation Development**

Run the scripts in order:

0. `0 - Refresh Backend Data.py` when backend data is missing or older than 30 days
1. `1 - Category Ideation Generator.py`
1B. `1B - Product Demand Ideation Generator.py` for the demand-weighted Step 1B flow
2. `2 - Ideation Template Generator.py`
3. `3 - Ideation Research Tool.py`

The scripts ask which category to run and list the available category options. Category owner is inferred from `templates/category_reference.csv`; parent categories are intentionally excluded.

Step 1 includes its selection rationale in the workbook. Check `Summary`, `Sources and Audit`, `Amazon Source Audit`, and `Run Audit` for:

- the decision tree used to choose ideations
- the reason these ideas were selected over other possible ideas
- row-level decision notes showing category fit, market signal, Sunco gap evidence, priority/confidence, and supplemental warnings

Use this as the answer when asked: "How did we determine those ideations would be successful over others?" The tool is not claiming guaranteed success; it is ranking higher-probability ideas for research based on competitor demand signals, Sunco coverage gaps, and PM actionability.

Step 1 also adds an `Existing SKU Line Review` sheet after the recommendation/audit sheets. This sheet is additive and does not change the existing Step 1 workbook contract used by Step 2 and Step 3. Line-review evidence is intentionally restricted to approved Postgres/Redshift snapshots; legacy local Line Review CSV/workbook sources are blocked. If no approved snapshot exists for the selected category, the sheet stays in the workbook with a clear source-policy note instead of fabricating rows.

Approved line-review snapshots should be JSON files under:

```text
backend/source_data/postgres_exports/line_reviews
backend/source_data/redshift_exports/line_reviews
```

Use a filename containing the category slug, such as `panels_line_review_2026-05-26.json`. The preferred shape is:

```json
{
  "source_system": "postgres",
  "category_slug": "panels",
  "generated_at": "2026-05-26T00:00:00Z",
  "sql": "SELECT ...",
  "rows": []
}
```

Step 1 writes the recommended Postgres query for each category to `backend/cache/ideation_data/<category_slug>/sql/line_review_postgres.sql`, and the SQL is included on `Run Audit`.

Step 2 and Step 3 use a clean handoff contract. Step 2 creates one row per candidate SKU when Step 1 evidence or PM action names distinct options such as bulb count, size, or form factor. It does not force a minimum row count. Step 3 then creates one final research workbook with one sheet per Step 2 row, so SKU-level permutations are researched separately without creating multiple workbooks.

Step 2 can still carry notes and validation guidance, but Step 3 only scores concise customer-facing feature and certification signals. Non-applicable values, `TBD`, `N/A`, optional bulb guidance, and supplier-validation prose are ignored or converted into cleaner checks such as `E12 socket`, `E26 socket`, `ETL`, and `UL`.

Section F is driven by a backend category feature-signal database at `backend/app/prd_research_tool/config/category_feature_signal_profiles.json`. For decorative chandelier rows, Step 3 now adds SKU-defining signals such as light count, size range, style/form factor, finish/color, material, mounting, and socket type before it scores competitor coverage.

The 1/2/3 workflow also preloads a local category intelligence SQLite database at `backend/source_data/category_intelligence/sunco_category_intelligence.sqlite`. This file is backend source/cache data and is intentionally not tracked by git. Rebuild it with:

```text
python -m backend.app.opportunity_engine.build_category_intelligence --root "."
```

The tracked schema and builder live in `backend/app/opportunity_engine`. The builder seeds categories, local gap evidence, available Stackline/cache inventory, category attribute defaults, and Section F feature-signal profiles. If a private Shopify/catalog export is available, place it under `backend/source_data/postgres_exports` as ignored source data before rebuilding.

To create the clean backend catalog/spec reference from the current Sunco specs CSV, run:

```text
python -m backend.app.opportunity_engine.clean_catalog_specs_reference --root "."
python -m backend.app.opportunity_engine.build_category_intelligence --root "."
```

The cleaner writes ignored backend files under `backend/source_data/catalog_specs`: one product reference CSV, one long-format SKU/spec attribute CSV, and a manifest. Those files can be regenerated from a newer Redshift/Postgres export later without changing the 1/2/3 workflow.

To create the clean backend SKU decoder reference from the current Manny SKU Decoder CSV, run:

```text
python -m backend.app.opportunity_engine.clean_sku_decoder_reference --root "."
python -m backend.app.opportunity_engine.build_category_intelligence --root "."
```

The SKU decoder cleaner writes ignored backend files under `backend/source_data/sku_decoder`. The decoder is used as category-matching support, not as the source of truth for product data. Postgres/Redshift still supply the actual line-review rows; the decoder helps the generated line-review SQL recognize SKU prefixes such as `VL` for Vanity, `WS` for Wall Sconce, `PN24` for Panels, and `VT` for Vapor Tight.

Step 3 has two phases:

1. Choose `Prepare new research session from latest Step 2 workbook`.
   - This creates a backend session and collection instructions.
   - This does **not** create the final research workbook yet.
   - When a Redshift Stackline cache exists for the category, Step 3 uses that first, then local Stackline CSVs, then web fallback.
   - Stackline-backed Amazon and Home Depot raw artifacts are autofilled locally when possible, so Claude only handles the remaining web-enrichment tasks.
   - Lowest-token option: if the remaining tasks are only optional brand-site enrichment, skip Claude and run Step 3 again with `Finalize latest prepared session`; the report will be based on available Stackline/local data and the workbook/session notes will say brand-site enrichment was skipped.
   - If competitor collection is still pending, the script prints the pending row/channel tasks and opens the session `instructions` folder.
   - Open `1 - COPY THIS PROMPT TO CLAUDE.md` and copy the full contents into Claude to start the raw collection handoff.
   - The prompt points Claude to `instructions/_support/CLAUDE_TASKS_LITE.json`; users do not need to open the support files directly.
2. Complete the raw competitor collection files listed in the instructions.
   - The files live under the session `raw` folder.
   - Required channels are usually `amazon`, `brick_and_mortar`, and `brand_sites`.
3. Run `3 - Ideation Research Tool.py` again and choose `Finalize latest prepared session`.
   - This validates raw files, normalizes competitors, analyzes the ideations, builds the report workbook, and publishes Excel output to `outputs/Research/Reports`.

Use `Check latest session status` any time Step 3 opens a backend folder but no workbook appears. That means the session is prepared, but collection/finalization is not complete.

## Low-Token Clean Run

To test the full workflow without using Claude tokens:

1. Run `1 - Category Ideation Generator.py`.
2. Run `2 - Ideation Template Generator.py`.
3. Run `3 - Ideation Research Tool.py` and choose `Prepare new research session from latest Step 2 workbook`.
4. If the script says Stackline/local artifacts were autofilled, run `3 - Ideation Research Tool.py` again and choose `Finalize latest prepared session`.

This produces a research workbook from local/Stackline evidence only. The tool writes explicit skipped-brand-site artifacts so the run can finish without pretending those brand pages were collected. Treat missing brand-site enrichment as a noted research gap, not as a completed full external web collection.

## Folder Rules

- User-facing gap workbooks are saved in `outputs/Ideations/Gap Workbooks`.
- User-facing PRD ideation workbooks are saved in `outputs/Ideations/PRD Ideation Workbooks`.
- User-facing research reports are saved in `outputs/Research/Reports`.
- Backend data, cache, logs, migrated sessions, and copied tool code live under `backend`.

## Data Freshness

Cached category data is reusable for 30 days. A refresh means pulling the latest approved Postgres/Redshift evidence into backend snapshot files so the user-facing workbook can run from a stable local copy instead of querying databases every time. If data is older than 30 days, Step 1 explains that a refresh is needed and creates a new timestamped run output. Users can also force refresh from the terminal prompt.

`0 - Refresh Backend Data.py` refreshes approved Postgres line-review snapshots under:

```text
backend/source_data/postgres_exports/line_reviews
```

These JSON snapshots are backend data and are intentionally ignored by git. They are the files Step 1 uses to populate `Existing SKU Line Review`.

Step 0 runs through local Postgres ODBC or a local Postgres connection string. It does not use Codex MCP. Configure one of these in the machine-local env file before running live refreshes:

```text
POSTGRES_ODBC_DSN=Postgres
POSTGRES_DATABASE=<database>
POSTGRES_USER=<user>
POSTGRES_PASSWORD=<password>
```

or:

```text
POSTGRES_DSN=DSN=Postgres;UID=<user>;PWD=<password>;Database=<database>
```

`1B - Product Demand Ideation Generator.py` adds the demand-weighted Step 1B flow. It refreshes ecommerce competitor evidence and Stackline/Amazon evidence through local Redshift ODBC, uses the local Sunco catalog cache or Postgres ODBC refresh for catalog coverage, writes a Step 1-compatible workbook, and publishes a Step 2 handoff copy under `outputs/Ideations/Gap Workbooks/<category>`.

## Credentials

Do not put secrets in this shared SharePoint folder. If a future live connector needs credentials, use a machine-local file:

```text
Windows: C:\Users\<user>\.sunco_ideation_development\.env
macOS: ~/.sunco_ideation_development/.env
```

The shared folder may contain `.env.example` files only.
