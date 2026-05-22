# Sunco Product Opportunity Engine

User-facing folder name: **Ideation Development**

Run the scripts in order:

1. `1 - Category Ideation Generator.py`
2. `2 - Ideation Template Generator.py`
3. `3 - Ideation Research Tool.py`

The scripts ask which category to run and list the available category options. Category owner is inferred from `templates/category_reference.csv`; parent categories are intentionally excluded.

Step 1 includes its selection rationale in the workbook. Check `Summary`, `Sources and Audit`, `Amazon Source Audit`, and `Run Audit` for:

- the decision tree used to choose ideations
- the reason these ideas were selected over other possible ideas
- row-level decision notes showing category fit, market signal, Sunco gap evidence, priority/confidence, and supplemental warnings

Use this as the answer when asked: "How did we determine those ideations would be successful over others?" The tool is not claiming guaranteed success; it is ranking higher-probability ideas for research based on competitor demand signals, Sunco coverage gaps, and PM actionability.

Step 2 and Step 3 use a clean handoff contract. Step 2 creates one row per candidate SKU when Step 1 evidence or PM action names distinct options such as bulb count, size, or form factor. It does not force a minimum row count. Step 3 then creates one final research workbook with one sheet per Step 2 row, so SKU-level permutations are researched separately without creating multiple workbooks.

Step 2 can still carry notes and validation guidance, but Step 3 only scores concise customer-facing feature and certification signals. Non-applicable values, `TBD`, `N/A`, optional bulb guidance, and supplier-validation prose are ignored or converted into cleaner checks such as `E12 socket`, `E26 socket`, `ETL`, and `UL`.

Section F is driven by a backend category feature-signal database at `backend/app/prd_research_tool/config/category_feature_signal_profiles.json`. For decorative chandelier rows, Step 3 now adds SKU-defining signals such as light count, size range, style/form factor, finish/color, material, mounting, and socket type before it scores competitor coverage.

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

Cached category data is reusable for 30 days. If data is older than 30 days, the script explains that a refresh is needed and creates a new timestamped run output. Users can also force refresh from the terminal prompt.

## Credentials

Do not put secrets in this shared SharePoint folder. If a future live connector needs credentials, use a machine-local file:

```text
Windows: C:\Users\<user>\.sunco_ideation_development\.env
macOS: ~/.sunco_ideation_development/.env
```

The shared folder may contain `.env.example` files only.
