# Product Demand Ideation

This is the isolated workspace for adding the combined product-demand + Stackline model to the Ideation Development tool.

It is intentionally separate from the existing scripts:

- `1 - Category Ideation Generator.py`
- `1B - Product Demand Ideation Generator.py`
- `2 - Ideation Template Generator.py`
- `3 - Ideation Research Tool.py`

`1B - Product Demand Ideation Generator.py` is the user-facing launcher for this isolated tree. Do not wire this into the existing Step 1 flow until the isolated model is validated.

## Standalone Requirement

The final tool must run outside of Codex.

Production data access should use the approved source connector for each source family. Local credentials stay outside the shared project folder, such as:

```text
C:\Users\<user>\.sunco_ideation_development\.env
```

Expected production path:

```text
approved source connector -> local snapshot/cache refresh -> workbook generator
```

Redshift ecommerce competitor evidence now prefers the local ODBC DSN:

```text
DSN=Redshift
```

If the Windows DSN does not persist credentials, create this local-only file:

```text
C:\Users\<user>\.sunco_ideation_development\.env
```

Supported Redshift options:

```text
REDSHIFT_ODBC_DSN=Redshift
REDSHIFT_USER=odbc_user
REDSHIFT_PASSWORD=<password from data team>
REDSHIFT_DATABASE=dev
```

Alternatively, use a full ODBC string:

```text
REDSHIFT_DSN=DSN=Redshift;UID=odbc_user;PWD=<password>;Database=dev
```

Postgres ODBC is not required for Manny's workstation. Sunco catalog coverage uses the saved local SQLite cache, a Postgres MCP refresh, or approved Postgres/Redshift export files.

Sunco catalog coverage is designed to run from a local SQLite cache:

```text
product_demand_ideation/cache/sunco_catalog.sqlite
```

Refresh it intentionally with:

```text
Refresh Product Demand Local Catalog Cache.py
```

Normal `1B` workbook runs read from that local cache. If the cache does not exist yet, the tool seeds it through Postgres MCP and then continues. For Sunco's current catalog change rate, a weekly or monthly refresh is enough for routine A/B testing; refresh before leadership readouts or after known catalog/product-status updates.

The SharePoint workbook `Sku's Classification.xlsx` / `PowerBI Families` can also be cached locally for the latest PM/category/Series ownership layer:

```text
product_demand_ideation/cache/sku_classification.sqlite
```

Refresh it from a locally synced/exported copy of the current SharePoint workbook:

```text
Refresh Product Demand SKU Classification Cache.py "C:\path\to\Sku's Classification.xlsx"
```

This cache is optional at runtime. When present, 1B enriches Sunco active-catalog coverage checks with PowerBI category, PM responsible, and Series. Treat the newest SharePoint workbook as the source of truth because other PMs may update it after local PM cleanup work.

Category profiles can be checked without generating workbooks:

```text
Audit Product Demand Category Profiles.py
```

This uses the local Redshift ODBC DSN and writes a compact CSV audit under `product_demand_ideation/profile_audits/`. Use it before expanding 1B to a new category so noisy filters are fixed before leadership sees the output.

## Model

The tool should use a combined model:

| Recommendation Area | Lead Signal | Required Support |
| --- | --- | --- |
| Main `Recommendations` tab | Competitor ecommerce/PDP inventory movement as demand proxy | Sunco catalog coverage + luminaire spec fit + Stackline/Amazon support |
| `Amazon Recommendations` tab | Stackline/Amazon demand | Sunco Amazon coverage + competitor inventory movement where available |

Stackline/Amazon data should remain part of every recommendation path.

The main recommendation tab is intentionally weighted toward Shopify/ecommerce launches, not Amazon launch priority:

- competitor ecommerce inventory movement = 35%
- Sunco coverage/gap = 25%
- luminaire performance fit = 20%
- Stackline/Amazon demand = 10%
- data quality = 10%

The Amazon recommendation tab stays Amazon/Stackline-led:

- Stackline/Amazon demand = 50%
- Sunco coverage/gap = 25%
- luminaire performance fit = 10%
- competitor ecommerce inventory movement = 5%
- data quality = 10%

The ecommerce layer is sourced from Redshift:

- `public.v_competitors_scrapping_latest` for latest competitor PDP title/spec/category/price/image/URL evidence
- `public.v_competitors_inventory_daily` for PDP inventory movement and price movement by URL

The Redshift refresh path is:

```text
local Amazon Redshift ODBC DSN -> pyodbc -> product_demand_ideation/experiments/<category>/exports/*_ecommerce_competitor_evidence_*.json
```

Normal 1B runs refresh ecommerce snapshots through local Redshift ODBC when the newest snapshot is not ODBC-sourced or is older than 24 hours. Use `PRODUCT_DEMAND_ECOMMERCE_SNAPSHOT_MAX_AGE_HOURS` to change that freshness window. Redshift MCP is not used by the integrated main tool.

When Redshift ecommerce PDP rows exist for the selected category, 1B uses them to lead the main `Recommendations` tab. Amazon-derived display rows stay in `Amazon Recommendations`.

For light-producing categories, the model also normalizes luminaire performance:

- lumens = user-facing brightness target
- wattage = installer/electrician load target
- efficacy = bridge between brightness and load

This prevents the model from treating a 75W competitor product and a 72W Sunco product as different opportunities when both serve the same 9000lm brightness tier.

## Current Pilot

Pilot category:

```text
Ceiling Panels
```

Reason:

- Known recent NPI work exists.
- Competitor inventory rows already show movement.
- The category has spec-level decisions PMs care about: CCT, wattage, lumens, size, emergency backup, and pack quantity.

Current launcher behavior:

- `1B - Product Demand Ideation Generator.py` uses the same category picker as the original Step 1 flow.
- `Smart Lighting` is currently added as a 1B-only category owned by Manny. It is sourced from the PowerBI Families high-level category and should not be added to the main Step 1 category reference until validated.
- The current validated pilot category is Panels/Ceiling Panels because inventory movement snapshots exist there.
- Do not promote the 1B model into the default Step 1 flow until the workbook contract, Sunco coverage check, Stackline/Amazon validation behavior, and product-demand weighting are approved.

Smart Lighting caveat:

- Main `Recommendations` can be ecommerce/PDP-demand led today.
- Sunco coverage uses the local catalog cache plus the PowerBI Families classification cache.
- `Amazon Recommendations` will stay empty until the main Step 1/Stackline category path is also built for Smart.

## Workbook Validation

Every `1B` run validates the generated workbook before reporting success.

Validation checks:

- required Step 1 sheet names exist
- Step 2-facing headers match the original Step 1 contract
- `Amazon Recommendations` stays data-empty until real Stackline/Amazon-led rows are connected
- `Amazon Recommendations` reuses Step 1 rows and evidence; 1B only appends the product-demand overlay
- Step 2 can parse usable rows
- output remains under `product_demand_ideation`
- Excel can open/save the file

## File Tree

```text
product_demand_ideation/
  README.md
  src/
    product_demand_cli.py
    step1b_generator.py
    ecommerce_evidence.py
    luminaire_performance.py
    sunco_catalog_coverage.py
    sku_classification_cache.py
  cache/
    sunco_catalog.sqlite
    sku_classification.sqlite
  docs/
    integration_plan.md
    data_contract.md
  config/
    categories/
      ceiling_panels.json
  sql/
    ceiling_panels/
      competitor_inventory_daily.sql
      competitor_latest_products.sql
      marketplace_validation.sql
      sunco_catalog_coverage.sql
  experiments/
    ceiling_panels/
      exports/
      outputs/
```

## Promotion Rule

This tree should remain isolated until the Ceiling Panels pilot is reviewed.

After validation, add only a small opt-in bridge to the existing Category Ideation Generator. The current default behavior should remain unchanged unless the new model is explicitly enabled.
