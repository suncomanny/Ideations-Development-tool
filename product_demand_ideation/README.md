# Product Demand Ideation

This folder contains the demand-weighted Step 1 model for the Ideation Development tool.

The user-facing workflow remains:

- `1 - Category Ideation Generator.py`
- `2 - Ideation Template Generator.py`
- `3 - Ideation Research Tool.py`

`1 - Category Ideation Generator.py` is the production launcher for Step 1.

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

Redshift ecommerce competitor evidence now requires Redshift MCP. Local Redshift ODBC is not used by the production Step 1 flow because workstation DSN issues can produce empty or stale reports.

Postgres ODBC is not required for Manny's workstation. Sunco catalog coverage uses the saved local SQLite cache, a Postgres MCP refresh, or approved Postgres/Redshift export files.

Sunco catalog coverage is designed to run from a local SQLite cache:

```text
product_demand_ideation/cache/sunco_catalog.sqlite
```

Refresh it intentionally with:

```text
Refresh Product Demand Local Catalog Cache.py
```

Normal Step 1 workbook runs read from that local cache. If the cache does not exist yet, the tool seeds it through Postgres MCP and then continues. For Sunco's current catalog change rate, a weekly or monthly refresh is enough for routine A/B testing; refresh before leadership readouts or after known catalog/product-status updates.

The SharePoint workbook `Sku's Classification.xlsx` / `PowerBI Families` can also be cached locally for the latest PM/category/Series ownership layer:

```text
backend/source_data/sharepoint_exports/sku_classification/sku_classification.sqlite
```

Place the latest SharePoint export here:

```text
backend/source_data/sharepoint_exports/sku_classification/SkuClassification_PowerBIFamilies_latest.xlsx
```

Then refresh it:

```text
Refresh Product Demand SKU Classification Cache.py
```

You can still pass an explicit workbook path if you are testing a temporary export. This cache is optional at runtime. When present, Step 1 enriches Sunco active-catalog coverage checks with PowerBI category, PM responsible, and Series. Step 0 also uses the cache through `templates/powerbi_category_designation_map.csv` to align generated line-review queries with the same reporting categories.

After the classification cache refresh, update the shared run-category reference:

```text
Refresh Category Reference From PowerBI.py
```

Treat the newest SharePoint workbook as the source of truth because other PMs may update it after local PM cleanup work.

Category profiles can be checked without generating workbooks:

```text
Audit Product Demand Category Profiles.py
```

This uses Redshift MCP and writes a compact CSV audit under `product_demand_ideation/profile_audits/`. Use it before expanding Step 1 to a new category so noisy filters are fixed before leadership sees the output.

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

- `public.vw_competitors_scraping_latest` for latest competitor PDP title/spec/category/price/image/URL evidence
- `public.vw_competitors_inventory_daily` for PDP inventory movement and price movement by URL

The Redshift refresh path is:

```text
Redshift MCP -> product_demand_ideation/experiments/<category>/exports/*_ecommerce_competitor_evidence_*.json
```

If Redshift MCP is unavailable, the tool stops before writing a workbook instead of falling back to local ODBC. Normal Step 1 runs refresh ecommerce snapshots when the newest snapshot is not Redshift MCP-sourced or is older than 24 hours. Use `PRODUCT_DEMAND_ECOMMERCE_SNAPSHOT_MAX_AGE_HOURS` to change that freshness window.

When Redshift ecommerce PDP rows exist for the selected category, Step 1 uses them to lead the main `Recommendations` tab. Amazon-derived display rows stay in `Amazon Recommendations`.

For light-producing categories, the model also normalizes luminaire performance:

- lumens = user-facing brightness target
- wattage = installer/electrician load target
- efficacy = bridge between brightness and load

This prevents the model from treating a 75W competitor product and a 72W Sunco product as different opportunities when both serve the same 9000lm brightness tier.

## Current Model

Pilot category:

```text
Ceiling Panels
```

Reason:

- Known recent NPI work exists.
- Competitor inventory rows already show movement.
- The category has spec-level decisions PMs care about: CCT, wattage, lumens, size, emergency backup, and pack quantity.

Current launcher behavior:

- `1 - Category Ideation Generator.py` uses the production category picker.
- `Smart Lighting` is included as a Step 1 category owned by Manny. It is sourced from the PowerBI Families high-level category.
- The current validated smoke categories include Panels/Ceiling Panels, Vaportights, Wraparounds, and Grow Lights.
- The demand-weighted model is the default Step 1 flow once this branch is promoted.

Smart Lighting caveat:

- Main `Recommendations` can be ecommerce/PDP-demand led today.
- Sunco coverage uses the local catalog cache plus the PowerBI Families classification cache.
- `Amazon Recommendations` will stay empty until the main Step 1/Stackline category path is also built for Smart.

## Workbook Validation

Every Step 1 run validates the generated workbook before reporting success.

Validation checks:

- required Step 1 sheet names exist
- Step 2-facing headers match the original Step 1 contract
- `Amazon Recommendations` stays Stackline/Amazon-led
- `Amazon Recommendations` reuses Step 1 rows and evidence; Step 1 appends demand scoring to the research notes
- Step 2 can parse usable rows
- output remains under `product_demand_ideation`
- Excel can open/save the file

## File Tree

```text
product_demand_ideation/
  README.md
  src/
    product_demand_cli.py
    category_ideation_generator.py
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
