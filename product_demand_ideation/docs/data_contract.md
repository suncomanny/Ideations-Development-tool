# Data Contract

## Source Families

The product-demand model depends on four source families.

## Connection Contract

The production version should keep database refresh separate from workbook generation.

Preferred production access:

- Redshift through MCP when running from Codex or another MCP-capable runner
- Redshift through a local ODBC DSN or `REDSHIFT_DSN` only as the fallback when MCP is unavailable
- Postgres through MCP for Sunco catalog and line-review refreshes
- secrets stored in the machine-local file, not in SharePoint:

```text
C:\Users\<user>\.sunco_ideation_development\.env
```

Development fallback:

- Redshift MCP, local Redshift ODBC fallback, Postgres MCP, and approved cache/export files are the supported integrated runtime sources.
- Workbook generation should read refreshed snapshot files where possible; the refresh step owns live connector access.

The workbook generator should read from local snapshot files written by a refresh step. That keeps database access separate from workbook creation and makes the tool easier to run, debug, and hand off.

## 1. Competitor Latest PDP / Spec Data

Primary source:

```text
public.v_competitors_scrapping_latest
```

Use for:

- PDP URL
- product name/title
- competitor domain
- brand
- SKU/model/MPN/GTIN
- price
- product category/type
- description/specification text
- wattage, lumens, CCT, voltage, dimming, reviews, rating where available

Known issue:

- The current Redshift path can hit a Parquet BOOLEAN error when selecting boolean fields.
- Avoid `select *`.
- If explicit-column reads still fail, ask data team to cast the boolean field safely in the view.

## 2. Competitor Inventory Movement

Primary source:

```text
public.v_competitors_inventory_daily
```

Use for:

- `stock_qty`
- `prev_stock_qty`
- `stock_qty_delta`
- `scrape_date`
- `prev_scrape_date`
- `days_since_prev`
- `stock_status`
- `availability`
- `price_delta`

Derived signals:

- total observed stock decrease
- decrease event count
- restock event count
- flat event count
- latest stock quantity
- inventory velocity score
- data quality/confidence flags

## 3. Stackline / Amazon Marketplace Validation

Primary sources:

```text
public.tb_scrapping_bsr_gold
public.competitorpricinganalysis
public.tb_stackline_atlas_products
```

Use for:

- Amazon recommendations
- validation for competitor inventory-led recommendations
- BSR/rank
- rating and review count
- price
- ASIN/product identity
- title/spec evidence

Rule:

- Stackline/Amazon data should be used across all models.

## 4. Sunco Catalog Coverage

Starting sources:

```text
public.sunco_pilot_amazon
public.sunco_pilot_shopify
```

Use for:

- direct match
- partial match
- missing gap
- already covered but weak
- Amazon vs Sunco.com coverage split

## Output Contract

User-facing output must be workbook format.

Required tabs:

- `Summary`
- `Recommendations`
- `Sources and Audit`
- `Amazon Recommendations`
- `Amazon Source Audit`

Backend/audit artifacts may include CSVs, but the main output should not be CSV-only.
