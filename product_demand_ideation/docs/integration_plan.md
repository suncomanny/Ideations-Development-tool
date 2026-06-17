# Integration Plan

## Objective

Add an inventory-led competitor demand model to the Ideation Development tool without disrupting the validated main workflow.

The new model should initially live only under:

```text
product_demand_ideation/
```

## Why This Is Separate

The existing tool already has a validated Category Ideation Generator. This adaptation introduces a new product-level signal: competitor inventory movement.

To avoid mixing outputs or changing behavior too early:

- do not modify the existing generator scripts during the isolated pilot
- do not write pilot outputs into existing validated output folders
- do not rename or change existing workbook templates

## Combined Recommendation Model

### Competitor Recommendations

Lead with:

- competitor PDP inventory movement

Validate with:

- Stackline/Amazon marketplace demand
- Sunco catalog coverage

Workbook destination:

- `Recommendations`

### Amazon Recommendations

Lead with:

- Stackline/Amazon demand

Validate with:

- competitor inventory movement where available
- Sunco catalog coverage

Workbook destination:

- `Amazon Recommendations`

## Pilot Scope

Category:

```text
Ceiling Panels
```

Required output:

- same workbook-style structure as the original validated tool
- no CSV-only deliverable
- CSVs may exist only as backend/audit artifacts

Required workbook tabs:

- `Summary`
- `Recommendations`
- `Sources and Audit`
- `Amazon Recommendations`
- `Amazon Source Audit`

## Future Bridge Into Existing Tool

Only after the pilot is validated, add a small opt-in bridge into:

```text
1 - Category Ideation Generator.py
```

The bridge should be disabled by default.

Possible switch:

```text
--enable-product-demand
```

Expected behavior:

- default generator remains unchanged
- product-demand model runs only when explicitly enabled
- outputs remain separated unless intentionally promoted


