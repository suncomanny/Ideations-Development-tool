# Clean Integration Strategy

## Recommendation

Do not make a full 1:1 copy of the entire Ideation Development tool.

Instead, create a separate product-demand branch of the tool with:

- a new isolated file tree
- a new user-facing launcher
- separate outputs
- separate config
- standalone database connectivity
- shared read-only utility reuse only where it is safe

This gives the cleanliness of a separate tool without the long-term problems of duplicating the whole codebase.

## Why Not Full 1:1 Copy?

A full copy sounds clean at first, but it creates drift:

- bug fixes in the original Step 1 would not automatically land in the copy
- workbook template improvements would need to be maintained twice
- category reference changes would need to be copied manually
- it becomes unclear which tool is the real source of truth
- two similar tools can produce different answers for reasons unrelated to the new inventory logic

The better pattern is:

```text
same repo + feature branch + isolated new engine + separate launcher + separate output folders
```

## Proposed User-Facing Shape

Keep the current scripts unchanged:

```text
0 - Refresh Backend Data.py
1 - Category Ideation Generator.py
2 - Ideation Template Generator.py
3 - Ideation Research Tool.py
```

Add a new launcher later:

```text
1B - Product Demand Ideation Generator.py
```

That gives PMs a clean separation:

- `1 - Category Ideation Generator.py` = original validated model
- `1B - Product Demand Ideation Generator.py` = new combined inventory + Stackline model

The new launcher must be able to run without Codex. Codex MCP can help build and validate the SQL, but the final PM workflow should use local ODBC/DSN connectivity or refreshed snapshots.

## Proposed File Tree

All new model work should remain under:

```text
product_demand_ideation/
```

Recommended structure:

```text
product_demand_ideation/
  README.md
  docs/
  config/
  sql/
  src/
    cli.py
    connections.py
    refresh_snapshots.py
    model.py
    parsers.py
    scoring.py
    workbook_writer.py
  experiments/
    ceiling_panels/
      exports/
      outputs/
```

## Output Separation

Original Step 1 output remains:

```text
outputs/Ideations/Gap Workbooks/
```

Product Demand output should go to:

```text
product_demand_ideation/experiments/<category_slug>/outputs/
```

Only after validation should promoted outputs move into the normal output tree.

## Data Access Separation

The tool should have two separate steps:

```text
Refresh source snapshots -> Generate workbook from snapshots
```

This matters because Redshift/Postgres credentials and workbook generation have different failure modes.

Recommended production flow:

```text
ODBC DSN or REDSHIFT_DSN -> product_demand_ideation/experiments/<category_slug>/exports/ -> workbook output
```

Recommended development fallback:

```text
Codex MCP -> product_demand_ideation/experiments/<category_slug>/exports/ -> workbook output
```

The workbook generator should not care whether the snapshot came from ODBC, DSN, or MCP, as long as the snapshot follows the same data contract.

## Integration Phases

### Phase 1: Shadow Tool

Build:

```text
1B - Product Demand Ideation Generator.py
```

This runs only the isolated product-demand engine.

It should:

- start with the Ceiling Panels pilot until workbook behavior is validated
- run the combined model
- output workbook in the original workbook format
- write only to the isolated product-demand output folder
- validate the workbook contract before reporting success

It should not call or modify the original Step 1 generator.

### Phase 1B: Category Picker Expansion

After the Ceiling Panels workbook contract is approved, update `1B - Product Demand Ideation Generator.py` to behave like the original Step 1 category picker.

The expanded behavior should:

- list available product-demand categories
- load category-specific config
- refresh or read that category's source snapshots
- write outputs under `product_demand_ideation/experiments/<category_slug>/outputs/`
- keep original Step 1 outputs untouched

### Phase 2: Backtest Against Original Step 1

For Ceiling Panels, compare:

- original Step 1 workbook
- product-demand workbook
- manual NPI decisions

Focus comparison on:

- recommended specs
- Sunco coverage classification
- Amazon/Stackline support
- competitor inventory movement
- PM action usefulness

### Phase 3: Optional Bridge

After validation, add an opt-in bridge to the original Step 1 flow.

The bridge should be disabled by default.

Possible behavior:

```text
Run original Step 1 only
Run Product Demand only
Run both and compare
```

Do not replace the original model until leadership accepts the new combined model.

## Branch Rule

All of this should stay on:

```text
codex/product-demand-ideation
```

Do not merge into `main` until:

- Ceiling Panels pilot works
- workbook format is approved
- data team fixes or confirms latest competitor PDP export behavior
- standalone ODBC/DSN refresh works outside Codex
- original Step 1 output remains unchanged in default mode



