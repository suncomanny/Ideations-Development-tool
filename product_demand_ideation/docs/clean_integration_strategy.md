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

Use the production Step 1 launcher:

```text
1 - Category Ideation Generator.py
```

That gives PMs the standard workflow:

- `1 - Category Ideation Generator.py` = demand-weighted category ideation model
- `2 - Ideation Template Generator.py` = PM-kept rows converted to SKU-level ideations
- `3 - Ideation Research Tool.py` = detailed research workbook per ideation row

The launcher must be able to run without Codex chat assistance where local connectors are available. The final PM workflow should use Redshift MCP, Redshift ODBC fallback, Postgres MCP, or approved refreshed snapshots.

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

Step 1 output remains:

```text
outputs/Ideations/Gap Workbooks/
```

Backend run copies can also go to:

```text
product_demand_ideation/experiments/<category_slug>/outputs/
```

The Step 2 handoff copy belongs in the normal output tree.

## Data Access Separation

The tool should have two separate steps:

```text
Refresh source snapshots -> Generate workbook from snapshots
```

This matters because Redshift, Postgres MCP, and workbook generation have different failure modes.

Recommended production flow:

```text
Redshift MCP -> product_demand_ideation/experiments/<category_slug>/exports/ -> workbook output
```

Recommended fallback:

```text
Redshift ODBC or approved source cache -> product_demand_ideation/experiments/<category_slug>/exports/ -> workbook output
```

The workbook generator should not care whether the snapshot came from Redshift MCP, Redshift ODBC, Postgres MCP, or an approved export, as long as the snapshot follows the same data contract.

## Integration Phases

### Phase 1: Production Step 1 Model

Build:

```text
1 - Category Ideation Generator.py
```

This runs the demand-weighted Step 1 engine.

It should:

- support the production category picker
- run the combined model
- output workbook in the original workbook format
- publish the Step 2 handoff copy to the normal gap workbook folder
- validate the workbook contract before reporting success

### Phase 2: Backtest Against Prior Step 1

For Ceiling Panels, compare:

- prior Step 1 workbook
- demand-weighted Step 1 workbook
- manual NPI decisions

Focus comparison on:

- recommended specs
- Sunco coverage classification
- Amazon/Stackline support
- competitor inventory movement
- PM action usefulness

### Phase 3: Deployment Hardening

After validation, keep the PM-facing process as Step 1, Step 2, and Step 3, and move legacy/experimental wording out of the production branch.

## Branch Rule

All of this should stay on:

```text
codex/integration-step2-category-profiles
```

Do not merge into `main` until:

- Ceiling Panels pilot works
- workbook format is approved
- data team fixes or confirms latest competitor PDP export behavior
- standalone Redshift ODBC and Postgres MCP refreshes work outside Codex chat
- original Step 1 output remains unchanged in default mode
