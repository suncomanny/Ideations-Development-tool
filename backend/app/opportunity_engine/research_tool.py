from __future__ import annotations

import json
import shutil
import subprocess
import sys
from datetime import datetime, timezone
from pathlib import Path
from typing import Any

from .categories import Category
from .ideation_template import latest_gap_workbook
from .paths import ProjectPaths
from .refresh_redshift_stackline_cache import ensure_stackline_cache_for_category
from .utils import newest_file, timestamp


def latest_prd_ideation_workbook(paths: ProjectPaths, category: Category) -> Path | None:
    return newest_file(
        paths.prd_ideation_category_outputs(category.slug),
        [f"{category.slug}_prd_ideations_*.xlsx"],
    ) or newest_file(paths.prd_ideation_outputs, [f"{category.slug}_prd_ideations_*.xlsx"])


def latest_session(paths: ProjectPaths, category: Category) -> Path | None:
    candidates = candidate_sessions(paths, category)
    if not candidates:
        return None
    return max(candidates, key=lambda p: p.stat().st_mtime)


def candidate_sessions(paths: ProjectPaths, category: Category) -> list[Path]:
    candidates = [
        path for path in paths.research_sessions.glob(f"{category.slug}_*")
        if path.is_dir() and (path / "manifest.json").exists()
    ]
    return candidates


def artifact_status_from_file(path: Path) -> str:
    if not path.exists():
        return "pending"
    if path.suffix.lower() != ".json":
        return "complete"
    try:
        payload = json.loads(path.read_text(encoding="utf-8"))
    except Exception:
        return "blocked"
    return str(payload.get("artifact_status") or "").strip().lower() or "pending"


def session_readiness(session_root: Path) -> dict[str, Any]:
    """Return file-derived readiness, avoiding stale manifest status after handoffs."""
    manifest_path = session_root / "manifest.json"
    if not manifest_path.exists():
        return {"raw_complete": False, "reported_complete": False, "row_count": 0}

    try:
        manifest = json.loads(manifest_path.read_text(encoding="utf-8"))
    except Exception:
        return {"raw_complete": False, "reported_complete": False, "row_count": 0}

    raw_stage_keys = ["amazon_collection", "brick_and_mortar_collection", "brand_site_collection"]
    rows = manifest.get("rows") or []
    raw_statuses: list[str] = []
    report_statuses: list[str] = []
    for row in rows:
        artifacts = row.get("artifacts") or {}
        for stage_key in raw_stage_keys:
            rel = artifacts.get(stage_key)
            raw_statuses.append(artifact_status_from_file(session_root / rel) if rel else "pending")
        report_rel = artifacts.get("reported")
        report_statuses.append(artifact_status_from_file(session_root / report_rel) if report_rel else "pending")

    raw_complete = bool(rows) and all(status == "complete" for status in raw_statuses)
    reported_complete = bool(rows) and all(status == "complete" for status in report_statuses)
    return {
        "raw_complete": raw_complete,
        "reported_complete": reported_complete,
        "row_count": len(rows),
        "raw_statuses": raw_statuses,
        "report_statuses": report_statuses,
    }


def best_session_to_finalize(paths: ProjectPaths, category: Category) -> Path | None:
    """Prefer a ready-to-finalize session over the newest accidentally prepared one."""
    ready = []
    for session in candidate_sessions(paths, category):
        readiness = session_readiness(session)
        if readiness.get("raw_complete") and not readiness.get("reported_complete"):
            ready.append(session)
    if ready:
        return max(ready, key=lambda p: p.stat().st_mtime)
    return latest_session(paths, category)


def _extract_json(stdout: str) -> dict[str, Any]:
    text = stdout.strip()
    if not text:
        return {}
    first = text.find("{")
    last = text.rfind("}")
    if first >= 0 and last >= first:
        return json.loads(text[first:last + 1])
    return {"raw_stdout": stdout}


def run_orchestrator(paths: ProjectPaths, args: list[str], log_name: str) -> dict[str, Any]:
    tool = paths.prd_tool / "tools" / "research_orchestrator.py"
    if not tool.exists():
        raise FileNotFoundError(f"Missing copied PRD Research orchestrator: {tool}")

    paths.logs.mkdir(parents=True, exist_ok=True)
    log_path = paths.logs / log_name
    command = [sys.executable, str(tool), *args]
    completed = subprocess.run(
        command,
        cwd=str(tool.parent),
        text=True,
        stdout=subprocess.PIPE,
        stderr=subprocess.PIPE,
        check=False,
    )
    log_path.write_text(
        "COMMAND\n"
        + " ".join(command)
        + "\n\nSTDOUT\n"
        + completed.stdout
        + "\n\nSTDERR\n"
        + completed.stderr,
        encoding="utf-8",
    )
    if completed.returncode != 0:
        raise RuntimeError(f"PRD Research tool failed. See log: {log_path}")
    result = _extract_json(completed.stdout)
    result["_log_path"] = str(log_path)
    return result


def run_session_manager(paths: ProjectPaths, args: list[str], log_name: str) -> dict[str, Any]:
    tool = paths.prd_tool / "tools" / "research_session_manager.py"
    if not tool.exists():
        raise FileNotFoundError(f"Missing copied PRD Research session manager: {tool}")

    paths.logs.mkdir(parents=True, exist_ok=True)
    log_path = paths.logs / log_name
    command = [sys.executable, str(tool), *args]
    completed = subprocess.run(
        command,
        cwd=str(tool.parent),
        text=True,
        stdout=subprocess.PIPE,
        stderr=subprocess.PIPE,
        check=False,
    )
    log_path.write_text(
        "COMMAND\n"
        + " ".join(command)
        + "\n\nSTDOUT\n"
        + completed.stdout
        + "\n\nSTDERR\n"
        + completed.stderr,
        encoding="utf-8",
    )
    if completed.returncode != 0:
        raise RuntimeError(f"PRD Research session manager failed. See log: {log_path}")
    result = _extract_json(completed.stdout)
    result["_log_path"] = str(log_path)
    return result


def utc_now() -> str:
    return datetime.now(timezone.utc).replace(microsecond=0).isoformat()


def read_json(path: Path) -> dict[str, Any]:
    return json.loads(path.read_text(encoding="utf-8"))


def write_json(path: Path, payload: dict[str, Any]) -> None:
    path.parent.mkdir(parents=True, exist_ok=True)
    path.write_text(json.dumps(payload, indent=2), encoding="utf-8")


def _seed_url(channel: str, seed: dict[str, Any]) -> str | None:
    url = seed.get("url")
    if url:
        return str(url)
    sku = seed.get("sku") or seed.get("retailer_sku")
    if not sku:
        return None
    if channel == "amazon":
        return f"https://www.amazon.com/dp/{sku}"
    if channel == "home_depot":
        return f"https://www.homedepot.com/p/{sku}"
    return None


def _stackline_raw_item(seed: dict[str, Any], channel: str, row_number: int, index: int) -> dict[str, Any] | None:
    title = seed.get("product_title") or seed.get("title")
    brand = seed.get("brand")
    url = _seed_url(channel, seed)
    if not title or not brand or not url:
        return None

    domain = "amazon.com" if channel == "amazon" else "homedepot.com" if channel == "home_depot" else None
    return {
        "candidate_id": f"stackline_r{row_number:03d}_{channel}_{index:03d}",
        "source_channel": channel,
        "source_domain": domain,
        "collection_method": "stackline_seed",
        "brand": str(brand),
        "product_title": str(title),
        "model_number": seed.get("model_number"),
        "sku": seed.get("sku") or seed.get("retailer_sku"),
        "variant": None,
        "pack_quantity": 1,
        "url": url,
        "price": seed.get("avg_retail_price"),
        "currency": "USD",
        "wattage": None,
        "lumens": None,
        "cct": None,
        "cri": None,
        "voltage": None,
        "dimmable": None,
        "dimming_type": None,
        "certifications": [],
        "features": [],
        "rating": seed.get("rating"),
        "review_count": seed.get("review_count"),
        "availability": None,
        "match_confidence": 0.85,
        "match_notes": "Autofilled from Stackline segment leader data. Use web enrichment only for missing specs, certifications, and listing-copy details.",
        "extraction_notes": "Local Stackline autofill; no live product page was fetched for this item.",
        "raw_observations": [
            f"Stackline units_sold={seed.get('units_sold')}",
            f"Stackline sales_share_pct={seed.get('sales_share_pct')}",
        ],
    }


def _write_stackline_raw_artifact(
    *,
    session_root: Path,
    manifest: dict[str, Any],
    row: dict[str, Any],
    packet: dict[str, Any],
    stage_key: str,
    channel: str,
    seeds: list[dict[str, Any]],
) -> bool:
    artifact_rel = row.get("artifacts", {}).get(stage_key)
    if not artifact_rel or not seeds:
        return False

    artifact_path = session_root / artifact_rel
    if artifact_path.exists():
        existing = read_json(artifact_path)
        if existing.get("artifact_status") == "complete" and existing.get("items"):
            return False
    else:
        existing = {}

    items = [
        item
        for index, seed in enumerate(seeds, start=1)
        if (item := _stackline_raw_item(seed, channel, row["row_number"], index))
    ]
    if not items:
        return False

    artifact_type = "amazon_raw" if stage_key == "amazon_collection" else "brick_and_mortar_raw"
    source_group = "amazon" if stage_key == "amazon_collection" else "brick_and_mortar"
    payload = {
        "schema_version": existing.get("schema_version") or "2026-04-15",
        "artifact_type": artifact_type,
        "artifact_status": "complete",
        "batch_id": manifest.get("batch_id") or session_root.name,
        "row_number": row["row_number"],
        "ideation_name": row.get("ideation_name") or packet.get("identity", {}).get("ideation_name"),
        "expected_owner": "codex",
        "source_channel_group": source_group,
        "packet_file": row.get("artifacts", {}).get("packet"),
        "queries_used": [],
        "items": items,
        "summary": {
            "autofill_source": "stackline_packet_seeds",
            "item_count": len(items),
            "channel": channel,
        },
        "notes": [
            "Autofilled locally by Codex from Stackline-backed packet seeds to reduce Claude collection cost.",
            "These are market-intelligence seeds, not live page fetches; use optional web enrichment for missing detailed specs/certifications.",
        ],
        "blocking_issues": [],
        "updated_at": utc_now(),
    }
    write_json(artifact_path, payload)
    return True


def prefill_stackline_raw_artifacts(session_root: Path) -> dict[str, Any]:
    """Use Stackline packet seeds to complete raw Amazon/Home Depot artifacts locally."""
    manifest_path = session_root / "manifest.json"
    if not manifest_path.exists():
        return {"prefilled_artifact_count": 0, "rows": []}

    manifest = read_json(manifest_path)
    changed_rows = []
    changed_count = 0
    for row in manifest.get("rows", []):
        packet_rel = row.get("artifacts", {}).get("packet")
        if not packet_rel:
            continue
        packet_path = session_root / packet_rel
        if not packet_path.exists():
            continue
        packet = read_json(packet_path)
        plan = packet.get("research_plan") or {}
        market = plan.get("market_intelligence") or {}
        if plan.get("collection_mode") != "stackline_first" and not market.get("matched"):
            continue

        row_changes = []
        amazon_seeds = (plan.get("amazon") or {}).get("competitor_seeds") or []
        if _write_stackline_raw_artifact(
            session_root=session_root,
            manifest=manifest,
            row=row,
            packet=packet,
            stage_key="amazon_collection",
            channel="amazon",
            seeds=amazon_seeds,
        ):
            changed_count += 1
            row_changes.append("amazon_collection")

        home_depot_seeds = (
            ((plan.get("brick_and_mortar") or {}).get("home_depot") or {}).get("competitor_seeds")
            or []
        )
        if _write_stackline_raw_artifact(
            session_root=session_root,
            manifest=manifest,
            row=row,
            packet=packet,
            stage_key="brick_and_mortar_collection",
            channel="home_depot",
            seeds=home_depot_seeds,
        ):
            changed_count += 1
            row_changes.append("brick_and_mortar_collection")

        if row_changes:
            changed_rows.append({"row_number": row["row_number"], "stages": row_changes})

    return {"prefilled_artifact_count": changed_count, "rows": changed_rows}


def complete_optional_brand_site_artifacts(session_root: Path) -> dict[str, Any]:
    """Allow a local-only finalize when Amazon and Home Depot evidence are already complete."""
    manifest_path = session_root / "manifest.json"
    if not manifest_path.exists():
        return {"completed_artifact_count": 0, "rows": []}

    manifest = read_json(manifest_path)
    changed_rows = []
    changed_count = 0
    for row in manifest.get("rows", []):
        artifacts = row.get("artifacts") or {}
        amazon_rel = artifacts.get("amazon_collection")
        brick_rel = artifacts.get("brick_and_mortar_collection")
        brand_rel = artifacts.get("brand_site_collection")
        if not amazon_rel or not brick_rel or not brand_rel:
            continue
        if artifact_status_from_file(session_root / amazon_rel) != "complete":
            continue
        if artifact_status_from_file(session_root / brick_rel) != "complete":
            continue
        brand_path = session_root / brand_rel
        if artifact_status_from_file(brand_path) == "complete":
            continue

        payload = {
            "schema_version": "2026-04-15",
            "artifact_type": "brand_sites_raw",
            "artifact_status": "complete",
            "batch_id": manifest.get("batch_id") or session_root.name,
            "row_number": row["row_number"],
            "ideation_name": row.get("ideation_name"),
            "expected_owner": "codex",
            "source_channel_group": "brand_sites",
            "packet_file": artifacts.get("packet"),
            "queries_used": [],
            "items": [],
            "summary": {
                "collection_skipped": True,
                "reason": "Local-only finalize requested/allowed after Amazon and Home Depot evidence were completed.",
                "impact": "Brand-site enrichment is absent; report confidence should remain directional where brand-site specs would matter.",
            },
            "notes": [
                "Brand-site collection was explicitly skipped for a no-Claude/local-only clean run.",
                "No brand product pages were fetched or fabricated in this artifact.",
                "Use Claude or a future web collector later if official brand-site specs are needed.",
            ],
            "blocking_issues": [],
            "updated_at": utc_now(),
        }
        write_json(brand_path, payload)
        changed_count += 1
        changed_rows.append(row["row_number"])

    return {"completed_artifact_count": changed_count, "rows": changed_rows}


def prepare_research_session(paths: ProjectPaths, category: Category, workbook: Path | None = None) -> dict[str, Any]:
    selected = workbook or latest_prd_ideation_workbook(paths, category)
    if selected is None:
        raise FileNotFoundError(f"No Step 2 PRD ideation workbook found for {category.run_name}. Run Step 2 first.")

    stackline_preflight = ensure_stackline_cache_for_category(paths, category)
    session_name = f"{category.slug}_{timestamp()}"
    args = [
        "prepare",
        str(selected),
        "--session-name",
        session_name,
        "--output-root",
        str(paths.research_sessions),
        "--limit",
        "10",
    ]
    result = run_orchestrator(paths, args, f"research_prepare_{session_name}.log")
    result["stackline_preflight"] = stackline_preflight
    session_root_value = (result.get("init") or {}).get("session_root")
    if session_root_value:
        session_root = Path(session_root_value)
        prefill = prefill_stackline_raw_artifacts(session_root)
        result["stackline_autofill"] = prefill
        if prefill.get("prefilled_artifact_count"):
            result["stackline_autofill_update"] = run_session_manager(
                paths,
                ["update", str(session_root)],
                f"research_autofill_update_{session_name}.log",
            )
    return result


def finalize_research_session(paths: ProjectPaths, session_root: Path) -> dict[str, Any]:
    local_only = complete_optional_brand_site_artifacts(session_root)
    if local_only.get("completed_artifact_count"):
        run_session_manager(
            paths,
            ["update", str(session_root)],
            f"research_local_only_update_{session_root.name}_{timestamp()}.log",
        )
    args = [
        "finalize",
        str(session_root),
    ]
    result = run_orchestrator(paths, args, f"research_finalize_{session_root.name}_{timestamp()}.log")
    if local_only.get("completed_artifact_count"):
        result["local_only_brand_site_completion"] = local_only
    return result


def session_status(paths: ProjectPaths, session_root: Path) -> dict[str, Any]:
    return run_orchestrator(paths, ["status", str(session_root), "--limit", "10"], f"research_status_{session_root.name}_{timestamp()}.log")


def write_research_collection_tasks(session_root: Path, status_result: dict[str, Any]) -> dict[str, Path | None]:
    status = status_result.get("status") or {}
    tasks = status.get("next_tasks") or []
    if not tasks:
        return {"codex_task": None, "claude_fallback": None, "task_bundle": None}

    instructions = session_root / "instructions"
    instructions.mkdir(parents=True, exist_ok=True)
    support = instructions / "_support"
    support.mkdir(parents=True, exist_ok=True)

    for helper_name in [
        "CLAUDE_NEXT.md",
        "CODEX_NEXT.md",
        "COLLECTOR_NEXT.md",
        "STEP4_PROMPT.md",
        "COPY_TO_CLAUDE.md",
        "1 - COPY THIS PROMPT TO CLAUDE.md",
    ]:
        helper_path = instructions / helper_name
        if helper_path.exists():
            destination = support / helper_name
            if destination.exists():
                destination.unlink()
            helper_path.replace(destination)

    codex_output = instructions / "1 - CODEX RESEARCH TASK.md"
    claude_output = instructions / "2 - OPTIONAL CLAUDE FALLBACK PROMPT.md"
    task_lines = []
    output_files = []
    packet_files = []
    for index, task in enumerate(tasks, start=1):
        packet = task.get("packet_file")
        out_file = task.get("output_file")
        task_lines.append(
            f"{index}. Row {task.get('row_number')} | {task.get('ideation_name')} | "
            f"channel `{task.get('channel')}` | packet `{packet}` | output `{out_file}`"
        )
        if packet:
            packet_files.append(str(packet))
        if out_file:
            output_files.append(str(out_file))

    row_list = ",".join(
        dict.fromkeys(str(task.get("row_number")) for task in tasks if task.get("row_number"))
    )
    manager = session_root.parents[1] / "app" / "prd_research_tool" / "tools" / "research_session_manager.py"
    task_bundle: dict[str, Any] = {
        "session_root": str(session_root),
        "rules": [
            "Complete only the tasks listed in this file.",
            "Use the row packet only if this task bundle is insufficient.",
            "Write only the listed output files.",
            "Do not normalize, dedupe, rank, analyze, or create the final workbook.",
            "Return control to the PM after raw collection so Step 3 can be finalized intentionally.",
        ],
        "schemas": {
            "collection_artifact": str(session_root / "schemas" / "collection-artifact.schema.json"),
            "competitor_result": str(session_root / "schemas" / "competitor-result.schema.json"),
        },
        "support_prompt": str(support / "STEP4_PROMPT.md"),
        "tasks": [],
        "validation_command": f'python "{manager}" validate "{session_root}" --rows {row_list}',
        "manifest_update_command": f'python "{manager}" update "{session_root}"',
    }

    for task in tasks:
        packet_path = session_root / str(task.get("packet_file"))
        packet_payload = read_json(packet_path) if packet_path.exists() else {}
        plan = packet_payload.get("research_plan") or {}
        channel = task.get("channel")
        channel_plan: dict[str, Any] = {}
        if channel == "amazon":
            amazon_plan = plan.get("amazon") or {}
            channel_plan = {
                "queries": (amazon_plan.get("queries") or [])[:3],
                "competitor_seed_count": len(amazon_plan.get("competitor_seeds") or []),
                "stackline_primary": amazon_plan.get("stackline_primary"),
            }
        elif channel == "brick_and_mortar":
            bm_plan = plan.get("brick_and_mortar") or {}
            home_depot = bm_plan.get("home_depot") or {}
            channel_plan = {
                "home_depot_queries": (home_depot.get("queries") or [])[:3],
                "home_depot_seed_count": len(home_depot.get("competitor_seeds") or []),
                "stackline_primary": home_depot.get("stackline_primary"),
            }
        elif channel == "brand_sites":
            brand_watchlist = []
            for entry in plan.get("brand_watchlist") or []:
                if not isinstance(entry, dict) or not entry.get("brand"):
                    continue
                brand_watchlist.append(
                    {
                        "brand": entry.get("brand"),
                        "source": entry.get("source"),
                        "priority": entry.get("priority"),
                    }
                )
                if len(brand_watchlist) >= 8:
                    break
            channel_plan = {
                "brand_watchlist": brand_watchlist,
                "task": "Find official brand/product pages for the closest competitors only; do not re-run Amazon/Home Depot collection already autofilled from Stackline.",
            }

        task_bundle["tasks"].append(
            {
                "row_number": task.get("row_number"),
                "ideation_name": task.get("ideation_name"),
                "channel": channel,
                "packet_file": str(packet_path),
                "output_file": str(session_root / str(task.get("output_file"))),
                "must_validate": {
                    "features": ((plan.get("must_validate") or {}).get("features") or [])[:12],
                    "certifications": ((plan.get("must_validate") or {}).get("certifications") or [])[:4],
                },
                "target_price_band": plan.get("target_price_band"),
                "channel_plan": channel_plan,
            }
        )

    task_bundle_path = support / "CODEX_RESEARCH_TASKS.json"
    write_json(task_bundle_path, task_bundle)

    codex_content = f"""# Codex Research Task

Complete the pending raw competitor collection tasks for this prepared Sunco Step 3 research session.

Session root:
`{session_root}`

Task bundle:
`{task_bundle_path}`

What to do:
1. Read the task bundle first.
2. Complete only the listed row/channel tasks.
3. Use the packet file only when the task bundle does not contain enough detail.
4. Verify each competitor/listing URL is real and relevant before writing it.
5. Write only the listed raw output files.
6. Run the validation command from the task bundle.
7. Run the manifest update command from the task bundle.
8. Stop after raw collection and validation. Do not normalize, analyze, or create the final workbook unless the PM explicitly asks.

Pending tasks:
{chr(10).join(task_lines)}

Rules:
- Use real product pages or real category/listing pages only.
- Do not invent products, prices, certifications, dimensions, pack sizes, URLs, or images.
- If a channel has no reliable match, write a blocked or complete-empty raw artifact with notes instead of fabricating.
- Prefer official brand/PDP pages for brand-site tasks.
- Keep Amazon/Home Depot tasks narrow to the ideation row and channel in the task bundle.

After Codex finishes these raw files, the PM should run:
`3 - Ideation Research Tool.py`

Then choose:
`Finalize latest prepared session`
"""
    codex_output.write_text(codex_content, encoding="utf-8")

    claude_content = f"""# Optional Claude Fallback Prompt

Run the Sunco PRD Research collection tasks listed in this task bundle:

`{task_bundle_path}`

Rules:
- Complete only the tasks in `CODEX_RESEARCH_TASKS.json`.
- Read the row packet only if the task bundle is insufficient.
- Write only the listed output files.
- Run the validation and update commands from the task bundle.
- Stop after raw collection. Do not normalize, analyze, or create the final workbook.
"""
    claude_output.write_text(claude_content, encoding="utf-8")
    return {"codex_task": codex_output, "claude_fallback": claude_output, "task_bundle": task_bundle_path}


def write_claude_collection_prompt(session_root: Path, status_result: dict[str, Any]) -> Path | None:
    """Backward-compatible wrapper for older callers.

    New Step 3 sessions are Codex-first. This returns the optional Claude fallback
    path only for callers that still expect a Claude prompt.
    """
    return write_research_collection_tasks(session_root, status_result).get("claude_fallback")


def _candidate_report_paths(session_root: Path) -> list[Path]:
    report_dirs = [
        session_root / "reports",
        session_root / "Research Reports",
        session_root / "output",
    ]
    files: list[Path] = []
    for folder in report_dirs:
        if folder.exists():
            files.extend(folder.rglob("*.xlsx"))
    files.extend(session_root.glob("*.xlsx"))
    return [path for path in files if path.is_file() and not path.name.startswith("~$")]


def _preferred_publish_report_paths(session_root: Path) -> list[Path]:
    """Return the user-facing Step 3 workbooks to publish for a session."""
    reports_dir = session_root / "reports"
    preferred = [
        reports_dir / f"{session_root.name}_completed_rows.xlsx",
    ]
    existing = [path for path in preferred if path.exists() and path.is_file()]
    if existing:
        return existing
    reports = _candidate_report_paths(session_root)
    if not reports:
        return []
    return sorted(reports, key=lambda p: p.stat().st_size, reverse=True)[:1]


def publish_combined_report(paths: ProjectPaths, session_root: Path) -> list[Path]:
    reports = _preferred_publish_report_paths(session_root)
    if not reports:
        return []

    paths.research_report_outputs.mkdir(parents=True, exist_ok=True)
    published: list[Path] = []
    for report in reports:
        destination = paths.research_report_outputs / report.name
        if destination.exists():
            destination = paths.research_report_outputs / f"{report.stem}_{timestamp()}{report.suffix}"
        shutil.copy2(report, destination)
        published.append(destination)
    return published
