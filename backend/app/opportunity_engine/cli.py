from __future__ import annotations

from pathlib import Path

from .categories import choose_category
from .gap_generator import generate_gap_workbook, load_category_data
from .ideation_template import generate_prd_ideation_workbook, latest_gap_workbook
from .paths import ProjectPaths
from .research_tool import (
    best_session_to_finalize,
    finalize_research_session,
    latest_prd_ideation_workbook,
    latest_session,
    prepare_research_session,
    publish_combined_report,
    session_status,
    write_research_collection_tasks,
)
from .utils import open_folder, open_path, prompt_choice, prompt_yes_no


def _print_issues(issues: list[str]) -> None:
    if not issues:
        print("Validation passed.")
        return
    print("\nValidation notes:")
    for issue in issues:
        print(f"- {issue}")


def _print_research_status(result: dict, session_root: Path) -> None:
    status = result.get("status") or {}
    summary = status.get("summary") or {}
    next_tasks = status.get("next_tasks") or []

    print("\nResearch session status:")
    print(f"- Session: {session_root}")
    if summary:
        row_count = summary.get("row_count")
        print(f"- Ideation rows: {row_count}")
        stage_counts = summary.get("stage_status_counts") or {}
        for stage in ["amazon_collection", "brick_and_mortar_collection", "brand_site_collection", "normalized", "analyzed", "reported"]:
            counts = stage_counts.get(stage)
            if counts:
                print(f"- {stage}: {counts}")

    if next_tasks:
        task_paths = write_research_collection_tasks(session_root, result)
        codex_task_path = task_paths.get("codex_task")
        claude_fallback_path = task_paths.get("claude_fallback")
        print("\nNo research workbook is created yet. Raw competitor collection is still pending.")
        print("If Stackline autofilled Amazon/Home Depot, the remaining tasks are optional web enrichment.")
        print("For the lowest-token run, skip web collection and choose 'Finalize latest prepared session' to build from available Stackline/local data.")
        print("Next collection tasks:")
        for task in next_tasks:
            print(
                f"- Row {task.get('row_number')} | {task.get('ideation_name')} | "
                f"{task.get('channel')} -> {task.get('output_file')}"
            )
        print("\nAsk Codex to complete this research task file:")
        if codex_task_path:
            print(codex_task_path)
        else:
            print(session_root / "instructions" / "1 - CODEX RESEARCH TASK.md")
        if claude_fallback_path:
            print("\nOptional fallback if Codex cannot complete the web collection:")
            print(claude_fallback_path)
        print("\nAfter those raw files are completed, run Step 3 again and choose:")
        print("Finalize latest prepared session")
    else:
        print("\nNo pending collection tasks were reported. Run Step 3 again and choose:")
        print("Finalize latest prepared session")


def run_step1(root: Path | str) -> None:
    paths = ProjectPaths.from_root(root)
    paths.ensure()
    category = choose_category(paths)
    force_refresh = prompt_yes_no("Force refresh even if cached data is under 30 days old?", default=False)
    data = load_category_data(paths, category)
    age = data.get("age_days")
    if age is None:
        print("\nNo dated cached data was found. A refresh should be run before final decision-making.")
    elif age > 30:
        print(f"\nCached data is {age} days old. A fresh collection should run; this script will still create a new timestamped workbook.")
    elif force_refresh:
        print(f"\nCached data is {age} days old. Force refresh was requested; this run will document that in the audit sheet.")
    output, issues = generate_gap_workbook(paths, category, force_refresh=force_refresh)
    print(f"\nStep 1 output:\n{output}")
    _print_issues(issues)
    open_path(output)


def run_step2(root: Path | str) -> None:
    paths = ProjectPaths.from_root(root)
    paths.ensure()
    category = choose_category(paths)
    latest = latest_gap_workbook(paths, category)
    selected = latest
    if latest is None:
        print("\nNo Step 1 output was found for this category.")
    else:
        print(f"\nLatest Step 1 workbook:\n{latest}")
        if not prompt_yes_no("Use this workbook?", default=True):
            selected = None

    if selected is None:
        manual = input("Paste the full path to the Step 1 gap workbook: ").strip().strip('"')
        selected = Path(manual).expanduser()
    output, issues = generate_prd_ideation_workbook(paths, category, selected)
    print(f"\nStep 2 output:\n{output}")
    _print_issues(issues)
    open_path(output)


def run_step3(root: Path | str) -> None:
    paths = ProjectPaths.from_root(root)
    paths.ensure()
    category = choose_category(paths)
    ready_session = best_session_to_finalize(paths, category)
    latest_existing_session = latest_session(paths, category)
    if ready_session:
        mode_options = [
            "Finalize ready session",
            "Check latest session status",
            "Prepare new research session from latest Step 2 workbook",
        ]
    else:
        mode_options = [
            "Prepare new research session from latest Step 2 workbook",
            "Finalize latest prepared session",
            "Check latest session status",
        ]
    mode = prompt_choice(
        mode_options,
        lambda item: item,
        "What should the research tool do?",
    )

    if mode.startswith("Prepare"):
        if ready_session:
            print(f"\nA session is already ready to finalize:\n{ready_session}")
            if not prompt_yes_no("Create a new research session anyway?", default=False):
                session = ready_session
                result = finalize_research_session(paths, session)
                published = publish_combined_report(paths, session)
                print(f"\nFinalized session:\n{session}")
                local_only = result.get("local_only_brand_site_completion") or {}
                if local_only.get("completed_artifact_count"):
                    print(
                        "Local-only finalize note: skipped brand-site enrichment for "
                        f"{local_only.get('completed_artifact_count')} row(s); no brand pages were fabricated."
                    )
                print(f"Log:\n{result.get('_log_path')}")
                if published:
                    print("\nPublished report(s):")
                    for report in published:
                        print(report)
                    open_path(published[0])
                else:
                    print("\nNo combined Excel report was found to publish yet. Check the session status for remaining collection tasks.")
                    open_folder(session)
                return

        latest = latest_prd_ideation_workbook(paths, category)
        if latest is None:
            raise FileNotFoundError(f"No Step 2 workbook found for {category.run_name}. Run Step 2 first.")
        print(f"\nLatest Step 2 workbook:\n{latest}")
        if not prompt_yes_no("Use this workbook?", default=True):
            manual = input("Paste the full path to the Step 2 PRD ideation workbook: ").strip().strip('"')
            latest = Path(manual).expanduser()
        result = prepare_research_session(paths, category, latest)
        session_root = Path(result.get("init", {}).get("session_root") or "")
        print(f"\nPrepared research session:\n{session_root}")
        autofill = result.get("stackline_autofill") or {}
        if autofill.get("prefilled_artifact_count"):
            print(
                "\nLocal Stackline autofill completed "
                f"{autofill.get('prefilled_artifact_count')} Amazon/Home Depot raw artifact(s)."
            )
        print(f"Log:\n{result.get('_log_path')}")
        if session_root.exists():
            status_result = session_status(paths, session_root)
            _print_research_status(status_result, session_root)
        if session_root.exists():
            task_file = session_root / "instructions" / "1 - CODEX RESEARCH TASK.md"
            open_path(task_file if task_file.exists() else session_root / "instructions")
        return

    session = ready_session if mode.startswith("Finalize ready") else latest_existing_session
    if session is None:
        raise FileNotFoundError(f"No prepared research session found for {category.run_name}.")

    if mode.startswith("Finalize"):
        print(f"\nSession selected for finalize:\n{session}")
        if not prompt_yes_no("Finalize this session?", default=True):
            manual = input("Paste the full path to the session folder: ").strip().strip('"')
            session = Path(manual).expanduser()
        result = finalize_research_session(paths, session)
        published = publish_combined_report(paths, session)
        print(f"\nFinalized session:\n{session}")
        local_only = result.get("local_only_brand_site_completion") or {}
        if local_only.get("completed_artifact_count"):
            print(
                "Local-only finalize note: skipped brand-site enrichment for "
                f"{local_only.get('completed_artifact_count')} row(s); no brand pages were fabricated."
            )
        print(f"Log:\n{result.get('_log_path')}")
        if published:
            print("\nPublished report(s):")
            for report in published:
                print(report)
            open_path(published[0])
        else:
            print("\nNo combined Excel report was found to publish yet. Check the session status for remaining collection tasks.")
            open_folder(session)
        return

    result = session_status(paths, session)
    print(f"\nStatus for:\n{session}")
    print(f"Log:\n{result.get('_log_path')}")
    _print_research_status(result, session)
    task_file = session / "instructions" / "1 - CODEX RESEARCH TASK.md"
    open_path(task_file if task_file.exists() else session / "instructions")
