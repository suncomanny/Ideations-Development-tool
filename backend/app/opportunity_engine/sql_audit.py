from __future__ import annotations

from pathlib import Path

from .paths import ProjectPaths


def collect_sql_text(paths: ProjectPaths, category_slug: str) -> str:
    sql_roots = [
        paths.cache / "ideation_data" / "_shared" / "sql",
        paths.cache / "ideation_data" / category_slug / "sql",
    ]

    chunks: list[str] = []
    for sql_root in sql_roots:
        if not sql_root.exists():
            continue
        for path in sorted(sql_root.glob("*.sql")):
            try:
                chunks.append(f"-- {path.name}\n{path.read_text(encoding='utf-8').strip()}")
            except UnicodeDecodeError:
                chunks.append(f"-- {path.name}\n<Could not read as UTF-8 text.>")
    return "\n\n".join(chunks).strip() or "SQL folder exists but contains no .sql files."
