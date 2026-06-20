from __future__ import annotations

from dataclasses import dataclass
from pathlib import Path


@dataclass(frozen=True)
class ProjectPaths:
    root: Path
    templates: Path
    outputs: Path
    backend: Path
    app: Path
    cache: Path
    logs: Path
    research_sessions: Path
    source_data: Path
    prd_tool: Path

    @classmethod
    def from_root(cls, root: Path | str) -> "ProjectPaths":
        resolved = Path(root).resolve()
        backend = resolved / "backend"
        return cls(
            root=resolved,
            templates=resolved / "templates",
            outputs=resolved / "outputs",
            backend=backend,
            app=backend / "app",
            cache=backend / "cache",
            logs=backend / "logs",
            research_sessions=backend / "research_sessions",
            source_data=backend / "source_data",
            prd_tool=backend / "app" / "prd_research_tool",
        )

    @property
    def gap_outputs(self) -> Path:
        return self.outputs / "Ideations" / "Gap Workbooks"

    def gap_category_outputs(self, category_slug: str) -> Path:
        return self.gap_outputs / category_slug

    @property
    def prd_ideation_outputs(self) -> Path:
        return self.outputs / "Ideations" / "PRD Ideation Workbooks"

    def prd_ideation_category_outputs(self, category_slug: str) -> Path:
        return self.prd_ideation_outputs / category_slug

    @property
    def research_report_outputs(self) -> Path:
        return self.outputs / "Research" / "Reports"

    def research_report_category_outputs(self, category_slug: str) -> Path:
        return self.research_report_outputs / category_slug

    @property
    def leadership_deck_outputs(self) -> Path:
        return self.outputs / "Leadership Decks" / "Gate 0"

    def leadership_deck_category_outputs(self, category_slug: str) -> Path:
        return self.leadership_deck_outputs / category_slug

    def ensure(self) -> None:
        for path in [
            self.templates,
            self.gap_outputs,
            self.prd_ideation_outputs,
            self.research_report_outputs,
            self.leadership_deck_outputs,
            self.cache / "ideation_data",
            self.cache / "gate0_decks",
            self.cache / "images",
            self.logs,
            self.research_sessions,
            self.backend / "legacy_research_sessions",
            self.source_data,
        ]:
            path.mkdir(parents=True, exist_ok=True)
