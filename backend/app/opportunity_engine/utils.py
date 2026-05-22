from __future__ import annotations

import os
import platform
import re
import subprocess
from datetime import datetime
from pathlib import Path
from typing import Iterable, Sequence, TypeVar


T = TypeVar("T")


def slugify(value: str) -> str:
    text = value.strip().lower().replace("&", "and")
    text = text.replace("+", "plus").replace("/", " ")
    text = re.sub(r"[^a-z0-9]+", "_", text)
    return text.strip("_") or "item"


def timestamp(fmt: str = "%Y-%m-%d_%H%M%S") -> str:
    return datetime.now().strftime(fmt)


def prompt_yes_no(question: str, default: bool = False) -> bool:
    suffix = "Y/n" if default else "y/N"
    answer = input(f"{question} ({suffix}): ").strip().lower()
    if not answer:
        return default
    return answer in {"y", "yes"}


def prompt_choice(items: Sequence[T], labeler, prompt: str) -> T:
    if not items:
        raise ValueError("No options are available.")

    print()
    for index, item in enumerate(items, start=1):
        print(f"{index:>2}. {labeler(item)}")

    while True:
        raw = input(f"\n{prompt} [1-{len(items)}]: ").strip()
        try:
            selected = int(raw)
        except ValueError:
            print("Enter a number from the list.")
            continue
        if 1 <= selected <= len(items):
            return items[selected - 1]
        print("That number is not in the list.")


def open_folder(path: Path) -> None:
    folder = path if path.is_dir() else path.parent
    try:
        system = platform.system().lower()
        if system == "windows":
            os.startfile(str(folder))  # type: ignore[attr-defined]
        elif system == "darwin":
            subprocess.Popen(["open", str(folder)])
        else:
            subprocess.Popen(["xdg-open", str(folder)])
    except Exception as exc:  # pragma: no cover - best-effort user convenience
        print(f"Could not open output folder automatically: {exc}")


def open_path(path: Path) -> None:
    """Open a folder, or open Explorer with a specific file selected."""
    try:
        system = platform.system().lower()
        if system == "windows":
            if path.is_file():
                subprocess.Popen(["explorer.exe", "/select,", str(path)])
            else:
                os.startfile(str(path))  # type: ignore[attr-defined]
        elif system == "darwin":
            if path.is_file():
                subprocess.Popen(["open", "-R", str(path)])
            else:
                subprocess.Popen(["open", str(path)])
        else:
            target = path.parent if path.is_file() else path
            subprocess.Popen(["xdg-open", str(target)])
    except Exception as exc:  # pragma: no cover - best-effort user convenience
        print(f"Could not open path automatically: {exc}")


def newest_file(folder: Path, patterns: Iterable[str]) -> Path | None:
    candidates: list[Path] = []
    for pattern in patterns:
        candidates.extend(folder.glob(pattern))
    files = [path for path in candidates if path.is_file() and not path.name.startswith("~$")]
    if not files:
        return None
    return max(files, key=lambda p: p.stat().st_mtime)
