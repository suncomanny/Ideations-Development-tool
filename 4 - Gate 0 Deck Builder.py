from pathlib import Path
import sys

ROOT = Path(__file__).resolve().parent
sys.path.insert(0, str(ROOT / "backend" / "app"))

from opportunity_engine.cli import run_gate0_deck


if __name__ == "__main__":
    run_gate0_deck(ROOT)
