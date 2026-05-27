from pathlib import Path
import sys

ROOT = Path(__file__).resolve().parent
sys.path.insert(0, str(ROOT / "backend" / "app"))

from opportunity_engine.refresh_line_review_snapshots import main


if __name__ == "__main__":
    main()
