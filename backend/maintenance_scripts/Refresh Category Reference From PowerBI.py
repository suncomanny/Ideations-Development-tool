from pathlib import Path
import sys


ROOT = Path(__file__).resolve().parents[2]
sys.path.insert(0, str(ROOT / "backend" / "app"))

from opportunity_engine.refresh_category_reference import main


if __name__ == "__main__":
    main()
