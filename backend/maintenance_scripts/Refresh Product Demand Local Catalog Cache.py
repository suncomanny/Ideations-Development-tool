from pathlib import Path
import sys


ROOT = Path(__file__).resolve().parent
sys.path.insert(0, str(ROOT / "product_demand_ideation" / "src"))

from product_demand_cli import refresh_local_sunco_catalog_cache


if __name__ == "__main__":
    refresh_local_sunco_catalog_cache(ROOT)
