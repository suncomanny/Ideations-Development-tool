from pathlib import Path
import sys

ROOT = Path(__file__).resolve().parent
sys.path.insert(0, str(ROOT / "product_demand_ideation" / "src"))

from product_demand_cli import run_product_demand_step1b


if __name__ == "__main__":
    run_product_demand_step1b(ROOT)
