from pathlib import Path
import sys


ROOT = Path(__file__).resolve().parents[2]
sys.path.insert(0, str(ROOT / "product_demand_ideation" / "src"))

from sku_classification_cache import default_classification_workbook_path, refresh_classification_cache_from_workbook


if __name__ == "__main__":
    workbook = Path(sys.argv[1]) if len(sys.argv) > 1 else default_classification_workbook_path(ROOT)
    output = refresh_classification_cache_from_workbook(ROOT, workbook)
    print("Product Demand SKU classification cache refresh complete:")
    print(f"  Source workbook: {workbook}")
    print(f"  Cache: {output}")
