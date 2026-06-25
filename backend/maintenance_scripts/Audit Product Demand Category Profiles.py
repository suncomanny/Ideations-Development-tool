from pathlib import Path
import sys


ROOT = Path(__file__).resolve().parent
sys.path.insert(0, str(ROOT / "product_demand_ideation" / "src"))

from profile_audit import audit_category_profiles


if __name__ == "__main__":
    output = audit_category_profiles(ROOT)
    print("Product Demand category profile audit complete:")
    print(f"  {output}")
