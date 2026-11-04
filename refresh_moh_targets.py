
import os
import json
import glob
from datetime import datetime
from pathlib import Path

import pandas as pd
import merch_kpi_export

# Existing reports folder
BASE = Path(
    r"G:\.shortcut-targets-by-id\1-OzYqibxMbDT38omFZxDLknsTKiBDbpt"
    r"\Business Analytics\Dashboard\Supporting_Folder_For_Complete_dashboard"
    r"\Merchandising_KPI_Dashboard"
)

# Dashboard JSON
JSON_FILE = Path(__file__).resolve().parent / "merch_kpi_data.json"

# Find the latest existing Excel report in each folder
REPORTS = {
    "a_class": ("a_class", "a_class_*.xlsx"),
    "b_class": ("b_class", "b_class_*.xlsx"),
    "new_sku_availability": ("new_sku_availability", "new_sku_availability_*.xlsx"),
    "moh": ("moh", "moh_*.xlsx"),
    "inventory_value_mrp": ("inventory_value", "inventory_value_mrp_*.xlsx"),
    "inventory_value_cp": ("inventory_value", "inventory_value_cp_*.xlsx"),
}

collected = {}

for prefix, (folder, pattern) in REPORTS.items():
    files = glob.glob(str(BASE / folder / pattern))

    if not files:
        raise FileNotFoundError(f"No existing report found: {pattern}")

    latest = max(files, key=os.path.getmtime)

    df = pd.read_excel(latest, sheet_name="Data")

    collected[prefix + ".xlsx"] = (df, latest)

    print(f"Loaded: {latest}")

# Preserve the original reporting date labels
with open(JSON_FILE, "r", encoding="utf-8") as f:
    old_payload = json.load(f)

generated = old_payload["metadata"]["generated_at"]
snapshot_date = datetime.fromisoformat(generated).date()

# Rebuild dashboard JSON using the updated MOH targets
warnings = merch_kpi_export.export(
    collected,
    str(JSON_FILE),
    today=snapshot_date
)

for warning in warnings:
    print("WARNING:", warning)

print("\nMOH targets refreshed successfully.")
print("Dashboard JSON:", JSON_FILE)
