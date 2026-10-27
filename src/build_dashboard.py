"""
Builds dashboard/index.html - an interactive pacing & delivery dashboard
for the DOOH plan, with the plan and delivery data embedded.

Run after build_plan.py:  python src/build_dashboard.py
Open dashboard/index.html in a browser (or via GitHub Pages).
"""
from pathlib import Path
import json
import pandas as pd

ROOT = Path(__file__).resolve().parents[1]
plan = pd.read_csv(ROOT / "data" / "media_plan.csv")
deliv = pd.read_csv(ROOT / "data" / "delivery_daily.csv")
deliv["day"] = (pd.to_datetime(deliv.date) - pd.Timestamp("2026-03-02")).dt.days + 1

data = {
    "flightDays": 28,
    "budget": 75000,
    "plan": plan[["venue_id", "metro", "publisher", "venue_type", "planned_spend",
                  "planned_impressions"]].values.tolist(),
    "deliv": deliv[["day", "venue_id", "delivered_impressions", "delivered_spend"]].values.tolist(),
}
html = (ROOT / "src" / "dashboard_template.html").read_text().replace(
    "__DATA__", json.dumps(data, separators=(",", ":")))
out = ROOT / "dashboard"
out.mkdir(exist_ok=True)
(out / "index.html").write_text(html)
print(f"wrote {out / 'index.html'} ({len(html)/1024:.0f} KB)")
