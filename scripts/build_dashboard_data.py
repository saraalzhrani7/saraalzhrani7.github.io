"""Builds data/dashboard.json for dashboard.html from the GoatCounter/listings Excel export.
Usage:  python scripts/build_dashboard_data.py awal-khatwa-data-en.xlsx data/dashboard.json
Needs:  pip install pandas openpyxl
"""
import sys, json
import pandas as pd

src = sys.argv[1] if len(sys.argv) > 1 else "awal-khatwa-data-en.xlsx"
out = sys.argv[2] if len(sys.argv) > 2 else "data/dashboard.json"
x = pd.read_excel(src, sheet_name=None)
L, C, H, A, S = x["Listings"], x["ListingCities"], x["HistoryDaily"], x["Activity"], x["ScreenSizes"]

def d(v): return pd.Timestamp(v).strftime("%Y-%m-%d")
def top(df, col, n=None):
    s = df.groupby(col).size().sort_values(ascending=False)
    s = s.head(n) if n else s
    return [{"name": str(k), "value": int(v)} for k, v in s.items()]

op = L[L.is_open == True]
apply_clicks = int(A[A.category == "apply"]["count"].sum())
applied = int(A[A.category == "applied"]["count"].sum())
pv = A[A.category == "page"]
scr = S.groupby("device")["count"].sum()
mobile = float(scr.get("Mobile", 0) / scr.sum()) if scr.sum() else 0
src_pv = pv[pv.source != "Internal"].groupby("source")["count"].sum().sort_values(ascending=False)
cv = A[(A.category == "cv") & A.item.isin(["CV Upload", "Job Match", "PDF Import", "CV Download"])]
cv = cv.groupby("item")["count"].sum()
cities = C[C.is_open == True].groupby("city")["id"].nunique().sort_values(ascending=False)
daily = pv.groupby("date")["count"].sum()
newl = op[op.new_this_week == True]

data = {
    "updated": d(max(H["date"].max(), A["date"].max())),
    "since": d(H["date"].min()),
    "kpi": {
        "open": int(len(op)),
        "new_this_week": int(len(newl)),
        "apply_clicks": apply_clicks,
        "applied": applied,
        "apply_rate": round(applied / apply_clicks, 4) if apply_clicks else 0,
        "page_views": int(pv["count"].sum()),
        "mobile_share": round(mobile, 4),
        "from_linkedin": int(pv[pv.source == "LinkedIn"]["count"].sum()),
    },
    "top_companies": top(op, "company", 6),
    "history": [{"date": d(r.date), "open": int(r.open)} for r in H.itertuples()],
    "cities": [{"name": k, "value": int(v)} for k, v in cities.items()],
    "types": top(op, "type"),
    "fields": top(op, "field", 8),
    "new_listings": [
        {"company": r.company, "title": r.title, "type": r.type, "city": r.city}
        for r in newl.itertuples()
    ],
    "sources": [{"name": k, "value": int(v)} for k, v in src_pv.items()],
    "cv_tools": [{"name": k, "value": int(cv.get(k, 0))} for k in ["CV Upload", "Job Match", "PDF Import", "CV Download"]],
    "daily_views": [{"date": d(k), "value": int(v)} for k, v in daily.items()],
}
with open(out, "w", encoding="utf-8") as f:
    json.dump(data, f, ensure_ascii=False, indent=1)
print("wrote", out, "| open:", data["kpi"]["open"], "| updated:", data["updated"])
