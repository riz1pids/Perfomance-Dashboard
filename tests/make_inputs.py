"""Synthetic but structurally faithful inputs for build_dashboard_data.py:
same file names, sheet names, header rows and column names it reads. Writes a
nested folder (as a Drive download does) and returns the expected totals."""
import json, os, sys
from datetime import date, timedelta
import numpy as np, pandas as pd

OUT = sys.argv[1]
SUB = os.path.join(OUT, "perfomance")          # gdown puts files in a subfolder
os.makedirs(SUB, exist_ok=True)
rng = np.random.default_rng(7)

DAYS_TY = [date(2026, 9, d) for d in range(1, 31)]
DAYS_LY = [date(2025, 9, d) for d in range(1, 31)]     # weekday pairing irrelevant here
LOADED = 12                                            # TY has 12 days; LY has the full month

cats = []
for div, c1s in {"COMPANY GOODS": ["PACKAGED FOOD", "BABY CARE"],
                 "FRESH": ["FISH", "VEGETABLE"]}.items():
    for c1 in c1s:
        for c3 in (["FROZEN FOOD", "NOODLES", "SNACKS"] if c1 == "PACKAGED FOOD" else
                   [c1 + " A", c1 + " B"]):
            cats.append((div, c1, c3))
brands = ["SHWAPNO", "BENGAL MEAT", "PRAN", "NESTLE", "ACI"]
arts = []
code = 1000001
for (div, c1, c3) in cats:
    for b in brands[: rng.integers(2, 5)]:
        for _ in range(3):
            arts.append((code, f"{b} {c3} item {code}", div, c1, c3, b)); code += 1
outlets = [f"D{n:03d}" for n in range(1, 41)] + ["N001", "N002"]   # N = new this year
same = [o for o in outlets if o.startswith("D")][:32]

def sales_rows(days, outs, year_ly):
    rows = []
    for d in days:
        for o in outs:
            pick = rng.choice(len(arts), size=25, replace=False)
            ff = int(rng.integers(150, 600))
            for k in pick:
                a = arts[k]
                q = float(rng.integers(1, 9))
                nsi = round(q * rng.uniform(40, 900), 2)
                rows.append((d.strftime("%m/%d/%Y"), o, a[0], nsi, round(nsi * rng.uniform(.08, .28), 2), ff, q))
    return pd.DataFrame(rows, columns=["Date", "Outlet Code", "Product Code", "POS NSI", "POS GPV", "Footfall", "POS Sales"])

ty = sales_rows(DAYS_TY[:LOADED], outlets, False)
ly = sales_rows(DAYS_LY, [o for o in outlets if o.startswith("D")], True)
ty.to_csv(os.path.join(SUB, "sept 2026.csv"), index=False)
ly.to_csv(os.path.join(SUB, "SEPT 2025.CSV"), index=False)            # case differs on purpose

# same-store list: header on Excel row 2
ss = pd.DataFrame({"SL": range(1, len(same) + 1), "CODE": same})
with pd.ExcelWriter(os.path.join(SUB, "SEPT'26 Same Store List.xlsx")) as xw:
    pd.DataFrame([["Same store list Sep'26"]]).to_excel(xw, index=False, header=False, startrow=0)
    ss.to_excel(xw, index=False, startrow=1)

# same-day map: row 3 TY dates, row 9 LY dates, data from column 1
grid = [[None] * 31 for _ in range(12)]
grid[3][0] = "TY"; grid[9][0] = "SPLY"
for i, (t, l) in enumerate(zip(DAYS_TY, DAYS_LY), start=1):
    grid[3][i] = pd.Timestamp(t); grid[9][i] = pd.Timestamp(l)
with pd.ExcelWriter(os.path.join(SUB, "SAME DAY Sep'26 Proposed.xlsx")) as xw:
    pd.DataFrame(grid).to_excel(xw, sheet_name="Sep'26", index=False, header=False)

# person map
crp_of = {c: ("CRP " + c[2][:6], "HEAD " + c[0][:5]) for c in cats}
pm = pd.DataFrame([{"Division[Article]": c[0], "Category Level 1[Article]": c[1],
                    "Category Level 3": c[2], "CRP": crp_of[c][0], "Head": crp_of[c][1]} for c in cats])
with pd.ExcelWriter(os.path.join(SUB, "Personwise Category Distribution.xlsx")) as xw:
    pm.to_excel(xw, sheet_name="raw", index=False)

# target: header on Excel row 2; sales block then GPV block, each day-dated
tcols = ["DIVISION", "Category Level 1[Article]", "Category Level 3", "CRP"]
trows = []
for c in cats:
    s_vals = [round(rng.uniform(2e5, 9e5), 2) for _ in DAYS_TY]
    g_vals = [round(v * .18, 2) for v in s_vals]
    trows.append([c[0], c[1], c[2], crp_of[c][0]] + s_vals + [None] + g_vals + [None])
header = tcols + [pd.Timestamp(d) for d in DAYS_TY] + ["Total"] + [pd.Timestamp(d) for d in DAYS_TY] + ["Total "]
with pd.ExcelWriter(os.path.join(SUB, "SEPT'26 Cat3 wise Target.xlsx")) as xw:
    pd.DataFrame([["Cat3 target"]]).to_excel(xw, index=False, header=False, startrow=0)
    pd.DataFrame(trows, columns=header).to_excel(xw, index=False, startrow=1)

ka = pd.DataFrame(arts, columns=["Article_Code", "Article_Name", "Division", "Category_Level_1",
                                 "Category_Level_3", "Brand_Name_Description"])
with pd.ExcelWriter(os.path.join(SUB, "Key_Article.xlsx")) as xw:
    ka.to_excel(xw, sheet_name="Key_Article", index=False)

fmts = ["Super Store", "Express", "Mega"]
ks = pd.DataFrame([{"Outlet_Code": o, "Outlet_Name": f"{o} Outlet", "Outlet_Format": fmts[i % 3],
                    "Outlet_Division": ["Dhaka", "Chattogram", "Khulna"][i % 3],
                    "Outlet_Zone": f"Zone {i % 5}", "Outlet_Cluster": ["Own", "Franchise"][i % 2],
                    "Outlet_Status": "ACTIVE", "Outlet_Size": 1000} for i, o in enumerate(outlets)])
ks.to_excel(os.path.join(SUB, "key_store.xlsx"), index=False)
os.rename(os.path.join(SUB, "key_store.xlsx"), os.path.join(SUB, "key_store.XLSX"))   # case differs on purpose

zone = pd.DataFrame([{"CODE": o, "Regional Head HR Name": f"Leader {i % 3}",
                      "Zonal HR Name": f"Zonal {i % 6}", "Area": f"Area {i % 4}"} for i, o in enumerate(outlets)])
with pd.ExcelWriter(os.path.join(SUB, "Zone Distribution Sep 2026 w location type.xlsx")) as xw:
    zone.to_excel(xw, sheet_name="Final_Zone Dis", index=False)

ly12 = ly[pd.to_datetime(ly["Date"]).dt.day <= LOADED]
exp = {"all": round(ty["POS NSI"].sum() / 1e7, 2),
       "same": round(ty[ty["Outlet Code"].isin(same)]["POS NSI"].sum() / 1e7, 2),
       "frozen": round(ty[ty["Product Code"].isin([a[0] for a in arts if a[4] == "FROZEN FOOD"])]["POS NSI"].sum() / 1e7, 2),
       "ly_all_loaded": round(ly12["POS NSI"].sum() / 1e7, 2)}
print(json.dumps(exp))
