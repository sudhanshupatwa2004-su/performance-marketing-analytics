"""Run every query in the SQL file against SQLite and compare headline KPIs with Python."""
import sqlite3, re, json, pandas as pd
sql = open("sql/wishcare_marketing_queries.sql").read()
sql = re.sub(r"/\*.*?\*/", "", sql, flags=re.S)
sql = "\n".join(l for l in sql.splitlines() if not l.strip().startswith("--"))
stmts = [s.strip() for s in sql.split(";") if s.strip()]
con = sqlite3.connect(":memory:")
results = []
for s in stmts:
    if s.upper().startswith(("CREATE DATABASE", "USE ")): continue
    if s.upper().startswith("CREATE TABLE") or s.upper().startswith("DROP"):
        con.execute(s); continue
    results.append((s.splitlines()[0][:60], pd.read_sql(s, con)))
    if len(results) == 0: pass
# load data after tables exist (re-run selects)
pd.read_csv("data/cleaned/campaign_performance_cleaned.csv").to_sql("campaign_performance", con, if_exists="append", index=False)
pd.read_csv("data/cleaned/ab_test_cleaned.csv").to_sql("ab_test", con, if_exists="append", index=False)
out = []
for s in stmts:
    if s.upper().startswith(("CREATE", "USE ", "DROP")): continue
    out.append(pd.read_sql(s, con))
print(f"{len(out)} SELECT queries ran OK")
for i in (0, 1, 2, 4): print(out[i].to_string(index=False)[:600], "\n")
py = json.load(open("data/cleaned/results.json"))["overall"]
q = out[4].iloc[0]
for k_sql, k_py in [("roas","ROAS"),("cac","CAC"),("ctr_pct","CTR_%"),("cpc","CPC"),("cpm","CPM"),("cvr_pct","CVR_%"),("aov","AOV")]:
    assert abs(q[k_sql] - round(py[k_py], 2)) < 0.006, (k_sql, q[k_sql], py[k_py])
print("SQL headline KPIs match Python exactly")
print(out[-2].to_string(index=False))
