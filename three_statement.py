# %% [markdown]
# # Three-Statement Financial Model
#
# **Income statement, balance sheet and cash flow, fully linked, in three scenarios.**
#
# The model every banking analyst builds first: five years of projections where
# the three statements tie out every year, a revolver funds any cash shortfall,
# and one switch moves the whole model between Base, Upside and Downside cases.
#
# | Module | Output |
# |---|---|
# | 1. Historicals | Three years of income statement and balance sheet (Yahoo Finance, or a synthetic demo company) |
# | 2. Drivers | Growth, margins, working-capital days, capex, D&A, financing, one set per scenario |
# | 3. Income statement | Revenue to net income, interest on opening balances (no circularity) |
# | 4. Balance sheet | Working capital from days, PP&E roll-forward, debt schedule, retained earnings |
# | 5. Cash flow | CFO, CFI, CFF, with a revolver that draws below minimum cash and repays when cash allows |
# | 6. Checks | Balance sheet balances and cash ties to the cash flow statement, every year |
# | 7. Credit & returns | Free cash flow, net debt, leverage, interest coverage by scenario |
# | 8. Exports | **Excel model with live formulas and a scenario switch** + charts |
#
# > Educational project. Historical figures from Yahoo Finance are reclassified into a standard template;
# > reconcile them with the latest 10-K before relying on the output.

# %%
import subprocess, sys
subprocess.run([sys.executable, "-m", "pip", "install", "-q", "yfinance", "xlsxwriter"], check=False)

# %%
import os, warnings, datetime as dt
import numpy as np
import pandas as pd
import matplotlib.pyplot as plt

warnings.filterwarnings("ignore")
pd.set_option("display.float_format", lambda x: f"{x:,.1f}")

# %% [markdown]
# ## 0. Configuration

# %%
TICKER = "GIS"
PROJECT_NAME = "Three-Statement Model"
AUTHOR = "Alessandro Radice"
DATA_MODE = os.environ.get("TSM_MODE", "live")   # "live" (Yahoo Finance) or "demo" (synthetic, offline)
SCENARIOS = ["Base", "Upside", "Downside"]
N = 5   # projection years

# Financing and policy assumptions (same in every scenario), $mm unless stated
POLICY = dict(
    tax_rate=0.24, rate_term=0.060, rate_revolver=0.070, rate_cash=0.030,
    min_cash=250.0, amortisation=150.0, dividends=150.0,
    da_pct=0.036, capex_pct=0.045,
)

OUTPUT_DIR = "tsm_output"
os.makedirs(OUTPUT_DIR, exist_ok=True)
TODAY = dt.date.today()

# %% [markdown]
# ## 1. Historicals
# A standard template: the balance sheet is reclassified so that it balances exactly in every historical year.

# %%
IS_ITEMS = ["Revenue", "COGS", "SG&A", "D&A", "Interest expense", "Interest income", "Tax"]
BS_ITEMS = ["Cash", "Accounts receivable", "Inventory", "Other current assets", "PP&E, net", "Goodwill & intangibles",
            "Accounts payable", "Other current liabilities", "Revolver", "Term debt", "Other long-term liabilities", "Equity"]

def demo_company():
    years = [2023, 2024, 2025]                      # fixed fiscal years for the synthetic demo
    rev = np.array([4200.0, 4480.0, 4750.0])
    is_ = pd.DataFrame({"Revenue": rev, "COGS": rev * [0.585, 0.580, 0.578], "SG&A": rev * [0.240, 0.238, 0.236],
                        "D&A": rev * 0.035, "Interest expense": [126.0, 123.0, 118.0], "Interest income": [6.0, 7.0, 8.0]},
                       index=years)
    ebt = is_["Revenue"] - is_["COGS"] - is_["SG&A"] - is_["D&A"] - is_["Interest expense"] + is_["Interest income"]
    is_["Tax"] = ebt * 0.24
    bs = pd.DataFrame({"Cash": [260.0, 280.0, 300.0], "Accounts receivable": [540.0, 585.0, 620.0],
                       "Inventory": [505.0, 535.0, 560.0], "Other current assets": [110.0, 115.0, 120.0],
                       "PP&E, net": [1380.0, 1415.0, 1450.0], "Goodwill & intangibles": [900.0, 900.0, 900.0],
                       "Accounts payable": [345.0, 362.0, 380.0], "Other current liabilities": [270.0, 280.0, 290.0],
                       "Revolver": [0.0, 0.0, 0.0], "Term debt": [2200.0, 2050.0, 1900.0],
                       "Other long-term liabilities": [250.0, 255.0, 260.0]}, index=years)
    assets = bs[BS_ITEMS[:6]].sum(axis=1)
    bs["Equity"] = assets - bs[BS_ITEMS[6:11]].sum(axis=1)
    return "Demo Co. (synthetic)", is_, bs

def live_company(tk):
    import yfinance as yf
    t = yf.Ticker(tk)
    inc, bal = t.income_stmt, t.balance_sheet
    def g(df, names, default=0.0):
        for n in names:
            if n in df.index:
                return pd.to_numeric(df.loc[n], errors="coerce").fillna(0.0) / 1e6
        return pd.Series(default, index=df.columns)
    cols = [c for c in inc.columns if c in bal.columns][:3][::-1]
    inc, bal = inc[cols], bal[cols]
    is_ = pd.DataFrame({
        "Revenue": g(inc, ["Total Revenue", "Operating Revenue"]), "COGS": g(inc, ["Cost Of Revenue"]),
        "SG&A": g(inc, ["Selling General And Administration", "Operating Expense"]),
        "D&A": g(inc, ["Reconciled Depreciation", "Depreciation And Amortization In Income Statement"]),
        "Interest expense": g(inc, ["Interest Expense", "Interest Expense Non Operating"]).abs(),
        "Interest income": g(inc, ["Interest Income", "Interest Income Non Operating"]).abs(),
        "Tax": g(inc, ["Tax Provision"])})
    cash = g(bal, ["Cash And Cash Equivalents"]); ar = g(bal, ["Accounts Receivable", "Receivables"])
    inv = g(bal, ["Inventory"]); tca = g(bal, ["Current Assets"]); ppe = g(bal, ["Net PPE"])
    gwi = g(bal, ["Goodwill And Other Intangible Assets"]); ta = g(bal, ["Total Assets"])
    ap = g(bal, ["Accounts Payable", "Payables"]); tcl = g(bal, ["Current Liabilities"])
    debt = g(bal, ["Total Debt"]); cdebt = g(bal, ["Current Debt", "Current Debt And Capital Lease Obligation"])
    eq = g(bal, ["Total Equity Gross Minority Interest", "Stockholders Equity"])
    bs = pd.DataFrame({"Cash": cash, "Accounts receivable": ar, "Inventory": inv,
                       "Other current assets": tca - cash - ar - inv,
                       "PP&E, net": ppe, "Goodwill & intangibles": gwi + (ta - tca - ppe - gwi),   # other long-term assets folded in
                       "Accounts payable": ap, "Other current liabilities": tcl - ap - cdebt,
                       "Revolver": 0.0, "Term debt": debt, "Other long-term liabilities": ta - eq - tcl - (debt - cdebt),
                       "Equity": eq})
    is_.index = bs.index = [pd.Timestamp(c).year for c in cols]
    if len(is_) < 3 or is_["Revenue"].iloc[-1] <= 0:
        raise ValueError("incomplete statements")
    cf = t.cashflow
    extra = {}
    if cf is not None and len(cf):
        c0 = cf.columns[0]
        pick = lambda names: next((abs(float(cf.loc[n, c0])) / 1e6 for n in names if n in cf.index and pd.notna(cf.loc[n, c0])), None)
        extra = {"capex": pick(["Capital Expenditure"]), "dividends": pick(["Cash Dividends Paid", "Common Stock Dividend Paid"])}
    extra["current_debt"] = float(cdebt.iloc[-1]) if len(cdebt) else None
    return t.info.get("longName", tk), is_, bs, extra

IS_DEMO = DATA_MODE != "live"
if not IS_DEMO:
    try:
        NAME, HIS, HBS, EXTRA = live_company(TICKER)
        # Policy defaults calibrated to the company's latest year (edit POLICY to override)
        rev_l = HIS["Revenue"].iloc[-1]
        POLICY["da_pct"] = round(HIS["D&A"].iloc[-1] / rev_l, 4)
        if EXTRA.get("capex"): POLICY["capex_pct"] = round(EXTRA["capex"] / rev_l, 4)
        if EXTRA.get("dividends"): POLICY["dividends"] = round(EXTRA["dividends"], 1)
        POLICY["amortisation"] = 0.0   # maturities assumed refinanced; set a figure to model mandatory paydown
        POLICY["min_cash"] = round(float(0.5 * HBS["Cash"].iloc[-1]), 1)
        POLICY.update({k: float(v) for k, v in POLICY.items()})
    except Exception as e:
        print(f"Live data unavailable ({e}). Falling back to synthetic DEMO data.")
        IS_DEMO = True
if IS_DEMO:
    NAME, HIS, HBS = demo_company()
DATA_LABEL = "SYNTHETIC DEMO DATA" if IS_DEMO else f"Yahoo Finance, reclassified; as of {TODAY:%d %b %Y}"
YEARS_H = list(HIS.index)
YEARS_P = [YEARS_H[-1] + i for i in range(1, N + 1)]
assert (HBS[BS_ITEMS[:6]].sum(axis=1) - HBS[BS_ITEMS[6:]].sum(axis=1)).abs().max() < 1e-6, "historical BS must balance"
print(NAME, "| historical years:", YEARS_H)
HIS.T

# %% [markdown]
# ## 2. Drivers by scenario
# Base-case operating drivers start from the last historical year (cost ratios trimmed slightly); Upside and Downside shift growth, margins and
# working-capital days. Edit any row.

# %%
last = HIS.iloc[-1]; lastbs = HBS.iloc[-1]
cogs0, sga0 = last["COGS"] / last["Revenue"], last["SG&A"] / last["Revenue"]
dso0 = lastbs["Accounts receivable"] / last["Revenue"] * 365
dio0 = lastbs["Inventory"] / last["COGS"] * 365
dpo0 = lastbs["Accounts payable"] / last["COGS"] * 365
DRIVERS = {
    "Base":     dict(growth=[0.050, 0.050, 0.045, 0.040, 0.035], cogs=[cogs0 - 0.003] * N, sga=[sga0 - 0.001] * N,
                     dso=[round(dso0)] * N, dio=[round(dio0)] * N, dpo=[round(dpo0)] * N),
    "Upside":   dict(growth=[0.080, 0.075, 0.065, 0.055, 0.045], cogs=[cogs0 - 0.008] * N, sga=[sga0 - 0.006] * N,
                     dso=[round(dso0) - 2] * N, dio=[round(dio0) - 4] * N, dpo=[round(dpo0) + 2] * N),
    "Downside": dict(growth=[-0.040, 0.000, 0.020, 0.025, 0.030], cogs=[cogs0 + 0.017] * N, sga=[sga0 + 0.014] * N,
                     dso=[round(dso0) + 5] * N, dio=[round(dio0) + 6] * N, dpo=[round(dpo0) - 4] * N),
}
pd.DataFrame({s: {k: v[0] for k, v in d.items()} for s, d in DRIVERS.items()})

# %% [markdown]
# ## 3–6. The linked model
# Interest is charged on opening balances, so the model has no circular reference.
# The revolver draws when cash would fall below the minimum and repays as soon as cash allows.

# %%
def run_model(scn, p=POLICY):
    d = DRIVERS[scn]
    IS = {k: [] for k in ["Revenue", "COGS", "Gross profit", "SG&A", "EBITDA", "D&A", "EBIT", "Interest expense",
                          "Interest income", "EBT", "Tax", "Net income"]}
    BS = {k: [] for k in BS_ITEMS}
    CF = {k: [] for k in ["Net income", "D&A", "(Increase) in receivables", "(Increase) in inventory",
                          "Increase in payables", "Cash from operations", "Capex", "Cash from investing",
                          "Term debt repayment", "Dividends", "Revolver draw / (repayment)", "Cash from financing",
                          "Net change in cash", "Beginning cash", "Ending cash"]}
    prev = HBS.iloc[-1].to_dict(); rev_prev = HIS["Revenue"].iloc[-1]
    for i in range(N):
        rev = rev_prev * (1 + d["growth"][i]); cogs = rev * d["cogs"][i]; sga = rev * d["sga"][i]
        ebitda = rev - cogs - sga; da = rev * p["da_pct"]; ebit = ebitda - da
        int_exp = prev["Term debt"] * p["rate_term"] + prev["Revolver"] * p["rate_revolver"]
        int_inc = prev["Cash"] * p["rate_cash"]
        ebt = ebit - int_exp + int_inc; tax = max(0.0, ebt) * p["tax_rate"]; ni = ebt - tax
        ar = rev * d["dso"][i] / 365; inv = cogs * d["dio"][i] / 365; ap = cogs * d["dpo"][i] / 365
        capex = rev * p["capex_pct"]
        term = max(0.0, prev["Term debt"] - p["amortisation"]); repay = prev["Term debt"] - term
        cfo = ni + da - (ar - prev["Accounts receivable"]) - (inv - prev["Inventory"]) + (ap - prev["Accounts payable"])
        pre = prev["Cash"] + cfo - capex - repay - p["dividends"]
        rev_chg = max(-prev["Revolver"], p["min_cash"] - pre)
        cash = pre + rev_chg; revolver = prev["Revolver"] + rev_chg
        row = dict(zip(IS.keys(), [rev, cogs, rev - cogs, sga, ebitda, da, ebit, int_exp, int_inc, ebt, tax, ni]))
        for k, v in row.items(): IS[k].append(v)
        cur = {"Cash": cash, "Accounts receivable": ar, "Inventory": inv, "Other current assets": prev["Other current assets"],
               "PP&E, net": prev["PP&E, net"] + capex - da, "Goodwill & intangibles": prev["Goodwill & intangibles"],
               "Accounts payable": ap, "Other current liabilities": prev["Other current liabilities"], "Revolver": revolver,
               "Term debt": term, "Other long-term liabilities": prev["Other long-term liabilities"],
               "Equity": prev["Equity"] + ni - p["dividends"]}
        for k, v in cur.items(): BS[k].append(v)
        cfrow = [ni, da, -(ar - prev["Accounts receivable"]), -(inv - prev["Inventory"]), ap - prev["Accounts payable"], cfo,
                 -capex, -capex, -repay, -p["dividends"], rev_chg, -repay - p["dividends"] + rev_chg,
                 cash - prev["Cash"], prev["Cash"], cash]
        for k, v in zip(CF.keys(), cfrow): CF[k].append(v)
        prev, rev_prev = cur, rev
    IS, BS, CF = (pd.DataFrame(x, index=YEARS_P).T for x in (IS, BS, CF))
    tot_a = BS.loc[BS_ITEMS[:6]].sum(); tot_le = BS.loc[BS_ITEMS[6:]].sum()
    checks = pd.DataFrame({"Balance sheet (A − L − E)": tot_a - tot_le,
                           "Cash (BS − CF)": BS.loc["Cash"] - CF.loc["Ending cash"]}).T
    net_debt = BS.loc["Term debt"] + BS.loc["Revolver"] - BS.loc["Cash"]
    kpi = pd.DataFrame({"Revenue growth": IS.loc["Revenue"].pct_change().fillna(IS.loc["Revenue"].iloc[0] / HIS["Revenue"].iloc[-1] - 1),
                        "EBITDA margin": IS.loc["EBITDA"] / IS.loc["Revenue"],
                        "Free cash flow": CF.loc["Cash from operations"] + CF.loc["Capex"],
                        "Net debt": net_debt, "Net debt / EBITDA": net_debt / IS.loc["EBITDA"],
                        "EBITDA / interest": IS.loc["EBITDA"] / IS.loc["Interest expense"],
                        "Revolver balance": BS.loc["Revolver"]}).T
    return dict(IS=IS, BS=BS, CF=CF, checks=checks, kpi=kpi)

RESULTS = {s: run_model(s) for s in SCENARIOS}
for s, r in RESULTS.items():
    print(f"{s:<9} max balance-sheet error: {r['checks'].abs().values.max():.2e}")
RESULTS["Base"]["IS"]

# %% [markdown]
# ## 7. Scenario comparison

# %%
comp = pd.DataFrame({s: {"Revenue, final year": r["IS"].loc["Revenue"].iloc[-1],
                         "EBITDA margin, final year": r["kpi"].loc["EBITDA margin"].iloc[-1],
                         "Cumulative free cash flow": r["kpi"].loc["Free cash flow"].sum(),
                         "Net debt / EBITDA, final year": r["kpi"].loc["Net debt / EBITDA"].iloc[-1],
                         "Peak revolver draw": r["BS"].loc["Revolver"].max(),
                         "Min EBITDA / interest": r["kpi"].loc["EBITDA / interest"].min()} for s, r in RESULTS.items()})
lev0 = (HBS["Term debt"].iloc[-1] + HBS["Revolver"].iloc[-1] - HBS["Cash"].iloc[-1]) / (last["Revenue"] - last["COGS"] - last["SG&A"])
print(f"Opening net debt / EBITDA: {lev0:.2f}x")
comp

# %% [markdown]
# ## 8. Charts

# %%
INK, GREY, LIGHT, MUTED = "#1A1A18", "#8A8980", "#C8C7BF", "#66655E"
plt.rcParams.update({"font.family": "DejaVu Sans", "axes.spines.top": False, "axes.spines.right": False,
                     "axes.edgecolor": "#b5b3ad", "xtick.color": MUTED, "ytick.color": MUTED})
fig, ax = plt.subplots(figsize=(9, 4.5))
styles = {"Base": (INK, "-"), "Upside": (GREY, "--"), "Downside": (LIGHT, "-")}
for s, r in RESULTS.items():
    y = [lev0] + list(r["kpi"].loc["Net debt / EBITDA"])
    c, ls = styles[s]
    ax.plot([YEARS_H[-1]] + YEARS_P, y, color=c, ls=ls, lw=2.5, marker="o", label=s)
    ax.annotate(f"{y[-1]:.1f}x", (YEARS_P[-1], y[-1]), xytext=(8, 0), textcoords="offset points", va="center", fontsize=9, color=INK)
ax.set_title("Net debt / EBITDA by scenario", loc="left", fontsize=12, color=INK)
ax.set_xticks([YEARS_H[-1]] + YEARS_P, [f"FY{str(y)[2:]}{'A' if y == YEARS_H[-1] else 'E'}" for y in [YEARS_H[-1]] + YEARS_P])
ax.yaxis.set_major_formatter(lambda v, _: f"{v:.1f}x"); ax.legend(frameon=False, fontsize=9); ax.margins(x=0.08)
fig.text(0.01, 0.01, f"Source: {DATA_LABEL}", fontsize=7, color=MUTED)
fig.tight_layout(rect=(0, 0.03, 1, 1))
fig.savefig(f"{OUTPUT_DIR}/leverage.png", dpi=200, metadata={"Title": "Net debt / EBITDA by scenario", "Author": AUTHOR, "Software": None})
plt.show()

# %% [markdown]
# ## 9. Excel export: live formulas and a scenario switch
# Change the scenario cell on `Inputs` (1 = Base, 2 = Upside, 3 = Downside) and all three statements recalculate.

# %%
import xlsxwriter
from xlsxwriter.utility import xl_col_to_name

XLSX = f"{OUTPUT_DIR}/{('DEMO' if IS_DEMO else TICKER)}_three_statement_model.xlsx"
wb = xlsxwriter.Workbook(XLSX)
wb.set_properties({"title": f"{NAME}: three-statement model", "author": AUTHOR, "manager": AUTHOR,
                   "subject": "Linked income statement, balance sheet and cash flow with scenarios",
                   "keywords": "financial model, three statement, investment banking",
                   "comments": "Generated by the Three-Statement Financial Model", "created": dt.datetime.now()})
F = dict(font_name="Arial", font_size=10); fmt = lambda **k: wb.add_format({**F, **k})
NUM = '#,##0.0;(#,##0.0);"–"'
f_title = fmt(bold=True, font_size=14); f_sub = fmt(italic=True, font_color="#66655E")
f_hdr = fmt(bold=True, font_color="white", bg_color="#1A1A18", align="center")
f_hdrE = fmt(bold=True, font_color="white", bg_color="#5C5B55", align="center")
f_in = fmt(font_color="#0000FF", num_format=NUM); f_inp = fmt(font_color="#0000FF", num_format="0.0%")
f_ind = fmt(font_color="#0000FF", num_format="0"); f_n = fmt(num_format=NUM); f_p = fmt(num_format="0.0%")
f_lk = fmt(font_color="#008000", num_format=NUM); f_lkp = fmt(font_color="#008000", num_format="0.0%")
f_lkd = fmt(font_color="#008000", num_format="0")
f_tot = fmt(num_format=NUM, bold=True, top=1); f_x = fmt(num_format='0.0"x"'); f_b = fmt(bold=True)
f_chk = fmt(num_format='0.000;-0.000;"OK"', bold=True, font_color="#008000")
f_sel = fmt(font_color="#0000FF", bold=True, bg_color="#FFF2CC", border=1, align="center")
H0 = 2                      # first historical column (C)
P0 = H0 + len(YEARS_H)      # first projection column (F)
COLS = [xl_col_to_name(H0 + i) for i in range(len(YEARS_H) + N)]
PC = COLS[len(YEARS_H):]    # projection column letters
LH = COLS[len(YEARS_H) - 1] # last historical column

def header(ws, row, label=""):
    ws.write(row, 1, label, f_b)
    for i, y in enumerate(YEARS_H + YEARS_P):
        ws.write(row, H0 + i, f"FY{str(y)[2:]}{'A' if y in YEARS_H else 'E'}", f_hdr if y in YEARS_H else f_hdrE)

def sheet(name, title):
    ws = wb.add_worksheet(name); ws.hide_gridlines(2)
    ws.set_column("A:A", 2); ws.set_column("B:B", 34); ws.set_column(H0, H0 + len(COLS), 11)
    ws.write("B2", title, f_title); ws.write("B3", f"{NAME} · $mm · Source: {DATA_LABEL}", f_sub)
    return ws

# ---------- Inputs ----------
wi = sheet("Inputs", "Inputs and scenario drivers"); wi.set_column("B:B", 50)
wi.write("B5", "Active scenario (1 = Base, 2 = Upside, 3 = Downside)", f_b); wi.write("C5", 1, f_sel)
wi.write_formula("D5", '=CHOOSE($C$5,"Base","Upside","Downside")', f_b)
SEL = "Inputs!$C$5"
header(wi, 7, "Operating drivers")
REFD = {}
r = 8
for key, lab, f_ in [("growth", "Revenue growth", f_inp), ("cogs", "COGS % of revenue", f_inp), ("sga", "SG&A % of revenue", f_inp),
                     ("dso", "Days sales outstanding", f_ind), ("dio", "Days inventory (on COGS)", f_ind), ("dpo", "Days payables (on COGS)", f_ind)]:
    wi.write(r, 1, lab, f_b); r += 1
    rows = {}
    for s in SCENARIOS:
        wi.write(r, 1, f"   {s}")
        for i in range(N):
            wi.write(r, P0 + i, DRIVERS[s][key][i], f_)
        rows[s] = r + 1; r += 1
    wi.write(r, 1, "   ► Live case", f_b)
    for i, c in enumerate(PC):
        wi.write_formula(r, P0 + i, f"=CHOOSE({SEL},{c}{rows['Base']},{c}{rows['Upside']},{c}{rows['Downside']})",
                         fmt(bold=True, num_format="0.0%" if f_ is f_inp else "0", bg_color="#ECEBE5"))
    REFD[key] = r + 1; r += 2
wi.write(r, 1, "Financing and policy (all scenarios)", f_b); r += 1
REFP = {}
for key, lab, f_ in [("tax_rate", "Tax rate", f_inp), ("rate_term", "Interest rate, term debt", f_inp),
                     ("rate_revolver", "Interest rate, revolver", f_inp), ("rate_cash", "Interest income on cash", f_inp),
                     ("min_cash", "Minimum cash balance", f_in), ("amortisation", "Term debt amortisation per year", f_in),
                     ("dividends", "Dividends per year", f_in), ("da_pct", "D&A % of revenue", f_inp), ("capex_pct", "Capex % of revenue", f_inp)]:
    wi.write(r, 1, lab); wi.write(r, 2, POLICY[key], f_); REFP[key] = f"Inputs!$C${r + 1}"; r += 1
D = lambda key, c: f"Inputs!{c}${REFD[key]}"
Pp = lambda key: REFP[key]

# ---------- Income statement ----------
wIS = sheet("Income Statement", "Income statement")
header(wIS, 5)
ISL = ["Revenue", "COGS", "Gross profit", "SG&A", "EBITDA", "D&A", "EBIT", "Interest expense", "Interest income", "EBT", "Tax", "Net income"]
RI = {k: 6 + i for i, k in enumerate(ISL)}
ISR = lambda k: RI[k] + 1
for k in ISL:
    wIS.write(RI[k], 1, k, f_b if k in ("Revenue", "EBITDA", "Net income") else None)
wIS.write(RI["Net income"] + 2, 1, "EBITDA margin"); wIS.write(RI["Net income"] + 3, 1, "Revenue growth")
# ---------- Balance sheet (define rows first for cross-references) ----------
wBS = sheet("Balance Sheet", "Balance sheet")
header(wBS, 5)
BSA = BS_ITEMS[:6]; BSL = BS_ITEMS[6:11]
RB = {}
r = 6
for k in BSA: RB[k] = r; r += 1
RB["Total assets"] = r; r += 2
for k in BSL: RB[k] = r; r += 1
RB["Total liabilities"] = r; r += 1
RB["Equity"] = r; r += 1
RB["Total liabilities & equity"] = r; r += 2
RB["Check"] = r
BR = lambda k: RB[k] + 1
# ---------- Cash flow rows ----------
wCF = sheet("Cash Flow", "Cash flow statement")
header(wCF, 5, "")
CFL = ["Net income", "D&A", "(Increase) in receivables", "(Increase) in inventory", "Increase in payables", "Cash from operations",
       "Capex", "Cash from investing", "Term debt repayment", "Dividends", "Cash before revolver", "Revolver draw / (repayment)",
       "Cash from financing", "Net change in cash", "Beginning cash", "Ending cash", "Check: ending cash vs balance sheet"]
RC = {k: 6 + i for i, k in enumerate(CFL)}
CR = lambda k: RC[k] + 1

# Historical values (blue)
for i, y in enumerate(YEARS_H):
    c = H0 + i
    for k in IS_ITEMS:
        wIS.write(RI[k], c, float(HIS.loc[y, k]), f_in)
    C = COLS[i]
    wIS.write_formula(RI["Gross profit"], c, f"={C}{ISR('Revenue')}-{C}{ISR('COGS')}", f_n)
    wIS.write_formula(RI["EBITDA"], c, f"={C}{ISR('Gross profit')}-{C}{ISR('SG&A')}", f_tot)
    wIS.write_formula(RI["EBIT"], c, f"={C}{ISR('EBITDA')}-{C}{ISR('D&A')}", f_n)
    wIS.write_formula(RI["EBT"], c, f"={C}{ISR('EBIT')}-{C}{ISR('Interest expense')}+{C}{ISR('Interest income')}", f_n)
    wIS.write_formula(RI["Net income"], c, f"={C}{ISR('EBT')}-{C}{ISR('Tax')}", f_tot)
    for k in BS_ITEMS:
        wBS.write(RB[k], c, float(HBS.loc[y, k]), f_in)

# Projections (formulas)
for i in range(N):
    c = P0 + i; C = PC[i]; P = COLS[len(YEARS_H) + i - 1]
    isr = lambda k: f"{C}{ISR(k)}"; bsr = lambda k: f"'Balance Sheet'!{C}{BR(k)}"; bsp = lambda k: f"'Balance Sheet'!{P}{BR(k)}"
    # IS
    wIS.write_formula(RI["Revenue"], c, f"={P}{ISR('Revenue')}*(1+{D('growth', C)})", f_n)
    wIS.write_formula(RI["COGS"], c, f"={isr('Revenue')}*{D('cogs', C)}", f_n)
    wIS.write_formula(RI["Gross profit"], c, f"={isr('Revenue')}-{isr('COGS')}", f_n)
    wIS.write_formula(RI["SG&A"], c, f"={isr('Revenue')}*{D('sga', C)}", f_n)
    wIS.write_formula(RI["EBITDA"], c, f"={isr('Gross profit')}-{isr('SG&A')}", f_tot)
    wIS.write_formula(RI["D&A"], c, f"={isr('Revenue')}*{Pp('da_pct')}", f_n)
    wIS.write_formula(RI["EBIT"], c, f"={isr('EBITDA')}-{isr('D&A')}", f_n)
    wIS.write_formula(RI["Interest expense"], c, f"={bsp('Term debt')}*{Pp('rate_term')}+{bsp('Revolver')}*{Pp('rate_revolver')}", f_n)
    wIS.write_formula(RI["Interest income"], c, f"={bsp('Cash')}*{Pp('rate_cash')}", f_n)
    wIS.write_formula(RI["EBT"], c, f"={isr('EBIT')}-{isr('Interest expense')}+{isr('Interest income')}", f_n)
    wIS.write_formula(RI["Tax"], c, f"=MAX(0,{isr('EBT')})*{Pp('tax_rate')}", f_n)
    wIS.write_formula(RI["Net income"], c, f"={isr('EBT')}-{isr('Tax')}", f_tot)
    # CF
    ISx = lambda k: f"'Income Statement'!{C}{ISR(k)}"
    cf = lambda k: f"{C}{CR(k)}"
    wCF.write_formula(RC["Net income"], c, f"={ISx('Net income')}", f_lk)
    wCF.write_formula(RC["D&A"], c, f"={ISx('D&A')}", f_lk)
    wCF.write_formula(RC["(Increase) in receivables"], c, f"={bsp('Accounts receivable')}-{bsr('Accounts receivable')}", f_n)
    wCF.write_formula(RC["(Increase) in inventory"], c, f"={bsp('Inventory')}-{bsr('Inventory')}", f_n)
    wCF.write_formula(RC["Increase in payables"], c, f"={bsr('Accounts payable')}-{bsp('Accounts payable')}", f_n)
    wCF.write_formula(RC["Cash from operations"], c, f"=SUM({C}{CR('Net income')}:{C}{CR('Increase in payables')})", f_tot)
    wCF.write_formula(RC["Capex"], c, f"=-{ISx('Revenue')}*{Pp('capex_pct')}", f_n)
    wCF.write_formula(RC["Cash from investing"], c, f"={cf('Capex')}", f_tot)
    wCF.write_formula(RC["Term debt repayment"], c, f"=-({bsp('Term debt')}-{bsr('Term debt')})", f_n)
    wCF.write_formula(RC["Dividends"], c, f"=-{Pp('dividends')}", f_n)
    wCF.write_formula(RC["Cash before revolver"], c, f"={bsp('Cash')}+{cf('Cash from operations')}+{cf('Cash from investing')}+{cf('Term debt repayment')}+{cf('Dividends')}", f_n)
    wCF.write_formula(RC["Revolver draw / (repayment)"], c, f"=MAX(-{bsp('Revolver')},{Pp('min_cash')}-{cf('Cash before revolver')})", f_n)
    wCF.write_formula(RC["Cash from financing"], c, f"={cf('Term debt repayment')}+{cf('Dividends')}+{cf('Revolver draw / (repayment)')}", f_tot)
    wCF.write_formula(RC["Net change in cash"], c, f"={cf('Cash from operations')}+{cf('Cash from investing')}+{cf('Cash from financing')}", f_n)
    wCF.write_formula(RC["Beginning cash"], c, f"={bsp('Cash')}", f_lk)
    wCF.write_formula(RC["Ending cash"], c, f"={cf('Beginning cash')}+{cf('Net change in cash')}", f_tot)
    wCF.write_formula(RC["Check: ending cash vs balance sheet"], c, f"=ROUND({cf('Ending cash')}-{bsr('Cash')},6)", f_chk)
    # BS
    b = lambda k: f"{C}{BR(k)}"; bp = lambda k: f"{P}{BR(k)}"; cfx = lambda k: f"'Cash Flow'!{C}{CR(k)}"
    wBS.write_formula(RB["Cash"], c, f"={cfx('Cash before revolver')}+{cfx('Revolver draw / (repayment)')}", f_lk)
    wBS.write_formula(RB["Accounts receivable"], c, f"={ISx('Revenue')}*{D('dso', C)}/365", f_n)
    wBS.write_formula(RB["Inventory"], c, f"={ISx('COGS')}*{D('dio', C)}/365", f_n)
    wBS.write_formula(RB["Other current assets"], c, f"={bp('Other current assets')}", f_n)
    wBS.write_formula(RB["PP&E, net"], c, f"={bp('PP&E, net')}+{ISx('Revenue')}*{Pp('capex_pct')}-{ISx('D&A')}", f_n)
    wBS.write_formula(RB["Goodwill & intangibles"], c, f"={bp('Goodwill & intangibles')}", f_n)
    wBS.write_formula(RB["Accounts payable"], c, f"={ISx('COGS')}*{D('dpo', C)}/365", f_n)
    wBS.write_formula(RB["Other current liabilities"], c, f"={bp('Other current liabilities')}", f_n)
    wBS.write_formula(RB["Revolver"], c, f"={bp('Revolver')}+{cfx('Revolver draw / (repayment)')}", f_n)
    wBS.write_formula(RB["Term debt"], c, f"=MAX(0,{bp('Term debt')}-{Pp('amortisation')})", f_n)
    wBS.write_formula(RB["Other long-term liabilities"], c, f"={bp('Other long-term liabilities')}", f_n)
    wBS.write_formula(RB["Equity"], c, f"={bp('Equity')}+{ISx('Net income')}-{Pp('dividends')}", f_n)
# Totals, ratios and checks for all columns
for i, C in enumerate(COLS):
    c = H0 + i
    wBS.write_formula(RB["Total assets"], c, f"=SUM({C}{BR('Cash')}:{C}{BR('Goodwill & intangibles')})", f_tot)
    wBS.write_formula(RB["Total liabilities"], c, f"=SUM({C}{BR('Accounts payable')}:{C}{BR('Other long-term liabilities')})", f_tot)
    wBS.write_formula(RB["Total liabilities & equity"], c, f"={C}{BR('Total liabilities')}+{C}{BR('Equity')}", f_tot)
    wBS.write_formula(RB["Check"], c, f"=ROUND({C}{BR('Total assets')}-{C}{BR('Total liabilities & equity')},6)", f_chk)
    wIS.write_formula(RI["Net income"] + 2, c, f"={C}{ISR('EBITDA')}/{C}{ISR('Revenue')}", f_p)
    if i > 0:
        wIS.write_formula(RI["Net income"] + 3, c, f"={C}{ISR('Revenue')}/{COLS[i - 1]}{ISR('Revenue')}-1", f_p)
for k in ["Total assets", "Total liabilities", "Equity", "Total liabilities & equity"]:
    wBS.write(RB[k], 1, k, f_b)
for k in BSA + BSL: wBS.write(RB[k], 1, k)
wBS.write(RB["Check"], 1, "Check: assets − liabilities − equity", f_b)
for k in CFL:
    wCF.write(RC[k], 1, k, f_b if k.startswith("Cash from") or k == "Ending cash" else None)

# ---------- Summary ----------
wS = sheet("Summary", "Summary: credit and cash flow (active scenario)")
wS.write_formula("B4", '="Active scenario: "&Inputs!$D$5', f_b)
header(wS, 5)
SL = ["Revenue", "EBITDA", "EBITDA margin", "Free cash flow (CFO − capex)", "Net debt", "Net debt / EBITDA", "EBITDA / interest", "Revolver balance",
      "Balance sheet check", "Cash check"]
for j, lab in enumerate(SL):
    rr = 6 + j; wS.write(rr, 1, lab, f_b if "check" in lab else None)
    for i, C in enumerate(COLS):
        c = H0 + i; proj = i >= len(YEARS_H)
        ISx = lambda k: f"'Income Statement'!{C}{ISR(k)}"; BSx = lambda k: f"'Balance Sheet'!{C}{BR(k)}"; CFx = lambda k: f"'Cash Flow'!{C}{CR(k)}"
        nd = f"({BSx('Term debt')}+{BSx('Revolver')}-{BSx('Cash')})"
        forms = {"Revenue": (f"={ISx('Revenue')}", f_lk), "EBITDA": (f"={ISx('EBITDA')}", f_lk),
                 "EBITDA margin": (f"={ISx('EBITDA')}/{ISx('Revenue')}", f_p),
                 "Free cash flow (CFO − capex)": (f"={CFx('Cash from operations')}+{CFx('Capex')}", f_n) if proj else None,
                 "Net debt": (f"={nd}", f_n), "Net debt / EBITDA": (f"={nd}/{ISx('EBITDA')}", f_x),
                 "EBITDA / interest": (f"={ISx('EBITDA')}/{ISx('Interest expense')}", f_x), "Revolver balance": (f"={BSx('Revolver')}", f_n),
                 "Balance sheet check": (f"={BSx('Check')}", f_chk), "Cash check": (f"={CFx('Check: ending cash vs balance sheet')}", f_chk) if proj else None}
        fo = forms[lab]
        if fo: wS.write_formula(rr, c, fo[0], fo[1])
wb.close()
print("Saved", XLSX)

# %% [markdown]
# ## 10. Download

# %%
try:
    from google.colab import files
    for fpath in [XLSX, f"{OUTPUT_DIR}/leverage.png"]:
        files.download(fpath)
except ImportError:
    print("Outputs saved in:", os.path.abspath(OUTPUT_DIR))
