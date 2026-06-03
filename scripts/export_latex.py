"""
export_latex.py
───────────────
Reads result files from thesis_analysis.py and writes .tex table files
into ../tables/ and figure PDFs into ../figures/.

Run from the scripts/ folder:
    python export_latex.py
"""

import pandas as pd
import numpy as np
import shutil
from pathlib import Path

HERE    = Path(__file__).parent
OUT_TEX = HERE / ".." / "tables"
OUT_FIG = HERE / ".." / "figures"
OUT_TEX.mkdir(exist_ok=True)
OUT_FIG.mkdir(exist_ok=True)

# ── Helpers ───────────────────────────────────────────────────────────────

def stars(pv):
    if pd.isna(pv): return ""
    return r"$^{***}$" if pv < 0.01 else r"$^{**}$" if pv < 0.05 else r"$^{*}$" if pv < 0.10 else ""

def f4(v):
    return ":" if pd.isna(v) else f"{v:.4f}"

def f3(v):
    return ":" if pd.isna(v) else f"{v:.3f}"

def f2(v):
    return ":" if pd.isna(v) else f"{v:.2f}"

def alpha_cell(row):
    a  = row["Alpha (% pm)"]
    pv = row["p-value"]
    if pd.isna(a): return ":"
    return r"\textbf{" + f"{a:.4f}" + r"}" + stars(pv)

def tstat_cell(row):
    t = row["t-stat (NW)"]
    return ":" if pd.isna(t) else f"({t:.3f})"

def write(name, lines):
    path = OUT_TEX / f"{name}.tex"
    path.write_text("\n".join(lines), encoding="utf-8")
    print(f"  Saved: {path}")

# Column spec used in all regression tables
TABCOLS  = r"l *{6}{>{\centering\arraybackslash}X}"
TABCOLS4 = r"l *{4}{>{\centering\arraybackslash}X}"

PORTS = ["P1", "P2", "P3", "P4", "P5", "HML"]

# ── Load result files ──────────────────────────────────────────────────────
print("Loading result files...")
desc    = pd.read_csv(HERE / "descriptive_table.csv")
firms   = pd.read_csv(HERE / "firms_quintiles.csv")
ret     = pd.read_csv(HERE / "portfolio_returns.csv", index_col=0, parse_dates=True)
ff5     = pd.read_csv(HERE / "results_global_FF5_EW.csv",  index_col=0)
ff6     = pd.read_csv(HERE / "results_global_FF6_EW.csv",  index_col=0)
vw      = pd.read_csv(HERE / "results_global_FF5_VW.csv",  index_col=0)
sp_xl   = pd.ExcelFile(HERE / "results_splitsample_FF5_EW.xlsx")
alt_xl  = pd.ExcelFile(HERE / "results_alternative_sorts_FF5.xlsx")
reg_xl  = pd.ExcelFile(HERE / "results_regional_FF5_EW.xlsx")
res_pre  = sp_xl.parse("Pre_COVID_2015_2019",  index_col=0)
res_post = sp_xl.parse("Post_COVID_2020_2025", index_col=0)
res_t3   = alt_xl.parse("Tercile_Sort",        index_col=0)
res_d10  = alt_xl.parse("Decile_Sort",         index_col=0)
sheet_map = {"North_America":"North America","Europe":"Europe",
             "Asia_Pacific":"Asia-Pacific","EM_ex_Asia":"EM ex-Asia"}
reg = {sheet_map[s]: reg_xl.parse(s, index_col=0) for s in reg_xl.sheet_names}

own         = firms["Ownership"]
firm_counts = firms.groupby("Region")["FirmID"].count()

FF5_COLS = ["Beta Mkt-RF", "Beta SMB", "Beta HML", "Beta RMW", "Beta CMA"]
FF6_COLS = FF5_COLS + ["Beta UMD"]

# Map column names from CSV to LaTeX labels
BETA_LABELS = {
    "Beta Mkt-RF": r"$\beta$ MKT-RF",
    "Beta SMB":    r"$\beta$ SMB",
    "Beta HML":    r"$\beta$ HML",
    "Beta RMW":    r"$\beta$ RMW",
    "Beta CMA":    r"$\beta$ CMA",
    "Beta UMD":    r"$\beta$ UMD (Momentum)",
}

def get_beta_cols(data):
    """Find actual beta column names in the dataframe."""
    cols = []
    for c in data.columns:
        if c.startswith("b ") or c.startswith("B ") or "Mkt" in c or "SMB" in c or "HML" in c or "RMW" in c or "CMA" in c or "UMD" in c:
            cols.append(c)
    return cols

def beta_label(col):
    """Convert column name like 'b Mkt-RF' to LaTeX label."""
    c = col.strip()
    if "Mkt" in c: return r"$\beta$ MKT-RF"
    if "SMB" in c: return r"$\beta$ SMB"
    if "HML" in c: return r"$\beta$ HML"
    if "RMW" in c: return r"$\beta$ RMW"
    if "CMA" in c: return r"$\beta$ CMA"
    if "UMD" in c: return r"$\beta$ UMD (Momentum)"
    return c

# ══════════════════════════════════════════════════════════════════════════
# TABLE 1 — Sample Characteristics
# ══════════════════════════════════════════════════════════════════════════
print("Building Table 1...")

ranges = {
    "P1": "20.00--26.46\\%", "P2": "26.48--34.04\\%", "P3": "34.05--43.71\\%",
    "P4": "43.74--56.84\\%", "P5": "56.87--99.81\\%"
}
labels_p = {"P1":"P1 (Low)","P2":"P2","P3":"P3","P4":"P4","P5":"P5 (High)"}
fc  = firms.groupby(["Portfolio","Region"])["FirmID"].count().unstack(fill_value=0)
mo  = firms.groupby("Portfolio")["Ownership"].mean()

L = []  # lines
L.append(r"% TABLE 1 — Sample Characteristics  [auto-generated]")
L.append(r"\begin{table}[htbp]")
L.append(r"\centering\small")
L.append(r"\caption{Sample Characteristics : Insider Ownership Quintile Portfolios}")
L.append(r"\label{tab:sample}")
L.append(r"")
L.append(r"\textit{Panel A: Firm Distribution by Ownership Quintile and Region}\\[4pt]")
L.append(r"\begin{adjustbox}{max width=\textwidth}")
L.append(r"\begin{tabular}{l l r r r r r r}")
L.append(r"\toprule")
L.append(r"\textbf{Portfolio} & \textbf{Ownership Range} & \textbf{N} & \textbf{Mean (\%)}"
         r" & \textbf{Asia-Pac.} & \textbf{Europe} & \textbf{N.\ Am.} & \textbf{EM ex-Asia} \\")
L.append(r"\midrule")

for p in ["P1","P2","P3","P4","P5"]:
    row = fc.loc[p] if p in fc.index else pd.Series(0, index=["Asia-Pacific","Europe","North America","EM ex-Asia"])
    n   = int(row.sum())
    ap  = int(row.get("Asia-Pacific",0))
    eu  = int(row.get("Europe",0))
    na  = int(row.get("North America",0))
    em  = int(row.get("EM ex-Asia",0))
    m   = f2(mo.get(p, np.nan))
    L.append(f"  {labels_p[p]} & {ranges[p]} & {n} & {m} & {ap} & {eu} & {na} & {em} \\\\")

L.append(r"\midrule")
tot_n  = len(firms)
tot_mo = f2(firms["Ownership"].mean())
L.append(f"  \\textbf{{Total}} & & \\textbf{{{tot_n:,}}} & {tot_mo}"
         f" & {int((firms['Region']=='Asia-Pacific').sum())}"
         f" & {int((firms['Region']=='Europe').sum())}"
         f" & {int((firms['Region']=='North America').sum())}"
         f" & {int((firms['Region']=='EM ex-Asia').sum())} \\\\")
L.append(r"\bottomrule")
L.append(r"\end{tabular}")
L.append(r"\end{adjustbox}")
L.append(r"")
L.append(r"\vspace{8pt}")
L.append(r"\textit{Panel B: Portfolio Return Summary Statistics}\\[4pt]")
L.append(r"\begin{tabularx}{\textwidth}{" + TABCOLS + "}")
L.append(r"\toprule")
L.append(r" & \textbf{P1 (Low)} & \textbf{P2} & \textbf{P3}"
         r" & \textbf{P4} & \textbf{P5 (High)} & \textbf{HML (P5$-$P1)} \\")
L.append(r"\midrule")

stats = [
    ("Mean return (\\% per month)",   lambda p: f"{ret[p].mean()*100:.3f}"),
    ("Mean return (\\% per year)",    lambda p: f"{ret[p].mean()*1200:.2f}"),
    ("Std.\\ dev.\\ (\\% per month)", lambda p: f"{ret[p].std()*100:.3f}"),
    ("Minimum (\\% per month)",       lambda p: f"{ret[p].min()*100:.3f}"),
    ("Maximum (\\% per month)",       lambda p: f"{ret[p].max()*100:.3f}"),
    ("Sharpe ratio (annualised)",     lambda p: f"{ret[p].mean()/ret[p].std()*np.sqrt(12):.3f}"),
    ("Observations (months)",         lambda p: str(int(ret[p].notna().sum()))),
]
for lbl, fn in stats:
    vals = " & ".join(fn(p) for p in PORTS)
    L.append(f"  {lbl} & {vals} \\\\")

L.append(r"\bottomrule")
L.append(r"\end{tabularx}")
L.append(r"")
L.append(r"\vspace{4pt}")
L.append(r"\begin{minipage}{\textwidth}")
L.append(r"\footnotesize\textit{Notes:} 1,285 firms sorted into equal-size quintiles by"
         r" insider ownership as of 2025 (Wood, 2025). Returns equal weighted, USD,"
         r" winsorised at the 1st/99th percentile. Sharpe ratio annualised"
         r" $= \bar{r}/\sigma \times \sqrt{12}$. Sample: January 2015 to December 2025 (132 months).")
L.append(r"\end{minipage}")
L.append(r"\end{table}")
write("table_01_sample", L)


# ══════════════════════════════════════════════════════════════════════════
# Helper: regression table  (Tables 2, 3, R3)
# ══════════════════════════════════════════════════════════════════════════

def reg_table(name, caption, label, data, note):
    beta_cols = get_beta_cols(data)
    L = []
    L.append(f"% {name}  [auto-generated]")
    L.append(r"\begin{table}[htbp]")
    L.append(r"\centering\small")
    L.append(f"\\caption{{{caption}}}")
    L.append(f"\\label{{{label}}}")
    L.append(r"\begin{tabularx}{\textwidth}{" + TABCOLS + "}")
    L.append(r"\toprule")
    L.append(r" & \textbf{P1} & \textbf{P2} & \textbf{P3}"
             r" & \textbf{P4} & \textbf{P5} & \textbf{HML} \\")
    L.append(r" & \textbf{(Low)} & & & & \textbf{(High)} & \textbf{(P5$-$P1)} \\")
    L.append(r"\midrule")

    # Alpha row
    alpha_vals = " & ".join(alpha_cell(data.loc[p]) if p in data.index else ":" for p in PORTS)
    L.append(r"\textbf{$\alpha$ (\% per month)} & " + alpha_vals + r" \\")
    tstat_vals = " & ".join(tstat_cell(data.loc[p]) if p in data.index else ":" for p in PORTS)
    L.append(r"\quad[$t$-statistic] & " + tstat_vals + r" \\")
    ann_vals = " & ".join(f3(data.loc[p,"Alpha (% pa)"]) if p in data.index else ":" for p in PORTS)
    L.append(r"$\alpha$ (\% per year) & " + ann_vals + r" \\[4pt]")

    # Factor loadings
    for bc in beta_cols:
        vals = " & ".join(f4(data.loc[p, bc]) if (p in data.index and bc in data.columns) else ":"
                          for p in PORTS)
        L.append(beta_label(bc) + " & " + vals + r" \\")

    L.append(r"\\[-2pt]")

    # Summary stats
    rsq_vals  = " & ".join(f4(data.loc[p,"Adj R\u00b2"]) if p in data.index else ":" for p in PORTS)
    nobs_vals = " & ".join(str(int(data.loc[p,"N months"])) if p in data.index else ":" for p in PORTS)

    # Try alternative column names
    rsq_col = "Adj R²" if "Adj R²" in data.columns else "Adj R2" if "Adj R2" in data.columns else None
    if rsq_col:
        rsq_vals = " & ".join(f4(data.loc[p, rsq_col]) if p in data.index else ":" for p in PORTS)
    nobs_col = "N months" if "N months" in data.columns else None
    if nobs_col:
        nobs_vals = " & ".join(str(int(data.loc[p, nobs_col])) if p in data.index else ":" for p in PORTS)

    L.append(r"Adj.\ $R^{2}$ & " + rsq_vals + r" \\")
    L.append(r"Observations (months) & " + nobs_vals + r" \\")
    L.append(r"\bottomrule")
    L.append(r"\end{tabularx}")
    L.append(r"\vspace{4pt}")
    L.append(r"\begin{minipage}{\textwidth}")
    L.append(r"\footnotesize\textit{Notes:} " + note)
    L.append(r"\end{minipage}")
    L.append(r"\end{table}")
    return L


print("Building Table 2...")
# Find actual column names
print("  FF5 columns:", list(ff5.columns))

write("table_02_ff5_global", reg_table(
    "TABLE 2",
    "Fama French Five Factor Alpha by Insider Ownership Quintile: Global Equal Weighted Portfolios",
    "tab:ff5_global",
    ff5,
    r"FF5 time-series regressions for ownership-sorted portfolios (Fahlenbrach, 2009)."
    r" P1 = 20.00--26.46\%; P5 = 56.87--99.81\%. HML is long P5, short P1."
    r" Equal-weighted returns winsorised at 1st/99th percentile."
    r" Newey-West $t$-statistics (6 lags). January 2015 to December 2025 (132 months)."
    r" $^{***}p{<}0.01$, $^{**}p{<}0.05$, $^{*}p{<}0.10$."
))

print("Building Table 3...")
write("table_03_ff6_robustness", reg_table(
    "TABLE 3",
    "Fama French Six Factor Alpha by Insider Ownership Quintile: Robustness Including Momentum",
    "tab:ff6",
    ff6,
    r"Extends Table~\ref{tab:ff5_global} with the momentum factor (UMD/WML) (Carhart, 1997)."
    r" HML alpha: $0.528\%$ pm (FF6) vs $0.530\%$ (FF5) : momentum does not absorb the alpha (H3)."
    r" $^{***}p{<}0.01$, $^{**}p{<}0.05$, $^{*}p{<}0.10$."
))


# ══════════════════════════════════════════════════════════════════════════
# TABLE 4 — Regional Results
# ══════════════════════════════════════════════════════════════════════════
print("Building Table 4...")

def panel_block(region, panel_label, data):
    n_firms = int(firm_counts.get(region, 0))
    L = []
    L.append(f"\\textit{{{panel_label}: {region} \\quad ($N = {n_firms}$ firms)}}\\\\[3pt]")
    L.append(r"\begin{tabularx}{\textwidth}{" + TABCOLS + "}")
    L.append(r"\toprule")
    L.append(r" & \textbf{P1 (Low)} & \textbf{P2} & \textbf{P3}"
             r" & \textbf{P4} & \textbf{P5 (High)} & \textbf{HML (P5$-$P1)} \\")
    L.append(r"\midrule")

    alpha_r = " & ".join(alpha_cell(data.loc[p]) if p in data.index else ":" for p in PORTS)
    tstat_r = " & ".join(tstat_cell(data.loc[p]) if p in data.index else ":" for p in PORTS)
    ann_r   = " & ".join(f3(data.loc[p,"Alpha (% pa)"]) if p in data.index else ":" for p in PORTS)

    rsq_col = "Adj R²" if "Adj R²" in data.columns else "Adj R2" if "Adj R2" in data.columns else None
    rsq_r = " & ".join(f4(data.loc[p, rsq_col]) if (p in data.index and rsq_col) else ":" for p in PORTS)

    L.append(r"\textbf{$\alpha$ (\% per month)} & " + alpha_r + r" \\")
    L.append(r"\quad[$t$-statistic] & " + tstat_r + r" \\")
    L.append(r"$\alpha$ (\% per year) & " + ann_r + r" \\")
    L.append(r"Adj.\ $R^{2}$ & " + rsq_r + r" \\")
    L.append(r"\bottomrule")
    L.append(r"\end{tabularx}")
    L.append(r"\vspace{6pt}")
    L.append(r"")
    return L

L = []
L.append(r"% TABLE 4 — Regional Results  [auto-generated]")
L.append(r"\begin{table}[htbp]")
L.append(r"\centering\small")
L.append(r"\caption{Regional Fama French Five Factor Alpha by Insider Ownership Quintile (H4 and H5)}")
L.append(r"\label{tab:regional}")
L.append(r"")
for region, panel in [("North America","Panel A"),("Europe","Panel B"),
                       ("Asia-Pacific","Panel C"),("EM ex-Asia","Panel D")]:
    data = reg.get(region)
    if data is not None:
        L.extend(panel_block(region, panel, data))
L.append(r"\begin{minipage}{\textwidth}")
L.append(r"\footnotesize\textit{Notes:} Panel A: North American FF5 factors;"
         r" Panel B: European; Panel C: Asia-Pacific ex-Japan;"
         r" Panel D: global Developed Markets factors (proxy)."
         r" Panel A ($N=141$): HML $\alpha = 1.768\%$ pm ($21.2\%$ pa), $t=1.681^{*}$."
         r" Panel B ($N=109$): HML $\alpha = -0.133\%$ pm, $t=-0.47$, confirms H5."
         r" Newey-West $t$-statistics (6 lags)."
         r" $^{***}p{<}0.01$, $^{**}p{<}0.05$, $^{*}p{<}0.10$.")
L.append(r"\end{minipage}")
L.append(r"\end{table}")
write("table_04_regional", L)


# ══════════════════════════════════════════════════════════════════════════
# TABLE 5 — Hypothesis Summary  (fully hardcoded)
# ══════════════════════════════════════════════════════════════════════════
print("Building Table 5...")
L = []
L.append(r"% TABLE 5 — Hypothesis Summary  [auto-generated]")
L.append(r"\begin{table}[htbp]")
L.append(r"\centering\small")
L.append(r"\caption{Summary of Hypotheses and Empirical Results}")
L.append(r"\label{tab:hypotheses}")
L.append(r"\begin{tabularx}{\textwidth}{c l p{4.4cm} p{4.4cm} c}")
L.append(r"\toprule")
L.append(r"\textbf{H} & \textbf{Hypothesis} & \textbf{Prediction} & \textbf{Evidence} & \textbf{Result} \\")
L.append(r"\midrule")

rows = [
    ("H1","Alpha hypothesis",
     "Positive significant FF5 alpha across all quintiles",
     r"All $\alpha>0$, $t>2.73^{***}$. P5: $1.45\%$ pm ($t=4.93$). HML: $0.53\%$ pm ($t=2.73$).",
     r"\textbf{\ding{51}} Supported"),
    ("H2","Monotonicity",
     r"$\alpha$ increases monotonically with ownership",
     r"P1=0.924\%, P2=1.135\%, P3=1.378\%, P4=1.090\%, P5=1.454\%. P4 dips below P3.",
     r"$\sim$ Partial"),
    ("H3","Robustness",
     "Alpha survives FF6 and value-weighting",
     r"FF6 HML: $0.528\%$ ($t=2.79^{***}$) vs FF5: $0.530\%$. UMD does not absorb alpha.",
     r"\textbf{\ding{51}} Supported"),
    ("H4","Regional robustness",
     "Positive alpha in all regional subsamples",
     r"P5$>0$ everywhere. NA HML significant ($t=1.68^{*}$). APAC $t=1.63$. EM $t=0.94$.",
     r"$\sim$ Partial"),
    ("H5","US vs.\\ Europe",
     "NA alpha exceeds EU alpha",
     r"NA HML $+1.768\%$ pm ($t=1.68^{*}$). EU HML $-0.133\%$ pm ($t=-0.47$).",
     r"\textbf{\ding{51}} Supported"),
]
for h, hyp, pred, ev, res in rows:
    L.append(f"  {h} & {hyp} & {pred} & {ev} & {res} \\\\[4pt]")

L.append(r"\bottomrule")
L.append(r"\end{tabularx}")
L.append(r"\vspace{4pt}")
L.append(r"\begin{minipage}{\textwidth}")
L.append(r"\footnotesize\textit{Notes:} \ding{51} = supported; $\sim$ = partially supported."
         r" $^{***}p{<}0.01$, $^{**}p{<}0.05$, $^{*}p{<}0.10$. Newey-West (6 lags)."
         r" 1,285 firms, January 2015 to December 2025 (132 months).")
L.append(r"\end{minipage}")
L.append(r"\end{table}")
write("table_05_hypotheses", L)


# ══════════════════════════════════════════════════════════════════════════
# TABLE R1 — Split Sample
# ══════════════════════════════════════════════════════════════════════════
print("Building Table R1...")

def ss_panel(panel_label, period, nobs, data):
    L = []
    L.append(f"\\textit{{{panel_label} : {period} ({nobs} months)}}\\\\[3pt]")
    L.append(r"\begin{tabularx}{\textwidth}{" + TABCOLS + "}")
    L.append(r"\toprule")
    L.append(r" & \textbf{P1 (Low)} & \textbf{P2} & \textbf{P3}"
             r" & \textbf{P4} & \textbf{P5 (High)} & \textbf{HML} \\")
    L.append(r"\midrule")
    alpha_r = " & ".join(alpha_cell(data.loc[p]) if p in data.index else ":" for p in PORTS)
    tstat_r = " & ".join(tstat_cell(data.loc[p]) if p in data.index else ":" for p in PORTS)
    ann_r   = " & ".join(f3(data.loc[p,"Alpha (% pa)"]) if p in data.index else ":" for p in PORTS)
    rsq_col = "Adj R²" if "Adj R²" in data.columns else "Adj R2" if "Adj R2" in data.columns else None
    rsq_r   = " & ".join(f4(data.loc[p, rsq_col]) if (p in data.index and rsq_col) else ":" for p in PORTS)
    L.append(r"\textbf{$\alpha$ (\% per month)} & " + alpha_r + r" \\")
    L.append(r"\quad[$t$-statistic] & " + tstat_r + r" \\")
    L.append(r"$\alpha$ (\% per year) & " + ann_r + r" \\")
    L.append(r"Adj.\ $R^{2}$ & " + rsq_r + r" \\")
    L.append(r"Observations & " + " & ".join([str(nobs)]*6) + r" \\")
    L.append(r"\bottomrule")
    L.append(r"\end{tabularx}")
    L.append(r"\vspace{6pt}")
    L.append(r"")
    return L

L = []
L.append(r"% TABLE R1 — Split Sample  [auto-generated]")
L.append(r"\begin{table}[htbp]")
L.append(r"\centering\small")
L.append(r"\caption{Split Sample Robustness: FF5 Alpha by Ownership Quintile}")
L.append(r"\label{tab:splitsample}")
L.append(r"")
L.extend(ss_panel("Panel A: Pre-COVID",  "January 2015 to December 2019", 60,  res_pre))
L.extend(ss_panel("Panel B: Post-COVID", "January 2020 to December 2025", 72, res_post))
L.append(r"\begin{minipage}{\textwidth}")
L.append(r"\footnotesize\textit{Notes:} Quintile assignments identical across panels (2025 snapshot)."
         r" Tests whether alpha is spuriously driven by end-of-sample ownership concentration."
         r" Newey-West $t$-statistics (6 lags). $^{***}p{<}0.01$, $^{**}p{<}0.05$, $^{*}p{<}0.10$.")
L.append(r"\end{minipage}")
L.append(r"\end{table}")
write("table_R1_splitsample", L)


# ══════════════════════════════════════════════════════════════════════════
# TABLE R2 — Alternative Sorts
# ══════════════════════════════════════════════════════════════════════════
print("Building Table R2...")

t1_break = own.quantile(1/3)
t2_break = own.quantile(2/3)
cnt_t3 = {
    "T1": int((own < t1_break).sum()),
    "T2": int(((own >= t1_break) & (own < t2_break)).sum()),
    "T3": int((own >= t2_break).sum())
}

def ta(p, d): return alpha_cell(d.loc[p]) if p in d.index else ":"
def tt(p, d): return tstat_cell(d.loc[p]) if p in d.index else ":"
def tan(p, d): return f3(d.loc[p,"Alpha (% pa)"]) if p in d.index else ":"
def trsq(p, d):
    c = "Adj R²" if "Adj R²" in d.columns else "Adj R2" if "Adj R2" in d.columns else None
    return f4(d.loc[p, c]) if (p in d.index and c) else ":"

L = []
L.append(r"% TABLE R2 — Alternative Sorts  [auto-generated]")
L.append(r"\begin{table}[htbp]")
L.append(r"\centering\small")
L.append(r"\caption{Alternative Portfolio Sort Specifications : FF5 Alpha Robustness}")
L.append(r"\label{tab:alt_sorts}")
L.append(r"")
L.append(f"\\textit{{Panel A: Tercile Sort (cut-offs: {t1_break:.2f}\\% and {t2_break:.2f}\\%)}}\\\\[3pt]")
L.append(r"\begin{tabularx}{\textwidth}{" + TABCOLS4 + "}")
L.append(r"\toprule")
L.append(r" & \textbf{T1 (Low)} & \textbf{T2 (Mid)} & \textbf{T3 (High)} & \textbf{HML (T3$-$T1)} \\")
L.append(r" & 20--31\% & 31--48\% & 48--100\% & \\")
L.append(r"\midrule")
L.append(f"  N firms & {cnt_t3['T1']} & {cnt_t3['T2']} & {cnt_t3['T3']} & --- \\\\")
L.append(r"\textbf{$\alpha$ (\% per month)} & " +
         f"{ta('T1',res_t3)} & {ta('T2',res_t3)} & {ta('T3',res_t3)} & {ta('HML',res_t3)} \\\\")
L.append(r"\quad[$t$-statistic] & " +
         f"{tt('T1',res_t3)} & {tt('T2',res_t3)} & {tt('T3',res_t3)} & {tt('HML',res_t3)} \\\\")
L.append(r"$\alpha$ (\% per year) & " +
         f"{tan('T1',res_t3)} & {tan('T2',res_t3)} & {tan('T3',res_t3)} & {tan('HML',res_t3)} \\\\")
L.append(r"Adj.\ $R^{2}$ & " +
         f"{trsq('T1',res_t3)} & {trsq('T2',res_t3)} & {trsq('T3',res_t3)} & {trsq('HML',res_t3)} \\\\")
L.append(r"Observations & 132 & 132 & 132 & 132 \\")
L.append(r"\bottomrule")
L.append(r"\end{tabularx}")
L.append(r"\vspace{8pt}")
L.append(r"")

# Decile panel
DECILES = [f"D{i}" for i in range(1,11)]
dec_ranges = {
    "D1":"20.0--23.4\\%","D2":"23.4--26.4\\%","D3":"26.5--29.6\\%",
    "D4":"29.6--34.0\\%","D5":"34.0--38.8\\%","D6":"38.8--43.6\\%",
    "D7":"43.7--49.9\\%","D8":"49.9--56.7\\%","D9":"56.8--68.3\\%","D10":"68.4--99.8\\%"
}
DALL = DECILES + ["HML"]

L.append(r"\textit{Panel B: Decile Sort ($\approx$128 firms per decile)}\\[3pt]")
L.append(r"\begin{adjustbox}{max width=\textwidth}")
L.append(r"\small")
L.append(r"\begin{tabular}{l " + "c " * 11 + "}")
L.append(r"\toprule")
hdr = " & ".join(f"\\textbf{{{d}}}" for d in DECILES) + r" & \textbf{HML} \\"
L.append("  & " + hdr)
rng = " & ".join(f"\\tiny {dec_ranges[d]}" for d in DECILES) + r" & D10$-$D1 \\"
L.append(r"\textit{Range} & " + rng)
L.append(r"\midrule")

alpha_d = " & ".join(ta(d, res_d10) for d in DALL)
L.append(r"\textbf{$\alpha$ (\% pm)} & " + alpha_d + r" \\")
tstat_d = " & ".join(tt(d, res_d10) for d in DALL)
L.append(r"\quad[$t$-stat] & " + tstat_d + r" \\")
ann_d   = " & ".join(tan(d, res_d10) for d in DALL)
L.append(r"$\alpha$ (\% pa) & " + ann_d + r" \\")
rsq_d   = " & ".join(trsq(d, res_d10) for d in DALL)
L.append(r"Adj.\ $R^{2}$ & " + rsq_d + r" \\")
obs_d   = " & ".join(["132"] * 11)
L.append(r"Obs. & " + obs_d + r" \\")
L.append(r"\bottomrule")
L.append(r"\end{tabular}")
L.append(r"\end{adjustbox}")
L.append(r"")
L.append(r"\vspace{4pt}")
L.append(r"\begin{minipage}{\textwidth}")
L.append(r"\footnotesize\textit{Notes:} Tests sensitivity of Table~\ref{tab:ff5_global} to"
         r" portfolio cut-off choice. Panel A: equal-size tercile sort."
         r" Panel B: equal-size decile sort ($\approx$128 firms each)."
         r" HML is long D10/T3 and short D1/T1."
         r" Newey-West $t$-statistics (6 lags). $^{***}p{<}0.01$, $^{**}p{<}0.05$, $^{*}p{<}0.10$.")
L.append(r"\end{minipage}")
L.append(r"\end{table}")
write("table_R2_alt_sorts", L)


# ══════════════════════════════════════════════════════════════════════════
# TABLE R3 — Value Weighted
# ══════════════════════════════════════════════════════════════════════════
print("Building Table R3...")
write("table_R3_vw", reg_table(
    "TABLE R3",
    "Fama French Five Factor Alpha: Value Weighted Portfolios (Robustness)",
    "tab:vw",
    vw,
    r"Replicates Table~\ref{tab:ff5_global} with value weighted returns."
    r" Weights: lagged market capitalisation (FactSet)."
    r" Tests whether equal weighted alpha is driven by smaller firms."
    r" $^{***}p{<}0.01$, $^{**}p{<}0.05$, $^{*}p{<}0.10$."
))


# ══════════════════════════════════════════════════════════════════════════
# Move figures
# ══════════════════════════════════════════════════════════════════════════
print("\nMoving figures...")
for pdf in HERE.glob("figure*.pdf"):
    shutil.copy2(pdf, OUT_FIG / pdf.name)
    print(f"  Copied: {pdf.name}")
for png in HERE.glob("figure*.png"):
    shutil.copy2(png, OUT_FIG / png.name)

print(f"\nDone. Tables in: {OUT_TEX.resolve()}")
print(f"Figures in:      {OUT_FIG.resolve()}")
