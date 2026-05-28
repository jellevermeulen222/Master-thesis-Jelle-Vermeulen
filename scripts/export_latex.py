"""
export_latex.py
───────────────
Reads all result files produced by thesis_analysis.py + thesis_output.py
and writes one .tex file per table into ../output/.
Also moves generated figure PDFs to ../figures/.

Run from the scripts/ folder:
    python export_latex.py

Output files
    ../output/table_01_sample.tex
    ../output/table_02_ff5_global.tex
    ../output/table_03_ff6_robustness.tex
    ../output/table_04_regional.tex
    ../output/table_05_hypotheses.tex
    ../output/table_R1_splitsample.tex
    ../output/table_R2_alt_sorts.tex
    ../output/table_R3_vw.tex
    ../figures/figure_*.pdf   (moved from scripts/)
"""

import pandas as pd
import numpy as np
import shutil
import os
from pathlib import Path

# ── Paths ─────────────────────────────────────────────────────────────────
HERE    = Path(__file__).parent          # scripts/
OUT_TEX = HERE / ".." / "output"        # output/
OUT_FIG = HERE / ".." / "figures"       # figures/
OUT_TEX.mkdir(exist_ok=True)
OUT_FIG.mkdir(exist_ok=True)

# ── Helpers ───────────────────────────────────────────────────────────────

def stars(pv):
    if pd.isna(pv): return ""
    return "$^{***}$" if pv < 0.01 else "$^{**}$" if pv < 0.05 else "$^{*}$" if pv < 0.10 else ""

def f4(v):
    """4 decimal places, or em-dash if missing."""
    return "---" if pd.isna(v) else f"{v:.4f}"

def f3(v):
    return "---" if pd.isna(v) else f"{v:.3f}"

def f2(v):
    return "---" if pd.isna(v) else f"{v:.2f}"

def alpha_cell(row):
    """Bold alpha with significance stars."""
    a  = row["Alpha (% pm)"]
    pv = row["p-value"]
    if pd.isna(a): return "---"
    return f"\\textbf{{{a:.4f}}}{stars(pv)}"

def tstat_cell(row):
    t = row["t-stat (NW)"]
    return "---" if pd.isna(t) else f"({t:.3f})"

def write(name, content):
    path = OUT_TEX / f"{name}.tex"
    path.write_text(content, encoding="utf-8")
    print(f"  Saved: {path}")

PORTS      = ["P1", "P2", "P3", "P4", "P5", "HML"]
PORT_HDRS  = (r"\textbf{P1} & \textbf{P2} & \textbf{P3} & "
              r"\textbf{P4} & \textbf{P5} & \textbf{HML} \\")
PORT_HDRS2 = (r"\textbf{P1 (Low)} & \textbf{P2} & \textbf{P3} & "
              r"\textbf{P4} & \textbf{P5 (High)} & \textbf{HML (P5$-$P1)} \\")

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

# ══════════════════════════════════════════════════════════════════════════
# TABLE 1 — Sample Characteristics
# ══════════════════════════════════════════════════════════════════════════
print("Building Table 1...")

ranges = {"P1":"20.00--26.46\\%","P2":"26.48--34.04\\%","P3":"34.05--43.71\\%",
          "P4":"43.74--56.84\\%","P5":"56.87--99.81\\%"}
labels = {"P1":"P1 (Low)","P2":"P2","P3":"P3","P4":"P4","P5":"P5 (High)"}
fc     = firms.groupby(["Portfolio","Region"])["FirmID"].count().unstack(fill_value=0)
mo     = firms.groupby("Portfolio")["Ownership"].mean()
regions_ord = ["Asia-Pacific","Europe","North America","EM ex-Asia"]

panelA_rows = ""
for p in ["P1","P2","P3","P4","P5"]:
    row = fc.loc[p] if p in fc.index else pd.Series(0, index=regions_ord)
    n   = int(row.sum())
    ap  = int(row.get("Asia-Pacific",0))
    eu  = int(row.get("Europe",0))
    na  = int(row.get("North America",0))
    em  = int(row.get("EM ex-Asia",0))
    m   = f2(mo.get(p, np.nan))
    panelA_rows += f"  {labels[p]} & {ranges[p]} & {n} & {m} & {ap} & {eu} & {na} & {em} \\\\\n"

tot_n  = len(firms)
tot_mo = f2(firms["Ownership"].mean())
tot_ap = int((firms["Region"]=="Asia-Pacific").sum())
tot_eu = int((firms["Region"]=="Europe").sum())
tot_na = int((firms["Region"]=="North America").sum())
tot_em = int((firms["Region"]=="EM ex-Asia").sum())

panelB_rows = ""
stat_defs = [
    ("Mean return (\\% per month)",  lambda p: f"{ret[p].mean()*100:.3f}"),
    ("Mean return (\\% per year)",   lambda p: f"{ret[p].mean()*1200:.2f}"),
    ("Std.\\ dev.\\ (\\% per month)",lambda p: f"{ret[p].std()*100:.3f}"),
    ("Minimum (\\% per month)",      lambda p: f"{ret[p].min()*100:.3f}"),
    ("Maximum (\\% per month)",      lambda p: f"{ret[p].max()*100:.3f}"),
    ("Sharpe ratio (annualised)",    lambda p: f"{ret[p].mean()/ret[p].std()*np.sqrt(12):.3f}"),
    ("Observations (months)",        lambda p: str(int(ret[p].notna().sum()))),
]
for lbl, fn in stat_defs:
    vals = " & ".join(fn(p) for p in PORTS)
    panelB_rows += f"  {lbl} & {vals} \\\\\n"

t1 = rf"""% TABLE 1 — Sample Characteristics  [auto-generated by export_latex.py]
\begin{{table}}[htbp]
\centering\small
\caption{{Sample Characteristics --- Insider Ownership Quintile Portfolios}}
\label{{tab:sample}}

\textit{{Panel A: Firm Distribution by Ownership Quintile and Region}}\\[4pt]
\begin{{tabularx}}{{\textwidth}}{{l l r r r r r r}}
\toprule
\textbf{{Portfolio}} & \textbf{{Ownership Range}} & \textbf{{N}} & \textbf{{Mean (\%)}}
  & \textbf{{Asia-Pac.}} & \textbf{{Europe}} & \textbf{{N.\ America}} & \textbf{{EM ex-Asia}} \\
\midrule
{panelA_rows.rstrip()}
\midrule
  \textbf{{Total}} & & \textbf{{{tot_n:,}}} & {tot_mo}
  & {tot_ap} & {tot_eu} & {tot_na} & {tot_em} \\
\bottomrule
\end{{tabularx}}

\vspace{{8pt}}
\textit{{Panel B: Portfolio Return Summary Statistics}}\\[4pt]
\begin{{tabularx}}{{\textwidth}}{{l *{{6}}{{>{{\\centering\\arraybackslash}}X}}}}
\toprule
 & \textbf{{P1 (Low)}} & \textbf{{P2}} & \textbf{{P3}}
 & \textbf{{P4}} & \textbf{{P5 (High)}} & \textbf{{HML (P5$-$P1)}} \\
\midrule
{panelB_rows.rstrip()}
\bottomrule
\end{{tabularx}}

\vspace{{4pt}}
\begin{{minipage}}{{\textwidth}}
\footnotesize\textit{{Notes:}} 1,285 firms from Wood (2025) sorted into equal-size
quintiles by insider ownership as of 2025. Returns are equal-weighted, USD,
winsorised at the 1st/99th percentile. Sharpe ratio annualised $= \bar{{r}} /
\sigma \times \sqrt{{12}}$. Sample: January 2015--December 2025 (132 months).
\end{{minipage}}
\end{{table}}
"""
write("table_01_sample", t1)


# ══════════════════════════════════════════════════════════════════════════
# Helper: generic regression table  (Tables 2, 3, R3)
# ══════════════════════════════════════════════════════════════════════════

def reg_table(label, caption, data, factor_cols, note):
    alpha_row  = " & ".join(alpha_cell(data.loc[p]) if p in data.index else "---" for p in PORTS)
    tstat_row  = " & ".join(tstat_cell(data.loc[p]) if p in data.index else "---" for p in PORTS)
    annalpha   = " & ".join(f3(data.loc[p,"Alpha (% pa)"]) if p in data.index else "---" for p in PORTS)
    rsq_row    = " & ".join(f4(data.loc[p,"Adj R²"])       if p in data.index else "---" for p in PORTS)
    nobs_row   = " & ".join(str(int(data.loc[p,"N months"])) if p in data.index else "---" for p in PORTS)

    fcol_labels = {"β Mkt-RF":"$\\beta$ MKT-RF","β SMB":"$\\beta$ SMB",
                   "β HML":"$\\beta$ HML","β RMW":"$\\beta$ RMW",
                   "β CMA":"$\\beta$ CMA","β UMD":"$\\beta$ UMD (Momentum)"}
    beta_rows = ""
    for fc in factor_cols:
        vals = " & ".join(f4(data.loc[p,fc]) if (p in data.index and fc in data.columns) else "---"
                          for p in PORTS)
        beta_rows += f"  {fcol_labels.get(fc,fc)} & {vals} \\\\\n"

    return rf"""% {label}  [auto-generated by export_latex.py]
\begin{{table}}[htbp]
\centering\small
\caption{{{caption}}}
\label{{{label.lower().replace(' ','_')}}}
\begin{{tabularx}}{{\textwidth}}{{l *{{6}}{{>{{\\centering\\arraybackslash}}X}}}}
\toprule
 & \textbf{{P1}} & \textbf{{P2}} & \textbf{{P3}}
 & \textbf{{P4}} & \textbf{{P5}} & \textbf{{HML}} \\
 & \textbf{{(Low)}} & & & & \textbf{{(High)}} & \textbf{{(P5$-$P1)}} \\
\midrule
\textbf{{$\alpha$ (\% per month)}} & {alpha_row} \\
\quad[$t$-statistic]               & {tstat_row} \\
$\alpha$ (\% per year)             & {annalpha}  \\[4pt]
{beta_rows.rstrip()}
\\[4pt]
  Adj.\ $R^{{2}}$          & {rsq_row}  \\
  Observations (months)    & {nobs_row} \\
\bottomrule
\end{{tabularx}}
\vspace{{4pt}}
\begin{{minipage}}{{\textwidth}}
\footnotesize\textit{{Notes:}} {note}
\end{{minipage}}
\end{{table}}
"""

FF5_COLS = ["β Mkt-RF","β SMB","β HML","β RMW","β CMA"]
FF6_COLS = FF5_COLS + ["β UMD"]

print("Building Table 2...")
write("table_02_ff5_global", reg_table(
    "tab:ff5\\_global",
    "Fama-French Five-Factor Alpha by Insider Ownership Quintile --- Global Equal-Weighted Portfolios",
    ff5, FF5_COLS,
    "FF5 time-series regressions for ownership-sorted portfolios "
    r"(Fahlenbrach, 2009). P1 = 20.00--26.46\%; P5 = 56.87--99.81\%. "
    "HML is long P5, short P1. Equal-weighted returns winsorised at 1st/99th percentile. "
    r"Newey-West $t$-statistics (6 lags). Sample: January 2015--December 2025 (132 months). "
    r"$^{***}p{<}0.01$, $^{**}p{<}0.05$, $^{*}p{<}0.10$."
))

print("Building Table 3...")
write("table_03_ff6_robustness", reg_table(
    "tab:ff6",
    "Fama-French Six-Factor Alpha by Insider Ownership Quintile --- Robustness Including Momentum",
    ff6, FF6_COLS,
    r"Extends Table~\ref{tab:ff5\_global} with the momentum factor (UMD/WML). "
    r"HML alpha: $0.528\%$ pm (FF6) vs $0.530\%$ (FF5) --- momentum does not absorb the alpha (H3). "
    r"$^{***}p{<}0.01$, $^{**}p{<}0.05$, $^{*}p{<}0.10$."
))


# ══════════════════════════════════════════════════════════════════════════
# TABLE 4 — Regional Results
# ══════════════════════════════════════════════════════════════════════════
print("Building Table 4...")

def panel_block(region, panel_label, data):
    n_firms = int(firm_counts.get(region, 0))
    alpha_r = " & ".join(alpha_cell(data.loc[p]) if p in data.index else "---" for p in PORTS)
    tstat_r = " & ".join(tstat_cell(data.loc[p]) if p in data.index else "---" for p in PORTS)
    ann_r   = " & ".join(f3(data.loc[p,"Alpha (% pa)"]) if p in data.index else "---" for p in PORTS)
    rsq_r   = " & ".join(f4(data.loc[p,"Adj R²"])       if p in data.index else "---" for p in PORTS)
    return rf"""
\textit{{{panel_label}: {region} \quad ($N = {n_firms}$ firms)}}\\[3pt]
\begin{{tabularx}}{{\textwidth}}{{l *{{6}}{{>{{\\centering\\arraybackslash}}X}}}}
\toprule
 & \textbf{{P1 (Low)}} & \textbf{{P2}} & \textbf{{P3}}
 & \textbf{{P4}} & \textbf{{P5 (High)}} & \textbf{{HML (P5$-$P1)}} \\
\midrule
\textbf{{$\alpha$ (\% per month)}} & {alpha_r} \\
\quad[$t$-statistic]               & {tstat_r} \\
$\alpha$ (\% per year)             & {ann_r}   \\
Adj.\ $R^{{2}}$                    & {rsq_r}   \\
\bottomrule
\end{{tabularx}}
\vspace{{6pt}}
"""

panels = (
    panel_block("North America", "Panel A", reg["North America"]) +
    panel_block("Europe",        "Panel B", reg["Europe"]) +
    panel_block("Asia-Pacific",  "Panel C", reg["Asia-Pacific"]) +
    panel_block("EM ex-Asia",    "Panel D (global factors as proxy)", reg["EM ex-Asia"])
)

t4 = rf"""% TABLE 4 — Regional Results  [auto-generated by export_latex.py]
\begin{{table}}[htbp]
\centering\small
\caption{{Regional Fama-French Five-Factor Alpha by Insider Ownership Quintile (H4 and H5)}}
\label{{tab:regional}}
{panels}
\begin{{minipage}}{{\textwidth}}
\footnotesize\textit{{Notes:}}
Panel A: Fama-French North American factors; Panel B: European; Panel C: Asia-Pacific ex-Japan;
Panel D: global Developed Markets factors (no EM ex-Asia regional factors published by French).
Panel A ($N=141$): HML $\alpha = 1.768\%$ pm ($21.2\%$ pa), $t=1.681^{{*}}$.
Panel B ($N=109$): HML $\alpha = -0.133\%$ pm, $t=-0.47$ --- confirms H5.
Newey-West $t$-statistics (6 lags). $^{{***}}p{{<}}0.01$, $^{{**}}p{{<}}0.05$, $^{{*}}p{{<}}0.10$.
\end{{minipage}}
\end{{table}}
"""
write("table_04_regional", t4)


# ══════════════════════════════════════════════════════════════════════════
# TABLE 5 — Hypothesis Summary  (fully hardcoded)
# ══════════════════════════════════════════════════════════════════════════
print("Building Table 5...")
t5 = r"""% TABLE 5 — Hypothesis Summary  [auto-generated by export_latex.py]
\begin{table}[htbp]
\centering\small
\caption{Summary of Hypotheses and Empirical Results}
\label{tab:hypotheses}
\begin{tabularx}{\textwidth}{c l p{4.6cm} p{4.6cm} c}
\toprule
\textbf{H} & \textbf{Hypothesis} & \textbf{Prediction} & \textbf{Evidence} & \textbf{Result} \\
\midrule
H1 & Alpha hypothesis
   & Positive significant FF5 alpha across all quintiles
   & All $\alpha>0$, $t>2.73^{***}$. P5: $1.45\%$ pm ($t=4.93$). HML: $0.53\%$ pm ($t=2.73$).
   & \textbf{\ding{51}} Supported \\[4pt]
H2 & Monotonicity
   & $\alpha$ increases monotonically with ownership
   & P1=0.924\%, P2=1.135\%, P3=1.378\%, P4=1.090\%, P5=1.454\%. P4 dips below P3.
   & $\sim$ Partial \\[4pt]
H3 & Robustness
   & Alpha survives FF6 and value-weighting
   & FF6 HML: $0.528\%$ ($t=2.79^{***}$) vs FF5: $0.530\%$. UMD does not absorb alpha.
   & \textbf{\ding{51}} Supported \\[4pt]
H4 & Regional robustness
   & Positive alpha in all regions
   & P5$>0$ everywhere. NA HML significant ($t=1.68^{*}$). APAC $t=1.63$. EM $t=0.94$.
   & $\sim$ Partial \\[4pt]
H5 & US vs.\ Europe
   & NA alpha exceeds EU alpha
   & NA HML $+1.768\%$ pm ($t=1.68^{*}$). EU HML $-0.133\%$ pm ($t=-0.47$).
   & \textbf{\ding{51}} Supported \\
\bottomrule
\end{tabularx}
\vspace{4pt}
\begin{minipage}{\textwidth}
\footnotesize\textit{Notes:} \ding{51} = supported; $\sim$ = partially supported.
$^{***}p{<}0.01$, $^{**}p{<}0.05$, $^{*}p{<}0.10$. Newey-West (6 lags).
1,285 firms, January 2015--December 2025 (132 months).
\end{minipage}
\end{table}
"""
write("table_05_hypotheses", t5)


# ══════════════════════════════════════════════════════════════════════════
# TABLE R1 — Split-Sample
# ══════════════════════════════════════════════════════════════════════════
print("Building Table R1...")

def ss_panel(label, period, data, nobs):
    alpha_r = " & ".join(alpha_cell(data.loc[p]) if p in data.index else "---" for p in PORTS)
    tstat_r = " & ".join(tstat_cell(data.loc[p]) if p in data.index else "---" for p in PORTS)
    ann_r   = " & ".join(f3(data.loc[p,"Alpha (% pa)"]) if p in data.index else "---" for p in PORTS)
    rsq_r   = " & ".join(f4(data.loc[p,"Adj R²"])       if p in data.index else "---" for p in PORTS)
    return rf"""
\textit{{{label} --- {period} ({nobs} months)}}\\[3pt]
\begin{{tabularx}}{{\textwidth}}{{l *{{6}}{{>{{\\centering\\arraybackslash}}X}}}}
\toprule
 & \textbf{{P1 (Low)}} & \textbf{{P2}} & \textbf{{P3}}
 & \textbf{{P4}} & \textbf{{P5 (High)}} & \textbf{{HML}} \\
\midrule
\textbf{{$\alpha$ (\% per month)}} & {alpha_r} \\
\quad[$t$-statistic]               & {tstat_r} \\
$\alpha$ (\% per year)             & {ann_r}   \\
Adj.\ $R^{{2}}$                    & {rsq_r}   \\
Observations                       & {" & ".join([str(nobs)]*6)} \\
\bottomrule
\end{{tabularx}}
\vspace{{6pt}}
"""

tR1 = rf"""% TABLE R1 — Split-Sample  [auto-generated by export_latex.py]
\begin{{table}}[htbp]
\centering\small
\caption{{Split-Sample Robustness --- FF5 Alpha by Ownership Quintile}}
\label{{tab:splitsample}}
{ss_panel("Panel A: Pre-COVID",  "January 2015--December 2019", res_pre,  60)}
{ss_panel("Panel B: Post-COVID", "January 2020--December 2025", res_post, 72)}
\begin{{minipage}}{{\textwidth}}
\footnotesize\textit{{Notes:}} Quintile assignments are identical across panels (2025 snapshot).
Tests whether the alpha pattern is spuriously driven by end-of-sample ownership concentration.
Newey-West $t$-statistics (6 lags). $^{{***}}p{{<}}0.01$, $^{{**}}p{{<}}0.05$, $^{{*}}p{{<}}0.10$.
\end{{minipage}}
\end{{table}}
"""
write("table_R1_splitsample", tR1)


# ══════════════════════════════════════════════════════════════════════════
# TABLE R2 — Alternative Sorts
# ══════════════════════════════════════════════════════════════════════════
print("Building Table R2...")

t1_break = own.quantile(1/3)
t2_break = own.quantile(2/3)
cnt_t3   = {"T1": int((own<t1_break).sum()),
            "T2": int(((own>=t1_break)&(own<t2_break)).sum()),
            "T3": int((own>=t2_break).sum())}

T3_PORTS = ["T1","T2","T3","HML"]
def t3_cell(p, data, fn):
    pk = "HML" if p=="HML" else p
    return fn(data.loc[pk]) if pk in data.index else "---"

def t3_alpha(p): return alpha_cell(res_t3.loc[p]) if p in res_t3.index else "---"
def t3_tstat(p): return tstat_cell(res_t3.loc[p]) if p in res_t3.index else "---"
def t3_ann(p):   return f3(res_t3.loc[p,"Alpha (% pa)"]) if p in res_t3.index else "---"
def t3_rsq(p):   return f4(res_t3.loc[p,"Adj R²"])       if p in res_t3.index else "---"
def t3_n(p):     return str(cnt_t3.get(p,"---"))

# Decile ranges (hardcoded from thesis_output.py)
dec_ranges = {"D1":"20.0--23.4\\%","D2":"23.4--26.4\\%","D3":"26.5--29.6\\%",
              "D4":"29.6--34.0\\%","D5":"34.0--38.8\\%","D6":"38.8--43.6\\%",
              "D7":"43.7--49.9\\%","D8":"49.9--56.7\\%","D9":"56.8--68.3\\%",
              "D10":"68.4--99.8\\%"}
DECILES = [f"D{i}" for i in range(1,11)]

def dcell(d, fn):
    return fn(res_d10.loc[d]) if d in res_d10.index else "---"

dec_hdr   = " & ".join(f"\\textbf{{{d}}}" for d in DECILES) + r" & \textbf{HML} \\"
dec_range = " & ".join(f"\\tiny {dec_ranges[d]}" for d in DECILES) + r" & D10$-$D1 \\"
dec_alpha = " & ".join(alpha_cell(res_d10.loc[d]) if d in res_d10.index else "---" for d in DECILES) \
            + " & " + (alpha_cell(res_d10.loc["HML"]) if "HML" in res_d10.index else "---") + r" \\"
dec_tstat = " & ".join(tstat_cell(res_d10.loc[d]) if d in res_d10.index else "---" for d in DECILES) \
            + " & " + (tstat_cell(res_d10.loc["HML"]) if "HML" in res_d10.index else "---") + r" \\"
dec_ann   = " & ".join(f3(res_d10.loc[d,"Alpha (% pa)"]) if d in res_d10.index else "---" for d in DECILES) \
            + " & " + (f3(res_d10.loc["HML","Alpha (% pa)"]) if "HML" in res_d10.index else "---") + r" \\"
dec_rsq   = " & ".join(f4(res_d10.loc[d,"Adj R²"]) if d in res_d10.index else "---" for d in DECILES) \
            + " & " + (f4(res_d10.loc["HML","Adj R²"]) if "HML" in res_d10.index else "---") + r" \\"

tR2 = rf"""% TABLE R2 — Alternative Sorts  [auto-generated by export_latex.py]
\begin{{table}}[htbp]
\centering\small
\caption{{Alternative Portfolio Sort Specifications --- FF5 Alpha Robustness}}
\label{{tab:alt_sorts}}

\textit{{Panel A: Tercile Sort (cut-offs: {t1_break:.2f}\% and {t2_break:.2f}\%)}}\\[3pt]
\begin{{tabularx}}{{\textwidth}}{{l *{{4}}{{>{{\\centering\\arraybackslash}}X}}}}
\toprule
 & \textbf{{T1 (Low)}} & \textbf{{T2 (Mid)}} & \textbf{{T3 (High)}} & \textbf{{HML (T3$-$T1)}} \\
 & 20--31\% & 31--48\% & 48--100\% & \\
\midrule
N firms          & {t3_n("T1")} & {t3_n("T2")} & {t3_n("T3")} & --- \\
\textbf{{$\alpha$ (\% pm)}} & {t3_alpha("T1")} & {t3_alpha("T2")} & {t3_alpha("T3")} & {t3_alpha("HML")} \\
\quad[$t$-stat]  & {t3_tstat("T1")} & {t3_tstat("T2")} & {t3_tstat("T3")} & {t3_tstat("HML")} \\
$\alpha$ (\% pa) & {t3_ann("T1")} & {t3_ann("T2")} & {t3_ann("T3")} & {t3_ann("HML")} \\
Adj.\ $R^{{2}}$  & {t3_rsq("T1")} & {t3_rsq("T2")} & {t3_rsq("T3")} & {t3_rsq("HML")} \\
\bottomrule
\end{{tabularx}}

\vspace{{8pt}}
\textit{{Panel B: Decile Sort ($\approx$128 firms per decile)}}\\[3pt]
\begin{{adjustbox}}{{max width=\textwidth}}
\small
\begin{{tabular}}{{l *{{11}}{{c}}}}
\toprule
 & {dec_hdr}
\textit{{Range}} & {dec_range}
\midrule
\textbf{{$\alpha$ (\% pm)}} & {dec_alpha}
\quad[$t$-stat] & {dec_tstat}
$\alpha$ (\% pa) & {dec_ann}
Adj.\ $R^{{2}}$ & {dec_rsq}
\bottomrule
\end{{tabular}}
\end{{adjustbox}}

\vspace{{4pt}}
\begin{{minipage}}{{\textwidth}}
\footnotesize\textit{{Notes:}} Alternative cut-offs to test sensitivity of Table~\ref{{tab:ff5_global}}.
Panel B decile ranges narrow at the bottom (D1: 20--23\%) due to clustering at the 20\% threshold.
HML is long D10 / T3 and short D1 / T1.
Newey-West $t$-statistics (6 lags). $^{{***}}p{{<}}0.01$, $^{{**}}p{{<}}0.05$, $^{{*}}p{{<}}0.10$.
\end{{minipage}}
\end{{table}}
"""
write("table_R2_alt_sorts", tR2)


# ══════════════════════════════════════════════════════════════════════════
# TABLE R3 — Value-Weighted
# ══════════════════════════════════════════════════════════════════════════
print("Building Table R3...")
write("table_R3_vw", reg_table(
    "tab:vw",
    "Fama-French Five-Factor Alpha --- Value-Weighted Portfolios (Robustness)",
    vw, FF5_COLS,
    r"Replicates Table~\ref{tab:ff5\_global} with value-weighted returns. "
    "Weights: lagged market capitalisation (FactSet). Tests whether equal-weighted alpha "
    r"is driven by smaller firms. $^{***}p{<}0.01$, $^{**}p{<}0.05$, $^{*}p{<}0.10$."
))


# ══════════════════════════════════════════════════════════════════════════
# Move figures to ../figures/
# ══════════════════════════════════════════════════════════════════════════
print("\nMoving figures...")
for pdf in HERE.glob("figure*.pdf"):
    dest = OUT_FIG / pdf.name
    shutil.copy2(pdf, dest)
    print(f"  Copied: {pdf.name} → figures/")
for png in HERE.glob("figure*.png"):
    dest = OUT_FIG / png.name
    shutil.copy2(png, dest)

print(f"\nDone. Tables in: {OUT_TEX.resolve()}")
print(f"Figures in:      {OUT_FIG.resolve()}")
