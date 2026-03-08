"""
Section 2: CUSIP 31418C5Z3 (Fannie Mae Pool MA3563)
- Derives monthly CPR from PoolTalk factor history
- Overlays with 30-yr PMMS from FRED
- Generates chart and analysis text
"""
import numpy as np
import pandas as pd
import matplotlib
matplotlib.use("Agg")
import matplotlib.pyplot as plt
import matplotlib.dates as mdates
import urllib.request
import os

OUT_DIR = r"c:\Users\nduta\OneDrive\Desktop\New folder (2)"
CHART_DIR = os.path.join(OUT_DIR, "charts")

# ── 1. Factor data from PoolTalk (scraped manually) ──────────────────────────
factor_data = {
    "2018-12": 1.00000000,
    "2019-01": 0.99666077, "2019-02": 0.99204106, "2019-03": 0.98601842,
    "2019-04": 0.97813946, "2019-05": 0.96071356, "2019-06": 0.93182598,
    "2019-07": 0.90169869, "2019-08": 0.84657600, "2019-09": 0.80047886,
    "2019-10": 0.74717638, "2019-11": 0.69430918, "2019-12": 0.65809120,
    "2020-01": 0.62228692, "2020-02": 0.59381142, "2020-03": 0.56476105,
    "2020-04": 0.52722780, "2020-05": 0.48195286, "2020-06": 0.44570577,
    "2020-07": 0.40733942, "2020-08": 0.37767858, "2020-09": 0.35023022,
    "2020-10": 0.32625636, "2020-11": 0.30208620, "2020-12": 0.28048911,
    "2021-01": 0.25993041, "2021-02": 0.24401320, "2021-03": 0.22785116,
    "2021-04": 0.20834577, "2021-05": 0.19196575, "2021-06": 0.18001966,
    "2021-07": 0.16744060, "2021-08": 0.15831874, "2021-09": 0.14970149,
    "2021-10": 0.14128669, "2021-11": 0.13331248, "2021-12": 0.12646906,
    "2022-01": 0.12076683, "2022-02": 0.11665040, "2022-03": 0.11232780,
    "2022-04": 0.10680204, "2022-05": 0.10351604, "2022-06": 0.10104890,
    "2022-07": 0.09913771, "2022-08": 0.09754058, "2022-09": 0.09565617,
    "2022-10": 0.09436273, "2022-11": 0.09345447, "2022-12": 0.09230293,
    "2023-01": 0.09123589, "2023-02": 0.09063405, "2023-03": 0.08979104,
    "2023-04": 0.08929145, "2023-05": 0.08847434, "2023-06": 0.08757917,
    "2023-07": 0.08663781, "2023-08": 0.08584490, "2023-09": 0.08499985,
    "2023-10": 0.08442349, "2023-11": 0.08359801, "2023-12": 0.08301750,
    "2024-01": 0.08255006, "2024-02": 0.08187656, "2024-03": 0.08113725,
    "2024-04": 0.08045630, "2024-05": 0.07962410, "2024-06": 0.07905292,
    "2024-07": 0.07849000, "2024-08": 0.07763867, "2024-09": 0.07704279,
    "2024-10": 0.07652208, "2024-11": 0.07563260, "2024-12": 0.07514356,
    "2025-01": 0.07468053, "2025-02": 0.07418124, "2025-03": 0.07369025,
    "2025-04": 0.07331939, "2025-05": 0.07269755, "2025-06": 0.07218977,
    "2025-07": 0.07147019, "2025-08": 0.07085694, "2025-09": 0.07038710,
    "2025-10": 0.06972338, "2025-11": 0.06885678, "2025-12": 0.06821332,
    "2026-01": 0.06766046, "2026-02": 0.06720381, "2026-03": 0.06669364,
}

df = pd.DataFrame(list(factor_data.items()), columns=["Month", "Factor"])
df["Date"] = pd.to_datetime(df["Month"])
df = df.sort_values("Date").reset_index(drop=True)

# ── 2. Derive CPR from sequential factor changes ──────────────────────────────
# CPR = 1 - (Factor_t / Factor_{t-1})^12
# This gives the annualized conditional prepayment rate from month-to-month factor decline
df["SMM"] = 1 - (df["Factor"] / df["Factor"].shift(1))
df["CPR"] = (1 - (1 - df["SMM"]) ** 12) * 100   # CPR in %
df = df.dropna(subset=["CPR"])

# Add pool age (months since issuance Dec-2018)
issue_date = pd.Timestamp("2018-12-01")
df["Age_Months"] = ((df["Date"].dt.year - issue_date.year)*12 +
                     (df["Date"].dt.month - issue_date.month))

print("Factor-derived CPR sample:")
print(df[["Date","Factor","SMM","CPR","Age_Months"]].head(15).to_string(index=False))
print(f"\nPeak CPR: {df['CPR'].max():.2f}% in {df.loc[df['CPR'].idxmax(),'Date'].strftime('%b %Y')}")
print(f"Latest CPR (Mar 2026): {df.iloc[-1]['CPR']:.2f}%")
print(f"Age at peak: {df.loc[df['CPR'].idxmax(),'Age_Months']} months")

# ── 3. Fetch PMMS from FRED ──────────────────────────────────────────────────
pmms_url = "https://fred.stlouisfed.org/graph/fredgraph.csv?id=MORTGAGE30US"
req = urllib.request.Request(pmms_url, headers={"User-Agent": "Mozilla/5.0"})
resp = urllib.request.urlopen(req, timeout=15)
pmms_raw = resp.read().decode()
pmms_lines = [l.split(",") for l in pmms_raw.strip().split("\n")[1:]]
df_pmms = pd.DataFrame(pmms_lines, columns=["Date","PMMS"])
df_pmms["Date"] = pd.to_datetime(df_pmms["Date"])
df_pmms["PMMS"] = pd.to_numeric(df_pmms["PMMS"], errors="coerce")
df_pmms = df_pmms.dropna()
# Filter to pool life
df_pmms = df_pmms[(df_pmms["Date"] >= "2019-01-01") & (df_pmms["Date"] <= "2026-03-31")]

print(f"\nPMMS: fetched {len(df_pmms)} observations")
print(f"PMMS range: {df_pmms['PMMS'].min():.2f}% to {df_pmms['PMMS'].max():.2f}%")

# ── 4. Chart ─────────────────────────────────────────────────────────────────
fig, ax1 = plt.subplots(figsize=(13, 6))

# CPR
color_cpr = "#2563EB"
ax1.bar(df["Date"], df["CPR"], width=20, color=color_cpr, alpha=0.75, label="Monthly CPR (%)")
ax1.set_xlabel("Date", fontsize=11)
ax1.set_ylabel("CPR (%)", color=color_cpr, fontsize=11)
ax1.tick_params(axis="y", labelcolor=color_cpr)
ax1.set_ylim(0, df["CPR"].max() * 1.25)

# PMMS on secondary axis
ax2 = ax1.twinx()
color_pmms = "#DC2626"
ax2.plot(df_pmms["Date"], df_pmms["PMMS"], color=color_pmms, lw=2,
         label="30-Yr Primary Mortgage Rate (PMMS)", alpha=0.9)
ax2.set_ylabel("30-Yr Mortgage Rate (%)", color=color_pmms, fontsize=11)
ax2.tick_params(axis="y", labelcolor=color_pmms)
ax2.set_ylim(2, df_pmms["PMMS"].max() + 1.5)

# Coupon line
coupon = 4.0
ax2.axhline(coupon, color="gray", lw=1.2, linestyle="--", alpha=0.7, label=f"Pool Coupon {coupon}%")

# Annotate peak
peak_idx = df["CPR"].idxmax()
peak_date = df.loc[peak_idx, "Date"]
peak_cpr  = df.loc[peak_idx, "CPR"]
ax1.annotate(f"Peak: {peak_cpr:.1f}%\n{peak_date.strftime('%b %Y')}",
             xy=(peak_date, peak_cpr),
             xytext=(peak_date + pd.DateOffset(months=4), peak_cpr * 0.95),
             fontsize=8.5, color=color_cpr,
             arrowprops=dict(arrowstyle="->", color=color_cpr, lw=1.2))

# Latest CPR
latest_date = df.iloc[-1]["Date"]
latest_cpr  = df.iloc[-1]["CPR"]
ax1.annotate(f"Latest: {latest_cpr:.1f}%\n{latest_date.strftime('%b %Y')}",
             xy=(latest_date, latest_cpr),
             xytext=(latest_date - pd.DateOffset(months=18), latest_cpr + 5),
             fontsize=8.5, color=color_cpr,
             arrowprops=dict(arrowstyle="->", color=color_cpr, lw=1.2))

# X-axis formatting
ax1.xaxis.set_major_formatter(mdates.DateFormatter("%Y-%m"))
ax1.xaxis.set_major_locator(mdates.MonthLocator(bymonth=[1,7]))
plt.xticks(rotation=45, ha="right", fontsize=8)

# Combined legend
lines1, labels1 = ax1.get_legend_handles_labels()
lines2, labels2 = ax2.get_legend_handles_labels()
ax1.legend(lines1 + lines2, labels1 + labels2, loc="upper right", fontsize=9)

ax1.set_title(
    "CUSIP 31418C5Z3 (Fannie Mae Pool MA3563) – Monthly CPR vs 30-Yr PMMS\n"
    "Coupon: 4.0% | Issued: Dec 2018 | Original UPB: $6.39B",
    fontweight="bold", fontsize=11
)
ax1.grid(axis="y", alpha=0.25)
ax1.set_xlim(pd.Timestamp("2019-01-01"), pd.Timestamp("2026-06-01"))

plt.tight_layout()
chart_path = os.path.join(CHART_DIR, "sec2_cusip_cpr_pmms.png")
fig.savefig(chart_path, bbox_inches="tight", dpi=150)
plt.close(fig)
print(f"\nChart saved: {chart_path}")

# ── 5. Print analysis stats ────────────────────────────────────────────────────
first_30 = df[df["Age_Months"] <= 30]
print("\n=== ANALYSIS STATS ===")
print(f"Issuance to month 30 CPR range: {first_30['CPR'].min():.1f}% to {first_30['CPR'].max():.1f}%")
print(f"Peak CPR: {peak_cpr:.2f}% in {peak_date.strftime('%B %Y')} (age {df.loc[peak_idx,'Age_Months']} months)")
print(f"Latest CPR: {latest_cpr:.2f}% ({latest_date.strftime('%B %Y')}, age {df.iloc[-1]['Age_Months']} months)")
# PMMS at peak
pmms_at_peak = df_pmms[df_pmms["Date"] <= peak_date].iloc[-1]["PMMS"]
print(f"30-yr PMMS near peak: {pmms_at_peak:.2f}%")
print(f"Pool coupon: {coupon}% — rate was {coupon-pmms_at_peak:.2f}% {'below' if pmms_at_peak<coupon else 'above'} market")
