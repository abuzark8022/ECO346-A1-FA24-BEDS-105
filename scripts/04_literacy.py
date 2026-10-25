#!/usr/bin/env python3
"""04_literacy.py -- Question 5, second language (Python).

Recomputes the Question 3 band table and redraws the Question 3 figure using
Python, saving it as figures/literacy_python.svg.  The two band tables must
agree to one decimal place.

Run from the repository root:  python scripts/04_literacy.py
"""

import os
import sys

import numpy as np
import pandas as pd
import matplotlib
matplotlib.use("Agg")
import matplotlib.pyplot as plt

# Keep text as real <text> elements in the SVG (not embedded glyph paths), so
# the file is readable and behaves like the R/svglite output.
matplotlib.rcParams["svg.fonttype"] = "none"

# Work from the repository root no matter where the script is called from.
ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
os.chdir(ROOT)


def weighted_mean(x, w):
    """Weighted mean: sum(w*x) / sum(w)."""
    return np.sum(w * x) / np.sum(w)


# --- Variant 5: Khyber Pakhtunkhwa, urban (prov_code 1, urban_rural 2) -----
hh = pd.read_csv("data/pslm_hh_subset.csv")
per = pd.read_csv("data/pslm_person_subset.csv")
d = per.merge(hh[["hhid", "hhweight", "prov_code", "urban_rural"]],
              on="hhid", how="inner")
d = d[(d["prov_code"] == 1) & (d["urban_rural"] == 2)]

# --- Literacy sample: persons aged 10-64 whose literate value is present ---
lsamp = d[(d["age"] >= 10) & (d["age"] <= 64) & d["literate"].notna()].copy()

# --- Five-year age bands with midpoints 12, 17, ..., 62 -------------------
edges = list(range(10, 66, 5))                 # 10,15,...,65
labels = [f"{a}-{a + 4}" for a in edges[:-1]]  # '10-14', ..., '60-64'
mids = {lab: a + 2 for lab, a in zip(labels, edges[:-1])}
lsamp["band"] = pd.cut(lsamp["age"], bins=edges, right=False, labels=labels)
lsamp["mid"] = lsamp["band"].map(mids).astype(float)
lsamp["sexf"] = lsamp["sex"].map({1: "Men", 2: "Women"})

# --- Weighted literacy rate and count for each band x sex -----------------
recs = []
for lab in labels:
    for sexf in ("Men", "Women"):
        g = lsamp[(lsamp["band"] == lab) & (lsamp["sexf"] == sexf)]
        recs.append(dict(
            band=lab,
            mid=mids[lab],
            sexf=sexf,
            n=int(len(g)),
            rate=100.0 * weighted_mean(g["literate"].to_numpy(float),
                                       g["hhweight"].to_numpy(float)) if len(g) else np.nan))
tab = pd.DataFrame(recs)
tab["thin"] = tab["n"] < 30

# --- Print the band table in the same shape as the R script ---------------
men = tab[tab["sexf"] == "Men"].reset_index(drop=True)
women = tab[tab["sexf"] == "Women"].reset_index(drop=True)
print("\n================ QUESTION 3 BAND TABLE (Python) ================")
print(f"{'band':>7} {'men_n':>6} {'men_rate':>9} {'women_n':>8} {'women_rate':>11}")
for i in range(len(men)):
    print(f"{men['band'][i]:>7} {men['n'][i]:>6} {men['rate'][i]:>9.1f} "
          f"{women['n'][i]:>8} {women['rate'][i]:>11.1f}")

# --- The band with the largest gap between men and women ------------------
gap_df = pd.DataFrame({"band": men["band"], "men": men["rate"], "women": women["rate"]})
gap_df["gap"] = (gap_df["men"] - gap_df["women"]).abs()
gi = int(gap_df["gap"].idxmax())
gap_band = gap_df["band"][gi]
gap_value = gap_df["gap"][gi]
gap_mid = float(mids[gap_band])

# --- Save the tidy table for the R-vs-Python check ------------------------
tab.to_csv("scripts/band_table_python.csv", index=False)

# --- The figure (mirrors the R redesign) ----------------------------------
fig, ax = plt.subplots(figsize=(8.5, 5.5))
colors = {"Men": "#2C7FB8", "Women": "#D95F02"}

for sexf in ("Men", "Women"):
    g = tab[tab["sexf"] == sexf].sort_values("mid")
    ax.plot(g["mid"], g["rate"], color=colors[sexf], linewidth=1.6, zorder=2)
    for _, row in g.iterrows():
        if row["thin"]:
            ax.plot(row["mid"], row["rate"], marker="o", markersize=7,
                    markerfacecolor="white", markeredgecolor=colors[sexf],
                    markeredgewidth=1.8, linestyle="none", zorder=3)
        else:
            ax.plot(row["mid"], row["rate"], marker="o", markersize=7,
                    markerfacecolor=colors[sexf], markeredgecolor=colors[sexf],
                    linestyle="none", zorder=3)
    last = g.dropna(subset=["rate"]).iloc[-1]
    ax.annotate(sexf, (last["mid"], last["rate"]),
                xytext=(6, 0), textcoords="offset points",
                color=colors[sexf], fontweight="bold", va="center", fontsize=11)

# Direct annotation of the widest-gap band
ax.plot([gap_mid, gap_mid], [gap_df["men"][gi], gap_df["women"][gi]],
        color="grey", linewidth=1.0, zorder=1)
ax.annotate(f"Largest gap at ages {gap_band}:\n{gap_value:.1f} points",
            (gap_mid, max(gap_df["men"][gi], gap_df["women"][gi]) + 4),
            ha="center", color="#333333", fontweight="bold", fontsize=9)

ax.set_xlim(9, 70)
ax.set_ylim(0, 100)
ax.set_xticks(range(12, 63, 10))
ax.set_yticks(range(0, 101, 20))
ax.set_xlabel("Age band midpoint (years)")
ax.set_ylabel("Weighted literacy rate (%)")
ax.set_title(
    "Women's literacy falls sharply with age while men's stays high, "
    f"peaking in a {gap_value:.0f}-point gap at ages {gap_band}",
    fontsize=10.5, fontweight="bold")
fig.text(0.01, 0.005,
         "Source: PSLM 2019-20 person and household files (course data repository).\n"
         "Sample: persons aged 10-64 with a non-missing literacy value in KP urban; "
         "weighted by the household weight (hhweight).\n"
         "Hollow markers mark thin cells: a band-sex group with fewer than 30 observations.",
         fontsize=6.5, color="#4D4D4D", ha="left", va="bottom")
ax.spines[["top", "right"]].set_visible(False)

fig.tight_layout(rect=[0, 0.10, 1, 1])
fig.savefig("figures/literacy_python.svg", format="svg")
plt.close(fig)

# --- Key numbers ----------------------------------------------------------
def rate(b, s):
    return float(tab[(tab["band"] == b) & (tab["sexf"] == s)]["rate"].iloc[0])

print("\n================ QUESTION 5 KEY NUMBERS (Python) ================")
print(f"Women 15-19 weighted literacy rate      : {rate('15-19', 'Women'):.1f} %")
print(f"Men   15-19 weighted literacy rate      : {rate('15-19', 'Men'):.1f} %")
print(f"Women 55-59 weighted literacy rate      : {rate('55-59', 'Women'):.1f} %")
print(f"Men   55-59 weighted literacy rate      : {rate('55-59', 'Men'):.1f} %")
print(f"Band with the largest gap               : {gap_band}")
print(f"Largest gap (points)                    : {gap_value:.1f}")
print("================================================================")
