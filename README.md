# ECO346-A1-FA24-BEDS-105

**Data Visualization Techniques for Economists - Assignment 1 (Fall 2026)**

| | |
|---|---|
| **Registration number** | FA24-BEDS-105 |
| **Variant** | **5** = Khyber Pakhtunkhwa, urban (prov_code 1, urban_rural 2) |
| **Second language (Question 5)** | Python (pandas + matplotlib) |
| **Data** | PSLM 2019-20 person and household files (subset), in `data/` |

The variant is the last two digits of the registration number, `05`, taken modulo 8: `05 = 0 x 8 + 5`, so variant 5. Every figure and number below uses only persons and households in **KP urban**. Population estimates (means, quantiles, rates, densities) are weighted by the household weight `hhweight`, attached from the household file by `hhid`. Sample counts ($n$) and the unweighted median are reported unweighted to distinguish sample size from population estimates.

---

## Environment setup and package installation

Before running the scripts for the first time on a fresh machine or lab computer, install the required packages:

```r
# In R console / RStudio:
install.packages(c("ggplot2", "svglite"))
```

```bash
# In terminal / Command Prompt (Python). On Windows, use `py` if `python` is not on PATH:
pip install numpy pandas matplotlib
py -m pip install numpy pandas matplotlib
```

Rebuilding the PowerPoint file is optional (the `.pptx` is already in `slides/`). It needs extra packages:

```bash
pip install python-pptx svglib reportlab lxml
```

---

## How to reproduce everything

After the packages above are installed, no internet connection is required. From the repository root:

```bash
Rscript scripts/run_all.R
```

That runs scripts 01-03 (R) and 04 (Python), checks that the band tables agree, and rebuilds the slides if the optional PowerPoint packages are installed. Figures and numbers do not depend on those extra packages. On Windows, the runner looks for `python`, `python3`, then `py`.

Individual scripts can also be executed independently in order:

| Script | Question | Output |
|---|---|---|
| `scripts/01_default.R` | 4 (default) | `figures/default.svg` |
| `scripts/02_distribution.R` | 2 | `figures/distribution.svg` |
| `scripts/03_literacy.R` | 3 (redesign) | `figures/literacy.svg` + `scripts/band_table_R.csv` |
| `scripts/04_literacy.py` | 5 | `figures/literacy_python.svg` + `scripts/band_table_python.csv` (automated test against R) |
| `scripts/build_slides.py` | 6 | `slides/FA24-BEDS-105.pptx` (optional rebuild of the 3-slide presentation) |

Every script includes explicit data validation asserting that the household file contains unique `hhid` keys, the merge preserves person record counts, survey weights (`hhweight`) are strictly positive, and categorical variables (`sex`, `literate`) follow valid codings.

All figures are written as **SVG** by the script that draws them. Every quantile, including the unweighted median, follows the assignment definition exactly: sort by the variable, accumulate the weight share, and take the smallest value at which the cumulative share reaches or exceeds *q*. Equal weights are used for the unweighted median. R's built-in `quantile()` and `median()` are not used.

---

## Question 2 - The distribution of education spending

![Weighted histogram of annual education spending on a student](figures/distribution.svg)

**The figure.** A weighted histogram of annual education spending on students in KP urban, drawn on a log scale (bins 0.1 log10 units wide, from Rs 100 to Rs 1,000,000). The vertical axis is density per log10 unit; the horizontal axis is labelled in rupees at each power of ten. Dashed lines mark the weighted median and the weighted 90th percentile.

**Reported numbers.**

| Quantity | Value |
|---|---|
| Students (ed_total > 0) | 1,061 |
| Students analysed (Rs 100 - Rs 1,000,000) | 1,060 |
| Students excluded (outside the range) | 1 |
| Weighted median | Rs 9,500 |
| Unweighted median | Rs 10,000 |
| Weighted 10th percentile | Rs 3,150 |
| Weighted 90th percentile | Rs 39,000 |
| Ratio of 90th to 10th percentile | 12.38 |
| Mean survey weight below unweighted median | 274.9 |
| Mean survey weight above unweighted median | 228.4 |

**What it shows (<= 80 words).** Education spending is right-skewed: the weighted population median is Rs 9,500, and the top tenth exceed Rs 39,000 (12.38 times the 10th percentile of Rs 3,150). Weighting pulls the centre down from the unweighted sample median of Rs 10,000 because students below that median have higher average survey weights (274.9 vs 228.4), so lower-spending households represent more of the urban KP population.

---

## Question 3 - The relationship between literacy and age

![Weighted literacy rate by age band, women and men](figures/literacy.svg)

**The figure.** The weighted literacy rate against five-year age bands (plotted at band midpoints), with women and men as two series of markers joined by lines. Each series is labelled directly at its end; there is no legend. Hollow markers mark thin cells (a band-sex group with fewer than 30 observations). The widest gap between women and men is annotated directly.

**Full band table.** Counts and weighted literacy rates (%) by age band and sex.

| Age band | Men, n | Men, % | Women, n | Women, % |
|---|---|---|---|---|
| 10-14 | 228 | 81.5 | 215 | 70.3 |
| 15-19 | 193 | 89.7 | 163 | 68.1 |
| 20-24 | 149 | 80.8 | 129 | 47.8 |
| 25-29 | 112 | 85.9 | 135 | 48.1 |
| 30-34 | 89 | 88.2 | 126 | 43.6 |
| 35-39 | 116 | 83.3 | 105 | 50.7 |
| 40-44 | 100 | 77.1 | 110 | 29.7 |
| 45-49 | 67 | 65.5 | 60 | 32.8 |
| 50-54 | 53 | 55.1 | 50 | 16.1 |
| 55-59 | 41 | 58.8 | 41 | 16.2 |
| 60-64 | 44 | 57.1 | 31 | 19.3 |

No band-sex cell falls below 30 observations, so no marker is hollow in this variant.

**Reported numbers.**

| Quantity | Value |
|---|---|
| Women aged 15-19 | 68.1 % |
| Men aged 15-19 | 89.7 % |
| Women aged 55-59 | 16.2 % |
| Men aged 55-59 | 58.8 % |
| Band with the largest gap | **40-44** (47.4 points) |

**Interpretation (<= 100 words).** Literacy is high among teenagers and lower in older bands, and men outpace women at every age. The gap is widest at ages 40-44 (47.4 points: men 77.1%, women 29.7%), not among the oldest adults. Women aged 15-19 have a literacy rate of 68.1% versus 16.2% at ages 55-59; men are at 89.7% and 58.8%. The pattern is consistent with later cohorts receiving more schooling, but a single cross-section cannot identify causes or separate cohort effects from aging.

---

## Question 4 - The makeover

The default figure (`figures/default.svg`, committed first) is `ggplot(data, aes(age, literate)) + geom_point()` with R's defaults untouched. The redesign is the Question 3 figure (`figures/literacy.svg`).

**Explanation of the redesign (<= 150 words).** Three changes, each tied to a principle from Chapters 1-4. First, the raw scatter is replaced by **conditional means**: the weighted literacy rate in each five-year band, so the age pattern is no longer hidden by thousands of overlapping 0/1 points (**overplotting**). Second, the redesign keeps **position**, the highest-ranked channel, for the two quantities that vary continuously - age and the literacy rate - and uses colour plus direct labels only for the categorical split by sex (**perceptual ranking of channels**). Third, the legend is dropped in favour of end labels, cutting non-data ink (**data-ink ratio**). The default figure still shows the actual spread and the discreteness of the 0/1 outcome, which the redesign hides by plotting one summary value per cell.

---

## Question 5 - A second language (Python)

`scripts/04_literacy.py` recomputes the Question 3 band table, saves it to `scripts/band_table_python.csv`, and redraws the figure in Python, saved as `figures/literacy_python.svg`. An automated verification step asserts that the counts and all eleven age-band weighted literacy rates agree with R within rounding tolerance (also cross-checked automatically in `scripts/run_all.R`).

**Reported numbers.** The literacy rate of **women aged 15-19** is **68.1%** from R and **68.1%** from Python.

**One difference and how it was fixed (<= 80 words).** The default axis range and tick spacing differed: ggplot expanded the panel around the data, while matplotlib padded the axes differently, so the two figures would not overlay. I set the same limits in both scripts (age 9-70, rate 0-100) and the same breaks (age 12, 22, ..., 62; rate 0, 20, ..., 100).

---

## Repository structure

```
ECO346-A1-FA24-BEDS-105/
+-- README.md
+-- data/       pslm_person_subset.csv, pslm_hh_subset.csv
+-- scripts/    01_default.R, 02_distribution.R, 03_literacy.R,
+               04_literacy.py, build_slides.py, run_all.R
+-- figures/    default.svg, distribution.svg, literacy.svg, literacy_python.svg
+-- slides/     FA24-BEDS-105.pptx
```
