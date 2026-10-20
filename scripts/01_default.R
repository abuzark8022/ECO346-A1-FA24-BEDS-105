# 01_default.R -- Question 4 baseline.
# Draws literate against age with R's defaults untouched and saves figures/default.svg.
# Run from the repository root:  Rscript scripts/01_default.R

library(ggplot2)
library(svglite)

if (basename(getwd()) == "scripts") setwd("..")

# --- Variant 5: Khyber Pakhtunkhwa, urban (prov_code 1, urban_rural 2) ----
hh  <- read.csv("data/pslm_hh_subset.csv")
per <- read.csv("data/pslm_person_subset.csv")
d   <- merge(per, hh[, c("hhid", "hhweight", "prov_code", "urban_rural")],
             by = "hhid")
d   <- d[d$prov_code == 1 & d$urban_rural == 2, ]

# --- Literacy sample: persons aged 10-64 whose literate value is present ---
literacy_sample <- subset(d, age >= 10 & age <= 64 & !is.na(literate))

# --- R's defaults, untouched ---------------------------------------------
p <- ggplot(literacy_sample, aes(x = age, y = literate)) + geom_point()

# --- Every figure is saved as SVG by the script that draws it ------------
ggsave("figures/default.svg", plot = p, device = svglite, width = 7, height = 5)
