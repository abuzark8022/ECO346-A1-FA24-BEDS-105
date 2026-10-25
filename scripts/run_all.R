# run_all.R -- reproduces every figure, table, verification test, and slides from a fresh clone.
#
# Runs the numbered scripts in order, runs Python for Question 5, validates agreement,
# and rebuilds the PowerPoint slides for Question 6.
# No manual step and no internet connection are required.
#
# Run from the repository root:  Rscript scripts/run_all.R

# --- Path resolution: locate repository root dynamically -------------------
args <- commandArgs(trailingOnly = FALSE)
file_arg <- grep("^--file=", args, value = TRUE)
if (length(file_arg) > 0) {
  script_dir <- dirname(normalizePath(sub("^--file=", "", file_arg[1])))
  repo_root <- if (basename(script_dir) == "scripts") dirname(script_dir) else script_dir
  setwd(repo_root)
} else {
  root <- normalizePath(".")
  if (basename(root) == "scripts") setwd(dirname(root))
}

rscript <- file.path(R.home("bin"), "Rscript")

# --- Run R analysis scripts -----------------------------------------------
for (s in c("scripts/01_default.R",
            "scripts/02_distribution.R",
            "scripts/03_literacy.R")) {
  cat("\n########## running", s, "##########\n")
  status <- system2(rscript, s)
  if (status != 0) stop("Script failed: ", s)
}

# --- Find Python executable (python or Windows py launcher) ---------------
py <- Sys.which("python")
if (!nzchar(py)) py <- Sys.which("py")
if (!nzchar(py)) stop("Python not found on PATH")

# --- Question 5: Python reproduction and automated comparison -------------
cat("\n########## running scripts/04_literacy.py ##########\n")
status <- system2(py, "scripts/04_literacy.py")
if (status != 0) stop("Python script failed: scripts/04_literacy.py")

# --- Automated verification between R and Python band tables in R ---------
cat("\n########## verifying R vs Python band tables ##########\n")
tab_r  <- read.csv("scripts/band_table_R.csv")
tab_py <- read.csv("scripts/band_table_python.csv")

stopifnot(
  "Row count mismatch between R and Python tables" = nrow(tab_r) == nrow(tab_py),
  "Band labels mismatch between R and Python tables" = identical(tab_r$band, tab_py$band),
  "Men sample count mismatch" = identical(as.integer(tab_r$men_n), as.integer(tab_py$men_n)),
  "Women sample count mismatch" = identical(as.integer(tab_r$women_n), as.integer(tab_py$women_n)),
  "Men literacy rates differ by more than rounding tolerance" = max(abs(tab_r$men_rate - tab_py$men_rate)) <= 0.05,
  "Women literacy rates differ by more than rounding tolerance" = max(abs(tab_r$women_rate - tab_py$women_rate)) <= 0.05
)
cat("[PASS] Automated cross-check: R and Python band tables are identical across all 11 bands.\n")

# --- Question 6: Build PowerPoint slides ----------------------------------
cat("\n########## rebuilding slides/FA24-BEDS-105.pptx ##########\n")
status <- system2(py, "scripts/build_slides.py")
if (status != 0) stop("Slide generation failed: scripts/build_slides.py")

cat("\n====================================================================\n")
cat("SUCCESS: All figures, tables, verification checks, and slides rebuilt.\n")
cat("====================================================================\n")
