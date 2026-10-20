# run_all.R -- reproduces every figure and number from a fresh clone.
#
# Runs the numbered scripts in order, then the Python script for Question 5.
# No manual step and no internet connection are required.
#
# Run from the repository root:  Rscript scripts/run_all.R

root <- normalizePath(".")
if (basename(root) == "scripts") { root <- dirname(root); setwd(root) }

rscript <- file.path(R.home("bin"), "Rscript")

for (s in c("scripts/01_default.R",
            "scripts/02_distribution.R",
            "scripts/03_literacy.R")) {
  cat("\n########## running", s, "##########\n")
  status <- system2(rscript, s)
  if (status != 0) stop("Script failed: ", s)
}

# Question 5 lives in Python; fall back to the "py" launcher on Windows.
cat("\n########## running scripts/04_literacy.py ##########\n")
py <- Sys.which("python")
if (!nzchar(py)) py <- Sys.which("py")
if (!nzchar(py)) stop("Python not found on PATH")
status <- system2(py, "scripts/04_literacy.py")
if (status != 0) stop("Python script failed")

cat("\nAll figures and numbers reproduced.\n")
