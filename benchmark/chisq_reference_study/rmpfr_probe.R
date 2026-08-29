#!/usr/bin/env Rscript
# =====================================================================
# Probe: at what shape does MPFR's incomplete gamma stop working?
#
# WHY THIS IS SHAPED ODDLY. mpfr_gamma_inc fails by tripping an internal
# C assertion, which calls abort(). An abort cannot be caught by tryCatch
# or withCallingHandlers - it terminates the R process immediately. So
# this script CANNOT collect results and report at the end.
#
# Instead it prints its intent BEFORE each attempt and flushes, then
# prints the result after. The last "attempting" line with no matching
# "ok" line is the shape that killed it. Read the boundary off the tail
# of the output.
#
# USAGE
#   Rscript rmpfr_probe.R [precision_bits]
# =====================================================================

suppressPackageStartupMessages(library(Rmpfr))

args <- commandArgs(trailingOnly = TRUE)
PREC <- if (length(args) >= 1) as.integer(args[1]) else 400L

cat(sprintf("Rmpfr %s, MPFR %s, precision %d bits\n",
            as.character(packageVersion("Rmpfr")),
            as.character(mpfrVersion()), PREC))
cat("Evaluating Q(a, a) = igamma(a, a)/gamma(a) at increasing shape.\n")
cat("The last 'attempting' line without a matching 'ok' is the failure shape.\n\n")

shapes <- c(1e3, 1e4, 1e5, 3e5, 5e5, 1e6, 3e6, 5e6, 7e6, 1e7,
            1.5e7, 2e7, 3e7, 5e7)

for (a_dbl in shapes) {
  cat(sprintf("attempting a = %.3e ... ", a_dbl)); flush.console()
  a <- mpfr(a_dbl, PREC)
  t0 <- Sys.time()
  v <- igamma(a, a) / gamma(a)
  dt <- as.numeric(difftime(Sys.time(), t0, units = "secs"))
  cat(sprintf("ok  Q=%s  (%.2fs)\n", format(v, digits = 12), dt))
  flush.console()
}

cat("\nAll probed shapes succeeded.\n")
