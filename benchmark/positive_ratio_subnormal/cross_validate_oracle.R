#!/usr/bin/env Rscript
# cross_validate_oracle.R
# =======================
#
# Independent high-precision evaluation of the reference functions this study
# depends on, using MPFR via Rmpfr.
#
# WHY THIS EXISTS
#   The positive-ratio study (#13) measures VBA against mpmath. A single-oracle
#   study is only as good as that oracle, and self-consistency at one precision
#   proves convergence, not correctness. MPFR is a genuinely independent C
#   implementation with no shared lineage with mpmath's pure Python.
#
# WHY THE PRECISION IS ADAPTIVE
#   MPFR's igamma is the UPPER incomplete gamma, so the regularized lower tail
#   is 1 - igamma(a,z)/gamma(a). When P is tiny that subtraction cancels
#   catastrophically: at 200 bits, a true P of 1E-157 comes back as exactly
#   zero. The precision is therefore chosen from the expected magnitude of P
#   and the result is recomputed at double that precision; a value that moves
#   is reported rather than trusted. This is the same complement-cancellation
#   that #13 addresses with -Expm1(LogP) instead of 1 - Exp(LogP).
#
# INPUT   points.csv with columns a_dec, z_dec, surface
# OUTPUT  r_oracle.csv with columns a_dec, z_dec, surface, value, bits, stable
#
# Usage:  Rscript cross_validate_oracle.R points.csv r_oracle.csv

args <- commandArgs(trailingOnly = TRUE)
if (length(args) < 2L) stop("usage: cross_validate_oracle.R <points.csv> <out.csv>")

suppressMessages(library(Rmpfr))

pts <- read.csv(args[1], colClasses = "character")

# Bits needed so that 1 - Q retains full information about a tiny P.
# log2(P) ~ a * log2(z) for small z, so scale the working precision by that.
bits_for <- function(a, z_exp10) {
  approx_log2P <- abs(as.numeric(a) * z_exp10 * log2(10))
  max(256, ceiling(4 * approx_log2P) + 256)
}

evaluate <- function(a_dec, z_dec, surface, bits) {
  a <- mpfr(a_dec, bits)
  z <- mpfr(z_dec, bits)
  Q <- igamma(a, z) / gamma(a)          # MPFR gives the upper tail
  if (surface == "cumulative") return(1 - Q)
  if (surface == "survival")   return(Q)
  # density of the standardized variable: z^(a-1) e^-z / Gamma(a)
  exp((a - 1) * log(z) - z - lgamma(a))
}

out <- data.frame()
for (i in seq_len(nrow(pts))) {
  a_dec <- pts$a_dec[i]; z_dec <- pts$z_dec[i]; surf <- pts$surface[i]
  z_exp10 <- log10(abs(as.numeric(z_dec)))
  if (!is.finite(z_exp10)) z_exp10 <- -320          # subnormal decimal input
  b <- bits_for(a_dec, z_exp10)

  v1 <- evaluate(a_dec, z_dec, surf, b)
  v2 <- evaluate(a_dec, z_dec, surf, 2L * b)

  # Stability: the low-precision answer must reproduce the high-precision one.
  stable <- if (v2 == 0) as.character(v1 == 0) else
            as.character(abs(v1 - v2) / abs(v2) < mpfr(10, 64)^-40)

  out <- rbind(out, data.frame(
    a_dec = a_dec, z_dec = z_dec, surface = surf,
    value = format(v2, digits = 40),
    bits = 2L * b, stable = stable, stringsAsFactors = FALSE))
}

write.csv(out, args[2], row.names = FALSE, quote = FALSE)
cat(sprintf("%s: %d points, Rmpfr %s, MPFR %s\n", args[2], nrow(out),
            as.character(packageVersion("Rmpfr")), as.character(mpfrVersion())))
