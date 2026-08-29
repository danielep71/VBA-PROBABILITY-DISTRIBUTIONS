#!/usr/bin/env Rscript
# =====================================================================
# Rmpfr cross-check of the frozen 69-point Chi-square reference set.
#
# This is the REQUIRED independent leg of the #22 feasibility checkpoint
# (v1.0.0 plan, Track A2 item 6). Until it runs and is verified, the
# checkpoint stays PROVISIONAL and no final agreement figure is recorded.
#
# IDENTICAL INPUTS. Probabilities are read from chisq_reference.json as
# hex float literals, so the exact IEEE-754 binary64 value crosses the
# language boundary with no decimal round-trip. 0.9 and 0.99 in particular
# are NOT re-parsed from decimal text.
#
# IDENTICAL MEASURES. Two, matching the Python side exactly:
#   quantile  - the Chi-square inverse CDF at the exact binary64 p
#   residual  - |area_direct - target| / min(p, 1-p)
# The small side is computed DIRECTLY (lower area for p <= 1/2, upper area
# for p > 1/2), so no complement subtraction enters either measure. This
# mirrors the Python route exactly.
#
# METHOD. MPFR's own incomplete gamma (Rmpfr::igamma, upper) at high
# precision, with the quantile re-derived independently by Newton from a
# Wilson-Hilferty seed. igamma is a different implementation in a different
# library from the Python series/CF route.
#
# USAGE
#   Rscript rmpfr_crosscheck.R [chisq_reference.json] [chisq_rmpfr.json]
#
# REQUIREMENTS
#   install.packages(c("Rmpfr", "jsonlite"))
#   Rmpfr >= 0.8-0 (needs MPFR >= 3.2 for mpfr_gamma_inc).
#
# FAIL-CLOSED. Any point that does not converge, or that stabilises at
# fewer than REQUIRED_DIGITS across the two precisions, is written with
# status REJECTED. A single rejection fails the checkpoint. Nothing is
# silently dropped and no point is skipped.
# =====================================================================

suppressPackageStartupMessages({
  library(Rmpfr)
  library(jsonlite)
})

args <- commandArgs(trailingOnly = TRUE)
in_path  <- if (length(args) >= 1) args[1] else "chisq_reference.json"
out_path <- if (length(args) >= 2) args[2] else "chisq_rmpfr.json"

# Precision pair, in BITS. Materially separated, mirroring the Python
# 60/120 decimal-digit pair (~200 / ~400 bits).
PREC_LOW  <- 200L
PREC_HIGH <- 400L
REQUIRED_DIGITS <- 40
MAX_NEWTON <- 100L

# ---------------------------------------------------------------------
# Exact binary64 from a hex float literal, e.g. "0x1.ccccccccccccdp-1".
# Parsed by hand so no decimal text ever represents the input.
# ---------------------------------------------------------------------
hex_to_mpfr <- function(hexstr, prec) {
  s <- hexstr
  neg <- substr(s, 1, 1) == "-"
  if (neg) s <- substring(s, 2)
  s <- sub("^0[xX]", "", s)
  parts <- strsplit(s, "[pP]")[[1]]
  mant <- parts[1]
  expo <- as.integer(parts[2])
  mp <- strsplit(mant, ".", fixed = TRUE)[[1]]
  int_part <- mp[1]
  frac_part <- if (length(mp) > 1) mp[2] else ""
  val <- mpfr(strtoi(int_part, 16L), prec)
  if (nchar(frac_part) > 0) {
    for (i in seq_len(nchar(frac_part))) {
      d <- strtoi(substr(frac_part, i, i), 16L)
      val <- val + mpfr(d, prec) * mpfr(2, prec)^(-4L * i)
    }
  }
  val <- val * mpfr(2, prec)^expo
  if (neg) val <- -val
  val
}

# ---------------------------------------------------------------------
# Regularized areas via MPFR's incomplete gamma.
#   Rmpfr::igamma(a, x) is the UPPER incomplete gamma Gamma(a, x),
#   so Q(a,x) = igamma(a,x) / gamma(a) and P = 1 - Q.
# The direct side is returned without forming the complement.
# ---------------------------------------------------------------------
reg_upper <- function(a, x) igamma(a, x) / gamma(a)

reg_area <- function(a, x, side, prec) {
  Q <- reg_upper(a, x)
  if (identical(side, "Q")) Q else (mpfr(1, prec) - Q)
}

log_pdf_chi2 <- function(df, q, prec) {
  k <- df / mpfr(2, prec)
  (k - 1) * log(q) - q / mpfr(2, prec) - lgamma(k) - k * log(mpfr(2, prec))
}

# Wilson-Hilferty seed. qnorm in double is ample for a seed; Newton
# supplies every digit that is actually reported.
wh_seed <- function(df_dbl, p_dbl, prec) {
  z <- qnorm(p_dbl)
  if (!is.finite(z)) z <- sign(p_dbl - 0.5) * 38
  t <- 2 / (9 * df_dbl)
  base <- 1 - t + z * sqrt(t)
  if (base <= 0) base <- 1 / (9 * df_dbl)
  mpfr(df_dbl, prec) * mpfr(base, prec)^3
}

invert_rmpfr <- function(df_dbl, p_hex, prec) {
  p <- hex_to_mpfr(p_hex, prec)
  a <- mpfr(df_dbl, prec) / mpfr(2, prec)
  half <- mpfr(0.5, prec)
  upper <- p > half
  side <- if (upper) "Q" else "P"
  target <- if (upper) (mpfr(1, prec) - p) else p

  q <- wh_seed(df_dbl, as.numeric(format(p, digits = 17)), prec)
  tol <- mpfr(10, prec)^(-(floor(prec * 0.301) - 8))

  for (step in seq_len(MAX_NEWTON)) {
    area <- reg_area(a, q / mpfr(2, prec), side, prec)
    resid <- area - target
    if (resid == 0) {
      return(list(q = q, steps = step, resid = abs(resid),
                  side = side, converged = TRUE))
    }
    dens <- exp(log_pdf_chi2(mpfr(df_dbl, prec), q, prec)) * q
    d_du <- if (upper) -dens else dens
    q_new <- exp(log(q) - resid / d_du)
    if (!is.finite(as.numeric(q_new)) || q_new <= 0) {
      return(list(q = q, steps = step, resid = NA, side = side,
                  converged = FALSE))
    }
    rel <- abs(q_new - q) / abs(q_new)
    q <- q_new
    if (rel < tol) {
      area <- reg_area(a, q / mpfr(2, prec), side, prec)
      return(list(q = q, steps = step, resid = abs(area - target),
                  side = side, converged = TRUE))
    }
  }
  list(q = q, steps = MAX_NEWTON, resid = NA, side = side, converged = FALSE)
}

agree_digits <- function(x, y, prec) {
  if (x == y) return(Inf)
  as.numeric(-log10(abs(x - y) / abs(x)))
}

# ---------------------------------------------------------------------
ref <- fromJSON(in_path, simplifyDataFrame = FALSE)
records <- ref$references
cat(sprintf("read %d reference points from %s\n", length(records), in_path))

t_start <- Sys.time()
out <- vector("list", length(records))
worst <- Inf
worst_at <- NA_character_
n_rejected <- 0L

for (i in seq_along(records)) {
  r <- records[[i]]
  p_hex <- r$probability$hex
  df_dbl <- r$df

  lo <- invert_rmpfr(df_dbl, p_hex, PREC_LOW)
  hi <- invert_rmpfr(df_dbl, p_hex, PREC_HIGH)

  if (!lo$converged || !hi$converged) {
    status <- "REJECTED"
    reason <- "Newton did not converge at one or both precisions"
    stab <- NA_real_
  } else {
    stab <- agree_digits(mpfr(lo$q, PREC_HIGH), hi$q, PREC_HIGH)
    if (is.finite(stab) && stab < REQUIRED_DIGITS) {
      status <- "REJECTED"
      reason <- sprintf("did not stabilise: %.4g digits, required %d",
                        stab, REQUIRED_DIGITS)
    } else {
      status <- "ACCEPTED"
      reason <- NA_character_
    }
  }
  if (status == "REJECTED") n_rejected <- n_rejected + 1L

  # Agreement against the Python primary, the checkpoint's headline measure.
  agree <- NA_real_
  if (status == "ACCEPTED" && !is.null(r$quantile)) {
    q_py <- mpfr(r$quantile, PREC_HIGH)
    agree <- agree_digits(q_py, hi$q, PREC_HIGH)
    if (is.finite(agree) && agree < worst) {
      worst <- agree
      worst_at <- sprintf("df=%.0e %s/%s p=%s", df_dbl, r$arm, r$band, p_hex)
    }
  }

  p_exact <- hex_to_mpfr(p_hex, PREC_HIGH)
  denom <- min(p_exact, mpfr(1, PREC_HIGH) - p_exact)
  tail_resid <- if (is.na(hi$resid)) NA_character_ else
    format(hi$resid / denom, digits = 6)

  out[[i]] <- list(
    df = df_dbl, arm = r$arm, band = r$band,
    probability_hex = p_hex,
    status = status, reason = reason,
    precision_pair_bits = c(PREC_LOW, PREC_HIGH),
    stabilisation_digits = if (is.finite(stab)) round(stab, 4) else NULL,
    quantile = format(hi$q, digits = 50),
    residual_side = hi$side,
    tail_residual = tail_resid,
    newton_steps = hi$steps,
    agreement_with_python_primary_digits =
      if (is.finite(agree)) round(agree, 4) else NULL
  )
  cat(sprintf("  [%2d/%d] df=%.0e %-8s %-8s %s agree=%s\n", i, length(records),
              df_dbl, r$arm, r$band, status,
              if (is.finite(agree)) sprintf("%.3f", agree) else "NA"))
}

runtime <- as.numeric(difftime(Sys.time(), t_start, units = "secs"))

payload <- list(
  checkpoint = "v1.0.0 plan Track A2 item 6 - Chi-square Rmpfr cross-check",
  status = if (n_rejected == 0L) "RMPFR LEG COMPLETE - AWAITING MAINTAINER VERIFICATION"
           else "RMPFR LEG FAILED",
  generated_utc = format(Sys.time(), "%Y-%m-%dT%H:%M:%SZ", tz = "UTC"),
  runtime_s_total = round(runtime, 3),
  versions = list(
    R = paste(R.version$major, R.version$minor, sep = "."),
    Rmpfr = as.character(packageVersion("Rmpfr")),
    mpfr = tryCatch(as.character(Rmpfr::mpfrVersion()), error = function(e) "unknown"),
    gmp = tryCatch(as.character(packageVersion("gmp")), error = function(e) "unknown"),
    platform = R.version$platform
  ),
  method = list(
    incomplete_gamma = "Rmpfr::igamma (MPFR mpfr_gamma_inc), upper",
    inversion = "Newton on log q from a Wilson-Hilferty seed",
    direct_side = "lower area for p <= 1/2, upper area for p > 1/2",
    measures = c("quantile", "tail_residual = |area - target| / min(p, 1-p)"),
    precision_pair_bits = c(PREC_LOW, PREC_HIGH),
    required_agreement_digits = REQUIRED_DIGITS
  ),
  convergence = list(
    points_total = length(records),
    accepted = length(records) - n_rejected,
    rejected = n_rejected,
    all_converged_and_stabilised = (n_rejected == 0L)
  ),
  agreement = list(
    measure = "primary (Python mpmath series/CF) vs Rmpfr, significant digits of the quantile",
    minimum_agreement_digits = if (is.finite(worst)) round(worst, 4) else NULL,
    minimum_agreement_at = worst_at
  ),
  points = out
)

write(toJSON(payload, auto_unbox = TRUE, pretty = TRUE, digits = NA), out_path)

cat(sprintf("\n%d/%d accepted, %d rejected\n",
            length(records) - n_rejected, length(records), n_rejected))
if (is.finite(worst)) {
  cat(sprintf("minimum primary-vs-Rmpfr agreement: %.4f significant digits\n", worst))
  cat(sprintf("  at %s\n", worst_at))
}
cat(sprintf("written to %s\n", out_path))
if (n_rejected > 0L) {
  cat("CHECKPOINT FAILS: at least one point did not converge or stabilise.\n")
  quit(status = 1L)
}
