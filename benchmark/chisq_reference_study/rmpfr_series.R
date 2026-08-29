#!/usr/bin/env Rscript
# =====================================================================
# Chi-square reference cross-check in PURE mpfr ARITHMETIC.
#
# WHY THIS EXISTS
#
# MPFR's own incomplete gamma (mpfr_gamma_inc, exposed as Rmpfr::igamma)
# aborts the process at shape a >= ~4.5E7:
#
#     a = 4.0E7  ok
#     a = 4.5E7  gamma_inc.c:289: MPFR assertion failed
#
# df = 1E8 needs a = 5E7, so rmpfr_crosscheck.R cannot reach it. mpmath's
# gammainc also fails at that shape (NoConvergence). Both third-party
# arbitrary-precision incomplete-gamma implementations available to this
# project are therefore unusable at df = 1E8.
#
# This script evaluates the incomplete gamma from first principles using
# only mpfr ARITHMETIC - add, multiply, divide, exp, log, lgamma - and
# never calls igamma. It restores a second implementation at df = 1E8.
#
# WHAT THIS IS, AND IS NOT
#
#   IS:     independent of MPFR's incomplete-gamma implementation, of its
#           internal strategy switching, and of mpmath's arithmetic
#           backend, rounding and evaluation order. A different library
#           computing every operation.
#
#   IS NOT: algorithmically independent. It uses the same converging
#           lower-series / upper-Lentz-CF route as the Python primary.
#           A shared *algorithmic* error would not be caught here.
#
# Record it as IMPLEMENTATION-INDEPENDENT, never as a fully independent
# confirmation. The genuinely algorithm-independent leg at df = 1E8
# remains the Python quadrature route in chisq_crosscheck.py.
#
# USAGE
#   Rscript rmpfr_series.R                          # timing probe only
#   Rscript rmpfr_series.R run                      # all df = 1E8 points
#   Rscript rmpfr_series.R run 1 6                  # points 1..6 only
#
# The probe runs FIRST and by default, because the series needs ~170,000
# terms per evaluation at a = 5E7 and the cost of an R-level mpfr loop
# cannot be assumed. Run the probe, read the projected total, then decide
# whether to run it whole or in ranges.
# =====================================================================

suppressPackageStartupMessages({
  library(Rmpfr)
  library(jsonlite)
})

args <- commandArgs(trailingOnly = TRUE)
MODE <- if (length(args) >= 1) args[1] else "probe"
FROM <- if (length(args) >= 2) as.integer(args[2]) else 1L
TO   <- if (length(args) >= 3) as.integer(args[3]) else NA_integer_

IN_PATH  <- "chisq_reference.json"
OUT_PATH <- "chisq_rmpfr_series_1e8.json"
DF_TARGET <- 1e8

PREC_LOW  <- 200L
PREC_HIGH <- 400L
REQUIRED_DIGITS <- 40
MAX_SERIES <- 4000000L
MAX_CF     <- 4000000L
MAX_NEWTON <- 100L

# ---------------------------------------------------------------------
# Exact binary64 from a hex float literal. Parsed by hand so the input
# never makes a decimal round-trip.
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
  val <- mpfr(strtoi(mp[1], 16L), prec)
  frac <- if (length(mp) > 1) mp[2] else ""
  if (nchar(frac) > 0) {
    for (i in seq_len(nchar(frac))) {
      d <- strtoi(substr(frac, i, i), 16L)
      val <- val + mpfr(d, prec) * mpfr(2, prec)^(-4L * i)
    }
  }
  val <- val * mpfr(2, prec)^expo
  if (neg) -val else val
}

# ---------------------------------------------------------------------
# log( x^a e^-x / Gamma(a) ), in log space to avoid overflow.
# ---------------------------------------------------------------------
log_prefactor <- function(a, x) a * log(x) - x - lgamma(a)

# Regularized P(a,x) by the converging lower series. Used for x < a+1.
lower_series <- function(a, x, prec) {
  one <- mpfr(1, prec)
  tol <- mpfr(2, prec)^(-(prec + 20L))
  term <- one
  total <- one
  n <- 0L
  repeat {
    n <- n + 1L
    if (n > MAX_SERIES) stop("lower series did not converge")
    term <- term * x / (a + n)
    total <- total + term
    if (abs(term) < abs(total) * tol) break
  }
  list(value = exp(log_prefactor(a, x) - log(a)) * total, iters = n)
}

# Regularized Q(a,x) by the modified Lentz continued fraction. x >= a+1.
upper_cf <- function(a, x, prec) {
  tiny <- mpfr(2, prec)^(-(2L * prec + 40L))
  tol  <- mpfr(2, prec)^(-(prec + 20L))
  one  <- mpfr(1, prec)

  b <- x + one - a
  cc <- one / tiny
  d <- if (b != 0) one / b else one / tiny
  h <- d
  i <- 0L
  repeat {
    i <- i + 1L
    if (i > MAX_CF) stop("upper CF did not converge")
    an <- -mpfr(i, prec) * (mpfr(i, prec) - a)
    b <- b + 2L
    d <- an * d + b
    if (abs(d) < tiny) d <- tiny
    cc <- b + an / cc
    if (abs(cc) < tiny) cc <- tiny
    d <- one / d
    delta <- d * cc
    h <- h * delta
    if (abs(delta - one) < tol) break
  }
  list(value = exp(log_prefactor(a, x)) * h, iters = i)
}

# Return the DIRECTLY computed side, never a complement.
reg_area <- function(a, x, side, prec) {
  if (x < a + 1) {
    r <- lower_series(a, x, prec)
    if (identical(side, "P")) return(r)
    return(list(value = mpfr(1, prec) - r$value, iters = r$iters))
  }
  r <- upper_cf(a, x, prec)
  if (identical(side, "Q")) return(r)
  list(value = mpfr(1, prec) - r$value, iters = r$iters)
}

log_pdf_chi2 <- function(df, q, prec) {
  k <- df / mpfr(2, prec)
  (k - 1) * log(q) - q / mpfr(2, prec) - lgamma(k) - k * log(mpfr(2, prec))
}

wh_seed <- function(df_dbl, p_dbl, prec) {
  z <- qnorm(p_dbl)
  if (!is.finite(z)) z <- sign(p_dbl - 0.5) * 38
  t <- 2 / (9 * df_dbl)
  base <- 1 - t + z * sqrt(t)
  if (base <= 0) base <- 1 / (9 * df_dbl)
  mpfr(df_dbl, prec) * mpfr(base, prec)^3
}

# Newton on log q. `seed` allows the high-precision pass to start from the
# low-precision answer. That is an EFFICIENCY choice within one route, not
# a borrowing of the Python result - it never sees the primary's value.
invert_series <- function(df_dbl, p_hex, prec, seed = NULL) {
  p <- hex_to_mpfr(p_hex, prec)
  a <- mpfr(df_dbl, prec) / mpfr(2, prec)
  half <- mpfr(1, prec) / mpfr(2, prec)
  upper <- p > half
  side <- if (upper) "Q" else "P"
  target <- if (upper) (mpfr(1, prec) - p) else p

  q <- if (is.null(seed)) wh_seed(df_dbl, as.numeric(format(p, digits = 17)), prec)
       else mpfr(seed, prec)
  tol <- mpfr(2, prec)^(-(prec - 30L))
  total_iters <- 0L

  for (step in seq_len(MAX_NEWTON)) {
    r <- reg_area(a, q / mpfr(2, prec), side, prec)
    total_iters <- total_iters + r$iters
    resid <- r$value - target
    if (resid == 0) {
      return(list(q = q, steps = step, resid = abs(resid),
                  side = side, iters = total_iters, converged = TRUE))
    }
    dens <- exp(log_pdf_chi2(mpfr(df_dbl, prec), q, prec)) * q
    d_du <- if (upper) -dens else dens
    q_new <- exp(log(q) - resid / d_du)
    if (q_new <= 0) {
      return(list(q = q, steps = step, resid = NA, side = side,
                  iters = total_iters, converged = FALSE))
    }
    rel <- abs(q_new - q) / abs(q_new)
    q <- q_new
    if (rel < tol) {
      r <- reg_area(a, q / mpfr(2, prec), side, prec)
      return(list(q = q, steps = step, resid = abs(r$value - target),
                  side = side, iters = total_iters, converged = TRUE))
    }
  }
  list(q = q, steps = MAX_NEWTON, resid = NA, side = side,
       iters = total_iters, converged = FALSE)
}

agree_digits <- function(x, y) {
  if (x == y) return(Inf)
  as.numeric(-log10(abs(x - y) / abs(x)))
}

# ---------------------------------------------------------------------
# Timing probe. The R-level mpfr loop cost cannot be assumed, so measure
# one central evaluation and project before committing to a long run.
# ---------------------------------------------------------------------
cat(sprintf("Rmpfr %s, MPFR %s\n", as.character(packageVersion("Rmpfr")),
            as.character(mpfrVersion())))
cat("Pure mpfr arithmetic - igamma is NEVER called.\n\n")

a_probe <- mpfr(DF_TARGET, PREC_LOW) / 2
cat(sprintf("probe: lower series at a = %.1e, %d bits ... ",
            DF_TARGET / 2, PREC_LOW)); flush.console()
t0 <- Sys.time()
pr <- lower_series(a_probe, a_probe, PREC_LOW)
dt_low <- as.numeric(difftime(Sys.time(), t0, units = "secs"))
cat(sprintf("%d terms, %.1fs\n", pr$iters, dt_low)); flush.console()

a_probe2 <- mpfr(DF_TARGET, PREC_HIGH) / 2
cat(sprintf("probe: lower series at a = %.1e, %d bits ... ",
            DF_TARGET / 2, PREC_HIGH)); flush.console()
t0 <- Sys.time()
pr2 <- lower_series(a_probe2, a_probe2, PREC_HIGH)
dt_high <- as.numeric(difftime(Sys.time(), t0, units = "secs"))
cat(sprintf("%d terms, %.1fs\n\n", pr2$iters, dt_high)); flush.console()

# ~12 of 23 points use the series; the rest use the CF, which is ~600
# iterations and negligible by comparison. Roughly 6 Newton steps at low
# precision and 2 at high (seeded).
est <- 12 * (6 * dt_low + 2 * dt_high)
cat(sprintf("projected df=1E8 block runtime: ~%.0f min\n", est / 60))
cat("  (series-bound points only; CF points are comparatively free)\n\n")

if (!identical(MODE, "run")) {
  cat("Probe only. Re-run with:  Rscript rmpfr_series.R run\n")
  cat("Or in ranges, e.g.:       Rscript rmpfr_series.R run 1 6\n")
  quit(status = 0L)
}

# ---------------------------------------------------------------------
ref <- fromJSON(IN_PATH, simplifyDataFrame = FALSE)
records <- Filter(function(r) isTRUE(all.equal(r$df, DF_TARGET)),
                  ref$references)
if (is.na(TO)) TO <- length(records)
sel <- seq(FROM, min(TO, length(records)))
cat(sprintf("df=1E8: %d points total, running %d..%d\n\n",
            length(records), FROM, max(sel)))

out <- list()
worst <- Inf
worst_at <- NA_character_
n_rej <- 0L
t_start <- Sys.time()

for (k in seq_along(sel)) {
  i <- sel[k]
  r <- records[[i]]
  p_hex <- r$probability$hex

  lo <- invert_series(DF_TARGET, p_hex, PREC_LOW)
  hi <- if (lo$converged)
          invert_series(DF_TARGET, p_hex, PREC_HIGH, seed = lo$q)
        else lo

  if (!lo$converged || !hi$converged) {
    status <- "REJECTED"; reason <- "Newton did not converge"; stab <- NA_real_
  } else {
    stab <- agree_digits(mpfr(lo$q, PREC_HIGH), hi$q)
    if (is.finite(stab) && stab < REQUIRED_DIGITS) {
      status <- "REJECTED"
      reason <- sprintf("did not stabilise: %.4g digits, required %d",
                        stab, REQUIRED_DIGITS)
    } else { status <- "ACCEPTED"; reason <- NA_character_ }
  }
  if (identical(status, "REJECTED")) n_rej <- n_rej + 1L

  agree <- NA_real_
  if (identical(status, "ACCEPTED") && !is.null(r$quantile)) {
    agree <- agree_digits(mpfr(r$quantile, PREC_HIGH), hi$q)
    if (is.finite(agree) && agree < worst) {
      worst <- agree
      worst_at <- sprintf("df=1E8 %s/%s p=%s", r$arm, r$band, p_hex)
    }
  }

  p_ex <- hex_to_mpfr(p_hex, PREC_HIGH)
  denom <- min(p_ex, mpfr(1, PREC_HIGH) - p_ex)

  out[[length(out) + 1L]] <- list(
    index = i, df = DF_TARGET, arm = r$arm, band = r$band,
    probability_hex = p_hex, status = status, reason = reason,
    precision_pair_bits = c(PREC_LOW, PREC_HIGH),
    stabilisation_digits = if (is.finite(stab)) round(stab, 4) else NULL,
    quantile = format(hi$q, digits = 50),
    residual_side = hi$side,
    tail_residual = if (is.na(hi$resid)) NA_character_ else
                      format(hi$resid / denom, digits = 6),
    newton_steps = hi$steps,
    route_iterations = hi$iters,
    agreement_with_python_primary_digits =
      if (is.finite(agree)) round(agree, 4) else NULL
  )

  cat(sprintf("  [%2d/%2d] idx=%2d %-8s %-8s %s agree=%s\n",
              k, length(sel), i, r$arm, r$band, status,
              if (is.finite(agree)) sprintf("%.3f", agree) else "NA"))
  flush.console()

  write(toJSON(list(
    checkpoint = "v1.0.0 plan Track A2 item 6 - df=1E8 mpfr-arithmetic cross-check",
    independence = paste("IMPLEMENTATION-INDEPENDENT ONLY - same series/CF",
                         "algorithm as the Python primary, different library,",
                         "arithmetic backend and rounding. NOT an",
                         "algorithm-independent confirmation."),
    igamma_used = FALSE,
    reason_igamma_unusable = paste("MPFR mpfr_gamma_inc aborts at a >= ~4.5E7",
                                   "(ok at 4.0E7); df=1E8 needs a = 5E7"),
    status = "PARTIAL - run in progress or terminated early",
    range = c(FROM, max(sel)),
    points_written = length(out),
    points = out
  ), auto_unbox = TRUE, pretty = TRUE, digits = NA), OUT_PATH)
}

runtime <- as.numeric(difftime(Sys.time(), t_start, units = "secs"))
cat(sprintf("\n%d/%d accepted, %d rejected, %.1f min\n",
            length(sel) - n_rej, length(sel), n_rej, runtime / 60))
if (is.finite(worst)) {
  cat(sprintf("minimum primary-vs-mpfr-series agreement: %.4f digits\n", worst))
  cat(sprintf("  at %s\n", worst_at))
}
cat(sprintf("written to %s\n", OUT_PATH))
