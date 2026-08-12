Attribute VB_Name = "M_STATS_PROBDIST_SPECIALFUNCS"
Option Explicit
Option Private Module

'==============================================================================
' M_STATS_PROBDIST_SPECIALFUNCS
'------------------------------------------------------------------------------
' PURPOSE
'   Provides the log-gamma, log-beta, regularized incomplete beta and
'   regularized incomplete gamma functions, together with their inverses, as
'   distribution-agnostic kernels with an explicit success/failure contract.
'
' WHY THIS EXISTS
'   Every distribution outside the normal family reduces to one of two objects:
'   the regularized incomplete beta function I_x(a, b) and the regularized
'   incomplete gamma functions P(a, x) and Q(a, x). Student t, chi-square, F,
'   gamma, beta, Poisson, binomial and the negative binomial are all thin
'   wrappers around them. Isolating those objects here means the hard numerics
'   are written, tested and fixed exactly once.
'
'   Every iterative routine here returns Boolean and reports non-convergence
'   through FailMsg. A continued fraction that hits its iteration cap and then
'   returns its partial sum as though nothing happened is the single most
'   dangerous thing a numerical library can do, because the caller has no way to
'   tell a converged answer from a wrong one.
'
' PUBLIC (PROJECT-SCOPED) SURFACE
'   Logarithmic gamma:
'     - PROB_LogGamma
'     - PROB_LogGammaDelta
'     - PROB_LogGammaHalfDiff
'     - PROB_LogBeta
'
'   Combinatorics:
'     - PROB_StirlingError
'     - PROB_LogChoose
'
'   Incomplete beta:
'     - PROB_TryBetaRegularized
'     - PROB_TryBetaContinuedFraction
'     - PROB_TryBetaInvRegularized
'
'   Incomplete gamma:
'     - PROB_TryGammaRegularizedP
'     - PROB_TryGammaRegularizedQ
'     - PROB_TryGammaSeriesP
'     - PROB_TryGammaContinuedFractionQ
'     - PROB_TryGammaInvP
'
'   Stable log-density kernels (CR-P1-01 / CR-P1-02):
'     - PROB_TryDeviancePart      Loader deviance bd0(x, m)
'     - PROB_TryGammaLogPdf       cancellation-free Gamma log-density
'     - PROB_TryBetaLogPdf        cancellation-free Beta log-density
'
'     These exist because the literal log-density and incomplete-function
'     prefactor forms subtract two quantities of size shape*Log(shape), which
'     loses the result entirely at large shape. Four public densities and both
'     incomplete functions route through them, so a fix lands once rather than
'     six times. Prefer the two log-pdf helpers over calling the deviance
'     directly: they encapsulate the decomposition their callers need.
'
' ALGORITHM PROVENANCE
'   - PROB_StirlingError / PROB_LogChoose:
'       Catherine Loader, "Fast and Accurate Computation of Binomial
'       Probabilities" (2000). The Stirling error delta(N) is O(1/(12N)) and
'       so is computed to full relative accuracy at every N; assembling
'       Log C(N,K) from three small deltas avoids subtracting two log-gammas
'       of size N*Log(N). Public; the arrangement used by R's dbinom.
'   - PROB_LogGamma:
'       Lanczos approximation, g = 7, n = 9, with the reflection formula for
'       z < 0.5. Measured relative error against 50-digit arithmetic is below
'       6.1E-14 for z in [1E-8, 1E+50]. Public, published; not proprietary.
'   - PROB_LogGammaHalfDiff:
'       Asymptotic expansion of Log(Gamma(z + 1/2)) - Log(Gamma(z)) for z >= 20,
'       direct difference below. The direct difference alone cancels: at z = 5E+5
'       it carries a relative error of 5.9E-11, because two numbers of size 6E+6
'       are being subtracted to produce a number of size 6. Standard result.
'   - PROB_TryBetaContinuedFraction:
'       Continued fraction evaluated by the modified Lentz method, in the
'       arrangement of Numerical Recipes (betacf).
'   - PROB_TryGammaSeriesP / PROB_TryGammaContinuedFractionQ:
'       Series expansion for x < a + 1, continued fraction for the upper tail,
'       in the arrangement of Numerical Recipes (gser / gcf).
'   - PROB_TryBetaInvRegularized:
'       Seed from the Carter / AS 109 normal approximation for a, b > 1, from
'       the leading series term otherwise; refined by Newton's method safeguarded
'       by bisection.
'   - PROB_TryGammaInvP:
'       Wilson-Hilferty seed refined by Newton's method safeguarded by bisection.
'   Nothing here is a newly-invented algorithm; the Try contract, the (X, Y)
'   argument pair and the iteration budgets are the local contribution.
'
' DESIGN PRINCIPLES
'   - Every iterative routine returns Boolean. False means the answer is unknown,
'     not approximately known. Result is left unchanged on failure.
'   - The incomplete beta takes BOTH X and Y = 1 - X. The caller supplies each
'     from a form that does not cancel, and the routine never re-derives one from
'     the other by subtraction. This is what makes Student t exact near zero and
'     what lets the F quantile reach 1E+34.
'   - The inverses always solve on whichever of the two tails is the smaller, so
'     that the quantity being driven to a target retains full relative precision.
'   - Kernels do not validate their callers' domains and do not write Status.
'     Each states its PRECONDITION and trusts it.
'
' NOTES
'   - Iteration budgets are generous because the cost is paid only in the rare
'     large-parameter case. The two kernels scale quite differently, which is why
'     one shared budget is not the same constraint for both:
'
'       incomplete gamma series   about 7 * Sqr(df) terms at its worst point
'                                 (x = a), drifting slowly downward with df
'       incomplete beta CF        well under 1 * Sqr(df) iterations, and the
'                                 ratio FALLS with df (about 0.9 at 1E4,
'                                 0.2 at 1E8) - it is far cheaper than the series
'
'     Typical degrees of freedom below 100 converge in a few dozen iterations.
'     The gamma series is therefore the binding kernel: it exhausts a 100000
'     budget while the beta CF is still in the low thousands.
'
'     THE REACHABLE RANGES ARE NOT RESTATED HERE. They follow from the budgets
'     above and the Sqr(df) growth rule, they differ between the two kernels, and
'     they have moved as measurement improved - the beta range in particular after
'     the CR-P1-02 prefactor repair. They are measured in benchmark/cdf_large_shape
'     rather than asserted in prose, which is where the previous figures in this
'     note drifted out of step with the code.
'
'     These convergence ranges are narrower than the representational validation
'     bound (PROB_PARAMETER_MAGNITUDE_GUARD): a parameter between a kernel's
'     convergence range and that bound is accepted by validation, attempted, and
'     then returns a clean parameter-named non-convergence error rather than a
'     wrong answer. The boundary is not a hard cliff because it depends on the
'     companion arguments, so it is documented and measured rather than enforced.
'
'     Reachability is also distinct from ACCURACY: a kernel can converge and
'     still be outside its measured accuracy envelope. The distribution modules
'     enforce that separately through PROB_*_MAX_DF and PROB_DENSITY_SHAPE_MAX.
'   - PROB_LogGamma is recursive through its reflection branch, exactly once.
'
' UPDATED
'   2026-08-11 - Surface list completed with the stable log-density kernels;
'                LogGammaDelta accuracy figure re-measured and contracted.
'==============================================================================

'==============================================================================
' PRIVATE CONSTANTS
'==============================================================================

Private Const PROB_BETA_MAX_ITER       As Long = 100000   'Lentz iterations, incomplete beta
Private Const PROB_GAMMA_MAX_ITER      As Long = 100000   'Series / Lentz iterations, incomplete gamma
'The ascending gamma series stops at machine epsilon rather than the shared
'PROB_NUM_EPS (3E-14). Measured: the tighter criterion buys about 145x accuracy
'(A = 1E6: 4.1E-12 -> 2.9E-14) for about 11% more terms, and it does NOT move
'the reachability edge - at the chi-square cap df 1E8 the series needs 41095
'terms at 3E-14 and 45946 here, both far inside PROB_GAMMA_MAX_ITER. The two
'continued fractions gain nothing from the same change (the beta CF is limited
'elsewhere and is unchanged above A = 1E6), so they keep PROB_NUM_EPS.
Private Const PROB_GAMMA_SERIES_EPS    As Double = PROB_MACH_EPS 'Stop for the ascending gamma series only; see the note above
Private Const PROB_INV_MAX_ITER        As Long = 200      'Safeguarded Newton iterations
Private Const PROB_BD0_MAX_ITER         As Long = 1000     'Loader deviance series iteration guard
Private Const PROB_BLP_TINY_PRODUCT As Double = 1E-300 'Below this the N*X / N*Y deviance argument is formed in log space to preserve the caller''s stable logs
Private Const PROB_IBETA_LOADER_MIN_SHAPE As Double = 1000# 'A + B at or above which the incomplete-beta factor uses the Loader decomposition instead of the literal form (measured crossover; see the header of PROB_TryBetaRegularized)
Public Const PROB_DENSITY_SHAPE_MAX As Double = 1E+20 'Validated large-shape density envelope (Gamma/Chi-square/Beta/F); measured, benchmark/density_large_shape
Private Const PROB_HALF_DIFF_CUTOFF    As Double = 20#    'Z at or above which the asymptotic half-difference wins
Private Const PROB_LOGBETA_STABLE_RATIO As Double = 0.1     'Small/Large below this uses the stable LogGamma difference (validated by the committed seam study and independent holdout)
Private Const PROB_BETAINV_ROUNDTRIP_TOL As Double = 0.000001    'Forward-probability residual above which no representable interior quantile exists (worst legitimate measured residual is 3.9E-10)

'Lanczos g = 7, n = 9 series coefficients. SINGLE SOURCE OF TRUTH shared by
'PROB_LogGamma and PROB_LogGammaDelta, which must evaluate the identical series.
Private Const PROB_LANCZOS_G  As Double = 7#
Private Const PROB_LANCZOS_P0 As Double = 0.99999999999981
Private Const PROB_LANCZOS_P1 As Double = 676.520368121885
Private Const PROB_LANCZOS_P2 As Double = -1259.1392167224
Private Const PROB_LANCZOS_P3 As Double = 771.323428777653
Private Const PROB_LANCZOS_P4 As Double = -176.615029162141
Private Const PROB_LANCZOS_P5 As Double = 12.5073432786869
Private Const PROB_LANCZOS_P6 As Double = -0.13857109526572
Private Const PROB_LANCZOS_P7 As Double = 9.98436957801957E-06
Private Const PROB_LANCZOS_P8 As Double = 1.50563273514931E-07

'Maclaurin coefficients of Log(Gamma(1 + X)) about X = 0:
'
'    Log(Gamma(1 + X)) = -EulerGamma * X
'                        + Sum(k = 2..) (-1)^k * Zeta(k) * X^k / k
'
'C1 is -EulerGamma; Ck is (-1)^k * Zeta(k) / k. Fixed table, never evaluated
'at run time. Twenty-six terms hold the series at the coefficient rounding
'floor throughout the retained interval; see PROB_TryLogGamma1p.
Private Const PROB_LG1P_SERIES_MAX As Double = 0.25   'Measured series/Lanczos seam
'C1 is written as a split expression rather than a fifteen-digit literal. It
'dominates the scaled-error floor and is the one coefficient that survives
'division by Shape in the scaled Gamma inverse, so a single ulp here is worth
'more than it is anywhere else in the table. The sum below is the exactly
'rounded Double for -0.5772156649015329; the fifteen-digit form is one ulp away
'and leaves a 2.3E-16 floor. Same pattern as PROB_DS_MAX_EXACT_INTEGER.
Private Const PROB_LG1P_C1  As Double = _
    -0.577215664901532 - 8.88178419700125E-16
Private Const PROB_LG1P_C2  As Double = 0.822467033424113
Private Const PROB_LG1P_C3  As Double = -0.400685634386531
Private Const PROB_LG1P_C4  As Double = 0.270580808427785
Private Const PROB_LG1P_C5  As Double = -0.207385551028674
Private Const PROB_LG1P_C6  As Double = 0.169557176997408
Private Const PROB_LG1P_C7  As Double = -0.144049896768846
Private Const PROB_LG1P_C8  As Double = 0.125509669524743
Private Const PROB_LG1P_C9  As Double = -0.111334265869565
Private Const PROB_LG1P_C10 As Double = 0.100099457512782
Private Const PROB_LG1P_C11 As Double = -0.090954017145829
Private Const PROB_LG1P_C12 As Double = 0.083353840546109
Private Const PROB_LG1P_C13 As Double = -7.69325164113522E-02
Private Const PROB_LG1P_C14 As Double = 7.14329462953613E-02
Private Const PROB_LG1P_C15 As Double = -6.66687058824205E-02
Private Const PROB_LG1P_C16 As Double = 0.062500955141213
Private Const PROB_LG1P_C17 As Double = -5.88239786586846E-02
Private Const PROB_LG1P_C18 As Double = 5.55557676274036E-02
Private Const PROB_LG1P_C19 As Double = -5.26316793796167E-02
Private Const PROB_LG1P_C20 As Double = 5.00000476981017E-02
Private Const PROB_LG1P_C21 As Double = -4.76190703301422E-02
Private Const PROB_LG1P_C22 As Double = 4.54545562932047E-02
Private Const PROB_LG1P_C23 As Double = -4.34782660530403E-02
Private Const PROB_LG1P_C24 As Double = 4.16666691503412E-02
Private Const PROB_LG1P_C25 As Double = -4.00000011921401E-02
Private Const PROB_LG1P_C26 As Double = 3.84615390346752E-02


'==============================================================================
' LOGARITHMIC GAMMA
'==============================================================================

Public Function PROB_LogGamma( _
    ByVal Z As Double) _
    As Double
'
'==============================================================================
' PROB_LogGamma
'------------------------------------------------------------------------------
' PURPOSE
'   Returns Log(Gamma(Z)) using the Lanczos approximation.
'
' PRECONDITION
'   Z > 0. Z = 0 raises a division/log error; Z < 0 is not supported. Callers in
'   this project validate strictly positive parameters before arriving here.
'
' ACCURACY
'   Under revision. The previous claim, relative error below 6.1E-14 across Z in
'   [1E-8, 1E+50], is withdrawn as a global statement for two reasons measured in
'   benchmark/loggamma1p_study:
'
'     - Log(Gamma(Z)) is zero at Z = 1 and Z = 2, so a single global RELATIVE
'       contract is ill-conditioned by construction. Relative error reaches
'       9.3E-14 near Z = 1.75, which is only about 7.9E-15 of absolute error on
'       a value of magnitude 0.084.
'     - The reflection path was genuinely defective for subnormal Z, reaching
'       6.2E-05 relative at the smallest positive Double. That defect is
'       corrected by the small-positive branch below.
'
'   The replacement is a set of regime-aware contracts keyed on absolute error
'   in the logarithm, which is the quantity downstream callers actually
'   propagate: the relative error of Exp(v) is approximately the absolute error
'   of v. Thresholds are frozen from the Phase 1 main grid and holdout.
'==============================================================================
'
'------------------------------------------------------------------------------
' DECLARE CONSTANTS
'------------------------------------------------------------------------------

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Dim X                   As Double          'Lanczos series accumulator
    Dim T                   As Double          'Lanczos shifted variable
    Dim Zm1                 As Double          'Z - 1

'------------------------------------------------------------------------------
' SMALL POSITIVE ARGUMENT
'------------------------------------------------------------------------------
    'Use Gamma(Z) = Gamma(1 + Z) / Z. This keeps the first-order -EulerGamma * Z
    'term, which the reflection route below loses, and it never forms
    'Sin(PROB_PI * Z) at an argument that has entered the subnormal range.
    'Measured in benchmark/loggamma1p_study: at the smallest positive Double the
    'reflection route reaches 6.2E-05 relative, this route 5.9E-17. The series
    'helper is called directly rather than through PROB_TryLogGamma1p, because
    'the branch condition already establishes the helper's precondition and a
    'Double-returning kernel has no channel to report a failure that cannot
    'happen.
        If Z <= PROB_LG1P_SERIES_MAX Then
            PROB_LogGamma = PROB_LogGamma1pSeries(Z) - Log(Z)
            Exit Function
        End If

'------------------------------------------------------------------------------
' REFLECTION FORMULA
'------------------------------------------------------------------------------
    'Use the reflection formula on the remaining small-positive interval. The
    'measured crossover is the series seam: below it the branch above wins by an
    'order of magnitude, above it the reflection wins by about two.
        If Z < 0.5 Then
            PROB_LogGamma = _
                Log(PROB_PI) - Log(Sin(PROB_PI * Z)) - PROB_LogGamma(1# - Z)
            Exit Function
        End If

'------------------------------------------------------------------------------
' LANCZOS APPROXIMATION
'------------------------------------------------------------------------------
    'Shift Z
        Zm1 = Z - 1#

    'Compute the Lanczos series
        X = PROB_LANCZOS_P0
        X = X + PROB_LANCZOS_P1 / (Zm1 + 1#)
        X = X + PROB_LANCZOS_P2 / (Zm1 + 2#)
        X = X + PROB_LANCZOS_P3 / (Zm1 + 3#)
        X = X + PROB_LANCZOS_P4 / (Zm1 + 4#)
        X = X + PROB_LANCZOS_P5 / (Zm1 + 5#)
        X = X + PROB_LANCZOS_P6 / (Zm1 + 6#)
        X = X + PROB_LANCZOS_P7 / (Zm1 + 7#)
        X = X + PROB_LANCZOS_P8 / (Zm1 + 8#)

    'Compute the shifted argument
        T = Zm1 + PROB_LANCZOS_G + 0.5

'------------------------------------------------------------------------------
' RETURN RESULT
'------------------------------------------------------------------------------
    'Return log-gamma
        PROB_LogGamma = _
            PROB_HALF_LOG_TWO_PI + _
            (Zm1 + 0.5) * Log(T) - _
            T + _
            Log(X)
End Function

Public Function PROB_TryLogGamma1p( _
    ByVal X As Double, _
    ByRef Result As Double, _
    ByRef FailMsg As String) _
    As Boolean
'
'==============================================================================
' PROB_TryLogGamma1p
'------------------------------------------------------------------------------
' PURPOSE
'   Returns Log(Gamma(1 + X)) for X >= 0 without ever forming 1 + X.
'
' WHY THIS EXISTS
'   The obvious spelling, PROB_LogGamma(1# + X), silently destroys the answer for
'   small X. In binary64, 1# + X rounds to exactly 1 for every X below 2^-53
'   (about 1.11E-16), so the first-order term -EulerGamma * X is lost outright
'   and the result collapses to zero. The loss is not confined to that boundary:
'   it is already worth a factor of four at X = 1E-14 and grows without limit as
'   X falls. Callers that divide the result by X - the scaled Gamma quantile in
'   particular - inherit the whole error, so a lost increment becomes a lost
'   quantile. This kernel evaluates the Maclaurin series in X directly, so the
'   increment is never asked to survive an addition to one.
'
' METHOD
'   For X at or below PROB_LG1P_SERIES_MAX the fixed Maclaurin table is summed by
'   Horner. The interval and the term count were chosen together from measured
'   error envelopes, not from the theoretical radius of convergence: twenty-six
'   terms hold the series at its coefficient rounding floor across the whole
'   retained interval, and the seam sits where the series is still roughly two
'   orders of magnitude better than the Lanczos route it hands over to.
'   Above the seam, 1# + X preserves the increment and PROB_LogGamma is used.
'
' RECURSION SAFETY
'   Once PROB_LogGamma routes small positive Z through this routine the call
'   graph will look circular. It is not. PROB_LogGamma calls this routine only
'   for a small positive Z inside the series regime. This routine delegates back
'   to PROB_LogGamma only when X exceeds the series crossover, in which case the
'   delegated argument is 1 + X, which is greater than one and therefore cannot
'   re-enter the small-positive PROB_LogGamma branch. The mutual call terminates
'   after at most one hop in each direction.
'
' INPUTS
'   X         Increment above one. Must be finite and non-negative.
'   Result    Receives Log(Gamma(1 + X)) on success.
'   FailMsg   Receives a diagnostic on failure.
'
' RETURNS
'   Boolean
'     TRUE  => Result holds Log(Gamma(1 + X)).
'     FALSE => FailMsg holds a diagnostic; Result is not contractual.
'
' ACCURACY
'   The contract metric is the SCALED absolute error,
'
'       Abs(Result - Reference) / X
'
'   because the scaled Gamma inverse divides this result by X. Ordinary absolute
'   error would look flattering and prove nothing about that caller.
'
'   Measured against 60-digit arithmetic, scaled absolute error is at or below
'   2.4E-16 across X in [3.9E-308, 0.25]. The floor is set by the fifteen-digit
'   coefficient literals, not by the series truncation.
'
' LIMITATION - SUBNORMAL RESULT
'   Below X = 3.9E-308 the product EulerGamma * X is itself subnormal, so the
'   returned Double cannot resolve it. Scaled error then degrades as 2^-1075 / X,
'   reaching 1.3E-14 at X = 1E-310, 1.4E-04 at X = 1E-320 and 4.2E-01 at the
'   smallest positive subnormal. This is a binary64 representability limit of the
'   OUTPUT, not a defect in the series: no evaluation order can place a value on
'   a grid that has no point near it. Callers requiring a scaled contract must
'   restrict X to the normal-result range.
'
' DEPENDENCIES
'   - PROB_IsFinite
'   - PROB_LogGamma
'==============================================================================
'
'------------------------------------------------------------------------------
' VALIDATE INPUTS
'------------------------------------------------------------------------------
    'Reject a non-finite increment
        If Not PROB_IsFinite(X) Then
            FailMsg = "Log-gamma increment must be a finite number"
            Exit Function
        End If

    'Reject a negative increment; this kernel covers the right of X = 0 only
        If X < 0# Then
            FailMsg = "Log-gamma increment must be non-negative"
            Exit Function
        End If

'------------------------------------------------------------------------------
' HANDLE THE EXACT ORIGIN
'------------------------------------------------------------------------------
    'Gamma(1) is one, so the logarithm is exactly zero
        If X = 0# Then
            Result = 0#
            PROB_TryLogGamma1p = True
            Exit Function
        End If

'------------------------------------------------------------------------------
' HAND OVER ABOVE THE MEASURED SEAM
'------------------------------------------------------------------------------
    'Above the seam the increment survives the addition, so use the Lanczos route.
    'The delegated argument is 1 + X > 1, so this cannot re-enter the
    'small-positive PROB_LogGamma branch; see RECURSION SAFETY above.
        If X > PROB_LG1P_SERIES_MAX Then
            Result = PROB_LogGamma(1# + X)
            PROB_TryLogGamma1p = True
            Exit Function
        End If

'------------------------------------------------------------------------------
' SUM THE MACLAURIN SERIES
'------------------------------------------------------------------------------
    'The branch above has established the helper's precondition
        Result = PROB_LogGamma1pSeries(X)

'------------------------------------------------------------------------------
' RETURN SUCCESS
'------------------------------------------------------------------------------
    'Report success
        PROB_TryLogGamma1p = True
End Function

Private Function PROB_LogGamma1pSeries( _
    ByVal X As Double) _
    As Double
'
'==============================================================================
' PROB_LogGamma1pSeries
'------------------------------------------------------------------------------
' PURPOSE
'   Evaluates Log(Gamma(1 + X)) through the fixed Maclaurin polynomial.
'
' PRECONDITION
'   0 <= X <= PROB_LG1P_SERIES_MAX. No validation is performed here.
'
' WHY THIS IS SEPARATE AND NON-TRY
'   Two callers need the polynomial and neither needs a failure channel.
'   PROB_TryLogGamma1p owns validation and dispatch, and calls this only after
'   the precondition holds. PROB_LogGamma returns a Double and so has no way to
'   report a failure; raising one would propagate through the public K_STATS_
'   wrappers and surface as #VALUE!, turning an internal invariant violation
'   into a user-visible error classification. Keeping the polynomial free of a
'   failure channel makes the impossible case impossible to express.
'
' TERMINATION
'   This helper calls nothing. It is the leaf that breaks what would otherwise
'   look like a cycle between PROB_LogGamma and PROB_TryLogGamma1p:
'
'     PROB_LogGamma(Z <= 0.25)      -> this helper                  (leaf)
'     PROB_TryLogGamma1p(X <= 0.25) -> this helper                  (leaf)
'     PROB_TryLogGamma1p(X > 0.25)  -> PROB_LogGamma(1 + X), 1 + X > 1.25,
'                                      which takes the Lanczos branch and can
'                                      never re-enter the small-Z route.
'
' ACCURACY
'   Scaled absolute error, Abs(Result - Reference) / X, at or below 1.4E-16
'   across X in [3.855E-308, 0.25]; see PROB_TryLogGamma1p and
'   benchmark/loggamma1p_study.
'==============================================================================
'
'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Dim Acc                 As Double          'Horner accumulator

'------------------------------------------------------------------------------
' COMPUTE
'------------------------------------------------------------------------------
    'Preconditions belong to the caller and are deliberately NOT re-checked here.
    'Both call sites establish them in the branch condition that selects this
    'helper, and a Double-returning routine has no channel through which to
    'report a violation. Outside the stated interval the polynomial simply
    'returns a truncated series value, which is why the interval is a
    'precondition rather than a guard.

    'Accumulate from the highest retained term downward
        Acc = PROB_LG1P_C26
        Acc = Acc * X + PROB_LG1P_C25
        Acc = Acc * X + PROB_LG1P_C24
        Acc = Acc * X + PROB_LG1P_C23
        Acc = Acc * X + PROB_LG1P_C22
        Acc = Acc * X + PROB_LG1P_C21
        Acc = Acc * X + PROB_LG1P_C20
        Acc = Acc * X + PROB_LG1P_C19
        Acc = Acc * X + PROB_LG1P_C18
        Acc = Acc * X + PROB_LG1P_C17
        Acc = Acc * X + PROB_LG1P_C16
        Acc = Acc * X + PROB_LG1P_C15
        Acc = Acc * X + PROB_LG1P_C14
        Acc = Acc * X + PROB_LG1P_C13
        Acc = Acc * X + PROB_LG1P_C12
        Acc = Acc * X + PROB_LG1P_C11
        Acc = Acc * X + PROB_LG1P_C10
        Acc = Acc * X + PROB_LG1P_C9
        Acc = Acc * X + PROB_LG1P_C8
        Acc = Acc * X + PROB_LG1P_C7
        Acc = Acc * X + PROB_LG1P_C6
        Acc = Acc * X + PROB_LG1P_C5
        Acc = Acc * X + PROB_LG1P_C4
        Acc = Acc * X + PROB_LG1P_C3
        Acc = Acc * X + PROB_LG1P_C2
        Acc = Acc * X + PROB_LG1P_C1

    'The final multiplication carries the leading X factor
        PROB_LogGamma1pSeries = Acc * X
End Function


Public Function PROB_LogGammaHalfDiff( _
    ByVal Z As Double) _
    As Double
'
'==============================================================================
' PROB_LogGammaHalfDiff
'------------------------------------------------------------------------------
' PURPOSE
'   Returns Log(Gamma(Z + 1/2)) - Log(Gamma(Z)) without cancellation.
'
' PRECONDITION
'   Z > 0.
'
' WHY THIS EXISTS
'   This difference appears in the Student t density (with Z = df/2) and in
'   Log(Beta(Z, 1/2)). Formed as a literal subtraction it cancels catastrophically
'   for large Z: at Z = 5E+5 the two log-gammas are each about 6.4E+6 and the
'   answer is about 6.6, so 14 of the 16 available digits are lost. The
'   asymptotic expansion computes the difference directly.
'
' ACCURACY
'   Relative error at or below 2E-14 across the tested range (Z > 0), measured
'   against a 50-digit mpmath reference; typically near machine epsilon, with the
'   worst case near Z = 1.6. The direct-difference branch is used only where it
'   is accurate.
'==============================================================================
'
'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Dim W                   As Double          '1 / Z
    Dim W2                  As Double          'W squared

'------------------------------------------------------------------------------
' DIRECT DIFFERENCE FOR SMALL Z
'------------------------------------------------------------------------------
    'Below the cutoff the subtraction is harmless and the asymptotic is not yet
    'converged
        If Z < PROB_HALF_DIFF_CUTOFF Then
            PROB_LogGammaHalfDiff = PROB_LogGamma(Z + 0.5) - PROB_LogGamma(Z)
            Exit Function
        End If

'------------------------------------------------------------------------------
' ASYMPTOTIC EXPANSION
'------------------------------------------------------------------------------
    'Compute the expansion variable
        W = 1# / Z
        W2 = W * W

    'Return 0.5*Log(Z) - 1/(8Z) + 1/(192 Z^3) - 1/(640 Z^5) + 17/(14336 Z^7)
        PROB_LogGammaHalfDiff = _
            0.5 * Log(Z) - _
            W / 8# + _
            W * W2 / 192# - _
            W * W2 * W2 / 640# + _
            17# * W * W2 * W2 * W2 / 14336#
End Function


Public Function PROB_LogGammaDelta( _
    ByVal LargeArg As Double, _
    ByVal Increment As Double) _
    As Double
'
'==============================================================================
' PROB_LogGammaDelta
'------------------------------------------------------------------------------
' PURPOSE
'   Returns LogGamma(LargeArg + Increment) - LogGamma(LargeArg) as one stable
'   expression, so the two large LogGamma values are never formed and subtracted.
'   Isolating the increment this way avoids the catastrophic cancellation that
'   otherwise wrecks Log(Beta) for unbalanced arguments.
'
' PRECONDITION
'   LargeArg >= 1 and Increment > 0. Intended for Increment <= LargeArg (the
'   unbalanced Beta regime). Accuracy is highest when Increment / LargeArg is
'   small; toward the balanced regime the caller should use the direct
'   three-log-gamma identity instead.
'
' METHOD
'   With the same Lanczos g = 7, n = 9 series A(z) used by PROB_LogGamma and
'   T = LargeArg + g - 1/2:
'
'       LogGamma(z+s) - LogGamma(z) =
'             s * Log(T)
'           + (z + s - 1/2) * Log1p(s / T)
'           - s
'           + Log1p( (A(z+s) - A(z)) / A(z) )
'
'   The 0.5*Log(2*Pi) term cancels in the difference and is absent here. The
'   series difference is formed directly, not by subtracting two series:
'
'       A(z+s) - A(z) = -s * SUM_k Pk / [ (z-1+k)(z+s-1+k) ]
'
'   so no cancellation occurs anywhere in the computation.
'
' ACCURACY
'   Held to LogGammaDelta.all.output in benchmark/accuracy_contracts.csv.
'   Measured worst relative error 6.9E-15 inside the documented regime
'   (Small / LargeArg <= 0.1) across Small in [0.25, ~10] and LargeArg up to
'   1E+50, validated against 50-digit
'   arithmetic. (VBA measurement is the authority; see benchmark/logbeta_study.)
'
' DEPENDENCIES
'   - PROB_Log1p
'==============================================================================
'
'------------------------------------------------------------------------------
' The Lanczos series uses the shared module-level PROB_LANCZOS_* coefficients.
'------------------------------------------------------------------------------

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Dim Az                  As Double          'Lanczos series A(LargeArg)
    Dim dA                  As Double          'A(LargeArg + Increment) - A(LargeArg)
    Dim T                   As Double          'Shifted argument LargeArg + g - 1/2
    Dim Zm1                 As Double          'LargeArg - 1

'------------------------------------------------------------------------------
' COMPUTE
'------------------------------------------------------------------------------
    Zm1 = LargeArg - 1#

    'Lanczos series A(LargeArg)
        Az = PROB_LANCZOS_P0
        Az = Az + PROB_LANCZOS_P1 / (Zm1 + 1#)
        Az = Az + PROB_LANCZOS_P2 / (Zm1 + 2#)
        Az = Az + PROB_LANCZOS_P3 / (Zm1 + 3#)
        Az = Az + PROB_LANCZOS_P4 / (Zm1 + 4#)
        Az = Az + PROB_LANCZOS_P5 / (Zm1 + 5#)
        Az = Az + PROB_LANCZOS_P6 / (Zm1 + 6#)
        Az = Az + PROB_LANCZOS_P7 / (Zm1 + 7#)
        Az = Az + PROB_LANCZOS_P8 / (Zm1 + 8#)

    'Direct series difference A(LargeArg + Increment) - A(LargeArg), no cancellation
        dA = PROB_LANCZOS_P1 / ((Zm1 + 1#) * (Zm1 + 1# + Increment))
        dA = dA + PROB_LANCZOS_P2 / ((Zm1 + 2#) * (Zm1 + 2# + Increment))
        dA = dA + PROB_LANCZOS_P3 / ((Zm1 + 3#) * (Zm1 + 3# + Increment))
        dA = dA + PROB_LANCZOS_P4 / ((Zm1 + 4#) * (Zm1 + 4# + Increment))
        dA = dA + PROB_LANCZOS_P5 / ((Zm1 + 5#) * (Zm1 + 5# + Increment))
        dA = dA + PROB_LANCZOS_P6 / ((Zm1 + 6#) * (Zm1 + 6# + Increment))
        dA = dA + PROB_LANCZOS_P7 / ((Zm1 + 7#) * (Zm1 + 7# + Increment))
        dA = dA + PROB_LANCZOS_P8 / ((Zm1 + 8#) * (Zm1 + 8# + Increment))
        dA = -Increment * dA

    'Shifted argument
        T = LargeArg + PROB_LANCZOS_G - 0.5

    'Stable difference (0.5*Log(2*Pi) cancels and is absent)
        PROB_LogGammaDelta = _
            Increment * Log(T) + _
            (LargeArg + Increment - 0.5) * PROB_Log1p(Increment / T) - _
            Increment + _
            PROB_Log1p(dA / Az)
End Function


Public Function PROB_LogBeta( _
    ByVal A As Double, _
    ByVal B As Double) _
    As Double
'
'==============================================================================
' PROB_LogBeta
'------------------------------------------------------------------------------
' PURPOSE
'   Returns Log(Beta(A, B)) = LogGamma(A) + LogGamma(B) - LogGamma(A + B).
'
' PRECONDITION
'   A > 0 and B > 0.
'
' NUMERICAL POLICY (two regimes)
'   - Half-integer cases use PROB_LogGammaHalfDiff.
'   - Unbalanced arguments (Small / Large < PROB_LOGBETA_STABLE_RATIO) use the
'     stable log-gamma difference:
'
'         Log(Beta) = LogGamma(Small) - PROB_LogGammaDelta(Large, Small)
'
'     PROB_LogGammaDelta forms LogGamma(Large + Small) - LogGamma(Large) as a
'     single expression, so the two large log-gamma values are never subtracted.
'     This is accurate across the whole unbalanced range, including extreme
'     ratios, and replaces the earlier one-term asymptotic branch.
'   - Balanced arguments use the defining identity
'     LogGamma(A) + LogGamma(B) - LogGamma(A + B).
'
' CROSSOVER
'   PROB_LOGBETA_STABLE_RATIO (0.1) is the switch between the two regimes. The
'   constant is validated by the committed VBA seam study (maximum error on each
'   side, continuity across the switch, non-integer Small, multiple absolute
'   scales, and symmetry after argument ordering) and by an independent holdout
'   that straddles the seam (ratios 0.099, 0.101, 0.11). The corresponding
'   PROB_LogBeta accuracy contracts are validated and frozen.
'
' DEPENDENCIES
'   - PROB_LogGamma, PROB_LogGammaHalfDiff, PROB_LogGammaDelta
'   - PROB_HALF_LOG_PI, PROB_LOGBETA_STABLE_RATIO
'==============================================================================
'
'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Dim LargeArg           As Double
    Dim SmallArg           As Double

'------------------------------------------------------------------------------
' HALF-INTEGER SHORTCUTS
'------------------------------------------------------------------------------
    'Log(Beta(A, 1/2)) = 0.5*Log(Pi) - (LogGamma(A + 1/2) - LogGamma(A))
        If B = 0.5 And A >= 1# Then
            PROB_LogBeta = PROB_HALF_LOG_PI - PROB_LogGammaHalfDiff(A)
            Exit Function
        End If

    'Beta is symmetric in its arguments
        If A = 0.5 And B >= 1# Then
            PROB_LogBeta = PROB_HALF_LOG_PI - PROB_LogGammaHalfDiff(B)
            Exit Function
        End If

'------------------------------------------------------------------------------
' ORDER ARGUMENTS
'------------------------------------------------------------------------------
        If A >= B Then
            LargeArg = A
            SmallArg = B
        Else
            LargeArg = B
            SmallArg = A
        End If

'------------------------------------------------------------------------------
' UNBALANCED ARGUMENTS
'------------------------------------------------------------------------------
    'For unbalanced arguments the literal three-log-gamma identity cancels
    'catastrophically. Compute Log(Beta) from the stable log-gamma difference,
    'which never forms and subtracts the two large log-gamma values:
    '    Log(Beta) = LogGamma(Small) - [LogGamma(Large + Small) - LogGamma(Large)]
    '
    'The delta kernel is only defined and measured for LargeArg >= 1, so that
    'precondition is enforced here rather than assumed. Both shapes far below
    'one is NOT the cancelling case: none of the three log-gamma values is
    'large, so the literal identity is well conditioned, while the delta
    'arrangement leaves its validated Lanczos regime and loses accuracy
    '(measured: ~2E-6 absolute at LargeArg = 1E-12, ~9E-3 at 1E-16).
    'Nested rather than a single And: VBA does not short-circuit, so the ratio
    'must not be formed until LargeArg is known to be at least one.
        If LargeArg >= 1# Then
            If SmallArg / LargeArg < PROB_LOGBETA_STABLE_RATIO Then
                PROB_LogBeta = _
                    PROB_LogGamma(SmallArg) - _
                    PROB_LogGammaDelta(LargeArg, SmallArg)
                Exit Function
            End If
        End If

'------------------------------------------------------------------------------
' GENERAL CASE
'------------------------------------------------------------------------------
        PROB_LogBeta = _
            PROB_LogGamma(A) + _
            PROB_LogGamma(B) - _
            PROB_LogGamma(A + B)
End Function


'==============================================================================
' REGULARIZED INCOMPLETE BETA
'==============================================================================

Public Function PROB_StirlingError( _
    ByVal n As Double) _
    As Double
'
'==============================================================================
' PROB_StirlingError
'------------------------------------------------------------------------------
' PURPOSE
'   Returns Loader's Stirling error delta(N), defined by
'       Log(N!) = (N + 0.5) * Log(N) - N + 0.5 * Log(2 * Pi) + delta(N)
'   equivalently  N! = Sqr(2 * Pi * N) * (N / e) ^ N * Exp(delta(N)).
'
' WHY THIS EXISTS
'   delta(N) is O(1 / (12 * N)) and is therefore computed with full relative
'   accuracy at every N. Any quantity that would otherwise be assembled by
'   subtracting two large log-gammas can instead be assembled from three small
'   deltas plus an exactly-computed leading term. PROB_LogChoose is the first
'   consumer; the binomial and Poisson mass functions will be the next.
'
' PRECONDITION
'   N >= 0.
'
' METHOD / PROVENANCE
'   Catherine Loader, "Fast and Accurate Computation of Binomial Probabilities"
'   (2000), the arrangement used by R's dbinom and dpois. Public.
'
'   - N on the half-integer grid at or below 15: an exact stored value. The
'     log-gamma route is accurate only to about 1E-12 RELATIVE there, and delta
'     is small, so a stored constant is both faster and better.
'   - N off the grid and at or below 15: the defining identity via PROB_LogGamma.
'   - N above 15: the asymptotic series in 1 / N, truncated by magnitude.
'   - N below 0.5: one upward recurrence to N + 1 (which uses the paths above),
'     delta(N) = delta(N + 1) + (N + 0.5) * Log((N + 1) / N) - 1.
'
' ACCURACY
'   The authoritative measured accuracy contract lives in
'   benchmark/accuracy_contracts.csv (StirlingError.all.output: 1E-13 absolute;
'   independent holdout worst 3.57E-14 - the earlier 3E-17 was overfit to its
'   grid). RELATIVE error is the wrong metric here: it reaches 1.5E-13 near
'   N = 501, where delta is 1.67E-04, so what propagates into a log-probability
'   is the absolute error.
'
'   The small-N table constants are written as a two-part sum, hi + lo, where hi
'   is the value to 15 significant digits and lo is the residual. VBA source
'   literals hold only about 15 significant digits, so a single literal of a
'   value near 0.15 holds only about 15 significant digits, so the residual term
'   restores the missing low-order bits at load time. Each part is itself a <= 15-digit literal that
'   the editor preserves.
'
' DEPENDENCIES
'   - PROB_LogGamma
'   - PROB_HALF_LOG_TWO_PI  (M_STATS_PROBDIST_CORE)
'==============================================================================
'
'------------------------------------------------------------------------------
' DECLARE CONSTANTS
'------------------------------------------------------------------------------
    'Coefficients of the asymptotic series delta(N) ~ S0/N - S1/N^3 + S2/N^5 ...
        Const S0 As Double = 8.33333333333333E-02     '1 / 12
        Const S1 As Double = 2.77777777777778E-03     '1 / 360
        Const S2 As Double = 7.93650793650794E-04     '1 / 1260
        Const S3 As Double = 5.95238095238095E-04     '1 / 1680
        Const S4 As Double = 8.41750841750842E-04     '1 / 1188

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Dim TwoN                As Double          'N doubled, to test the half-integer grid
    Dim NSquared            As Double          'N * N, the series variable

'------------------------------------------------------------------------------
' SMALL ARGUMENT
'------------------------------------------------------------------------------
    'Below the smallest tabulated positive point the correction is not used
        If n < 0.5 Then
            'Below the tabulated domain recurse up one unit, from
            'Log Gamma(n) = Log Gamma(n + 1) - Log(n):
            '  delta(n) = delta(n + 1) + (n + 0.5) * Log((n + 1) / n) - 1
            'delta(0) is 0 by convention; n + 1 >= 1 lands on the valid paths below.
            If n <= 0# Then
                PROB_StirlingError = 0#
            Else
                PROB_StirlingError = _
                    PROB_StirlingError(n + 1#) + (n + 0.5) * Log((n + 1#) / n) - 1#
            End If
            Exit Function
        End If

'------------------------------------------------------------------------------
' TABULATED REGION
'------------------------------------------------------------------------------
    'Exact stored values on the half-integer grid up to 15
        If n <= 15# Then
            TwoN = 2# * n

            If TwoN = Int(TwoN) Then
                Select Case CLng(TwoN)
            Case 0: PROB_StirlingError = 0#                        'delta(0)
            Case 1: PROB_StirlingError = 0.153426409720027 + 3.45291383939271E-16   'delta(0.5)
            Case 2: PROB_StirlingError = 8.10614667953273E-02 - 4.17803297364056E-17 'delta(1)
            Case 3: PROB_StirlingError = 5.48141210519177E-02 - 4.61038612976516E-17 'delta(1.5)
            Case 4: PROB_StirlingError = 4.13406959554093E-02 - 5.90617791859288E-18 'delta(2)
            Case 5: PROB_StirlingError = 3.31628735199363E-02 - 1.25148894902589E-17 'delta(2.5)
            Case 6: PROB_StirlingError = 2.76779256849983E-02 + 3.91487892927462E-17 'delta(3)
            Case 7: PROB_StirlingError = 2.37461636562975E-02 - 4.02866972090991E-18 'delta(3.5)
            Case 8: PROB_StirlingError = 2.07906721037651E-02 - 6.88847722823215E-18 'delta(4)
            Case 9: PROB_StirlingError = 1.84884505326732E-02 - 1.47692206425174E-17 'delta(4.5)
            Case 10: PROB_StirlingError = 1.66446911898212E-02 - 7.83680513462641E-18 'delta(5)
            Case 11: PROB_StirlingError = 1.51349732219174E-02 - 2.11264861631178E-17 'delta(5.5)
            Case 12: PROB_StirlingError = 1.38761288230707E-02 + 4.79987457270238E-17 'delta(6)
            Case 13: PROB_StirlingError = 1.28104652429202E-02 + 2.69242506552811E-17 'delta(6.5)
            Case 14: PROB_StirlingError = 1.18967099458918E-02 - 2.99049442758823E-17 'delta(7)
            Case 15: PROB_StirlingError = 1.11045597582069E-02 + 1.73266307551973E-17 'delta(7.5)
            Case 16: PROB_StirlingError = 1.04112652619721E-02 - 3.50252143286747E-18 'delta(8)
            Case 17: PROB_StirlingError = 9.7994161261588E-03 + 3.29839037340201E-18   'delta(8.5)
            Case 18: PROB_StirlingError = 9.25546218271273E-03 + 2.9177286366331E-18   'delta(9)
            Case 19: PROB_StirlingError = 8.76870013413939E-03 - 4.53704495273054E-18  'delta(9.5)
            Case 20: PROB_StirlingError = 8.33056343336287E-03 + 1.25646931865963E-18  'delta(10)
            Case 21: PROB_StirlingError = 7.93411456431402E-03 + 5.47249562490943E-19  'delta(10.5)
            Case 22: PROB_StirlingError = 7.57367548795184E-03 + 7.94972024211595E-19  'delta(11)
            Case 23: PROB_StirlingError = 7.24455430132038E-03 + 3.17954619660155E-18  'delta(11.5)
            Case 24: PROB_StirlingError = 6.94284010720953E-03 - 1.34335847336525E-19  'delta(12)
            Case 25: PROB_StirlingError = 6.66524703270768E-03 + 2.4423561808954E-18   'delta(12.5)
            Case 26: PROB_StirlingError = 6.40899418800421E-03 - 2.93156036891702E-18  'delta(13)
            Case 27: PROB_StirlingError = 6.17171226303946E-03 - 2.35246539520223E-18  'delta(13.5)
            Case 28: PROB_StirlingError = 5.95137011275885E-03 - 2.26437558395353E-18  'delta(14)
            Case 29: PROB_StirlingError = 5.74621651301012E-03 - 4.31797389752291E-18  'delta(14.5)
            Case 30: PROB_StirlingError = 5.5547335519628E-03 + 1.37103868995979E-18   'delta(15)
    'Unreachable while 0.5 <= N <= 15 and TwoN is integral. Present so that a
    'broken invariant produces a correct number rather than a silent zero.
                    Case Else
                        PROB_StirlingError = PROB_LogGamma(n + 1#) - _
                                             (n + 0.5) * Log(n) + n - PROB_HALF_LOG_TWO_PI
                End Select

                Exit Function
            End If

    'Off the grid: the defining identity, well conditioned at small N
            PROB_StirlingError = PROB_LogGamma(n + 1#) - _
                                 (n + 0.5) * Log(n) + n - PROB_HALF_LOG_TWO_PI
            Exit Function
        End If

'------------------------------------------------------------------------------
' ASYMPTOTIC SERIES
'------------------------------------------------------------------------------
    'Truncate by magnitude; each cut sits below the Double round-off
        NSquared = n * n

        If n > 500# Then
            PROB_StirlingError = (S0 - S1 / NSquared) / n
        ElseIf n > 80# Then
            PROB_StirlingError = (S0 - (S1 - S2 / NSquared) / NSquared) / n
        ElseIf n > 35# Then
            PROB_StirlingError = (S0 - (S1 - (S2 - S3 / NSquared) / NSquared) / NSquared) / n
        Else
            PROB_StirlingError = (S0 - (S1 - (S2 - (S3 - S4 / NSquared) / NSquared) / NSquared) / NSquared) / n
        End If
End Function

Public Function PROB_TryDeviancePart( _
    ByVal X As Double, _
    ByVal MeanPart As Double, _
    ByRef Result As Double, _
    ByRef FailMsg As String) _
    As Boolean
'
'==============================================================================
' PROB_TryDeviancePart
'------------------------------------------------------------------------------
' PURPOSE
'   Computes Loader's deviance component
'
'       bd0(X, MeanPart) = X * Log(X / MeanPart) + MeanPart - X
'
'   without cancellation when X is close to MeanPart.
'
' PREFER THE LOG-PDF HELPERS
'   This is a low-level primitive with a fragile contract. The Beta case needs
'   the DECOMPOSITION bd0(A, N*X) + bd0(B, N*Y) rather than a raw deviance, and
'   passing the wrong arrangement produces a plausible wrong number rather than
'   an error. PROB_TryGammaLogPdf and PROB_TryBetaLogPdf encapsulate the correct
'   arrangements and should be used in preference to calling this directly.
'
' PRECONDITION
'   X >= 0 and MeanPart > 0.
'
' METHOD
'   Uses Loader's convergent odd-power series when
'   Abs(X-MeanPart) < 0.1 * (X+MeanPart); otherwise uses the direct expression
'   with Log(X)-Log(MeanPart), avoiding overflow in X/MeanPart.
'
' RETURNS
'   Boolean
'     TRUE  => Result contains the non-negative deviance component.
'     FALSE => The bounded series did not converge.
'
' UPDATED
'   2026-07-19 - Loader bd0 implementation with an explicit iteration guard.
'==============================================================================
'
'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Dim Difference          As Double          'X - MeanPart
    Dim SumArguments        As Double          'X + MeanPart
    Dim V                   As Double          'Scaled difference
    Dim V2                  As Double          'V squared
    Dim Ej                  As Double          'Series numerator
    Dim Term                As Double          'Current series term
    Dim SumValue            As Double          'Current series sum
    Dim NewSum              As Double          'Updated series sum
    Dim ScaleValue          As Double          'Convergence scale
    Dim IterIdx             As Long            'Iteration index

'------------------------------------------------------------------------------
' HANDLE X = 0
'------------------------------------------------------------------------------
    'The limiting deviance component is MeanPart
        If X <= 0# Then
            Result = MeanPart
            PROB_TryDeviancePart = True
            Exit Function
        End If

'------------------------------------------------------------------------------
' CHOOSE NUMERICAL BRANCH
'------------------------------------------------------------------------------
        Difference = X - MeanPart
        SumArguments = X + MeanPart

    'Use the convergent series near equality
        If Abs(Difference) < 0.1 * SumArguments Then
            V = Difference / SumArguments
            V2 = V * V
            SumValue = Difference * V
            Ej = 2# * X * V

            For IterIdx = 1 To PROB_BD0_MAX_ITER
                Ej = Ej * V2
                Term = Ej / (2# * CDbl(IterIdx) + 1#)
                NewSum = SumValue + Term

                If NewSum = SumValue Then
                    Result = NewSum
                    PROB_TryDeviancePart = True
                    Exit Function
                End If

                ScaleValue = Abs(NewSum)
                If ScaleValue < 1# Then ScaleValue = 1#

                If Abs(Term) <= PROB_MACH_EPS * ScaleValue Then
                    Result = NewSum
                    PROB_TryDeviancePart = True
                    Exit Function
                End If

                SumValue = NewSum
            Next IterIdx

            FailMsg = "Loader deviance series failed to converge in " & _
                      PROB_BD0_MAX_ITER & " iterations"
            Exit Function
        End If

'------------------------------------------------------------------------------
' DIRECT BRANCH
'------------------------------------------------------------------------------
    'Away from equality the direct expression is well conditioned
        Result = X * (Log(X) - Log(MeanPart)) + MeanPart - X

    'Clamp a tiny negative round-off to the mathematical lower bound zero
        If Result < 0# Then
            If Abs(Result) <= PROB_MACH_EPS * (X + MeanPart) Then
                Result = 0#
            Else
                FailMsg = "Loader deviance calculation produced a negative value"
                Exit Function
            End If
        End If

        PROB_TryDeviancePart = True
End Function


Public Function PROB_TryGammaLogPdf( _
    ByVal StandardX As Double, _
    ByVal Shape As Double, _
    ByVal LogRatio As Double, _
    ByVal LogScale As Double, _
    ByRef LogPdf As Double, _
    ByRef FailMsg As String) _
    As Boolean
'
'==============================================================================
' PROB_TryGammaLogPdf
'------------------------------------------------------------------------------
' PURPOSE
'   Stable Gamma log-density via Loader''s deviance and StirlingError, free of
'   the (Shape-1)*Log(y) - y - LogGamma(Shape) cancellation that loses accuracy
'   at large Shape near the mode.
'
'       log f = -bd0(Shape, StandardX) + 0.5*Log(Shape) - LogRatio
'               - 0.5*Log(2*Pi) - StirlingError(Shape) - LogScale
'
' PRECONDITION
'   StandardX > 0, Shape > 0. Caller handles X = 0, validation and the X/Scale
'   overflow (density zero); this routine computes the interior log-density.
'
' INPUTS
'   StandardX   X / ScaleParam (the standardized variate y)
'   Shape       Gamma shape parameter
'   LogRatio    Log(X) - Log(ScaleParam), i.e. Log(StandardX), formed by the
'               caller without dividing extreme operands
'   LogScale    Log(ScaleParam)
'
' RETURNS
'   Boolean
'     TRUE  => LogPdf holds the natural-log density.
'     FALSE => The deviance series did not converge (FailMsg set).
'
' DEPENDENCIES
'   - PROB_TryDeviancePart
'   - PROB_StirlingError
'   - PROB_HALF_LOG_TWO_PI
'
' UPDATED
'   2026-07-25
'==============================================================================
'
'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Dim Deviance            As Double          'Loader deviance bd0(Shape, y)
'------------------------------------------------------------------------------
' COMPUTE
'------------------------------------------------------------------------------
    If Not PROB_TryDeviancePart(Shape, StandardX, Deviance, FailMsg) Then
        PROB_TryGammaLogPdf = False
        Exit Function
    End If

    LogPdf = -Deviance _
             + 0.5 * Log(Shape) _
             - LogRatio _
             - PROB_HALF_LOG_TWO_PI _
             - PROB_StirlingError(Shape) _
             - LogScale
'------------------------------------------------------------------------------
' RETURN SUCCESS
'------------------------------------------------------------------------------
    PROB_TryGammaLogPdf = True
End Function


Public Function PROB_TryBetaLogPdf( _
    ByVal X As Double, _
    ByVal Y As Double, _
    ByVal Alpha As Double, _
    ByVal BetaShape As Double, _
    ByVal LogX As Double, _
    ByVal LogY As Double, _
    ByRef LogPdf As Double, _
    ByRef FailMsg As String) _
    As Boolean
'
'==============================================================================
' PROB_TryBetaLogPdf
'------------------------------------------------------------------------------
' PURPOSE
'   Stable Beta log-density via Loader''s deviance and StirlingError, free of the
'   (a-1)*Log(x) + (b-1)*Log(1-x) - LogBeta(a,b) cancellation at large shapes.
'   The deviance is the Loader DECOMPOSITION
'
'       D = bd0(Alpha, N*X) + bd0(BetaShape, N*Y),   N = Alpha + BetaShape
'
'   evaluated by PROB_TryDeviancePart; forming the raw a*Log(a/(N*x)) +
'   b*Log(b/(N*y)) instead would reintroduce the cancellation.
'
'       log f = -D + 0.5*(Log(Alpha) + Log(BetaShape) - Log(N))
'               - LogX - LogY - 0.5*Log(2*Pi)
'               - StirlingError(Alpha) - StirlingError(BetaShape)
'               + StirlingError(N)
'
' WHY THE COMPLEMENT IS AN ARGUMENT
'   Y is supplied by the caller and NEVER reconstructed as 1 - X here. When X
'   rounds to 1 its true complement can be far below half an ulp of 1, yet the
'   caller may hold it exactly (the F wrapper''s logistic pair). Recomputing
'   1 - X at that point destroys the complement and, through the deviance,
'   the entire density (CR-P1-01B). When a product N*X or N*Y falls below
'   PROB_BLP_TINY_PRODUCT the matching deviance is formed directly in log
'   space from the caller''s stable LogX / LogY, so an underflowed product
'   cannot corrupt the result either.
'
' PRECONDITION
'   X >= 0, Y >= 0, X + Y = 1 up to rounding, Alpha > 0, BetaShape > 0. The
'   caller validates and handles exact endpoint densities.
'
' INPUTS
'   X           Evaluation point (or the F wrapper''s Beta variate U)
'   Y           Its complement, formed stably by the caller
'   Alpha       First Beta shape
'   BetaShape   Second Beta shape (named to avoid the Beta_Density argument)
'   LogX        Log(X), formed stably by the caller
'   LogY        Log(Y), formed stably by the caller (e.g. PROB_Log1p(-X))
'
' RETURNS
'   Boolean
'     TRUE  => LogPdf holds the natural-log density.
'     FALSE => A deviance series did not converge (FailMsg set).
'
' DEPENDENCIES
'   - PROB_TryDeviancePart
'   - PROB_StirlingError
'   - PROB_HALF_LOG_TWO_PI
'
' UPDATED
'   2026-07-25 - CR-P1-01B: complement passed explicitly, log-form fallback
'==============================================================================
'
'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Dim n                   As Double          'Alpha + BetaShape
    Dim MX                  As Double          'N * X, first deviance argument
    Dim MY                  As Double          'N * Y, second deviance argument
    Dim DevAlpha            As Double          'bd0(Alpha, N*X)
    Dim DevBeta             As Double          'bd0(BetaShape, N*Y)
'------------------------------------------------------------------------------
' COMPUTE
'------------------------------------------------------------------------------
    n = Alpha + BetaShape
    MX = n * X
    MY = n * Y

    'First deviance: log-form when the product underflows, series otherwise.
    'bd0(a, m) = a * (Log(a) - Log(m)) + m - a with Log(m) = Log(N) + LogX.
    If MX < PROB_BLP_TINY_PRODUCT Then
        DevAlpha = Alpha * (Log(Alpha) - (Log(n) + LogX)) + MX - Alpha
    ElseIf Not PROB_TryDeviancePart(Alpha, MX, DevAlpha, FailMsg) Then
        PROB_TryBetaLogPdf = False
        Exit Function
    End If

    'Second deviance, symmetrically.
    If MY < PROB_BLP_TINY_PRODUCT Then
        DevBeta = BetaShape * (Log(BetaShape) - (Log(n) + LogY)) + MY - BetaShape
    ElseIf Not PROB_TryDeviancePart(BetaShape, MY, DevBeta, FailMsg) Then
        PROB_TryBetaLogPdf = False
        Exit Function
    End If

    LogPdf = -(DevAlpha + DevBeta) _
             + 0.5 * (Log(Alpha) + Log(BetaShape) - Log(n)) _
             - LogX _
             - LogY _
             - PROB_HALF_LOG_TWO_PI _
             - PROB_StirlingError(Alpha) _
             - PROB_StirlingError(BetaShape) _
             + PROB_StirlingError(n)
'------------------------------------------------------------------------------
' RETURN SUCCESS
'------------------------------------------------------------------------------
    PROB_TryBetaLogPdf = True
End Function



Public Function PROB_LogChoose( _
    ByVal n As Double, _
    ByVal K As Double) _
    As Double
'
'==============================================================================
' PROB_LogChoose
'------------------------------------------------------------------------------
' PURPOSE
'   Returns Log(C(N, K)), the natural logarithm of the binomial coefficient.
'
' WHY THIS EXISTS
'   C(N, K) overflows a Double at N = 1030 while Log(C(N, K)) stays finite to
'   N = 1E+308. Every discrete mass function, and the hypergeometric in
'   particular, needs the logarithm rather than the coefficient.
'
' PRECONDITION
'   0 <= K <= N. Callers validate; this kernel does not.
'
' METHOD / PROVENANCE
'   The Stirling decomposition
'       Log C(N,K) = 0.5 * Log(N / (2*Pi*K*(N-K)))
'                  + K * Log1p((N-K)/K) + (N-K) * Log1p(K/(N-K))
'                  + delta(N) - delta(K) - delta(N-K)
'   where delta is PROB_StirlingError. Every term is computed directly; nothing
'   large is subtracted from anything large.
'
' WHY NOT THE OBVIOUS ROUTES
'   -Log(N+1) - PROB_LogBeta(N-K+1, K+1) is exact algebra and numerically poor:
'   LogBeta subtracts two log-gammas of size N*Log(N), so its absolute error is
'   about 1.4E-09 at N = 1E+6, and the answer at K = 3 is only 39.65. Measured
'   relative error 3.4E-12 there, and 2.0E+00 at N = 2^53, K = 1, where N + 1
'   rounds back to N. Three PROB_LogGamma calls fail the same way. The product
'   form Prod (N-M+i)/i is accurate but costs Min(K, N-K) logarithms.
'
' ACCURACY
'   Relative error at or below 3.2E-16 across N in [2, 2^53] and all K.
'==============================================================================
'
'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Dim j                   As Double          'N - K, the complementary count
    Dim LeadingTerm         As Double          'The Sqr(N / (2*Pi*K*J)) factor, logged
    Dim EntropyTerm         As Double          'N * H(K/N), the dominant term

'------------------------------------------------------------------------------
' BOUNDARY CASES
'------------------------------------------------------------------------------
    'C(N,0) = C(N,N) = 1, so the logarithm is exactly zero
        If K <= 0# Or K >= n Then
            PROB_LogChoose = 0#
            Exit Function
        End If

'------------------------------------------------------------------------------
' COMPUTE
'------------------------------------------------------------------------------
    'Complementary count
        j = n - K

    'Leading term, expanded so that K * J never overflows
        LeadingTerm = 0.5 * (Log(n) - Log(PROB_TWO_PI) - Log(K) - Log(j))

    'Entropy term. Both logarithms are of a ratio at least one, so neither
    'cancels; Log1p carries the case where that ratio is close to one
        EntropyTerm = K * PROB_Log1p(j / K) + j * PROB_Log1p(K / j)

    'Assemble with the three small Stirling corrections
        PROB_LogChoose = LeadingTerm + EntropyTerm + _
                         PROB_StirlingError(n) - PROB_StirlingError(K) - PROB_StirlingError(j)
End Function


Public Function PROB_TryBetaRegularized( _
    ByVal X As Double, _
    ByVal Y As Double, _
    ByVal A As Double, _
    ByVal B As Double, _
    ByRef Result As Double, _
    ByRef FailMsg As String) _
    As Boolean
'
'==============================================================================
' PROB_TryBetaRegularized
'------------------------------------------------------------------------------
' PURPOSE
'   Computes the regularized incomplete beta function I_X(A, B).
'
' PRECONDITION
'   A > 0, B > 0, X >= 0, Y >= 0, and X + Y = 1 in exact arithmetic.
'
' WHY TWO ARGUMENTS
'   The caller passes both X and its complement Y, each computed from a form that
'   does not cancel. This routine never forms 1 - X or 1 - Y internally. That one
'   change is what makes the Student t CDF exact near zero: with a single
'   argument, X = df / (df + x^2) rounds to exactly 1 as soon as x^2/df drops
'   below 1.1E-16, and the CDF collapses to exactly 0.5, losing eight digits.
'
' RETURNS
'   Boolean
'     TRUE  => Result holds I_X(A, B).
'     FALSE => the continued fraction did not converge; FailMsg says so and
'              Result is left unchanged.
'
' DEPENDENCIES
'   - PROB_LogBeta
'   - PROB_TryExp
'   - PROB_TryBetaContinuedFraction
'==============================================================================
'
'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Dim LogBt               As Double          'Log of the common beta factor
    Dim Bt                  As Double          'Common beta factor
    Dim CFValue             As Double          'Continued-fraction value
    Dim Value               As Double          'Working result

'------------------------------------------------------------------------------
' HANDLE BOUNDARIES
'------------------------------------------------------------------------------
    'Return boundary values exactly
        If X <= 0# Then
            Result = 0#
            PROB_TryBetaRegularized = True
            Exit Function
        End If

        If Y <= 0# Then
            Result = 1#
            PROB_TryBetaRegularized = True
            Exit Function
        End If

'------------------------------------------------------------------------------
' COMPUTE COMMON FACTOR
'------------------------------------------------------------------------------
    'Compute the log of X^A * Y^B / Beta(A, B); this factor is symmetric under
    'the simultaneous swap (X,A) <-> (Y,B), so one value serves both branches
    'CR-P1-02 regime dispatch. The literal form cancels two quantities of size
    'N*Log(N), so above the crossover it degrades without bound (at A = B = 1E4
    'it is already wrong by 8.5E-13, at 1E16 by 8.5E-01). The Loader
    'decomposition removes that subtraction. Below the crossover the literal
    'form is measured to be at least as accurate - both sit at ~1E-15 near
    'A + B = 200 - and it is the form the frozen tiny and unbalanced contracts
    'were validated against, so it is retained there. The density carries A-1
    'and B-1 powers while this factor carries A and B, hence the two added logs.
        If A + B >= PROB_IBETA_LOADER_MIN_SHAPE Then
            If Not PROB_TryBetaLogPdf(X, Y, A, B, Log(X), Log(Y), LogBt, FailMsg) Then
                Exit Function
            End If

            LogBt = LogBt + Log(X) + Log(Y)
        Else
            LogBt = A * Log(X) + B * Log(Y) - PROB_LogBeta(A, B)
        End If

    'Exponentiate; underflow to zero is a valid result at the far edges
        If Not PROB_TryExp(LogBt, Bt) Then
            FailMsg = "Incomplete beta factor overflowed for A = " & A & _
                      ", B = " & B
            Exit Function
        End If

'------------------------------------------------------------------------------
' EVALUATE CONTINUED FRACTION
'------------------------------------------------------------------------------
    'Use the direct expansion where it converges, the reflected one elsewhere;
    'note that the reflected branch consumes Y directly and never 1 - X
        If X < (A + 1#) / (A + B + 2#) Then
            If Not PROB_TryBetaContinuedFraction(A, B, X, CFValue, FailMsg) Then Exit Function
            Value = Bt * CFValue / A
        Else
            If Not PROB_TryBetaContinuedFraction(B, A, Y, CFValue, FailMsg) Then Exit Function
            Value = 1# - Bt * CFValue / B
        End If

'------------------------------------------------------------------------------
' RETURN SUCCESS
'------------------------------------------------------------------------------
    'Clamp small numerical overshoots at the closed unit interval
        If Value < 0# Then Value = 0#
        If Value > 1# Then Value = 1#

    'Return the regularized value
        Result = Value
    'Return success
        PROB_TryBetaRegularized = True
End Function


Public Function PROB_TryBetaContinuedFraction( _
    ByVal A As Double, _
    ByVal B As Double, _
    ByVal X As Double, _
    ByRef Result As Double, _
    ByRef FailMsg As String) _
    As Boolean
'
'==============================================================================
' PROB_TryBetaContinuedFraction
'------------------------------------------------------------------------------
' PURPOSE
'   Evaluates the incomplete beta continued fraction by the modified Lentz
'   method.
'
' PRECONDITION
'   A > 0, B > 0, 0 < X < 1, and X < (A + 1) / (A + B + 2).
'
' RETURNS
'   Boolean
'     TRUE  => Result holds the continued-fraction value.
'     FALSE => PROB_BETA_MAX_ITER was exhausted without meeting PROB_NUM_EPS.
'              Result is left unchanged; a partial sum is never returned.
'==============================================================================
'
'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Dim Qab                 As Double          'A + B
    Dim Qap                 As Double          'A + 1
    Dim Qam                 As Double          'A - 1
    Dim c                   As Double          'Lentz c accumulator
    Dim D                   As Double          'Lentz d accumulator
    Dim h                   As Double          'Continued-fraction value
    Dim Aa                  As Double          'Coefficient
    Dim Del                 As Double          'Multiplicative increment
    Dim M                   As Long            'Iteration index
    Dim M2                  As Long            '2 * iteration index

'------------------------------------------------------------------------------
' INITIALIZE
'------------------------------------------------------------------------------
    'Initialize constants
        Qab = A + B
        Qap = A + 1#
        Qam = A - 1#

    'Initialize Lentz's method
        c = 1#
        D = 1# - Qab * X / Qap

        If Abs(D) < PROB_FPMIN Then D = PROB_FPMIN

        D = 1# / D
        h = D

'------------------------------------------------------------------------------
' ITERATE CONTINUED FRACTION
'------------------------------------------------------------------------------
    'Loop over continued-fraction terms
        For M = 1 To PROB_BETA_MAX_ITER
            M2 = 2 * M

            'Even step
                Aa = M * (B - M) * X / ((Qam + M2) * (A + M2))

                D = 1# + Aa * D
                If Abs(D) < PROB_FPMIN Then D = PROB_FPMIN

                c = 1# + Aa / c
                If Abs(c) < PROB_FPMIN Then c = PROB_FPMIN

                D = 1# / D
                h = h * D * c

            'Odd step
                Aa = -(A + M) * (Qab + M) * X / ((A + M2) * (Qap + M2))

                D = 1# + Aa * D
                If Abs(D) < PROB_FPMIN Then D = PROB_FPMIN

                c = 1# + Aa / c
                If Abs(c) < PROB_FPMIN Then c = PROB_FPMIN

                D = 1# / D
                Del = D * c
                h = h * Del

            'Return on convergence
                If Abs(Del - 1#) <= PROB_NUM_EPS Then
                    Result = h
                    PROB_TryBetaContinuedFraction = True
                    Exit Function
                End If
        Next M

'------------------------------------------------------------------------------
' REPORT NON-CONVERGENCE
'------------------------------------------------------------------------------
    'Never return a partial sum
        FailMsg = "Incomplete beta continued fraction failed to converge in " & _
                  PROB_BETA_MAX_ITER & " iterations for A = " & A & ", B = " & B
End Function


Public Function PROB_TryBetaInvRegularized( _
    ByVal Probability As Double, _
    ByVal ComplementProbability As Double, _
    ByVal A As Double, _
    ByVal B As Double, _
    ByRef ResultX As Double, _
    ByRef ResultY As Double, _
    ByRef FailMsg As String) _
    As Boolean
'
'==============================================================================
' PROB_TryBetaInvRegularized
'------------------------------------------------------------------------------
' PURPOSE
'   Solves I_X(A, B) = Probability for X in (0, 1), returning both X and its
'   complement Y = 1 - X.
'
' PRECONDITION
'   A > 0, B > 0, 0 < Probability < 1, and
'   Probability + ComplementProbability = 1 in exact arithmetic. The caller
'   supplies the complement because 1 - Probability is exact only when
'   Probability >= 0.5 (Sterbenz), and the far tail depends on it.
'
' BEHAVIOR
'   - Solves on whichever tail is the smaller, so the quantity driven to its
'     target keeps full relative precision. Without this the upper tail loses
'     everything: I_X - Probability with both near 1 is pure cancellation.
'   - Returns both X and Y so that callers such as the F quantile, which needs
'     X / Y, do not have to re-derive the small one by subtraction. This is what
'     lets F.INV reach 1E+34.
'   - Seeds from the AS 109 normal approximation when A > 1 and B > 1, and from
'     the leading series term otherwise, then runs Newton's method safeguarded by
'     a bisection bracket, which cannot diverge.
'
' RETURNS
'   Boolean
'     TRUE  => ResultX and ResultY hold the solution.
'     FALSE => an inner incomplete beta evaluation failed; FailMsg says so.
'
' DEPENDENCIES
'   - PROB_NormalInvCDFRaw, PROB_LogBeta, PROB_TryExp
'   - PROB_TryBetaRegularized
'==============================================================================
'
'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Dim SolveDirect         As Boolean         'TRUE when solving for X, FALSE for Y
    Dim RoundTrip           As Double          'Forward probability at the root
    Dim TailMass            As Double          'Min(Probability, Complement)
    Dim Sa                  As Double          'Shape parameter of the solved tail
    Dim Sb                  As Double          'Other shape parameter
    Dim Target              As Double          'Probability mass being matched
    Dim LogBetaAB           As Double          'Log(Beta(Sa, Sb))
    Dim U                   As Double          'Current iterate
    Dim UNew                As Double          'Proposed iterate
    Dim Low                 As Double          'Bisection lower bound
    Dim High                As Double          'Bisection upper bound
    Dim Ibeta               As Double          'I_U(Sa, Sb)
    Dim Residual            As Double          'Ibeta - Target
    Dim Density             As Double          'Beta density at U
    Dim LogDensity          As Double          'Log of the beta density at U
    Dim LogSeed             As Double          'Log of the series-inverted seed
    Dim Z                   As Double          'Normal seed
    Dim R                   As Double          'AS 109 working value
    Dim S1                  As Double          'AS 109 working value
    Dim S2                  As Double          'AS 109 working value
    Dim HH                  As Double          'AS 109 working value
    Dim W                   As Double          'AS 109 working value
    Dim ExpTwoW             As Double          'Exp(2 * W), overflow-guarded seed factor
    Dim Converged           As Boolean         'TRUE once the iterate has settled
    Dim IterIdx             As Long            'Iteration index

'------------------------------------------------------------------------------
' HANDLE BOUNDARIES
'------------------------------------------------------------------------------
    'Return boundary values exactly
        If Probability <= 0# Then
            ResultX = 0#
            ResultY = 1#
            PROB_TryBetaInvRegularized = True
            Exit Function
        End If

        If ComplementProbability <= 0# Then
            ResultX = 1#
            ResultY = 0#
            PROB_TryBetaInvRegularized = True
            Exit Function
        End If

'------------------------------------------------------------------------------
' ORIENT ONTO THE SMALLER TAIL
'------------------------------------------------------------------------------
    'Solve I_U(Sa, Sb) = Target where Target <= 0.5, then unwind
        If Probability <= ComplementProbability Then
            SolveDirect = True
            Sa = A
            Sb = B
            Target = Probability
        Else
            SolveDirect = False
            Sa = B
            Sb = A
            Target = ComplementProbability
        End If

    'Cache the log-beta for the density evaluations
        LogBetaAB = PROB_LogBeta(Sa, Sb)

'------------------------------------------------------------------------------
' SEED
'------------------------------------------------------------------------------
    'Use the AS 109 normal approximation when both shapes exceed one
        If Sa > 1# And Sb > 1# Then
            Z = PROB_NormalInvCDFRaw(Target)
            R = (Z * Z - 3#) / 6#
            S1 = 1# / (2# * Sa - 1#)
            S2 = 1# / (2# * Sb - 1#)
            HH = 2# / (S1 + S2)
            W = Z * Sqr(HH + R) / HH - (S2 - S1) * (R + 5# / 6# - 2# / (3# * HH))
            'Guard the module's one raw exponential; W is bounded above here
            '(Target <= 0.5 forces Z <= 0) so overflow is not reachable in
            'practice, and the seed clamp below recovers U on any failure
            If PROB_TryExp(2# * W, ExpTwoW) Then U = Sa / (Sa + Sb * ExpTwoW)
    'Otherwise invert the leading series term I_U ~ U^Sa / (Sa * Beta(Sa, Sb))
        Else
            LogSeed = (Log(Target) + Log(Sa) + LogBetaAB) / Sa
            If Not PROB_TryExp(LogSeed, U) Then U = 0.5
        End If

    'Force the seed strictly inside the open unit interval
        If U <= 0# Or U >= 1# Or Not PROB_IsFinite(U) Then U = 0.5 * Target + 0.25

'------------------------------------------------------------------------------
' SAFEGUARDED NEWTON
'------------------------------------------------------------------------------
    'Initialize the bracket
        Low = 0#
        High = 1#

    'Iterate
        For IterIdx = 1 To PROB_INV_MAX_ITER
            'Evaluate the objective
                If Not PROB_TryBetaRegularized(U, 1# - U, Sa, Sb, Ibeta, FailMsg) Then Exit Function
                Residual = Ibeta - Target

            'Tighten the bracket; I_U is increasing in U
                If Residual < 0# Then
                    If U > Low Then Low = U
                Else
                    If U < High Then High = U
                End If

            'Evaluate the beta density, the derivative of the objective
                Density = 0#
                If U > 0# And U < 1# Then
                    LogDensity = (Sa - 1#) * Log(U) + (Sb - 1#) * Log(1# - U) - LogBetaAB
                    If Not PROB_TryExp(LogDensity, Density) Then Density = 0#
                End If

            'Take a Newton step, falling back to bisection when it is unusable
                If Density <= 0# Then
                    UNew = 0.5 * (Low + High)
                Else
                    UNew = U - Residual / Density
                    If UNew <= Low Or UNew >= High Then UNew = 0.5 * (Low + High)
                End If

            'Return on convergence, including the case where the iterate has
            'settled onto a single Double and can no longer move
                If Abs(UNew - U) <= PROB_MACH_EPS * Abs(UNew) Or UNew = U Then
                    U = UNew
                    Converged = True
                    Exit For
                End If

            'Advance
                U = UNew
        Next IterIdx

'------------------------------------------------------------------------------
' REPORT NON-CONVERGENCE
'------------------------------------------------------------------------------
    'Never return an unsettled iterate
        If Not Converged Then
            FailMsg = "Incomplete beta inverse failed to converge in " & _
                      PROB_INV_MAX_ITER & " iterations for A = " & A & ", B = " & B
            Exit Function
        End If

'------------------------------------------------------------------------------
' RETURN SUCCESS
'------------------------------------------------------------------------------
    'Unwind the orientation, keeping the small side exact
        If SolveDirect Then
            ResultX = U
            ResultY = 1# - U
        Else
            ResultX = 1# - U
            ResultY = U
        End If

'------------------------------------------------------------------------------
' VALIDATE REPRESENTABILITY
'------------------------------------------------------------------------------
    'For sufficiently small shapes the distribution is numerically
    'indistinguishable in binary64 from a two-point limit whose interior CDF
    'plateau is about B / (A + B). Away from that plateau the mathematical
    'quantile lies below the smallest positive Double, or closer to one than the
    'nearest representable interior Double, so no interior root exists to
    'return. Verify the solved root reproduces the requested probability and
    'report the boundary explicitly instead of returning an arbitrary interior
    'value. Checked only when both shapes are below one - the only regime where
    'the collapse occurs - so the common path is unaffected.
    'See BetaInverse.InteriorQuantileRepresentability.
        If A < 1# And B < 1# Then

            If Not PROB_TryBetaRegularized( _
                ResultX, ResultY, A, B, RoundTrip, FailMsg) Then Exit Function

            TailMass = Probability
            If ComplementProbability < TailMass Then TailMass = ComplementProbability

            If Abs(RoundTrip - Probability) > _
               PROB_BETAINV_ROUNDTRIP_TOL * TailMass Then

                FailMsg = "No representable interior quantile: with A = " & A & _
                          " and B = " & B & " the distribution collapses to its" & _
                          " two-point limit in binary64, so the quantile for" & _
                          " probability " & Probability & " is not representable"
                Exit Function

            End If

        End If

    'Return success
        PROB_TryBetaInvRegularized = True
End Function


'==============================================================================
' REGULARIZED INCOMPLETE GAMMA
'==============================================================================

Public Function PROB_TryGammaRegularizedP( _
    ByVal A As Double, _
    ByVal X As Double, _
    ByRef Result As Double, _
    ByRef FailMsg As String) _
    As Boolean
'
'==============================================================================
' PROB_TryGammaRegularizedP
'------------------------------------------------------------------------------
' PURPOSE
'   Computes the regularized lower incomplete gamma function P(A, X).
'
' PRECONDITION
'   A > 0 and X >= 0.
'
' RETURNS
'   Boolean
'     TRUE  => Result holds P(A, X).
'     FALSE => an inner series or continued fraction failed to converge.
'==============================================================================
'
'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Dim Value               As Double          'Working result

'------------------------------------------------------------------------------
' HANDLE BOUNDARY
'------------------------------------------------------------------------------
    'Return the boundary value exactly
        If X <= 0# Then
            Result = 0#
            PROB_TryGammaRegularizedP = True
            Exit Function
        End If

'------------------------------------------------------------------------------
' COMPUTE
'------------------------------------------------------------------------------
    'Use the series expansion in the lower region and the continued fraction above
        If X < A + 1# Then
            If Not PROB_TryGammaSeriesP(A, X, Value, FailMsg) Then Exit Function
        Else
            If Not PROB_TryGammaContinuedFractionQ(A, X, Value, FailMsg) Then Exit Function
            Value = 1# - Value
        End If

'------------------------------------------------------------------------------
' RETURN SUCCESS
'------------------------------------------------------------------------------
    'Clamp small numerical overshoots
        If Value < 0# Then Value = 0#
        If Value > 1# Then Value = 1#

    'Return the regularized value
        Result = Value
    'Return success
        PROB_TryGammaRegularizedP = True
End Function


Public Function PROB_TryGammaRegularizedQ( _
    ByVal A As Double, _
    ByVal X As Double, _
    ByRef Result As Double, _
    ByRef FailMsg As String) _
    As Boolean
'
'==============================================================================
' PROB_TryGammaRegularizedQ
'------------------------------------------------------------------------------
' PURPOSE
'   Computes the regularized upper incomplete gamma function Q(A, X) = 1 - P(A, X).
'
' PRECONDITION
'   A > 0 and X >= 0.
'
' WHY THIS EXISTS SEPARATELY
'   Q is not usefully recovered as 1 - P. For a chi-square with 10 degrees of
'   freedom at x = 200, Q is 1.6E-37 while 1 - P evaluates to exactly zero. Any
'   right-tail p-value has to come from here.
'==============================================================================
'
'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Dim Value               As Double          'Working result

'------------------------------------------------------------------------------
' HANDLE BOUNDARY
'------------------------------------------------------------------------------
    'Return the boundary value exactly
        If X <= 0# Then
            Result = 1#
            PROB_TryGammaRegularizedQ = True
            Exit Function
        End If

'------------------------------------------------------------------------------
' COMPUTE
'------------------------------------------------------------------------------
    'Use the series expansion in the lower region and the continued fraction above
        If X < A + 1# Then
            If Not PROB_TryGammaSeriesP(A, X, Value, FailMsg) Then Exit Function
            Value = 1# - Value
        Else
            If Not PROB_TryGammaContinuedFractionQ(A, X, Value, FailMsg) Then Exit Function
        End If

'------------------------------------------------------------------------------
' RETURN SUCCESS
'------------------------------------------------------------------------------
    'Clamp small numerical overshoots
        If Value < 0# Then Value = 0#
        If Value > 1# Then Value = 1#

    'Return the regularized value
        Result = Value
    'Return success
        PROB_TryGammaRegularizedQ = True
End Function


Private Function PROB_TryGammaPrefactor( _
    ByVal A As Double, _
    ByVal X As Double, _
    ByRef Factor As Double, _
    ByRef FailMsg As String) _
    As Boolean
'
'==============================================================================
' PROB_TryGammaPrefactor
'------------------------------------------------------------------------------
' PURPOSE
'   Returns the incomplete-gamma prefactor X^A * Exp(-X) / Gamma(A) without the
'   cancellation of the literal form -X + A*Log(X) - LogGamma(A).
'
' WHY THIS EXISTS
'   The literal form subtracts two quantities of size A*Log(A) to leave a modest
'   logarithm. LogGamma carries a RELATIVE error contract, so at A = 1E12 the
'   absolute error already reaches ~2E-3, and by A = 1E16 the prefactor is wrong
'   by e^46 - a silent, catastrophic error in every probability built on it.
'   Routing through the stable Loader log-density removes the subtraction:
'
'       Log(X^A * Exp(-X) / Gamma(A)) = GammaLogPdf(X; A, scale 1) + Log(X)
'
'   because the density carries A - 1 powers of X and the prefactor carries A.
'
' INPUTS
'   A               Shape parameter, strictly positive
'   X               Evaluation point, strictly positive
'
' RETURNS
'   Boolean
'     TRUE  => Factor holds the prefactor; underflow to zero is a valid result.
'     FALSE => The log-density failed or the prefactor overflowed (FailMsg set).
'
' DEPENDENCIES
'   - PROB_TryGammaLogPdf
'   - PROB_TryExp
'
' UPDATED
'   2026-07-29 - CR-P1-02: cancellation-free incomplete-gamma prefactor
'==============================================================================
'
'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Dim LogPdf              As Double          'Stable Gamma log-density at X
'------------------------------------------------------------------------------
' COMPUTE
'------------------------------------------------------------------------------
    If Not PROB_TryGammaLogPdf(X, A, Log(X), 0#, LogPdf, FailMsg) Then
        Exit Function
    End If

    If Not PROB_TryExp(LogPdf + Log(X), Factor) Then
        FailMsg = "Incomplete gamma prefactor overflowed for A = " & A
        Exit Function
    End If
'------------------------------------------------------------------------------
' RETURN SUCCESS
'------------------------------------------------------------------------------
    PROB_TryGammaPrefactor = True
End Function


Public Function PROB_TryGammaSeriesP( _
    ByVal A As Double, _
    ByVal X As Double, _
    ByRef Result As Double, _
    ByRef FailMsg As String) _
    As Boolean
'
'==============================================================================
' PROB_TryGammaSeriesP
'------------------------------------------------------------------------------
' PURPOSE
'   Evaluates P(A, X) by the lower incomplete gamma series expansion.
'
' PRECONDITION
'   A > 0 and 0 < X < A + 1.
'
' RETURNS
'   Boolean
'     TRUE  => Result holds P(A, X).
'     FALSE => PROB_GAMMA_MAX_ITER was exhausted without meeting PROB_NUM_EPS.
'              Result is left unchanged; a partial sum is never returned.
'
' NOTE
'   The term count grows like 2.4 * Sqr(A) at the worst point X = A. At A = 800
'   (a chi-square with 1600 degrees of freedom) it exceeds 200 terms, which is
'   where the previous 200-iteration budget began returning silently wrong
'   answers, reaching a 37 percent error by df = 100000.
'==============================================================================
'
'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Dim Ap                  As Double          'A plus iteration index
    Dim SumValue            As Double          'Series sum
    Dim Del                 As Double          'Series increment
    Dim Factor              As Double          'Exp(-X + A*Log(X) - LogGamma(A))
    Dim IterIdx             As Long            'Iteration index

'------------------------------------------------------------------------------
' INITIALIZE
'------------------------------------------------------------------------------
    'Initialize the first term
        Ap = A
        SumValue = 1# / A
        Del = SumValue

'------------------------------------------------------------------------------
' SERIES ITERATION
'------------------------------------------------------------------------------
    'Loop over terms
        For IterIdx = 1 To PROB_GAMMA_MAX_ITER
            Ap = Ap + 1#
            Del = Del * X / Ap
            SumValue = SumValue + Del

            'Return on convergence
                If Abs(Del) <= Abs(SumValue) * PROB_GAMMA_SERIES_EPS Then
                    'Cancellation-free prefactor (CR-P1-02)
                    If Not PROB_TryGammaPrefactor(A, X, Factor, FailMsg) Then
                        Exit Function
                    End If

                    Result = SumValue * Factor
                    PROB_TryGammaSeriesP = True
                    Exit Function
                End If
        Next IterIdx

'------------------------------------------------------------------------------
' REPORT NON-CONVERGENCE
'------------------------------------------------------------------------------
    'Never return a partial sum
        FailMsg = "Incomplete gamma series failed to converge in " & _
                  PROB_GAMMA_MAX_ITER & " iterations for A = " & A
End Function


Public Function PROB_TryGammaContinuedFractionQ( _
    ByVal A As Double, _
    ByVal X As Double, _
    ByRef Result As Double, _
    ByRef FailMsg As String) _
    As Boolean
'
'==============================================================================
' PROB_TryGammaContinuedFractionQ
'------------------------------------------------------------------------------
' PURPOSE
'   Evaluates Q(A, X), the regularized upper incomplete gamma function, by the
'   modified Lentz continued fraction.
'
' PRECONDITION
'   A > 0 and X >= A + 1.
'
' RETURNS
'   Boolean
'     TRUE  => Result holds Q(A, X).
'     FALSE => PROB_GAMMA_MAX_ITER was exhausted without meeting PROB_NUM_EPS.
'==============================================================================
'
'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Dim B                   As Double          'Continued-fraction b value
    Dim c                   As Double          'Lentz c accumulator
    Dim D                   As Double          'Lentz d accumulator
    Dim h                   As Double          'Continued-fraction value
    Dim An                  As Double          'Coefficient
    Dim Del                 As Double          'Multiplicative increment
    Dim Factor              As Double          'Exp(-X + A*Log(X) - LogGamma(A))
    Dim IterIdx             As Long            'Iteration index

'------------------------------------------------------------------------------
' INITIALIZE
'------------------------------------------------------------------------------
    'Initialize the continued fraction
        B = X + 1# - A

        If Abs(B) < PROB_FPMIN Then B = PROB_FPMIN

        c = 1# / PROB_FPMIN
        D = 1# / B
        h = D

'------------------------------------------------------------------------------
' CONTINUED FRACTION ITERATION
'------------------------------------------------------------------------------
    'Loop over continued-fraction terms
        For IterIdx = 1 To PROB_GAMMA_MAX_ITER
            An = -CDbl(IterIdx) * (CDbl(IterIdx) - A)
            B = B + 2#

            D = An * D + B
            If Abs(D) < PROB_FPMIN Then D = PROB_FPMIN

            c = B + An / c
            If Abs(c) < PROB_FPMIN Then c = PROB_FPMIN

            D = 1# / D
            Del = D * c
            h = h * Del

            'Return on convergence
                If Abs(Del - 1#) <= PROB_NUM_EPS Then
                    'Cancellation-free prefactor (CR-P1-02)
                    If Not PROB_TryGammaPrefactor(A, X, Factor, FailMsg) Then
                        Exit Function
                    End If

                    Result = Factor * h
                    PROB_TryGammaContinuedFractionQ = True
                    Exit Function
                End If
        Next IterIdx

'------------------------------------------------------------------------------
' REPORT NON-CONVERGENCE
'------------------------------------------------------------------------------
    'Never return a partial value
        FailMsg = "Incomplete gamma continued fraction failed to converge in " & _
                  PROB_GAMMA_MAX_ITER & " iterations for A = " & A
End Function


Public Function PROB_TryGammaInvP( _
    ByVal Probability As Double, _
    ByVal ComplementProbability As Double, _
    ByVal A As Double, _
    ByRef Result As Double, _
    ByRef FailMsg As String) _
    As Boolean
'
'==============================================================================
' PROB_TryGammaInvP
'------------------------------------------------------------------------------
' PURPOSE
'   Solves P(A, X) = Probability for X > 0.
'
' PRECONDITION
'   A > 0, 0 < Probability < 1, and
'   Probability + ComplementProbability = 1 in exact arithmetic.
'
' BEHAVIOR
'   - Drives the smaller of P and Q onto its target, so the residual never
'     consists of two nearly equal numbers being subtracted.
'   - Seeds from the Wilson-Hilferty cube-root normal approximation, then runs
'     Newton's method safeguarded by a bisection bracket.
'
' RETURNS
'   Boolean
'     TRUE  => Result holds the quantile.
'     FALSE => an inner incomplete gamma evaluation failed; FailMsg says so.
'
' DEPENDENCIES
'   - PROB_NormalInvCDFRaw, PROB_LogGamma, PROB_TryExp
'   - PROB_TryGammaRegularizedP, PROB_TryGammaRegularizedQ
'==============================================================================
'
'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Dim SolveLower          As Boolean         'TRUE when matching P, FALSE when matching Q
    Dim Target              As Double          'Probability mass being matched
    Dim X                   As Double          'Current iterate
    Dim XNew                As Double          'Proposed iterate
    Dim Low                 As Double          'Bisection lower bound
    Dim High                As Double          'Bisection upper bound
    Dim HasHigh             As Boolean         'TRUE once an upper bound is known
    Dim Value               As Double          'P or Q at the iterate
    Dim Residual            As Double          'Signed distance to the target
    Dim Density             As Double          'Gamma density at the iterate
    Dim LogDensity          As Double          'Stable log-density for the Newton step
    Dim Z                   As Double          'Normal seed
    Dim T                   As Double          'Wilson-Hilferty working value
    Dim Converged           As Boolean         'TRUE once the iterate has settled
    Dim IterIdx             As Long            'Iteration index

'------------------------------------------------------------------------------
' HANDLE BOUNDARIES
'------------------------------------------------------------------------------
    'Return the lower boundary exactly
        If Probability <= 0# Then
            Result = 0#
            PROB_TryGammaInvP = True
            Exit Function
        End If

    'Refuse the degenerate upper boundary
        If ComplementProbability <= 0# Then
            FailMsg = "Gamma quantile is unbounded at Probability = 1"
            Exit Function
        End If

'------------------------------------------------------------------------------
' ORIENT ONTO THE SMALLER TAIL
'------------------------------------------------------------------------------
    'Match P below the median and Q above it
        If Probability <= ComplementProbability Then
            SolveLower = True
            Target = Probability
        Else
            SolveLower = False
            Target = ComplementProbability
        End If

'------------------------------------------------------------------------------
' SEED
'------------------------------------------------------------------------------
    'Wilson-Hilferty: X ~ A * (1 - 1/(9A) + Z/Sqr(9A))^3
        If Probability <= 0.5 Then
            Z = PROB_NormalInvCDFRaw(Probability)
        Else
            Z = -PROB_NormalInvCDFRaw(ComplementProbability)
        End If

        T = 1# - 1# / (9# * A) + Z / Sqr(9# * A)
        X = A * T * T * T

    'Fall back to the leading series term for small shape or a nonsense seed
        If A < 1# Or X <= 0# Or Not PROB_IsFinite(X) Then
            If SolveLower Then
                If Not PROB_TryExp((Log(Probability) + PROB_LogGamma(A + 1#)) / A, X) Then X = A
            Else
                X = A
            End If
        End If

    'Force a strictly positive seed
        If X <= 0# Or Not PROB_IsFinite(X) Then X = 0.00000001

'------------------------------------------------------------------------------
' SAFEGUARDED NEWTON
'------------------------------------------------------------------------------
    'Initialize the bracket; the upper bound is discovered by expansion
        Low = 0#
        High = 0#
        HasHigh = False

    'Iterate
        For IterIdx = 1 To PROB_INV_MAX_ITER
            'Evaluate the objective, always increasing in X
                If SolveLower Then
                    If Not PROB_TryGammaRegularizedP(A, X, Value, FailMsg) Then Exit Function
                    Residual = Value - Target
                Else
                    If Not PROB_TryGammaRegularizedQ(A, X, Value, FailMsg) Then Exit Function
                    Residual = Target - Value
                End If

            'Tighten the bracket
                If Residual < 0# Then
                    If X > Low Then Low = X
                Else
                    If (Not HasHigh) Or X < High Then High = X
                    HasHigh = True
                End If

            'Evaluate the gamma density, the derivative of the objective
                'Cancellation-free density (CR-P1-02); a failure leaves Density
                'at zero, which routes the step to the bisection fallback below.
                If Not PROB_TryGammaLogPdf(X, A, Log(X), 0#, LogDensity, FailMsg) Then
                    Density = 0#
                ElseIf Not PROB_TryExp(LogDensity, Density) Then
                    Density = 0#
                End If

            'Take a Newton step, falling back to bisection or expansion
                If Density <= 0# Then
                    If HasHigh Then
                        XNew = 0.5 * (Low + High)
                    Else
                        XNew = 2# * X
                    End If
                Else
                    XNew = X - Residual / Density

                    If HasHigh Then
                        If XNew <= Low Or XNew >= High Then XNew = 0.5 * (Low + High)
                    ElseIf XNew <= Low Then
                        XNew = 2# * X
                    End If
                End If

            'Return on convergence, including the case where the iterate has
            'settled onto a single Double and can no longer move
                If Abs(XNew - X) <= PROB_EPS * Abs(XNew) Or XNew = X Then
                    X = XNew
                    Converged = True
                    Exit For
                End If

            'Advance
                X = XNew
        Next IterIdx

'------------------------------------------------------------------------------
' REPORT NON-CONVERGENCE
'------------------------------------------------------------------------------
    'Never return an unsettled iterate
        If Not Converged Then
            FailMsg = "Incomplete gamma inverse failed to converge in " & _
                      PROB_INV_MAX_ITER & " iterations for A = " & A
            Exit Function
        End If

'------------------------------------------------------------------------------
' RETURN SUCCESS
'------------------------------------------------------------------------------
    'Return the quantile
        Result = X
    'Return success
        PROB_TryGammaInvP = True
End Function


