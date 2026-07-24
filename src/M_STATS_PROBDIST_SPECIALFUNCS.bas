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
'     large-parameter case. The incomplete gamma series needs about 2.4 * Sqr(df)
'     terms at its worst point (x = a); the beta continued fraction needs about
'     0.27 * Sqr(df). Typical degrees of freedom below 100 converge in under 70
'     iterations. PROB_GAMMA_MAX_ITER = 100000 covers df up to roughly 1E+9;
'     PROB_BETA_MAX_ITER = 100000 covers df up to roughly 1E+7. These convergence
'     ranges are narrower than the 1E100 representational validation bound
'     (PROB_PARAMETER_MAGNITUDE_GUARD): a parameter between a kernel's convergence range and
'     1E100 is accepted by validation, attempted, and then returns a clean
'     parameter-named non-convergence error rather than a wrong answer. The
'     ranges are approximate because the true boundary depends on the companion
'     arguments, so they are documented rather than enforced as hard cliffs.
'   - PROB_LogGamma is recursive through its reflection branch, exactly once.
'
' UPDATED
'   2026-07-21
'==============================================================================

'==============================================================================
' PRIVATE CONSTANTS
'==============================================================================

Private Const PROB_BETA_MAX_ITER       As Long = 100000   'Lentz iterations, incomplete beta
Private Const PROB_GAMMA_MAX_ITER      As Long = 100000   'Series / Lentz iterations, incomplete gamma
Private Const PROB_INV_MAX_ITER        As Long = 200      'Safeguarded Newton iterations
Private Const PROB_HALF_DIFF_CUTOFF    As Double = 20#    'Z at or above which the asymptotic half-difference wins
Private Const PROB_LOGBETA_STABLE_RATIO As Double = 0.1     'Small/Large below this uses the stable LogGamma difference (validated by the committed seam study and independent holdout)

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
'   Relative error below 6.1E-14 across Z in [1E-8, 1E+50], measured against
'   50-digit arithmetic.
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
' REFLECTION FORMULA
'------------------------------------------------------------------------------
    'Use the reflection formula for small positive Z
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
'   Relative error at or below ~5E-15 for Increment / LargeArg <= 0.1, across
'   Increment in [0.25, ~10] and LargeArg up to 1E+50, validated against 50-digit
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
'
' ACCURACY
'   Absolute error at or below 3E-17 for every N >= 0.5. RELATIVE error is the
'   wrong metric here: it reaches 1.5E-13 near N = 501, where delta is 1.67E-04.
'   What propagates into a log-probability is the absolute error.
'
'   The small-N table constants are written as a two-part sum, hi + lo, where hi
'   is the value to 15 significant digits and lo is the residual. VBA source
'   literals hold only about 15 significant digits, so a single literal of a
'   value near 0.15 could not reach 3E-17; the residual term restores the missing
'   low-order bits at load time. Each part is itself a <= 15-digit literal that
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
            PROB_StirlingError = 0#
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
        LogBt = A * Log(X) + B * Log(Y) - PROB_LogBeta(A, B)

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
                If Abs(Del) <= Abs(SumValue) * PROB_NUM_EPS Then
                    If Not PROB_TryExp(-X + A * Log(X) - PROB_LogGamma(A), Factor) Then
                        FailMsg = "Incomplete gamma series prefactor overflowed for A = " & A
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
                    If Not PROB_TryExp(-X + A * Log(X) - PROB_LogGamma(A), Factor) Then
                        FailMsg = "Incomplete gamma prefactor overflowed for A = " & A
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
                If Not PROB_TryExp(-X + (A - 1#) * Log(X) - PROB_LogGamma(A), Density) Then Density = 0#

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




