Attribute VB_Name = "M_STATS_PROBDIST_SPECIALFUNCS"
Option Explicit
Option Private Module

'==============================================================================
' M_STATS_PROBDIST_SPECIALFUNCS
'------------------------------------------------------------------------------
' PURPOSE
'   Provides the log-gamma, log-beta, regularized incomplete beta and
'   regularized incomplete gamma functions, together with their inverses, and
'   the standard normal upper tail, as distribution-agnostic kernels with an
'   explicit success/failure contract.
'
' WHY THIS EXISTS
'   Every distribution outside the normal family reduces to one of two objects:
'   the regularized incomplete beta function I_x(a, b) and the regularized
'   incomplete gamma functions P(a, x) and Q(a, x). Student t, chi-square, F,
'   gamma, beta, Poisson, binomial and the negative binomial are all thin
'   wrappers around them. Isolating those objects here means the hard numerics
'   are written, tested and fixed exactly once.
'
'   The standard normal upper tail is here for the same reason. The normal
'   family needs it, and so does the large-df Student-t route (#34), both to
'   full relative precision; two copies would be two chances to disagree.
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
'     - PROB_TryLogGamma1p
'
'   Validated envelopes:
'     - PROB_DENSITY_SHAPE_MAX
'
'   Positive-ratio dispatch (#13):
'     - PROB_PositiveRatioClass, PROB_ClassifyPositiveRatio
'     - PROB_PRS_CUTOFF_48BIT, PROB_PRS_CUTOFF_40BIT
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
'   Standard normal upper tail (#72):
'     - PROB_NormalUpperTail
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
'       z < 0.5 and the LogGamma1p series route for small positive z. The
'       previous global claim (relative error below 6.1E-14 for z in
'       [1E-8, 1E+50]) is withdrawn; accuracy is now stated per regime by
'       the LogGamma.* contracts in benchmark/accuracy_contracts.csv, with
'       evidence in benchmark/loggamma_regimes_study. Public, published;
'       not proprietary.
'   - PROB_TryLogGamma1p:
'       Maclaurin series for Log(Gamma(1 + X)) on the small-X interior,
'       avoiding the 1 + X collapse of the general route. Contracts
'       LogGamma1p.small.scaled_abs and LogGamma1p.series_seam.scaled_abs;
'       evidence in benchmark/loggamma1p_study. Public, published.
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
'   - PROB_NormalUpperTail:
'       A degree-(7, 7) rational approximation fitted in this repository by a
'       preregistered, deterministic procedure
'       (benchmark/normal_tail_study/derive_kernel.py) below Z = 3, and the
'       classical Laplace continued fraction for the Mills ratio at and above
'       it. No third-party coefficients are used.
'   Nothing here is a newly-invented algorithm; the Try contract, the (X, Y)
'   argument pair and the iteration budgets are the local contribution.
'
' DESIGN PRINCIPLES
'   - Every iterative routine returns Boolean. False means the answer is unknown,
'     not approximately known. On a False return the ByRef result is not
'     contractual and must not be read (CORE invariant I3).
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
'   2026-10-13 - Standard normal upper tail added (#72).
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

'Positive-ratio dispatch cutoffs for ICR-P1-01A (#13). Both derive from
'PROB_MIN_NORMAL by exact division by a power of two, so no subnormal literal
'is ever parsed: a VBA source literal cannot be relied on to denote one, and
'this study has already been bitten by a hand-written constant that was one
'ULP wrong.
'
'The comparison is STRICT less-than, and that is load-bearing. / 16# is
'2 ^ -1026, so the largest 48-significand-bit subnormal routes to the log path
'and the smallest 49-bit value does not; / 4096# is 2 ^ -1034 and reproduces
'the 40-bit Gamma cumulative boundary the same way. Both boundaries were
'confirmed on an independent, preregistered holdout - see
'benchmark/positive_ratio_subnormal/HOLDOUT_DESIGN.md.
Public Const PROB_PRS_CUTOFF_48BIT As Double = PROB_MIN_NORMAL / 16#       'Gamma/ChiSquare density, ChiSquare cumulative
Public Const PROB_PRS_CUTOFF_40BIT As Double = PROB_MIN_NORMAL / 4096#     'Gamma cumulative

'Classification of a standardized argument for positive-ratio dispatch.
'
'THREE STATES, NOT A CUTOFF COMPARISON. Hard underflow is tested
'independently of the cutoff and is NOT the bottom of the precision scale:
'a stored zero carries no information at all, whereas a positive subnormal
'carries some. Phase A established that distinction, and collapsing it into
'StandardX <= Cutoff would erase it - which is why this is an Enum and not
'a Boolean. The first Enum in this project; the alternative was a Boolean
'that a later refactor could not have been stopped from conflating.
Public Enum PROB_PositiveRatioClass
    PROB_PRC_DIRECT = 0                'Existing path; nothing was lost
    PROB_PRC_PRECISION_RECOVERY = 1    'Positive subnormal below the cutoff
    PROB_PRC_HARD_UNDERFLOW = 2        'Positive ratio stored as zero
End Enum
Private Const PROB_IBETA_LOADER_MIN_SHAPE As Double = 1000# 'A + B at or above which the incomplete-beta factor uses the Loader decomposition instead of the literal form (measured crossover; see the COMPUTE COMMON FACTOR comment in PROB_TryBetaRegularized and benchmark/cdf_large_shape)
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
'at run time. Twenty-six terms hold the series at the arithmetic rounding
'floor throughout the retained interval; see PROB_TryLogGamma1p.
Private Const PROB_LG1P_SERIES_MAX As Double = 0.25   'Measured series/Lanczos seam
'C1 is written as a split expression rather than a fifteen-digit literal. It
'dominates the scaled-error floor and is the one coefficient that survives
'division by Shape in the scaled Gamma inverse, so a single ulp here is worth
'more than it is anywhere else in the table. The sum below is the exactly
'rounded Double for -0.5772156649015329; the fifteen-digit form is 8 ulps away
'and raises the scaled-error floor. Same pattern as PROB_DS_MAX_EXACT_INTEGER.
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

'Standard normal upper tail (#72). Every value below is fixed by the
'preregistered study in benchmark/normal_tail_study, which records each one
'to 40 digits in kernel_coefficients.json. A constant that a 15-digit literal
'cannot hold is written as the sum of two such literals, and that sum is the
'correctly rounded Double; test_vba_kernel_source.py checks every pair against
'the study. Do not retype, round or "tidy" these values.
'
'P and D are the numerator and denominator of the degree-(7, 7) rational
'approximation to Q(Z) * Exp(Z ^ 2 / 2) on 0 <= Z < 3, with D(0) = 1. Every
'coefficient except P7 is positive; P7 is -2.5E-10, and its whole effect on
'[0, 3] is a cancellation factor of 1 + 1.5E-7 (Amendment 1 of the study).
Private Const PROB_NT_P0 As Double = 0.5
Private Const PROB_NT_P1 As Double = 0.480184367590059 - 3.88578058618805E-16
Private Const PROB_NT_P2 As Double = 0.23576729183073 + 3.33066907387547E-16
Private Const PROB_NT_P3 As Double = 6.89088968314517E-02
Private Const PROB_NT_P4 As Double = 1.24964409031359E-02 + 4.16333634234434E-17
Private Const PROB_NT_P5 As Double = 1.31859652187628E-03 + 1.0842021724855E-18
Private Const PROB_NT_P6 As Double = 6.34118830319193E-05 + 1.35525271560688E-20
Private Const PROB_NT_P7 As Double = -2.54113287998232E-10 + 1.55096364853693E-25
Private Const PROB_NT_D1 As Double = 1.75825329598299 - 4.88498130835069E-15
Private Const PROB_NT_D2 As Double = 1.37441774250694 + 4.44089209850063E-16
Private Const PROB_NT_D3 As Double = 0.621279362780297 - 3.33066907387547E-16
Private Const PROB_NT_D4 As Double = 0.176120941664172 - 3.60822483003176E-16
Private Const PROB_NT_D5 As Double = 3.14745662345996E-02 + 3.46944695195361E-17
Private Const PROB_NT_D6 As Double = 3.30581981247468E-03
Private Const PROB_NT_D7 As Double = 1.58922498110447E-04 - 2.71050543121376E-19

Private Const PROB_NT_INV_SQRT_TWO_PI As Double = _
    0.398942280401433 - 2.77555756156289E-16
Private Const PROB_NT_SPLIT        As Double = 3#     'Rational below, continued fraction at and above
Private Const PROB_NT_CF_LEVELS    As Long = 60       'Laplace continued-fraction depth; truncation 2.0E-18 at Z = 3
Private Const PROB_NT_UNDERFLOW_Z  As Double = 38.5   'At approximately 38.49 the one-sided tail rounds to zero in Double


'==============================================================================
' LOGARITHMIC GAMMA
'==============================================================================

' BEGIN public Gamma bit types
Private Type PROB_DD_Number
    Value As Double
End Type
Private Type PROB_DD_Words
    Low As Long
    High As Long
End Type
' END public Gamma bit types

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
'   Held to the LogGamma.small_positive, LogGamma.near_zero and LogGamma.general
'   contracts in benchmark/accuracy_contracts.csv (evidence:
'   benchmark/loggamma_regimes_study). The former claim, relative error below
'   6.1E-14 across Z in [1E-8, 1E+50], was withdrawn as a global statement for
'   two reasons measured in that study:
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
    'Measured in benchmark/loggamma_regimes_study: at the smallest positive Double
    'the reflection route reaches 6.2E-05 relative; this route is held to
    'LogGamma.small_positive.log_abs. The series helper is called directly
    'rather than through PROB_TryLogGamma1p, because the branch condition
    'already establishes the helper's precondition and a Double-returning kernel
    'has no channel to report a failure that cannot happen.
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
'   X falls. Callers that divide the result by X inherit the whole error, so a
'   lost increment becomes a lost quantile. This kernel is provided for such
'   callers, chiefly the scaled Gamma quantile seed in PROB_TryGammaInvP, which
'   has not yet been rerouted onto it (#14). It evaluates the Maclaurin series
'   in X directly, so the increment is never asked to survive an addition to
'   one.
'
' METHOD
'   For X at or below PROB_LG1P_SERIES_MAX the fixed Maclaurin table is summed by
'   Horner. The interval and the term count were chosen together from measured
'   error envelopes, not from the theoretical radius of convergence: twenty-six
'   terms hold the series at its arithmetic rounding floor across the whole
'   retained interval, and the seam sits where the series is still roughly two
'   orders of magnitude better than the Lanczos route it hands over to.
'   Above the seam, 1# + X preserves the increment and PROB_LogGamma is used.
'
' RECURSION SAFETY
'   PROB_LogGamma does not call this routine; both share the leaf
'   PROB_LogGamma1pSeries. This routine delegates to PROB_LogGamma only when X
'   exceeds the series crossover, with argument 1 + X > 1.25, which takes the
'   Lanczos branch and cannot re-enter the small-Z route.
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
'   because the scaled Gamma inverse it is built for divides this result by X
'   (#14). Ordinary absolute error would look flattering and prove nothing
'   about that caller.
'
'   Scaled absolute error is held to LogGamma1p.small.scaled_abs in
'   benchmark/accuracy_contracts.csv (evidence: benchmark/loggamma1p_study). The
'   floor is the rounding of the final Acc * X, not the series truncation or the
'   coefficient literals.
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
'   Scaled absolute error, Abs(Result - Reference) / X, is held to
'   LogGamma1p.small.scaled_abs in benchmark/accuracy_contracts.csv; see
'   PROB_TryLogGamma1p and benchmark/loggamma1p_study.
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
'   for large Z: at Z = 5E+5 the two log-gammas are each about 6.1E+6 and the
'   answer is about 6.6, so about six of the sixteen available digits are lost.
'   The asymptotic expansion computes the difference directly.
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
'   arithmetic. (VBA measurement is the authority; see benchmark/delta_seam_study.)
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

Public Function PROB_ClassifyPositiveRatio( _
    ByVal MathematicallyPositive As Boolean, _
    ByVal StandardX As Double, _
    ByVal Cutoff As Double) _
    As PROB_PositiveRatioClass
'
'==============================================================================
' PROB_ClassifyPositiveRatio
'------------------------------------------------------------------------------
' PURPOSE
'   Decides which evaluation path a standardized argument requires, for
'   ICR-P1-01A (#13).
'
' WHY THE DECISION LIVES AT THE CALLER
'   PROB_TryGammaRegularizedP and _Q receive only the rounded StandardX. A
'   mathematically positive ratio that underflowed is indistinguishable there
'   from a genuine support value of zero, so their boundary handling is
'   CORRECT given what they are told and must not be made conditional. Only
'   the caller still holds the fact that the input was positive.
'
'   This function does not compute anything. It classifies, so that the
'   caller decides and the shared kernels still own the arithmetic.
'
' THE THREE STATES
'     MathematicallyPositive and StandardX = 0   -> PROB_PRC_HARD_UNDERFLOW
'     StandardX > 0 and StandardX < Cutoff       -> PROB_PRC_PRECISION_RECOVERY
'     otherwise                                  -> PROB_PRC_DIRECT
'
'   Hard underflow is tested FIRST and independently of Cutoff. It is not the
'   bottom of the precision scale: a stored zero retains no information about
'   the ratio, while a positive subnormal retains some. Writing this as
'   StandardX <= Cutoff would give the same answers today and silently erase
'   the distinction the moment a cutoff moved.
'
' WHY THE COMPARISON IS STRICT
'   Cutoff is a power of two derived from PROB_MIN_NORMAL, and the holdout
'   confirmed the boundary in retained significand bits. PROB_MIN_NORMAL / 16#
'   is 2 ^ -1026, so a strict < routes the largest 48-bit subnormal and leaves
'   the smallest 49-bit value on the direct path. A <= would move both
'   boundaries by one bucket, past the confirmed evidence.
'
' ARGUMENTS
'   MathematicallyPositive  True when the caller has established that the
'                           exact ratio is strictly positive. At every #13
'                           call site this holds because X > 0 is guaranteed
'                           before standardization and the scale is validated
'                           positive and finite.
'   StandardX               The stored standardized argument.
'   Cutoff                  PROB_PRS_CUTOFF_48BIT or PROB_PRS_CUTOFF_40BIT,
'                           per the surface being evaluated. The two families
'                           do not share a cumulative boundary.
'
' RETURNS
'   A PROB_PositiveRatioClass value. Never raises.
'
' UPDATED
'   2026-08-15
'==============================================================================
'
'------------------------------------------------------------------------------
' CLASSIFY
'------------------------------------------------------------------------------
    'Hard underflow first, and never via the cutoff: a stored zero is a
    'different kind of loss from a positive subnormal, not a smaller one
        If MathematicallyPositive And StandardX = 0# Then
            PROB_ClassifyPositiveRatio = PROB_PRC_HARD_UNDERFLOW
            Exit Function
        End If

    'Strict less-than: see WHY THE COMPARISON IS STRICT above
        If StandardX > 0# Then
            If StandardX < Cutoff Then
                PROB_ClassifyPositiveRatio = PROB_PRC_PRECISION_RECOVERY
                Exit Function
            End If
        End If

        PROB_ClassifyPositiveRatio = PROB_PRC_DIRECT
End Function


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
'   deltas plus an exactly-computed leading term. Consumers: PROB_LogChoose,
'   PROB_TryGammaLogPdf, PROB_TryBetaLogPdf, and the binomial and Poisson mass
'   functions in M_STATS_PROBDIST_DISCRETE.
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
'   - N off the grid and at or below 15: upward recurrence to the asymptotic
'     region. This avoids importing the absolute error of PROB_LogGamma into
'     Gamma-series prefactors whose public contract is tighter.
'   - N above 15: the asymptotic series in 1 / N, truncated by magnitude.
'   - N below 0.5: the same upward recurrence, with the first step evaluated as
'     Log1p(N) - Log(N) so that the ratio is never formed. Subsequent steps use
'     PROB_Log1p(1 / N), where the reciprocal is bounded by two.
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
'   literals hold only about 15 significant digits, so the residual term
'   restores the missing low-order bits at load time. Each part is itself a
'   <= 15-digit literal that the editor preserves.
'
' DOMAIN
'   Finite for every N from the smallest positive Double upward. The earlier
'   arrangement formed (N + 1) / N and raised an overflow below N = 2 ^ -1024,
'   which reached the public surface as a refused density for a shape whose
'   true density was small and representable.
'
' DEPENDENCIES
'   - PROB_Log1p            (M_STATS_PROBDIST_CORE)
'   - PROB_HALF_LOG_TWO_PI  (M_STATS_PROBDIST_CORE)
'
' UPDATED
'   2026-08-28 - #23: off-grid values recur to the asymptotic anchor instead of
'                importing the LogGamma error floor into Gamma-series P.
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
            'delta(0) is 0 by convention. An off-grid n + 1 continues upward
            'until it reaches the asymptotic anchor; a half-integer lands on the
            'stored table. This is bounded by sixteen calls.
            'The log of the ratio is taken as Log1p(n) - Log(n) and the ratio
            'itself is never formed. Forming it overflows once 1 / n leaves
            'the Double range, at n = 2 ^ -1024, which is inside the accepted
            'shape domain: the guard has no lower cutoff. Measured in
            'benchmark/stirling_overflow_probe. PROB_Log1p returns n
            'unchanged once 1 + n rounds to one, so nothing here needs the
            'increment to survive either.
            If n <= 0# Then
                PROB_StirlingError = 0#
            Else
                PROB_StirlingError = _
                    PROB_StirlingError(n + 1#) + _
                    (n + 0.5) * (PROB_Log1p(n) - Log(n)) - 1#
            End If
            Exit Function
        End If

'------------------------------------------------------------------------------
' TABULATED REGION
'------------------------------------------------------------------------------
    'Use exact stored values on the half-integer grid up to 15
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
    'broken invariant continues through the accurate off-grid path rather than
    'returning a silent zero.
                    Case Else
                        PROB_StirlingError = _
                            PROB_StirlingError(n + 1#) + _
                            (n + 0.5) * PROB_Log1p(1# / n) - 1#
                End Select

                Exit Function
            End If

    'Off the grid: recur upward to the asymptotic anchor. The previous defining
    'identity called PROB_LogGamma and contributed about 1E-14 absolute error
    'to the Stirling correction. That is inside StirlingError's own 1E-13
    'contract but becomes the same relative error in an incomplete-gamma
    'prefactor, breaching the tighter Gamma/Chi-square CDF contract (#23).
    'Here n >= 0.5, so 1 / n is bounded by two and cannot overflow.
            PROB_StirlingError = _
                PROB_StirlingError(n + 1#) + _
                (n + 0.5) * PROB_Log1p(1# / n) - 1#
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
'   - PROB_TryBetaLogPdf (A + B >= PROB_IBETA_LOADER_MIN_SHAPE)
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
'     FALSE => an inner incomplete beta evaluation failed, Newton/bisection did
'              not settle within PROB_INV_MAX_ITER, or (both shapes below one)
'              no representable interior quantile exists; FailMsg says which.
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
' PURPOSE
'   Computes the regularized Gamma P tail through the scaled two-Double pair.
' PRECONDITION
'   Finite A > 0 and X >= 0. X = 0 is returned exactly.
' METHOD
'   Lower Q is reconstructed directly from log P, before binary64 rounding.
'   The frozen wide kernel handles the region outside the extended interior.
' ERROR POLICY
'   False clears Result and reports invalid inputs, arithmetic failure or cap.
'   Retains the established 100000-iteration cap diagnostic wording.
' UPDATED
'   2026-11-08 - #24/#26 public integration candidate; acceptance pending.
'==============================================================================
    Dim P As Double, Q As Double
    Result = 0#: FailMsg = ""
    If Not PROB_DD_WG_Pair(A, X, P, Q, FailMsg) Then
        If InStr(1, FailMsg, "continued fraction cap 100000", vbBinaryCompare) > 0 Then
            FailMsg = "Incomplete gamma continued fraction failed to converge in 100000 iterations for A = " & A
        ElseIf InStr(1, FailMsg, "lower series cap 100000", vbBinaryCompare) > 0 Then
            FailMsg = "Incomplete gamma series failed to converge in 100000 iterations for A = " & A
        ElseIf Len(FailMsg) = 0 Then
            FailMsg = "Gamma kernel requires finite A > 0 and X >= 0"
        End If
        Exit Function
    End If
    If Not PROB_IsFinite(P) Or P < 0# Or P > 1# Then
        FailMsg = "Gamma kernel produced an invalid probability"
        Exit Function
    End If
    Result = P
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
' PURPOSE
'   Computes the regularized Gamma Q tail through the scaled two-Double pair.
' PRECONDITION
'   Finite A > 0 and X >= 0. X = 0 is returned exactly.
' METHOD
'   Lower Q is reconstructed directly from log P, before binary64 rounding.
'   The frozen wide kernel handles the region outside the extended interior.
' ERROR POLICY
'   False clears Result and reports invalid inputs, arithmetic failure or cap.
'   Retains the established 100000-iteration cap diagnostic wording.
' UPDATED
'   2026-11-08 - #24/#26 public integration candidate; acceptance pending.
'==============================================================================
    Dim P As Double, Q As Double
    Result = 0#: FailMsg = ""
    If Not PROB_DD_WG_Pair(A, X, P, Q, FailMsg) Then
        If InStr(1, FailMsg, "continued fraction cap 100000", vbBinaryCompare) > 0 Then
            FailMsg = "Incomplete gamma continued fraction failed to converge in 100000 iterations for A = " & A
        ElseIf InStr(1, FailMsg, "lower series cap 100000", vbBinaryCompare) > 0 Then
            FailMsg = "Incomplete gamma series failed to converge in 100000 iterations for A = " & A
        ElseIf Len(FailMsg) = 0 Then
            FailMsg = "Gamma kernel requires finite A > 0 and X >= 0"
        End If
        Exit Function
    End If
    If Not PROB_IsFinite(Q) Or Q < 0# Or Q > 1# Then
        FailMsg = "Gamma kernel produced an invalid probability"
        Exit Function
    End If
    Result = Q
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
'     FALSE => PROB_GAMMA_MAX_ITER was exhausted without meeting
'              PROB_GAMMA_SERIES_EPS, or the prefactor failed (FailMsg set).
'              Result is left unchanged; a partial sum is never returned.
'
' NOTE
'   The term count grows like roughly 7 * Sqr(A) at the worst point X = A. At
'   A = 800 (a chi-square with 1600 degrees of freedom) it exceeds 200 terms,
'   which is where the previous 200-iteration budget began returning silently
'   wrong answers, reaching a 37 percent error by df = 100000.
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


Private Function PROB_GammaSmallLogGamma1p(ByVal A As Double) As Double
'
'==============================================================================
' PROB_GammaSmallLogGamma1p
'------------------------------------------------------------------------------
' PURPOSE
'   Evaluates Log(Gamma(1+A)) with exact endpoint zeros for the small-shape Q path.
' PRECONDITION
'   0 <= A <= 1. The caller establishes this interval.
' METHOD
'   Degree-32 Taylor polynomial at 0.5 for Log(Gamma(1+A))/(A*(A-1)).
'   Exact-arithmetic truncation is bounded by 1/(8*3^32); this is not a
'   bound on the full rounded survival computation. Coefficients were generated
'   by independent 100/140-digit constructions; see issue #26, candidate B6.
' ERROR POLICY
'   No independent failure channel; finite polynomial on the stated interval.
' CALLED FROM
'   PROB_TryGammaContinuedFractionQ
' UPDATED
'   2026-11-08 - #26: B6 candidate, pending exact-source Excel validation.
'==============================================================================
    Dim Acc As Double
    Dim U As Double
    Dim Product As Double

    'Each coefficient is a two-part literal preserving the frozen binary64 value.
    Const C0 As Double = 4.83128950540981E-01 + (-1.11022302462516E-16)
    Const C1 As Double = -1.45959895914306E-01 + (-8.32667268468867E-17)
    Const C2 As Double = 6.29114010745649E-02 + (4.16333634234434E-17)
    Const C3 As Double = -3.13084875010110E-02 + (0.00000000000000E+00)
    Const C4 As Double = 1.67970986311869E-02 + (-2.08166817117217E-17)
    Const C5 As Double = -9.42562244846982E-03 + (-1.73472347597681E-18)
    Const C6 As Double = 5.44645784455033E-03 + (-3.46944695195361E-18)
    Const C7 As Double = -3.21207407719414E-03 + (3.03576608295941E-18)
    Const C8 As Double = 1.92291614029804E-03 + (-3.68628738645072E-18)
    Const C9 As Double = -1.16441447410474E-03 + (2.16840434497101E-19)
    Const C10 As Double = 7.11522258031626E-04 + (-3.25260651745651E-19)
    Const C11 As Double = -4.37995120948431E-04 + (5.42101086242752E-20)
    Const C12 As Double = 2.71277099973908E-04 + (3.79470760369927E-19)
    Const C13 As Double = -1.68894325207693E-04 + (1.89735380184963E-19)
    Const C14 As Double = 1.05623637963857E-04 + (1.35525271560688E-20)
    Const C15 As Double = -6.63135482599680E-05 + (1.35525271560688E-20)
    Const C16 As Double = 4.17769653553246E-05 + (-4.74338450462408E-20)
    Const C17 As Double = -2.63996995905211E-05 + (0.00000000000000E+00)
    Const C18 As Double = 1.67282243428494E-05 + (4.40457132572236E-20)
    Const C19 As Double = -1.06260564852470E-05 + (-3.04931861011548E-20)
    Const C20 As Double = 6.76496373796090E-06 + (-5.08219768352580E-21)
    Const C21 As Double = -4.31562117239537E-06 + (2.54109884176290E-21)
    Const C22 As Double = 2.75822905390482E-06 + (0.00000000000000E+00)
    Const C23 As Double = -1.76587736954239E-06 + (6.35274710440725E-22)
    Const C24 As Double = 1.13233727790091E-06 + (-4.02340649945793E-21)
    Const C25 As Double = -7.27150972981638E-07 + (-3.17637355220363E-22)
    Const C26 As Double = 4.67585459593438E-07 + (1.05879118406788E-22)
    Const C27 As Double = -3.01053697400556E-07 + (-5.29395592033938E-23)
    Const C28 As Double = 1.94060256838287E-07 + (-2.64697796016969E-23)
    Const C29 As Double = -1.25229209086299E-07 + (-2.11758236813575E-22)
    Const C30 As Double = 8.08948668443836E-08 + (1.32348898008484E-23)
    Const C31 As Double = -5.23064498832708E-08 + (3.97046694025453E-23)
    Const C32 As Double = 3.38519414938090E-08 + (3.97046694025453E-23)

    U = A - 0.5
    Acc = C32
    Acc = Acc * U + C31
    Acc = Acc * U + C30
    Acc = Acc * U + C29
    Acc = Acc * U + C28
    Acc = Acc * U + C27
    Acc = Acc * U + C26
    Acc = Acc * U + C25
    Acc = Acc * U + C24
    Acc = Acc * U + C23
    Acc = Acc * U + C22
    Acc = Acc * U + C21
    Acc = Acc * U + C20
    Acc = Acc * U + C19
    Acc = Acc * U + C18
    Acc = Acc * U + C17
    Acc = Acc * U + C16
    Acc = Acc * U + C15
    Acc = Acc * U + C14
    Acc = Acc * U + C13
    Acc = Acc * U + C12
    Acc = Acc * U + C11
    Acc = Acc * U + C10
    Acc = Acc * U + C9
    Acc = Acc * U + C8
    Acc = Acc * U + C7
    Acc = Acc * U + C6
    Acc = Acc * U + C5
    Acc = Acc * U + C4
    Acc = Acc * U + C3
    Acc = Acc * U + C2
    Acc = Acc * U + C1
    Acc = Acc * U + C0
    Product = A * (A - 1#)
    PROB_GammaSmallLogGamma1p = Product * Acc
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
'     FALSE => Iteration cap or exponential failure; Result is not contractual.
' BEHAVIOR
'   For 0 < A <= 1, stop at exact 2^-52 and evaluate the same finite fraction
'   backward with the B6 split prefactor. For A > 1, retain PROB_NUM_EPS and
'   the existing forward recurrence and prefactor.
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
    Dim BackIdx             As Long            'Backward finite-fraction index
    Dim Tail                As Double          'Backward accumulator
    Dim Epsilon             As Double          'Route-specific stop
    Dim LogGamma1p          As Double          'Small-shape log gamma
    Dim Exponent            As Double          'Unscaled exponential argument
    Dim Correction          As Double          'Positive correction factor

'------------------------------------------------------------------------------
' INITIALIZE
'------------------------------------------------------------------------------
    'Initialize the continued fraction
        Epsilon = PROB_NUM_EPS
        If A <= 1# Then Epsilon = 2# ^ (-52)
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
                If Abs(Del - 1#) <= Epsilon Then
                    'B6 re-evaluates exactly this finite fraction, without extra terms.
                    If A <= 1# Then
                        Tail = 0#
                        For BackIdx = IterIdx To 1 Step -1
                            Tail = (-CDbl(BackIdx) * (CDbl(BackIdx) - A)) / _
                                (X + 1# - A + 2# * CDbl(BackIdx) + Tail)
                        Next BackIdx
                        h = 1# / (X + 1# - A + Tail)

                        'Q(A,X) <= Exp(-X) for A<=1; exponential underflow is safe.
                        If Not PROB_TryExp(-X, Factor) Then
                            FailMsg = "Small-shape gamma exponential failed"
                            Exit Function
                        End If
                        If Factor = 0# Then
                            Result = 0#
                            PROB_TryGammaContinuedFractionQ = True
                            Exit Function
                        End If

                        'Preserve the tested Double stores and split exponential.
                        LogGamma1p = PROB_GammaSmallLogGamma1p(A)
                        Exponent = A * Log(X) - LogGamma1p
                        If Not PROB_TryExp(Exponent, Correction) Then
                            FailMsg = "Small-shape gamma correction overflowed"
                            Exit Function
                        End If
                        Factor = A * Factor
                        Factor = Factor * Correction
                        Result = Factor * h
                        PROB_TryGammaContinuedFractionQ = True
                        Exit Function
                    End If
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
'     FALSE => an inner incomplete gamma evaluation failed, Newton/bisection
'              did not settle within PROB_INV_MAX_ITER, or Probability = 1
'              (the quantile is unbounded); FailMsg says which.
'
' DEPENDENCIES
'   - PROB_NormalInvCDFRaw, PROB_LogGamma, PROB_TryExp
'   - PROB_TryGammaRegularizedP, PROB_TryGammaRegularizedQ
'   - PROB_TryGammaLogPdf
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


'==============================================================================
' STANDARD NORMAL UPPER TAIL
'==============================================================================

Public Function PROB_NormalUpperTail( _
    ByVal Z As Double) _
    As Double
'
'==============================================================================
' PROB_NormalUpperTail
'------------------------------------------------------------------------------
' PURPOSE
'   Returns Q(Z) = P(N(0,1) > Z) for Z >= 0 without subtracting from one, to
'   full relative precision wherever Q(Z) is a normal Double.
'
' WHY THIS EXISTS
'   The normal family's previous kernel combined a Hart rational approximation
'   below Z = 7.07 with a sixteen-level continued fraction above it. It was
'   accurate in absolute terms only, and its exponential factor rounded Z ^ 2
'   first (#72). The Student-t large-df route (#34) needs the same tail to full
'   relative precision, so the repaired kernel lives here, once, for both.
'
' METHOD
'   Q(Z) = E(Z) * F(Z), with E(Z) = Exp(-Z ^ 2 / 2) from PROB_ExpNegHalfSquare,
'   whose main argument is exact, and F(Z) = Q(Z) * Exp(Z ^ 2 / 2):
'
'     0 <= Z < 3      F is the rational P(Z) / D(Z), each evaluated by Horner
'                     in Z. Every coefficient but one is positive, so there is
'                     no cancellation to speak of.
'     3 <= Z < 38.5   F = (1 / Sqr(2 * Pi)) / d, where d is the Laplace
'                     continued fraction Z + 1/(Z + 2/(Z + 3/(Z + ...))),
'                     evaluated backwards over 60 levels.
'     Z >= 38.5       0: the one-sided tail rounds to zero in Double.
'
'   Every operation is written in the order the binary64 model in
'   benchmark/normal_tail_study evaluates it. Do not reassociate.
'
' ACCURACY
'   Not stated here. The binary64 model of this exact arithmetic, its
'   validation and holdout, and the predictions for every affected public
'   function are recorded in benchmark/normal_tail_study. Contracts are in
'   benchmark/accuracy_contracts.csv and change only with Excel evidence.
'
' PRECONDITION
'   Z >= 0. Callers reflect negative arguments.
'
' UPDATED
'   2026-10-13 - New (#72); replaces PROB_NormalUpperTailPositive in
'                M_STATS_PROBDIST_NORMALFAMILY.
'==============================================================================
'
'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Dim Exponential         As Double          'Exp(-Z ^ 2 / 2), split exponent
    Dim SumP                As Double          'Rational numerator accumulator
    Dim SumD                As Double          'Rational denominator accumulator
    Dim Denominator         As Double          'Continued-fraction denominator
    Dim Level               As Long            'Continued-fraction level

'------------------------------------------------------------------------------
' HANDLE TRUE DOUBLE UNDERFLOW
'------------------------------------------------------------------------------
    'The one-sided tail rounds to zero at approximately Z = 38.49
        If Z >= PROB_NT_UNDERFLOW_Z Then
            PROB_NormalUpperTail = 0#
            Exit Function
        End If

'------------------------------------------------------------------------------
' CALCULATE COMMON EXPONENTIAL FACTOR
'------------------------------------------------------------------------------
    'The underflow cutoff above keeps Z inside the helper's precondition
        Exponential = PROB_ExpNegHalfSquare(Z)

'------------------------------------------------------------------------------
' CENTRAL REGION
'------------------------------------------------------------------------------
    'Rational approximation to Q(Z) * Exp(Z ^ 2 / 2) below the split
        If Z < PROB_NT_SPLIT Then
            SumP = PROB_NT_P7 * Z + PROB_NT_P6
            SumP = SumP * Z + PROB_NT_P5
            SumP = SumP * Z + PROB_NT_P4
            SumP = SumP * Z + PROB_NT_P3
            SumP = SumP * Z + PROB_NT_P2
            SumP = SumP * Z + PROB_NT_P1
            SumP = SumP * Z + PROB_NT_P0

            SumD = PROB_NT_D7 * Z + PROB_NT_D6
            SumD = SumD * Z + PROB_NT_D5
            SumD = SumD * Z + PROB_NT_D4
            SumD = SumD * Z + PROB_NT_D3
            SumD = SumD * Z + PROB_NT_D2
            SumD = SumD * Z + PROB_NT_D1
            SumD = SumD * Z + 1#

            PROB_NormalUpperTail = Exponential * (SumP / SumD)

'------------------------------------------------------------------------------
' TAIL
'------------------------------------------------------------------------------
        Else
        'Evaluate the Laplace continued fraction backwards, then scale by the
        'Mills ratio phi(Z) / Denominator
            Denominator = Z

            For Level = PROB_NT_CF_LEVELS To 1 Step -1
                Denominator = Z + CDbl(Level) / Denominator
            Next Level

            PROB_NormalUpperTail = _
                Exponential * (PROB_NT_INV_SQRT_TWO_PI / Denominator)
        End If
End Function


' BEGIN generated public Gamma integration
Private Function PROB_DD_SP_FromWords(ByVal HighWord As Long, ByVal LowWord As Long) As Double
'==============================================================================
' PURPOSE
'   PROB_DD_SP_FromWords: Gamma exponent-separated arithmetic.
' NOTES
'   Explicit Double stores; Windows little-endian word fixtures.
'   Operation graph retained from frozen Excel port b8efaf3.
' ERROR POLICY
'   Try functions return False; test failures remain visible.
' UPDATED
'   2026-11-08
'==============================================================================
    Dim Words As PROB_DD_Words, Number As PROB_DD_Number
    Words.High = HighWord
    Words.Low = LowWord
    LSet Number = Words
    PROB_DD_SP_FromWords = Number.Value
End Function

Private Function PROB_DD_SP_Negative(ByVal A As Double) As Boolean
'==============================================================================
' PURPOSE
'   PROB_DD_SP_Negative: Gamma exponent-separated arithmetic.
' NOTES
'   Explicit Double stores; Windows little-endian word fixtures.
'   Operation graph retained from frozen Excel port b8efaf3.
' ERROR POLICY
'   Try functions return False; test failures remain visible.
' UPDATED
'   2026-11-08
'==============================================================================
    Dim Number As PROB_DD_Number, Words As PROB_DD_Words
    Number.Value = A
    LSet Words = Number
    PROB_DD_SP_Negative = (Words.High < 0)
End Function

Private Function PROB_DD_SP_SignedZero(ByVal Negative As Boolean) As Double
'==============================================================================
' PURPOSE
'   PROB_DD_SP_SignedZero: Gamma exponent-separated arithmetic.
' NOTES
'   Explicit Double stores; Windows little-endian word fixtures.
'   Operation graph retained from frozen Excel port b8efaf3.
' ERROR POLICY
'   Try functions return False; test failures remain visible.
' UPDATED
'   2026-11-08
'==============================================================================
    If Negative Then PROB_DD_SP_SignedZero = PROB_DD_SP_FromWords(&H80000000, 0)
End Function

Private Function PROB_DD_SP_Shift(ByVal X As Double, ByVal Exponent As Long) As Double
'==============================================================================
' PURPOSE
'   PROB_DD_SP_Shift: Gamma exponent-separated arithmetic.
' NOTES
'   Explicit Double stores; Windows little-endian word fixtures.
'   Operation graph retained from frozen Excel port b8efaf3.
' ERROR POLICY
'   Try functions return False; test failures remain visible.
' UPDATED
'   2026-11-08
'==============================================================================
    Dim J As Long
    If Exponent < -1074 Or Exponent > 1024 Then Err.Raise 5
    If Exponent >= 0 Then
        For J = 1 To Exponent
            X = X * 2#
        Next J
    Else
        For J = 1 To -Exponent
            X = X * 0.5
        Next J
    End If
    PROB_DD_SP_Shift = X
End Function

Private Sub PROB_DD_SP_Frexp(ByVal X As Double, ByRef Mantissa As Double, ByRef Exponent As Long)
'==============================================================================
' PURPOSE
'   PROB_DD_SP_Frexp: Gamma exponent-separated arithmetic.
' NOTES
'   Explicit Double stores; Windows little-endian word fixtures.
'   Operation graph retained from frozen Excel port b8efaf3.
' ERROR POLICY
'   Try functions return False; test failures remain visible.
' UPDATED
'   2026-11-08
'==============================================================================
    Exponent = 0
    Mantissa = X
    If X = 0# Then Exit Sub
    Do While Abs(Mantissa) >= 1#
        Mantissa = Mantissa * 0.5
        Exponent = Exponent + 1
    Loop
    Do While Abs(Mantissa) < 0.5
        Mantissa = Mantissa * 2#
        Exponent = Exponent - 1
    Loop
End Sub

Private Function PROB_DD_SP_TwoSum(ByVal A As Double, ByVal B As Double) As Variant
'==============================================================================
' PURPOSE
'   PROB_DD_SP_TwoSum: Gamma exponent-separated arithmetic.
' NOTES
'   Explicit Double stores; Windows little-endian word fixtures.
'   Operation graph retained from frozen Excel port b8efaf3.
' ERROR POLICY
'   Try functions return False; test failures remain visible.
' UPDATED
'   2026-11-08
'==============================================================================
    Dim S As Double, V As Double, AV As Double, AR As Double, BR As Double
    S = A + B
    V = S - A
    AV = S - V
    AR = A - AV
    BR = B - V
    PROB_DD_SP_TwoSum = Array(S, AR + BR)
End Function

Private Function PROB_DD_SP_Product(ByVal A As Double, ByVal B As Double, _
    ByRef H As Double, ByRef L As Double, ByRef Exponent As Long) As Boolean
'==============================================================================
' PURPOSE
'   PROB_DD_SP_Product: Gamma exponent-separated arithmetic.
' NOTES
'   Explicit Double stores; Windows little-endian word fixtures.
'   Operation graph retained from frozen Excel port b8efaf3.
' ERROR POLICY
'   Try functions return False; test failures remain visible.
' UPDATED
'   2026-11-08
'==============================================================================
    Dim AM As Double, BM As Double, AE As Long, BE As Long, Shift As Long
    Dim CA As Double, CB As Double, AH As Double, BH As Double
    Dim AL As Double, BL As Double, P As Double, E As Double, T As Double
    On Error GoTo Err_Handler
    H = 0#: L = 0#: Exponent = 0
    If Not PROB_IsFinite(A) Then Exit Function
    If Not PROB_IsFinite(B) Then Exit Function
    If A = 0# Or B = 0# Then
        H = PROB_DD_SP_SignedZero(PROB_DD_SP_Negative(A) Xor PROB_DD_SP_Negative(B))
        PROB_DD_SP_Product = True
        Exit Function
    End If
    PROB_DD_SP_Frexp A, AM, AE
    PROB_DD_SP_Frexp B, BM, BE
    P = AM * BM
    CA = 134217729# * AM
    CB = 134217729# * BM
    T = CA - AM
    AH = CA - T
    T = CB - BM
    BH = CB - T
    AL = AM - AH
    BL = BM - BH
    T = AH * BH
    E = T - P
    T = AH * BL
    E = E + T
    T = AL * BH
    E = E + T
    T = AL * BL
    E = E + T
    PROB_DD_SP_Frexp P, H, Shift
    L = PROB_DD_SP_Shift(E, -Shift)
    Exponent = AE + BE + Shift
    PROB_DD_SP_Product = True
    Exit Function
Err_Handler:
    H = 0#: L = 0#: Exponent = 0
    PROB_DD_SP_Product = False
End Function

Private Function PROB_DD_SGD_Try(ByVal Operation As Long, _
    ByVal AH As Double, ByVal AL As Double, ByVal AE As Long, _
    ByVal BH As Double, ByVal BL As Double, ByVal BE As Long, _
    ByRef H As Double, ByRef L As Double, ByRef E As Long, Optional ByRef Diagnostic As String = "") As Boolean
'==============================================================================
' PURPOSE
'   PROB_DD_SGD_Try: Gamma exponent-separated arithmetic.
' NOTES
'   Explicit Double stores; Windows little-endian word fixtures.
'   Operation graph retained from frozen Excel port b8efaf3.
' ERROR POLICY
'   Try functions return False; test failures remain visible.
' UPDATED
'   2026-11-08
'==============================================================================
    Dim Pair As Variant, Gap As Long, Shift As Long, Factor As Double
    Dim T As Double, TE As Long, RawH As Double, RawL As Double
    On Error GoTo Err_Handler
    H = 0#: L = 0#: E = 0
    If Not PROB_DD_SGD_Valid(AH, AL, AE) Then
        Diagnostic = "input A H=" & PROB_DD_SG_Bits(AH) & " L=" & PROB_DD_SG_Bits(AL) & " sum=" & PROB_DD_SG_Bits(AH + AL)
        Exit Function
    End If
    If Not PROB_DD_SGD_Valid(BH, BL, BE) Then
        Diagnostic = "input B H=" & PROB_DD_SG_Bits(BH) & " L=" & PROB_DD_SG_Bits(BL) & " sum=" & PROB_DD_SG_Bits(BH + BL)
        Exit Function
    End If
    If Operation < 1 Or Operation > 4 Then Exit Function
    If Operation = 2 Then
        BH = -BH: BL = -BL
        Operation = 1
    End If
    If Operation = 1 Then
        If AH = 0# Then
            If BH = 0# Then
                H = PROB_DD_SP_SignedZero(PROB_DD_SP_Negative(AH) And PROB_DD_SP_Negative(BH))
            Else
                H = BH: L = BL: E = BE
            End If
            PROB_DD_SGD_Try = True
            Exit Function
        End If
        If BH = 0# Then
            H = AH: L = AL: E = AE
            PROB_DD_SGD_Try = True
            Exit Function
        End If
        If AE < BE Then
            T = AH: AH = BH: BH = T
            T = AL: AL = BL: BL = T
            TE = AE: AE = BE: BE = TE
        End If
        Gap = AE - BE
        If Gap > 120 Then
            H = AH: L = AL: E = AE
            PROB_DD_SGD_Try = True
            Exit Function
        End If
        'One multiplication rounds a tiny aligned low component only once.
        Factor = PROB_DD_SP_Shift(1#, -Gap)
        BH = BH * Factor
        BL = BL * Factor
        Pair = PROB_DD_PV2_Add(Array(AH, AL), Array(BH, BL))
        E = AE
    Else
        If Operation = 4 And BH = 0# Then Exit Function
        If AH = 0# Or BH = 0# Then
            H = PROB_DD_SP_SignedZero(PROB_DD_SP_Negative(AH) Xor PROB_DD_SP_Negative(BH))
            PROB_DD_SGD_Try = True
            Exit Function
        End If
        If Operation = 3 Then
            Pair = PROB_DD_PV2_Mul(Array(AH, AL), Array(BH, BL))
            E = AE + BE
        Else
            Pair = PROB_DD_PV2_Div(Array(AH, AL), Array(BH, BL))
            E = AE - BE
        End If
    End If
    Pair = PROB_DD_PV2_Sum(CDbl(Pair(0)), CDbl(Pair(1)))
    RawH = Pair(0): RawL = Pair(1)
    If RawH = 0# Then
        H = RawH: L = 0#: E = 0
    Else
        PROB_DD_SP_Frexp RawH, H, Shift
        'If RawH is subnormal, the nonoverlapping RawL must be zero.
        If RawL <> 0# Then
            Factor = PROB_DD_SP_Shift(1#, -Shift)
            L = RawL * Factor
        End If
        E = E + Shift
    End If
    PROB_DD_SGD_Canonicalize H, L, E
    If Not PROB_DD_SGD_Valid(H, L, E) Then
        Diagnostic = "output H=" & PROB_DD_SG_Bits(H) & " L=" & PROB_DD_SG_Bits(L) & " E=" & CStr(E) & " sum=" & PROB_DD_SG_Bits(H + L)
        GoTo Err_Handler
    End If
    PROB_DD_SGD_Try = True
    Exit Function
Err_Handler:
    If Len(Diagnostic) = 0 Then Diagnostic = "runtime " & Err.Description
    H = 0#: L = 0#: E = 0
    PROB_DD_SGD_Try = False
End Function

Private Function PROB_DD_SGD_Round(ByVal H As Double, ByVal L As Double, _
    ByVal Exponent As Long, ByRef Result As Double) As Boolean
'==============================================================================
' PURPOSE
'   PROB_DD_SGD_Round: Gamma exponent-separated arithmetic.
' NOTES
'   Explicit Double stores; Windows little-endian word fixtures.
'   Operation graph retained from frozen Excel port b8efaf3.
' ERROR POLICY
'   Try functions return False; test failures remain visible.
' UPDATED
'   2026-11-08
'==============================================================================
    Dim Units As Double, LowUnits As Double, Nearest As Double
    Dim SumValue As Double, R As Double, Residual As Double, SignValue As Double
    Dim Pair As Variant, Odd As Boolean
    On Error GoTo Err_Handler
    Result = 0#
    If Not PROB_IsFinite(H) Then Exit Function
    If Not PROB_IsFinite(L) Then Exit Function
    If H = 0# Then
        If L <> 0# Then Exit Function
        Result = H
        PROB_DD_SGD_Round = True
        Exit Function
    End If
    If Not PROB_DD_SGD_Valid(H, L, Exponent) Then Exit Function
    If Exponent > 1024 Then Exit Function
    If Exponent >= -1021 Then
        Result = PROB_DD_SP_Shift(H, Exponent)
        If Not PROB_IsFinite(Result) Then GoTo Err_Handler
        PROB_DD_SGD_Round = True
        Exit Function
    End If
    If Exponent < -1074 Then
        Result = PROB_DD_SP_SignedZero(H < 0#)
        PROB_DD_SGD_Round = True
        Exit Function
    End If
    SignValue = 1#
    If H < 0# Then SignValue = -1#
    Units = PROB_DD_SP_Shift(Abs(H), Exponent + 1074)
    LowUnits = PROB_DD_SP_Shift(SignValue * L, Exponent + 1074)
    Nearest = VBA.Round(Units, 0)
    R = Units - Nearest
    Pair = PROB_DD_SP_TwoSum(R, LowUnits)
    R = Pair(0)
    Residual = Pair(1)
    'Nearest can exceed Long: parity is evaluated with exact Double arithmetic.
    Odd = (Nearest - 2# * Fix(Nearest * 0.5) <> 0#)
    If R > 0.5 Or (R = 0.5 And (Residual > 0# Or (Residual = 0# And Odd))) Then
        Nearest = Nearest + 1#
    ElseIf R < -0.5 Or (R = -0.5 And (Residual < 0# Or (Residual = 0# And Odd))) Then
        Nearest = Nearest - 1#
    End If
    If Nearest = 0# Then
        Result = PROB_DD_SP_SignedZero(H < 0#)
    Else
        Result = PROB_DD_SP_Shift(Nearest, -1074)
        Result = SignValue * Result
    End If
    PROB_DD_SGD_Round = True
    Exit Function
Err_Handler:
    Result = 0#
    PROB_DD_SGD_Round = False
End Function

Private Function PROB_DD_SGD_Half(ByVal H As Double, ByVal L As Double) As Double
'==============================================================================
' PURPOSE
'   PROB_DD_SGD_Half: Gamma exponent-separated arithmetic.
' NOTES
'   Explicit Double stores; Windows little-endian word fixtures.
'   Operation graph retained from frozen Excel port b8efaf3.
' ERROR POLICY
'   Try functions return False; test failures remain visible.
' UPDATED
'   2026-11-08
'==============================================================================
    If Abs(H) = 0.5 And ((H > 0# And L < 0#) Or (H < 0# And L > 0#)) Then
        PROB_DD_SGD_Half = PROB_DD_SP_FromWords(&H3C800000&, 0)
    Else
        PROB_DD_SGD_Half = PROB_DD_SP_FromWords(&H3C900000&, 0)
    End If
End Function

Private Function PROB_DD_SGD_Even(ByVal H As Double) As Boolean
'==============================================================================
' PURPOSE
'   PROB_DD_SGD_Even: Gamma exponent-separated arithmetic.
' NOTES
'   Explicit Double stores; Windows little-endian word fixtures.
'   Operation graph retained from frozen Excel port b8efaf3.
' ERROR POLICY
'   Try functions return False; test failures remain visible.
' UPDATED
'   2026-11-08
'==============================================================================
    Dim N As PROB_DD_Number, W As PROB_DD_Words
    N.Value = H
    LSet W = N
    PROB_DD_SGD_Even = ((W.Low And 1&) = 0)
End Function

Private Function PROB_DD_SGD_Valid(ByVal H As Double, ByVal L As Double, ByVal E As Long) As Boolean
'==============================================================================
' PURPOSE
'   PROB_DD_SGD_Valid: Gamma exponent-separated arithmetic.
' NOTES
'   Explicit Double stores; Windows little-endian word fixtures.
'   Operation graph retained from frozen Excel port b8efaf3.
' ERROR POLICY
'   Try functions return False; test failures remain visible.
' UPDATED
'   2026-11-08
'==============================================================================
    Dim Half As Double
    If E < -32767 Or E > 32767 Then Exit Function
    If Not PROB_IsFinite(H) Or Not PROB_IsFinite(L) Then Exit Function
    If H = 0# Then
        PROB_DD_SGD_Valid = (L = 0# And E = 0)
        Exit Function
    End If
    If Abs(H) < 0.5 Or Abs(H) >= 1# Then Exit Function
    Half = PROB_DD_SGD_Half(H, L)
    PROB_DD_SGD_Valid = (Abs(L) < Half Or (Abs(L) = Half And PROB_DD_SGD_Even(H)))
End Function

Private Sub PROB_DD_SGD_Canonicalize(ByRef H As Double, ByRef L As Double, ByRef E As Long)
'==============================================================================
' PURPOSE
'   PROB_DD_SGD_Canonicalize: Gamma exponent-separated arithmetic.
' NOTES
'   Explicit Double stores; Windows little-endian word fixtures.
'   Operation graph retained from frozen Excel port b8efaf3.
' ERROR POLICY
'   Try functions return False; test failures remain visible.
' UPDATED
'   2026-11-08
'==============================================================================
    Dim Half As Double, N As PROB_DD_Number, W As PROB_DD_Words, OldH As Double, StepValue As Double
    Dim IncreaseMagnitude As Boolean
    If H = 0# Then Exit Sub
    Half = PROB_DD_SGD_Half(H, L)
    If Abs(L) < Half Or (Abs(L) = Half And PROB_DD_SGD_Even(H)) Then Exit Sub
    If Abs(L) > 2# * Half Then Err.Raise 5, "PROB_DD_SGD_Canonicalize", "residual exceeds repair range"
    OldH = H
    IncreaseMagnitude = ((H > 0# And L > 0#) Or (H < 0# And L < 0#))
    N.Value = H
    LSet W = N
    If IncreaseMagnitude Then
        If W.Low = -1 Then
            W.Low = 0: W.High = W.High + 1
        ElseIf W.Low = &H7FFFFFFF Then
            W.Low = &H80000000
        Else
            W.Low = W.Low + 1
        End If
    Else
        If W.Low = 0 Then
            W.Low = -1: W.High = W.High - 1
        ElseIf W.Low = &H80000000 Then
            W.Low = &H7FFFFFFF
        Else
            W.Low = W.Low - 1
        End If
    End If
    LSet N = W
    H = N.Value
    StepValue = H - OldH
    L = L - StepValue
    If Abs(H) >= 1# Then
        H = H * 0.5: L = L * 0.5: E = E + 1
    ElseIf Abs(H) < 0.5 Then
        H = H * 2#: L = L * 2#: E = E - 1
    End If
End Sub

Private Function PROB_DD_SG_Bits(ByVal X As Double) As String
'==============================================================================
' PURPOSE
'   PROB_DD_SG_Bits: Gamma exponent-separated arithmetic.
' NOTES
'   Explicit Double stores; Windows little-endian word fixtures.
'   Operation graph retained from frozen Excel port b8efaf3.
' ERROR POLICY
'   Try functions return False; test failures remain visible.
' UPDATED
'   2026-11-08
'==============================================================================
    Dim N As PROB_DD_Number, W As PROB_DD_Words
    N.Value = X
    LSet W = N
    PROB_DD_SG_Bits = Hex$(W.High) & ":" & Hex$(W.Low)
End Function

Private Function PROB_DD_SG_Coerce(ByVal V As Variant) As Variant
'==============================================================================
' PURPOSE
'   PROB_DD_SG_Coerce: Gamma exponent-separated arithmetic.
' NOTES
'   Explicit Double stores; Windows little-endian word fixtures.
'   Operation graph retained from frozen Excel port b8efaf3.
' ERROR POLICY
'   Try functions return False; test failures remain visible.
' UPDATED
'   2026-11-08
'==============================================================================
    Dim H As Double, E As Long
    If IsArray(V) Then
        PROB_DD_SG_Coerce = V
    Else
        If Not PROB_IsFinite(CDbl(V)) Then Err.Raise 5
        PROB_DD_SP_Frexp CDbl(V), H, E
        PROB_DD_SG_Coerce = Array(H, 0#, E)
    End If
End Function

Private Function PROB_DD_SG_Op(ByVal Operation As Long, ByVal A As Variant, ByVal B As Variant) As Variant
'==============================================================================
' PURPOSE
'   PROB_DD_SG_Op: Gamma exponent-separated arithmetic.
' NOTES
'   Explicit Double stores; Windows little-endian word fixtures.
'   Operation graph retained from frozen Excel port b8efaf3.
' ERROR POLICY
'   Try functions return False; test failures remain visible.
' UPDATED
'   2026-11-08
'==============================================================================
    Dim U As Variant, V As Variant, H As Double, L As Double, E As Long, Fault As String
    U = PROB_DD_SG_Coerce(A): V = PROB_DD_SG_Coerce(B)
    If Not PROB_DD_SGD_Try(Operation, CDbl(U(0)), CDbl(U(1)), CLng(U(2)), _
        CDbl(V(0)), CDbl(V(1)), CLng(V(2)), H, L, E, Fault) Then
        Err.Raise 5, "PROB_DD_SG_Op", Fault & " operation " & CStr(Operation) & " A=" & CStr(U(0)) & ";" & CStr(U(1)) & ";" & CStr(U(2)) & _
            " B=" & CStr(V(0)) & ";" & CStr(V(1)) & ";" & CStr(V(2))
    End If
    PROB_DD_SG_Op = Array(H, L, E)
End Function

Private Function PROB_DD_SG_Add(ByVal A As Variant, ByVal B As Variant) As Variant
'==============================================================================
' PURPOSE
'   PROB_DD_SG_Add: Gamma exponent-separated arithmetic.
' NOTES
'   Explicit Double stores; Windows little-endian word fixtures.
'   Operation graph retained from frozen Excel port b8efaf3.
' ERROR POLICY
'   Try functions return False; test failures remain visible.
' UPDATED
'   2026-11-08
'==============================================================================
    PROB_DD_SG_Add = PROB_DD_SG_Op(1, A, B)
End Function

Private Function PROB_DD_SG_Sub(ByVal A As Variant, ByVal B As Variant) As Variant
'==============================================================================
' PURPOSE
'   PROB_DD_SG_Sub: Gamma exponent-separated arithmetic.
' NOTES
'   Explicit Double stores; Windows little-endian word fixtures.
'   Operation graph retained from frozen Excel port b8efaf3.
' ERROR POLICY
'   Try functions return False; test failures remain visible.
' UPDATED
'   2026-11-08
'==============================================================================
    PROB_DD_SG_Sub = PROB_DD_SG_Op(2, A, B)
End Function

Private Function PROB_DD_SG_Mul(ByVal A As Variant, ByVal B As Variant) As Variant
'==============================================================================
' PURPOSE
'   PROB_DD_SG_Mul: Gamma exponent-separated arithmetic.
' NOTES
'   Explicit Double stores; Windows little-endian word fixtures.
'   Operation graph retained from frozen Excel port b8efaf3.
' ERROR POLICY
'   Try functions return False; test failures remain visible.
' UPDATED
'   2026-11-08
'==============================================================================
    PROB_DD_SG_Mul = PROB_DD_SG_Op(3, A, B)
End Function

Private Function PROB_DD_SG_Div(ByVal A As Variant, ByVal B As Variant) As Variant
'==============================================================================
' PURPOSE
'   PROB_DD_SG_Div: Gamma exponent-separated arithmetic.
' NOTES
'   Explicit Double stores; Windows little-endian word fixtures.
'   Operation graph retained from frozen Excel port b8efaf3.
' ERROR POLICY
'   Try functions return False; test failures remain visible.
' UPDATED
'   2026-11-08
'==============================================================================
    PROB_DD_SG_Div = PROB_DD_SG_Op(4, A, B)
End Function

Private Function PROB_DD_SG_Neg(ByVal A As Variant) As Variant
'==============================================================================
' PURPOSE
'   PROB_DD_SG_Neg: Gamma exponent-separated arithmetic.
' NOTES
'   Explicit Double stores; Windows little-endian word fixtures.
'   Operation graph retained from frozen Excel port b8efaf3.
' ERROR POLICY
'   Try functions return False; test failures remain visible.
' UPDATED
'   2026-11-08
'==============================================================================
    PROB_DD_SG_Neg = Array(-CDbl(A(0)), -CDbl(A(1)), CLng(A(2)))
End Function

Private Function PROB_DD_SG_Value(ByVal A As Variant) As Double
'==============================================================================
' PURPOSE
'   PROB_DD_SG_Value: Gamma exponent-separated arithmetic.
' NOTES
'   Explicit Double stores; Windows little-endian word fixtures.
'   Operation graph retained from frozen Excel port b8efaf3.
' ERROR POLICY
'   Try functions return False; test failures remain visible.
' UPDATED
'   2026-11-08
'==============================================================================
    Dim Result As Double
    If Not PROB_DD_SGD_Round(CDbl(A(0)), CDbl(A(1)), CLng(A(2)), Result) Then Err.Raise 5, "PROB_DD_SG_Value", "rounding refused"
    PROB_DD_SG_Value = Result
End Function

Private Function PROB_DD_SG_Small(ByVal Term As Variant, ByVal Total As Variant) As Boolean
'==============================================================================
' PURPOSE
'   PROB_DD_SG_Small: Gamma exponent-separated arithmetic.
' NOTES
'   Explicit Double stores; Windows little-endian word fixtures.
'   Operation graph retained from frozen Excel port b8efaf3.
' ERROR POLICY
'   Try functions return False; test failures remain visible.
' UPDATED
'   2026-11-08
'==============================================================================
    PROB_DD_SG_Small = (CDbl(Term(0)) = 0# Or _
        (CDbl(Total(0)) <> 0# And CLng(Term(2)) < CLng(Total(2)) - 108))
End Function

Private Function PROB_DD_SG_Ln2() As Variant
'==============================================================================
' PURPOSE
'   PROB_DD_SG_Ln2: Gamma exponent-separated arithmetic.
' NOTES
'   Explicit Double stores; Windows little-endian word fixtures.
'   Operation graph retained from frozen Excel port b8efaf3.
' ERROR POLICY
'   Try functions return False; test failures remain visible.
' UPDATED
'   2026-11-08
'==============================================================================
    PROB_DD_SG_Ln2 = Array(PROB_DD_SP_FromWords(&H3FE62E42&, &HFEFA39EF&), PROB_DD_SP_FromWords(&H3C7ABC9E&, &H3B39803F&), 0)
End Function

Private Function PROB_DD_SG_Log(ByVal X As Double) As Variant
'==============================================================================
' PURPOSE
'   PROB_DD_SG_Log: Gamma exponent-separated arithmetic.
' NOTES
'   Explicit Double stores; Windows little-endian word fixtures.
'   Operation graph retained from frozen Excel port b8efaf3.
' ERROR POLICY
'   Try functions return False; test failures remain visible.
' UPDATED
'   2026-11-08
'==============================================================================
    Dim F As Double, E As Long, Z As Variant, Square As Variant
    Dim Term As Variant, Total As Variant, Delta As Variant, N As Long
    If Not PROB_IsFinite(X) Then Err.Raise 5
    If X <= 0# Then Err.Raise 5
    PROB_DD_SP_Frexp X, F, E
    Z = PROB_DD_SG_Div(PROB_DD_SG_Sub(F, 1#), PROB_DD_SG_Add(F, 1#))
    Square = PROB_DD_SG_Mul(Z, Z)
    Term = Z: Total = Z
    For N = 3 To 1000 Step 2
        Term = PROB_DD_SG_Mul(Term, Square)
        Delta = PROB_DD_SG_Div(Term, N)
        Total = PROB_DD_SG_Add(Total, Delta)
        If PROB_DD_SG_Small(Delta, Total) Then
            PROB_DD_SG_Log = PROB_DD_SG_Add(PROB_DD_SG_Mul(Total, 2#), PROB_DD_SG_Mul(PROB_DD_SG_Ln2(), E))
            Exit Function
        End If
    Next N
    Err.Raise 5
End Function

Private Function PROB_DD_SG_Exp(ByVal Z As Variant) As Variant
'==============================================================================
' PURPOSE
'   PROB_DD_SG_Exp: Gamma exponent-separated arithmetic.
' NOTES
'   Explicit Double stores; Windows little-endian word fixtures.
'   Operation graph retained from frozen Excel port b8efaf3.
' ERROR POLICY
'   Try functions return False; test failures remain visible.
' UPDATED
'   2026-11-08
'==============================================================================
    Dim Value As Double, N As Long, R As Variant, Term As Variant, Total As Variant, K As Long
    Value = PROB_DD_SG_Value(Z)
    If Value < -1500# Or Value > 8# Then Err.Raise 5
    N = CLng(Round(Value / PROB_DD_SP_FromWords(&H3FE62E42&, &HFEFA39EF&), 0))
    R = PROB_DD_SG_Sub(Z, PROB_DD_SG_Mul(PROB_DD_SG_Ln2(), N))
    Term = PROB_DD_SG_Coerce(1#): Total = Term
    For K = 1 To 1000
        Term = PROB_DD_SG_Div(PROB_DD_SG_Mul(Term, R), K)
        Total = PROB_DD_SG_Add(Total, Term)
        If PROB_DD_SG_Small(Term, Total) Then
            PROB_DD_SG_Exp = Array(CDbl(Total(0)), CDbl(Total(1)), CLng(Total(2)) + N)
            Exit Function
        End If
    Next K
    Err.Raise 5
End Function

Private Function PROB_DD_SG_Expm1(ByVal Z As Variant) As Variant
'==============================================================================
' PURPOSE
'   PROB_DD_SG_Expm1: Gamma exponent-separated arithmetic.
' NOTES
'   Explicit Double stores; Windows little-endian word fixtures.
'   Operation graph retained from frozen Excel port b8efaf3.
' ERROR POLICY
'   Try functions return False; test failures remain visible.
' UPDATED
'   2026-11-08
'==============================================================================
    Dim Term As Variant, Total As Variant, K As Long
    If Abs(PROB_DD_SG_Value(Z)) > 0.5 Then Err.Raise 5
    Term = Z: Total = Z
    For K = 2 To 1000
        Term = PROB_DD_SG_Div(PROB_DD_SG_Mul(Term, Z), K)
        Total = PROB_DD_SG_Add(Total, Term)
        If PROB_DD_SG_Small(Term, Total) Then
            PROB_DD_SG_Expm1 = Total
            Exit Function
        End If
    Next K
    Err.Raise 5
End Function

Private Function PROB_DD_SG_Log1p(ByVal Z As Variant) As Variant
'==============================================================================
' PURPOSE
'   PROB_DD_SG_Log1p: Gamma exponent-separated arithmetic.
' NOTES
'   Explicit Double stores; Windows little-endian word fixtures.
'   Operation graph retained from frozen Excel port b8efaf3.
' ERROR POLICY
'   Try functions return False; test failures remain visible.
' UPDATED
'   2026-11-08
'==============================================================================
    Dim Value As Double, T As Variant, Square As Variant, Term As Variant
    Dim Total As Variant, Delta As Variant, K As Long
    Value = PROB_DD_SG_Value(Z)
    If Value < -0.8 Or Value > 0# Then Err.Raise 5
    T = PROB_DD_SG_Div(Z, PROB_DD_SG_Add(Z, 2#))
    Square = PROB_DD_SG_Mul(T, T)
    Term = T: Total = T
    For K = 3 To 2000 Step 2
        Term = PROB_DD_SG_Mul(Term, Square)
        Delta = PROB_DD_SG_Div(Term, K)
        Total = PROB_DD_SG_Add(Total, Delta)
        If PROB_DD_SG_Small(Delta, Total) Then
            PROB_DD_SG_Log1p = PROB_DD_SG_Mul(Total, 2#)
            Exit Function
        End If
    Next K
    Err.Raise 5
End Function

Private Function PROB_DD_SG_LogGamma(ByVal A As Variant) As Variant
'==============================================================================
' PURPOSE
'   PROB_DD_SG_LogGamma: Gamma exponent-separated arithmetic.
' NOTES
'   Explicit Double stores; Windows little-endian word fixtures.
'   Operation graph retained from frozen Excel port b8efaf3.
' ERROR POLICY
'   Try functions return False; test failures remain visible.
' UPDATED
'   2026-11-08
'==============================================================================
    Dim U As Variant, V As Variant
    U = PROB_DD_SG_Sub(A, 0.5)
    V = Array(PROB_DD_SP_FromWords(&H3FE96E5E&, &H1B3C5AF7&), PROB_DD_SP_FromWords(&HBC880518&, &HFEB901F7&), -49)
    V = PROB_DD_SG_Add(PROB_DD_SG_Mul(V, U), Array(PROB_DD_SP_FromWords(&HBFE35572&, &HAE72C2EE&), PROB_DD_SP_FromWords(&H3C826A5E&, &HAF086A60&), -48))
    V = PROB_DD_SG_Add(PROB_DD_SG_Mul(V, U), Array(PROB_DD_SP_FromWords(&H3FED66F3&, &H18AE5EA7&), PROB_DD_SP_FromWords(&H3C7FDC39&, &HB2705763&), -48))
    V = PROB_DD_SG_Add(PROB_DD_SG_Mul(V, U), Array(PROB_DD_SP_FromWords(&HBFE65C75&, &HB71D73A7&), PROB_DD_SP_FromWords(&HBC7C4960&, &H0B916A88&), -47))
    V = PROB_DD_SG_Add(PROB_DD_SG_Mul(V, U), Array(PROB_DD_SP_FromWords(&H3FE10279&, &H2689D188&), PROB_DD_SP_FromWords(&H3C8720CC&, &H2B7BE6FB&), -46))
    V = PROB_DD_SG_Add(PROB_DD_SG_Mul(V, U), Array(PROB_DD_SP_FromWords(&HBFE9E20D&, &H2BB58450&), PROB_DD_SP_FromWords(&HBC7BD20C&, &H6A387A6E&), -46))
    V = PROB_DD_SG_Add(PROB_DD_SG_Mul(V, U), Array(PROB_DD_SP_FromWords(&H3FE3B25E&, &H708B456B&), PROB_DD_SP_FromWords(&H3C8D35FF&, &H74EE2CC0&), -45))
    V = PROB_DD_SG_Add(PROB_DD_SG_Mul(V, U), Array(PROB_DD_SP_FromWords(&HBFEDFC0C&, &H83BBF352&), PROB_DD_SP_FromWords(&HBC5D534E&, &H18864249&), -45))
    V = PROB_DD_SG_Add(PROB_DD_SG_Mul(V, U), Array(PROB_DD_SP_FromWords(&H3FE6D3F4&, &HCE91C72E&), PROB_DD_SP_FromWords(&HBC713AEB&, &H4D8F0CF8&), -44))
    V = PROB_DD_SG_Add(PROB_DD_SG_Mul(V, U), Array(PROB_DD_SP_FromWords(&HBFE16227&, &HE2371AE4&), PROB_DD_SP_FromWords(&HBC8DB3F1&, &H0D7CC516&), -43))
    V = PROB_DD_SG_Add(PROB_DD_SG_Mul(V, U), Array(PROB_DD_SP_FromWords(&H3FEA7B27&, &HAB6726A4&), PROB_DD_SP_FromWords(&HBC83770C&, &HA96E43B7&), -43))
    V = PROB_DD_SG_Add(PROB_DD_SG_Mul(V, U), Array(PROB_DD_SP_FromWords(&HBFE42CC5&, &HCCB85BA1&), PROB_DD_SP_FromWords(&HBC750D12&, &H4F368B45&), -42))
    V = PROB_DD_SG_Add(PROB_DD_SG_Mul(V, U), Array(PROB_DD_SP_FromWords(&H3FEEBFA5&, &H2AE009C6&), PROB_DD_SP_FromWords(&HBC50F730&, &H98D885B0&), -42))
    V = PROB_DD_SG_Add(PROB_DD_SG_Mul(V, U), Array(PROB_DD_SP_FromWords(&HBFE77025&, &HFFB629FC&), PROB_DD_SP_FromWords(&H3C7A8C91&, &HFBD4CC51&), -41))
    V = PROB_DD_SG_Add(PROB_DD_SG_Mul(V, U), Array(PROB_DD_SP_FromWords(&H3FE1DED2&, &HD825BB6F&), PROB_DD_SP_FromWords(&HBC7E343D&, &HD497F856&), -40))
    V = PROB_DD_SG_Add(PROB_DD_SG_Mul(V, U), Array(PROB_DD_SP_FromWords(&HBFEB4216&, &HA140F01E&), PROB_DD_SP_FromWords(&HBC8C6D43&, &HF025D8FC&), -40))
    V = PROB_DD_SG_Add(PROB_DD_SG_Mul(V, U), Array(PROB_DD_SP_FromWords(&H3FE4CB70&, &HD3097185&), PROB_DD_SP_FromWords(&H3C846D36&, &H7A4A7A07&), -39))
    V = PROB_DD_SG_Add(PROB_DD_SG_Mul(V, U), Array(PROB_DD_SP_FromWords(&HBFEFBCAF&, &H3A802F65&), PROB_DD_SP_FromWords(&HBC71601D&, &H7CA4BACB&), -39))
    V = PROB_DD_SG_Add(PROB_DD_SG_Mul(V, U), Array(PROB_DD_SP_FromWords(&H3FE839E0&, &HDACB4830&), PROB_DD_SP_FromWords(&H3C62D497&, &HE7B31E9A&), -38))
    V = PROB_DD_SG_Add(PROB_DD_SG_Mul(V, U), Array(PROB_DD_SP_FromWords(&HBFE27FA0&, &H40C5DA37&), PROB_DD_SP_FromWords(&HBC8B28EB&, &H572E2B88&), -37))
    V = PROB_DD_SG_Add(PROB_DD_SG_Mul(V, U), Array(PROB_DD_SP_FromWords(&H3FEC426C&, &HA67FEF49&), PROB_DD_SP_FromWords(&H3C822C46&, &H74A1157A&), -37))
    V = PROB_DD_SG_Add(PROB_DD_SG_Mul(V, U), Array(PROB_DD_SP_FromWords(&HBFE597BF&, &HC48AC196&), PROB_DD_SP_FromWords(&H3C8BAEE5&, &HC54842E9&), -36))
    V = PROB_DD_SG_Add(PROB_DD_SG_Mul(V, U), Array(PROB_DD_SP_FromWords(&H3FE08130&, &H11D7D91A&), PROB_DD_SP_FromWords(&HBC8D550A&, &HF0099556&), -35))
    V = PROB_DD_SG_Add(PROB_DD_SG_Mul(V, U), Array(PROB_DD_SP_FromWords(&HBFE93D7E&, &HA20F6C7F&), PROB_DD_SP_FromWords(&HBC812DE8&, &HCBCD3D82&), -35))
    V = PROB_DD_SG_Add(PROB_DD_SG_Mul(V, U), Array(PROB_DD_SP_FromWords(&H3FE34E99&, &HD340CA87&), PROB_DD_SP_FromWords(&HBC8416AF&, &H310C8312&), -34))
    V = PROB_DD_SG_Add(PROB_DD_SG_Mul(V, U), Array(PROB_DD_SP_FromWords(&HBFED8C80&, &H723BC225&), PROB_DD_SP_FromWords(&HBC8E0DF9&, &H422BE676&), -34))
    V = PROB_DD_SG_Add(PROB_DD_SG_Mul(V, U), Array(PROB_DD_SP_FromWords(&H3FE69F02&, &HF35C279F&), PROB_DD_SP_FromWords(&HBC8DBF3F&, &H0099800A&), -33))
    V = PROB_DD_SG_Add(PROB_DD_SG_Mul(V, U), Array(PROB_DD_SP_FromWords(&HBFE15338&, &HD62BEBF4&), PROB_DD_SP_FromWords(&HBC8451AE&, &H0489B4E2&), -32))
    V = PROB_DD_SG_Add(PROB_DD_SG_Mul(V, U), Array(PROB_DD_SP_FromWords(&H3FEA8CBF&, &H58F242AB&), PROB_DD_SP_FromWords(&HBC747A52&, &HCD3A6641&), -32))
    V = PROB_DD_SG_Add(PROB_DD_SG_Mul(V, U), Array(PROB_DD_SP_FromWords(&HBFE45A45&, &H41411B2D&), PROB_DD_SP_FromWords(&H3C809250&, &H85654C7D&), -31))
    V = PROB_DD_SG_Add(PROB_DD_SG_Mul(V, U), Array(PROB_DD_SP_FromWords(&H3FEF381E&, &H2C0150C7&), PROB_DD_SP_FromWords(&H3C8403A5&, &H76642969&), -31))
    V = PROB_DD_SG_Add(PROB_DD_SG_Mul(V, U), Array(PROB_DD_SP_FromWords(&HBFE7F4C1&, &H8C998D6B&), PROB_DD_SP_FromWords(&H3C16A630&, &HA93DF275&), -30))
    V = PROB_DD_SG_Add(PROB_DD_SG_Mul(V, U), Array(PROB_DD_SP_FromWords(&H3FE2647E&, &HB6C95516&), PROB_DD_SP_FromWords(&HBC8ADEDC&, &HBF885061&), -29))
    V = PROB_DD_SG_Add(PROB_DD_SG_Mul(V, U), Array(PROB_DD_SP_FromWords(&HBFEC4210&, &H32D6F34B&), PROB_DD_SP_FromWords(&H3C8704DC&, &HFCDFA893&), -29))
    V = PROB_DD_SG_Add(PROB_DD_SG_Mul(V, U), Array(PROB_DD_SP_FromWords(&H3FE5B86D&, &HB198417B&), PROB_DD_SP_FromWords(&H3C7E2380&, &H4C8C8260&), -28))
    V = PROB_DD_SG_Add(PROB_DD_SG_Mul(V, U), Array(PROB_DD_SP_FromWords(&HBFE0B4A4&, &H4934AF42&), PROB_DD_SP_FromWords(&HBC85F670&, &H3DFCD0A8&), -27))
    V = PROB_DD_SG_Add(PROB_DD_SG_Mul(V, U), Array(PROB_DD_SP_FromWords(&H3FE9B6CA&, &H551B03C7&), PROB_DD_SP_FromWords(&H3C759D8F&, &H95F83956&), -27))
    V = PROB_DD_SG_Add(PROB_DD_SG_Mul(V, U), Array(PROB_DD_SP_FromWords(&HBFE3CDBA&, &H9595A081&), PROB_DD_SP_FromWords(&H3C8E2DB5&, &H54614458&), -26))
    V = PROB_DD_SG_Add(PROB_DD_SG_Mul(V, U), Array(PROB_DD_SP_FromWords(&H3FEE8689&, &H4A5D3179&), PROB_DD_SP_FromWords(&HBC096647&, &H8B98A9FE&), -26))
    V = PROB_DD_SG_Add(PROB_DD_SG_Mul(V, U), Array(PROB_DD_SP_FromWords(&HBFE78B4D&, &HBB5132FE&), PROB_DD_SP_FromWords(&H3C71A3B4&, &HA3DE5DB3&), -25))
    V = PROB_DD_SG_Add(PROB_DD_SG_Mul(V, U), Array(PROB_DD_SP_FromWords(&H3FE22C93&, &H4E31CF80&), PROB_DD_SP_FromWords(&HBC56E213&, &HE5457B50&), -24))
    V = PROB_DD_SG_Add(PROB_DD_SG_Mul(V, U), Array(PROB_DD_SP_FromWords(&HBFEC14F1&, &H98682057&), PROB_DD_SP_FromWords(&HBC5C1779&, &H819E660D&), -24))
    V = PROB_DD_SG_Add(PROB_DD_SG_Mul(V, U), Array(PROB_DD_SP_FromWords(&H3FE5B70D&, &H8C2D3282&), PROB_DD_SP_FromWords(&H3C741106&, &HA2C5BF7F&), -23))
    V = PROB_DD_SG_Add(PROB_DD_SG_Mul(V, U), Array(PROB_DD_SP_FromWords(&HBFE0CED7&, &HC5B04032&), PROB_DD_SP_FromWords(&HBC55C891&, &H4F796ECC&), -22))
    V = PROB_DD_SG_Add(PROB_DD_SG_Mul(V, U), Array(PROB_DD_SP_FromWords(&H3FEA0BDC&, &H123131CA&), PROB_DD_SP_FromWords(&HBC730167&, &H4563F2DA&), -22))
    V = PROB_DD_SG_Add(PROB_DD_SG_Mul(V, U), Array(PROB_DD_SP_FromWords(&HBFE43410&, &H29DBA039&), PROB_DD_SP_FromWords(&HBC535ECE&, &H10F3741A&), -21))
    V = PROB_DD_SG_Add(PROB_DD_SG_Mul(V, U), Array(PROB_DD_SP_FromWords(&H3FEF610E&, &H9965B00C&), PROB_DD_SP_FromWords(&HBC7EFCD8&, &H596275EC&), -21))
    V = PROB_DD_SG_Add(PROB_DD_SG_Mul(V, U), Array(PROB_DD_SP_FromWords(&HBFE8662D&, &HE65ECA58&), PROB_DD_SP_FromWords(&H3C68815F&, &HE55CE225&), -20))
    V = PROB_DD_SG_Add(PROB_DD_SG_Mul(V, U), Array(PROB_DD_SP_FromWords(&H3FE2FF5A&, &H00ECFEF1&), PROB_DD_SP_FromWords(&H3C69D037&, &H362FA1FF&), -19))
    V = PROB_DD_SG_Add(PROB_DD_SG_Mul(V, U), Array(PROB_DD_SP_FromWords(&HBFEDA062&, &HB3773FE2&), PROB_DD_SP_FromWords(&HBC8405C6&, &H0165D191&), -19))
    V = PROB_DD_SG_Add(PROB_DD_SG_Mul(V, U), Array(PROB_DD_SP_FromWords(&H3FE72340&, &H755AF21F&), PROB_DD_SP_FromWords(&H3C66DDFF&, &HF394B316&), -18))
    V = PROB_DD_SG_Add(PROB_DD_SG_Mul(V, U), Array(PROB_DD_SP_FromWords(&HBFE219DC&, &HEA3F7CEE&), PROB_DD_SP_FromWords(&H3C86E68A&, &HB8EDD7FC&), -17))
    V = PROB_DD_SG_Add(PROB_DD_SG_Mul(V, U), Array(PROB_DD_SP_FromWords(&H3FEC5FD3&, &H12A5D28A&), PROB_DD_SP_FromWords(&HBC8E392F&, &H1C7FA64A&), -17))
    V = PROB_DD_SG_Add(PROB_DD_SG_Mul(V, U), Array(PROB_DD_SP_FromWords(&HBFE648D2&, &H1536C0B3&), PROB_DD_SP_FromWords(&H3C88C931&, &HEEBB3135&), -16))
    V = PROB_DD_SG_Add(PROB_DD_SG_Mul(V, U), Array(PROB_DD_SP_FromWords(&H3FE18A72&, &HD2D506DD&), PROB_DD_SP_FromWords(&HBC8B73BB&, &HC28C2A0D&), -15))
    V = PROB_DD_SG_Add(PROB_DD_SG_Mul(V, U), Array(PROB_DD_SP_FromWords(&HBFEBAE9D&, &H8AB69059&), PROB_DD_SP_FromWords(&HBC83EB8A&, &H48F2B0AA&), -15))
    V = PROB_DD_SG_Add(PROB_DD_SG_Mul(V, U), Array(PROB_DD_SP_FromWords(&H3FE5E735&, &H99737474&), PROB_DD_SP_FromWords(&HBC7A754B&, &H20F04961&), -14))
    V = PROB_DD_SG_Add(PROB_DD_SG_Mul(V, U), Array(PROB_DD_SP_FromWords(&HBFE1623A&, &H15904D43&), PROB_DD_SP_FromWords(&H3C86D799&, &H89C80AEC&), -13))
    V = PROB_DD_SG_Add(PROB_DD_SG_Mul(V, U), Array(PROB_DD_SP_FromWords(&H3FEBB048&, &H486FE2BA&), PROB_DD_SP_FromWords(&H3C81CC08&, &H9C09215D&), -13))
    V = PROB_DD_SG_Add(PROB_DD_SG_Mul(V, U), Array(PROB_DD_SP_FromWords(&HBFE62327&, &H34DCCB6A&), PROB_DD_SP_FromWords(&HBC6B6203&, &H0AECE0A0&), -12))
    V = PROB_DD_SG_Add(PROB_DD_SG_Mul(V, U), Array(PROB_DD_SP_FromWords(&H3FE1C746&, &H45C54A48&), PROB_DD_SP_FromWords(&HBC62A4CE&, &H93235B7C&), -11))
    V = PROB_DD_SG_Add(PROB_DD_SG_Mul(V, U), Array(PROB_DD_SP_FromWords(&HBFECB456&, &HB864572F&), PROB_DD_SP_FromWords(&HBC879CE2&, &HBFA33CA4&), -11))
    V = PROB_DD_SG_Add(PROB_DD_SG_Mul(V, U), Array(PROB_DD_SP_FromWords(&H3FE750AE&, &H6A104859&), PROB_DD_SP_FromWords(&HBC87737F&, &HB30C7F3C&), -10))
    V = PROB_DD_SG_Add(PROB_DD_SG_Mul(V, U), Array(PROB_DD_SP_FromWords(&HBFE313E8&, &H85750BAC&), PROB_DD_SP_FromWords(&H3C716907&, &HB67084CE&), -9))
    V = PROB_DD_SG_Add(PROB_DD_SG_Mul(V, U), Array(PROB_DD_SP_FromWords(&H3FEF814B&, &H7BDFBBCC&), PROB_DD_SP_FromWords(&HBC57AB64&, &H7A72E512&), -9))
    V = PROB_DD_SG_Add(PROB_DD_SG_Mul(V, U), Array(PROB_DD_SP_FromWords(&HBFEA5035&, &H23A4E3E4&), PROB_DD_SP_FromWords(&H3C8630EE&, &H57ED3E2C&), -8))
    V = PROB_DD_SG_Add(PROB_DD_SG_Mul(V, U), Array(PROB_DD_SP_FromWords(&H3FE64F06&, &H652465F7&), PROB_DD_SP_FromWords(&H3C8CAC2D&, &H53771289&), -7))
    V = PROB_DD_SG_Add(PROB_DD_SG_Mul(V, U), Array(PROB_DD_SP_FromWords(&HBFE34DBD&, &HA148F3D6&), PROB_DD_SP_FromWords(&H3C89911A&, &H6BE1F65D&), -6))
    V = PROB_DD_SG_Add(PROB_DD_SG_Mul(V, U), Array(PROB_DD_SP_FromWords(&H3FE13342&, &H35278FAD&), PROB_DD_SP_FromWords(&H3C8D9D86&, &H2155642B&), -5))
    V = PROB_DD_SG_Add(PROB_DD_SG_Mul(V, U), Array(PROB_DD_SP_FromWords(&HBFE007AA&, &H83CEE1DF&), PROB_DD_SP_FromWords(&H3C7125D1&, &HC90250D0&), -4))
    V = PROB_DD_SG_Add(PROB_DD_SG_Mul(V, U), Array(PROB_DD_SP_FromWords(&H3FE01AF6&, &H2A2929E8&), PROB_DD_SP_FromWords(&HBC84F325&, &HBDDAFAE3&), -3))
    V = PROB_DD_SG_Add(PROB_DD_SG_Mul(V, U), Array(PROB_DD_SP_FromWords(&HBFE2AED0&, &H59BD608A&), PROB_DD_SP_FromWords(&HBC4CD3D2&, &HCA77B63A&), -2))
    V = PROB_DD_SG_Add(PROB_DD_SG_Mul(V, U), Array(PROB_DD_SP_FromWords(&H3FEEEB95&, &HB094C191&), PROB_DD_SP_FromWords(&H3C834686&, &H3F58B075&), -1))
    PROB_DD_SG_LogGamma = PROB_DD_SG_Mul(PROB_DD_SG_Mul(A, PROB_DD_SG_Sub(A, 1#)), V)
End Function

Private Function PROB_DD_WG_Stirling(ByVal A As Variant) As Variant
'==============================================================================
' PURPOSE
'   PROB_DD_WG_Stirling: Gamma exponent-separated arithmetic.
' NOTES
'   Explicit Double stores; Windows little-endian word fixtures.
'   Operation graph retained from frozen Excel port b8efaf3.
' ERROR POLICY
'   Try functions return False; test failures remain visible.
' UPDATED
'   2026-11-08
'==============================================================================
    Dim Numerators As Variant, Denominators As Variant, Inverse As Variant, Square As Variant, Term As Variant
    Dim K As Long
    Numerators = Array(1#, -1#, 1#, -1#, 1#, -691#, 1#, -3617#, 43867#, -174611#)
    Denominators = Array(12#, 360#, 1260#, 1680#, 1188#, 360360#, 156#, 122400#, 244188#, 125400#)
    Inverse = PROB_DD_SG_Div(1#, A): Square = PROB_DD_SG_Mul(Inverse, Inverse)
    Term = PROB_DD_SG_Div(Numerators(9), Denominators(9))
    For K = 8 To 0 Step -1
        Term = PROB_DD_SG_Add(PROB_DD_SG_Mul(Term, Square), PROB_DD_SG_Div(Numerators(K), Denominators(K)))
    Next K
    PROB_DD_WG_Stirling = PROB_DD_SG_Mul(Term, Inverse)
End Function

Private Function PROB_DD_WG_Deviance(ByVal A As Double, ByVal X As Double) As Variant
'==============================================================================
' PURPOSE
'   PROB_DD_WG_Deviance: Gamma exponent-separated arithmetic.
' NOTES
'   Explicit Double stores; Windows little-endian word fixtures.
'   Operation graph retained from frozen Excel port b8efaf3.
' ERROR POLICY
'   Try functions return False; test failures remain visible.
' UPDATED
'   2026-11-08
'==============================================================================
    Dim AA As Variant, XX As Variant, U As Variant, Term As Variant, Total As Variant, Delta As Variant, Value As Variant
    Dim N As Long, Converged As Boolean
    AA = PROB_DD_SG_Coerce(A): XX = PROB_DD_SG_Coerce(X)
    U = PROB_DD_SG_Div(PROB_DD_SG_Sub(XX, AA), AA)
    If Abs(PROB_DD_SG_Value(U)) <= 0.5 Then
        Term = PROB_DD_SG_Mul(U, U): Total = PROB_DD_SG_Div(Term, 2#)
        For N = 3 To 1000
            Term = PROB_DD_SG_Mul(Term, PROB_DD_SG_Neg(U))
            Delta = PROB_DD_SG_Div(Term, N)
            Total = PROB_DD_SG_Add(Total, Delta)
            If PROB_DD_SG_Small(Delta, Total) Then
                Value = PROB_DD_SG_Mul(AA, Total)
                Converged = True
                Exit For
            End If
        Next N
        If Not Converged Then Err.Raise 5, "PROB_DD_WG_Deviance", "series cap"
    Else
        Value = PROB_DD_SG_Sub(PROB_DD_SG_Sub(XX, AA), PROB_DD_SG_Mul(AA, PROB_DD_SG_Sub(PROB_DD_SG_Log(X), PROB_DD_SG_Log(A))))
    End If
    If CDbl(Value(0)) < 0# Then Err.Raise 5, "PROB_DD_WG_Deviance", "negative deviance"
    PROB_DD_WG_Deviance = Value
End Function

Private Sub PROB_DD_WG_Upper(ByVal A As Double, ByVal X As Double, ByVal LP As Variant, ByRef P As Double, ByRef Q As Double)
'==============================================================================
' PURPOSE
'   PROB_DD_WG_Upper: Gamma exponent-separated arithmetic.
' NOTES
'   Explicit Double stores; Windows little-endian word fixtures.
'   Operation graph retained from frozen Excel port b8efaf3.
' ERROR POLICY
'   Try functions return False; test failures remain visible.
' UPDATED
'   2026-11-08
'==============================================================================
    Dim AA As Variant, XX As Variant, B As Variant, C As Variant, D As Variant, AN As Variant
    Dim Delta As Variant, Tail As Variant, H As Variant, QQ As Variant
    Dim N As Long, K As Long, Converged As Boolean
    AA = PROB_DD_SG_Coerce(A): XX = PROB_DD_SG_Coerce(X)
    B = PROB_DD_SG_Sub(PROB_DD_SG_Add(XX, 1#), AA)
    C = PROB_DD_SG_Coerce(1E+300): D = PROB_DD_SG_Div(1#, B)
    For N = 1 To 100000
        AN = PROB_DD_SG_Mul(-CDbl(N), PROB_DD_SG_Sub(N, AA))
        B = PROB_DD_SG_Add(B, 2#)
        D = PROB_DD_SG_Div(1#, PROB_DD_SG_Add(PROB_DD_SG_Mul(AN, D), B))
        C = PROB_DD_SG_Add(B, PROB_DD_SG_Div(AN, C))
        Delta = PROB_DD_SG_Sub(PROB_DD_SG_Mul(C, D), 1#)
        If CDbl(Delta(0)) = 0# Or CLng(Delta(2)) < -100 Then
            Converged = True
            Exit For
        End If
    Next N
    If Not Converged Then Err.Raise 5, "PROB_DD_WG_Upper", "continued fraction cap 100000"
    Tail = PROB_DD_SG_Coerce(0#)
    For K = N To 1 Step -1
        Tail = PROB_DD_SG_Div(PROB_DD_SG_Mul(-CDbl(K), PROB_DD_SG_Sub(K, AA)), _
            PROB_DD_SG_Add(PROB_DD_SG_Add(PROB_DD_SG_Sub(PROB_DD_SG_Add(XX, 1#), AA), 2# * K), Tail))
    Next K
    H = PROB_DD_SG_Div(1#, PROB_DD_SG_Add(PROB_DD_SG_Sub(PROB_DD_SG_Add(XX, 1#), AA), Tail))
    QQ = PROB_DD_SG_Mul(PROB_DD_SG_Mul(PROB_DD_SG_Exp(LP), AA), H)
    P = PROB_DD_SG_Value(PROB_DD_SG_Sub(1#, QQ)): Q = PROB_DD_SG_Value(QQ)
End Sub

Private Function PROB_DD_WG_Pair(ByVal A As Double, ByVal X As Double, ByRef P As Double, ByRef Q As Double, _
    Optional ByRef Diagnostic As String = "") As Boolean
'==============================================================================
' PURPOSE
'   PROB_DD_WG_Pair: Gamma exponent-separated arithmetic.
' NOTES
'   Explicit Double stores; Windows little-endian word fixtures.
'   Operation graph retained from frozen Excel port b8efaf3.
' ERROR POLICY
'   Try functions return False; test failures remain visible.
' UPDATED
'   2026-11-08
'==============================================================================
    Dim AA As Variant, XX As Variant, Dev As Variant, LP As Variant, Constant As Variant
    Dim Term As Variant, Total As Variant, Ratio As Variant, Gap As Variant, Remainder As Variant, Guarded As Variant
    Dim PP As Variant, QQ As Variant, N As Long, Converged As Boolean, Underflow As Boolean, Stage As String
    On Error GoTo Err_Handler
    P = 0#: Q = 0#: Diagnostic = "": Stage = "inputs"
    If Not PROB_IsFinite(A) Then Exit Function
    If Not PROB_IsFinite(X) Then Exit Function
    If A <= 0# Or X < 0# Then Exit Function
    If X = 0# Then
        Q = 1#: PROB_DD_WG_Pair = True
        Exit Function
    End If
    If A <= 128# And X <= 1024# Then
        PROB_DD_WG_Pair = PROB_DD_EG_Pair(A, X, P, Q, Diagnostic)
        Exit Function
    End If
    AA = PROB_DD_SG_Coerce(A): XX = PROB_DD_SG_Coerce(X)
    If A <= 128# Then
        If X >= 1280# Then
            P = 1#: PROB_DD_WG_Pair = True
            Exit Function
        End If
        Stage = "small-shape upper"
        LP = PROB_DD_SG_Sub(PROB_DD_SG_Sub(PROB_DD_SG_Mul(PROB_DD_SG_Log(X), AA), PROB_DD_EG_LogGamma(AA)), XX)
        PROB_DD_WG_Upper A, X, LP, P, Q
        PROB_DD_WG_Pair = True
        Exit Function
    End If
    Stage = "deviance"
    Dev = PROB_DD_WG_Deviance(A, X)
    Underflow = (CLng(Dev(2)) >= 11)
    If Not Underflow Then Underflow = (PROB_DD_SG_Value(Dev) > 750#)
    If Underflow Then
        If X < A Then Q = 1# Else P = 1#
        PROB_DD_WG_Pair = True
        Exit Function
    End If
    Constant = PROB_DD_SG_Add(PROB_DD_SP_FromWords(&H3FED67F1&, &HC864BEB5&), PROB_DD_SP_FromWords(&HBC865B5A&, &H1B7FF5DF&))
    LP = PROB_DD_SG_Sub(PROB_DD_SG_Sub(PROB_DD_SG_Sub(PROB_DD_SG_Neg(Dev), PROB_DD_SG_Mul(PROB_DD_SG_Log(A), 0.5)), Constant), PROB_DD_WG_Stirling(AA))
    If X >= A + 1# Then
        Stage = "upper"
        PROB_DD_WG_Upper A, X, LP, P, Q
        PROB_DD_WG_Pair = True
        Exit Function
    End If
    Stage = "lower geometric remainder"
    Term = PROB_DD_SG_Coerce(1#): Total = Term
    For N = 1 To 100000
        Term = PROB_DD_SG_Div(PROB_DD_SG_Mul(Term, XX), PROB_DD_SG_Add(AA, N))
        Total = PROB_DD_SG_Add(Total, Term)
        Ratio = PROB_DD_SG_Div(XX, PROB_DD_SG_Add(PROB_DD_SG_Add(AA, N), 1#))
        Gap = PROB_DD_SG_Sub(1#, Ratio)
        If CDbl(Gap(0)) > 0# Then
            Remainder = PROB_DD_SG_Div(PROB_DD_SG_Mul(Term, Ratio), Gap)
            Guarded = PROB_DD_SG_Mul(Remainder, 4#)
            If CDbl(Guarded(0)) = 0# Or CLng(Guarded(2)) < CLng(Total(2)) - 68 Then
                Converged = True
                Exit For
            End If
        End If
    Next N
    If Not Converged Then Err.Raise 5, "PROB_DD_WG_Pair", "lower series cap 100000"
    Stage = "lower exponential"
    PROB_DD_EG_LowerPair PROB_DD_SG_Add(LP, PROB_DD_EG_LogPair(Total)), PP, QQ
    P = PROB_DD_SG_Value(PP): Q = PROB_DD_SG_Value(QQ)
    PROB_DD_WG_Pair = True
    Exit Function
Err_Handler:
    Diagnostic = Stage & ": " & Err.Source & ": " & Err.Description
    P = 0#: Q = 0#: PROB_DD_WG_Pair = False
End Function

Private Function PROB_DD_EG_LogPair(ByVal A As Variant) As Variant
'==============================================================================
' PURPOSE
'   PROB_DD_EG_LogPair: Gamma exponent-separated arithmetic.
' NOTES
'   Explicit Double stores; Windows little-endian word fixtures.
'   Operation graph retained from frozen Excel port b8efaf3.
' ERROR POLICY
'   Try functions return False; test failures remain visible.
' UPDATED
'   2026-11-08
'==============================================================================
    Dim F As Variant, Z As Variant, Square As Variant, Term As Variant, Total As Variant, Delta As Variant
    Dim N As Long
    If CDbl(A(0)) <= 0# Then Err.Raise 5
    F = Array(CDbl(A(0)), CDbl(A(1)), 0&)
    Z = PROB_DD_SG_Div(PROB_DD_SG_Sub(F, 1#), PROB_DD_SG_Add(F, 1#))
    Square = PROB_DD_SG_Mul(Z, Z): Term = Z: Total = Z
    For N = 3 To 1000 Step 2
        Term = PROB_DD_SG_Mul(Term, Square)
        Delta = PROB_DD_SG_Div(Term, N)
        Total = PROB_DD_SG_Add(Total, Delta)
        If PROB_DD_SG_Small(Delta, Total) Then
            PROB_DD_EG_LogPair = PROB_DD_SG_Add(PROB_DD_SG_Mul(Total, 2#), PROB_DD_SG_Mul(PROB_DD_SG_Ln2(), CLng(A(2))))
            Exit Function
        End If
    Next N
    Err.Raise 5, "PROB_DD_EG_LogPair", "iteration cap"
End Function

Private Function PROB_DD_EG_LogGamma(ByVal A As Variant) As Variant
'==============================================================================
' PURPOSE
'   PROB_DD_EG_LogGamma: Gamma exponent-separated arithmetic.
' NOTES
'   Explicit Double stores; Windows little-endian word fixtures.
'   Operation graph retained from frozen Excel port b8efaf3.
' ERROR POLICY
'   Try functions return False; test failures remain visible.
' UPDATED
'   2026-11-08
'==============================================================================
    Dim N As Long, K As Long, B As Variant, LG As Variant
    If PROB_DD_SG_Value(A) > 1# Then N = CLng(Fix(PROB_DD_SG_Value(A)))
    B = PROB_DD_SG_Sub(A, N)
    LG = PROB_DD_SG_LogGamma(B)
    For K = 1 To N
        LG = PROB_DD_SG_Add(LG, PROB_DD_EG_LogPair(PROB_DD_SG_Add(B, K)))
    Next K
    PROB_DD_EG_LogGamma = LG
End Function

Private Sub PROB_DD_EG_LowerPair(ByVal LP As Variant, ByRef PP As Variant, ByRef QQ As Variant)
'==============================================================================
' PURPOSE
'   PROB_DD_EG_LowerPair: Gamma exponent-separated arithmetic.
' NOTES
'   Explicit Double stores; Windows little-endian word fixtures.
'   Operation graph retained from frozen Excel port b8efaf3.
' ERROR POLICY
'   Try functions return False; test failures remain visible.
' UPDATED
'   2026-11-08
'==============================================================================
    Dim Value As Double, Z As Variant, Halvings As Long, K As Long
    Value = PROB_DD_SG_Value(LP)
    If CDbl(LP(0)) > 0# Then Err.Raise 5, "PROB_DD_EG_LowerPair", "positive log probability"
    If Value < -750# Then
        PP = PROB_DD_SG_Coerce(0#): QQ = PROB_DD_SG_Coerce(1#)
        Exit Sub
    End If
    Z = LP
    Do While PROB_DD_SG_Value(Z) < -0.5
        Z = PROB_DD_SG_Mul(Z, 0.5)
        Halvings = Halvings + 1
    Loop
    QQ = PROB_DD_SG_Neg(PROB_DD_SG_Expm1(Z))
    For K = 1 To Halvings
        QQ = PROB_DD_SG_Mul(QQ, PROB_DD_SG_Sub(2#, QQ))
    Next K
    If Value >= -0.5 Then
        PP = PROB_DD_SG_Sub(1#, QQ)
    Else
        PP = PROB_DD_SG_Exp(LP)
    End If
End Sub

Private Function PROB_DD_EG1_Pair(ByVal A As Double, ByVal X As Double, ByRef P As Double, ByRef Q As Double, _
    Optional ByRef Diagnostic As String = "") As Boolean
'==============================================================================
' PURPOSE
'   PROB_DD_EG1_Pair: Gamma exponent-separated arithmetic.
' NOTES
'   Explicit Double stores; Windows little-endian word fixtures.
'   Operation graph retained from frozen Excel port b8efaf3.
' ERROR POLICY
'   Try functions return False; test failures remain visible.
' UPDATED
'   2026-11-08
'==============================================================================
    Dim AA As Variant, XX As Variant, LG As Variant, Term As Variant, Total As Variant
    Dim Delta As Variant, LP As Variant, PP As Variant, QQ As Variant
    Dim B As Variant, C As Variant, D As Variant, AN As Variant, Tail As Variant, H As Variant
    Dim N As Long, K As Long, Converged As Boolean, Stage As String
    On Error GoTo Err_Handler
    P = 0#: Q = 0#: Diagnostic = "": Stage = "inputs"
    If Not PROB_IsFinite(A) Then Exit Function
    If Not PROB_IsFinite(X) Then Exit Function
    If A <= 0# Or A > 128# Or X <= 0# Or X > 1024# Then Exit Function
    AA = PROB_DD_SG_Coerce(A): XX = PROB_DD_SG_Coerce(X)
    Stage = "loggamma"
    LG = PROB_DD_EG_LogGamma(AA)
    If X < A + 1# Then
        If A <= 1# Then
            Stage = "alternating lower series"
            Term = PROB_DD_SG_Neg(XX)
            Total = PROB_DD_SG_Div(Term, PROB_DD_SG_Add(AA, 1#))
            For N = 2 To 1000
                Term = PROB_DD_SG_Div(PROB_DD_SG_Mul(Term, PROB_DD_SG_Neg(XX)), N)
                Delta = PROB_DD_SG_Div(Term, PROB_DD_SG_Add(AA, N))
                Total = PROB_DD_SG_Add(Total, Delta)
                If PROB_DD_SG_Small(Delta, Total) Then
                    Converged = True
                    Exit For
                End If
            Next N
            If Not Converged Then Err.Raise 5
            LP = PROB_DD_SG_Add(PROB_DD_SG_Sub(PROB_DD_SG_Mul(PROB_DD_SG_Log(X), AA), LG), PROB_DD_SG_Log1p(PROB_DD_SG_Mul(Total, AA)))
        Else
            Stage = "positive lower series"
            Term = PROB_DD_SG_Coerce(1#): Total = Term
            For N = 1 To 10000
                Term = PROB_DD_SG_Div(PROB_DD_SG_Mul(Term, XX), PROB_DD_SG_Add(AA, N))
                Total = PROB_DD_SG_Add(Total, Term)
                If PROB_DD_SG_Small(Term, Total) Then
                    Converged = True
                    Exit For
                End If
            Next N
            If Not Converged Then Err.Raise 5
            LP = PROB_DD_SG_Add(PROB_DD_SG_Sub(PROB_DD_SG_Sub(PROB_DD_SG_Mul(PROB_DD_SG_Log(X), AA), XX), LG), PROB_DD_EG_LogPair(Total))
        End If
        Stage = "lower exponential"
        PROB_DD_EG_LowerPair LP, PP, QQ
    Else
        Stage = "upper forward"
        B = PROB_DD_SG_Sub(PROB_DD_SG_Add(XX, 1#), AA)
        C = PROB_DD_SG_Coerce(1E+300): D = PROB_DD_SG_Div(1#, B)
        For N = 1 To 100000
            AN = PROB_DD_SG_Mul(-CDbl(N), PROB_DD_SG_Sub(N, AA))
            B = PROB_DD_SG_Add(B, 2#)
            D = PROB_DD_SG_Div(1#, PROB_DD_SG_Add(PROB_DD_SG_Mul(AN, D), B))
            C = PROB_DD_SG_Add(B, PROB_DD_SG_Div(AN, C))
            Delta = PROB_DD_SG_Sub(PROB_DD_SG_Mul(C, D), 1#)
            If CDbl(Delta(0)) = 0# Or CLng(Delta(2)) < -100 Then
                Converged = True
                Exit For
            End If
        Next N
        If Not Converged Then Err.Raise 5
        Stage = "upper backward"
        Tail = PROB_DD_SG_Coerce(0#)
        For K = N To 1 Step -1
            Tail = PROB_DD_SG_Div(PROB_DD_SG_Mul(-CDbl(K), PROB_DD_SG_Sub(K, AA)), _
                PROB_DD_SG_Add(PROB_DD_SG_Add(PROB_DD_SG_Sub(PROB_DD_SG_Add(XX, 1#), AA), 2# * K), Tail))
        Next K
        H = PROB_DD_SG_Div(1#, PROB_DD_SG_Add(PROB_DD_SG_Sub(PROB_DD_SG_Add(XX, 1#), AA), Tail))
        Stage = "upper exponential"
        QQ = PROB_DD_SG_Mul(PROB_DD_SG_Mul(PROB_DD_SG_Exp(PROB_DD_SG_Sub(PROB_DD_SG_Sub(PROB_DD_SG_Mul(PROB_DD_SG_Log(X), AA), LG), XX)), AA), H)
        PP = PROB_DD_SG_Sub(1#, QQ)
    End If
    Stage = "final rounding"
    P = PROB_DD_SG_Value(PP): Q = PROB_DD_SG_Value(QQ)
    PROB_DD_EG1_Pair = True
    Exit Function
Err_Handler:
    Diagnostic = Stage & ": " & Err.Source & ": " & Err.Description
    P = 0#: Q = 0#
    PROB_DD_EG1_Pair = False
End Function

Private Function PROB_DD_EG_Pair(ByVal A As Double, ByVal X As Double, ByRef P As Double, ByRef Q As Double, _
    Optional ByRef Diagnostic As String = "") As Boolean
'==============================================================================
' PURPOSE
'   PROB_DD_EG_Pair: Gamma exponent-separated arithmetic.
' NOTES
'   Explicit Double stores; Windows little-endian word fixtures.
'   Operation graph retained from frozen Excel port b8efaf3.
' ERROR POLICY
'   Try functions return False; test failures remain visible.
' UPDATED
'   2026-11-08
'==============================================================================
    Dim H As Double, L As Double, E As Long
    On Error GoTo Err_Handler
    P = 0#: Q = 0#: Diagnostic = ""
    If Not PROB_IsFinite(A) Then Exit Function
    If Not PROB_IsFinite(X) Then Exit Function
    If A = 2# And X > 0# And X <= PROB_DD_SP_FromWords(&H1FF00000&, 0) Then
        If Not PROB_DD_SP_Product(X, X, H, L, E) Then Err.Raise 5
        E = E - 1
        PROB_DD_SGD_Canonicalize H, L, E
        If Not PROB_DD_SGD_Round(H, L, E, P) Then Err.Raise 5
        Q = 1#
        PROB_DD_EG_Pair = True
    Else
        PROB_DD_EG_Pair = PROB_DD_EG1_Pair(A, X, P, Q, Diagnostic)
    End If
    Exit Function
Err_Handler:
    P = 0#: Q = 0#: Diagnostic = "shape-two subnormal: " & Err.Description
    PROB_DD_EG_Pair = False
End Function

Private Function PROB_DD_PV2_H(ByVal V As Variant) As Double
'==============================================================================
' PURPOSE
'   PROB_DD_PV2_H: test-only arithmetic port of the frozen Precision v2 model.
' NOTES
'   Explicit Double stores preserve the operation boundaries under test.
'   This restricted diagnostic implementation is not a public kernel.
' ERROR POLICY
'   Errors propagate to PROB_DD_PV2_Pair, which returns False without acceptance.
' UPDATED
'   2026-11-08
'==============================================================================
    If IsArray(V) Then
        PROB_DD_PV2_H = CDbl(V(0))
    Else
        PROB_DD_PV2_H = CDbl(V)
    End If
End Function

Private Function PROB_DD_PV2_L(ByVal V As Variant) As Double
'==============================================================================
' PURPOSE
'   PROB_DD_PV2_L: test-only arithmetic port of the frozen Precision v2 model.
' NOTES
'   Explicit Double stores preserve the operation boundaries under test.
'   This restricted diagnostic implementation is not a public kernel.
' ERROR POLICY
'   Errors propagate to PROB_DD_PV2_Pair, which returns False without acceptance.
' UPDATED
'   2026-11-08
'==============================================================================
    If IsArray(V) Then PROB_DD_PV2_L = CDbl(V(1))
End Function

Private Function PROB_DD_PV2_Value(ByVal V As Variant) As Double
'==============================================================================
' PURPOSE
'   PROB_DD_PV2_Value: test-only arithmetic port of the frozen Precision v2 model.
' NOTES
'   Explicit Double stores preserve the operation boundaries under test.
'   This restricted diagnostic implementation is not a public kernel.
' ERROR POLICY
'   Errors propagate to PROB_DD_PV2_Pair, which returns False without acceptance.
' UPDATED
'   2026-11-08
'==============================================================================
    Dim H As Double, L As Double
    H = PROB_DD_PV2_H(V)
    L = PROB_DD_PV2_L(V)
    PROB_DD_PV2_Value = H + L
End Function

Private Function PROB_DD_PV2_Sum(ByVal A As Double, ByVal B As Double) As Variant
'==============================================================================
' PURPOSE
'   PROB_DD_PV2_Sum: test-only arithmetic port of the frozen Precision v2 model.
' NOTES
'   Explicit Double stores preserve the operation boundaries under test.
'   This restricted diagnostic implementation is not a public kernel.
' ERROR POLICY
'   Errors propagate to PROB_DD_PV2_Pair, which returns False without acceptance.
' UPDATED
'   2026-11-08
'==============================================================================
    Dim S As Double, V As Double, AV As Double, AR As Double, BR As Double
    S = A + B
    V = S - A
    AV = S - V
    AR = A - AV
    BR = B - V
    PROB_DD_PV2_Sum = Array(S, AR + BR)
End Function

Private Function PROB_DD_PV2_Add(ByVal A As Variant, ByVal B As Variant) As Variant
'==============================================================================
' PURPOSE
'   PROB_DD_PV2_Add: test-only arithmetic port of the frozen Precision v2 model.
' NOTES
'   Explicit Double stores preserve the operation boundaries under test.
'   This restricted diagnostic implementation is not a public kernel.
' ERROR POLICY
'   Errors propagate to PROB_DD_PV2_Pair, which returns False without acceptance.
' UPDATED
'   2026-11-08
'==============================================================================
    Dim U As Variant, V As Variant, S As Double, E As Double, F As Double
    U = PROB_DD_PV2_Sum(PROB_DD_PV2_H(A), PROB_DD_PV2_H(B))
    V = PROB_DD_PV2_Sum(PROB_DD_PV2_L(A), PROB_DD_PV2_L(B))
    S = U(0)
    E = U(1)
    E = E + CDbl(V(0))
    F = V(1)
    U = PROB_DD_PV2_Sum(S, E)
    S = U(0)
    E = U(1)
    E = E + F
    PROB_DD_PV2_Add = PROB_DD_PV2_Sum(S, E)
End Function

Private Function PROB_DD_PV2_Neg(ByVal A As Variant) As Variant
'==============================================================================
' PURPOSE
'   PROB_DD_PV2_Neg: test-only arithmetic port of the frozen Precision v2 model.
' NOTES
'   Explicit Double stores preserve the operation boundaries under test.
'   This restricted diagnostic implementation is not a public kernel.
' ERROR POLICY
'   Errors propagate to PROB_DD_PV2_Pair, which returns False without acceptance.
' UPDATED
'   2026-11-08
'==============================================================================
    PROB_DD_PV2_Neg = Array(-PROB_DD_PV2_H(A), -PROB_DD_PV2_L(A))
End Function

Private Function PROB_DD_PV2_Sub(ByVal A As Variant, ByVal B As Variant) As Variant
'==============================================================================
' PURPOSE
'   PROB_DD_PV2_Sub: test-only arithmetic port of the frozen Precision v2 model.
' NOTES
'   Explicit Double stores preserve the operation boundaries under test.
'   This restricted diagnostic implementation is not a public kernel.
' ERROR POLICY
'   Errors propagate to PROB_DD_PV2_Pair, which returns False without acceptance.
' UPDATED
'   2026-11-08
'==============================================================================
    PROB_DD_PV2_Sub = PROB_DD_PV2_Add(A, PROB_DD_PV2_Neg(B))
End Function

Private Function PROB_DD_PV2_Mul(ByVal A As Variant, ByVal B As Variant) As Variant
'==============================================================================
' PURPOSE
'   PROB_DD_PV2_Mul: test-only arithmetic port of the frozen Precision v2 model.
' NOTES
'   Explicit Double stores preserve the operation boundaries under test.
'   This restricted diagnostic implementation is not a public kernel.
' ERROR POLICY
'   Errors propagate to PROB_DD_PV2_Pair, which returns False without acceptance.
' UPDATED
'   2026-11-08
'==============================================================================
    Dim AH As Double, AL As Double, BH As Double, BL As Double
    Dim CA As Double, CB As Double, HA As Double, HB As Double
    Dim LA As Double, LB As Double, P As Double, E As Double, T As Double
    AH = PROB_DD_PV2_H(A): AL = PROB_DD_PV2_L(A): BH = PROB_DD_PV2_H(B): BL = PROB_DD_PV2_L(B)
    P = AH * BH
    CA = 134217729# * AH
    CB = 134217729# * BH
    T = CA - AH
    HA = CA - T
    T = CB - BH
    HB = CB - T
    LA = AH - HA
    LB = BH - HB
    T = HA * HB
    E = T - P
    T = HA * LB
    E = E + T
    T = LA * HB
    E = E + T
    T = LA * LB
    E = E + T
    T = AH * BL
    E = E + T
    T = AL * BH
    E = E + T
    T = AL * BL
    E = E + T
    PROB_DD_PV2_Mul = PROB_DD_PV2_Sum(P, E)
End Function

Private Function PROB_DD_PV2_Div(ByVal A As Variant, ByVal B As Variant) As Variant
'==============================================================================
' PURPOSE
'   PROB_DD_PV2_Div: test-only arithmetic port of the frozen Precision v2 model.
' NOTES
'   Explicit Double stores preserve the operation boundaries under test.
'   This restricted diagnostic implementation is not a public kernel.
' ERROR POLICY
'   Errors propagate to PROB_DD_PV2_Pair, which returns False without acceptance.
' UPDATED
'   2026-11-08
'==============================================================================
    Dim Q As Variant, R As Variant, T As Double, J As Long
    T = PROB_DD_PV2_H(A) / PROB_DD_PV2_H(B)
    Q = Array(T, 0#)
    For J = 1 To 2
        R = PROB_DD_PV2_Sub(A, PROB_DD_PV2_Mul(Q, B))
        T = PROB_DD_PV2_H(R) / PROB_DD_PV2_H(B)
        Q = PROB_DD_PV2_Add(Q, T)
    Next J
    PROB_DD_PV2_Div = Q
End Function

Private Function PROB_DD_PV2_Ln2() As Variant
'==============================================================================
' PURPOSE
'   PROB_DD_PV2_Ln2: test-only arithmetic port of the frozen Precision v2 model.
' NOTES
'   Explicit Double stores preserve the operation boundaries under test.
'   This restricted diagnostic implementation is not a public kernel.
' ERROR POLICY
'   Errors propagate to PROB_DD_PV2_Pair, which returns False without acceptance.
' UPDATED
'   2026-11-08
'==============================================================================
    PROB_DD_PV2_Ln2 = Array((6.93147180559945E-01 + (3.33066907387547E-16)), (2.31904681384630E-17 + (-3.08148791101958E-33)))
End Function

Private Function PROB_DD_PV2_Log(ByVal X As Double) As Variant
'==============================================================================
' PURPOSE
'   PROB_DD_PV2_Log: test-only arithmetic port of the frozen Precision v2 model.
' NOTES
'   Explicit Double stores preserve the operation boundaries under test.
'   This restricted diagnostic implementation is not a public kernel.
' ERROR POLICY
'   Errors propagate to PROB_DD_PV2_Pair, which returns False without acceptance.
' UPDATED
'   2026-11-08
'==============================================================================
    Dim F As Double, E As Long, N As Long
    Dim Z As Variant, S As Variant, T As Variant, D As Variant, Z2 As Variant
    If X = 1# Then
        PROB_DD_PV2_Log = Array(0#, 0#)
        Exit Function
    End If
    If X <= 0# Then Err.Raise 5, , "PV2 Log domain"
    F = X
    Do While F < 0.5
        F = F * 2#
        E = E - 1
    Loop
    Do While F >= 1#
        F = F / 2#
        E = E + 1
    Loop
    Z = PROB_DD_PV2_Div(PROB_DD_PV2_Sub(F, 1#), PROB_DD_PV2_Add(F, 1#))
    Z2 = PROB_DD_PV2_Mul(Z, Z)
    T = Z: S = Z
    For N = 3 To 1000 Step 2
        T = PROB_DD_PV2_Mul(T, Z2)
        D = PROB_DD_PV2_Div(T, CDbl(N))
        S = PROB_DD_PV2_Add(S, D)
        If Abs(PROB_DD_PV2_H(D)) <= Abs(PROB_DD_PV2_H(S)) * 2# ^ (-105) Then
            PROB_DD_PV2_Log = PROB_DD_PV2_Add(PROB_DD_PV2_Mul(S, 2#), PROB_DD_PV2_Mul(PROB_DD_PV2_Ln2(), CDbl(E)))
            Exit Function
        End If
    Next N
    Err.Raise 5, , "PV2 Log cap"
End Function

Private Function PROB_DD_PV2_Log1p(ByVal Z As Variant) As Variant
'==============================================================================
' PURPOSE
'   PROB_DD_PV2_Log1p: test-only arithmetic port of the frozen Precision v2 model.
' NOTES
'   Explicit Double stores preserve the operation boundaries under test.
'   This restricted diagnostic implementation is not a public kernel.
' ERROR POLICY
'   Errors propagate to PROB_DD_PV2_Pair, which returns False without acceptance.
' UPDATED
'   2026-11-08
'==============================================================================
    Dim T As Variant, S As Variant, D As Variant, Z2 As Variant, N As Long
    If PROB_DD_PV2_Value(Z) < -0.8 Or PROB_DD_PV2_Value(Z) > 0# Then Err.Raise 5, , "PV2 Log1p domain"
    T = PROB_DD_PV2_Div(Z, PROB_DD_PV2_Add(Z, 2#))
    Z2 = PROB_DD_PV2_Mul(T, T)
    S = T
    For N = 3 To 2000 Step 2
        T = PROB_DD_PV2_Mul(T, Z2)
        D = PROB_DD_PV2_Div(T, CDbl(N))
        S = PROB_DD_PV2_Add(S, D)
        If Abs(PROB_DD_PV2_H(D)) <= Abs(PROB_DD_PV2_H(S)) * 2# ^ (-105) Then
            PROB_DD_PV2_Log1p = PROB_DD_PV2_Mul(S, 2#)
            Exit Function
        End If
    Next N
    Err.Raise 5, , "PV2 Log1p cap"
End Function

Private Function PROB_DD_PV2_Expm1(ByVal Z As Variant) As Variant
'==============================================================================
' PURPOSE
'   PROB_DD_PV2_Expm1: test-only arithmetic port of the frozen Precision v2 model.
' NOTES
'   Explicit Double stores preserve the operation boundaries under test.
'   This restricted diagnostic implementation is not a public kernel.
' ERROR POLICY
'   Errors propagate to PROB_DD_PV2_Pair, which returns False without acceptance.
' UPDATED
'   2026-11-08
'==============================================================================
    Dim T As Variant, S As Variant, N As Long
    If PROB_DD_PV2_Value(Z) < -8# Or PROB_DD_PV2_Value(Z) > 0# Then Err.Raise 5, , "PV2 Expm1 domain"
    T = Z: S = Z
    For N = 2 To 1000
        T = PROB_DD_PV2_Div(PROB_DD_PV2_Mul(T, Z), CDbl(N))
        S = PROB_DD_PV2_Add(S, T)
        If Abs(PROB_DD_PV2_H(T)) <= Abs(PROB_DD_PV2_H(S)) * 2# ^ (-105) Then
            PROB_DD_PV2_Expm1 = S
            Exit Function
        End If
    Next N
    Err.Raise 5, , "PV2 Expm1 cap"
End Function

Private Function PROB_DD_PV2_Exp(ByVal Z As Variant) As Variant
'==============================================================================
' PURPOSE
'   PROB_DD_PV2_Exp: test-only arithmetic port of the frozen Precision v2 model.
' NOTES
'   Explicit Double stores preserve the operation boundaries under test.
'   This restricted diagnostic implementation is not a public kernel.
' ERROR POLICY
'   Errors propagate to PROB_DD_PV2_Pair, which returns False without acceptance.
' UPDATED
'   2026-11-08
'==============================================================================
    Dim R As Variant, T As Variant, S As Variant, N As Long, K As Long
    Dim Ratio As Double, ScaleFactor As Double, H As Double, L As Double
    If PROB_DD_PV2_Value(Z) < -64# Or PROB_DD_PV2_Value(Z) > 8# Then Err.Raise 5, , "PV2 Exp domain"
    Ratio = PROB_DD_PV2_H(Z) / PROB_DD_PV2_H(PROB_DD_PV2_Ln2())
    N = CLng(Ratio)
    R = PROB_DD_PV2_Sub(Z, PROB_DD_PV2_Mul(PROB_DD_PV2_Ln2(), CDbl(N)))
    T = Array(1#, 0#): S = T
    For K = 1 To 1000
        T = PROB_DD_PV2_Div(PROB_DD_PV2_Mul(T, R), CDbl(K))
        S = PROB_DD_PV2_Add(S, T)
        If Abs(PROB_DD_PV2_H(T)) <= Abs(PROB_DD_PV2_H(S)) * 2# ^ (-105) Then
            ScaleFactor = 2# ^ N
            H = PROB_DD_PV2_H(S) * ScaleFactor
            L = PROB_DD_PV2_L(S) * ScaleFactor
            PROB_DD_PV2_Exp = Array(H, L)
            Exit Function
        End If
    Next K
    Err.Raise 5, , "PV2 Exp cap"
End Function

Private Function PROB_DD_PV2_LogGamma(ByVal A As Double) As Variant
'==============================================================================
' PURPOSE
'   PROB_DD_PV2_LogGamma: test-only arithmetic port of the frozen Precision v2 model.
' NOTES
'   Explicit Double stores preserve the operation boundaries under test.
'   This restricted diagnostic implementation is not a public kernel.
' ERROR POLICY
'   Errors propagate to PROB_DD_PV2_Pair, which returns False without acceptance.
' UPDATED
'   2026-11-08
'==============================================================================
    Dim U As Variant, V As Variant
    U = PROB_DD_PV2_Sub(A, 0.5)
    V = Array((1.41171093978875E-15 + (-7.88860905221012E-31)), (-7.40170704349773E-32 + (1.09476442525376E-47)))
    V = PROB_DD_PV2_Add(PROB_DD_PV2_Mul(V, U), Array((-2.14648089856621E-15 + (5.12759588393658E-30)), (1.13494325987459E-31 + (6.56858655152258E-47))))
    V = PROB_DD_PV2_Add(PROB_DD_PV2_Mul(V, U), Array((3.26429404138876E-15 + (-2.36658271566304E-30)), (9.81769918162968E-32 + (2.18952885050753E-47))))
    V = PROB_DD_PV2_Add(PROB_DD_PV2_Mul(V, U), Array((-4.96517742317239E-15 + (-7.88860905221012E-31)), (-1.74329766137388E-31 + (2.40848173555828E-46))))
    V = PROB_DD_PV2_Add(PROB_DD_PV2_Mul(V, U), Array((7.55380696239768E-15 + (1.57772181044202E-30)), (5.70152063703829E-31 + (0.00000000000000E+00))))
    V = PROB_DD_PV2_Add(PROB_DD_PV2_Mul(V, U), Array((-1.14943670002178E-14 + (3.62876016401665E-29)), (-3.42914161585886E-31 + (-2.62743462060903E-46))))
    V = PROB_DD_PV2_Add(PROB_DD_PV2_Mul(V, U), Array((1.74942314273340E-14 + (2.52435489670724E-29)), (1.44021000372386E-30 + (4.90454462513686E-45))))
    V = PROB_DD_PV2_Add(PROB_DD_PV2_Mul(V, U), Array((-2.66316444074713E-14 + (-1.57772181044202E-29)), (-1.80731795250046E-31 + (-1.09476442525376E-46))))
    V = PROB_DD_PV2_Add(PROB_DD_PV2_Mul(V, U), Array((4.05505925851729E-14 + (-2.52435489670724E-29)), (-8.49512127749938E-31 + (-1.75162308040602E-46))))
    V = PROB_DD_PV2_Add(PROB_DD_PV2_Mul(V, U), Array((-6.17583178419540E-14 + (-2.52435489670724E-29)), (-5.85786353840856E-30 + (-2.10194769648723E-45))))
    V = PROB_DD_PV2_Add(PROB_DD_PV2_Mul(V, U), Array((9.40796740393935E-14 + (0.00000000000000E+00)), (-3.83880166798765E-30 + (3.50324616081204E-45))))
    V = PROB_DD_PV2_Add(PROB_DD_PV2_Mul(V, U), Array((-1.43351237964261E-13 + (1.00974195868290E-28)), (-4.15158968604589E-30 + (2.80259692864963E-45))))
    V = PROB_DD_PV2_Add(PROB_DD_PV2_Mul(V, U), Array((2.18482043164793E-13 + (1.51461293802434E-28)), (-8.36467935207240E-31 + (-1.75162308040602E-46))))
    V = PROB_DD_PV2_Add(PROB_DD_PV2_Mul(V, U), Array((-3.33075147079755E-13 + (-5.04870979341448E-29)), (1.04717745291746E-29 + (3.50324616081204E-44))))
    V = PROB_DD_PV2_Add(PROB_DD_PV2_Mul(V, U), Array((5.07907450613959E-13 + (1.00974195868290E-28)), (-2.38268087848993E-29 + (4.20389539297445E-44))))
    V = PROB_DD_PV2_Add(PROB_DD_PV2_Mul(V, U), Array((-7.74723440736884E-13 + (-5.04870979341448E-28)), (-4.48496106067019E-29 + (0.00000000000000E+00))))
    V = PROB_DD_PV2_Add(PROB_DD_PV2_Mul(V, U), Array((1.18204129155139E-12 + (-3.43312265952184E-27)), (6.44550241526769E-29 + (4.48415508583941E-44))))
    V = PROB_DD_PV2_Add(PROB_DD_PV2_Mul(V, U), Array((-1.80404235692326E-12 + (-2.22143230910237E-27)), (-2.74136263274511E-29 + (1.68155815718978E-44))))
    V = PROB_DD_PV2_Add(PROB_DD_PV2_Mul(V, U), Array((2.75418725094515E-12 + (-3.63507105125842E-27)), (2.97092004439325E-29 + (-1.12103877145985E-44))))
    V = PROB_DD_PV2_Add(PROB_DD_PV2_Mul(V, U), Array((-4.20608080662836E-12 + (-2.42338070083895E-27)), (-3.42805388259430E-28 + (2.69049305150365E-43))))
    V = PROB_DD_PV2_Add(PROB_DD_PV2_Mul(V, U), Array((6.42545964487383E-12 + (1.61558713389263E-27)), (2.29374877444704E-28 + (-4.93257059442336E-43))))
    V = PROB_DD_PV2_Add(PROB_DD_PV2_Mul(V, U), Array((-9.81925490738286E-12 + (0.00000000000000E+00)), (6.98822051228334E-28 + (-1.79366203433577E-43))))
    V = PROB_DD_PV2_Add(PROB_DD_PV2_Mul(V, U), Array((1.50108823940263E-11 + (4.20052654812084E-26)), (-1.48089756038230E-27 + (-2.69049305150365E-42))))
    V = PROB_DD_PV2_Add(PROB_DD_PV2_Mul(V, U), Array((-2.29558404652975E-11 + (-3.87740912134232E-26)), (-8.67334689656528E-28 + (-3.58732406867153E-43))))
    V = PROB_DD_PV2_Add(PROB_DD_PV2_Mul(V, U), Array((3.51192915084393E-11 + (-1.29246970711411E-26)), (-2.02843131271761E-27 + (-3.94605647553868E-42))))
    V = PROB_DD_PV2_Add(PROB_DD_PV2_Mul(V, U), Array((-5.37490176317862E-11 + (0.00000000000000E+00)), (-3.03473751543268E-27 + (-4.66352128927299E-42))))
    V = PROB_DD_PV2_Add(PROB_DD_PV2_Mul(V, U), Array((8.22952234483210E-11 + (-2.58493941422821E-26)), (-6.00736993421394E-27 + (-1.43492962746861E-42))))
    V = PROB_DD_PV2_Add(PROB_DD_PV2_Mul(V, U), Array((-1.26056591454892E-10 + (1.29246970711411E-25)), (-8.20680360314836E-27 + (-1.43492962746861E-42))))
    V = PROB_DD_PV2_Add(PROB_DD_PV2_Mul(V, U), Array((1.93175181123118E-10 + (-4.39439700418796E-25)), (-4.13546401752092E-27 + (3.58732406867153E-42))))
    V = PROB_DD_PV2_Add(PROB_DD_PV2_Mul(V, U), Array((-2.96169589941651E-10 + (2.06795153138257E-25)), (1.33863843386909E-26 + (2.86985925493723E-42))))
    V = PROB_DD_PV2_Add(PROB_DD_PV2_Mul(V, U), Array((4.54299303042099E-10 + (1.03397576569128E-25)), (1.61673771521419E-26 + (1.43492962746861E-41))))
    V = PROB_DD_PV2_Add(PROB_DD_PV2_Mul(V, U), Array((-6.97213642028249E-10 + (2.06795153138257E-25)), (2.85872842316712E-28 + (3.58732406867153E-43))))
    V = PROB_DD_PV2_Add(PROB_DD_PV2_Mul(V, U), Array((1.07058780863553E-09 + (-3.72231275648862E-24)), (-8.68234486230219E-26 + (2.29588740394978E-41))))
    V = PROB_DD_PV2_Add(PROB_DD_PV2_Mul(V, U), Array((-1.64483555539652E-09 + (-4.13590306276514E-24)), (7.43783907961386E-26 + (-3.44383110592467E-41))))
    V = PROB_DD_PV2_Add(PROB_DD_PV2_Mul(V, U), Array((2.52859012613536E-09 + (-2.48154183765908E-24)), (9.73833157704985E-26 + (-2.29588740394978E-41))))
    V = PROB_DD_PV2_Add(PROB_DD_PV2_Mul(V, U), Array((-3.88958300575891E-09 + (-2.48154183765908E-24)), (-2.83860600444586E-25 + (-3.21424236552969E-40))))
    V = PROB_DD_PV2_Add(PROB_DD_PV2_Mul(V, U), Array((5.98701295630897E-09 + (-2.48154183765908E-24)), (1.39686715089691E-25 + (-4.59177480789956E-40))))
    V = PROB_DD_PV2_Add(PROB_DD_PV2_Mul(V, U), Array((-9.22178304786582E-09 + (4.96308367531817E-24)), (7.80097185002356E-25 + (-4.59177480789956E-40))))
    V = PROB_DD_PV2_Add(PROB_DD_PV2_Mul(V, U), Array((1.42145587068719E-08 + (4.13590306276514E-23)), (-2.56469703585928E-27 + (-4.30478888240584E-42))))
    V = PROB_DD_PV2_Add(PROB_DD_PV2_Mul(V, U), Array((-2.19272029043474E-08 + (2.64697796016969E-23)), (4.55969744306379E-25 + (-4.59177480789956E-40))))
    V = PROB_DD_PV2_Add(PROB_DD_PV2_Mul(V, U), Array((3.38519414938090E-08 + (3.97046694025453E-23)), (-2.95757343411792E-25 + (4.59177480789956E-40))))
    V = PROB_DD_PV2_Add(PROB_DD_PV2_Mul(V, U), Array((-5.23064498832708E-08 + (3.97046694025453E-23)), (-3.63076684141467E-25 + (4.59177480789956E-40))))
    V = PROB_DD_PV2_Add(PROB_DD_PV2_Mul(V, U), Array((8.08948668443836E-08 + (1.32348898008484E-23)), (2.07482824619517E-24 + (-7.34683969263930E-40))))
    V = PROB_DD_PV2_Add(PROB_DD_PV2_Mul(V, U), Array((-1.25229209086299E-07 + (-2.11758236813575E-22)), (-1.12617886205366E-24 + (-7.34683969263930E-40))))
    V = PROB_DD_PV2_Add(PROB_DD_PV2_Mul(V, U), Array((1.94060256838287E-07 + (-2.64697796016969E-23)), (-3.93024156918979E-24 + (-4.40810381558358E-39))))
    V = PROB_DD_PV2_Add(PROB_DD_PV2_Mul(V, U), Array((-3.01053697400556E-07 + (-5.29395592033938E-23)), (-2.00284536762141E-24 + (-2.20405190779179E-39))))
    V = PROB_DD_PV2_Add(PROB_DD_PV2_Mul(V, U), Array((4.67585459593438E-07 + (1.05879118406788E-22)), (-1.28162025011814E-23 + (1.32243114467507E-38))))
    V = PROB_DD_PV2_Add(PROB_DD_PV2_Mul(V, U), Array((-7.27150972981638E-07 + (-3.17637355220363E-22)), (1.01351832794266E-23 + (4.70197740328915E-38))))
    V = PROB_DD_PV2_Add(PROB_DD_PV2_Mul(V, U), Array((1.13233727790091E-06 + (-4.02340649945793E-21)), (2.13522964307009E-23 + (2.05711511393900E-38))))
    V = PROB_DD_PV2_Add(PROB_DD_PV2_Mul(V, U), Array((-1.76587736954239E-06 + (6.35274710440725E-22)), (-6.62490692106423E-23 + (-1.17549435082229E-38))))
    V = PROB_DD_PV2_Add(PROB_DD_PV2_Mul(V, U), Array((2.75822905390482E-06 + (0.00000000000000E+00)), (3.78305871026079E-23 + (-1.17549435082229E-38))))
    V = PROB_DD_PV2_Add(PROB_DD_PV2_Mul(V, U), Array((-4.31562117239537E-06 + (2.54109884176290E-21)), (3.03086311685494E-22 + (1.88079096131566E-37))))
    V = PROB_DD_PV2_Add(PROB_DD_PV2_Mul(V, U), Array((6.76496373796090E-06 + (-5.08219768352580E-21)), (-4.00003039013541E-22 + (0.00000000000000E+00))))
    V = PROB_DD_PV2_Add(PROB_DD_PV2_Mul(V, U), Array((-1.06260564852470E-05 + (-3.04931861011548E-20)), (6.56077790924962E-22 + (2.82118644197349E-37))))
    V = PROB_DD_PV2_Add(PROB_DD_PV2_Mul(V, U), Array((1.67282243428494E-05 + (4.40457132572236E-20)), (-1.45330121238327E-21 + (-7.52316384526264E-37))))
    V = PROB_DD_PV2_Add(PROB_DD_PV2_Mul(V, U), Array((-2.63996995905211E-05 + (0.00000000000000E+00)), (-1.05456019154754E-21 + (2.25694915357879E-36))))
    V = PROB_DD_PV2_Add(PROB_DD_PV2_Mul(V, U), Array((4.17769653553246E-05 + (-4.74338450462408E-20)), (-1.40068426065889E-21 + (-2.44502824971036E-36))))
    V = PROB_DD_PV2_Add(PROB_DD_PV2_Mul(V, U), Array((-6.63135482599680E-05 + (1.35525271560688E-20)), (4.83702115030649E-21 + (3.00926553810506E-36))))
    V = PROB_DD_PV2_Add(PROB_DD_PV2_Mul(V, U), Array((1.05623637963857E-04 + (1.35525271560688E-20)), (3.76866268963272E-21 + (-4.51389830715758E-36))))
    V = PROB_DD_PV2_Add(PROB_DD_PV2_Mul(V, U), Array((-1.68894325207693E-04 + (1.89735380184963E-19)), (-2.89927296270812E-21 + (4.13774011489445E-36))))
    V = PROB_DD_PV2_Add(PROB_DD_PV2_Mul(V, U), Array((2.71277099973908E-04 + (3.79470760369927E-19)), (-3.94797336214631E-21 + (-3.76158192263132E-36))))
    V = PROB_DD_PV2_Add(PROB_DD_PV2_Mul(V, U), Array((-4.37995120948431E-04 + (5.42101086242752E-20)), (-2.00008491458486E-20 + (2.10648587667354E-35))))
    V = PROB_DD_PV2_Add(PROB_DD_PV2_Mul(V, U), Array((7.11522258031626E-04 + (-3.25260651745651E-19)), (-3.97278226896129E-20 + (-3.00926553810506E-35))))
    V = PROB_DD_PV2_Add(PROB_DD_PV2_Mul(V, U), Array((-1.16441447410474E-03 + (2.16840434497101E-19)), (2.94941512886342E-20 + (3.00926553810506E-35))))
    V = PROB_DD_PV2_Add(PROB_DD_PV2_Mul(V, U), Array((1.92291614029804E-03 + (-3.68628738645072E-18)), (-1.00244239888136E-20 + (-2.25694915357879E-35))))
    V = PROB_DD_PV2_Add(PROB_DD_PV2_Mul(V, U), Array((-3.21207407719414E-03 + (3.03576608295941E-18)), (1.50372992274185E-19 + (1.68518870133883E-34))))
    V = PROB_DD_PV2_Add(PROB_DD_PV2_Mul(V, U), Array((5.44645784455033E-03 + (-3.46944695195361E-18)), (3.88585737757394E-19 + (-1.92592994438724E-34))))
    V = PROB_DD_PV2_Add(PROB_DD_PV2_Mul(V, U), Array((-9.42562244846982E-03 + (-1.73472347597681E-18)), (6.92989757614234E-19 + (-9.62964972193618E-35))))
    V = PROB_DD_PV2_Add(PROB_DD_PV2_Mul(V, U), Array((1.67970986311869E-02 + (-2.08166817117217E-17)), (1.60545014319234E-18 + (-4.81482486096809E-33))))
    V = PROB_DD_PV2_Add(PROB_DD_PV2_Mul(V, U), Array((-3.13084875010110E-02 + (0.00000000000000E+00)), (9.29580431618060E-19 + (3.85185988877447E-34))))
    V = PROB_DD_PV2_Add(PROB_DD_PV2_Mul(V, U), Array((6.29114010745649E-02 + (4.16333634234434E-17)), (-4.54276257267888E-18 + (-3.08148791101958E-33))))
    V = PROB_DD_PV2_Add(PROB_DD_PV2_Mul(V, U), Array((-1.45959895914306E-01 + (-8.32667268468867E-17)), (-7.81369195792092E-19 + (-1.92592994438724E-34))))
    V = PROB_DD_PV2_Add(PROB_DD_PV2_Mul(V, U), Array((4.83128950540981E-01 + (-1.11022302462516E-16)), (1.67188189971785E-17 + (6.16297582203915E-33))))
    PROB_DD_PV2_LogGamma = PROB_DD_PV2_Mul(PROB_DD_PV2_Mul(A, PROB_DD_PV2_Sub(A, 1#)), V)
End Function

Private Function PROB_DD_PV2_Pair(ByVal A As Double, ByVal X As Double, ByRef P As Double, ByRef Q As Double) As Boolean
'==============================================================================
' PURPOSE
'   PROB_DD_PV2_Pair: test-only arithmetic port of the frozen Precision v2 model.
' NOTES
'   Explicit Double stores preserve the operation boundaries under test.
'   This restricted diagnostic implementation is not a public kernel.
' ERROR POLICY
'   Errors propagate to PROB_DD_PV2_Pair, which returns False without acceptance.
' UPDATED
'   2026-11-08
'==============================================================================
    Dim LG As Variant, T As Variant, S As Variant, D As Variant, LP As Variant
    Dim B As Variant, C As Variant, H As Variant, AN As Variant, Delta As Variant
    Dim Tail As Variant, E As Variant, V As Variant, N As Long, K As Long
    On Error GoTo Err_Handler
    If A < 2# ^ (-400) Or A > 1# Or X < 2# ^ (-400) Then GoTo Err_Handler
    LG = PROB_DD_PV2_LogGamma(A)
    If X < A + 1# Then
        T = Array(-X, 0#)
        S = PROB_DD_PV2_Div(T, PROB_DD_PV2_Add(A, 1#))
        For N = 2 To 1000
            T = PROB_DD_PV2_Div(PROB_DD_PV2_Mul(T, -X), CDbl(N))
            D = PROB_DD_PV2_Div(T, PROB_DD_PV2_Add(A, CDbl(N)))
            S = PROB_DD_PV2_Add(S, D)
            If Abs(PROB_DD_PV2_H(D)) <= Abs(PROB_DD_PV2_H(S)) * 2# ^ (-105) Then Exit For
        Next N
        If N > 1000 Then GoTo Err_Handler
        LP = PROB_DD_PV2_Add(PROB_DD_PV2_Sub(PROB_DD_PV2_Mul(PROB_DD_PV2_Log(X), A), LG), PROB_DD_PV2_Log1p(PROB_DD_PV2_Mul(S, A)))
        If PROB_DD_PV2_Value(LP) > 0# Then GoTo Err_Handler
        If PROB_DD_PV2_Value(LP) >= -0.5 Then
            V = PROB_DD_PV2_Expm1(LP)
            P = PROB_DD_PV2_Value(PROB_DD_PV2_Add(V, 1#))
            Q = PROB_DD_PV2_Value(PROB_DD_PV2_Neg(V))
        Else
            V = PROB_DD_PV2_Exp(LP)
            P = PROB_DD_PV2_Value(V)
            Q = PROB_DD_PV2_Value(PROB_DD_PV2_Sub(1#, V))
        End If
    Else
        If X > 32# Then GoTo Err_Handler
        B = PROB_DD_PV2_Sub(PROB_DD_PV2_Add(X, 1#), A)
        C = Array(1E+300, 0#)
        D = PROB_DD_PV2_Div(1#, B)
        H = D
        For N = 1 To 100000
            AN = PROB_DD_PV2_Mul(CDbl(-N), PROB_DD_PV2_Sub(CDbl(N), A))
            B = PROB_DD_PV2_Add(B, 2#)
            D = PROB_DD_PV2_Add(PROB_DD_PV2_Mul(AN, D), B)
            C = PROB_DD_PV2_Add(B, PROB_DD_PV2_Div(AN, C))
            If Abs(PROB_DD_PV2_H(D)) < 1E-300 Or Abs(PROB_DD_PV2_H(C)) < 1E-300 Then GoTo Err_Handler
            D = PROB_DD_PV2_Div(1#, D)
            Delta = PROB_DD_PV2_Mul(C, D)
            H = PROB_DD_PV2_Mul(H, Delta)
            If Abs(PROB_DD_PV2_Value(PROB_DD_PV2_Sub(Delta, 1#))) <= 2# ^ (-100) Then Exit For
        Next N
        If N > 100000 Then GoTo Err_Handler
        Tail = Array(0#, 0#)
        For K = N To 1 Step -1
            B = PROB_DD_PV2_Add(PROB_DD_PV2_Add(PROB_DD_PV2_Sub(PROB_DD_PV2_Add(X, 1#), A), CDbl(2 * K)), Tail)
            Tail = PROB_DD_PV2_Div(PROB_DD_PV2_Mul(CDbl(-K), PROB_DD_PV2_Sub(CDbl(K), A)), B)
        Next K
        H = PROB_DD_PV2_Div(1#, PROB_DD_PV2_Add(PROB_DD_PV2_Sub(PROB_DD_PV2_Add(X, 1#), A), Tail))
        E = PROB_DD_PV2_Sub(PROB_DD_PV2_Sub(PROB_DD_PV2_Mul(PROB_DD_PV2_Log(X), A), LG), X)
        V = PROB_DD_PV2_Mul(PROB_DD_PV2_Mul(PROB_DD_PV2_Exp(E), A), H)
        P = PROB_DD_PV2_Value(PROB_DD_PV2_Sub(1#, V))
        Q = PROB_DD_PV2_Value(V)
    End If
    PROB_DD_PV2_Pair = True
    Exit Function
Err_Handler:
    Err.Clear
    PROB_DD_PV2_Pair = False
End Function

' END generated public Gamma integration

