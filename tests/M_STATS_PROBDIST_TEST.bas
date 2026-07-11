Attribute VB_Name = "M_STATS_PROBDIST_TEST"

Option Explicit

'==============================================================================
' M_STATS_PROBDIST_TEST
'------------------------------------------------------------------------------
' PURPOSE
'   Single self-checking test harness for the whole probability-distribution
'   stack: M_STATS_CORE, M_STATS_SPECIALFUNC, M_STATS_PROBDIST_NORMALFAMILY and
'   M_STATS_PROBDIST_TFAMILY. Verifies known values, symmetry, inverse
'   round-trips, the survival surface, lognormal moments, the arithmetic-to-log
'   parameter round-trip, interval probabilities, and the error/overflow
'   contract.
'
' HOW TO RUN
'   From the VBA IDE Immediate window (Ctrl+G), pick a scope:
'
'       Test_STATS_PROBDIST_RunAll            'everything
'       Test_STATS_PROBDIST_RunCore           'M_STATS_CORE + M_STATS_SPECIALFUNC
'       Test_STATS_PROBDIST_RunNormalFamily   'normal + lognormal
'       Test_STATS_PROBDIST_RunTFamily        't, chi-square, F
'
'   All four are argument-less Public Subs, so they also appear in the Excel
'   macro list (Alt+F8).
'
'   Results print with Debug.Print. Only failures print a detail line; passing
'   assertions are silent, so a clean run shows the suite headers, the section
'   headers and the final summary.
'
' WHY ONE MODULE
'   The two harnesses that preceded this one each carried their own copies of
'   mTestCount, AssertClose, AssertIsError and RecordResult, and each declared a
'   Private Sub named Test_ErrorContract. Two counters means two summaries and no
'   single answer to "is the library green". The section subs are prefixed
'   Test_Core_, Test_NF_ and Test_TF_ so that nothing collides, and the counters
'   and assertion helpers exist exactly once.
'
' SCOPE MAP
'   Core suite          -> M_STATS_CORE, M_STATS_SPECIALFUNC
'   NormalFamily suite  -> M_STATS_PROBDIST_NORMALFAMILY
'   TFamily suite       -> M_STATS_PROBDIST_TFAMILY
'   A suite run does not implicitly run the suites it depends on. Run Core first
'   when a lower-layer change is suspected; RunAll runs the three in dependency
'   order, so the first FAIL line is normally the deepest one.
'
' DEPENDENCIES
'   - M_STATS_CORE
'   - M_STATS_SPECIALFUNC
'   - M_STATS_PROBDIST_NORMALFAMILY
'   - M_STATS_PROBDIST_TFAMILY
'
' NOTES
'   - Public distribution functions return Variant and may return CVErr, so the
'     assertion helpers accept Variant and route through IsError.
'   - Reference constants were computed in 30- to 60-digit arithmetic and are
'     quoted to 16 or 17 significant figures.
'   - AssertClose compares on absolute tolerance and is used where the expected
'     value is of order one. AssertRelClose compares on relative tolerance and is
'     the only meaningful test for a survival probability of order 1E-37, a
'     quantile of order 1E+34, or a normal quantile at Probability = 1E-300.
'   - TOL_TIGHT is the machine-precision target. TOL_LOOSE covers the raw
'     (unrefined) fast inverse and the lognormal moment round-trips.
'     TOL_LARGE_DF covers the log-gamma error floor above df ~ 1E+4.
'
' REGRESSION REGISTRY
'   Each assertion tagged REGRESSION below fails on a defect that was actually
'   shipped. Do not weaken their tolerances.
'
'   Core / SpecialFunc
'     C1  Log1p at X = 1E-8, the seam of the old Taylor threshold, where the old
'         implementation carried a 6E-9 relative error.
'     C2  The gamma series and beta continued fraction must report
'         non-convergence rather than returning a partial sum.
'
'   NormalFamily
'     N1  PROB_NormalInvCDFRaw must return its value. A find/replace once ate the
'         "=" sign in the Acklam kernel, so it silently returned 0# on every
'         branch: InverseCumulative(0.975) came back 1.1906 instead of 1.9600,
'         and InverseCumulativeFast returned 0# for every input.
'     N2  The Halley step must be skipped once PROB_NormalCDF has saturated.
'         Past Abs(Z) = 37 the residual degenerates and the step turned a 9.7E-11
'         relative error at Probability = 1E-300 into 4.9E-04.
'     N3  Lognormal_Variance and Lognormal_StdDev must return 0 when the
'         exponential underflows. VBA's And is not short-circuit, so the old
'         one-line overflow guard divided by zero and returned CVErr(xlErrValue).
'     N4  PROB_INV_SQRT_TWO_PI must be the correctly rounded Double. The old
'         literal was 5 ulp low, biasing every normal density by 7E-16.
'
'   TFamily
'     T1  Chi-square cumulative at df = 1600, 5000, 1E+6. The old 200-iteration
'         incomplete gamma series silently returned its partial sum: a 37 percent
'         error at df = 1E+5.
'     T2  Student t cumulative at x = 1E-8. The old beta argument rounded to
'         exactly 1 and the CDF collapsed to exactly 0.5, losing eight digits.
'     T3  Student t survival at x = 20, df = 30. 1 - CDF is exactly zero there.
'     T4  Student t inverse at p = 1E-14, df = 1. The old 1E+12 bracket cap
'         refused a legal input whose answer is -3.18E+13.
'     T5  Student t inverse near the median, p = 0.5 + 1E-10. The old CDF was flat
'         at 0.5 across |x| < 1E-8 and the quantile came back 30 times too large.
'
' UPDATED
'   2026-07-09
'==============================================================================

'==============================================================================
' MODULE-LEVEL TEST STATE
'==============================================================================

Private mTestCount          As Long            'Total assertions executed
Private mPassCount          As Long            'Assertions passed
Private mFailCount          As Long            'Assertions failed

'==============================================================================
' TEST TOLERANCES
'==============================================================================

Private Const TOL_TIGHT     As Double = 0.0000000001      '1E-10, absolute, machine-precision paths
Private Const TOL_LOOSE     As Double = 0.000001          '1E-6,  absolute, fast inverse / moment round-trips
Private Const TOL_LARGE_DF  As Double = 0.000000001       '1E-9,  absolute, large-df log-gamma error floor
Private Const TOL_ULP       As Double = 1E-16              '1E-16, absolute, correctly-rounded constants
Private Const TOL_FEW_ULP   As Double = 5E-16              '5E-16, absolute, constants vs a runtime-computed reference

Private Const TOL_REL_TIGHT As Double = 0.0000000001      '1E-10, relative, tails and quantiles
Private Const TOL_REL_LOOSE As Double = 0.000001          '1E-6,  relative, extreme-parameter corners
Private Const TOL_REL_TAIL  As Double = 0.000000001       '1E-9,  relative, normal quantile past the CDF split


'==============================================================================
' PUBLIC ENTRY POINTS
'==============================================================================

Public Sub Test_STATS_PROBDIST_RunAll()
'
'==============================================================================
' Test_STATS_PROBDIST_RunAll
'------------------------------------------------------------------------------
' PURPOSE
'   Runs every suite in dependency order and prints one PASS/FAIL summary.
'==============================================================================
'
    BeginRun "ALL SUITES"
    RunCoreSuite
    RunNormalFamilySuite
    RunTFamilySuite
    RunContinuousSuite
    EndRun
End Sub


Public Sub Test_STATS_PROBDIST_RunCore()
'
'==============================================================================
' Test_STATS_PROBDIST_RunCore
'------------------------------------------------------------------------------
' PURPOSE
'   Runs the M_STATS_CORE and M_STATS_SPECIALFUNC suite only.
'==============================================================================
'
    BeginRun "M_STATS_CORE + M_STATS_SPECIALFUNC"
    RunCoreSuite
    EndRun
End Sub


Public Sub Test_STATS_PROBDIST_RunNormalFamily()
'
'==============================================================================
' Test_STATS_PROBDIST_RunNormalFamily
'------------------------------------------------------------------------------
' PURPOSE
'   Runs the M_STATS_PROBDIST_NORMALFAMILY suite only.
'==============================================================================
'
    BeginRun "M_STATS_PROBDIST_NORMALFAMILY"
    RunNormalFamilySuite
    EndRun
End Sub


Public Sub Test_STATS_PROBDIST_RunTFamily()
'
'==============================================================================
' Test_STATS_PROBDIST_RunTFamily
'------------------------------------------------------------------------------
' PURPOSE
'   Runs the M_STATS_PROBDIST_TFAMILY suite only.
'==============================================================================
'
    BeginRun "M_STATS_PROBDIST_TFAMILY"
    RunTFamilySuite
    EndRun
End Sub


Public Sub Test_STATS_PROBDIST_RunContinuous()
'
'==============================================================================
' Test_STATS_PROBDIST_RunContinuous
'------------------------------------------------------------------------------
' PURPOSE
'   Runs the M_STATS_PROBDIST_CONTINUOUS suite only.
'==============================================================================
'
    BeginRun "M_STATS_PROBDIST_CONTINUOUS"
    RunContinuousSuite
    EndRun
End Sub


'==============================================================================
' SUITE DRIVERS
'==============================================================================

Private Sub RunCoreSuite()
    Debug.Print "== SUITE: M_STATS_CORE + M_STATS_SPECIALFUNC"
    Test_Core_Constants
    Test_Core_Log1p
    Test_Core_TryExp
    Test_Core_LogGamma
    Test_Core_NormalInvRaw
End Sub


Private Sub RunNormalFamilySuite()
    Debug.Print "== SUITE: M_STATS_PROBDIST_NORMALFAMILY"
    Test_NF_StandardNormalDensity
    Test_NF_StandardNormalCumulative
    Test_NF_StandardNormalInverse
    Test_NF_InverseTails
    Test_NF_InverseRoundTrips
    Test_NF_Symmetry
    Test_NF_GeneralNormal
    Test_NF_ZScore
    Test_NF_IntervalProbability
    Test_NF_FastInverse
    Test_NF_LognormalCore
    Test_NF_LognormalMoments
    Test_NF_LognormalUnderflow
    Test_NF_ParameterRoundTrip
    Test_NF_ErrorContract
    Test_NF_OverflowContract
End Sub


Private Sub RunTFamilySuite()
    Debug.Print "== SUITE: M_STATS_PROBDIST_TFAMILY"
    Test_TF_StudentTDensity
    Test_TF_StudentTCumulative
    Test_TF_StudentTCentralRegion
    Test_TF_StudentTSurvival
    Test_TF_StudentTInverse
    Test_TF_StudentTRoundTrips
    Test_TF_StudentTSymmetry
    Test_TF_ChiSquareDensity
    Test_TF_ChiSquareCumulative
    Test_TF_ChiSquareLargeDF
    Test_TF_ChiSquareSurvival
    Test_TF_ChiSquareInverse
    Test_TF_FDensity
    Test_TF_FCumulative
    Test_TF_FSurvival
    Test_TF_FInverse
    Test_TF_ErrorContract
    Test_TF_SupportEdges
End Sub


Private Sub RunContinuousSuite()
    Debug.Print "== SUITE: M_STATS_PROBDIST_CONTINUOUS"
    Test_CN_GammaDensity
    Test_CN_GammaCumulative
    Test_CN_GammaSurvival
    Test_CN_GammaInverse
    Test_CN_GammaMoments
    Test_CN_BetaDensity
    Test_CN_BetaCumulative
    Test_CN_BetaSurvival
    Test_CN_BetaInverse
    Test_CN_BetaMoments
    Test_CN_ExponentialDensity
    Test_CN_ExponentialCumulative
    Test_CN_ExponentialSurvival
    Test_CN_ExponentialInverse
    Test_CN_WeibullDensity
    Test_CN_WeibullCumulative
    Test_CN_WeibullSurvival
    Test_CN_WeibullInverse
    Test_CN_WeibullMoments
    Test_CN_UniformDensity
    Test_CN_UniformCumulative
    Test_CN_UniformSurvival
    Test_CN_UniformInverse
    Test_CN_CrossFamilyIdentities
    Test_CN_RoundTrips
    Test_CN_ErrorContract
    Test_CN_SupportEdges
End Sub


'==============================================================================
' RUN HARNESS
'==============================================================================

Private Sub BeginRun( _
    ByVal Title As String)
'
'==============================================================================
' BeginRun
'------------------------------------------------------------------------------
' PURPOSE
'   Resets the counters and prints the run header.
'==============================================================================
'
    'Reset counters
        mTestCount = 0
        mPassCount = 0
        mFailCount = 0

    'Header
        Debug.Print String(70, "=")
        Debug.Print Title & " - test run " & Format(Now, "yyyy-mm-dd hh:nn:ss")
        Debug.Print String(70, "=")
End Sub


Private Sub EndRun()
'
'==============================================================================
' EndRun
'------------------------------------------------------------------------------
' PURPOSE
'   Prints the PASS/FAIL summary for the run.
'==============================================================================
'
    'Summary
        Debug.Print String(70, "-")
        Debug.Print "TOTAL  " & mTestCount & _
                    "   PASS " & mPassCount & _
                    "   FAIL " & mFailCount

    'Verdict
        If mFailCount = 0 Then
            Debug.Print "RESULT: ALL TESTS PASSED"
        Else
            Debug.Print "RESULT: " & mFailCount & " TEST(S) FAILED"
        End If

        Debug.Print String(70, "=")
End Sub


'==============================================================================
' SUITE - CORE AND SPECIAL FUNCTIONS
'==============================================================================

Private Sub Test_Core_Constants()
    Debug.Print "-- Core constants"

    'Pi must be the correctly rounded Double, not the 14-figure truncation
    AssertClose "PROB_PI", PROB_PI, 4# * Atn(1#), TOL_ULP

    'Half-log-two-pi, used by the Lanczos log-gamma. The reference is itself
    'computed at run time from two rounded operations, so it is good to a few ulp
    'only; TOL_ULP would fail on the correctly rounded constant.
    AssertClose "PROB_HALF_LOG_TWO_PI", PROB_HALF_LOG_TWO_PI, _
        0.5 * Log(2# * PROB_PI), TOL_FEW_ULP
End Sub


Private Sub Test_Core_Log1p()
    Debug.Print "-- Core PROB_Log1p"

    'REGRESSION C1: the old Log1p handed off to Log(1 + X) at |X| < 1E-8 and was
    'therefore 6E-9 wrong at exactly that seam
    AssertRelClose "Log1p(1e-8)", PROB_Log1p(0.00000001), _
        9.99999995E-09, TOL_REL_TIGHT
    AssertRelClose "Log1p(1e-12)", PROB_Log1p(0.000000000001), _
        9.999999999995E-13, TOL_REL_TIGHT
    AssertRelClose "Log1p(1)", PROB_Log1p(1#), 0.693147180559945, TOL_REL_TIGHT
    AssertRelClose "Log1p(1e300)", PROB_Log1p(1E+300), 690.775527898214, TOL_REL_TIGHT
End Sub


Private Sub Test_Core_TryExp()
    Dim ExpResult As Double

    Debug.Print "-- Core PROB_TryExp contract"

    'Overflow must fail, never return a clamped sentinel
    AssertTrue "TryExp overflow rejected", (Not PROB_TryExp(710#, ExpResult))

    'Underflow is a valid zero, not a failure
    AssertTrue "TryExp underflow accepted", PROB_TryExp(-1000#, ExpResult)
    AssertClose "TryExp underflow value", ExpResult, 0#, 0#

    'Regular exponential
    AssertTrue "TryExp regular accepted", PROB_TryExp(1#, ExpResult)
    AssertClose "TryExp regular value", ExpResult, 2.71828182845905, TOL_TIGHT
End Sub


Private Sub Test_Core_LogGamma()
    Debug.Print "-- Core log-gamma"

    'Known exact values
    AssertClose "LogGamma(1)", PROB_LogGamma(1#), 0#, 0.00000000000001
    AssertClose "LogGamma(0.5)", PROB_LogGamma(0.5), 0.5723649429247, 0.00000000000001
    AssertClose "LogGamma(10)", PROB_LogGamma(10#), 12.8018274800815, 0.0000000001

    'The half-difference must not cancel at large argument. Written as a literal
    'subtraction of two log-gammas this carries a 5.9E-11 relative error at
    'Z = 5E+5, because two numbers of size 6.4E+6 produce an answer of size 6.6.
    AssertRelClose "LogGammaHalfDiff(3)", PROB_LogGammaHalfDiff(3#), _
        0.507826421787129, TOL_REL_TIGHT
    AssertRelClose "LogGammaHalfDiff(500000)", PROB_LogGammaHalfDiff(500000#), _
        6.56118143870216, TOL_REL_TIGHT

    'Log-beta half-integer shortcut must agree with the general route
    AssertRelClose "LogBeta(5, 0.5)", PROB_LogBeta(5#, 0.5), _
        -0.207395194346071, TOL_REL_TIGHT
    AssertRelClose "LogBeta(2, 3)", PROB_LogBeta(2#, 3#), _
        -2.484906649788, TOL_REL_TIGHT
End Sub


Private Sub Test_Core_NormalInvRaw()
    Debug.Print "-- Core PROB_NormalInvCDFRaw (shared seed kernel)"

    'REGRESSION N1: the raw kernel must actually return its value. When the "="
    'was lost from the assignment it returned 0# on every branch, and every
    'inverse-normal caller in the project silently produced garbage.
    AssertTrue "raw kernel is not a zero stub", (Abs(PROB_NormalInvCDFRaw(0.975)) > 1#)

    'Raw Acklam accuracy is ~1.15E-9, so a loose tolerance is correct here
    AssertClose "raw InvPhi(0.975)", PROB_NormalInvCDFRaw(0.975), 1.95996398454005, TOL_LOOSE
    AssertClose "raw InvPhi(0.025)", PROB_NormalInvCDFRaw(0.025), -1.95996398454005, TOL_LOOSE
    AssertClose "raw InvPhi(0.5)", PROB_NormalInvCDFRaw(0.5), 0#, TOL_LOOSE

    'Each of the three branches must be exercised
    AssertTrue "raw lower branch", (PROB_NormalInvCDFRaw(0.001) < -3#)
    AssertTrue "raw upper branch", (PROB_NormalInvCDFRaw(0.999) > 3#)
End Sub


'==============================================================================
' SUITE - NORMAL FAMILY
'==============================================================================

Private Sub Test_NF_StandardNormalDensity()
    Debug.Print "-- Standard normal density"

    'REGRESSION N4: at Z = 0 the density IS the constant, so this assertion pins
    'PROB_INV_SQRT_TWO_PI to the correctly rounded Double. The old literal
    '0.398942280401433 is 5 ulp high and misses by 2.8E-16.
    AssertClose "phi(0) correctly rounded", _
        K_STATS_NormalStandard_Density(0#), 0.398942280401433, TOL_ULP

    AssertClose "phi(1)", K_STATS_NormalStandard_Density(1#), 0.241970724519143, TOL_TIGHT
    AssertClose "phi(-1)", K_STATS_NormalStandard_Density(-1#), 0.241970724519143, TOL_TIGHT
    AssertClose "phi(2)", K_STATS_NormalStandard_Density(2#), 0.053990966513188, TOL_TIGHT

    'Symmetry is exact, not approximate
    AssertClose "phi symmetric", _
        CDbl(K_STATS_NormalStandard_Density(-1.7)) - _
        CDbl(K_STATS_NormalStandard_Density(1.7)), 0#, 0#

    'Far tail underflows to a valid zero
    AssertClose "phi far tail = 0", K_STATS_NormalStandard_Density(50#), 0#, 0#
End Sub


Private Sub Test_NF_StandardNormalCumulative()
    Debug.Print "-- Standard normal cumulative"
    AssertClose "Phi(0)", K_STATS_NormalStandard_Cumulative(0#), 0.5, TOL_TIGHT
    AssertClose "Phi(1)", K_STATS_NormalStandard_Cumulative(1#), 0.841344746068543, TOL_TIGHT
    AssertClose "Phi(2)", K_STATS_NormalStandard_Cumulative(2#), 0.977249868051821, TOL_TIGHT
    AssertClose "Phi(-1)", K_STATS_NormalStandard_Cumulative(-1#), 0.158655253931457, TOL_TIGHT
    AssertClose "Phi(1.959963984540054)", _
        K_STATS_NormalStandard_Cumulative(1.95996398454005), 0.975, TOL_TIGHT

    'Both sides of the rational / continued-fraction split at Z = Sqr(50)
    AssertClose "Phi split seam", _
        CDbl(K_STATS_NormalStandard_Cumulative(-7.07106781)) - _
        CDbl(K_STATS_NormalStandard_Cumulative(-7.07106782)), 0#, 1E-16

    'Saturation past the tail cutoff
    AssertClose "Phi(-40) saturates", K_STATS_NormalStandard_Cumulative(-40#), 0#, 0#
    AssertClose "Phi(40) saturates", K_STATS_NormalStandard_Cumulative(40#), 1#, 0#
End Sub


Private Sub Test_NF_StandardNormalInverse()
    Debug.Print "-- Standard normal inverse"

    'REGRESSION N1: the zero-stub Acklam kernel returned 1.1906 here
    AssertClose "InvPhi(0.975)", _
        K_STATS_NormalStandard_InverseCumulative(0.975), 1.95996398454005, TOL_TIGHT

    AssertClose "InvPhi(0.5)", K_STATS_NormalStandard_InverseCumulative(0.5), 0#, TOL_TIGHT
    AssertClose "InvPhi(0.95)", _
        K_STATS_NormalStandard_InverseCumulative(0.95), 1.64485362695147, TOL_TIGHT
    AssertClose "InvPhi(0.025)", _
        K_STATS_NormalStandard_InverseCumulative(0.025), -1.95996398454005, TOL_TIGHT
    AssertClose "InvPhi(0.005)", _
        K_STATS_NormalStandard_InverseCumulative(0.005), -2.5758293035489, TOL_TIGHT
End Sub


Private Sub Test_NF_InverseTails()
    Debug.Print "-- Standard normal inverse, deep tails"

    'REGRESSION N2: past Abs(Z) = PROB_CDF_TAIL_CUTOFF the CDF returns exactly 0,
    'so the Halley residual degenerates to -Probability. Refining anyway turned
    'a 9.7E-11 relative error into 4.9E-04; the guard restores the raw estimate.
    AssertRelClose "InvPhi(1e-300)", _
        K_STATS_NormalStandard_InverseCumulative(1E-300), _
        -37.0470962993614, TOL_REL_TAIL

    'Inside the cutoff, Halley works and the accuracy floor is the RELATIVE
    'accuracy of PROB_NormalCDF past its Sqr(50) split, which is about 1E-10.
    'The Hart/West arrangement is 1E-15 in ABSOLUTE terms, not relative.
    AssertRelClose "InvPhi(1e-100)", _
        K_STATS_NormalStandard_InverseCumulative(1E-100), _
        -21.2734535609655, TOL_REL_TAIL
    AssertRelClose "InvPhi(1e-20)", _
        K_STATS_NormalStandard_InverseCumulative(1E-20), _
        -9.26234008979836, TOL_REL_TAIL
    AssertRelClose "InvPhi(1e-6)", _
        K_STATS_NormalStandard_InverseCumulative(0.000001), _
        -4.75342430882289, TOL_REL_TIGHT
End Sub


Private Sub Test_NF_InverseRoundTrips()
    Dim ZValues As Variant
    Dim I As Long
    Dim Z As Double
    Dim P As Variant
    Dim Back As Variant

    Debug.Print "-- Inverse round-trips  InvPhi(Phi(z)) = z"
    ZValues = Array(-3#, -2#, -0.5, 0#, 0.5, 2#, 3#)
    For I = LBound(ZValues) To UBound(ZValues)
        Z = ZValues(I)
        P = K_STATS_NormalStandard_Cumulative(Z)
        If IsError(P) Then
            RecordResult "roundtrip z=" & Z & " (CDF errored)", False
        Else
            Back = K_STATS_NormalStandard_InverseCumulative(CDbl(P))
            AssertClose "roundtrip z=" & Z, Back, Z, TOL_TIGHT
        End If
    Next I
End Sub


Private Sub Test_NF_Symmetry()
    Dim ZValues As Variant
    Dim I As Long
    Dim Z As Double
    Dim Lo As Variant
    Dim Hi As Variant

    Debug.Print "-- Symmetry  Phi(z) + Phi(-z) = 1"
    ZValues = Array(0.25, 1#, 2.5, 4#)
    For I = LBound(ZValues) To UBound(ZValues)
        Z = ZValues(I)
        Lo = K_STATS_NormalStandard_Cumulative(-Z)
        Hi = K_STATS_NormalStandard_Cumulative(Z)
        If IsError(Lo) Or IsError(Hi) Then
            RecordResult "symmetry z=" & Z & " (errored)", False
        Else
            AssertClose "symmetry z=" & Z, CDbl(Lo) + CDbl(Hi), 1#, TOL_TIGHT
        End If
    Next I

    'Inverse antisymmetry
    AssertClose "InvPhi(p) = -InvPhi(1-p)", _
        CDbl(K_STATS_NormalStandard_InverseCumulative(0.02)) + _
        CDbl(K_STATS_NormalStandard_InverseCumulative(0.98)), 0#, TOL_TIGHT
End Sub


Private Sub Test_NF_GeneralNormal()
    Debug.Print "-- General normal (mean=10, sd=2)"
    'Density at the mean equals phi(0)/sd
    AssertClose "density@mean", _
        K_STATS_Normal_Density(10#, 10#, 2#), 0.398942280401433 / 2#, TOL_TIGHT
    'CDF at the mean is 0.5
    AssertClose "cdf@mean", K_STATS_Normal_Cumulative(10#, 10#, 2#), 0.5, TOL_TIGHT
    'CDF one sd above the mean equals Phi(1)
    AssertClose "cdf@mean+sd", _
        K_STATS_Normal_Cumulative(12#, 10#, 2#), 0.841344746068543, TOL_TIGHT
    'Inverse at 0.975 is mean + 1.959963984540054 * sd
    AssertClose "inv@0.975", _
        K_STATS_Normal_InverseCumulative(0.975, 10#, 2#), _
        10# + 1.95996398454005 * 2#, TOL_TIGHT
End Sub


Private Sub Test_NF_ZScore()
    Debug.Print "-- Z-score"
    AssertClose "z(10,4,2)", K_STATS_Normal_ZScore(10#, 4#, 2#), 3#, TOL_TIGHT
    AssertClose "z(4,4,2)", K_STATS_Normal_ZScore(4#, 4#, 2#), 0#, TOL_TIGHT
    AssertClose "z(1,4,2)", K_STATS_Normal_ZScore(1#, 4#, 2#), -1.5, TOL_TIGHT
End Sub


Private Sub Test_NF_IntervalProbability()
    Debug.Print "-- Interval probability"
    'Standard: symmetric 95% band
    AssertClose "std P(-1.96..1.96)=0.95", _
        K_STATS_NormalStandard_IntervalProbability(-1.95996398454005, 1.95996398454005), _
        0.95, TOL_TIGHT
    'Standard: +/- 1 sd
    AssertClose "std P(-1..1)", _
        K_STATS_NormalStandard_IntervalProbability(-1#, 1#), 0.682689492137086, TOL_TIGHT
    'Standard: equal bounds -> 0
    AssertClose "std P(1..1)=0", _
        K_STATS_NormalStandard_IntervalProbability(1#, 1#), 0#, TOL_TIGHT
    'General: mean 10 sd 2, 8..12 equals P(-1..1)
    AssertClose "gen P(8..12 | 10,2)", _
        K_STATS_Normal_IntervalProbability(8#, 12#, 10#, 2#), 0.682689492137086, TOL_TIGHT
    'General: full standard band via defaults
    AssertClose "gen P(-1..1 | 0,1)", _
        K_STATS_Normal_IntervalProbability(-1#, 1#), 0.682689492137086, TOL_TIGHT
End Sub


Private Sub Test_NF_FastInverse()
    Dim R As Double

    Debug.Print "-- Fast inverse (raw Acklam, ~1E-9)"

    'REGRESSION N1: the zero-stub kernel made this return 0# for every input,
    'which would have made every Monte Carlo shock identical
    AssertTrue "fast inverse is not a zero stub", _
        (Abs(K_STATS_NormalStandard_InverseCumulativeFast(0.975)) > 1#)

    'Accuracy against known quantile (loose tolerance)
    AssertClose "fast InvPhi(0.975)", _
        K_STATS_NormalStandard_InverseCumulativeFast(0.975), 1.95996398454005, TOL_LOOSE
    AssertClose "fast InvPhi(0.5)", _
        K_STATS_NormalStandard_InverseCumulativeFast(0.5), 0#, TOL_LOOSE
    'Endpoint clipping: p=0 must not error and must return a large negative number
    R = K_STATS_NormalStandard_InverseCumulativeFast(0#)
    AssertTrue "fast InvPhi(0) clipped negative", (R < -5#)
    'Endpoint clipping: p=1 must return a large positive number
    R = K_STATS_NormalStandard_InverseCumulativeFast(1#)
    AssertTrue "fast InvPhi(1) clipped positive", (R > 5#)
End Sub


Private Sub Test_NF_LognormalCore()
    Debug.Print "-- Lognormal core (MeanLog=0, StdDevLog=1)"
    'Density at x=1 equals phi(0)/(1*1)
    AssertClose "logn density@1", _
        K_STATS_Lognormal_Density(1#, 0#, 1#), 0.398942280401433, TOL_TIGHT
    'CDF at x=1 equals Phi(0)=0.5
    AssertClose "logn cdf@1", K_STATS_Lognormal_Cumulative(1#, 0#, 1#), 0.5, TOL_TIGHT
    'CDF at x<=0 returns 0
    AssertClose "logn cdf@0", K_STATS_Lognormal_Cumulative(0#, 0#, 1#), 0#, TOL_TIGHT
    AssertClose "logn cdf@-5", K_STATS_Lognormal_Cumulative(-5#, 0#, 1#), 0#, TOL_TIGHT
    'Inverse at p=0.5 returns Exp(0)=1
    AssertClose "logn inv@0.5", _
        K_STATS_Lognormal_InverseCumulative(0.5, 0#, 1#), 1#, TOL_TIGHT
End Sub


Private Sub Test_NF_LognormalMoments()
    Dim VarV As Variant
    Dim SdV As Variant

    Debug.Print "-- Lognormal moments (MeanLog=0, StdDevLog=1)"
    'Mean = Exp(0.5)
    AssertClose "logn mean", _
        K_STATS_Lognormal_Mean(0#, 1#), 1.64872127070013, TOL_TIGHT
    'Variance = (e-1)*e
    AssertClose "logn variance", _
        K_STATS_Lognormal_Variance(0#, 1#), 4.6707742704716, TOL_TIGHT
    'StdDev = Sqr(variance)
    AssertClose "logn stddev", _
        K_STATS_Lognormal_StdDev(0#, 1#), 2.16119741589509, TOL_TIGHT
    'Consistency: StdDev^2 == Variance
    VarV = K_STATS_Lognormal_Variance(0#, 1#)
    SdV = K_STATS_Lognormal_StdDev(0#, 1#)
    If IsError(VarV) Or IsError(SdV) Then
        RecordResult "logn stddev^2 == variance (errored)", False
    Else
        AssertClose "logn stddev^2 == variance", CDbl(SdV) * CDbl(SdV), CDbl(VarV), TOL_LOOSE
    End If
End Sub


Private Sub Test_NF_LognormalUnderflow()
    Debug.Print "-- Lognormal underflow (REGRESSION N3)"

    'VBA's And is not short-circuit. The old one-line guard
    '    If ExpShift > 0# And Factor > PROB_DOUBLE_MAX / ExpShift Then
    'evaluated the division even when ExpShift had underflowed to zero, raising
    'run-time error 11 and returning CVErr(xlErrValue) on a perfectly valid input.
    'The stated policy is that underflow of an exponential is a valid zero.
    AssertClose "logn variance underflow -> 0", _
        K_STATS_Lognormal_Variance(-400#, 1#), 0#, 0#
    AssertClose "logn variance deep underflow -> 0", _
        K_STATS_Lognormal_Variance(-1000#, 1#), 0#, 0#
    AssertClose "logn stddev underflow -> 0", _
        K_STATS_Lognormal_StdDev(-800#, 1#), 0#, 0#
    AssertClose "logn mean underflow -> 0", _
        K_STATS_Lognormal_Mean(-1000#, 1#), 0#, 0#

    'Overflow must still fail loudly; the guard is not simply removed
    AssertIsError "logn variance still overflows", K_STATS_Lognormal_Variance(1000#, 1#)
End Sub


Private Sub Test_NF_ParameterRoundTrip()
    Dim P As Variant
    Dim MeanLog As Double
    Dim StdDevLog As Double

    Debug.Print "-- Parameter round-trip (Mean=2, StdDev=0.5)"
    P = K_STATS_Lognormal_ParametersFromMeanStdDev(2#, 0.5)
    If IsError(P) Then
        RecordResult "param conversion returned error", False
        Exit Sub
    End If

    MeanLog = P(1, 1)
    StdDevLog = P(1, 2)

    'Check the converted log-space parameters
    AssertClose "MeanLog", MeanLog, 0.662834806151016, TOL_LOOSE
    AssertClose "StdDevLog", StdDevLog, 0.246221445044987, TOL_LOOSE

    'Round-trip: feeding the log params back must recover Mean and StdDev.
    'This ties K_STATS_Lognormal_StdDev to the conversion.
    AssertClose "roundtrip Mean", _
        K_STATS_Lognormal_Mean(MeanLog, StdDevLog), 2#, TOL_LOOSE
    AssertClose "roundtrip StdDev", _
        K_STATS_Lognormal_StdDev(MeanLog, StdDevLog), 0.5, TOL_LOOSE
End Sub


Private Sub Test_NF_ErrorContract()
    Dim Diag As String

    Debug.Print "-- Normal family error contract (invalid domains must return CVErr)"
    'Zero / negative standard deviation
    AssertIsError "normal density sd=0", K_STATS_Normal_Density(0#, 0#, 0#)
    AssertIsError "normal cdf sd<0", K_STATS_Normal_Cumulative(0#, 0#, -1#)
    'Probability at or outside the open unit interval
    AssertIsError "inv p=0", K_STATS_NormalStandard_InverseCumulative(0#)
    AssertIsError "inv p=1", K_STATS_NormalStandard_InverseCumulative(1#)
    AssertIsError "inv p=1.5", K_STATS_NormalStandard_InverseCumulative(1.5)
    'Lognormal domain
    AssertIsError "logn density x=0", K_STATS_Lognormal_Density(0#, 0#, 1#)
    AssertIsError "logn density sd=0", K_STATS_Lognormal_Density(1#, 0#, 0#)
    'Parameter conversion rejects StdDev = 0 and non-positive Mean
    AssertIsError "param StdDev=0", K_STATS_Lognormal_ParametersFromMeanStdDev(2#, 0#)
    AssertIsError "param Mean=0", K_STATS_Lognormal_ParametersFromMeanStdDev(0#, 1#)
    'Reversed interval bounds
    AssertIsError "std interval reversed", _
        K_STATS_NormalStandard_IntervalProbability(1#, -1#)
    AssertIsError "gen interval reversed", _
        K_STATS_Normal_IntervalProbability(5#, 1#, 0#, 1#)

    'Status must be populated on failure and cleared on success
    Diag = "stale"
    AssertIsError "normal cdf sd=0 with status", K_STATS_Normal_Cumulative(0#, 0#, 0#, Diag)
    AssertTrue "NF status populated on failure", (Len(Diag) > 0 And Diag <> "stale")

    Diag = "stale"
    AssertClose "normal cdf ok with status", K_STATS_Normal_Cumulative(0#, 0#, 1#, Diag), 0.5, TOL_TIGHT
    AssertTrue "NF status cleared on success", (Len(Diag) = 0)
End Sub


Private Sub Test_NF_OverflowContract()
    Debug.Print "-- Normal family overflow contract (CVErr(xlErrNum), not a sentinel)"
    'Exp argument well beyond the Double range
    AssertIsError "logn mean overflow", K_STATS_Lognormal_Mean(1000#, 1#)
    AssertIsError "logn variance overflow", K_STATS_Lognormal_Variance(1000#, 1#)
    AssertIsError "logn stddev overflow", K_STATS_Lognormal_StdDev(1000#, 1#)
    'Quantile that exponentiates past the Double range
    AssertIsError "logn inverse overflow", _
        K_STATS_Lognormal_InverseCumulative(0.999999999, 700#, 20#)
    'Underflow must remain a valid (tiny/zero) result, not an error
    AssertTrue "logn mean underflow ok", _
        (Not IsError(K_STATS_Lognormal_Mean(-1000#, 1#)))
End Sub


'==============================================================================
' SUITE - T FAMILY: STUDENT T
'==============================================================================

Private Sub Test_TF_StudentTDensity()
    Debug.Print "-- Student t density"
    AssertClose "t pdf(0,1)", K_STATS_StudentT_Density(0#, 1#), 0.318309886183791, TOL_TIGHT
    AssertClose "t pdf(0,5)", K_STATS_StudentT_Density(0#, 5#), 0.379606689822494, TOL_TIGHT
    AssertClose "t pdf(1,5)", K_STATS_StudentT_Density(1#, 5#), 0.219679797350981, TOL_TIGHT
    AssertClose "t pdf(2,10)", K_STATS_StudentT_Density(2#, 10#), 6.11457663212182E-02, TOL_TIGHT
    AssertClose "t pdf(0,30)", K_STATS_StudentT_Density(0#, 30#), 0.395632184894098, TOL_TIGHT

    'Symmetry
    AssertClose "t pdf symmetric", _
        CDbl(K_STATS_StudentT_Density(-2#, 7#)) - CDbl(K_STATS_StudentT_Density(2#, 7#)), 0#, 0#

    'Large df must converge on the standard normal density; the old log-gamma
    'subtraction lost seven digits here
    AssertClose "t pdf(0,1e6) -> phi(0)", K_STATS_StudentT_Density(0#, 1000000#), _
        0.398942180665875, TOL_LARGE_DF

    'Far tail underflows to a valid zero, not an error
    AssertClose "t pdf far tail = 0", K_STATS_StudentT_Density(1E+50, 30#), 0#, 0#
End Sub


Private Sub Test_TF_StudentTCumulative()
    Debug.Print "-- Student t cumulative"
    AssertClose "t cdf(0,5)", K_STATS_StudentT_Cumulative(0#, 5#), 0.5, 0#
    AssertClose "t cdf(1,5)", K_STATS_StudentT_Cumulative(1#, 5#), 0.818391266175439, TOL_TIGHT
    AssertClose "t cdf(2,10)", K_STATS_StudentT_Cumulative(2#, 10#), 0.96330598261463, TOL_TIGHT
    AssertClose "t cdf(-1,1)", K_STATS_StudentT_Cumulative(-1#, 1#), 0.25, TOL_TIGHT
    AssertClose "t cdf(2.5,3)", K_STATS_StudentT_Cumulative(2.5, 3#), 0.956146676495967, TOL_TIGHT

    'The df = 2 closed form must agree with the general beta route at df = 2.0000001
    AssertClose "t cdf df=2 continuity", _
        CDbl(K_STATS_StudentT_Cumulative(1.5, 2#)) - _
        CDbl(K_STATS_StudentT_Cumulative(1.5, 2.0000001)), 0#, 0.0000001
End Sub


Private Sub Test_TF_StudentTCentralRegion()
    Debug.Print "-- Student t cumulative, central region (REGRESSION T2)"

    'The old implementation returned exactly 0.5 for every |x| below about 1E-8,
    'because df/(df + x*x) rounded to exactly 1. These three assertions fail on
    'the old code by roughly 4E-9 absolute.
    AssertRelClose "t cdf(-1e-8,1)", K_STATS_StudentT_Cumulative(-0.00000001, 1#), _
        0.499999996816901, TOL_REL_TIGHT
    AssertRelClose "t cdf(-1e-8,5)", K_STATS_StudentT_Cumulative(-0.00000001, 5#), _
        0.499999996203933, TOL_REL_TIGHT
    AssertRelClose "t cdf(1e-8,30)", K_STATS_StudentT_Cumulative(0.00000001, 30#), _
        0.500000003956322, TOL_REL_TIGHT

    'The CDF must be strictly monotone across the origin, not flat
    AssertTrue "t cdf strictly increasing at 0", _
        (CDbl(K_STATS_StudentT_Cumulative(0.000000001, 5#)) > _
         CDbl(K_STATS_StudentT_Cumulative(0#, 5#)))
End Sub


Private Sub Test_TF_StudentTSurvival()
    Debug.Print "-- Student t survival (REGRESSION T3)"

    'These are the p-values that 1 - CDF cannot express
    AssertRelClose "t sf(20,30)", K_STATS_StudentT_Survival(20#, 30#), _
        3.37454183289E-19, TOL_REL_TIGHT
    AssertRelClose "t sf(10,5)", K_STATS_StudentT_Survival(10#, 5#), _
        8.54737878715E-05, TOL_REL_TIGHT
    AssertRelClose "t sf(30,5)", K_STATS_StudentT_Survival(30#, 5#), _
        3.85932431025E-07, TOL_REL_TIGHT

    'The CDF-based route really does collapse, which is why Survival exists
    AssertTrue "1 - cdf(20,30) is exactly zero", _
        ((1# - CDbl(K_STATS_StudentT_Cumulative(20#, 30#))) = 0#)

    'Survival and cumulative must sum to one in the well-conditioned region
    AssertClose "t sf + cdf = 1", _
        CDbl(K_STATS_StudentT_Survival(1.3, 9#)) + _
        CDbl(K_STATS_StudentT_Cumulative(1.3, 9#)), 1#, TOL_TIGHT

    'Negative arguments are accepted, unlike Excel's T.DIST.RT
    AssertClose "t sf(-1,1)", K_STATS_StudentT_Survival(-1#, 1#), 0.75, TOL_TIGHT
End Sub


Private Sub Test_TF_StudentTInverse()
    Debug.Print "-- Student t inverse"
    AssertClose "t inv(0.975,10)", K_STATS_StudentT_InverseCumulative(0.975, 10#), _
        2.22813885198627, TOL_TIGHT
    AssertClose "t inv(0.95,5)", K_STATS_StudentT_InverseCumulative(0.95, 5#), _
        2.01504837333302, TOL_TIGHT
    AssertClose "t inv(0.005,20)", K_STATS_StudentT_InverseCumulative(0.005, 20#), _
        -2.84533970978611, TOL_TIGHT
    AssertClose "t inv(0.5,7)", K_STATS_StudentT_InverseCumulative(0.5, 7#), 0#, 0#

    'REGRESSION T4: the old 1E+12 bracket cap returned CVErr(xlErrNum) here
    AssertRelClose "t inv(1e-14,1)", K_STATS_StudentT_InverseCumulative(0.00000000000001, 1#), _
        -31830988618379.1, TOL_REL_TIGHT
    AssertTrue "t inv(1e-14,0.5) is a number", _
        (Not IsError(K_STATS_StudentT_InverseCumulative(0.00000000000001, 0.5)))

    'REGRESSION T5: the old bisection could not resolve inside the flat spot and
    'returned about 1.05E-8 instead of 3.14E-10. The expected value is the exact
    'quantile of the Double 0.5000000001, not of the decimal 0.5 + 1E-10.
    AssertRelClose "t inv(0.5+1e-10,1)", _
        K_STATS_StudentT_InverseCumulative(0.5000000001, 1#), _
        3.14159291352633E-10, TOL_REL_TIGHT

    'The same case with a df that has no closed form; the median inverse is
    'conditioned by the spacing of Doubles near 0.5, so 1E-6 relative is the floor
    AssertRelClose "t inv(0.5+1e-10,5)", _
        K_STATS_StudentT_InverseCumulative(0.5000000001, 5#), _
        2.63430571834E-10, TOL_REL_LOOSE

    'Exact closed form at df = 2
    AssertRelClose "t inv(0.975,2)", K_STATS_StudentT_InverseCumulative(0.975, 2#), _
        4.30265272974946, TOL_REL_TIGHT
End Sub


Private Sub Test_TF_StudentTRoundTrips()
    Dim Quantile As Variant

    Debug.Print "-- Student t inverse round-trips"
    AssertRoundTripT "t round-trip p=0.001 df=3", 0.001, 3#
    AssertRoundTripT "t round-trip p=0.025 df=8", 0.025, 8#
    AssertRoundTripT "t round-trip p=0.3 df=1", 0.3, 1#
    AssertRoundTripT "t round-trip p=0.7 df=2", 0.7, 2#
    AssertRoundTripT "t round-trip p=0.999 df=40", 0.999, 40#
    AssertRoundTripT "t round-trip p=0.975 df=1000", 0.975, 1000#

    'Survival round-trip in the far tail, the case the CDF cannot see
    Quantile = K_STATS_StudentT_InverseCumulative(0.00000001, 6#)
    AssertRelClose "t sf(inv(1e-8,6)) = 1e-8", _
        K_STATS_StudentT_Survival(-CDbl(Quantile), 6#), 0.00000001, TOL_REL_TIGHT
End Sub


Private Sub Test_TF_StudentTSymmetry()
    Debug.Print "-- Student t symmetry"
    AssertClose "t cdf(-x) = 1 - cdf(x)", _
        CDbl(K_STATS_StudentT_Cumulative(-1.7, 11#)) + _
        CDbl(K_STATS_StudentT_Cumulative(1.7, 11#)), 1#, TOL_TIGHT
    AssertClose "t inv(p) = -inv(1-p)", _
        CDbl(K_STATS_StudentT_InverseCumulative(0.02, 13#)) + _
        CDbl(K_STATS_StudentT_InverseCumulative(0.98, 13#)), 0#, TOL_TIGHT
    AssertClose "t sf(x) = cdf(-x)", _
        CDbl(K_STATS_StudentT_Survival(2.2, 4#)) - _
        CDbl(K_STATS_StudentT_Cumulative(-2.2, 4#)), 0#, 0#
End Sub


'==============================================================================
' SUITE - T FAMILY: CHI-SQUARE
'==============================================================================

Private Sub Test_TF_ChiSquareDensity()
    Debug.Print "-- Chi-square density"
    AssertClose "chi2 pdf(1,2)", K_STATS_ChiSquare_Density(1#, 2#), 0.303265329856317, TOL_TIGHT
    AssertClose "chi2 pdf(4,4)", K_STATS_ChiSquare_Density(4#, 4#), 0.135335283236613, TOL_TIGHT
    AssertClose "chi2 pdf(2,1)", K_STATS_ChiSquare_Density(2#, 1#), 0.103776874355149, TOL_TIGHT
    AssertClose "chi2 pdf(-1,3)", K_STATS_ChiSquare_Density(-1#, 3#), 0#, 0#
    AssertClose "chi2 pdf(0,2)", K_STATS_ChiSquare_Density(0#, 2#), 0.5, 0#
    AssertClose "chi2 pdf(0,3)", K_STATS_ChiSquare_Density(0#, 3#), 0#, 0#
    AssertIsError "chi2 pdf(0,1) unbounded", K_STATS_ChiSquare_Density(0#, 1#)
End Sub


Private Sub Test_TF_ChiSquareCumulative()
    Debug.Print "-- Chi-square cumulative"
    AssertClose "chi2 cdf(3.84,1)", K_STATS_ChiSquare_Cumulative(3.84, 1#), _
        0.949956478751295, TOL_TIGHT
    AssertClose "chi2 cdf(11.07,5)", K_STATS_ChiSquare_Cumulative(11.07, 5#), _
        0.949990381377595, TOL_TIGHT
    AssertClose "chi2 cdf(1,2)", K_STATS_ChiSquare_Cumulative(1#, 2#), _
        0.393469340287367, TOL_TIGHT
    AssertClose "chi2 cdf(0,5)", K_STATS_ChiSquare_Cumulative(0#, 5#), 0#, 0#
    AssertClose "chi2 cdf(-3,5)", K_STATS_ChiSquare_Cumulative(-3#, 5#), 0#, 0#

    'Series and continued-fraction branches must agree across the x = a + 1 seam
    AssertClose "chi2 branch seam", _
        CDbl(K_STATS_ChiSquare_Cumulative(11.99999999, 10#)) - _
        CDbl(K_STATS_ChiSquare_Cumulative(12.00000001, 10#)), 0#, 0.00000001
End Sub


Private Sub Test_TF_ChiSquareLargeDF()
    Debug.Print "-- Chi-square, large degrees of freedom (REGRESSION T1)"

    'The old 200-iteration series stopped converging at about df = 1600 and
    'returned its partial sum with no diagnostic. Errors on the old code:
    '   df = 1600   8.8E-12      df = 5000   7.4E-05
    '   df = 10000  4.8E-03      df = 1E+5   3.7E-01
    AssertRelClose "chi2 cdf(1600,1600)", K_STATS_ChiSquare_Cumulative(1600#, 1600#), _
        0.504701612421641, TOL_REL_TIGHT
    AssertRelClose "chi2 cdf(5000,5000)", K_STATS_ChiSquare_Cumulative(5000#, 5000#), _
        0.502659621107655, TOL_REL_TIGHT
    AssertRelClose "chi2 cdf(1e6,1e6)", K_STATS_ChiSquare_Cumulative(1000000#, 1000000#), _
        0.500188063196606, TOL_REL_LOOSE

    'A converged answer, or an honest error - never a silent partial sum
    AssertTrue "chi2 cdf(1e6) is not a partial sum", _
        (Abs(CDbl(K_STATS_ChiSquare_Cumulative(1000000#, 1000000#)) - 0.5) < 0.001)
End Sub


Private Sub Test_TF_ChiSquareSurvival()
    Debug.Print "-- Chi-square survival"
    AssertRelClose "chi2 sf(200,10)", K_STATS_ChiSquare_Survival(200#, 10#), _
        1.6139305337E-37, TOL_REL_TIGHT
    AssertRelClose "chi2 sf(100,1)", K_STATS_ChiSquare_Survival(100#, 1#), _
        1.52397060483E-23, TOL_REL_TIGHT
    AssertRelClose "chi2 sf(3.84,1)", K_STATS_ChiSquare_Survival(3.84, 1#), _
        0.0500435212487, TOL_REL_TIGHT
    AssertClose "chi2 sf(0,5)", K_STATS_ChiSquare_Survival(0#, 5#), 1#, 0#
    AssertClose "chi2 sf + cdf = 1", _
        CDbl(K_STATS_ChiSquare_Survival(7#, 4#)) + _
        CDbl(K_STATS_ChiSquare_Cumulative(7#, 4#)), 1#, TOL_TIGHT

    'The CDF-based route collapses; that is why Survival exists
    AssertTrue "1 - chi2 cdf(200,10) is exactly zero", _
        ((1# - CDbl(K_STATS_ChiSquare_Cumulative(200#, 10#))) = 0#)
End Sub


Private Sub Test_TF_ChiSquareInverse()
    Debug.Print "-- Chi-square inverse"
    AssertClose "chi2 inv(0.95,1)", K_STATS_ChiSquare_InverseCumulative(0.95, 1#), _
        3.84145882069412, TOL_TIGHT
    AssertClose "chi2 inv(0.95,5)", K_STATS_ChiSquare_InverseCumulative(0.95, 5#), _
        11.0704976935164, TOL_TIGHT
    AssertClose "chi2 inv(0.5,4)", K_STATS_ChiSquare_InverseCumulative(0.5, 4#), _
        3.35669398003332, TOL_TIGHT
    AssertRelClose "chi2 inv(0.99,5000)", K_STATS_ChiSquare_InverseCumulative(0.99, 5000#), _
        5235.57183813011, TOL_REL_TIGHT

    'Round-trips, including the far upper tail where a CDF-side solve would be
    'driving two nearly equal numbers together
    AssertRoundTripChi2 "chi2 round-trip p=1e-9 df=3", 0.000000001, 3#
    AssertRoundTripChi2 "chi2 round-trip p=0.025 df=17", 0.025, 17#
    AssertRoundTripChi2 "chi2 round-trip p=0.999999999 df=2", 0.999999999, 2#
    AssertRoundTripChi2 "chi2 round-trip p=0.5 df=0.5", 0.5, 0.5
End Sub


'==============================================================================
' SUITE - T FAMILY: F
'==============================================================================

Private Sub Test_TF_FDensity()
    Debug.Print "-- F density"
    AssertClose "F pdf(1,10,10)", K_STATS_F_Density(1#, 10#, 10#), 0.615234375, TOL_TIGHT
    AssertClose "F pdf(2,4,8)", K_STATS_F_Density(2#, 4#, 8#), 0.15625, TOL_TIGHT
    AssertClose "F pdf(-1,3,3)", K_STATS_F_Density(-1#, 3#, 3#), 0#, 0#
    AssertClose "F pdf(0,2,5)", K_STATS_F_Density(0#, 2#, 5#), 1#, 0#
    AssertClose "F pdf(0,3,5)", K_STATS_F_Density(0#, 3#, 5#), 0#, 0#
    AssertIsError "F pdf(0,1,5) unbounded", K_STATS_F_Density(0#, 1#, 5#)
End Sub


Private Sub Test_TF_FCumulative()
    Debug.Print "-- F cumulative"
    AssertClose "F cdf(2.5,5,10)", K_STATS_F_Cumulative(2.5, 5#, 10#), _
        0.89799772335573, TOL_TIGHT
    AssertClose "F cdf(1,1,1)", K_STATS_F_Cumulative(1#, 1#, 1#), 0.5, TOL_TIGHT
    AssertClose "F cdf(4.96,3,10)", K_STATS_F_Cumulative(4.96, 3#, 10#), _
        0.976863670854344, TOL_TIGHT
    AssertClose "F cdf(0,4,4)", K_STATS_F_Cumulative(0#, 4#, 4#), 0#, 0#
    AssertClose "F cdf(-2,4,4)", K_STATS_F_Cumulative(-2#, 4#, 4#), 0#, 0#

    'The reciprocal identity: F(x; a, b) = 1 - F(1/x; b, a)
    AssertClose "F reciprocal identity", _
        CDbl(K_STATS_F_Cumulative(3#, 6#, 9#)) + _
        CDbl(K_STATS_F_Cumulative(1# / 3#, 9#, 6#)), 1#, TOL_TIGHT

    'Large equal degrees of freedom: the old 200-iteration beta continued fraction
    'silently stopped converging at about df = 5E+5
    AssertRelClose "F cdf(1,1e5,1e5)", K_STATS_F_Cumulative(1#, 100000#, 100000#), _
        0.5, TOL_REL_LOOSE
End Sub


Private Sub Test_TF_FSurvival()
    Debug.Print "-- F survival"
    AssertRelClose "F sf(100,5,5)", K_STATS_F_Survival(100#, 5#, 5#), _
        5.24291335785E-05, TOL_REL_TIGHT
    AssertRelClose "F sf(2.5,5,10)", K_STATS_F_Survival(2.5, 5#, 10#), _
        0.10200227664427, TOL_REL_TIGHT
    AssertClose "F sf(0,4,4)", K_STATS_F_Survival(0#, 4#, 4#), 1#, 0#
    AssertClose "F sf + cdf = 1", _
        CDbl(K_STATS_F_Survival(2#, 7#, 12#)) + _
        CDbl(K_STATS_F_Cumulative(2#, 7#, 12#)), 1#, TOL_TIGHT
End Sub


Private Sub Test_TF_FInverse()
    Debug.Print "-- F inverse"
    AssertClose "F inv(0.95,3,10)", K_STATS_F_InverseCumulative(0.95, 3#, 10#), _
        3.70826481904684, TOL_TIGHT
    AssertClose "F inv(0.95,5,20)", K_STATS_F_InverseCumulative(0.95, 5#, 20#), _
        2.71088983720969, TOL_TIGHT
    AssertClose "F inv(0.5,1,1)", K_STATS_F_InverseCumulative(0.5, 1#, 1#), 1#, TOL_TIGHT

    'The extreme upper tail. The answer is 8.4623534263E+34; it survives only
    'because the beta solver returns both the root and its complement.
    AssertRelClose "F inv(1-1e-9,0.5,0.5)", _
        K_STATS_F_InverseCumulative(0.999999999, 0.5, 0.5), 8.4623534263E+34, TOL_REL_LOOSE

    'Round-trips
    AssertRoundTripF "F round-trip p=0.001 df=(2,3)", 0.001, 2#, 3#
    AssertRoundTripF "F round-trip p=0.5 df=(1,1)", 0.5, 1#, 1#
    AssertRoundTripF "F round-trip p=0.99 df=(10,4)", 0.99, 10#, 4#
    AssertRoundTripF "F round-trip p=0.975 df=(100,1000)", 0.975, 100#, 1000#
End Sub


'==============================================================================
' SUITE - T FAMILY: CONTRACT
'==============================================================================

Private Sub Test_TF_ErrorContract()
    Dim Diag As String

    Debug.Print "-- T family error contract (must return CVErr, never a sentinel)"

    'Non-positive or non-finite degrees of freedom
    AssertIsError "t pdf df=0", K_STATS_StudentT_Density(1#, 0#)
    AssertIsError "t cdf df<0", K_STATS_StudentT_Cumulative(1#, -3#)
    AssertIsError "t sf df=0", K_STATS_StudentT_Survival(1#, 0#)
    AssertIsError "chi2 cdf df=0", K_STATS_ChiSquare_Cumulative(1#, 0#)
    AssertIsError "chi2 sf df<0", K_STATS_ChiSquare_Survival(1#, -1#)
    AssertIsError "F cdf df1=0", K_STATS_F_Cumulative(1#, 0#, 5#)
    AssertIsError "F cdf df2<0", K_STATS_F_Cumulative(1#, 5#, -2#)
    AssertIsError "F sf df1=0", K_STATS_F_Survival(1#, 0#, 5#)

    'Non-finite evaluation points
    AssertIsError "t cdf x too large", K_STATS_StudentT_Cumulative(1E+200, 5#)
    AssertIsError "chi2 cdf x too large", K_STATS_ChiSquare_Cumulative(1E+200, 5#)

    'Probabilities outside the open unit interval
    AssertIsError "t inv p=0", K_STATS_StudentT_InverseCumulative(0#, 5#)
    AssertIsError "t inv p=1", K_STATS_StudentT_InverseCumulative(1#, 5#)
    AssertIsError "t inv p>1", K_STATS_StudentT_InverseCumulative(1.5, 5#)
    AssertIsError "chi2 inv p=0", K_STATS_ChiSquare_InverseCumulative(0#, 5#)
    AssertIsError "chi2 inv p=1", K_STATS_ChiSquare_InverseCumulative(1#, 5#)
    AssertIsError "F inv p=0", K_STATS_F_InverseCumulative(0#, 5#, 5#)
    AssertIsError "F inv p=1", K_STATS_F_InverseCumulative(1#, 5#, 5#)

    'Status must be populated on failure and cleared on success
    Diag = "stale"
    AssertIsError "t cdf df=0 with status", K_STATS_StudentT_Cumulative(1#, 0#, Diag)
    AssertTrue "TF status populated on failure", (Len(Diag) > 0 And Diag <> "stale")

    Diag = "stale"
    AssertClose "t cdf ok with status", K_STATS_StudentT_Cumulative(0#, 5#, Diag), 0.5, 0#
    AssertTrue "TF status cleared on success", (Len(Diag) = 0)
End Sub


Private Sub Test_TF_SupportEdges()
    Debug.Print "-- T family support edges and non-integer degrees of freedom"

    'Fractional degrees of freedom are legal
    AssertTrue "t cdf df=0.5 ok", (Not IsError(K_STATS_StudentT_Cumulative(1#, 0.5)))
    AssertTrue "chi2 cdf df=0.1 ok", (Not IsError(K_STATS_ChiSquare_Cumulative(1#, 0.1)))
    AssertTrue "F cdf df=(0.5,0.5) ok", (Not IsError(K_STATS_F_Cumulative(1#, 0.5, 0.5)))

    'Every probability lies in the closed unit interval
    AssertInUnitInterval "t cdf(-40,3)", K_STATS_StudentT_Cumulative(-40#, 3#)
    AssertInUnitInterval "t sf(40,3)", K_STATS_StudentT_Survival(40#, 3#)
    AssertInUnitInterval "chi2 cdf(1e9,2)", K_STATS_ChiSquare_Cumulative(1000000000#, 2#)
    AssertInUnitInterval "chi2 sf(1e-9,2)", K_STATS_ChiSquare_Survival(0.000000001, 2#)
    AssertInUnitInterval "F cdf(1e9,3,3)", K_STATS_F_Cumulative(1000000000#, 3#, 3#)
    AssertInUnitInterval "F sf(1e-9,3,3)", K_STATS_F_Survival(0.000000001, 3#, 3#)

    'Chi-square with df = 2 is exponential: cdf(x) = 1 - Exp(-x/2)
    AssertClose "chi2 df=2 is exponential", _
        CDbl(K_STATS_ChiSquare_Cumulative(3#, 2#)), 1# - Exp(-1.5), TOL_TIGHT

    'Student t with df = 1 is Cauchy: cdf(x) = 0.5 + Atn(x)/Pi
    AssertClose "t df=1 is Cauchy", _
        CDbl(K_STATS_StudentT_Cumulative(2.7, 1#)), 0.5 + Atn(2.7) / (4# * Atn(1#)), TOL_TIGHT
End Sub


'==============================================================================
' ROUND-TRIP HELPERS
'==============================================================================

Private Sub AssertRoundTripT( _
    ByVal TestName As String, _
    ByVal Probability As Double, _
    ByVal DegreesFreedom As Double)
'
'==============================================================================
' AssertRoundTripT
'------------------------------------------------------------------------------
' PURPOSE
'   Passes when Cumulative(InverseCumulative(p)) reproduces p to TOL_REL_TIGHT.
'==============================================================================
'
    'Declare
        Dim Quantile As Variant
        Dim Recovered As Variant

    'Invert then re-evaluate
        Quantile = K_STATS_StudentT_InverseCumulative(Probability, DegreesFreedom)
        If IsError(Quantile) Then
            RecordResult TestName & " -> inverse returned error", False
            Exit Sub
        End If

        Recovered = K_STATS_StudentT_Cumulative(CDbl(Quantile), DegreesFreedom)
        AssertRelClose TestName, Recovered, Probability, TOL_REL_TIGHT
End Sub


Private Sub AssertRoundTripChi2( _
    ByVal TestName As String, _
    ByVal Probability As Double, _
    ByVal DegreesFreedom As Double)
'
'==============================================================================
' AssertRoundTripChi2
'------------------------------------------------------------------------------
' PURPOSE
'   Passes when Cumulative(InverseCumulative(p)) reproduces p to TOL_REL_TIGHT.
'==============================================================================
'
    'Declare
        Dim Quantile As Variant
        Dim Recovered As Variant

    'Invert then re-evaluate
        Quantile = K_STATS_ChiSquare_InverseCumulative(Probability, DegreesFreedom)
        If IsError(Quantile) Then
            RecordResult TestName & " -> inverse returned error", False
            Exit Sub
        End If

        Recovered = K_STATS_ChiSquare_Cumulative(CDbl(Quantile), DegreesFreedom)
        AssertRelClose TestName, Recovered, Probability, TOL_REL_TIGHT
End Sub


Private Sub AssertRoundTripF( _
    ByVal TestName As String, _
    ByVal Probability As Double, _
    ByVal DegreesFreedom1 As Double, _
    ByVal DegreesFreedom2 As Double)
'
'==============================================================================
' AssertRoundTripF
'------------------------------------------------------------------------------
' PURPOSE
'   Passes when Cumulative(InverseCumulative(p)) reproduces p to TOL_REL_TIGHT.
'==============================================================================
'
    'Declare
        Dim Quantile As Variant
        Dim Recovered As Variant

    'Invert then re-evaluate
        Quantile = K_STATS_F_InverseCumulative(Probability, DegreesFreedom1, DegreesFreedom2)
        If IsError(Quantile) Then
            RecordResult TestName & " -> inverse returned error", False
            Exit Sub
        End If

        Recovered = K_STATS_F_Cumulative(CDbl(Quantile), DegreesFreedom1, DegreesFreedom2)
        AssertRelClose TestName, Recovered, Probability, TOL_REL_TIGHT
End Sub


'==============================================================================
' ASSERTION HELPERS
'==============================================================================

Private Sub AssertClose( _
    ByVal TestName As String, _
    ByVal Actual As Variant, _
    ByVal Expected As Double, _
    ByVal Tolerance As Double)
'
'==============================================================================
' AssertClose
'------------------------------------------------------------------------------
' PURPOSE
'   Passes when Actual is a number within an absolute Tolerance of Expected. A
'   CVErr Actual is treated as a failure.
'==============================================================================
'
    'Reject error returns outright
        If IsError(Actual) Then
            RecordResult TestName & " -> returned error, expected " & Expected, False
            Exit Sub
        End If
    'Compare within tolerance
        If Abs(CDbl(Actual) - Expected) <= Tolerance Then
            RecordResult TestName, True
        Else
            RecordResult TestName & " -> got " & CDbl(Actual) & _
                         ", expected " & Expected & " (tol " & Tolerance & ")", False
        End If
End Sub


Private Sub AssertRelClose( _
    ByVal TestName As String, _
    ByVal Actual As Variant, _
    ByVal Expected As Double, _
    ByVal RelativeTolerance As Double)
'
'==============================================================================
' AssertRelClose
'------------------------------------------------------------------------------
' PURPOSE
'   Passes when Actual is a number within a relative RelativeTolerance of
'   Expected. This is the only meaningful comparison for a survival probability
'   of order 1E-37, a quantile of order 1E+34, or a normal quantile at
'   Probability = 1E-300, where an absolute tolerance is either vacuous or
'   unsatisfiable.
'==============================================================================
'
    'Declare
        Dim Difference As Double

    'Reject error returns outright
        If IsError(Actual) Then
            RecordResult TestName & " -> returned error, expected " & Expected, False
            Exit Sub
        End If

    'Fall back to an absolute comparison at zero
        If Expected = 0# Then
            AssertClose TestName, Actual, 0#, RelativeTolerance
            Exit Sub
        End If

    'Compare within relative tolerance
        Difference = Abs(CDbl(Actual) - Expected) / Abs(Expected)

        If Difference <= RelativeTolerance Then
            RecordResult TestName, True
        Else
            RecordResult TestName & " -> got " & CDbl(Actual) & _
                         ", expected " & Expected & " (rel err " & Difference & _
                         ", tol " & RelativeTolerance & ")", False
        End If
End Sub


Private Sub AssertInUnitInterval( _
    ByVal TestName As String, _
    ByVal Actual As Variant)
'
'==============================================================================
' AssertInUnitInterval
'------------------------------------------------------------------------------
' PURPOSE
'   Passes when Actual is a number in the closed interval [0, 1].
'==============================================================================
'
    'Reject error returns outright
        If IsError(Actual) Then
            RecordResult TestName & " -> returned error, expected a probability", False
            Exit Sub
        End If

    'Check the closed unit interval
        If CDbl(Actual) >= 0# And CDbl(Actual) <= 1# Then
            RecordResult TestName, True
        Else
            RecordResult TestName & " -> got " & CDbl(Actual) & ", outside [0, 1]", False
        End If
End Sub


Private Sub AssertTrue( _
    ByVal TestName As String, _
    ByVal Condition As Boolean)
'
'==============================================================================
' AssertTrue
'------------------------------------------------------------------------------
' PURPOSE
'   Passes when Condition is True.
'==============================================================================
'
    RecordResult TestName, Condition
End Sub


Private Sub AssertIsError( _
    ByVal TestName As String, _
    ByVal Actual As Variant)
'
'==============================================================================
' AssertIsError
'------------------------------------------------------------------------------
' PURPOSE
'   Passes when Actual is a CVErr value (an expected domain / overflow failure).
'==============================================================================
'
    If IsError(Actual) Then
        RecordResult TestName, True
    Else
        RecordResult TestName & " -> expected error, got " & CStr(Actual), False
    End If
End Sub


Private Sub RecordResult( _
    ByVal TestName As String, _
    ByVal Passed As Boolean)
'
'==============================================================================
' RecordResult
'------------------------------------------------------------------------------
' PURPOSE
'   Updates the pass/fail counters and prints a line only on failure.
'==============================================================================
'
    'Count the assertion
        mTestCount = mTestCount + 1
    'Tally and report
        If Passed Then
            mPassCount = mPassCount + 1
        Else
            mFailCount = mFailCount + 1
            Debug.Print "   FAIL: " & TestName
        End If
End Sub




'==============================================================================
' CONTINUOUS FAMILY SECTIONS
'   Reference values pre-verified against Python mpmath at 50 significant digits
'   before this module was written. Kernel-driven values (Gamma / Beta CDF, both
'   inverses) carry the same tolerances as the T family; closed-form values
'   (Exponential, Weibull, Uniform, all moments) are checked to machine
'   precision. Left-tail values built on PROB_Expm1 / PROB_Log1p are checked
'   relatively, since an absolute tolerance would be vacuous at 1E-10.
'==============================================================================

Private Sub Test_CN_GammaDensity()
    Debug.Print "-- Gamma density"
    AssertClose "gamma pdf(3,2.5,1.5)", K_STATS_Gamma_Density(3#, 2.5, 1.5), _
        0.19196788093578, TOL_TIGHT
    AssertClose "gamma pdf(0.5,0.5,2)", K_STATS_Gamma_Density(0.5, 0.5, 2#), _
        0.439391289467722, TOL_TIGHT
    AssertClose "gamma pdf(-1,2.5,1.5)=0", K_STATS_Gamma_Density(-1#, 2.5, 1.5), 0#, 0#
End Sub


Private Sub Test_CN_GammaCumulative()
    Debug.Print "-- Gamma cumulative"
    AssertClose "gamma cdf(3,2.5,1.5)", K_STATS_Gamma_Cumulative(3#, 2.5, 1.5), _
        0.45058404864722, TOL_TIGHT
    AssertClose "gamma cdf(1e-6,0.5,2)", K_STATS_Gamma_Cumulative(0.000001, 0.5, 2#), _
        7.97884427822125E-04, TOL_TIGHT
    AssertClose "gamma cdf(0,2.5,1.5)=0", K_STATS_Gamma_Cumulative(0#, 2.5, 1.5), 0#, 0#
    AssertClose "gamma cdf(-5,2.5,1.5)=0", K_STATS_Gamma_Cumulative(-5#, 2.5, 1.5), 0#, 0#
End Sub


Private Sub Test_CN_GammaSurvival()
    Debug.Print "-- Gamma survival"
    AssertClose "gamma sf(3,2.5,1.5)", K_STATS_Gamma_Survival(3#, 2.5, 1.5), _
        0.54941595135278, TOL_TIGHT
    AssertClose "gamma sf(0,2.5,1.5)=1", K_STATS_Gamma_Survival(0#, 2.5, 1.5), 1#, 0#
    AssertClose "gamma sf(-5,2.5,1.5)=1", K_STATS_Gamma_Survival(-5#, 2.5, 1.5), 1#, 0#

    'CDF and survival must sum to one on the support
    AssertClose "gamma cdf+sf=1", _
        CDbl(K_STATS_Gamma_Cumulative(3#, 2.5, 1.5)) + _
        CDbl(K_STATS_Gamma_Survival(3#, 2.5, 1.5)), 1#, TOL_TIGHT
End Sub


Private Sub Test_CN_GammaInverse()
    Debug.Print "-- Gamma inverse"
    AssertRelClose "gamma inv(0.7,2.5,1.5)", K_STATS_Gamma_InverseCumulative(0.7, 2.5, 1.5), _
        4.54832248811618, TOL_REL_TIGHT
    AssertRelClose "gamma inv(0.05,2.5,1.5)", K_STATS_Gamma_InverseCumulative(0.05, 2.5, 1.5), _
        0.859107169546327, TOL_REL_TIGHT
End Sub


Private Sub Test_CN_GammaMoments()
    Debug.Print "-- Gamma moments"
    AssertClose "gamma mean(2.5,1.5)", K_STATS_Gamma_Mean(2.5, 1.5), 3.75, TOL_TIGHT
    AssertClose "gamma var(2.5,1.5)", K_STATS_Gamma_Variance(2.5, 1.5), 5.625, TOL_TIGHT
    AssertClose "gamma std(2.5,1.5)", K_STATS_Gamma_StdDev(2.5, 1.5), _
        2.37170824512628, TOL_TIGHT
End Sub


Private Sub Test_CN_BetaDensity()
    Debug.Print "-- Beta density"
    AssertClose "beta pdf(0.3,2,5)", K_STATS_Beta_Density(0.3, 2#, 5#), _
        2.1609, TOL_TIGHT
    AssertRelClose "beta pdf(0.999999,2,3)", K_STATS_Beta_Density(0.999999, 2#, 3#), _
        1.1999988E-11, TOL_REL_TAIL
    AssertClose "beta pdf(-0.1,2,5)=0", K_STATS_Beta_Density(-0.1, 2#, 5#), 0#, 0#
    AssertClose "beta pdf(1.1,2,5)=0", K_STATS_Beta_Density(1.1, 2#, 5#), 0#, 0#
End Sub


Private Sub Test_CN_BetaCumulative()
    Debug.Print "-- Beta cumulative"
    AssertClose "beta cdf(0.3,2,5)", K_STATS_Beta_Cumulative(0.3, 2#, 5#), _
        0.579825, TOL_TIGHT
    AssertClose "beta cdf(0,2,5)=0", K_STATS_Beta_Cumulative(0#, 2#, 5#), 0#, 0#
    AssertClose "beta cdf(1,2,5)=1", K_STATS_Beta_Cumulative(1#, 2#, 5#), 1#, 0#
    AssertClose "beta cdf(-0.2,2,5)=0", K_STATS_Beta_Cumulative(-0.2, 2#, 5#), 0#, 0#
    AssertClose "beta cdf(1.2,2,5)=1", K_STATS_Beta_Cumulative(1.2, 2#, 5#), 1#, 0#
End Sub


Private Sub Test_CN_BetaSurvival()
    Debug.Print "-- Beta survival"
    AssertClose "beta sf(0.3,2,5)", K_STATS_Beta_Survival(0.3, 2#, 5#), _
        0.420175, TOL_TIGHT
    AssertClose "beta sf(0,2,5)=1", K_STATS_Beta_Survival(0#, 2#, 5#), 1#, 0#
    AssertClose "beta sf(1,2,5)=0", K_STATS_Beta_Survival(1#, 2#, 5#), 0#, 0#

    'CDF and survival must sum to one on the support
    AssertClose "beta cdf+sf=1", _
        CDbl(K_STATS_Beta_Cumulative(0.3, 2#, 5#)) + _
        CDbl(K_STATS_Beta_Survival(0.3, 2#, 5#)), 1#, TOL_TIGHT
End Sub


Private Sub Test_CN_BetaInverse()
    Debug.Print "-- Beta inverse"
    AssertRelClose "beta inv(0.6,2,5)", K_STATS_Beta_InverseCumulative(0.6, 2#, 5#), _
        0.309444427545314, TOL_REL_TIGHT
End Sub


Private Sub Test_CN_BetaMoments()
    Debug.Print "-- Beta moments"
    AssertClose "beta mean(2,5)", K_STATS_Beta_Mean(2#, 5#), _
        0.285714285714286, TOL_TIGHT
    AssertClose "beta var(2,5)", K_STATS_Beta_Variance(2#, 5#), _
        2.55102040816327E-02, TOL_TIGHT
    AssertClose "beta std(2,5)", K_STATS_Beta_StdDev(2#, 5#), _
        0.159719141249985, TOL_TIGHT
End Sub


Private Sub Test_CN_ExponentialDensity()
    Debug.Print "-- Exponential density"
    AssertClose "exp pdf(1,2)", K_STATS_Exponential_Density(1#, 2#), _
        0.270670566473225, TOL_TIGHT
    AssertClose "exp pdf(0,2)=lambda", K_STATS_Exponential_Density(0#, 2#), 2#, TOL_TIGHT
    AssertClose "exp pdf(-1,2)=0", K_STATS_Exponential_Density(-1#, 2#), 0#, 0#
End Sub


Private Sub Test_CN_ExponentialCumulative()
    Debug.Print "-- Exponential cumulative"
    AssertClose "exp cdf(1,2)", K_STATS_Exponential_Cumulative(1#, 2#), _
        0.864664716763387, TOL_TIGHT
    'Left tail through PROB_Expm1: absolute tolerance would be vacuous here
    AssertRelClose "exp cdf(1e-10,1)", K_STATS_Exponential_Cumulative(0.0000000001, 1#), _
        9.9999999995E-11, TOL_REL_TAIL
    AssertClose "exp cdf(0,2)=0", K_STATS_Exponential_Cumulative(0#, 2#), 0#, 0#
    AssertClose "exp cdf(-1,2)=0", K_STATS_Exponential_Cumulative(-1#, 2#), 0#, 0#
End Sub


Private Sub Test_CN_ExponentialSurvival()
    Debug.Print "-- Exponential survival"
    AssertClose "exp sf(1,2)", K_STATS_Exponential_Survival(1#, 2#), _
        0.135335283236613, TOL_TIGHT
    AssertClose "exp sf(0,2)=1", K_STATS_Exponential_Survival(0#, 2#), 1#, 0#
    AssertClose "exp sf(-1,2)=1", K_STATS_Exponential_Survival(-1#, 2#), 1#, 0#

    'CDF and survival must sum to one
    AssertClose "exp cdf+sf=1", _
        CDbl(K_STATS_Exponential_Cumulative(1#, 2#)) + _
        CDbl(K_STATS_Exponential_Survival(1#, 2#)), 1#, TOL_TIGHT
End Sub


Private Sub Test_CN_ExponentialInverse()
    Debug.Print "-- Exponential inverse"
    AssertClose "exp inv(0.5,2)", K_STATS_Exponential_InverseCumulative(0.5, 2#), _
        0.346573590279973, TOL_TIGHT
    'Left tail through PROB_Log1p
    AssertRelClose "exp inv(1e-12,1)", K_STATS_Exponential_InverseCumulative(0.000000000001, 1#), _
        1.0000000000005E-12, TOL_REL_TAIL
End Sub


Private Sub Test_CN_WeibullDensity()
    Debug.Print "-- Weibull density"
    AssertClose "weibull pdf(1,1.5,2)", K_STATS_Weibull_Density(1#, 1.5, 2#), _
        0.372391688219422, TOL_TIGHT
    AssertClose "weibull pdf(-1,1.5,2)=0", K_STATS_Weibull_Density(-1#, 1.5, 2#), 0#, 0#
End Sub


Private Sub Test_CN_WeibullCumulative()
    Debug.Print "-- Weibull cumulative"
    AssertClose "weibull cdf(1,1.5,2)", K_STATS_Weibull_Cumulative(1#, 1.5, 2#), _
        0.29781149867344, TOL_TIGHT
    'Left tail through PROB_Expm1
    AssertRelClose "weibull cdf(1e-10,1,1)", K_STATS_Weibull_Cumulative(0.0000000001, 1#, 1#), _
        9.9999999995E-11, TOL_REL_TAIL
    AssertClose "weibull cdf(0,1.5,2)=0", K_STATS_Weibull_Cumulative(0#, 1.5, 2#), 0#, 0#
    AssertClose "weibull cdf(-3,1.5,2)=0", K_STATS_Weibull_Cumulative(-3#, 1.5, 2#), 0#, 0#
End Sub


Private Sub Test_CN_WeibullSurvival()
    Debug.Print "-- Weibull survival"
    AssertClose "weibull sf(1,1.5,2)", K_STATS_Weibull_Survival(1#, 1.5, 2#), _
        0.70218850132656, TOL_TIGHT
    AssertClose "weibull sf(0,1.5,2)=1", K_STATS_Weibull_Survival(0#, 1.5, 2#), 1#, 0#

    'CDF and survival must sum to one on the support
    AssertClose "weibull cdf+sf=1", _
        CDbl(K_STATS_Weibull_Cumulative(1#, 1.5, 2#)) + _
        CDbl(K_STATS_Weibull_Survival(1#, 1.5, 2#)), 1#, TOL_TIGHT
End Sub


Private Sub Test_CN_WeibullInverse()
    Debug.Print "-- Weibull inverse"
    AssertClose "weibull inv(0.4,1.5,2)", K_STATS_Weibull_InverseCumulative(0.4, 1.5, 2#), _
        1.27804195727092, TOL_TIGHT
End Sub


Private Sub Test_CN_WeibullMoments()
    Debug.Print "-- Weibull moments"
    AssertClose "weibull mean(1.5,2)", K_STATS_Weibull_Mean(1.5, 2#), _
        1.80549058590187, TOL_TIGHT
    AssertClose "weibull var(1.5,2)", K_STATS_Weibull_Variance(1.5, 2#), _
        1.50276113925573, TOL_TIGHT
    AssertClose "weibull std(1.5,2)", K_STATS_Weibull_StdDev(1.5, 2#), _
        1.22587158350935, TOL_TIGHT
End Sub


Private Sub Test_CN_UniformDensity()
    Debug.Print "-- Uniform density"
    AssertClose "uniform pdf(3,2,5)", K_STATS_Uniform_Density(3#, 2#, 5#), _
        0.333333333333333, TOL_TIGHT
    AssertClose "uniform pdf(2,2,5) edge", K_STATS_Uniform_Density(2#, 2#, 5#), _
        0.333333333333333, TOL_TIGHT
    AssertClose "uniform pdf(1,2,5)=0", K_STATS_Uniform_Density(1#, 2#, 5#), 0#, 0#
    AssertClose "uniform pdf(6,2,5)=0", K_STATS_Uniform_Density(6#, 2#, 5#), 0#, 0#
End Sub


Private Sub Test_CN_UniformCumulative()
    Debug.Print "-- Uniform cumulative"
    AssertClose "uniform cdf(3,2,5)", K_STATS_Uniform_Cumulative(3#, 2#, 5#), _
        0.333333333333333, TOL_TIGHT
    AssertClose "uniform cdf(1,2,5)=0", K_STATS_Uniform_Cumulative(1#, 2#, 5#), 0#, 0#
    AssertClose "uniform cdf(6,2,5)=1", K_STATS_Uniform_Cumulative(6#, 2#, 5#), 1#, 0#
End Sub


Private Sub Test_CN_UniformSurvival()
    Debug.Print "-- Uniform survival"
    AssertClose "uniform sf(3,2,5)", K_STATS_Uniform_Survival(3#, 2#, 5#), _
        0.666666666666667, TOL_TIGHT
    AssertClose "uniform sf(1,2,5)=1", K_STATS_Uniform_Survival(1#, 2#, 5#), 1#, 0#
    AssertClose "uniform sf(6,2,5)=0", K_STATS_Uniform_Survival(6#, 2#, 5#), 0#, 0#

    'CDF and survival must sum to one on the support
    AssertClose "uniform cdf+sf=1", _
        CDbl(K_STATS_Uniform_Cumulative(3#, 2#, 5#)) + _
        CDbl(K_STATS_Uniform_Survival(3#, 2#, 5#)), 1#, 0#
End Sub


Private Sub Test_CN_UniformInverse()
    Debug.Print "-- Uniform inverse"
    AssertClose "uniform inv(0.25,2,5)", K_STATS_Uniform_InverseCumulative(0.25, 2#, 5#), _
        2.75, TOL_TIGHT
End Sub


Private Sub Test_CN_CrossFamilyIdentities()
    Dim BetaArg As Double

    Debug.Print "-- Cross-family identities (self-checks against independent oracles)"

    'Identity 1 (marshalling): Chi-square(v) is Gamma(shape v/2, scale 2)
    AssertClose "id chi2(3,5) via TFAMILY", K_STATS_ChiSquare_Cumulative(3#, 5#), _
        0.300014164121372, TOL_TIGHT
    AssertClose "id chi2(3,5)=gamma(3,2.5,2)", K_STATS_Gamma_Cumulative(3#, 2.5, 2#), _
        0.300014164121372, TOL_TIGHT

    'Identity 2 (marshalling): F(d1,d2) maps to Beta(d1/2, d2/2) at d1 x /(d1 x + d2)
    BetaArg = 5# * 2.5 / (5# * 2.5 + 10#)
    AssertClose "id F(2.5,5,10) via TFAMILY", K_STATS_F_Cumulative(2.5, 5#, 10#), _
        0.89799772335573, TOL_TIGHT
    AssertClose "id F(2.5,5,10)=beta(arg,2.5,5)", K_STATS_Beta_Cumulative(BetaArg, 2.5, 5#), _
        0.89799772335573, TOL_TIGHT

    'Identity 3 (real kernel test): Exponential(rate L) is Gamma(1, 1/L).
    'Note the RECIPROCAL: rate 2 maps to scale 0.5. This pits the incomplete-gamma
    'kernel at shape = 1 against the closed-form PROB_Expm1 CDF.
    AssertClose "id exp(1,2) closed form", K_STATS_Exponential_Cumulative(1#, 2#), _
        0.864664716763387, TOL_TIGHT
    AssertClose "id exp(1,2)=gamma(1,1,0.5) kernel", K_STATS_Gamma_Cumulative(1#, 1#, 0.5), _
        0.864664716763387, TOL_TIGHT

    'Identity 4 (real kernel test): Uniform(0,1) is Beta(1,1); both equal x
    AssertClose "id uniform(0.37,0,1)", K_STATS_Uniform_Cumulative(0.37, 0#, 1#), _
        0.37, TOL_TIGHT
    AssertClose "id beta(0.37,1,1) kernel", K_STATS_Beta_Cumulative(0.37, 1#, 1#), _
        0.37, TOL_TIGHT

    'Identity 5 (real kernel test): Chi-square(2) is Exponential(rate 1/2). The
    'gamma kernel (via chi-square) is checked against the closed-form Exponential.
    AssertClose "id chi2(2.4,2) kernel", K_STATS_ChiSquare_Cumulative(2.4, 2#), _
        0.698805788087798, TOL_TIGHT
    AssertClose "id exp(2.4,0.5) closed form", K_STATS_Exponential_Cumulative(2.4, 0.5), _
        0.698805788087798, TOL_TIGHT
End Sub


Private Sub Test_CN_RoundTrips()
    Debug.Print "-- Continuous inverse round-trips (CDF of quantile returns the probability)"

    AssertClose "gamma roundtrip p=0.7", _
        CDbl(K_STATS_Gamma_Cumulative( _
            CDbl(K_STATS_Gamma_InverseCumulative(0.7, 2.5, 1.5)), 2.5, 1.5)), 0.7, TOL_LOOSE
    AssertClose "beta roundtrip p=0.6", _
        CDbl(K_STATS_Beta_Cumulative( _
            CDbl(K_STATS_Beta_InverseCumulative(0.6, 2#, 5#)), 2#, 5#)), 0.6, TOL_LOOSE
    AssertClose "exp roundtrip p=0.35", _
        CDbl(K_STATS_Exponential_Cumulative( _
            CDbl(K_STATS_Exponential_InverseCumulative(0.35, 2#)), 2#)), 0.35, TOL_TIGHT
    AssertClose "weibull roundtrip p=0.4", _
        CDbl(K_STATS_Weibull_Cumulative( _
            CDbl(K_STATS_Weibull_InverseCumulative(0.4, 1.5, 2#)), 1.5, 2#)), 0.4, TOL_TIGHT
    AssertClose "uniform roundtrip p=0.25", _
        CDbl(K_STATS_Uniform_Cumulative( _
            CDbl(K_STATS_Uniform_InverseCumulative(0.25, 2#, 5#)), 2#, 5#)), 0.25, TOL_TIGHT
End Sub


Private Sub Test_CN_ErrorContract()
    Dim Diag As String

    Debug.Print "-- Continuous error contract (must return CVErr, never a sentinel)"

    'Density poles at boundaries with shape < 1 return CVErr(xlErrNum)
    AssertIsError "gamma pdf pole(0,0.5,2)", K_STATS_Gamma_Density(0#, 0.5, 2#)
    AssertIsError "beta pdf pole(0,0.5,2)", K_STATS_Beta_Density(0#, 0.5, 2#)
    AssertIsError "beta pdf pole(1,2,0.5)", K_STATS_Beta_Density(1#, 2#, 0.5)
    AssertIsError "weibull pdf pole(0,0.5,2)", K_STATS_Weibull_Density(0#, 0.5, 2#)

    'Invalid parameters
    AssertIsError "gamma pdf shape<0", K_STATS_Gamma_Density(1#, -2#, 1.5)
    AssertIsError "gamma cdf scale=0", K_STATS_Gamma_Cumulative(1#, 2.5, 0#)
    AssertIsError "beta cdf alpha=0", K_STATS_Beta_Cumulative(0.3, 0#, 5#)
    AssertIsError "exp pdf lambda<0", K_STATS_Exponential_Density(1#, -2#)
    AssertIsError "weibull cdf shape=0", K_STATS_Weibull_Cumulative(1#, 0#, 2#)
    AssertIsError "uniform pdf hi<=lo", K_STATS_Uniform_Density(3#, 5#, 2#)
    AssertIsError "uniform pdf hi=lo", K_STATS_Uniform_Density(3#, 2#, 2#)

    'Non-finite evaluation points
    AssertIsError "gamma cdf x huge", K_STATS_Gamma_Cumulative(1E+200, 2.5, 1.5)

    'Probabilities outside the open unit interval
    AssertIsError "gamma inv p=0", K_STATS_Gamma_InverseCumulative(0#, 2.5, 1.5)
    AssertIsError "gamma inv p=1", K_STATS_Gamma_InverseCumulative(1#, 2.5, 1.5)
    AssertIsError "beta inv p>1", K_STATS_Beta_InverseCumulative(1.5, 2#, 5#)
    AssertIsError "exp inv p=1", K_STATS_Exponential_InverseCumulative(1#, 2#)
    AssertIsError "weibull inv p=0", K_STATS_Weibull_InverseCumulative(0#, 1.5, 2#)
    AssertIsError "uniform inv p=1", K_STATS_Uniform_InverseCumulative(1#, 2#, 5#)

    'Status must be populated on failure and cleared on success
    Diag = "stale"
    AssertIsError "gamma cdf scale=0 with status", K_STATS_Gamma_Cumulative(1#, 2.5, 0#, Diag)
    AssertTrue "CN status populated on failure", (Len(Diag) > 0 And Diag <> "stale")

    Diag = "stale"
    AssertClose "gamma cdf ok with status", K_STATS_Gamma_Cumulative(3#, 2.5, 1.5, Diag), _
        0.45058404864722, TOL_TIGHT
    AssertTrue "CN status cleared on success", (Len(Diag) = 0)
End Sub


Private Sub Test_CN_SupportEdges()
    Debug.Print "-- Continuous support edges (finite boundary densities)"

    'Gamma origin: 1/Scale at shape 1, zero above, both finite (not poles)
    AssertClose "gamma pdf(0,1,2)=1/2", K_STATS_Gamma_Density(0#, 1#, 2#), 0.5, TOL_TIGHT
    AssertClose "gamma pdf(0,2,1.5)=0", K_STATS_Gamma_Density(0#, 2#, 1.5), 0#, 0#

    'Beta endpoints: Beta at 0 when alpha=1, Alpha at 1 when beta=1
    AssertClose "beta pdf(0,1,3)=3", K_STATS_Beta_Density(0#, 1#, 3#), 3#, TOL_TIGHT
    AssertClose "beta pdf(1,2,1)=2", K_STATS_Beta_Density(1#, 2#, 1#), 2#, TOL_TIGHT
    AssertClose "beta pdf(0,2,5)=0", K_STATS_Beta_Density(0#, 2#, 5#), 0#, 0#
    AssertClose "beta pdf(1,2,5)=0", K_STATS_Beta_Density(1#, 2#, 5#), 0#, 0#

    'Weibull origin: 1/Scale at shape 1, zero above
    AssertClose "weibull pdf(0,1,2)=1/2", K_STATS_Weibull_Density(0#, 1#, 2#), 0.5, TOL_TIGHT
    AssertClose "weibull pdf(0,2,1.5)=0", K_STATS_Weibull_Density(0#, 2#, 1.5), 0#, 0#

    'All CDFs live in the unit interval
    AssertInUnitInterval "gamma cdf in [0,1]", K_STATS_Gamma_Cumulative(3#, 2.5, 1.5)
    AssertInUnitInterval "beta cdf in [0,1]", K_STATS_Beta_Cumulative(0.3, 2#, 5#)
    AssertInUnitInterval "exp cdf in [0,1]", K_STATS_Exponential_Cumulative(1#, 2#)
    AssertInUnitInterval "weibull cdf in [0,1]", K_STATS_Weibull_Cumulative(1#, 1.5, 2#)
    AssertInUnitInterval "uniform cdf in [0,1]", K_STATS_Uniform_Cumulative(3#, 2#, 5#)
End Sub


