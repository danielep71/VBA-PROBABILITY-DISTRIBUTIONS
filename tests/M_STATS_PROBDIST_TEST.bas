Attribute VB_Name = "M_STATS_PROBDIST_TEST"
Option Explicit

'==============================================================================
' M_STATS_PROBDIST_TEST
'------------------------------------------------------------------------------
' PURPOSE
'   Provides the single consolidated self-checking test harness for the complete
'   probability-distribution library:
'
'     - M_STATS_PROBDIST_CORE
'     - M_STATS_PROBDIST_SPECIALFUNCS
'     - M_STATS_PROBDIST_NORMALFAMILY
'     - M_STATS_PROBDIST_TFAMILY
'     - M_STATS_PROBDIST_CONTINUOUS
'
'   The harness verifies known values, support behavior, complement identities,
'   symmetry, inverse round-trips, extreme tails, moment formulas, numerical
'   overflow and underflow policy, diagnostic-status behavior, and exact CVErr
'   classification for predictable numerical failures.
'
' WHY THIS EXISTS
'   A numerical library needs one authoritative green-or-red result. Earlier
'   test modules duplicated counters and assertion helpers, which produced
'   fragmented summaries and allowed inconsistent error policies. This module
'   owns the counters, assertion layer, suite order, regression registry, and
'   final verdict for the entire distribution stack.
'
' HOW TO RUN
'   Run one of the following argument-less Public procedures from the Immediate
'   window, the Macros dialog, or another VBA procedure:
'
'       Test_STATS_PROBDIST_RunAll
'       Test_STATS_PROBDIST_RunCore
'       Test_STATS_PROBDIST_RunNormalFamily
'       Test_STATS_PROBDIST_RunTFamily
'       Test_STATS_PROBDIST_RunContinuous
'
'   Results are written with Debug.Print. Passing assertions are silent.
'   Failures print one detailed line, followed by a consolidated summary.
'
' SUITE ORDER
'   RunAll executes suites in dependency order:
'
'       1. Core and reusable special-function kernels
'       2. Normal and lognormal family
'       3. Student t, chi-square and F family
'       4. Gamma, Beta, Exponential, Weibull and Uniform family
'
' TEST DESIGN
'   - Public UDFs return Variant and may return CVErr, so assertion helpers
'     accept Variant and reject unexpected worksheet errors explicitly.
'   - Exact binary equality is used for constants that are intended to match
'     independently evaluated VBA Double values.
'   - Absolute tolerances are used for order-one values and exact support edges.
'   - Relative tolerances are used for deep tails, extreme quantiles and values
'     spanning many orders of magnitude.
'   - Error-code assertions distinguish #NUM! from #VALUE! where the numerical
'     contract requires predictable failures to return xlErrNum.
'   - Regression tolerances document the accuracy contract and must not be
'     weakened merely to make a failing implementation appear green.
'
' TOLERANCE POLICY
'   - TOL_ABS_TIGHT is a ten-decimal-place absolute tolerance. It is not a claim
'     of machine precision.
'   - TOL_ABS_LOOSE covers raw approximations and moment round-trips.
'   - TOL_ABS_LARGE_DF covers the large-parameter special-function error floor.
'   - TOL_ABS_ULP and TOL_ABS_FEW_ULP are used only for constant-level checks.
'   - TOL_REL_TIGHT, TOL_REL_LOOSE and TOL_REL_TAIL cover relative comparisons.
'
' REGRESSION REGISTRY
'   Core / Special Functions
'     C1  PROB_Log1p must remain accurate at the former Taylor-series seam.
'     C2  Incomplete-beta and incomplete-gamma kernels are tested directly for
'         known values, complements, inverse round-trips and paired arguments.
'     C3  PROB_LogBeta must remain stable for extremely unbalanced arguments.
'
'   Normal Family
'     N1  The raw Acklam inverse-normal kernel must return its computed value.
'     N2  Deep-tail inverse-normal refinement must not degrade saturated tails.
'     N3  Lognormal moments must treat exponential underflow as a valid zero.
'     N4  The normal density constant must equal the correctly rounded Double.
'     N5  Same-tail normal interval probabilities must not collapse to zero.
'     N6  Tiny lognormal variance and parameter conversion must use Expm1/Log1p.
'
'   T Family
'     T1  Large-degree chi-square calculations must never return partial sums.
'     T2  Student t central probabilities must retain displacement from 0.5.
'     T3  Student t survival must preserve probabilities invisible to 1 - CDF.
'     T4  Legal large Student t quantiles must not be rejected by an arbitrary
'         bracket ceiling.
'     T5  Near-median Student t quantiles must resolve the local CDF slope.
'     T6  Small-degree and tiny-argument branches must return #NUM! for genuine
'         overflow, never an unexpected #VALUE!.
'     T7  Extreme F ratios must be assembled without intermediate overflow.
'
'   Continuous Family
'     D1  Gamma scale-ratio overflow must saturate to the mathematical limit.
'     D2  Exponential products and inverses must use guarded arithmetic.
'     D3  Weibull moments must remain positive for very large shape.
'     D4  Weibull tiny-shape failures must be classified as #NUM!.
'     D5  Beta must validate both shape parameters under the supported-domain
'         policy.
'     D6  Uniform calculations must support the full finite Double range,
'         including opposite-sign bounds whose width exceeds Double maximum.
'
' ERROR POLICY
'   - A failed assertion increments the failure counter and prints one line.
'   - The test harness itself raises no MsgBox.
'   - Unexpected errors inside production UDFs are expected to return #VALUE!.
'   - Predictable domain, non-convergence and overflow failures are expected to
'     return #NUM! where the public contract states that requirement.
'
' DEPENDENCIES
'   - M_STATS_PROBDIST_CORE
'   - M_STATS_PROBDIST_SPECIALFUNCS
'   - M_STATS_PROBDIST_NORMALFAMILY
'   - M_STATS_PROBDIST_TFAMILY
'   - M_STATS_PROBDIST_CONTINUOUS
'
' PUBLIC SURFACE
'   - Test_STATS_PROBDIST_RunAll
'   - Test_STATS_PROBDIST_RunCore
'   - Test_STATS_PROBDIST_RunNormalFamily
'   - Test_STATS_PROBDIST_RunTFamily
'   - Test_STATS_PROBDIST_RunContinuous
'
' NOTES
'   - Reference values were prepared with high-precision arithmetic and rounded
'     to values suitable for comparison with IEEE-754 Double results.
'   - The production harness does not force artificial iteration limits to test
'     non-convergence. It verifies the public success path and failure
'     classification through reachable numerical cases.
'
' UPDATED
'   2026-07-11 - House-style normalization and consolidated regression coverage.
'==============================================================================

'==============================================================================
' MODULE-LEVEL TEST STATE
'==============================================================================

Private mTestCount          As Long            'Total assertions executed
Private mPassCount          As Long            'Assertions passed
Private mFailCount          As Long            'Assertions failed
Private mFailureLog         As String          'Accumulated failure lines for CI reporting

'==============================================================================
' TEST TOLERANCES
'==============================================================================

Private Const TOL_ABS_TIGHT     As Double = 0.0000000001  '1E-10 absolute
Private Const TOL_ABS_LOOSE     As Double = 0.000001      '1E-6 absolute
Private Const TOL_ABS_LARGE_DF  As Double = 0.000000001   '1E-9 absolute
Private Const TOL_ABS_ULP       As Double = 1E-16          'Constant-level absolute
Private Const TOL_ABS_FEW_ULP   As Double = 5E-16          'Few-ULP absolute

Private Const TOL_REL_TIGHT     As Double = 0.0000000001  '1E-10 relative
Private Const TOL_REL_LOOSE     As Double = 0.000001      '1E-6 relative
Private Const TOL_REL_TAIL      As Double = 0.000000001   '1E-9 relative


'==============================================================================
' PUBLIC ENTRY POINTS
'==============================================================================

Public Sub Test_STATS_PROBDIST_RunAll()
'
'==============================================================================
' Test_STATS_PROBDIST_RunAll
'------------------------------------------------------------------------------
' PURPOSE
'   Runs all four suites in dependency order and prints one consolidated result.
'
' BEHAVIOR
'   - Resets the shared counters.
'   - Runs the selected suite or suites.
'   - Prints one consolidated PASS/FAIL summary.
'
' OUTPUTS
'   - Diagnostic output in the VBA Immediate window.
'
' DEPENDENCIES
'   - BeginRun
'   - Selected suite drivers
'   - EndRun
'
' CALLED FROM
'   - VBA Immediate window
'   - Excel Macros dialog
'   - Other VBA procedures
'
' UPDATED
'   2026-07-11
'==============================================================================
'
    'Initialize the consolidated test run
        BeginRun "ALL SUITES"

    'Run the lower-level numerical infrastructure first
        RunCoreSuite

    'Run the distribution-family suites
        RunNormalFamilySuite
        RunTFamilySuite
        RunContinuousSuite
        RunDiscreteSuite

    'Print the consolidated result
        EndRun
End Sub


Public Sub Test_STATS_PROBDIST_RunCore()
'
'==============================================================================
' Test_STATS_PROBDIST_RunCore
'------------------------------------------------------------------------------
' PURPOSE
'   Runs only the Core and Special Functions suite.
'
' BEHAVIOR
'   - Resets the shared counters.
'   - Runs the selected suite or suites.
'   - Prints one consolidated PASS/FAIL summary.
'
' OUTPUTS
'   - Diagnostic output in the VBA Immediate window.
'
' DEPENDENCIES
'   - BeginRun
'   - Selected suite drivers
'   - EndRun
'
' CALLED FROM
'   - VBA Immediate window
'   - Excel Macros dialog
'   - Other VBA procedures
'
' UPDATED
'   2026-07-11
'==============================================================================
'
    'Initialize the selected test run
        BeginRun "M_STATS_PROBDIST_CORE + M_STATS_PROBDIST_SPECIALFUNCS"

    'Run the selected suite
        RunCoreSuite

    'Print the result
        EndRun
End Sub


Public Sub Test_STATS_PROBDIST_RunNormalFamily()
'
'==============================================================================
' Test_STATS_PROBDIST_RunNormalFamily
'------------------------------------------------------------------------------
' PURPOSE
'   Runs only the Normal and Lognormal family suite.
'
' BEHAVIOR
'   - Resets the shared counters.
'   - Runs the selected suite or suites.
'   - Prints one consolidated PASS/FAIL summary.
'
' OUTPUTS
'   - Diagnostic output in the VBA Immediate window.
'
' DEPENDENCIES
'   - BeginRun
'   - Selected suite drivers
'   - EndRun
'
' CALLED FROM
'   - VBA Immediate window
'   - Excel Macros dialog
'   - Other VBA procedures
'
' UPDATED
'   2026-07-11
'==============================================================================
'
    'Initialize the selected test run
        BeginRun "M_STATS_PROBDIST_NORMALFAMILY"

    'Run the selected suite
        RunNormalFamilySuite

    'Print the result
        EndRun
End Sub


Public Sub Test_STATS_PROBDIST_RunTFamily()
'
'==============================================================================
' Test_STATS_PROBDIST_RunTFamily
'------------------------------------------------------------------------------
' PURPOSE
'   Runs only the Student t, chi-square and F suite.
'
' BEHAVIOR
'   - Resets the shared counters.
'   - Runs the selected suite or suites.
'   - Prints one consolidated PASS/FAIL summary.
'
' OUTPUTS
'   - Diagnostic output in the VBA Immediate window.
'
' DEPENDENCIES
'   - BeginRun
'   - Selected suite drivers
'   - EndRun
'
' CALLED FROM
'   - VBA Immediate window
'   - Excel Macros dialog
'   - Other VBA procedures
'
' UPDATED
'   2026-07-11
'==============================================================================
'
    'Initialize the selected test run
        BeginRun "M_STATS_PROBDIST_TFAMILY"

    'Run the selected suite
        RunTFamilySuite

    'Print the result
        EndRun
End Sub


Public Sub Test_STATS_PROBDIST_RunContinuous()
'
'==============================================================================
' Test_STATS_PROBDIST_RunContinuous
'------------------------------------------------------------------------------
' PURPOSE
'   Runs only the Gamma, Beta, Exponential, Weibull and Uniform suite.
'
' BEHAVIOR
'   - Resets the shared counters.
'   - Runs the selected suite or suites.
'   - Prints one consolidated PASS/FAIL summary.
'
' OUTPUTS
'   - Diagnostic output in the VBA Immediate window.
'
' DEPENDENCIES
'   - BeginRun
'   - Selected suite drivers
'   - EndRun
'
' CALLED FROM
'   - VBA Immediate window
'   - Excel Macros dialog
'   - Other VBA procedures
'
' UPDATED
'   2026-07-11
'==============================================================================
'
    'Initialize the selected test run
        BeginRun "M_STATS_PROBDIST_CONTINUOUS"

    'Run the selected suite
        RunContinuousSuite

    'Print the result
        EndRun
End Sub


Public Sub Test_STATS_PROBDIST_RunDiscrete()
'
'==============================================================================
' Test_STATS_PROBDIST_RunDiscrete
'------------------------------------------------------------------------------
' PURPOSE
'   Runs only the Binomial, Poisson and Geometric suite.
'
' BEHAVIOR
'   - Resets the shared counters.
'   - Runs the selected suite or suites.
'   - Prints one consolidated PASS/FAIL summary.
'
' OUTPUTS
'   - Diagnostic output in the VBA Immediate window.
'
' DEPENDENCIES
'   - BeginRun
'   - Selected suite drivers
'   - EndRun
'
' CALLED FROM
'   - VBA Immediate window
'   - Excel Macros dialog
'   - Other VBA procedures
'
' UPDATED
'   2026-07-19
'==============================================================================
'
    'Initialize the selected test run
        BeginRun "M_STATS_PROBDIST_DISCRETE"

    'Run the selected suite
        RunDiscreteSuite

    'Print the result
        EndRun
End Sub


Private Sub RunCoreSuite()
'
'==============================================================================
' RunCoreSuite
'------------------------------------------------------------------------------
' PURPOSE
'   Runs the Core and Special Functions test sections.
'
' BEHAVIOR
'   - Prints one section heading.
'   - Executes silent passing assertions.
'   - Records detailed output only when an assertion fails.
'
' DEPENDENCIES
'   - Test section procedures in this module
'
' CALLED FROM
'   - Public test entry points
'
' UPDATED
'   2026-07-11
'==============================================================================
'
    Debug.Print "== SUITE: M_STATS_PROBDIST_CORE + M_STATS_PROBDIST_SPECIALFUNCS"
    Test_Core_Constants
    Test_Core_Log1p
    Test_Core_TryExp
    Test_Core_AffineAndStandardize
    Test_Core_LogGamma
    Test_Core_SpecialFunctionKernels
    Test_Core_NormalInvRaw
End Sub


Private Sub RunNormalFamilySuite()
'
'==============================================================================
' RunNormalFamilySuite
'------------------------------------------------------------------------------
' PURPOSE
'   Runs the Normal and Lognormal family test sections.
'
' BEHAVIOR
'   - Prints one section heading.
'   - Executes silent passing assertions.
'   - Records detailed output only when an assertion fails.
'
' DEPENDENCIES
'   - Test section procedures in this module
'
' CALLED FROM
'   - Public test entry points
'
' UPDATED
'   2026-07-11
'==============================================================================
'
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
    Test_NF_Survival
    Test_NF_MagnitudePolicy
    Test_NF_InverseSurvival
    Test_NF_FastInverse
    Test_NF_LognormalCore
    Test_NF_LognormalMoments
    Test_NF_LognormalUnderflow
    Test_NF_ParameterRoundTrip
    Test_NF_ErrorContract
    Test_NF_OverflowContract
End Sub


Private Sub RunTFamilySuite()
'
'==============================================================================
' RunTFamilySuite
'------------------------------------------------------------------------------
' PURPOSE
'   Runs the Student t, chi-square and F family test sections.
'
' BEHAVIOR
'   - Prints one section heading.
'   - Executes silent passing assertions.
'   - Records detailed output only when an assertion fails.
'
' DEPENDENCIES
'   - Test section procedures in this module
'
' CALLED FROM
'   - Public test entry points
'
' UPDATED
'   2026-07-11
'==============================================================================
'
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
'
'==============================================================================
' RunContinuousSuite
'------------------------------------------------------------------------------
' PURPOSE
'   Runs the Gamma, Beta, Exponential, Weibull and Uniform test sections.
'
' BEHAVIOR
'   - Prints one section heading.
'   - Executes silent passing assertions.
'   - Records detailed output only when an assertion fails.
'
' DEPENDENCIES
'   - Test section procedures in this module
'
' CALLED FROM
'   - Public test entry points
'
' UPDATED
'   2026-07-11
'==============================================================================
'
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


Private Sub RunDiscreteSuite()
'
'==============================================================================
' RunDiscreteSuite
'------------------------------------------------------------------------------
' PURPOSE
'   Runs the Binomial, Poisson and Geometric test sections.
'
' BEHAVIOR
'   - Prints one section heading.
'   - Executes silent passing assertions.
'   - Records detailed output only when an assertion fails.
'
' DEPENDENCIES
'   - Test section procedures in this module
'
' CALLED FROM
'   - Public test entry points
'
' UPDATED
'   2026-07-19
'==============================================================================
'
    Debug.Print "== SUITE: M_STATS_PROBDIST_DISCRETE"
    Test_DS_BinomialPMF
    Test_DS_BinomialCumulative
    Test_DS_BinomialSurvival
    Test_DS_BinomialInverse
    Test_DS_BinomialMoments
    Test_DS_PoissonPMF
    Test_DS_PoissonCumulative
    Test_DS_PoissonSurvival
    Test_DS_PoissonInverse
    Test_DS_PoissonMoments
    Test_DS_GeometricPMF
    Test_DS_GeometricCumulative
    Test_DS_GeometricSurvival
    Test_DS_GeometricInverse
    Test_DS_GeometricMoments
    Test_DS_ErrorContract
    Test_DS_SupportEdges
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
'   Resets the shared counters and prints the test-run header.
'
' INPUTS
'   Title   Human-readable run or suite title.
'
' BEHAVIOR
'   - Resets total, pass and failure counters to zero.
'   - Prints a timestamped heading in the Immediate window.
'
' DEPENDENCIES
'   - Module-level test counters
'
' CALLED FROM
'   - Public test entry points
'
' UPDATED
'   2026-07-11
'==============================================================================
'
'------------------------------------------------------------------------------
' RESET STATE
'------------------------------------------------------------------------------
    'Reset all assertion counters
        mTestCount = 0
        mPassCount = 0
        mFailCount = 0
        mFailureLog = vbNullString

'------------------------------------------------------------------------------
' PRINT HEADER
'------------------------------------------------------------------------------
    'Print the run separator and timestamp
        Debug.Print String(70, "=")
        Debug.Print _
            Title & _
            " - test run " & _
            Format$(Now, "yyyy-mm-dd hh:nn:ss")
        Debug.Print String(70, "=")
End Sub


Private Sub EndRun()
'
'==============================================================================
' EndRun
'------------------------------------------------------------------------------
' PURPOSE
'   Prints the consolidated assertion counts and final test verdict.
'
' BEHAVIOR
'   - Prints total, passed and failed assertion counts.
'   - Prints a green verdict when the failure count is zero.
'   - Prints the number of failed assertions otherwise.
'
' DEPENDENCIES
'   - Module-level test counters
'
' CALLED FROM
'   - Public test entry points
'
' UPDATED
'   2026-07-11
'==============================================================================
'
'------------------------------------------------------------------------------
' PRINT COUNTS
'------------------------------------------------------------------------------
    'Print the summary separator
        Debug.Print String(70, "-")

    'Print the assertion counters
        Debug.Print _
            "TOTAL  " & mTestCount & _
            "   PASS " & mPassCount & _
            "   FAIL " & mFailCount

'------------------------------------------------------------------------------
' PRINT VERDICT
'------------------------------------------------------------------------------
    'Print the final green or red result
        If mFailCount = 0 Then
            Debug.Print "RESULT: ALL TESTS PASSED"
        Else
            Debug.Print _
                "RESULT: " & _
                mFailCount & _
                " TEST(S) FAILED"
        End If

    'Close the run output
        Debug.Print String(70, "=")
End Sub


Private Sub Test_Core_Constants()
'
'==============================================================================
' Test_Core_Constants
'------------------------------------------------------------------------------
' PURPOSE
'   Verifies that shared mathematical constants equal independently evaluated
'   VBA Double references at the exact binary level.
'
' WHY
'   The VBA editor canonicalizes long decimal literals. Exact equality catches
'   a source literal that displays plausibly but maps to the wrong Double.
'
' DEPENDENCIES
'   - PROB_PI
'   - PROB_TWO_PI
'   - PROB_HALF_LOG_TWO_PI
'   - PROB_HALF_LOG_PI
'   - AssertExactlyEqual
'
' CALLED FROM
'   - RunCoreSuite
'
' UPDATED
'   2026-07-11
'==============================================================================
'
'------------------------------------------------------------------------------
' EXECUTE ASSERTIONS
'------------------------------------------------------------------------------
    'Print the test-section heading
        Debug.Print "-- Core constants"

    'Verify pi against four times arctangent of one
        AssertExactlyEqual _
            "PROB_PI exact", _
            PROB_PI, _
            4# * Atn(1#)

    'Verify two-pi against eight times arctangent of one
        AssertExactlyEqual _
            "PROB_TWO_PI exact", _
            PROB_TWO_PI, _
            8# * Atn(1#)

    'Verify one-half log of two-pi against an independent runtime expression
        AssertExactlyEqual _
            "PROB_HALF_LOG_TWO_PI exact", _
            PROB_HALF_LOG_TWO_PI, _
            0.5 * Log(8# * Atn(1#))

    'Verify one-half log of pi against an independent runtime expression
        AssertExactlyEqual _
            "PROB_HALF_LOG_PI exact", _
            PROB_HALF_LOG_PI, _
            0.5 * Log(4# * Atn(1#))
End Sub


Private Sub Test_Core_Log1p()
'
'==============================================================================
' Test_Core_Log1p
'------------------------------------------------------------------------------
' PURPOSE
'   Verifies the cancellation-resistant PROB_Log1p implementation.
'
' BEHAVIOR
'   - Prints one section heading.
'   - Executes silent passing assertions.
'   - Records detailed output only when an assertion fails.
'
' DEPENDENCIES
'   - Production functions under test
'   - Shared assertion helpers in this module
'
' CALLED FROM
'   - RunCoreSuite
'
' UPDATED
'   2026-07-11
'==============================================================================
'
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
'
'==============================================================================
' Test_Core_TryExp
'------------------------------------------------------------------------------
' PURPOSE
'   Verifies exponential and guarded-arithmetic Try contracts.
'
' BEHAVIOR
'   - Prints one section heading.
'   - Executes silent passing assertions.
'   - Records detailed output only when an assertion fails.
'
' DEPENDENCIES
'   - Production functions under test
'   - Shared assertion helpers in this module
'
' CALLED FROM
'   - RunCoreSuite
'
' UPDATED
'   2026-07-11
'==============================================================================
'
    Dim ExpResult As Double
    Dim ArithmeticResult As Double

    Debug.Print "-- Core exponential and arithmetic Try contracts"

    AssertTrue "TryExp 709 accepted", PROB_TryExp(709#, ExpResult)
    AssertRelClose "TryExp 709 value", ExpResult, 8.21840746155497E+307, 0.000000000001

    '709.5 is finite and was incorrectly rejected by the former 709 cutoff.
    AssertTrue "TryExp 709.5 accepted", PROB_TryExp(709.5, ExpResult)
    AssertTrue "TryExp 709.5 positive finite", (ExpResult > 0# And PROB_IsFinite(ExpResult))
    AssertTrue "TryExp 709.78 accepted", PROB_TryExp(709.78, ExpResult)
    AssertTrue "TryExp 709.79 overflow rejected", (Not PROB_TryExp(709.79, ExpResult))
    AssertTrue "TryExp 710 overflow rejected", (Not PROB_TryExp(710#, ExpResult))

    AssertTrue "true finiteness accepts 1E200", PROB_IsFinite(1E+200)
    AssertTrue "supported magnitude rejects 1E200", _
        (Not PROB_IsWithinSupportedMagnitude(1E+200))

    AssertTrue "TryAdd finite", PROB_TryAdd(1E+200, -1E+200, ExpResult)
    AssertClose "TryAdd cancellation", ExpResult, 0#, 0#
    AssertTrue "TryAdd overflow rejected", _
        (Not PROB_TryAdd(PROB_DOUBLE_MAX, PROB_DOUBLE_MAX, ExpResult))

    AssertTrue "TryExp underflow accepted", PROB_TryExp(-1000#, ExpResult)
    AssertClose "TryExp underflow value", ExpResult, 0#, 0#

    AssertTrue "TryExp regular accepted", PROB_TryExp(1#, ExpResult)
    AssertClose "TryExp regular value", ExpResult, 2.71828182845905, TOL_ABS_TIGHT

    AssertTrue "TryMultiply regular", PROB_TryMultiply(1E+150, 1E+150, ArithmeticResult)
    AssertRelClose "TryMultiply value", ArithmeticResult, 1E+300, TOL_REL_TIGHT
    AssertTrue "TryMultiply overflow rejected", _
        (Not PROB_TryMultiply(1E+200, 1E+200, ArithmeticResult))

    AssertTrue "TryDivide regular", PROB_TryDivide(1#, 4#, ArithmeticResult)
    AssertClose "TryDivide value", ArithmeticResult, 0.25, 0#
    AssertTrue "TryDivide overflow rejected", _
        (Not PROB_TryDivide(1E+308, 1E-308, ArithmeticResult))
End Sub


Private Sub Test_Core_AffineAndStandardize()
'
'==============================================================================
' Test_Core_AffineAndStandardize
'------------------------------------------------------------------------------
' PURPOSE
'   Verifies the guarded location-scale reconstruction PROB_TryAffineTransform
'   and the guarded standardization PROB_TryStandardize: correct values for
'   ordinary inputs, FALSE on overflow, and FALSE on a zero scale.
'
' BEHAVIOR
'   - Prints one section heading.
'   - Executes silent passing assertions.
'   - Records detailed output only when an assertion fails.
'
' DEPENDENCIES
'   - Production functions under test
'   - Shared assertion helpers in this module
'
' CALLED FROM
'   - RunCoreSuite
'
' UPDATED
'   2026-07-17
'==============================================================================
'
    Dim R                   As Double          'Guarded result

    Debug.Print "-- Core affine reconstruction and standardization Try contracts"

    'Affine: Offset + Scale * Value
    AssertTrue "affine ordinary accepted", _
        PROB_TryAffineTransform(100#, 15#, 2.326, R)
    AssertClose "affine value 100 + 15*2.326", R, 134.89, TOL_ABS_TIGHT
    AssertTrue "affine negative scale", _
        PROB_TryAffineTransform(0#, -2#, 3#, R)
    AssertClose "affine -2*3", R, -6#, TOL_ABS_TIGHT

    'Affine overflow of the product is rejected
    AssertTrue "affine product overflow rejected", _
        (Not PROB_TryAffineTransform(0#, 1E+308, 8#, R))
    'Affine overflow of the sum is rejected
    AssertTrue "affine sum overflow rejected", _
        (Not PROB_TryAffineTransform(1E+308, 1#, 1E+308, R))

    'Standardize: (Value - Location) / ScaleParam
    AssertTrue "standardize ordinary accepted", _
        PROB_TryStandardize(10#, 4#, 2#, R)
    AssertClose "standardize (10-4)/2", R, 3#, TOL_ABS_TIGHT
    AssertTrue "standardize negative result", _
        PROB_TryStandardize(1#, 1.96, 1#, R)
    AssertClose "standardize (1-1.96)/1", R, -0.96, TOL_ABS_TIGHT

    'Zero scale is rejected (division guard)
    AssertTrue "standardize zero scale rejected", _
        (Not PROB_TryStandardize(1#, 0#, 0#, R))

    'The closed gap: a tiny scale that would overflow returns FALSE, not a fault
    AssertTrue "standardize tiny scale overflow rejected", _
        (Not PROB_TryStandardize(9E+99, 0#, 1E-300, R))

    'Underflow of the quotient to zero is a valid success
    AssertTrue "standardize underflow accepted", _
        PROB_TryStandardize(1#, 0#, 1E+300, R)
    AssertClose "standardize underflow value", R, 0#, TOL_ABS_TIGHT
End Sub


Private Sub Test_Core_LogGamma()
'
'==============================================================================
' Test_Core_LogGamma
'------------------------------------------------------------------------------
' PURPOSE
'   Verifies log-gamma, half-difference and log-beta calculations.
'
' BEHAVIOR
'   - Prints one section heading.
'   - Executes silent passing assertions.
'   - Records detailed output only when an assertion fails.
'
' DEPENDENCIES
'   - Production functions under test
'   - Shared assertion helpers in this module
'
' CALLED FROM
'   - RunCoreSuite
'
' UPDATED
'   2026-07-11
'==============================================================================
'
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

    'Extremely unbalanced arguments must not lose LogGamma(Small) when the
    'two enormous large-argument log-gammas cancel.
    AssertClose "LogBeta extreme unbalanced", _
        PROB_LogBeta(5E+98, 5E-100), _
        228.64907138697, _
        0.000000000001

    AssertClose "LogBeta extreme unbalanced symmetric", _
        PROB_LogBeta(5E-100, 5E+98), _
        228.64907138697, _
        0.000000000001

    'PROB_LogGammaDelta structural identities (validate the kernel independently
    'of LogBeta): Delta(z,0)=0, Delta(z,1)=Log(z), and the composition law
    'Delta(z,s+t)=Delta(z,s)+Delta(z+s,t).
    AssertClose "LogGammaDelta(z,0)=0 (z=2.5)", _
        PROB_LogGammaDelta(2.5, 0#), 0#, TOL_ABS_TIGHT
    AssertClose "LogGammaDelta(z,0)=0 (z=1E6)", _
        PROB_LogGammaDelta(1000000#, 0#), 0#, TOL_ABS_TIGHT
    AssertRelClose "LogGammaDelta(z,1)=Log(z) (z=2.5)", _
        PROB_LogGammaDelta(2.5, 1#), Log(2.5), TOL_REL_TIGHT
    AssertRelClose "LogGammaDelta(z,1)=Log(z) (z=1E6)", _
        PROB_LogGammaDelta(1000000#, 1#), Log(1000000#), TOL_REL_TIGHT
    AssertRelClose "LogGammaDelta composition (z=2.5)", _
        PROB_LogGammaDelta(2.5, 0.7), _
        PROB_LogGammaDelta(2.5, 0.3) + PROB_LogGammaDelta(2.8, 0.4), _
        TOL_REL_TIGHT
    AssertRelClose "LogGammaDelta composition (z=1E4)", _
        PROB_LogGammaDelta(10000#, 5.75), _
        PROB_LogGammaDelta(10000#, 0.7) + PROB_LogGammaDelta(10000.7, 5.05), _
        TOL_REL_TIGHT

    'LogBeta unbalanced regression: the middle band (ratio ~1E-2 to 1E-13) that
    'the previous one-term asymptotic could not reach is now handled by the stable
    'log-gamma difference. References are 50-digit mpmath values.
    AssertRelClose "LogBeta unbalanced (8E5, 0.8)", _
        PROB_LogBeta(800000#, 0.8), -10.7218338269202, TOL_REL_TIGHT
    AssertRelClose "LogBeta unbalanced (1E10, 0.7)", _
        PROB_LogBeta(10000000000#, 0.7), -15.8572284044162, TOL_REL_TIGHT
    AssertRelClose "LogBeta unbalanced (1E6, 2.5)", _
        PROB_LogBeta(1000000#, 2.5), -34.2540953994365, TOL_REL_TIGHT
End Sub


Private Sub Test_Core_SpecialFunctionKernels()
'
'==============================================================================
' Test_Core_SpecialFunctionKernels
'------------------------------------------------------------------------------
' PURPOSE
'   Verifies the reusable incomplete-gamma and incomplete-beta kernels directly.
'
' WHY
'   Distribution-level tests can hide a marshalling defect behind another
'   wrapper. Direct kernel tests isolate lower-layer failures before the public
'   distribution suites run.
'
' BEHAVIOR
'   - Checks known lower and upper incomplete-gamma values.
'   - Checks the P + Q complement identity.
'   - Checks incomplete-beta and swapped-complement values.
'   - Checks direct gamma and beta inverse round-trips.
'
' DEPENDENCIES
'   - PROB_TryGammaRegularizedP
'   - PROB_TryGammaRegularizedQ
'   - PROB_TryGammaInvP
'   - PROB_TryBetaRegularized
'   - PROB_TryBetaInvRegularized
'   - Shared assertion helpers
'
' CALLED FROM
'   - RunCoreSuite
'
' UPDATED
'   2026-07-11
'==============================================================================
'
'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Dim GammaP             As Double          'Regularized lower gamma value
    Dim GammaQ             As Double          'Regularized upper gamma value
    Dim GammaRoot          As Double          'Inverse-gamma root
    Dim BetaValue          As Double          'Regularized beta value
    Dim BetaComplement     As Double          'Swapped beta complement
    Dim BetaRoot           As Double          'Inverse-beta root
    Dim BetaRootComplement As Double          'Complement of inverse-beta root
    Dim Recovered          As Double          'Recovered probability
    Dim FailMsg            As String          'Kernel diagnostic message

'------------------------------------------------------------------------------
' EXECUTE GAMMA ASSERTIONS
'------------------------------------------------------------------------------
    'Print the test-section heading
        Debug.Print "-- Core special-function kernels"

    'Evaluate the regularized lower incomplete gamma
        FailMsg = vbNullString

        AssertTrue _
            "gamma P kernel succeeds", _
            PROB_TryGammaRegularizedP( _
                2.5, _
                3#, _
                GammaP, _
                FailMsg)

    'Check the known lower-tail value
        AssertRelClose _
            "gamma P(2.5,3)", _
            GammaP, _
            0.693781081586722, _
            TOL_REL_TIGHT

    'Evaluate the regularized upper incomplete gamma
        FailMsg = vbNullString

        AssertTrue _
            "gamma Q kernel succeeds", _
            PROB_TryGammaRegularizedQ( _
                2.5, _
                3#, _
                GammaQ, _
                FailMsg)

    'Check the known upper-tail value
        AssertRelClose _
            "gamma Q(2.5,3)", _
            GammaQ, _
            0.306218918413278, _
            TOL_REL_TIGHT

    'Check the lower-plus-upper complement identity
        AssertClose _
            "gamma P + Q = 1", _
            GammaP + GammaQ, _
            1#, _
            TOL_ABS_TIGHT

    'Invert the lower incomplete gamma at probability 0.7
        FailMsg = vbNullString

        AssertTrue _
            "gamma inverse kernel succeeds", _
            PROB_TryGammaInvP( _
                0.7, _
                0.3, _
                2.5, _
                GammaRoot, _
                FailMsg)

    'Re-evaluate the recovered gamma probability
        FailMsg = vbNullString

        AssertTrue _
            "gamma inverse recovery succeeds", _
            PROB_TryGammaRegularizedP( _
                2.5, _
                GammaRoot, _
                Recovered, _
                FailMsg)

    'Check the inverse round-trip
        AssertRelClose _
            "gamma inverse kernel round-trip", _
            Recovered, _
            0.7, _
            TOL_REL_TIGHT

'------------------------------------------------------------------------------
' EXECUTE BETA ASSERTIONS
'------------------------------------------------------------------------------
    'Evaluate I_x(2,5) at x = 0.3
        FailMsg = vbNullString

        AssertTrue _
            "beta kernel succeeds", _
            PROB_TryBetaRegularized( _
                0.3, _
                0.7, _
                2#, _
                5#, _
                BetaValue, _
                FailMsg)

    'Check the known regularized-beta value
        AssertRelClose _
            "beta I(0.3;2,5)", _
            BetaValue, _
            0.579825, _
            TOL_REL_TIGHT

    'Evaluate the swapped complement directly
        FailMsg = vbNullString

        AssertTrue _
            "beta complement kernel succeeds", _
            PROB_TryBetaRegularized( _
                0.7, _
                0.3, _
                5#, _
                2#, _
                BetaComplement, _
                FailMsg)

    'Check the direct complement identity
        AssertClose _
            "beta value + complement = 1", _
            BetaValue + BetaComplement, _
            1#, _
            TOL_ABS_TIGHT

    'Invert the regularized beta at probability 0.6
        FailMsg = vbNullString

        AssertTrue _
            "beta inverse kernel succeeds", _
            PROB_TryBetaInvRegularized( _
                0.6, _
                0.4, _
                2#, _
                5#, _
                BetaRoot, _
                BetaRootComplement, _
                FailMsg)

    'Check that the returned root pair remains complementary
        AssertClose _
            "beta inverse root pair sums to one", _
            BetaRoot + BetaRootComplement, _
            1#, _
            TOL_ABS_TIGHT

    'Re-evaluate the recovered beta probability
        FailMsg = vbNullString

        AssertTrue _
            "beta inverse recovery succeeds", _
            PROB_TryBetaRegularized( _
                BetaRoot, _
                BetaRootComplement, _
                2#, _
                5#, _
                Recovered, _
                FailMsg)

    'Check the inverse round-trip
        AssertRelClose _
            "beta inverse kernel round-trip", _
            Recovered, _
            0.6, _
            TOL_REL_TIGHT
End Sub


Private Sub Test_Core_NormalInvRaw()
'
'==============================================================================
' Test_Core_NormalInvRaw
'------------------------------------------------------------------------------
' PURPOSE
'   Verifies the shared raw inverse-normal seed kernel.
'
' BEHAVIOR
'   - Prints one section heading.
'   - Executes silent passing assertions.
'   - Records detailed output only when an assertion fails.
'
' DEPENDENCIES
'   - Production functions under test
'   - Shared assertion helpers in this module
'
' CALLED FROM
'   - RunCoreSuite
'
' UPDATED
'   2026-07-11
'==============================================================================
'
    Debug.Print "-- Core PROB_NormalInvCDFRaw (shared seed kernel)"

    'REGRESSION N1: the raw kernel must actually return its value. When the "="
    'was lost from the assignment it returned 0# on every branch, and every
    'inverse-normal caller in the project silently produced garbage.
    AssertTrue "raw kernel is not a zero stub", (Abs(PROB_NormalInvCDFRaw(0.975)) > 1#)

    'Raw Acklam accuracy is ~1.15E-9, so a loose tolerance is correct here
    AssertClose "raw InvPhi(0.975)", PROB_NormalInvCDFRaw(0.975), 1.95996398454005, TOL_ABS_LOOSE
    AssertClose "raw InvPhi(0.025)", PROB_NormalInvCDFRaw(0.025), -1.95996398454005, TOL_ABS_LOOSE
    AssertClose "raw InvPhi(0.5)", PROB_NormalInvCDFRaw(0.5), 0#, TOL_ABS_LOOSE

    'Each of the three branches must be exercised
    AssertTrue "raw lower branch", (PROB_NormalInvCDFRaw(0.001) < -3#)
    AssertTrue "raw upper branch", (PROB_NormalInvCDFRaw(0.999) > 3#)
End Sub


'==============================================================================
' SUITE - NORMAL FAMILY
'==============================================================================

Private Sub Test_NF_StandardNormalDensity()
'
'==============================================================================
' Test_NF_StandardNormalDensity
'------------------------------------------------------------------------------
' PURPOSE
'   Verifies standard-normal density values, symmetry and tail underflow.
'
' BEHAVIOR
'   - Prints one section heading.
'   - Executes silent passing assertions.
'   - Records detailed output only when an assertion fails.
'
' DEPENDENCIES
'   - Production functions under test
'   - Shared assertion helpers in this module
'
' CALLED FROM
'   - RunNormalFamilySuite
'
' UPDATED
'   2026-07-11
'==============================================================================
'
    Debug.Print "-- Standard normal density"

    'REGRESSION N4: calculate the oracle independently at run time and require
    'the exact binary Double value.
        AssertExactlyEqual _
            "phi(0) correctly rounded", _
            K_STATS_NormalStandard_Density(0#), _
            1# / Sqr(8# * Atn(1#))

    AssertClose "phi(1)", K_STATS_NormalStandard_Density(1#), 0.241970724519143, TOL_ABS_TIGHT
    AssertClose "phi(-1)", K_STATS_NormalStandard_Density(-1#), 0.241970724519143, TOL_ABS_TIGHT
    AssertClose "phi(2)", K_STATS_NormalStandard_Density(2#), 0.053990966513188, TOL_ABS_TIGHT

    'Symmetry is exact, not approximate
    AssertClose "phi symmetric", _
        CDbl(K_STATS_NormalStandard_Density(-1.7)) - _
        CDbl(K_STATS_NormalStandard_Density(1.7)), 0#, 0#

    'Far tail underflows to a valid zero
    AssertClose "phi far tail = 0", K_STATS_NormalStandard_Density(50#), 0#, 0#
End Sub


Private Sub Test_NF_StandardNormalCumulative()
'
'==============================================================================
' Test_NF_StandardNormalCumulative
'------------------------------------------------------------------------------
' PURPOSE
'   Verifies standard-normal cumulative probabilities and deep tails.
'
' BEHAVIOR
'   - Prints one section heading.
'   - Executes silent passing assertions.
'   - Records detailed output only when an assertion fails.
'
' DEPENDENCIES
'   - Production functions under test
'   - Shared assertion helpers in this module
'
' CALLED FROM
'   - RunNormalFamilySuite
'
' UPDATED
'   2026-07-11
'==============================================================================
'
    Debug.Print "-- Standard normal cumulative"
    AssertClose "Phi(0)", K_STATS_NormalStandard_Cumulative(0#), 0.5, TOL_ABS_TIGHT
    AssertClose "Phi(1)", K_STATS_NormalStandard_Cumulative(1#), 0.841344746068543, TOL_ABS_TIGHT
    AssertClose "Phi(2)", K_STATS_NormalStandard_Cumulative(2#), 0.977249868051821, TOL_ABS_TIGHT
    AssertClose "Phi(-1)", K_STATS_NormalStandard_Cumulative(-1#), 0.158655253931457, TOL_ABS_TIGHT
    AssertClose "Phi(1.959963984540054)", _
        K_STATS_NormalStandard_Cumulative(1.95996398454005), 0.975, TOL_ABS_TIGHT

    'Both sides of the rational / continued-fraction split at Z = Sqr(50)
    AssertClose "Phi split seam", _
        CDbl(K_STATS_NormalStandard_Cumulative(-7.07106781)) - _
        CDbl(K_STATS_NormalStandard_Cumulative(-7.07106782)), 0#, 1E-16

    'Representable deep tails must not be cut off prematurely.
    AssertRelClose "Phi(-37.5) representable tail", _
        K_STATS_NormalStandard_Cumulative(-37.5), _
        4.60535300958195E-308, TOL_REL_TAIL

    AssertRelClose "Phi(-38) subnormal tail", _
        K_STATS_NormalStandard_Cumulative(-38#), _
        2.88542835100396E-316, TOL_REL_LOOSE

    'Only genuinely unrepresentable tails saturate.
    AssertClose "Phi(-40) saturates", K_STATS_NormalStandard_Cumulative(-40#), 0#, 0#
    AssertClose "Phi(40) saturates", K_STATS_NormalStandard_Cumulative(40#), 1#, 0#
End Sub


Private Sub Test_NF_StandardNormalInverse()
'
'==============================================================================
' Test_NF_StandardNormalInverse
'------------------------------------------------------------------------------
' PURPOSE
'   Verifies standard-normal inverse cumulative values.
'
' BEHAVIOR
'   - Prints one section heading.
'   - Executes silent passing assertions.
'   - Records detailed output only when an assertion fails.
'
' DEPENDENCIES
'   - Production functions under test
'   - Shared assertion helpers in this module
'
' CALLED FROM
'   - RunNormalFamilySuite
'
' UPDATED
'   2026-07-11
'==============================================================================
'
    Debug.Print "-- Standard normal inverse"

    'REGRESSION N1: the zero-stub Acklam kernel returned 1.1906 here
    AssertClose "InvPhi(0.975)", _
        K_STATS_NormalStandard_InverseCumulative(0.975), 1.95996398454005, TOL_ABS_TIGHT

    AssertClose "InvPhi(0.5)", K_STATS_NormalStandard_InverseCumulative(0.5), 0#, TOL_ABS_TIGHT
    AssertClose "InvPhi(0.95)", _
        K_STATS_NormalStandard_InverseCumulative(0.95), 1.64485362695147, TOL_ABS_TIGHT
    AssertClose "InvPhi(0.025)", _
        K_STATS_NormalStandard_InverseCumulative(0.025), -1.95996398454005, TOL_ABS_TIGHT
    AssertClose "InvPhi(0.005)", _
        K_STATS_NormalStandard_InverseCumulative(0.005), -2.5758293035489, TOL_ABS_TIGHT
End Sub


Private Sub Test_NF_InverseTails()
'
'==============================================================================
' Test_NF_InverseTails
'------------------------------------------------------------------------------
' PURPOSE
'   Verifies inverse-normal accuracy in deep representable tails.
'
' BEHAVIOR
'   - Prints one section heading.
'   - Executes silent passing assertions.
'   - Records detailed output only when an assertion fails.
'
' DEPENDENCIES
'   - Production functions under test
'   - Shared assertion helpers in this module
'
' CALLED FROM
'   - RunNormalFamilySuite
'
' UPDATED
'   2026-07-11
'==============================================================================
'
    Debug.Print "-- Standard normal inverse, deep tails"

    'REGRESSION N2: the tail kernel must preserve the representable CDF at the
    'Acklam seed and the Halley guard must still reject genuinely rounded
    'endpoints. The deep-tail quantile must remain accurate.
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
'
'==============================================================================
' Test_NF_InverseRoundTrips
'------------------------------------------------------------------------------
' PURPOSE
'   Verifies standard-normal CDF and inverse round-trips.
'
' BEHAVIOR
'   - Prints one section heading.
'   - Executes silent passing assertions.
'   - Records detailed output only when an assertion fails.
'
' DEPENDENCIES
'   - Production functions under test
'   - Shared assertion helpers in this module
'
' CALLED FROM
'   - RunNormalFamilySuite
'
' UPDATED
'   2026-07-11
'==============================================================================
'
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
            AssertClose "roundtrip z=" & Z, Back, Z, TOL_ABS_TIGHT
        End If
    Next I
End Sub


Private Sub Test_NF_Symmetry()
'
'==============================================================================
' Test_NF_Symmetry
'------------------------------------------------------------------------------
' PURPOSE
'   Verifies normal CDF symmetry and inverse antisymmetry.
'
' BEHAVIOR
'   - Prints one section heading.
'   - Executes silent passing assertions.
'   - Records detailed output only when an assertion fails.
'
' DEPENDENCIES
'   - Production functions under test
'   - Shared assertion helpers in this module
'
' CALLED FROM
'   - RunNormalFamilySuite
'
' UPDATED
'   2026-07-11
'==============================================================================
'
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
            AssertClose "symmetry z=" & Z, CDbl(Lo) + CDbl(Hi), 1#, TOL_ABS_TIGHT
        End If
    Next I

    'Inverse antisymmetry
    AssertClose "InvPhi(p) = -InvPhi(1-p)", _
        CDbl(K_STATS_NormalStandard_InverseCumulative(0.02)) + _
        CDbl(K_STATS_NormalStandard_InverseCumulative(0.98)), 0#, TOL_ABS_TIGHT
End Sub


Private Sub Test_NF_GeneralNormal()
'
'==============================================================================
' Test_NF_GeneralNormal
'------------------------------------------------------------------------------
' PURPOSE
'   Verifies general-normal density, CDF and inverse parameterization.
'
' BEHAVIOR
'   - Prints one section heading.
'   - Executes silent passing assertions.
'   - Records detailed output only when an assertion fails.
'
' DEPENDENCIES
'   - Production functions under test
'   - Shared assertion helpers in this module
'
' CALLED FROM
'   - RunNormalFamilySuite
'
' UPDATED
'   2026-07-11
'==============================================================================
'
    Debug.Print "-- General normal (mean=10, sd=2)"
    'Density at the mean equals phi(0)/sd
    AssertClose "density@mean", _
        K_STATS_Normal_Density(10#, 10#, 2#), 0.398942280401433 / 2#, TOL_ABS_TIGHT
    'CDF at the mean is 0.5
    AssertClose "cdf@mean", K_STATS_Normal_Cumulative(10#, 10#, 2#), 0.5, TOL_ABS_TIGHT
    'CDF one sd above the mean equals Phi(1)
    AssertClose "cdf@mean+sd", _
        K_STATS_Normal_Cumulative(12#, 10#, 2#), 0.841344746068543, TOL_ABS_TIGHT
    'Inverse at 0.975 is mean + 1.959963984540054 * sd
    AssertClose "inv@0.975", _
        K_STATS_Normal_InverseCumulative(0.975, 10#, 2#), _
        10# + 1.95996398454005 * 2#, TOL_ABS_TIGHT
End Sub


Private Sub Test_NF_ZScore()
'
'==============================================================================
' Test_NF_ZScore
'------------------------------------------------------------------------------
' PURPOSE
'   Verifies the general-normal Z-score transformation.
'
' BEHAVIOR
'   - Prints one section heading.
'   - Executes silent passing assertions.
'   - Records detailed output only when an assertion fails.
'
' DEPENDENCIES
'   - Production functions under test
'   - Shared assertion helpers in this module
'
' CALLED FROM
'   - RunNormalFamilySuite
'
' UPDATED
'   2026-07-11
'==============================================================================
'
    Debug.Print "-- Z-score"
    AssertClose "z(10,4,2)", K_STATS_Normal_ZScore(10#, 4#, 2#), 3#, TOL_ABS_TIGHT
    AssertClose "z(4,4,2)", K_STATS_Normal_ZScore(4#, 4#, 2#), 0#, TOL_ABS_TIGHT
    AssertClose "z(1,4,2)", K_STATS_Normal_ZScore(1#, 4#, 2#), -1.5, TOL_ABS_TIGHT
End Sub


Private Sub Test_NF_IntervalProbability()
'
'==============================================================================
' Test_NF_IntervalProbability
'------------------------------------------------------------------------------
' PURPOSE
'   Verifies central and same-tail normal interval probabilities.
'
' BEHAVIOR
'   - Prints one section heading.
'   - Executes silent passing assertions.
'   - Records detailed output only when an assertion fails.
'
' DEPENDENCIES
'   - Production functions under test
'   - Shared assertion helpers in this module
'
' CALLED FROM
'   - RunNormalFamilySuite
'
' UPDATED
'   2026-07-11
'==============================================================================
'
    Debug.Print "-- Interval probability"
    'Standard: symmetric 95% band
    AssertClose "std P(-1.96..1.96)=0.95", _
        K_STATS_NormalStandard_IntervalProbability(-1.95996398454005, 1.95996398454005), _
        0.95, TOL_ABS_TIGHT
    'Standard: +/- 1 sd
    AssertClose "std P(-1..1)", _
        K_STATS_NormalStandard_IntervalProbability(-1#, 1#), 0.682689492137086, TOL_ABS_TIGHT
    'Standard: equal bounds -> 0
    AssertClose "std P(1..1)=0", _
        K_STATS_NormalStandard_IntervalProbability(1#, 1#), 0#, TOL_ABS_TIGHT
    'General: mean 10 sd 2, 8..12 equals P(-1..1)
    AssertClose "gen P(8..12 | 10,2)", _
        K_STATS_Normal_IntervalProbability(8#, 12#, 10#, 2#), 0.682689492137086, TOL_ABS_TIGHT
    'General: full standard band via defaults
    AssertClose "gen P(-1..1 | 0,1)", _
        K_STATS_Normal_IntervalProbability(-1#, 1#), 0.682689492137086, TOL_ABS_TIGHT

    'REGRESSION: direct CDF subtraction returned zero because Phi(9) and
    'Phi(10) both round to one. Tail-oriented branching preserves the mass.
    AssertRelClose "std P(9..10) positive tail", _
        K_STATS_NormalStandard_IntervalProbability(9#, 10#), _
        1.1285122074236E-19, TOL_REL_TAIL

    'The negative-tail branch must preserve the symmetric interval mass.
    AssertRelClose "std P(-10..-9) negative tail", _
        K_STATS_NormalStandard_IntervalProbability(-10#, -9#), _
        1.1285122074236E-19, TOL_REL_TAIL

    'General-normal standardization must reach the same stable kernel.
    AssertRelClose "gen P(18..20 | 0,2) positive tail", _
        K_STATS_Normal_IntervalProbability(18#, 20#, 0#, 2#), _
        1.1285122074236E-19, TOL_REL_TAIL
End Sub


Private Sub Test_NF_FastInverse()
'
'==============================================================================
' Test_NF_FastInverse
'------------------------------------------------------------------------------
' PURPOSE
'   Verifies the raw fast inverse-normal worksheet surface.
'
' BEHAVIOR
'   - Prints one section heading.
'   - Executes silent passing assertions.
'   - Records detailed output only when an assertion fails.
'
' DEPENDENCIES
'   - Production functions under test
'   - Shared assertion helpers in this module
'
' CALLED FROM
'   - RunNormalFamilySuite
'
' UPDATED
'   2026-07-11
'==============================================================================
'
    Dim R As Double

    Debug.Print "-- Fast inverse (raw Acklam, ~1E-9)"

    'REGRESSION N1: the zero-stub kernel made this return 0# for every input,
    'which would have made every Monte Carlo shock identical
    AssertTrue "fast inverse is not a zero stub", _
        (Abs(K_STATS_NormalStandard_InverseCumulativeFast(0.975)) > 1#)

    'Accuracy against known quantile (loose tolerance)
    AssertClose "fast InvPhi(0.975)", _
        K_STATS_NormalStandard_InverseCumulativeFast(0.975), 1.95996398454005, TOL_ABS_LOOSE
    AssertClose "fast InvPhi(0.5)", _
        K_STATS_NormalStandard_InverseCumulativeFast(0.5), 0#, TOL_ABS_LOOSE
    'Endpoint clipping: p=0 must not error and must return a large negative number
    R = K_STATS_NormalStandard_InverseCumulativeFast(0#)
    AssertTrue "fast InvPhi(0) clipped negative", (R < -5#)
    'Endpoint clipping: p=1 must return a large positive number
    R = K_STATS_NormalStandard_InverseCumulativeFast(1#)
    AssertTrue "fast InvPhi(1) clipped positive", (R > 5#)
End Sub


Private Sub Test_NF_LognormalCore()
'
'==============================================================================
' Test_NF_LognormalCore
'------------------------------------------------------------------------------
' PURPOSE
'   Verifies lognormal density, CDF and inverse behavior.
'
' BEHAVIOR
'   - Prints one section heading.
'   - Executes silent passing assertions.
'   - Records detailed output only when an assertion fails.
'
' DEPENDENCIES
'   - Production functions under test
'   - Shared assertion helpers in this module
'
' CALLED FROM
'   - RunNormalFamilySuite
'
' UPDATED
'   2026-07-11
'==============================================================================
'
    Debug.Print "-- Lognormal core (MeanLog=0, StdDevLog=1)"
    'Density at x=1 equals phi(0)/(1*1)
    AssertClose "logn density@1", _
        K_STATS_Lognormal_Density(1#, 0#, 1#), 0.398942280401433, TOL_ABS_TIGHT
    'CDF at x=1 equals Phi(0)=0.5
    AssertClose "logn cdf@1", K_STATS_Lognormal_Cumulative(1#, 0#, 1#), 0.5, TOL_ABS_TIGHT
    'Density at x<=0 returns 0: a positive-support density is zero outside its
    'support, matching the CDF returning 0 and the survival returning 1 there.
    AssertClose "logn density@0 = 0", _
        K_STATS_Lognormal_Density(0#, 0#, 1#), 0#, TOL_ABS_TIGHT
    AssertClose "logn density@-5 = 0", _
        K_STATS_Lognormal_Density(-5#, 0#, 1#), 0#, TOL_ABS_TIGHT
    'CDF at x<=0 returns 0
    AssertClose "logn cdf@0", K_STATS_Lognormal_Cumulative(0#, 0#, 1#), 0#, TOL_ABS_TIGHT
    AssertClose "logn cdf@-5", K_STATS_Lognormal_Cumulative(-5#, 0#, 1#), 0#, TOL_ABS_TIGHT
    'Inverse at p=0.5 returns Exp(0)=1
    AssertClose "logn inv@0.5", _
        K_STATS_Lognormal_InverseCumulative(0.5, 0#, 1#), 1#, TOL_ABS_TIGHT
End Sub


Private Sub Test_NF_LognormalMoments()
'
'==============================================================================
' Test_NF_LognormalMoments
'------------------------------------------------------------------------------
' PURPOSE
'   Verifies lognormal mean, variance and standard deviation.
'
' BEHAVIOR
'   - Prints one section heading.
'   - Executes silent passing assertions.
'   - Records detailed output only when an assertion fails.
'
' DEPENDENCIES
'   - Production functions under test
'   - Shared assertion helpers in this module
'
' CALLED FROM
'   - RunNormalFamilySuite
'
' UPDATED
'   2026-07-11
'==============================================================================
'
    Dim VarV As Variant
    Dim SdV As Variant

    Debug.Print "-- Lognormal moments (MeanLog=0, StdDevLog=1)"
    'Mean = Exp(0.5)
    AssertClose "logn mean", _
        K_STATS_Lognormal_Mean(0#, 1#), 1.64872127070013, TOL_ABS_TIGHT
    'Variance = (e-1)*e
    AssertClose "logn variance", _
        K_STATS_Lognormal_Variance(0#, 1#), 4.6707742704716, TOL_ABS_TIGHT
    'StdDev = Sqr(variance)
    AssertClose "logn stddev", _
        K_STATS_Lognormal_StdDev(0#, 1#), 2.16119741589509, TOL_ABS_TIGHT
    'Consistency: StdDev^2 == Variance
    VarV = K_STATS_Lognormal_Variance(0#, 1#)
    SdV = K_STATS_Lognormal_StdDev(0#, 1#)
    If IsError(VarV) Or IsError(SdV) Then
        RecordResult "logn stddev^2 == variance (errored)", False
    Else
        AssertClose "logn stddev^2 == variance", CDbl(SdV) * CDbl(SdV), CDbl(VarV), TOL_ABS_LOOSE
    End If

    'REGRESSION: Exp(sigma^2) - 1 rounded to zero at sigma = 1E-8.
    AssertRelClose "logn tiny-sigma variance", _
        K_STATS_Lognormal_Variance(0#, 0.00000001), _
        1E-16, TOL_REL_TIGHT

    AssertRelClose "logn tiny-sigma stddev", _
        K_STATS_Lognormal_StdDev(0#, 0.00000001), _
        0.00000001, TOL_REL_TIGHT
End Sub


Private Sub Test_NF_LognormalUnderflow()
'
'==============================================================================
' Test_NF_LognormalUnderflow
'------------------------------------------------------------------------------
' PURPOSE
'   Verifies lognormal underflow and overflow classification.
'
' BEHAVIOR
'   - Prints one section heading.
'   - Executes silent passing assertions.
'   - Records detailed output only when an assertion fails.
'
' DEPENDENCIES
'   - Production functions under test
'   - Shared assertion helpers in this module
'
' CALLED FROM
'   - RunNormalFamilySuite
'
' UPDATED
'   2026-07-11
'==============================================================================
'
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

    Debug.Print "-- Lognormal moment log-reconstruction (REGRESSION N4)"

    'Magnitudes that are extreme but REPRESENTABLE were previously lost: the old
    'code evaluated the large and small exponential factors separately, so an
    'intermediate factor overflowed (-> #NUM!) or underflowed (-> 0) even though
    'the final moment is finite. The single-log reconstruction returns it.
    'Overflow-side: small sigma pulls the large exponential back into range
    '(old code returned #NUM!).
    AssertRelClose "logn variance overflow-recovery", _
        K_STATS_Lognormal_Variance(354.995, 0.1), 2.24520206650824E+306, TOL_REL_TIGHT
    AssertRelClose "logn stddev overflow-recovery", _
        K_STATS_Lognormal_StdDev(354.995, 0.1), 1.49839983532708E+153, TOL_REL_TIGHT
    'Underflow-side: large sigma lifts the tiny exponential back into range
    '(old code returned 0).
    AssertRelClose "logn variance underflow-recovery", _
        K_STATS_Lognormal_Variance(-425#, 10#), 5.11195194865116E-283, TOL_REL_TIGHT
    AssertRelClose "logn stddev underflow-recovery", _
        K_STATS_Lognormal_StdDev(-425#, 10#), 7.14979156944533E-142, TOL_REL_TIGHT

    'Degenerate zero log-variance: X is constant, so both moments are exactly 0.
    AssertClose "logn variance sigma=0 -> 0", _
        K_STATS_Lognormal_Variance(2#, 0#), 0#, 0#
    AssertClose "logn stddev sigma=0 -> 0", _
        K_STATS_Lognormal_StdDev(2#, 0#), 0#, 0#
End Sub


Private Sub Test_NF_ParameterRoundTrip()
'
'==============================================================================
' Test_NF_ParameterRoundTrip
'------------------------------------------------------------------------------
' PURPOSE
'   Verifies arithmetic-to-log parameter conversion and round-trips.
'
' BEHAVIOR
'   - Prints one section heading.
'   - Executes silent passing assertions.
'   - Records detailed output only when an assertion fails.
'
' DEPENDENCIES
'   - Production functions under test
'   - Shared assertion helpers in this module
'
' CALLED FROM
'   - RunNormalFamilySuite
'
' UPDATED
'   2026-07-11
'==============================================================================
'
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
    AssertClose "MeanLog", MeanLog, 0.662834806151016, TOL_ABS_LOOSE
    AssertClose "StdDevLog", StdDevLog, 0.246221445044987, TOL_ABS_LOOSE

    'Degenerate point mass: StdDev = 0 is a valid conversion, not a domain error.
    'MeanLog = Log(Mean), StdDevLog = 0, which reproduces the input moments.
    P = K_STATS_Lognormal_ParametersFromMeanStdDev(2#, 0#)
    If IsError(P) Then
        RecordResult "param StdDev=0 returned error", False
    Else
        AssertClose "param StdDev=0 -> MeanLog = Log(2)", P(1, 1), 0.693147180559945, TOL_ABS_TIGHT
        AssertClose "param StdDev=0 -> StdDevLog = 0", P(1, 2), 0#, TOL_ABS_TIGHT
    End If
    'Negative StdDev is still rejected
    AssertIsError "param StdDev<0", K_STATS_Lognormal_ParametersFromMeanStdDev(2#, -1#)

    'Round-trip: feeding the log params back must recover Mean and StdDev.
    'This ties K_STATS_Lognormal_StdDev to the conversion.
    AssertClose "roundtrip Mean", _
        K_STATS_Lognormal_Mean(MeanLog, StdDevLog), 2#, TOL_ABS_LOOSE
    AssertClose "roundtrip StdDev", _
        K_STATS_Lognormal_StdDev(MeanLog, StdDevLog), 0.5, TOL_ABS_LOOSE

    'REGRESSION: Log(1 + CV^2) rounded to zero for CV = 1E-10.
    P = K_STATS_Lognormal_ParametersFromMeanStdDev(1#, 0.0000000001)

    If IsError(P) Then
        RecordResult "tiny-CV parameter conversion returned error", False
    Else
        AssertRelClose "tiny-CV StdDevLog", _
            P(1, 2), 0.0000000001, TOL_REL_TIGHT
        AssertClose "tiny-CV MeanLog", _
            P(1, 1), -5E-21, 1E-20
    End If
End Sub


Private Sub Test_NF_ErrorContract()
'
'==============================================================================
' Test_NF_ErrorContract
'------------------------------------------------------------------------------
' PURPOSE
'   Verifies normal-family domain errors and diagnostic-status behavior.
'
' BEHAVIOR
'   - Prints one section heading.
'   - Executes silent passing assertions.
'   - Records detailed output only when an assertion fails.
'
' DEPENDENCIES
'   - Production functions under test
'   - Shared assertion helpers in this module
'
' CALLED FROM
'   - RunNormalFamilySuite
'
' UPDATED
'   2026-07-11
'==============================================================================
'
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
    AssertIsError "logn density sd=0", K_STATS_Lognormal_Density(1#, 0#, 0#)
    'Survival honours the same domain contract as the CDF
    AssertIsError "normal survival sd=0", K_STATS_Normal_Survival(0#, 0#, 0#)
    AssertIsError "logn survival sd=0", K_STATS_Lognormal_Survival(1#, 0#, 0#)
    'Inverse survival honours the same probability and parameter contract
    AssertIsError "std inv survival p=0", K_STATS_NormalStandard_InverseSurvival(0#)
    AssertIsError "std inv survival p=1", K_STATS_NormalStandard_InverseSurvival(1#)
    AssertIsError "normal inv survival sd=0", K_STATS_Normal_InverseSurvival(0.025, 0#, 0#)
    AssertIsError "logn inv survival sd=0", K_STATS_Lognormal_InverseSurvival(0.5, 0#, 0#)
    'Parameter conversion rejects StdDev = 0 and non-positive Mean
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
    AssertClose "normal cdf ok with status", K_STATS_Normal_Cumulative(0#, 0#, 1#, Diag), 0.5, TOL_ABS_TIGHT
    AssertTrue "NF status cleared on success", (Len(Diag) = 0)
End Sub


Private Sub Test_NF_Survival()
'
'==============================================================================
' Test_NF_Survival
'------------------------------------------------------------------------------
' PURPOSE
'   Verifies normal-family upper-tail survival Q(x) = 1 - F(x), including the
'   deep tails where a 1 - CDF subtraction collapses to zero.
'
' BEHAVIOR
'   - Prints one section heading.
'   - Executes silent passing assertions.
'   - Records detailed output only when an assertion fails.
'
' DEPENDENCIES
'   - Production functions under test
'   - Shared assertion helpers in this module
'
' CALLED FROM
'   - RunNormalFamilySuite
'
' UPDATED
'   2026-07-12
'==============================================================================
'
    Debug.Print "-- Normal family survival (upper tail)"

    'Standard survival known values
    AssertClose "std Q(0)", K_STATS_NormalStandard_Survival(0#), 0.5, TOL_ABS_TIGHT
    AssertClose "std Q(1)", K_STATS_NormalStandard_Survival(1#), 0.158655253931457, TOL_ABS_TIGHT
    AssertClose "std Q(1.96)", _
        K_STATS_NormalStandard_Survival(1.95996398454005), 0.025, TOL_ABS_TIGHT
    AssertClose "std Q(-1.96)", _
        K_STATS_NormalStandard_Survival(-1.95996398454005), 0.975, TOL_ABS_TIGHT

    'Deep right tails: the p-values that 1 - CDF cannot express
    AssertRelClose "std Q(6) tail", _
        K_STATS_NormalStandard_Survival(6#), 9.86587645037695E-10, TOL_REL_TAIL
    AssertRelClose "std Q(9) tail", _
        K_STATS_NormalStandard_Survival(9#), 1.12858840595383E-19, TOL_REL_TAIL

    'Survival preserves representable tails to the same depth as the CDF, by
    'symmetry Q(z) = Phi(-z).
    AssertRelClose "std Q(37.5) representable tail", _
        K_STATS_NormalStandard_Survival(37.5), 4.60535300958195E-308, TOL_REL_TAIL
    AssertRelClose "std Q(38) subnormal tail", _
        K_STATS_NormalStandard_Survival(38#), 2.88542835100396E-316, TOL_REL_LOOSE

    'The CDF-based route really does collapse, which is why Survival exists.
    AssertTrue "1 - Phi(9) is exactly zero", _
        ((1# - CDbl(K_STATS_NormalStandard_Cumulative(9#))) = 0#)

    'Survival and cumulative sum to one in the well-conditioned region.
    AssertClose "std Q + Phi = 1", _
        CDbl(K_STATS_NormalStandard_Survival(1.3)) + _
        CDbl(K_STATS_NormalStandard_Cumulative(1.3)), 1#, TOL_ABS_TIGHT

    'General normal survival
    AssertClose "gen Q(12 | 10,2)", _
        K_STATS_Normal_Survival(12#, 10#, 2#), 0.158655253931457, TOL_ABS_TIGHT
    AssertClose "gen Q(mean)", K_STATS_Normal_Survival(10#, 10#, 2#), 0.5, TOL_ABS_TIGHT

    'Lognormal survival
    AssertClose "logn Q(1 | 0,1)", _
        K_STATS_Lognormal_Survival(1#, 0#, 1#), 0.5, TOL_ABS_TIGHT
    'Non-positive points carry the full mass above them (mirrors CDF = 0 there)
    AssertClose "logn Q(0) = 1", K_STATS_Lognormal_Survival(0#, 0#, 1#), 1#, TOL_ABS_TIGHT
    AssertClose "logn Q(-5) = 1", K_STATS_Lognormal_Survival(-5#, 0#, 1#), 1#, TOL_ABS_TIGHT

    'Lognormal complement  Q(x) + F(x) = 1
    AssertClose "logn Q + F = 1 at x=2", _
        CDbl(K_STATS_Lognormal_Survival(2#, 0#, 1#)) + _
        CDbl(K_STATS_Lognormal_Cumulative(2#, 0#, 1#)), 1#, TOL_ABS_TIGHT
End Sub


Private Sub Test_NF_MagnitudePolicy()
'
'==============================================================================
' Test_NF_MagnitudePolicy
'------------------------------------------------------------------------------
' PURPOSE
'   Locks the Core magnitude split in place for the normal family: the standard
'   normal routines accept any finite argument, while the routines that genuinely
'   need the 1E100 restriction keep rejecting beyond it.
'
' WHY
'   1E200 is a finite Double. The standard-normal kernels cut off at Abs(Z) > 37
'   or 38 and return 0 or 1 without arithmetic, so there is nothing for a
'   magnitude bound to protect. The lognormal moment routines square StdDevLog
'   and the general-normal routines standardize by a division, so they keep the
'   conservative domain.
'
' CALLED FROM
'   - RunNormalFamilySuite
'
' UPDATED
'   2026-07-17
'==============================================================================
'
    Debug.Print "-- Normal family magnitude policy (finite vs supported magnitude)"

    'The Core predicates really do differ
    AssertTrue "PROB_IsFinite accepts 1E200", PROB_IsFinite(1E+200)
    AssertTrue "supported magnitude rejects 1E200", _
        (Not PROB_IsWithinSupportedMagnitude(1E+200))

    'Standard normal: any finite Z is accepted and saturates through the cutoffs
    AssertClose "std density(1E200) = 0", _
        K_STATS_NormalStandard_Density(1E+200), 0#, TOL_ABS_TIGHT
    AssertClose "std density(-1E200) = 0", _
        K_STATS_NormalStandard_Density(-1E+200), 0#, TOL_ABS_TIGHT
    AssertClose "std cumulative(1E200) = 1", _
        K_STATS_NormalStandard_Cumulative(1E+200), 1#, TOL_ABS_TIGHT
    AssertClose "std cumulative(-1E200) = 0", _
        K_STATS_NormalStandard_Cumulative(-1E+200), 0#, TOL_ABS_TIGHT
    AssertClose "std survival(1E200) = 0", _
        K_STATS_NormalStandard_Survival(1E+200), 0#, TOL_ABS_TIGHT
    AssertClose "std survival(-1E200) = 1", _
        K_STATS_NormalStandard_Survival(-1E+200), 1#, TOL_ABS_TIGHT
    AssertClose "std interval(-1E200, 1E200) = 1", _
        K_STATS_NormalStandard_IntervalProbability(-1E+200, 1E+200), 1#, TOL_ABS_TIGHT

    'Routines that need the restriction still enforce it
    AssertIsError "logn mean rejects MeanLog 1E200", _
        K_STATS_Lognormal_Mean(1E+200, 1#)
    AssertIsError "logn mean rejects StdDevLog 1E200", _
        K_STATS_Lognormal_Mean(0#, 1E+200)
    AssertIsError "normal density rejects X 1E200", _
        K_STATS_Normal_Density(1E+200, 0#, 1#)

    'Standardization overflow (formerly the KNOWN GAP) is now a clean numeric
    'error via PROB_TryStandardize, not an unexpected xlErrValue.
    AssertIsError "normal density tiny StdDev overflow -> xlErrNum", _
        K_STATS_Normal_Density(9E+99, 0#, 1E-300)
End Sub


Private Sub Test_NF_InverseSurvival()
'
'==============================================================================
' Test_NF_InverseSurvival
'------------------------------------------------------------------------------
' PURPOSE
'   Verifies the normal-family inverse survival (upper-tail quantile), including
'   the small exceedance probabilities where InverseCumulative(1 - q) fails
'   because 1 - q has rounded to exactly one.
'
' BEHAVIOR
'   - Prints one section heading.
'   - Executes silent passing assertions.
'   - Records detailed output only when an assertion fails.
'
' DEPENDENCIES
'   - Production functions under test
'   - Shared assertion helpers in this module
'
' CALLED FROM
'   - RunNormalFamilySuite
'
' UPDATED
'   2026-07-16
'==============================================================================
'
    Dim QValues             As Variant         'Exceedance probabilities
    Dim I                   As Long            'Loop index
    Dim Q                   As Double          'One exceedance probability
    Dim Z                   As Variant         'Recovered quantile
    Dim Back                As Variant         'Round-tripped probability
    Dim OverflowStatus      As String          'Captured diagnostic message
    Dim OverflowResult      As Variant         'Result of an overflow call

    Debug.Print "-- Normal family inverse survival (upper-tail quantile)"

    'Known one-sided critical values
    AssertClose "std isf(0.5) = 0", _
        K_STATS_NormalStandard_InverseSurvival(0.5), 0#, TOL_ABS_TIGHT
    AssertClose "std isf(0.025) = 1.96", _
        K_STATS_NormalStandard_InverseSurvival(0.025), 1.95996398454005, TOL_ABS_TIGHT
    AssertClose "std isf(0.05) = 1.645", _
        K_STATS_NormalStandard_InverseSurvival(0.05), 1.64485362695147, TOL_ABS_TIGHT
    AssertClose "std isf(0.001) = 3.09", _
        K_STATS_NormalStandard_InverseSurvival(0.001), 3.09023230616781, TOL_ABS_TIGHT
    'Half of 0.1% in each tail: the 99.9% central interval half-width
    AssertClose "std isf(0.0005) = 3.2905 sigma", _
        K_STATS_NormalStandard_InverseSurvival(0.0005), 3.29052673149189, TOL_ABS_TIGHT

    'Deep tails: exceedance probabilities far below the central range
    AssertRelClose "std isf(1E-10)", _
        K_STATS_NormalStandard_InverseSurvival(0.0000000001), 6.36134090240406, TOL_REL_TAIL
    AssertRelClose "std isf(1E-15)", _
        K_STATS_NormalStandard_InverseSurvival(0.000000000000001), 7.941345326171, TOL_REL_TAIL
    AssertRelClose "std isf(1E-18)", _
        K_STATS_NormalStandard_InverseSurvival(1E-18), 8.75729034878232, TOL_REL_TAIL

    'The composed route really does fail, which is why this function exists:
    '1 - 1E-18 rounds to exactly one, which the inverse cumulative must reject.
    AssertTrue "1 - 1E-18 rounds to exactly one", ((1# - 1E-18) = 1#)
    AssertIsError "InverseCumulative(1 - 1E-18) fails", _
        K_STATS_NormalStandard_InverseCumulative(1# - 1E-18)

    'Reflection identity  isf(q) = -InverseCumulative(q)
    QValues = Array(0.3, 0.025, 0.000001, 0.000000000000001)
    For I = LBound(QValues) To UBound(QValues)
        Q = QValues(I)
        Z = K_STATS_NormalStandard_InverseSurvival(Q)
        Back = K_STATS_NormalStandard_InverseCumulative(Q)
        If IsError(Z) Or IsError(Back) Then
            RecordResult "isf = -invcdf at q=" & Q & " (errored)", False
        Else
            AssertClose "isf = -invcdf at q=" & Q, CDbl(Z), -CDbl(Back), TOL_ABS_TIGHT
        End If
    Next I

    'Round-trip  Survival(isf(q)) = q
    QValues = Array(0.25, 0.025, 0.001, 0.000001)
    For I = LBound(QValues) To UBound(QValues)
        Q = QValues(I)
        Z = K_STATS_NormalStandard_InverseSurvival(Q)
        If IsError(Z) Then
            RecordResult "isf round-trip q=" & Q & " (errored)", False
        Else
            Back = K_STATS_NormalStandard_Survival(CDbl(Z))
            If IsError(Back) Then
                RecordResult "isf round-trip q=" & Q & " (survival errored)", False
            Else
                AssertRelClose "isf round-trip q=" & Q, CDbl(Back), Q, TOL_REL_TAIL
            End If
        End If
    Next I

    'General normal
    AssertClose "gen isf(0.025 | 10,2)", _
        K_STATS_Normal_InverseSurvival(0.025, 10#, 2#), 13.9199279690801, TOL_ABS_TIGHT
    AssertClose "gen isf(0.5 | 10,2) = mean", _
        K_STATS_Normal_InverseSurvival(0.5, 10#, 2#), 10#, TOL_ABS_TIGHT

    'Lognormal
    AssertClose "logn isf(0.5 | 0,1) = 1", _
        K_STATS_Lognormal_InverseSurvival(0.5, 0#, 1#), 1#, TOL_ABS_TIGHT
    AssertClose "logn isf(0.025 | 0,1)", _
        K_STATS_Lognormal_InverseSurvival(0.025, 0#, 1#), 7.09907138423134, TOL_ABS_TIGHT

    'Lognormal round-trip  Survival(isf(q)) = q
    Z = K_STATS_Lognormal_InverseSurvival(0.001, 0#, 1#)
    If IsError(Z) Then
        RecordResult "logn isf round-trip (errored)", False
    Else
        AssertRelClose "logn isf round-trip q=0.001", _
            K_STATS_Lognormal_Survival(CDbl(Z), 0#, 1#), 0.001, TOL_REL_TAIL
    End If

'------------------------------------------------------------------------------
' OVERFLOW AND BOUNDARY CONTRACT
'------------------------------------------------------------------------------
    'General normal near the supported magnitude boundary is a SUCCESS, not an
    'error: StdDev = 9E99 passes the 1E100 cap and 9E99 * isf(0.025) stays finite.
    AssertRelClose "normal isf near-boundary rescaling is finite", _
        K_STATS_Normal_InverseSurvival(0.025, 0#, 9E+99), _
        1.76396758608605E+100, TOL_REL_TIGHT

    'Lognormal inverse survival: the exponential of the reconstructed log-quantile
    'overflows. MeanLog = 700, StdDevLog = 20, isf(1E-100) ~ 21.27, so the
    'log-quantile ~ 1125 and Exp(1125) overflows. This must be a clean xlErrNum.
    AssertErrorCode "lognormal isf exponential overflow is #NUM", _
        K_STATS_Lognormal_InverseSurvival(1E-100, 700#, 20#), xlErrNum

    'The overflow must also populate the diagnostic Status with an explanatory
    'message rather than failing silently.
    OverflowStatus = vbNullString
    OverflowResult = K_STATS_Lognormal_InverseSurvival(1E-100, 700#, 20#, OverflowStatus)
    AssertTrue "lognormal isf overflow is an error", IsError(OverflowResult)
    AssertTrue "lognormal isf overflow status mentions overflow", _
        (InStr(1, LCase$(OverflowStatus), "overflow") > 0)

    'A valid inverse survival must leave Status empty (no false diagnostics).
    OverflowStatus = "sentinel"
    OverflowResult = K_STATS_Lognormal_InverseSurvival(0.5, 0#, 1#, OverflowStatus)
    AssertTrue "lognormal isf success clears status", (OverflowStatus = vbNullString)
End Sub


Private Sub Test_NF_OverflowContract()
'
'==============================================================================
' Test_NF_OverflowContract
'------------------------------------------------------------------------------
' PURPOSE
'   Verifies normal-family overflow and underflow policy.
'
' BEHAVIOR
'   - Prints one section heading.
'   - Executes silent passing assertions.
'   - Records detailed output only when an assertion fails.
'
' DEPENDENCIES
'   - Production functions under test
'   - Shared assertion helpers in this module
'
' CALLED FROM
'   - RunNormalFamilySuite
'
' UPDATED
'   2026-07-11
'==============================================================================
'
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
'
'==============================================================================
' Test_TF_StudentTDensity
'------------------------------------------------------------------------------
' PURPOSE
'   Verifies Student t density values and large-parameter behavior.
'
' BEHAVIOR
'   - Prints one section heading.
'   - Executes silent passing assertions.
'   - Records detailed output only when an assertion fails.
'
' DEPENDENCIES
'   - Production functions under test
'   - Shared assertion helpers in this module
'
' CALLED FROM
'   - RunTFamilySuite
'
' UPDATED
'   2026-07-11
'==============================================================================
'
    Debug.Print "-- Student t density"
    AssertClose "t pdf(0,1)", K_STATS_StudentT_Density(0#, 1#), 0.318309886183791, TOL_ABS_TIGHT
    AssertClose "t pdf(0,5)", K_STATS_StudentT_Density(0#, 5#), 0.379606689822494, TOL_ABS_TIGHT
    AssertClose "t pdf(1,5)", K_STATS_StudentT_Density(1#, 5#), 0.219679797350981, TOL_ABS_TIGHT
    AssertClose "t pdf(2,10)", K_STATS_StudentT_Density(2#, 10#), 6.11457663212182E-02, TOL_ABS_TIGHT
    AssertClose "t pdf(0,30)", K_STATS_StudentT_Density(0#, 30#), 0.395632184894098, TOL_ABS_TIGHT

    'Symmetry
    AssertClose "t pdf symmetric", _
        CDbl(K_STATS_StudentT_Density(-2#, 7#)) - CDbl(K_STATS_StudentT_Density(2#, 7#)), 0#, 0#

    'Large df must converge on the standard normal density; the old log-gamma
    'subtraction lost seven digits here
    AssertClose "t pdf(0,1e6) -> phi(0)", K_STATS_StudentT_Density(0#, 1000000#), _
        0.398942180665875, TOL_ABS_LARGE_DF

    'Far tail underflows to a valid zero, not an error
    AssertClose "t pdf far tail = 0", K_STATS_StudentT_Density(1E+50, 30#), 0#, 0#
    'Small-df and huge-X paths must remain numeric, not fall into runtime errors.
    AssertTrue "t pdf tiny df at zero numeric", _
        (Not IsError(K_STATS_StudentT_Density(0#, 0.000000000001)))
    AssertClose "t pdf huge x underflows", _
        K_STATS_StudentT_Density(1E+200, 5#), 0#, 0#
End Sub


Private Sub Test_TF_StudentTCumulative()
'
'==============================================================================
' Test_TF_StudentTCumulative
'------------------------------------------------------------------------------
' PURPOSE
'   Verifies Student t cumulative probabilities and closed-form seams.
'
' BEHAVIOR
'   - Prints one section heading.
'   - Executes silent passing assertions.
'   - Records detailed output only when an assertion fails.
'
' DEPENDENCIES
'   - Production functions under test
'   - Shared assertion helpers in this module
'
' CALLED FROM
'   - RunTFamilySuite
'
' UPDATED
'   2026-07-11
'==============================================================================
'
    Debug.Print "-- Student t cumulative"
    AssertClose "t cdf(0,5)", K_STATS_StudentT_Cumulative(0#, 5#), 0.5, 0#
    AssertClose "t cdf(1,5)", K_STATS_StudentT_Cumulative(1#, 5#), 0.818391266175439, TOL_ABS_TIGHT
    AssertClose "t cdf(2,10)", K_STATS_StudentT_Cumulative(2#, 10#), 0.96330598261463, TOL_ABS_TIGHT
    AssertClose "t cdf(-1,1)", K_STATS_StudentT_Cumulative(-1#, 1#), 0.25, TOL_ABS_TIGHT
    AssertClose "t cdf(2.5,3)", K_STATS_StudentT_Cumulative(2.5, 3#), 0.956146676495967, TOL_ABS_TIGHT

    'The df = 2 closed form must agree with the general beta route at df = 2.0000001
    AssertClose "t cdf df=2 continuity", _
        CDbl(K_STATS_StudentT_Cumulative(1.5, 2#)) - _
        CDbl(K_STATS_StudentT_Cumulative(1.5, 2.0000001)), 0#, 0.0000001
    'Tiny Cauchy argument formerly overflowed in 1 / AbsX.
    AssertClose "t cdf Cauchy tiny x", _
        K_STATS_StudentT_Cumulative(9.99988867182683E-321, 1#), 0.5, 0#

    'Tiny-argument local series at small degrees of freedom.
    AssertClose "t cdf tiny x small df", _
        K_STATS_StudentT_Cumulative(0.0000000001, 0.1), _
        0.500000000014809, 0.000000000000002
End Sub


Private Sub Test_TF_StudentTCentralRegion()
'
'==============================================================================
' Test_TF_StudentTCentralRegion
'------------------------------------------------------------------------------
' PURPOSE
'   Verifies Student t central-region precision near zero.
'
' BEHAVIOR
'   - Prints one section heading.
'   - Executes silent passing assertions.
'   - Records detailed output only when an assertion fails.
'
' DEPENDENCIES
'   - Production functions under test
'   - Shared assertion helpers in this module
'
' CALLED FROM
'   - RunTFamilySuite
'
' UPDATED
'   2026-07-11
'==============================================================================
'
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
'
'==============================================================================
' Test_TF_StudentTSurvival
'------------------------------------------------------------------------------
' PURPOSE
'   Verifies Student t right-tail survival probabilities.
'
' BEHAVIOR
'   - Prints one section heading.
'   - Executes silent passing assertions.
'   - Records detailed output only when an assertion fails.
'
' DEPENDENCIES
'   - Production functions under test
'   - Shared assertion helpers in this module
'
' CALLED FROM
'   - RunTFamilySuite
'
' UPDATED
'   2026-07-11
'==============================================================================
'
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
        CDbl(K_STATS_StudentT_Cumulative(1.3, 9#)), 1#, TOL_ABS_TIGHT

    'Negative arguments are accepted, unlike Excel's T.DIST.RT
    AssertClose "t sf(-1,1)", K_STATS_StudentT_Survival(-1#, 1#), 0.75, TOL_ABS_TIGHT
End Sub


Private Sub Test_TF_StudentTInverse()
'
'==============================================================================
' Test_TF_StudentTInverse
'------------------------------------------------------------------------------
' PURPOSE
'   Verifies Student t inverse cumulative values and special branches.
'
' BEHAVIOR
'   - Prints one section heading.
'   - Executes silent passing assertions.
'   - Records detailed output only when an assertion fails.
'
' DEPENDENCIES
'   - Production functions under test
'   - Shared assertion helpers in this module
'
' CALLED FROM
'   - RunTFamilySuite
'
' UPDATED
'   2026-07-11
'==============================================================================
'
    Debug.Print "-- Student t inverse"
    AssertClose "t inv(0.975,10)", K_STATS_StudentT_InverseCumulative(0.975, 10#), _
        2.22813885198627, TOL_ABS_TIGHT
    AssertClose "t inv(0.95,5)", K_STATS_StudentT_InverseCumulative(0.95, 5#), _
        2.01504837333302, TOL_ABS_TIGHT
    AssertClose "t inv(0.005,20)", K_STATS_StudentT_InverseCumulative(0.005, 20#), _
        -2.84533970978611, TOL_ABS_TIGHT
    AssertClose "t inv(0.5,7)", K_STATS_StudentT_InverseCumulative(0.5, 7#), 0#, 0#

    'REGRESSION T4: the old 1E+12 bracket cap returned CVErr(xlErrNum) here
    AssertRelClose "t inv(1e-14,1)", K_STATS_StudentT_InverseCumulative(0.00000000000001, 1#), _
        -31830988618379.1, TOL_REL_TIGHT
    AssertTrue "t inv(1e-14,0.5) is a number", _
        (Not IsError(K_STATS_StudentT_InverseCumulative(0.00000000000001, 0.5)))
    AssertRoundTripT "t inverse small-df beta branch", 0.75, 0.1

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
'
'==============================================================================
' Test_TF_StudentTRoundTrips
'------------------------------------------------------------------------------
' PURPOSE
'   Verifies Student t inverse and cumulative round-trips.
'
' BEHAVIOR
'   - Prints one section heading.
'   - Executes silent passing assertions.
'   - Records detailed output only when an assertion fails.
'
' DEPENDENCIES
'   - Production functions under test
'   - Shared assertion helpers in this module
'
' CALLED FROM
'   - RunTFamilySuite
'
' UPDATED
'   2026-07-11
'==============================================================================
'
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
'
'==============================================================================
' Test_TF_StudentTSymmetry
'------------------------------------------------------------------------------
' PURPOSE
'   Verifies Student t symmetry identities.
'
' BEHAVIOR
'   - Prints one section heading.
'   - Executes silent passing assertions.
'   - Records detailed output only when an assertion fails.
'
' DEPENDENCIES
'   - Production functions under test
'   - Shared assertion helpers in this module
'
' CALLED FROM
'   - RunTFamilySuite
'
' UPDATED
'   2026-07-11
'==============================================================================
'
    Debug.Print "-- Student t symmetry"
    AssertClose "t cdf(-x) = 1 - cdf(x)", _
        CDbl(K_STATS_StudentT_Cumulative(-1.7, 11#)) + _
        CDbl(K_STATS_StudentT_Cumulative(1.7, 11#)), 1#, TOL_ABS_TIGHT
    AssertClose "t inv(p) = -inv(1-p)", _
        CDbl(K_STATS_StudentT_InverseCumulative(0.02, 13#)) + _
        CDbl(K_STATS_StudentT_InverseCumulative(0.98, 13#)), 0#, TOL_ABS_TIGHT
    AssertClose "t sf(x) = cdf(-x)", _
        CDbl(K_STATS_StudentT_Survival(2.2, 4#)) - _
        CDbl(K_STATS_StudentT_Cumulative(-2.2, 4#)), 0#, 0#
End Sub


'==============================================================================
' SUITE - T FAMILY: CHI-SQUARE
'==============================================================================

Private Sub Test_TF_ChiSquareDensity()
'
'==============================================================================
' Test_TF_ChiSquareDensity
'------------------------------------------------------------------------------
' PURPOSE
'   Verifies chi-square density values and support edges.
'
' BEHAVIOR
'   - Prints one section heading.
'   - Executes silent passing assertions.
'   - Records detailed output only when an assertion fails.
'
' DEPENDENCIES
'   - Production functions under test
'   - Shared assertion helpers in this module
'
' CALLED FROM
'   - RunTFamilySuite
'
' UPDATED
'   2026-07-11
'==============================================================================
'
    Debug.Print "-- Chi-square density"
    AssertClose "chi2 pdf(1,2)", K_STATS_ChiSquare_Density(1#, 2#), 0.303265329856317, TOL_ABS_TIGHT
    AssertClose "chi2 pdf(4,4)", K_STATS_ChiSquare_Density(4#, 4#), 0.135335283236613, TOL_ABS_TIGHT
    AssertClose "chi2 pdf(2,1)", K_STATS_ChiSquare_Density(2#, 1#), 0.103776874355149, TOL_ABS_TIGHT
    AssertClose "chi2 pdf(-1,3)", K_STATS_ChiSquare_Density(-1#, 3#), 0#, 0#
    AssertClose "chi2 pdf(0,2)", K_STATS_ChiSquare_Density(0#, 2#), 0.5, 0#
    AssertClose "chi2 pdf(0,3)", K_STATS_ChiSquare_Density(0#, 3#), 0#, 0#
    AssertIsError "chi2 pdf(0,1) unbounded", K_STATS_ChiSquare_Density(0#, 1#)
End Sub


Private Sub Test_TF_ChiSquareCumulative()
'
'==============================================================================
' Test_TF_ChiSquareCumulative
'------------------------------------------------------------------------------
' PURPOSE
'   Verifies chi-square cumulative probabilities.
'
' BEHAVIOR
'   - Prints one section heading.
'   - Executes silent passing assertions.
'   - Records detailed output only when an assertion fails.
'
' DEPENDENCIES
'   - Production functions under test
'   - Shared assertion helpers in this module
'
' CALLED FROM
'   - RunTFamilySuite
'
' UPDATED
'   2026-07-11
'==============================================================================
'
    Debug.Print "-- Chi-square cumulative"
    AssertClose "chi2 cdf(3.84,1)", K_STATS_ChiSquare_Cumulative(3.84, 1#), _
        0.949956478751295, TOL_ABS_TIGHT
    AssertClose "chi2 cdf(11.07,5)", K_STATS_ChiSquare_Cumulative(11.07, 5#), _
        0.949990381377595, TOL_ABS_TIGHT
    AssertClose "chi2 cdf(1,2)", K_STATS_ChiSquare_Cumulative(1#, 2#), _
        0.393469340287367, TOL_ABS_TIGHT
    AssertClose "chi2 cdf(0,5)", K_STATS_ChiSquare_Cumulative(0#, 5#), 0#, 0#
    AssertClose "chi2 cdf(-3,5)", K_STATS_ChiSquare_Cumulative(-3#, 5#), 0#, 0#

    'Series and continued-fraction branches must agree across the x = a + 1 seam
    AssertClose "chi2 branch seam", _
        CDbl(K_STATS_ChiSquare_Cumulative(11.99999999, 10#)) - _
        CDbl(K_STATS_ChiSquare_Cumulative(12.00000001, 10#)), 0#, 0.00000001
End Sub


Private Sub Test_TF_ChiSquareLargeDF()
'
'==============================================================================
' Test_TF_ChiSquareLargeDF
'------------------------------------------------------------------------------
' PURPOSE
'   Verifies chi-square behavior at large degrees of freedom.
'
' BEHAVIOR
'   - Prints one section heading.
'   - Executes silent passing assertions.
'   - Records detailed output only when an assertion fails.
'
' DEPENDENCIES
'   - Production functions under test
'   - Shared assertion helpers in this module
'
' CALLED FROM
'   - RunTFamilySuite
'
' UPDATED
'   2026-07-11
'==============================================================================
'
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
'
'==============================================================================
' Test_TF_ChiSquareSurvival
'------------------------------------------------------------------------------
' PURPOSE
'   Verifies chi-square survival probabilities.
'
' BEHAVIOR
'   - Prints one section heading.
'   - Executes silent passing assertions.
'   - Records detailed output only when an assertion fails.
'
' DEPENDENCIES
'   - Production functions under test
'   - Shared assertion helpers in this module
'
' CALLED FROM
'   - RunTFamilySuite
'
' UPDATED
'   2026-07-11
'==============================================================================
'
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
        CDbl(K_STATS_ChiSquare_Cumulative(7#, 4#)), 1#, TOL_ABS_TIGHT

    'The CDF-based route collapses; that is why Survival exists
    AssertTrue "1 - chi2 cdf(200,10) is exactly zero", _
        ((1# - CDbl(K_STATS_ChiSquare_Cumulative(200#, 10#))) = 0#)
End Sub


Private Sub Test_TF_ChiSquareInverse()
'
'==============================================================================
' Test_TF_ChiSquareInverse
'------------------------------------------------------------------------------
' PURPOSE
'   Verifies chi-square inverse values and round-trips.
'
' BEHAVIOR
'   - Prints one section heading.
'   - Executes silent passing assertions.
'   - Records detailed output only when an assertion fails.
'
' DEPENDENCIES
'   - Production functions under test
'   - Shared assertion helpers in this module
'
' CALLED FROM
'   - RunTFamilySuite
'
' UPDATED
'   2026-07-11
'==============================================================================
'
    Debug.Print "-- Chi-square inverse"
    AssertClose "chi2 inv(0.95,1)", K_STATS_ChiSquare_InverseCumulative(0.95, 1#), _
        3.84145882069412, TOL_ABS_TIGHT
    AssertClose "chi2 inv(0.95,5)", K_STATS_ChiSquare_InverseCumulative(0.95, 5#), _
        11.0704976935164, TOL_ABS_TIGHT
    AssertClose "chi2 inv(0.5,4)", K_STATS_ChiSquare_InverseCumulative(0.5, 4#), _
        3.35669398003332, TOL_ABS_TIGHT
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
'
'==============================================================================
' Test_TF_FDensity
'------------------------------------------------------------------------------
' PURPOSE
'   Verifies F-density values, support edges and extreme ratios.
'
' BEHAVIOR
'   - Prints one section heading.
'   - Executes silent passing assertions.
'   - Records detailed output only when an assertion fails.
'
' DEPENDENCIES
'   - Production functions under test
'   - Shared assertion helpers in this module
'
' CALLED FROM
'   - RunTFamilySuite
'
' UPDATED
'   2026-07-11
'==============================================================================
'
    Debug.Print "-- F density"
    AssertClose "F pdf(1,10,10)", K_STATS_F_Density(1#, 10#, 10#), 0.615234375, TOL_ABS_TIGHT
    AssertClose "F pdf(2,4,8)", K_STATS_F_Density(2#, 4#, 8#), 0.15625, TOL_ABS_TIGHT
    AssertClose "F pdf(-1,3,3)", K_STATS_F_Density(-1#, 3#, 3#), 0#, 0#
    AssertClose "F pdf(0,2,5)", K_STATS_F_Density(0#, 2#, 5#), 1#, 0#
    AssertClose "F pdf(0,3,5)", K_STATS_F_Density(0#, 3#, 5#), 0#, 0#
    AssertIsError "F pdf(0,1,5) unbounded", K_STATS_F_Density(0#, 1#, 5#)
    AssertClose "F pdf extreme positive log-ratio", _
        K_STATS_F_Density(1E+308, 1E+99, 1E-99), 0#, 0#
End Sub


Private Sub Test_TF_FCumulative()
'
'==============================================================================
' Test_TF_FCumulative
'------------------------------------------------------------------------------
' PURPOSE
'   Verifies F cumulative probabilities and reciprocal identities.
'
' BEHAVIOR
'   - Prints one section heading.
'   - Executes silent passing assertions.
'   - Records detailed output only when an assertion fails.
'
' DEPENDENCIES
'   - Production functions under test
'   - Shared assertion helpers in this module
'
' CALLED FROM
'   - RunTFamilySuite
'
' UPDATED
'   2026-07-11
'==============================================================================
'
    Debug.Print "-- F cumulative"
    AssertClose "F cdf(2.5,5,10)", K_STATS_F_Cumulative(2.5, 5#, 10#), _
        0.89799772335573, TOL_ABS_TIGHT
    AssertClose "F cdf(1,1,1)", K_STATS_F_Cumulative(1#, 1#, 1#), 0.5, TOL_ABS_TIGHT
    AssertClose "F cdf(4.96,3,10)", K_STATS_F_Cumulative(4.96, 3#, 10#), _
        0.976863670854344, TOL_ABS_TIGHT
    AssertClose "F cdf(0,4,4)", K_STATS_F_Cumulative(0#, 4#, 4#), 0#, 0#
    AssertClose "F cdf(-2,4,4)", K_STATS_F_Cumulative(-2#, 4#, 4#), 0#, 0#

    'The reciprocal identity: F(x; a, b) = 1 - F(1/x; b, a)
    AssertClose "F reciprocal identity", _
        CDbl(K_STATS_F_Cumulative(3#, 6#, 9#)) + _
        CDbl(K_STATS_F_Cumulative(1# / 3#, 9#, 6#)), 1#, TOL_ABS_TIGHT

    'Large equal degrees of freedom: the old 200-iteration beta continued fraction
    'silently stopped converging at about df = 5E+5
    AssertRelClose "F cdf(1,1e5,1e5)", K_STATS_F_Cumulative(1#, 100000#, 100000#), _
        0.5, TOL_REL_LOOSE
    AssertClose "F cdf extreme positive log-ratio", _
        K_STATS_F_Cumulative(1E+308, 1E+99, 1E-99), 1#, 0#
End Sub


Private Sub Test_TF_FSurvival()
'
'==============================================================================
' Test_TF_FSurvival
'------------------------------------------------------------------------------
' PURPOSE
'   Verifies F survival probabilities and complements.
'
' BEHAVIOR
'   - Prints one section heading.
'   - Executes silent passing assertions.
'   - Records detailed output only when an assertion fails.
'
' DEPENDENCIES
'   - Production functions under test
'   - Shared assertion helpers in this module
'
' CALLED FROM
'   - RunTFamilySuite
'
' UPDATED
'   2026-07-11
'==============================================================================
'
    Debug.Print "-- F survival"
    AssertRelClose "F sf(100,5,5)", K_STATS_F_Survival(100#, 5#, 5#), _
        5.24291335785E-05, TOL_REL_TIGHT
    AssertRelClose "F sf(2.5,5,10)", K_STATS_F_Survival(2.5, 5#, 10#), _
        0.10200227664427, TOL_REL_TIGHT
    AssertClose "F sf(0,4,4)", K_STATS_F_Survival(0#, 4#, 4#), 1#, 0#
    AssertClose "F sf + cdf = 1", _
        CDbl(K_STATS_F_Survival(2#, 7#, 12#)) + _
        CDbl(K_STATS_F_Cumulative(2#, 7#, 12#)), 1#, TOL_ABS_TIGHT
    AssertClose "F sf extreme positive log-ratio", _
        K_STATS_F_Survival(1E+308, 1E+99, 1E-99), 0#, 0#
End Sub


Private Sub Test_TF_FInverse()
'
'==============================================================================
' Test_TF_FInverse
'------------------------------------------------------------------------------
' PURPOSE
'   Verifies F inverse values and round-trips.
'
' BEHAVIOR
'   - Prints one section heading.
'   - Executes silent passing assertions.
'   - Records detailed output only when an assertion fails.
'
' DEPENDENCIES
'   - Production functions under test
'   - Shared assertion helpers in this module
'
' CALLED FROM
'   - RunTFamilySuite
'
' UPDATED
'   2026-07-11
'==============================================================================
'
    Debug.Print "-- F inverse"
    AssertClose "F inv(0.95,3,10)", K_STATS_F_InverseCumulative(0.95, 3#, 10#), _
        3.70826481904684, TOL_ABS_TIGHT
    AssertClose "F inv(0.95,5,20)", K_STATS_F_InverseCumulative(0.95, 5#, 20#), _
        2.71088983720969, TOL_ABS_TIGHT
    AssertClose "F inv(0.5,1,1)", K_STATS_F_InverseCumulative(0.5, 1#, 1#), 1#, TOL_ABS_TIGHT

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
'
'==============================================================================
' Test_TF_ErrorContract
'------------------------------------------------------------------------------
' PURPOSE
'   Verifies T-family domain, overflow and error-code contracts.
'
' BEHAVIOR
'   - Prints one section heading.
'   - Executes silent passing assertions.
'   - Records detailed output only when an assertion fails.
'
' DEPENDENCIES
'   - Production functions under test
'   - Shared assertion helpers in this module
'
' CALLED FROM
'   - RunTFamilySuite
'
' UPDATED
'   2026-07-11
'==============================================================================
'
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
    AssertClose "t cdf x huge saturates", K_STATS_StudentT_Cumulative(1E+200, 5#), 1#, 0#
    AssertClose "chi2 cdf x huge saturates", K_STATS_ChiSquare_Cumulative(1E+200, 5#), 1#, 0#

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
    AssertErrorCode "t inverse tiny-df predictable overflow is #NUM", _
        K_STATS_StudentT_InverseCumulative(0.75, 1E-200), xlErrNum
    AssertErrorCode "Cauchy inverse overflow is #NUM", _
        K_STATS_StudentT_InverseCumulative(9.99988867182683E-321, 1#), xlErrNum

    AssertErrorCode "F inverse predictable overflow is #NUM", _
        K_STATS_F_InverseCumulative(0.999999999999999, 1#, 0.01), xlErrNum
End Sub


Private Sub Test_TF_SupportEdges()
'
'==============================================================================
' Test_TF_SupportEdges
'------------------------------------------------------------------------------
' PURPOSE
'   Verifies T-family support edges and fractional degrees of freedom.
'
' BEHAVIOR
'   - Prints one section heading.
'   - Executes silent passing assertions.
'   - Records detailed output only when an assertion fails.
'
' DEPENDENCIES
'   - Production functions under test
'   - Shared assertion helpers in this module
'
' CALLED FROM
'   - RunTFamilySuite
'
' UPDATED
'   2026-07-11
'==============================================================================
'
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
        CDbl(K_STATS_ChiSquare_Cumulative(3#, 2#)), 1# - Exp(-1.5), TOL_ABS_TIGHT

    'Student t with df = 1 is Cauchy: cdf(x) = 0.5 + Atn(x)/Pi
    AssertClose "t df=1 is Cauchy", _
        CDbl(K_STATS_StudentT_Cumulative(2.7, 1#)), 0.5 + Atn(2.7) / (4# * Atn(1#)), TOL_ABS_TIGHT
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

Private Sub AssertExactlyEqual( _
    ByVal TestName As String, _
    ByVal Actual As Variant, _
    ByVal Expected As Double)
'
'==============================================================================
' AssertExactlyEqual
'------------------------------------------------------------------------------
' PURPOSE
'   Passes only when Actual and Expected are the same binary Double value.
'
' WHY
'   Decimal formatting in VBA exposes approximately 15 significant digits and
'   can hide a one-ULP difference. Direct equality is therefore required for
'   constants intended to match an independently evaluated Double exactly.
'
' INPUTS
'   TestName   Assertion label printed on failure.
'   Actual     Numeric Variant returned by the code under test.
'   Expected   Independently evaluated Double reference.
'
' BEHAVIOR
'   - Rejects a CVErr Actual.
'   - Uses direct Double equality.
'   - Reports the numerical difference when equality fails.
'
' DEPENDENCIES
'   - RecordResult
'
' CALLED FROM
'   - Constant and exact-kernel tests
'
' UPDATED
'   2026-07-11
'==============================================================================
'
'------------------------------------------------------------------------------
' VALIDATE ACTUAL VALUE
'------------------------------------------------------------------------------
    'Reject worksheet errors
        If IsError(Actual) Then
            RecordResult _
                TestName & " -> returned error, expected exact equality", _
                False

            Exit Sub
        End If

'------------------------------------------------------------------------------
' COMPARE
'------------------------------------------------------------------------------
    'Pass only when both values map to the same Double
        If CDbl(Actual) = Expected Then
            RecordResult TestName, True
        Else
            RecordResult _
                TestName & _
                " -> got " & CStr(CDbl(Actual)) & _
                ", expected " & CStr(Expected) & _
                ", difference " & CStr(CDbl(Actual) - Expected), _
                False
        End If
End Sub


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
'   Passes when Actual is numeric and lies within an absolute tolerance of the
'   expected value.
'
' INPUTS
'   TestName   Assertion label printed on failure.
'   Actual     Numeric Variant returned by the code under test.
'   Expected   Double reference value.
'   Tolerance  Non-negative absolute tolerance.
'
' BEHAVIOR
'   - Rejects CVErr and non-numeric values.
'   - Forms the difference through PROB_TryAdd so an extreme mismatch cannot
'     overflow the test harness itself.
'   - Records one assertion result.
'
' DEPENDENCIES
'   - PROB_TryAdd
'   - RecordResult
'
' CALLED FROM
'   - Test procedures throughout this module
'
' UPDATED
'   2026-07-11
'==============================================================================
'
'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Dim ActualValue         As Double          'Actual numeric value
    Dim Difference          As Double          'Actual minus expected

'------------------------------------------------------------------------------
' VALIDATE INPUTS
'------------------------------------------------------------------------------
    'Reject worksheet errors
        If IsError(Actual) Then
            RecordResult _
                TestName & _
                " -> returned error, expected " & CStr(Expected), _
                False

            Exit Sub
        End If

    'Reject non-numeric values
        If Not IsNumeric(Actual) Then
            RecordResult _
                TestName & " -> returned a non-numeric value", _
                False

            Exit Sub
        End If

    'Reject an invalid tolerance supplied by the harness
        If Tolerance < 0# Then
            RecordResult _
                TestName & " -> test tolerance must be non-negative", _
                False

            Exit Sub
        End If

'------------------------------------------------------------------------------
' COMPUTE DIFFERENCE
'------------------------------------------------------------------------------
    'Convert the actual value once
        ActualValue = CDbl(Actual)

    'Form Actual - Expected without allowing the assertion helper to overflow
        If Not PROB_TryAdd( _
            ActualValue, _
            -Expected, _
            Difference) Then

            RecordResult _
                TestName & _
                " -> absolute difference overflowed; got " & _
                CStr(ActualValue) & _
                ", expected " & CStr(Expected), _
                False

            Exit Sub
        End If

'------------------------------------------------------------------------------
' RECORD RESULT
'------------------------------------------------------------------------------
    'Compare the absolute difference with the tolerance
        If Abs(Difference) <= Tolerance Then
            RecordResult TestName, True
        Else
            RecordResult _
                TestName & _
                " -> got " & CStr(ActualValue) & _
                ", expected " & CStr(Expected) & _
                " (abs err " & CStr(Abs(Difference)) & _
                ", tol " & CStr(Tolerance) & ")", _
                False
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
'   Passes when Actual is numeric and lies within a relative tolerance of the
'   expected value.
'
' INPUTS
'   TestName          Assertion label printed on failure.
'   Actual            Numeric Variant returned by the code under test.
'   Expected          Double reference value.
'   RelativeTolerance Non-negative relative tolerance.
'
' BEHAVIOR
'   - Rejects CVErr and non-numeric values.
'   - Falls back to AssertClose when Expected is zero.
'   - Computes Actual / Expected - 1 through guarded division, avoiding an
'     overflowing subtraction between extreme values.
'
' DEPENDENCIES
'   - PROB_TryDivide
'   - AssertClose
'   - RecordResult
'
' CALLED FROM
'   - Tail, quantile and extreme-parameter tests
'
' UPDATED
'   2026-07-11
'==============================================================================
'
'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Dim ActualValue         As Double          'Actual numeric value
    Dim Ratio               As Double          'Actual divided by expected
    Dim RelativeError       As Double          'Absolute relative error

'------------------------------------------------------------------------------
' VALIDATE INPUTS
'------------------------------------------------------------------------------
    'Reject worksheet errors
        If IsError(Actual) Then
            RecordResult _
                TestName & _
                " -> returned error, expected " & CStr(Expected), _
                False

            Exit Sub
        End If

    'Reject non-numeric values
        If Not IsNumeric(Actual) Then
            RecordResult _
                TestName & " -> returned a non-numeric value", _
                False

            Exit Sub
        End If

    'Reject an invalid tolerance supplied by the harness
        If RelativeTolerance < 0# Then
            RecordResult _
                TestName & _
                " -> relative tolerance must be non-negative", _
                False

            Exit Sub
        End If

'------------------------------------------------------------------------------
' HANDLE ZERO REFERENCE
'------------------------------------------------------------------------------
    'Use an absolute comparison when the reference value is zero
        If Expected = 0# Then
            AssertClose _
                TestName, _
                Actual, _
                0#, _
                RelativeTolerance

            Exit Sub
        End If

'------------------------------------------------------------------------------
' COMPUTE RELATIVE ERROR
'------------------------------------------------------------------------------
    'Convert the actual value once
        ActualValue = CDbl(Actual)

    'Form the ratio without allowing the assertion helper to overflow
        If Not PROB_TryDivide( _
            ActualValue, _
            Expected, _
            Ratio) Then

            RecordResult _
                TestName & _
                " -> relative ratio overflowed; got " & _
                CStr(ActualValue) & _
                ", expected " & CStr(Expected), _
                False

            Exit Sub
        End If

    'Compute the relative error from the finite ratio
        RelativeError = Abs(Ratio - 1#)

'------------------------------------------------------------------------------
' RECORD RESULT
'------------------------------------------------------------------------------
    'Compare the relative error with the tolerance
        If RelativeError <= RelativeTolerance Then
            RecordResult TestName, True
        Else
            RecordResult _
                TestName & _
                " -> got " & CStr(ActualValue) & _
                ", expected " & CStr(Expected) & _
                " (rel err " & CStr(RelativeError) & _
                ", tol " & CStr(RelativeTolerance) & ")", _
                False
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
'   Passes when Actual is numeric and belongs to the closed interval [0, 1].
'
' INPUTS
'   TestName   Assertion label printed on failure.
'   Actual     Probability-like Variant returned by the code under test.
'
' BEHAVIOR
'   - Rejects CVErr and non-numeric values.
'   - Records one assertion result.
'
' DEPENDENCIES
'   - RecordResult
'
' CALLED FROM
'   - Probability support and range tests
'
' UPDATED
'   2026-07-11
'==============================================================================
'
'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Dim ActualValue         As Double          'Actual numeric value

'------------------------------------------------------------------------------
' VALIDATE ACTUAL VALUE
'------------------------------------------------------------------------------
    'Reject worksheet errors
        If IsError(Actual) Then
            RecordResult _
                TestName & _
                " -> returned error, expected a probability", _
                False

            Exit Sub
        End If

    'Reject non-numeric values
        If Not IsNumeric(Actual) Then
            RecordResult _
                TestName & " -> returned a non-numeric value", _
                False

            Exit Sub
        End If

'------------------------------------------------------------------------------
' RECORD RESULT
'------------------------------------------------------------------------------
    'Convert once and test the closed unit interval
        ActualValue = CDbl(Actual)

        If ActualValue >= 0# And ActualValue <= 1# Then
            RecordResult TestName, True
        Else
            RecordResult _
                TestName & _
                " -> got " & CStr(ActualValue) & _
                ", outside [0, 1]", _
                False
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


Private Sub AssertErrorCode( _
    ByVal TestName As String, _
    ByVal Actual As Variant, _
    ByVal ExpectedErrorCode As Long)
'
'==============================================================================
' AssertErrorCode
'------------------------------------------------------------------------------
' PURPOSE
'   Passes only when Actual is the requested worksheet CVErr code.
'
' INPUTS
'   TestName         Assertion label printed on failure.
'   Actual           Variant returned by the code under test.
'   ExpectedErrorCode Excel error code, such as xlErrNum or xlErrValue.
'
' BEHAVIOR
'   - Rejects a non-error Actual.
'   - Compares the localized string representations of the two Variant/Error
'     values because direct numeric coercion of a Variant/Error raises a type
'     mismatch in VBA.
'
' DEPENDENCIES
'   - RecordResult
'
' CALLED FROM
'   - Numerical-contract tests
'
' UPDATED
'   2026-07-11
'==============================================================================
'
'------------------------------------------------------------------------------
' VALIDATE ACTUAL VALUE
'------------------------------------------------------------------------------
    'Reject a non-error return
        If Not IsError(Actual) Then
            RecordResult _
                TestName & _
                " -> expected error, got " & CStr(Actual), _
                False

            Exit Sub
        End If

'------------------------------------------------------------------------------
' COMPARE ERROR CODES
'------------------------------------------------------------------------------
    'Compare the requested and actual worksheet error values
        If CStr(Actual) = CStr(CVErr(ExpectedErrorCode)) Then
            RecordResult TestName, True
        Else
            RecordResult _
                TestName & _
                " -> got " & CStr(Actual) & _
                ", expected " & CStr(CVErr(ExpectedErrorCode)), _
                False
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
            mFailureLog = mFailureLog & TestName & vbCrLf
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
'
'==============================================================================
' Test_CN_GammaDensity
'------------------------------------------------------------------------------
' PURPOSE
'   Verifies Gamma density values and ratio-overflow limits.
'
' BEHAVIOR
'   - Prints one section heading.
'   - Executes silent passing assertions.
'   - Records detailed output only when an assertion fails.
'
' DEPENDENCIES
'   - Production functions under test
'   - Shared assertion helpers in this module
'
' CALLED FROM
'   - RunContinuousSuite
'
' UPDATED
'   2026-07-11
'==============================================================================
'
    Debug.Print "-- Gamma density"
    AssertClose "gamma pdf(3,2.5,1.5)", K_STATS_Gamma_Density(3#, 2.5, 1.5), _
        0.19196788093578, TOL_ABS_TIGHT
    AssertClose "gamma pdf(0.5,0.5,2)", K_STATS_Gamma_Density(0.5, 0.5, 2#), _
        0.439391289467722, TOL_ABS_TIGHT
    AssertClose "gamma pdf(-1,2.5,1.5)=0", K_STATS_Gamma_Density(-1#, 2.5, 1.5), 0#, 0#
    AssertClose "gamma pdf ratio overflow tends zero", _
        K_STATS_Gamma_Density(1E+308, 2#, 9.99988867182683E-321), 0#, 0#
End Sub


Private Sub Test_CN_GammaCumulative()
'
'==============================================================================
' Test_CN_GammaCumulative
'------------------------------------------------------------------------------
' PURPOSE
'   Verifies Gamma cumulative probabilities and support limits.
'
' BEHAVIOR
'   - Prints one section heading.
'   - Executes silent passing assertions.
'   - Records detailed output only when an assertion fails.
'
' DEPENDENCIES
'   - Production functions under test
'   - Shared assertion helpers in this module
'
' CALLED FROM
'   - RunContinuousSuite
'
' UPDATED
'   2026-07-11
'==============================================================================
'
    Debug.Print "-- Gamma cumulative"
    AssertClose "gamma cdf(3,2.5,1.5)", K_STATS_Gamma_Cumulative(3#, 2.5, 1.5), _
        0.45058404864722, TOL_ABS_TIGHT
    AssertClose "gamma cdf(1e-6,0.5,2)", K_STATS_Gamma_Cumulative(0.000001, 0.5, 2#), _
        7.97884427822125E-04, TOL_ABS_TIGHT
    AssertClose "gamma cdf(0,2.5,1.5)=0", K_STATS_Gamma_Cumulative(0#, 2.5, 1.5), 0#, 0#
    AssertClose "gamma cdf(-5,2.5,1.5)=0", K_STATS_Gamma_Cumulative(-5#, 2.5, 1.5), 0#, 0#
    AssertClose "gamma cdf ratio overflow tends one", _
        K_STATS_Gamma_Cumulative(1E+308, 2#, 9.99988867182683E-321), 1#, 0#
End Sub


Private Sub Test_CN_GammaSurvival()
'
'==============================================================================
' Test_CN_GammaSurvival
'------------------------------------------------------------------------------
' PURPOSE
'   Verifies Gamma survival probabilities and complements.
'
' BEHAVIOR
'   - Prints one section heading.
'   - Executes silent passing assertions.
'   - Records detailed output only when an assertion fails.
'
' DEPENDENCIES
'   - Production functions under test
'   - Shared assertion helpers in this module
'
' CALLED FROM
'   - RunContinuousSuite
'
' UPDATED
'   2026-07-11
'==============================================================================
'
    Debug.Print "-- Gamma survival"
    AssertClose "gamma sf(3,2.5,1.5)", K_STATS_Gamma_Survival(3#, 2.5, 1.5), _
        0.54941595135278, TOL_ABS_TIGHT
    AssertClose "gamma sf(0,2.5,1.5)=1", K_STATS_Gamma_Survival(0#, 2.5, 1.5), 1#, 0#
    AssertClose "gamma sf(-5,2.5,1.5)=1", K_STATS_Gamma_Survival(-5#, 2.5, 1.5), 1#, 0#

    'CDF and survival must sum to one on the support
    AssertClose "gamma cdf+sf=1", _
        CDbl(K_STATS_Gamma_Cumulative(3#, 2.5, 1.5)) + _
        CDbl(K_STATS_Gamma_Survival(3#, 2.5, 1.5)), 1#, TOL_ABS_TIGHT
    AssertClose "gamma sf ratio overflow tends zero", _
        K_STATS_Gamma_Survival(1E+308, 2#, 9.99988867182683E-321), 0#, 0#
End Sub


Private Sub Test_CN_GammaInverse()
'
'==============================================================================
' Test_CN_GammaInverse
'------------------------------------------------------------------------------
' PURPOSE
'   Verifies Gamma inverse cumulative values.
'
' BEHAVIOR
'   - Prints one section heading.
'   - Executes silent passing assertions.
'   - Records detailed output only when an assertion fails.
'
' DEPENDENCIES
'   - Production functions under test
'   - Shared assertion helpers in this module
'
' CALLED FROM
'   - RunContinuousSuite
'
' UPDATED
'   2026-07-11
'==============================================================================
'
    Debug.Print "-- Gamma inverse"
    AssertRelClose "gamma inv(0.7,2.5,1.5)", K_STATS_Gamma_InverseCumulative(0.7, 2.5, 1.5), _
        4.54832248811618, TOL_REL_TIGHT
    AssertRelClose "gamma inv(0.05,2.5,1.5)", K_STATS_Gamma_InverseCumulative(0.05, 2.5, 1.5), _
        0.859107169546327, TOL_REL_TIGHT
End Sub


Private Sub Test_CN_GammaMoments()
'
'==============================================================================
' Test_CN_GammaMoments
'------------------------------------------------------------------------------
' PURPOSE
'   Verifies Gamma mean, variance and standard deviation.
'
' BEHAVIOR
'   - Prints one section heading.
'   - Executes silent passing assertions.
'   - Records detailed output only when an assertion fails.
'
' DEPENDENCIES
'   - Production functions under test
'   - Shared assertion helpers in this module
'
' CALLED FROM
'   - RunContinuousSuite
'
' UPDATED
'   2026-07-11
'==============================================================================
'
    Debug.Print "-- Gamma moments"
    AssertClose "gamma mean(2.5,1.5)", K_STATS_Gamma_Mean(2.5, 1.5), 3.75, TOL_ABS_TIGHT
    AssertClose "gamma var(2.5,1.5)", K_STATS_Gamma_Variance(2.5, 1.5), 5.625, TOL_ABS_TIGHT
    AssertClose "gamma std(2.5,1.5)", K_STATS_Gamma_StdDev(2.5, 1.5), _
        2.37170824512628, TOL_ABS_TIGHT
End Sub


Private Sub Test_CN_BetaDensity()
'
'==============================================================================
' Test_CN_BetaDensity
'------------------------------------------------------------------------------
' PURPOSE
'   Verifies Beta density values and support behavior.
'
' BEHAVIOR
'   - Prints one section heading.
'   - Executes silent passing assertions.
'   - Records detailed output only when an assertion fails.
'
' DEPENDENCIES
'   - Production functions under test
'   - Shared assertion helpers in this module
'
' CALLED FROM
'   - RunContinuousSuite
'
' UPDATED
'   2026-07-11
'==============================================================================
'
    Debug.Print "-- Beta density"
    AssertClose "beta pdf(0.3,2,5)", K_STATS_Beta_Density(0.3, 2#, 5#), _
        2.1609, TOL_ABS_TIGHT
    AssertRelClose "beta pdf(0.999999,2,3)", K_STATS_Beta_Density(0.999999, 2#, 3#), _
        1.1999988E-11, TOL_REL_TAIL
    AssertClose "beta pdf(-0.1,2,5)=0", K_STATS_Beta_Density(-0.1, 2#, 5#), 0#, 0#
    AssertClose "beta pdf(1.1,2,5)=0", K_STATS_Beta_Density(1.1, 2#, 5#), 0#, 0#
End Sub


Private Sub Test_CN_BetaCumulative()
'
'==============================================================================
' Test_CN_BetaCumulative
'------------------------------------------------------------------------------
' PURPOSE
'   Verifies Beta cumulative probabilities and support edges.
'
' BEHAVIOR
'   - Prints one section heading.
'   - Executes silent passing assertions.
'   - Records detailed output only when an assertion fails.
'
' DEPENDENCIES
'   - Production functions under test
'   - Shared assertion helpers in this module
'
' CALLED FROM
'   - RunContinuousSuite
'
' UPDATED
'   2026-07-11
'==============================================================================
'
    Debug.Print "-- Beta cumulative"
    AssertClose "beta cdf(0.3,2,5)", K_STATS_Beta_Cumulative(0.3, 2#, 5#), _
        0.579825, TOL_ABS_TIGHT
    AssertClose "beta cdf(0,2,5)=0", K_STATS_Beta_Cumulative(0#, 2#, 5#), 0#, 0#
    AssertClose "beta cdf(1,2,5)=1", K_STATS_Beta_Cumulative(1#, 2#, 5#), 1#, 0#
    AssertClose "beta cdf(-0.2,2,5)=0", K_STATS_Beta_Cumulative(-0.2, 2#, 5#), 0#, 0#
    AssertClose "beta cdf(1.2,2,5)=1", K_STATS_Beta_Cumulative(1.2, 2#, 5#), 1#, 0#
End Sub


Private Sub Test_CN_BetaSurvival()
'
'==============================================================================
' Test_CN_BetaSurvival
'------------------------------------------------------------------------------
' PURPOSE
'   Verifies Beta survival probabilities and complements.
'
' BEHAVIOR
'   - Prints one section heading.
'   - Executes silent passing assertions.
'   - Records detailed output only when an assertion fails.
'
' DEPENDENCIES
'   - Production functions under test
'   - Shared assertion helpers in this module
'
' CALLED FROM
'   - RunContinuousSuite
'
' UPDATED
'   2026-07-11
'==============================================================================
'
    Debug.Print "-- Beta survival"
    AssertClose "beta sf(0.3,2,5)", K_STATS_Beta_Survival(0.3, 2#, 5#), _
        0.420175, TOL_ABS_TIGHT
    AssertClose "beta sf(0,2,5)=1", K_STATS_Beta_Survival(0#, 2#, 5#), 1#, 0#
    AssertClose "beta sf(1,2,5)=0", K_STATS_Beta_Survival(1#, 2#, 5#), 0#, 0#

    'CDF and survival must sum to one on the support
    AssertClose "beta cdf+sf=1", _
        CDbl(K_STATS_Beta_Cumulative(0.3, 2#, 5#)) + _
        CDbl(K_STATS_Beta_Survival(0.3, 2#, 5#)), 1#, TOL_ABS_TIGHT
End Sub


Private Sub Test_CN_BetaInverse()
'
'==============================================================================
' Test_CN_BetaInverse
'------------------------------------------------------------------------------
' PURPOSE
'   Verifies Beta inverse cumulative values.
'
' BEHAVIOR
'   - Prints one section heading.
'   - Executes silent passing assertions.
'   - Records detailed output only when an assertion fails.
'
' DEPENDENCIES
'   - Production functions under test
'   - Shared assertion helpers in this module
'
' CALLED FROM
'   - RunContinuousSuite
'
' UPDATED
'   2026-07-11
'==============================================================================
'
    Debug.Print "-- Beta inverse"
    AssertRelClose "beta inv(0.6,2,5)", K_STATS_Beta_InverseCumulative(0.6, 2#, 5#), _
        0.309444427545314, TOL_REL_TIGHT
End Sub


Private Sub Test_CN_BetaMoments()
'
'==============================================================================
' Test_CN_BetaMoments
'------------------------------------------------------------------------------
' PURPOSE
'   Verifies Beta moments, including extreme balanced shapes.
'
' BEHAVIOR
'   - Prints one section heading.
'   - Executes silent passing assertions.
'   - Records detailed output only when an assertion fails.
'
' DEPENDENCIES
'   - Production functions under test
'   - Shared assertion helpers in this module
'
' CALLED FROM
'   - RunContinuousSuite
'
' UPDATED
'   2026-07-11
'==============================================================================
'
    Debug.Print "-- Beta moments"
    AssertClose "beta mean(2,5)", K_STATS_Beta_Mean(2#, 5#), _
        0.285714285714286, TOL_ABS_TIGHT
    AssertClose "beta var(2,5)", K_STATS_Beta_Variance(2#, 5#), _
        2.55102040816327E-02, TOL_ABS_TIGHT
    AssertClose "beta std(2,5)", K_STATS_Beta_StdDev(2#, 5#), _
        0.159719141249985, TOL_ABS_TIGHT

    'Balanced extreme shapes must retain the correct small variance.
        AssertClose _
            "beta mean balanced extreme shapes", _
            K_STATS_Beta_Mean(1E+99, 1E+99), _
            0.5, _
            TOL_ABS_TIGHT

        AssertRelClose _
            "beta variance balanced extreme shapes", _
            K_STATS_Beta_Variance(1E+99, 1E+99), _
            1.25E-100, _
            TOL_REL_LOOSE

        AssertRelClose _
            "beta stddev balanced extreme shapes", _
            K_STATS_Beta_StdDev(1E+99, 1E+99), _
            1.11803398874989E-50, _
            TOL_REL_LOOSE
End Sub


Private Sub Test_CN_ExponentialDensity()
'
'==============================================================================
' Test_CN_ExponentialDensity
'------------------------------------------------------------------------------
' PURPOSE
'   Verifies Exponential density values and product overflow.
'
' BEHAVIOR
'   - Prints one section heading.
'   - Executes silent passing assertions.
'   - Records detailed output only when an assertion fails.
'
' DEPENDENCIES
'   - Production functions under test
'   - Shared assertion helpers in this module
'
' CALLED FROM
'   - RunContinuousSuite
'
' UPDATED
'   2026-07-11
'==============================================================================
'
    Debug.Print "-- Exponential density"
    AssertClose "exp pdf(1,2)", K_STATS_Exponential_Density(1#, 2#), _
        0.270670566473225, TOL_ABS_TIGHT
    AssertClose "exp pdf(0,2)=lambda", K_STATS_Exponential_Density(0#, 2#), 2#, TOL_ABS_TIGHT
    AssertClose "exp pdf(-1,2)=0", K_STATS_Exponential_Density(-1#, 2#), 0#, 0#
    AssertClose "exp pdf product overflow tends zero", _
        K_STATS_Exponential_Density(1E+308, 1E+308), 0#, 0#
End Sub


Private Sub Test_CN_ExponentialCumulative()
'
'==============================================================================
' Test_CN_ExponentialCumulative
'------------------------------------------------------------------------------
' PURPOSE
'   Verifies Exponential cumulative probabilities and left-tail precision.
'
' BEHAVIOR
'   - Prints one section heading.
'   - Executes silent passing assertions.
'   - Records detailed output only when an assertion fails.
'
' DEPENDENCIES
'   - Production functions under test
'   - Shared assertion helpers in this module
'
' CALLED FROM
'   - RunContinuousSuite
'
' UPDATED
'   2026-07-11
'==============================================================================
'
    Debug.Print "-- Exponential cumulative"
    AssertClose "exp cdf(1,2)", K_STATS_Exponential_Cumulative(1#, 2#), _
        0.864664716763387, TOL_ABS_TIGHT
    'Left tail through PROB_Expm1: absolute tolerance would be vacuous here
    AssertRelClose "exp cdf(1e-10,1)", K_STATS_Exponential_Cumulative(0.0000000001, 1#), _
        9.9999999995E-11, TOL_REL_TAIL
    AssertClose "exp cdf(0,2)=0", K_STATS_Exponential_Cumulative(0#, 2#), 0#, 0#
    AssertClose "exp cdf(-1,2)=0", K_STATS_Exponential_Cumulative(-1#, 2#), 0#, 0#
    AssertClose "exp cdf product overflow tends one", _
        K_STATS_Exponential_Cumulative(1E+308, 1E+308), 1#, 0#
End Sub


Private Sub Test_CN_ExponentialSurvival()
'
'==============================================================================
' Test_CN_ExponentialSurvival
'------------------------------------------------------------------------------
' PURPOSE
'   Verifies Exponential survival probabilities and complements.
'
' BEHAVIOR
'   - Prints one section heading.
'   - Executes silent passing assertions.
'   - Records detailed output only when an assertion fails.
'
' DEPENDENCIES
'   - Production functions under test
'   - Shared assertion helpers in this module
'
' CALLED FROM
'   - RunContinuousSuite
'
' UPDATED
'   2026-07-11
'==============================================================================
'
    Debug.Print "-- Exponential survival"
    AssertClose "exp sf(1,2)", K_STATS_Exponential_Survival(1#, 2#), _
        0.135335283236613, TOL_ABS_TIGHT
    AssertClose "exp sf(0,2)=1", K_STATS_Exponential_Survival(0#, 2#), 1#, 0#
    AssertClose "exp sf(-1,2)=1", K_STATS_Exponential_Survival(-1#, 2#), 1#, 0#

    'CDF and survival must sum to one
    AssertClose "exp cdf+sf=1", _
        CDbl(K_STATS_Exponential_Cumulative(1#, 2#)) + _
        CDbl(K_STATS_Exponential_Survival(1#, 2#)), 1#, TOL_ABS_TIGHT
    AssertClose "exp sf product overflow tends zero", _
        K_STATS_Exponential_Survival(1E+308, 1E+308), 0#, 0#
End Sub


Private Sub Test_CN_ExponentialInverse()
'
'==============================================================================
' Test_CN_ExponentialInverse
'------------------------------------------------------------------------------
' PURPOSE
'   Verifies Exponential inverse cumulative values and small probabilities.
'
' BEHAVIOR
'   - Prints one section heading.
'   - Executes silent passing assertions.
'   - Records detailed output only when an assertion fails.
'
' DEPENDENCIES
'   - Production functions under test
'   - Shared assertion helpers in this module
'
' CALLED FROM
'   - RunContinuousSuite
'
' UPDATED
'   2026-07-11
'==============================================================================
'
    Debug.Print "-- Exponential inverse"
    AssertClose "exp inv(0.5,2)", K_STATS_Exponential_InverseCumulative(0.5, 2#), _
        0.346573590279973, TOL_ABS_TIGHT
    'Left tail through PROB_Log1p
    AssertRelClose "exp inv(1e-12,1)", K_STATS_Exponential_InverseCumulative(0.000000000001, 1#), _
        1.0000000000005E-12, TOL_REL_TAIL
End Sub


Private Sub Test_CN_WeibullDensity()
'
'==============================================================================
' Test_CN_WeibullDensity
'------------------------------------------------------------------------------
' PURPOSE
'   Verifies Weibull density values and support behavior.
'
' BEHAVIOR
'   - Prints one section heading.
'   - Executes silent passing assertions.
'   - Records detailed output only when an assertion fails.
'
' DEPENDENCIES
'   - Production functions under test
'   - Shared assertion helpers in this module
'
' CALLED FROM
'   - RunContinuousSuite
'
' UPDATED
'   2026-07-11
'==============================================================================
'
    Debug.Print "-- Weibull density"
    AssertClose "weibull pdf(1,1.5,2)", K_STATS_Weibull_Density(1#, 1.5, 2#), _
        0.372391688219422, TOL_ABS_TIGHT
    AssertClose "weibull pdf(-1,1.5,2)=0", K_STATS_Weibull_Density(-1#, 1.5, 2#), 0#, 0#
End Sub


Private Sub Test_CN_WeibullCumulative()
'
'==============================================================================
' Test_CN_WeibullCumulative
'------------------------------------------------------------------------------
' PURPOSE
'   Verifies Weibull cumulative probabilities and left-tail precision.
'
' BEHAVIOR
'   - Prints one section heading.
'   - Executes silent passing assertions.
'   - Records detailed output only when an assertion fails.
'
' DEPENDENCIES
'   - Production functions under test
'   - Shared assertion helpers in this module
'
' CALLED FROM
'   - RunContinuousSuite
'
' UPDATED
'   2026-07-11
'==============================================================================
'
    Debug.Print "-- Weibull cumulative"
    AssertClose "weibull cdf(1,1.5,2)", K_STATS_Weibull_Cumulative(1#, 1.5, 2#), _
        0.29781149867344, TOL_ABS_TIGHT
    'Left tail through PROB_Expm1
    AssertRelClose "weibull cdf(1e-10,1,1)", K_STATS_Weibull_Cumulative(0.0000000001, 1#, 1#), _
        9.9999999995E-11, TOL_REL_TAIL
    AssertClose "weibull cdf(0,1.5,2)=0", K_STATS_Weibull_Cumulative(0#, 1.5, 2#), 0#, 0#
    AssertClose "weibull cdf(-3,1.5,2)=0", K_STATS_Weibull_Cumulative(-3#, 1.5, 2#), 0#, 0#
End Sub


Private Sub Test_CN_WeibullSurvival()
'
'==============================================================================
' Test_CN_WeibullSurvival
'------------------------------------------------------------------------------
' PURPOSE
'   Verifies Weibull survival probabilities and complements.
'
' BEHAVIOR
'   - Prints one section heading.
'   - Executes silent passing assertions.
'   - Records detailed output only when an assertion fails.
'
' DEPENDENCIES
'   - Production functions under test
'   - Shared assertion helpers in this module
'
' CALLED FROM
'   - RunContinuousSuite
'
' UPDATED
'   2026-07-11
'==============================================================================
'
    Debug.Print "-- Weibull survival"
    AssertClose "weibull sf(1,1.5,2)", K_STATS_Weibull_Survival(1#, 1.5, 2#), _
        0.70218850132656, TOL_ABS_TIGHT
    AssertClose "weibull sf(0,1.5,2)=1", K_STATS_Weibull_Survival(0#, 1.5, 2#), 1#, 0#

    'CDF and survival must sum to one on the support
    AssertClose "weibull cdf+sf=1", _
        CDbl(K_STATS_Weibull_Cumulative(1#, 1.5, 2#)) + _
        CDbl(K_STATS_Weibull_Survival(1#, 1.5, 2#)), 1#, TOL_ABS_TIGHT
End Sub


Private Sub Test_CN_WeibullInverse()
'
'==============================================================================
' Test_CN_WeibullInverse
'------------------------------------------------------------------------------
' PURPOSE
'   Verifies Weibull inverse cumulative values.
'
' BEHAVIOR
'   - Prints one section heading.
'   - Executes silent passing assertions.
'   - Records detailed output only when an assertion fails.
'
' DEPENDENCIES
'   - Production functions under test
'   - Shared assertion helpers in this module
'
' CALLED FROM
'   - RunContinuousSuite
'
' UPDATED
'   2026-07-11
'==============================================================================
'
    Debug.Print "-- Weibull inverse"
    AssertClose "weibull inv(0.4,1.5,2)", K_STATS_Weibull_InverseCumulative(0.4, 1.5, 2#), _
        1.27804195727092, TOL_ABS_TIGHT
End Sub


Private Sub Test_CN_WeibullMoments()
'
'==============================================================================
' Test_CN_WeibullMoments
'------------------------------------------------------------------------------
' PURPOSE
'   Verifies stable Weibull moments across ordinary and extreme shapes.
'
' BEHAVIOR
'   - Prints one section heading.
'   - Executes silent passing assertions.
'   - Records detailed output only when an assertion fails.
'
' DEPENDENCIES
'   - Production functions under test
'   - Shared assertion helpers in this module
'
' CALLED FROM
'   - RunContinuousSuite
'
' UPDATED
'   2026-07-11
'==============================================================================
'
    Debug.Print "-- Weibull moments"
    AssertClose "weibull mean(1.5,2)", K_STATS_Weibull_Mean(1.5, 2#), _
        1.80549058590187, TOL_ABS_TIGHT
    AssertClose "weibull var(1.5,2)", K_STATS_Weibull_Variance(1.5, 2#), _
        1.50276113925573, TOL_ABS_TIGHT
    AssertClose "weibull std(1.5,2)", K_STATS_Weibull_StdDev(1.5, 2#), _
        1.22587158350935, TOL_ABS_TIGHT

    'REGRESSION: direct subtraction of two Gamma values collapsed to zero for
    'large Shape. The asymptotic branch preserves the small positive factor.
    AssertRelClose "weibull var(shape=1e8,scale=2)", _
        K_STATS_Weibull_Variance(100000000#, 2#), _
        6.57973609526982E-16, TOL_REL_TIGHT

    AssertRelClose "weibull std(shape=1e8,scale=2)", _
        K_STATS_Weibull_StdDev(100000000#, 2#), _
        2.56509962677277E-08, TOL_REL_TIGHT

    'The log-domain scale adjustment must avoid an intermediate scale^2
    'overflow where the final variance remains representable.
    AssertRelClose "weibull balanced large scale and shape", _
        K_STATS_Weibull_Variance(1E+99, 1E+99), _
        1.64493406684823, TOL_REL_LOOSE
End Sub


Private Sub Test_CN_UniformDensity()
'
'==============================================================================
' Test_CN_UniformDensity
'------------------------------------------------------------------------------
' PURPOSE
'   Verifies Uniform density values across ordinary and extreme bounds.
'
' BEHAVIOR
'   - Prints one section heading.
'   - Executes silent passing assertions.
'   - Records detailed output only when an assertion fails.
'
' DEPENDENCIES
'   - Production functions under test
'   - Shared assertion helpers in this module
'
' CALLED FROM
'   - RunContinuousSuite
'
' UPDATED
'   2026-07-11
'==============================================================================
'
    Debug.Print "-- Uniform density"
    AssertClose "uniform pdf(3,2,5)", K_STATS_Uniform_Density(3#, 2#, 5#), _
        0.333333333333333, TOL_ABS_TIGHT
    AssertClose "uniform pdf(2,2,5) edge", K_STATS_Uniform_Density(2#, 2#, 5#), _
        0.333333333333333, TOL_ABS_TIGHT
    AssertClose "uniform pdf(1,2,5)=0", K_STATS_Uniform_Density(1#, 2#, 5#), 0#, 0#
    AssertClose "uniform pdf(6,2,5)=0", K_STATS_Uniform_Density(6#, 2#, 5#), 0#, 0#

    'Opposite-sign finite bounds may have a mathematical width above Double max.
        AssertRelClose _
            "uniform pdf full finite range", _
            K_STATS_Uniform_Density(0#, -1E+308, 1E+308), _
            5E-309, _
            TOL_REL_LOOSE
End Sub


Private Sub Test_CN_UniformCumulative()
'
'==============================================================================
' Test_CN_UniformCumulative
'------------------------------------------------------------------------------
' PURPOSE
'   Verifies Uniform cumulative probabilities across finite bounds.
'
' BEHAVIOR
'   - Prints one section heading.
'   - Executes silent passing assertions.
'   - Records detailed output only when an assertion fails.
'
' DEPENDENCIES
'   - Production functions under test
'   - Shared assertion helpers in this module
'
' CALLED FROM
'   - RunContinuousSuite
'
' UPDATED
'   2026-07-11
'==============================================================================
'
    Debug.Print "-- Uniform cumulative"
    AssertClose "uniform cdf(3,2,5)", K_STATS_Uniform_Cumulative(3#, 2#, 5#), _
        0.333333333333333, TOL_ABS_TIGHT
    AssertClose "uniform cdf(1,2,5)=0", K_STATS_Uniform_Cumulative(1#, 2#, 5#), 0#, 0#
    AssertClose "uniform cdf(6,2,5)=1", K_STATS_Uniform_Cumulative(6#, 2#, 5#), 1#, 0#

    'Scaled coordinates must preserve the midpoint of an extreme support.
        AssertClose _
            "uniform cdf full finite range midpoint", _
            K_STATS_Uniform_Cumulative(0#, -1E+308, 1E+308), _
            0.5, _
            TOL_ABS_TIGHT
End Sub


Private Sub Test_CN_UniformSurvival()
'
'==============================================================================
' Test_CN_UniformSurvival
'------------------------------------------------------------------------------
' PURPOSE
'   Verifies Uniform survival probabilities and complements.
'
' BEHAVIOR
'   - Prints one section heading.
'   - Executes silent passing assertions.
'   - Records detailed output only when an assertion fails.
'
' DEPENDENCIES
'   - Production functions under test
'   - Shared assertion helpers in this module
'
' CALLED FROM
'   - RunContinuousSuite
'
' UPDATED
'   2026-07-11
'==============================================================================
'
    Debug.Print "-- Uniform survival"
    AssertClose "uniform sf(3,2,5)", K_STATS_Uniform_Survival(3#, 2#, 5#), _
        0.666666666666667, TOL_ABS_TIGHT
    AssertClose "uniform sf(1,2,5)=1", K_STATS_Uniform_Survival(1#, 2#, 5#), 1#, 0#
    AssertClose "uniform sf(6,2,5)=0", K_STATS_Uniform_Survival(6#, 2#, 5#), 0#, 0#

    'CDF and survival must sum to one on the support
    AssertClose "uniform cdf+sf=1", _
        CDbl(K_STATS_Uniform_Cumulative(3#, 2#, 5#)) + _
        CDbl(K_STATS_Uniform_Survival(3#, 2#, 5#)), 1#, 0#

    'The direct right-tail calculation must preserve the extreme-support midpoint.
        AssertClose _
            "uniform sf full finite range midpoint", _
            K_STATS_Uniform_Survival(0#, -1E+308, 1E+308), _
            0.5, _
            TOL_ABS_TIGHT
End Sub


Private Sub Test_CN_UniformInverse()
'
'==============================================================================
' Test_CN_UniformInverse
'------------------------------------------------------------------------------
' PURPOSE
'   Verifies stable Uniform inverse interpolation.
'
' BEHAVIOR
'   - Prints one section heading.
'   - Executes silent passing assertions.
'   - Records detailed output only when an assertion fails.
'
' DEPENDENCIES
'   - Production functions under test
'   - Shared assertion helpers in this module
'
' CALLED FROM
'   - RunContinuousSuite
'
' UPDATED
'   2026-07-11
'==============================================================================
'
    Debug.Print "-- Uniform inverse"
    AssertClose "uniform inv(0.25,2,5)", K_STATS_Uniform_InverseCumulative(0.25, 2#, 5#), _
        2.75, TOL_ABS_TIGHT

    'The stable convex combination must not form an overflowing support width.
        AssertRelClose _
            "uniform inv full finite range p=0.25", _
            K_STATS_Uniform_InverseCumulative( _
                0.25, _
                -1E+308, _
                1E+308), _
            -5E+307, _
            TOL_REL_TIGHT

        AssertRelClose _
            "uniform inv full finite range p=0.75", _
            K_STATS_Uniform_InverseCumulative( _
                0.75, _
                -1E+308, _
                1E+308), _
            5E+307, _
            TOL_REL_TIGHT
End Sub


Private Sub Test_CN_CrossFamilyIdentities()
'
'==============================================================================
' Test_CN_CrossFamilyIdentities
'------------------------------------------------------------------------------
' PURPOSE
'   Verifies distribution identities across independent public surfaces.
'
' BEHAVIOR
'   - Prints one section heading.
'   - Executes silent passing assertions.
'   - Records detailed output only when an assertion fails.
'
' DEPENDENCIES
'   - Production functions under test
'   - Shared assertion helpers in this module
'
' CALLED FROM
'   - RunContinuousSuite
'
' UPDATED
'   2026-07-11
'==============================================================================
'
    Dim BetaArg As Double

    Debug.Print "-- Cross-family identities (self-checks against independent oracles)"

    'Identity 1 (marshalling): Chi-square(v) is Gamma(shape v/2, scale 2)
    AssertClose "id chi2(3,5) via TFAMILY", K_STATS_ChiSquare_Cumulative(3#, 5#), _
        0.300014164121372, TOL_ABS_TIGHT
    AssertClose "id chi2(3,5)=gamma(3,2.5,2)", K_STATS_Gamma_Cumulative(3#, 2.5, 2#), _
        0.300014164121372, TOL_ABS_TIGHT

    'Identity 2 (marshalling): F(d1,d2) maps to Beta(d1/2, d2/2) at d1 x /(d1 x + d2)
    BetaArg = 5# * 2.5 / (5# * 2.5 + 10#)
    AssertClose "id F(2.5,5,10) via TFAMILY", K_STATS_F_Cumulative(2.5, 5#, 10#), _
        0.89799772335573, TOL_ABS_TIGHT
    AssertClose "id F(2.5,5,10)=beta(arg,2.5,5)", K_STATS_Beta_Cumulative(BetaArg, 2.5, 5#), _
        0.89799772335573, TOL_ABS_TIGHT

    'Identity 3 (real kernel test): Exponential(rate L) is Gamma(1, 1/L).
    'Note the RECIPROCAL: rate 2 maps to scale 0.5. This pits the incomplete-gamma
    'kernel at shape = 1 against the closed-form PROB_Expm1 CDF.
    AssertClose "id exp(1,2) closed form", K_STATS_Exponential_Cumulative(1#, 2#), _
        0.864664716763387, TOL_ABS_TIGHT
    AssertClose "id exp(1,2)=gamma(1,1,0.5) kernel", K_STATS_Gamma_Cumulative(1#, 1#, 0.5), _
        0.864664716763387, TOL_ABS_TIGHT

    'Identity 4 (real kernel test): Uniform(0,1) is Beta(1,1); both equal x
    AssertClose "id uniform(0.37,0,1)", K_STATS_Uniform_Cumulative(0.37, 0#, 1#), _
        0.37, TOL_ABS_TIGHT
    AssertClose "id beta(0.37,1,1) kernel", K_STATS_Beta_Cumulative(0.37, 1#, 1#), _
        0.37, TOL_ABS_TIGHT

    'Identity 5 (real kernel test): Chi-square(2) is Exponential(rate 1/2). The
    'gamma kernel (via chi-square) is checked against the closed-form Exponential.
    AssertClose "id chi2(2.4,2) kernel", K_STATS_ChiSquare_Cumulative(2.4, 2#), _
        0.698805788087798, TOL_ABS_TIGHT
    AssertClose "id exp(2.4,0.5) closed form", K_STATS_Exponential_Cumulative(2.4, 0.5), _
        0.698805788087798, TOL_ABS_TIGHT
End Sub


Private Sub Test_CN_RoundTrips()
'
'==============================================================================
' Test_CN_RoundTrips
'------------------------------------------------------------------------------
' PURPOSE
'   Verifies continuous-family inverse and cumulative round-trips.
'
' BEHAVIOR
'   - Prints one section heading.
'   - Executes silent passing assertions.
'   - Records detailed output only when an assertion fails.
'
' DEPENDENCIES
'   - Production functions under test
'   - Shared assertion helpers in this module
'
' CALLED FROM
'   - RunContinuousSuite
'
' UPDATED
'   2026-07-11
'==============================================================================
'
    Debug.Print "-- Continuous inverse round-trips (CDF of quantile returns the probability)"

    AssertClose "gamma roundtrip p=0.7", _
        CDbl(K_STATS_Gamma_Cumulative( _
            CDbl(K_STATS_Gamma_InverseCumulative(0.7, 2.5, 1.5)), 2.5, 1.5)), 0.7, TOL_ABS_LOOSE
    AssertClose "beta roundtrip p=0.6", _
        CDbl(K_STATS_Beta_Cumulative( _
            CDbl(K_STATS_Beta_InverseCumulative(0.6, 2#, 5#)), 2#, 5#)), 0.6, TOL_ABS_LOOSE
    AssertClose "exp roundtrip p=0.35", _
        CDbl(K_STATS_Exponential_Cumulative( _
            CDbl(K_STATS_Exponential_InverseCumulative(0.35, 2#)), 2#)), 0.35, TOL_ABS_TIGHT
    AssertClose "weibull roundtrip p=0.4", _
        CDbl(K_STATS_Weibull_Cumulative( _
            CDbl(K_STATS_Weibull_InverseCumulative(0.4, 1.5, 2#)), 1.5, 2#)), 0.4, TOL_ABS_TIGHT
    AssertClose "uniform roundtrip p=0.25", _
        CDbl(K_STATS_Uniform_Cumulative( _
            CDbl(K_STATS_Uniform_InverseCumulative(0.25, 2#, 5#)), 2#, 5#)), 0.25, TOL_ABS_TIGHT
End Sub


Private Sub Test_CN_ErrorContract()
'
'==============================================================================
' Test_CN_ErrorContract
'------------------------------------------------------------------------------
' PURPOSE
'   Verifies continuous-family domains, error codes and diagnostics.
'
' BEHAVIOR
'   - Prints one section heading.
'   - Executes silent passing assertions.
'   - Records detailed output only when an assertion fails.
'
' DEPENDENCIES
'   - Production functions under test
'   - Shared assertion helpers in this module
'
' CALLED FROM
'   - RunContinuousSuite
'
' UPDATED
'   2026-07-11
'==============================================================================
'
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
    AssertClose "gamma cdf x huge saturates", K_STATS_Gamma_Cumulative(1E+200, 2.5, 1.5), 1#, 0#

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
        0.45058404864722, TOL_ABS_TIGHT
    AssertTrue "CN status cleared on success", (Len(Diag) = 0)
    'Full-range finite rates and scales are valid; the 1E100 cap now applies
    'only to algorithmic shape parameters.
    AssertRelClose "gamma full-range scale accepted", _
        K_STATS_Gamma_Mean(1#, 1E+200), 1E+200, TOL_REL_TIGHT
    AssertRelClose "exponential full-range rate accepted", _
        K_STATS_Exponential_Density(0#, 1E+200), 1E+200, TOL_REL_TIGHT
    AssertRelClose "weibull full-range scale accepted", _
        K_STATS_Weibull_Mean(1#, 1E+200), 1E+200, TOL_REL_TIGHT
    AssertErrorCode "gamma shape at supported boundary is #NUM", _
        K_STATS_Gamma_Mean(PROB_PARAMETER_MAGNITUDE_GUARD, 1#), xlErrNum

    'Predictable arithmetic failures must be #NUM, never the unexpected #VALUE.
    AssertErrorCode "gamma density origin overflow is #NUM", _
        K_STATS_Gamma_Density(0#, 1#, 9.99988867182683E-321), xlErrNum
    'A scale of 1E+308 is still valid because the shape-two median is about
    '1.67835 and the rescaled quantile remains below the Double maximum.
    AssertRelClose "gamma inverse near Double maximum remains finite", _
        K_STATS_Gamma_InverseCumulative(0.5, 2#, 1E+308), _
        1.67834699001668E+308, _
        TOL_REL_TIGHT

    'Using the largest finite Double as the scale makes the same rescaling
    'mathematically exceed the representable range and must return #NUM.
    AssertErrorCode "gamma inverse rescale overflow is #NUM", _
        K_STATS_Gamma_InverseCumulative(0.5, 2#, PROB_DOUBLE_MAX), xlErrNum
    AssertErrorCode "exponential inverse overflow is #NUM", _
        K_STATS_Exponential_InverseCumulative(0.5, 9.99988867182683E-321), xlErrNum
    AssertErrorCode "weibull inverse overflow is #NUM", _
        K_STATS_Weibull_InverseCumulative(0.9, 1E-100, 1#), xlErrNum
    AssertErrorCode "weibull density origin overflow is #NUM", _
        K_STATS_Weibull_Density(0#, 1#, 9.99988867182683E-321), xlErrNum
    AssertErrorCode "weibull mean tiny-shape overflow is #NUM", _
        K_STATS_Weibull_Mean(9.99988867182683E-321, 1#), xlErrNum
    AssertErrorCode "weibull variance tiny-shape overflow is #NUM", _
        K_STATS_Weibull_Variance(9.99988867182683E-321, 1#), xlErrNum
    AssertErrorCode "weibull stddev tiny-shape overflow is #NUM", _
        K_STATS_Weibull_StdDev(9.99988867182683E-321, 1#), xlErrNum

    'Both Beta arguments are algorithmic shape parameters and must respect the
    'supported-magnitude contract.
        AssertErrorCode _
            "beta second shape at supported boundary is #NUM", _
            K_STATS_Beta_Cumulative( _
                0.5, _
                1#, _
                PROB_PARAMETER_MAGNITUDE_GUARD), _
            xlErrNum

    'Full finite Uniform bounds are valid even when their width is not directly
    'representable as a Double.
        AssertClose _
            "uniform full finite bounds accepted", _
            K_STATS_Uniform_Cumulative( _
                0#, _
                -1E+308, _
                1E+308), _
            0.5, _
            TOL_ABS_TIGHT
End Sub


Private Sub Test_CN_SupportEdges()
'
'==============================================================================
' Test_CN_SupportEdges
'------------------------------------------------------------------------------
' PURPOSE
'   Verifies continuous-family finite support-edge behavior.
'
' BEHAVIOR
'   - Prints one section heading.
'   - Executes silent passing assertions.
'   - Records detailed output only when an assertion fails.
'
' DEPENDENCIES
'   - Production functions under test
'   - Shared assertion helpers in this module
'
' CALLED FROM
'   - RunContinuousSuite
'
' UPDATED
'   2026-07-11
'==============================================================================
'
    Debug.Print "-- Continuous support edges (finite boundary densities)"

    'Gamma origin: 1/Scale at shape 1, zero above, both finite (not poles)
    AssertClose "gamma pdf(0,1,2)=1/2", K_STATS_Gamma_Density(0#, 1#, 2#), 0.5, TOL_ABS_TIGHT
    AssertClose "gamma pdf(0,2,1.5)=0", K_STATS_Gamma_Density(0#, 2#, 1.5), 0#, 0#

    'Beta endpoints: Beta at 0 when alpha=1, Alpha at 1 when beta=1
    AssertClose "beta pdf(0,1,3)=3", K_STATS_Beta_Density(0#, 1#, 3#), 3#, TOL_ABS_TIGHT
    AssertClose "beta pdf(1,2,1)=2", K_STATS_Beta_Density(1#, 2#, 1#), 2#, TOL_ABS_TIGHT
    AssertClose "beta pdf(0,2,5)=0", K_STATS_Beta_Density(0#, 2#, 5#), 0#, 0#
    AssertClose "beta pdf(1,2,5)=0", K_STATS_Beta_Density(1#, 2#, 5#), 0#, 0#

    'Weibull origin: 1/Scale at shape 1, zero above
    AssertClose "weibull pdf(0,1,2)=1/2", K_STATS_Weibull_Density(0#, 1#, 2#), 0.5, TOL_ABS_TIGHT
    AssertClose "weibull pdf(0,2,1.5)=0", K_STATS_Weibull_Density(0#, 2#, 1.5), 0#, 0#

    'All CDFs live in the unit interval
    AssertInUnitInterval "gamma cdf in [0,1]", K_STATS_Gamma_Cumulative(3#, 2.5, 1.5)
    AssertInUnitInterval "beta cdf in [0,1]", K_STATS_Beta_Cumulative(0.3, 2#, 5#)
    AssertInUnitInterval "exp cdf in [0,1]", K_STATS_Exponential_Cumulative(1#, 2#)
    AssertInUnitInterval "weibull cdf in [0,1]", K_STATS_Weibull_Cumulative(1#, 1.5, 2#)
    AssertInUnitInterval "uniform cdf in [0,1]", K_STATS_Uniform_Cumulative(3#, 2#, 5#)
End Sub


Private Sub Test_DS_BinomialPMF()
'
'==============================================================================
' Test_DS_BinomialPMF
'------------------------------------------------------------------------------
' PURPOSE
'   Verifies Binomial mass values, count truncation and degenerate p.
'
' BEHAVIOR
'   - Prints one section heading.
'   - Executes silent passing assertions.
'   - Records detailed output only when an assertion fails.
'
' CALLED FROM
'   - RunDiscreteSuite
'
' UPDATED
'   2026-07-19
'==============================================================================
'
    Debug.Print "-- Binomial PMF"
    AssertClose "binom pmf(7,20,.35)", K_STATS_Binomial_PMF(7#, 20#, 0.35), _
        0.184401186383931, TOL_ABS_TIGHT
    AssertClose "binom pmf(0,20,.35)", K_STATS_Binomial_PMF(0#, 20#, 0.35), _
        1.8124545836335E-04, TOL_ABS_TIGHT
    AssertRelClose "binom pmf(20,20,.35)", K_STATS_Binomial_PMF(20#, 20#, 0.35), _
        7.60958350158805E-10, TOL_REL_TIGHT
    AssertClose "binom pmf truncates k=7.9", K_STATS_Binomial_PMF(7.9, 20#, 0.35), _
        0.184401186383931, TOL_ABS_TIGHT
    AssertClose "binom pmf p=0 at k=0", K_STATS_Binomial_PMF(0#, 20#, 0#), _
        1#, 0#
    AssertClose "binom pmf p=1 at k=n", K_STATS_Binomial_PMF(20#, 20#, 1#), _
        1#, 0#
End Sub


Private Sub Test_DS_BinomialCumulative()
'
'==============================================================================
' Test_DS_BinomialCumulative
'------------------------------------------------------------------------------
' PURPOSE
'   Verifies Binomial left-tail probabilities via the beta identity.
'
' BEHAVIOR
'   - Prints one section heading.
'   - Executes silent passing assertions.
'   - Records detailed output only when an assertion fails.
'
' CALLED FROM
'   - RunDiscreteSuite
'
' UPDATED
'   2026-07-19
'==============================================================================
'
    Debug.Print "-- Binomial cumulative"
    AssertClose "binom cdf(7,20,.35)", K_STATS_Binomial_Cumulative(7#, 20#, 0.35), _
        0.601026604603164, TOL_ABS_TIGHT
    AssertClose "binom cdf(480,1000,.5)", K_STATS_Binomial_Cumulative(480#, 1000#, 0.5), _
        0.10872414660207, TOL_ABS_TIGHT
    AssertClose "binom cdf(50,10000,.004)", K_STATS_Binomial_Cumulative(50#, 10000#, 0.004), _
        0.947726326387241, TOL_ABS_TIGHT
    AssertClose "binom cdf(20,20,.35)=1", K_STATS_Binomial_Cumulative(20#, 20#, 0.35), _
        1#, 0#
End Sub


Private Sub Test_DS_BinomialSurvival()
'
'==============================================================================
' Test_DS_BinomialSurvival
'------------------------------------------------------------------------------
' PURPOSE
'   Verifies Binomial right-tail probabilities computed directly.
'
' BEHAVIOR
'   - Prints one section heading.
'   - Executes silent passing assertions.
'   - Records detailed output only when an assertion fails.
'
' CALLED FROM
'   - RunDiscreteSuite
'
' UPDATED
'   2026-07-19
'==============================================================================
'
    Debug.Print "-- Binomial survival"
    AssertClose "binom sf(7,20,.35)", K_STATS_Binomial_Survival(7#, 20#, 0.35), _
        0.398973395396836, TOL_ABS_TIGHT
    AssertClose "binom sf(480,1000,.5)", K_STATS_Binomial_Survival(480#, 1000#, 0.5), _
        0.89127585339793, TOL_ABS_TIGHT
    AssertRelClose "binom sf(150,200,.5) tail", K_STATS_Binomial_Survival(150#, 200#, 0.5), _
        1.37214281800074E-13, TOL_REL_TAIL
    AssertClose "binom sf(20,20,.35)=0", K_STATS_Binomial_Survival(20#, 20#, 0.35), _
        0#, 0#
End Sub


Private Sub Test_DS_BinomialInverse()
'
'==============================================================================
' Test_DS_BinomialInverse
'------------------------------------------------------------------------------
' PURPOSE
'   Verifies the Binomial quantile as the least k with CDF at least p.
'
' BEHAVIOR
'   - Prints one section heading.
'   - Executes silent passing assertions.
'   - Records detailed output only when an assertion fails.
'
' CALLED FROM
'   - RunDiscreteSuite
'
' UPDATED
'   2026-07-19
'==============================================================================
'
    Debug.Print "-- Binomial inverse"
    AssertClose "binom inv(.601,20,.35)=7", K_STATS_Binomial_InverseCumulative(0.601, 20#, 0.35), _
        7#, 0#
    AssertClose "binom inv(.5,1000,.5)=500", K_STATS_Binomial_InverseCumulative(0.5, 1000#, 0.5), _
        500#, 0#
    AssertClose "binom inv(.975,20,.35)=11", K_STATS_Binomial_InverseCumulative(0.975, 20#, 0.35), _
        11#, 0#
    AssertClose "binom inv(.05,20,.35)=4", K_STATS_Binomial_InverseCumulative(0.05, 20#, 0.35), _
        4#, 0#
End Sub


Private Sub Test_DS_BinomialMoments()
'
'==============================================================================
' Test_DS_BinomialMoments
'------------------------------------------------------------------------------
' PURPOSE
'   Verifies Binomial mean, variance and standard deviation.
'
' BEHAVIOR
'   - Prints one section heading.
'   - Executes silent passing assertions.
'   - Records detailed output only when an assertion fails.
'
' CALLED FROM
'   - RunDiscreteSuite
'
' UPDATED
'   2026-07-19
'==============================================================================
'
    Debug.Print "-- Binomial moments"
    AssertClose "binom mean(20,.35)", K_STATS_Binomial_Mean(20#, 0.35), _
        7#, TOL_ABS_TIGHT
    AssertClose "binom var(20,.35)", K_STATS_Binomial_Variance(20#, 0.35), _
        4.55, TOL_ABS_TIGHT
    AssertClose "binom std(20,.35)", K_STATS_Binomial_StdDev(20#, 0.35), _
        2.13307290077015, TOL_ABS_TIGHT
End Sub


Private Sub Test_DS_PoissonPMF()
'
'==============================================================================
' Test_DS_PoissonPMF
'------------------------------------------------------------------------------
' PURPOSE
'   Verifies Poisson mass values including a deep-tail case Excel returns as zero.
'
' BEHAVIOR
'   - Prints one section heading.
'   - Executes silent passing assertions.
'   - Records detailed output only when an assertion fails.
'
' CALLED FROM
'   - RunDiscreteSuite
'
' UPDATED
'   2026-07-19
'==============================================================================
'
    Debug.Print "-- Poisson PMF"
    AssertClose "pois pmf(7,3)", K_STATS_Poisson_PMF(7#, 3#), _
        2.16040314524838E-02, TOL_ABS_TIGHT
    AssertClose "pois pmf(160,150)", K_STATS_Poisson_PMF(160#, 150#), _
        2.27495558810425E-02, TOL_ABS_TIGHT
    AssertRelClose "pois pmf(2,50) tail", K_STATS_Poisson_PMF(2#, 50#), _
        2.4109373099549E-19, TOL_REL_TAIL
    AssertRelClose "pois pmf(0,700) deep tail", K_STATS_Poisson_PMF(0#, 700#), _
        9.85967654375977E-305, TOL_REL_TAIL
    AssertClose "pois pmf mean=0 at k=0", K_STATS_Poisson_PMF(0#, 0#), _
        1#, 0#
End Sub


Private Sub Test_DS_PoissonCumulative()
'
'==============================================================================
' Test_DS_PoissonCumulative
'------------------------------------------------------------------------------
' PURPOSE
'   Verifies Poisson left-tail probabilities via the gamma identity.
'
' BEHAVIOR
'   - Prints one section heading.
'   - Executes silent passing assertions.
'   - Records detailed output only when an assertion fails.
'
' CALLED FROM
'   - RunDiscreteSuite
'
' UPDATED
'   2026-07-19
'==============================================================================
'
    Debug.Print "-- Poisson cumulative"
    AssertClose "pois cdf(7,3)", K_STATS_Poisson_Cumulative(7#, 3#), _
        0.988095496143643, TOL_ABS_TIGHT
    AssertClose "pois cdf(160,150)", K_STATS_Poisson_Cumulative(160#, 150#), _
        0.805398685507147, TOL_ABS_TIGHT
    AssertRelClose "pois cdf(2,50) tail", K_STATS_Poisson_Cumulative(2#, 50#), _
        2.50930355220106E-19, TOL_REL_TAIL
    AssertClose "pois cdf mean=0=1", K_STATS_Poisson_Cumulative(0#, 0#), _
        1#, 0#
End Sub


Private Sub Test_DS_PoissonSurvival()
'
'==============================================================================
' Test_DS_PoissonSurvival
'------------------------------------------------------------------------------
' PURPOSE
'   Verifies Poisson right-tail probabilities computed directly.
'
' BEHAVIOR
'   - Prints one section heading.
'   - Executes silent passing assertions.
'   - Records detailed output only when an assertion fails.
'
' CALLED FROM
'   - RunDiscreteSuite
'
' UPDATED
'   2026-07-19
'==============================================================================
'
    Debug.Print "-- Poisson survival"
    AssertClose "pois sf(7,3)", K_STATS_Poisson_Survival(7#, 3#), _
        1.19045038563574E-02, TOL_ABS_TIGHT
    AssertClose "pois sf(160,150)", K_STATS_Poisson_Survival(160#, 150#), _
        0.194601314492853, TOL_ABS_TIGHT
    AssertClose "pois sf mean=0=0", K_STATS_Poisson_Survival(0#, 0#), _
        0#, 0#
End Sub


Private Sub Test_DS_PoissonInverse()
'
'==============================================================================
' Test_DS_PoissonInverse
'------------------------------------------------------------------------------
' PURPOSE
'   Verifies the Poisson quantile via exponential-search bracketing.
'
' BEHAVIOR
'   - Prints one section heading.
'   - Executes silent passing assertions.
'   - Records detailed output only when an assertion fails.
'
' CALLED FROM
'   - RunDiscreteSuite
'
' UPDATED
'   2026-07-19
'==============================================================================
'
    Debug.Print "-- Poisson inverse"
    AssertClose "pois inv(.5,3)=3", K_STATS_Poisson_InverseCumulative(0.5, 3#), _
        3#, 0#
    AssertClose "pois inv(.975,3)=7", K_STATS_Poisson_InverseCumulative(0.975, 3#), _
        7#, 0#
    AssertClose "pois inv(.999,3)=10", K_STATS_Poisson_InverseCumulative(0.999, 3#), _
        10#, 0#
    AssertClose "pois inv(.5,150)=150", K_STATS_Poisson_InverseCumulative(0.5, 150#), _
        150#, 0#
End Sub


Private Sub Test_DS_PoissonMoments()
'
'==============================================================================
' Test_DS_PoissonMoments
'------------------------------------------------------------------------------
' PURPOSE
'   Verifies Poisson mean, variance and standard deviation.
'
' BEHAVIOR
'   - Prints one section heading.
'   - Executes silent passing assertions.
'   - Records detailed output only when an assertion fails.
'
' CALLED FROM
'   - RunDiscreteSuite
'
' UPDATED
'   2026-07-19
'==============================================================================
'
    Debug.Print "-- Poisson moments"
    AssertClose "pois mean(3)", K_STATS_Poisson_Mean(3#), _
        3#, 0#
    AssertClose "pois var(3)", K_STATS_Poisson_Variance(3#), _
        3#, 0#
    AssertClose "pois std(3)", K_STATS_Poisson_StdDev(3#), _
        1.73205080756888, TOL_ABS_TIGHT
End Sub


Private Sub Test_DS_GeometricPMF()
'
'==============================================================================
' Test_DS_GeometricPMF
'------------------------------------------------------------------------------
' PURPOSE
'   Verifies Geometric mass values under the failures-before-success convention.
'
' BEHAVIOR
'   - Prints one section heading.
'   - Executes silent passing assertions.
'   - Records detailed output only when an assertion fails.
'
' CALLED FROM
'   - RunDiscreteSuite
'
' UPDATED
'   2026-07-19
'==============================================================================
'
    Debug.Print "-- Geometric PMF"
    AssertClose "geo pmf(0,.3)", K_STATS_Geometric_PMF(0#, 0.3), _
        0.3, TOL_ABS_TIGHT
    AssertClose "geo pmf(5,.3)", K_STATS_Geometric_PMF(5#, 0.3), _
        0.050421, TOL_ABS_TIGHT
    AssertRelClose "geo pmf(100,.3) tail", K_STATS_Geometric_PMF(100#, 0.3), _
        9.70342952887429E-17, TOL_REL_TAIL
    AssertClose "geo pmf(50,1e-6)", K_STATS_Geometric_PMF(50#, 0.000001), _
        9.9995000122498E-07, TOL_ABS_TIGHT
    AssertClose "geo pmf p=1 at k=0", K_STATS_Geometric_PMF(0#, 1#), _
        1#, 0#
End Sub


Private Sub Test_DS_GeometricCumulative()
'
'==============================================================================
' Test_DS_GeometricCumulative
'------------------------------------------------------------------------------
' PURPOSE
'   Verifies Geometric left-tail probabilities via Expm1.
'
' BEHAVIOR
'   - Prints one section heading.
'   - Executes silent passing assertions.
'   - Records detailed output only when an assertion fails.
'
' CALLED FROM
'   - RunDiscreteSuite
'
' UPDATED
'   2026-07-19
'==============================================================================
'
    Debug.Print "-- Geometric cumulative"
    AssertClose "geo cdf(0,.3)", K_STATS_Geometric_Cumulative(0#, 0.3), _
        0.3, TOL_ABS_TIGHT
    AssertClose "geo cdf(5,.3)", K_STATS_Geometric_Cumulative(5#, 0.3), _
        0.882351, TOL_ABS_TIGHT
    AssertRelClose "geo cdf(50,1e-6) small", K_STATS_Geometric_Cumulative(50#, 0.000001), _
        5.09987250208248E-05, TOL_REL_TAIL
End Sub


Private Sub Test_DS_GeometricSurvival()
'
'==============================================================================
' Test_DS_GeometricSurvival
'------------------------------------------------------------------------------
' PURPOSE
'   Verifies Geometric right-tail probabilities computed directly.
'
' BEHAVIOR
'   - Prints one section heading.
'   - Executes silent passing assertions.
'   - Records detailed output only when an assertion fails.
'
' CALLED FROM
'   - RunDiscreteSuite
'
' UPDATED
'   2026-07-19
'==============================================================================
'
    Debug.Print "-- Geometric survival"
    AssertClose "geo sf(0,.3)", K_STATS_Geometric_Survival(0#, 0.3), _
        0.7, TOL_ABS_TIGHT
    AssertClose "geo sf(5,.3)", K_STATS_Geometric_Survival(5#, 0.3), _
        0.117649, TOL_ABS_TIGHT
    AssertClose "geo sf(50,1e-6)", K_STATS_Geometric_Survival(50#, 0.000001), _
        0.999949001274979, TOL_ABS_TIGHT
End Sub


Private Sub Test_DS_GeometricInverse()
'
'==============================================================================
' Test_DS_GeometricInverse
'------------------------------------------------------------------------------
' PURPOSE
'   Verifies the Geometric quantile via closed-form seed and correction.
'
' BEHAVIOR
'   - Prints one section heading.
'   - Executes silent passing assertions.
'   - Records detailed output only when an assertion fails.
'
' CALLED FROM
'   - RunDiscreteSuite
'
' UPDATED
'   2026-07-19
'==============================================================================
'
    Debug.Print "-- Geometric inverse"
    AssertClose "geo inv(.5,.3)=1", K_STATS_Geometric_InverseCumulative(0.5, 0.3), _
        1#, 0#
    AssertClose "geo inv(.9,.3)=6", K_STATS_Geometric_InverseCumulative(0.9, 0.3), _
        6#, 0#
    AssertClose "geo inv(.999,.3)=19", K_STATS_Geometric_InverseCumulative(0.999, 0.3), _
        19#, 0#
    AssertClose "geo inv(.5,1e-6)=693146", K_STATS_Geometric_InverseCumulative(0.5, 0.000001), _
        693146#, 0#
End Sub


Private Sub Test_DS_GeometricMoments()
'
'==============================================================================
' Test_DS_GeometricMoments
'------------------------------------------------------------------------------
' PURPOSE
'   Verifies Geometric mean, variance and standard deviation.
'
' BEHAVIOR
'   - Prints one section heading.
'   - Executes silent passing assertions.
'   - Records detailed output only when an assertion fails.
'
' CALLED FROM
'   - RunDiscreteSuite
'
' UPDATED
'   2026-07-19
'==============================================================================
'
    Debug.Print "-- Geometric moments"
    AssertClose "geo mean(.3)", K_STATS_Geometric_Mean(0.3), _
        2.33333333333333, TOL_ABS_TIGHT
    AssertClose "geo var(.3)", K_STATS_Geometric_Variance(0.3), _
        7.77777777777778, TOL_ABS_TIGHT
    AssertClose "geo std(.3)", K_STATS_Geometric_StdDev(0.3), _
        2.78886675511359, TOL_ABS_TIGHT
End Sub


Private Sub Test_DS_ErrorContract()
'
'==============================================================================
' Test_DS_ErrorContract
'------------------------------------------------------------------------------
' PURPOSE
'   Verifies discrete-family domains, error codes and diagnostics.
'
' BEHAVIOR
'   - Prints one section heading.
'   - Executes silent passing assertions.
'   - Records detailed output only when an assertion fails.
'
' CALLED FROM
'   - RunDiscreteSuite
'
' UPDATED
'   2026-07-19
'==============================================================================
'
    Dim Diag As String

    Debug.Print "-- Discrete error contract (must return CVErr, never a sentinel)"

    'Invalid counts and probabilities
    AssertIsError "binom pmf trials<0", K_STATS_Binomial_PMF(1#, -5#, 0.5)
    AssertIsError "binom pmf k>n", K_STATS_Binomial_PMF(6#, 5#, 0.5)
    AssertIsError "binom pmf p>1", K_STATS_Binomial_PMF(1#, 5#, 1.5)
    AssertIsError "binom pmf p<0", K_STATS_Binomial_PMF(1#, 5#, -0.1)
    AssertIsError "binom cdf k<0", K_STATS_Binomial_Cumulative(-1#, 5#, 0.5)
    AssertIsError "pois pmf mean<0", K_STATS_Poisson_PMF(1#, -2#)
    AssertIsError "pois pmf k<0", K_STATS_Poisson_PMF(-1#, 3#)
    AssertIsError "geo pmf p=0", K_STATS_Geometric_PMF(1#, 0#)
    AssertIsError "geo pmf p>1", K_STATS_Geometric_PMF(1#, 1.5)

    'Inverse probabilities outside the open unit interval
    AssertIsError "binom inv p=0", K_STATS_Binomial_InverseCumulative(0#, 20#, 0.35)
    AssertIsError "binom inv p=1", K_STATS_Binomial_InverseCumulative(1#, 20#, 0.35)
    AssertIsError "pois inv p=1", K_STATS_Poisson_InverseCumulative(1#, 3#)
    AssertIsError "geo inv p=0", K_STATS_Geometric_InverseCumulative(0#, 0.3)
    AssertIsError "geo inv p>1", K_STATS_Geometric_InverseCumulative(1.5, 0.3)

    'Status must be populated on failure and cleared on success
    Diag = "stale"
    AssertIsError "binom pmf p>1 with status", K_STATS_Binomial_PMF(1#, 5#, 1.5, Diag)
    AssertTrue "DS status populated on failure", (Len(Diag) > 0 And Diag <> "stale")
    Diag = "stale"
    AssertClose "binom pmf ok with status", K_STATS_Binomial_PMF(7#, 20#, 0.35, Diag), _
        0.184401186383931, TOL_ABS_TIGHT
    AssertTrue "DS status cleared on success", (Len(Diag) = 0)
End Sub


Private Sub Test_DS_SupportEdges()
'
'==============================================================================
' Test_DS_SupportEdges
'------------------------------------------------------------------------------
' PURPOSE
'   Verifies degenerate parameters and support boundaries across the family.
'
' BEHAVIOR
'   - Prints one section heading.
'   - Executes silent passing assertions.
'   - Records detailed output only when an assertion fails.
'
' CALLED FROM
'   - RunDiscreteSuite
'
' UPDATED
'   2026-07-19
'==============================================================================
'
    Debug.Print "-- Discrete support edges"
    AssertClose "binom p=0 cdf=1", K_STATS_Binomial_Cumulative(0#, 20#, 0#), _
        1#, 0#
    AssertClose "binom p=1 sf(19)=1", K_STATS_Binomial_Survival(19#, 20#, 1#), _
        1#, 0#
    AssertClose "pois mean=0 pmf(1)=0", K_STATS_Poisson_PMF(1#, 0#), _
        0#, 0#
    AssertClose "geo p=1 cdf(0)=1", K_STATS_Geometric_Cumulative(0#, 1#), _
        1#, 0#
    AssertClose "geo p=1 sf(0)=0", K_STATS_Geometric_Survival(0#, 1#), _
        0#, 0#
    AssertClose "geo p=1 inv=0", K_STATS_Geometric_InverseCumulative(0.5, 1#), _
        0#, 0#
End Sub


