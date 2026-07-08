Attribute VB_Name = "M_STATS_PROBDIST_TEST"
Option Explicit

'==============================================================================
' M_STATS_PROBDIST_NORMALFAMILY_TESTS
'------------------------------------------------------------------------------
' PURPOSE
'   Self-checking test harness for M_STATS_PROBDIST_NORMALFAMILY. Verifies known
'   values, inverse round-trips, symmetry, lognormal moments, the arithmetic-to-
'   log parameter round-trip, interval probabilities, and the error/overflow
'   contract.
'
' HOW TO RUN
'   From the VBA IDE Immediate window (Ctrl+G):
'       Test_STATS_PROBDIST_RunAll
'   Results and a PASS/FAIL summary are printed with Debug.Print. Only failures
'   print a detail line; passing tests are silent, so a clean run shows just the
'   section headers and the final summary.
'
' DEPENDENCIES
'   - M_STATS_PROBDIST_NORMALFAMILY (module under test)
'
' NOTES
'   - Public distribution functions return Variant and may return CVErr, so the
'     assertion helpers accept Variant and route through IsError.
'   - Reference constants are standard textbook values quoted to ~15 figures.
'   - TOL_TIGHT is used where the library targets machine precision; TOL_LOOSE
'     is used for the raw (unrefined) fast inverse and for moment round-trips.
'
' UPDATED
'   2026-07-08
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

Private Const TOL_TIGHT     As Double = 0.0000000001      '1E-10, machine-precision paths
Private Const TOL_LOOSE     As Double = 0.000001          '1E-6, fast inverse / round-trips


Public Sub Test_STATS_PROBDIST_NormalFamily_RunAll()
'
'==============================================================================
' Test_STATS_PROBDIST_RunAll
'------------------------------------------------------------------------------
' PURPOSE
'   Runs the full test suite and prints a PASS/FAIL summary.
'==============================================================================
'
'------------------------------------------------------------------------------
' INITIALIZE
'------------------------------------------------------------------------------
    'Reset counters
        mTestCount = 0
        mPassCount = 0
        mFailCount = 0
    'Header
        Debug.Print String(70, "=")
        Debug.Print "M_STATS_PROBDIST_NORMALFAMILY - test run " & Format(Now, "yyyy-mm-dd hh:nn:ss")
        Debug.Print String(70, "=")
'------------------------------------------------------------------------------
' RUN SECTIONS
'------------------------------------------------------------------------------
    Test_StandardNormalDensity
    Test_StandardNormalCumulative
    Test_StandardNormalInverse
    Test_InverseRoundTrips
    Test_Symmetry
    Test_GeneralNormal
    Test_ZScore
    Test_IntervalProbability
    Test_FastInverse
    Test_LognormalCore
    Test_LognormalMoments
    Test_ParameterRoundTrip
    Test_ErrorContract
    Test_OverflowContract
'------------------------------------------------------------------------------
' SUMMARY
'------------------------------------------------------------------------------
    Debug.Print String(70, "-")
    Debug.Print "TOTAL  " & mTestCount & _
                "   PASS " & mPassCount & _
                "   FAIL " & mFailCount
    If mFailCount = 0 Then
        Debug.Print "RESULT: ALL TESTS PASSED"
    Else
        Debug.Print "RESULT: " & mFailCount & " TEST(S) FAILED"
    End If
    Debug.Print String(70, "=")
End Sub


'==============================================================================
' TEST SECTIONS
'==============================================================================

Private Sub Test_StandardNormalDensity()
    Debug.Print "-- Standard normal density"
    AssertClose "phi(0)", K_STATS_NormalStandard_Density(0#), 0.398942280401433, TOL_TIGHT
    AssertClose "phi(1)", K_STATS_NormalStandard_Density(1#), 0.241970724519143, TOL_TIGHT
    AssertClose "phi(-1)", K_STATS_NormalStandard_Density(-1#), 0.241970724519143, TOL_TIGHT
    AssertClose "phi(2)", K_STATS_NormalStandard_Density(2#), 0.053990966513188, TOL_TIGHT
End Sub


Private Sub Test_StandardNormalCumulative()
    Debug.Print "-- Standard normal cumulative"
    AssertClose "Phi(0)", K_STATS_NormalStandard_Cumulative(0#), 0.5, TOL_TIGHT
    AssertClose "Phi(1)", K_STATS_NormalStandard_Cumulative(1#), 0.841344746068543, TOL_TIGHT
    AssertClose "Phi(2)", K_STATS_NormalStandard_Cumulative(2#), 0.977249868051821, TOL_TIGHT
    AssertClose "Phi(-1)", K_STATS_NormalStandard_Cumulative(-1#), 0.158655253931457, TOL_TIGHT
    AssertClose "Phi(1.959963984540054)", _
        K_STATS_NormalStandard_Cumulative(1.95996398454005), 0.975, TOL_TIGHT
End Sub


Private Sub Test_StandardNormalInverse()
    Debug.Print "-- Standard normal inverse"
    AssertClose "InvPhi(0.5)", K_STATS_NormalStandard_InverseCumulative(0.5), 0#, TOL_TIGHT
    AssertClose "InvPhi(0.975)", _
        K_STATS_NormalStandard_InverseCumulative(0.975), 1.95996398454005, TOL_TIGHT
    AssertClose "InvPhi(0.95)", _
        K_STATS_NormalStandard_InverseCumulative(0.95), 1.64485362695147, TOL_TIGHT
    AssertClose "InvPhi(0.025)", _
        K_STATS_NormalStandard_InverseCumulative(0.025), -1.95996398454005, TOL_TIGHT
End Sub


Private Sub Test_InverseRoundTrips()
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


Private Sub Test_Symmetry()
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
End Sub


Private Sub Test_GeneralNormal()
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


Private Sub Test_ZScore()
    Debug.Print "-- Z-score"
    AssertClose "z(10,4,2)", K_STATS_Normal_ZScore(10#, 4#, 2#), 3#, TOL_TIGHT
    AssertClose "z(4,4,2)", K_STATS_Normal_ZScore(4#, 4#, 2#), 0#, TOL_TIGHT
    AssertClose "z(1,4,2)", K_STATS_Normal_ZScore(1#, 4#, 2#), -1.5, TOL_TIGHT
End Sub


Private Sub Test_IntervalProbability()
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


Private Sub Test_FastInverse()
    Dim R As Double
    Debug.Print "-- Fast inverse (raw Acklam, ~1E-9)"
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


Private Sub Test_LognormalCore()
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


Private Sub Test_LognormalMoments()
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


Private Sub Test_ParameterRoundTrip()
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


Private Sub Test_ErrorContract()
    Debug.Print "-- Error contract (invalid domains must return CVErr)"
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
    'Parameter conversion rejects StdDev = 0 (new behavior) and non-positive Mean
    AssertIsError "param StdDev=0", K_STATS_Lognormal_ParametersFromMeanStdDev(2#, 0#)
    AssertIsError "param Mean=0", K_STATS_Lognormal_ParametersFromMeanStdDev(0#, 1#)
    'Reversed interval bounds
    AssertIsError "std interval reversed", _
        K_STATS_NormalStandard_IntervalProbability(1#, -1#)
    AssertIsError "gen interval reversed", _
        K_STATS_Normal_IntervalProbability(5#, 1#, 0#, 1#)
End Sub


Private Sub Test_OverflowContract()
    Debug.Print "-- Overflow contract (must return CVErr(xlErrNum), not a sentinel)"
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
'   Passes when Actual is a number within Tolerance of Expected. A CVErr Actual
'   is treated as a failure.
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


