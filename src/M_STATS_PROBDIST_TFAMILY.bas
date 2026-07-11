Attribute VB_Name = "M_STATS_PROBDIST_TFAMILY"
Option Explicit

'==============================================================================
' M_STATS_PROBDIST_TFAMILY
'------------------------------------------------------------------------------
' PURPOSE
'   Provides the Student t, chi-square and F distributions - density, left-tail
'   cumulative, right-tail survival and inverse cumulative - for worksheet
'   formulas, VBA models and simulation engines.
'
' WHY THIS EXISTS
'   These three distributions are one family, not three: every one of them is a
'   reduction of the regularized incomplete beta or incomplete gamma function, and
'   they are the distributions of the classical test statistics. A t-statistic, a
'   likelihood-ratio statistic and a variance ratio are the three things a
'   regression or ANOVA actually produces, so their tails are what a model needs.
'
'   Excel covers most of this surface (see WORKSHEET EQUIVALENTS below) but the
'   worksheet functions are awkward to call from VBA: Application.WorksheetFunction
'   raises a trappable runtime error on bad input instead of returning CVErr, and
'   carries per-call marshalling overhead. This module gives one VBA-native,
'   consistently-validated surface that is equally callable from a cell and from a
'   tight numerical loop.
'
'   Above all, the survival functions exist. Without them a right-tail p-value
'   below about 1E-16 cannot be expressed at all: 1 - CDF is exactly 1 there. The
'   Student t survival at x = 20 with 30 degrees of freedom is 3.37E-19; computed
'   as 1 - CDF it is zero.
'
' PUBLIC FUNCTIONS
'   Student t:
'     - K_STATS_StudentT_Density
'     - K_STATS_StudentT_Cumulative
'     - K_STATS_StudentT_Survival
'     - K_STATS_StudentT_InverseCumulative
'
'   Chi-square:
'     - K_STATS_ChiSquare_Density
'     - K_STATS_ChiSquare_Cumulative
'     - K_STATS_ChiSquare_Survival
'     - K_STATS_ChiSquare_InverseCumulative
'
'   F:
'     - K_STATS_F_Density
'     - K_STATS_F_Cumulative
'     - K_STATS_F_Survival
'     - K_STATS_F_InverseCumulative
'
' WORKSHEET EQUIVALENTS (native Excel, 2010+)
'   K_STATS_StudentT_Density              T.DIST(x, df, FALSE)
'   K_STATS_StudentT_Cumulative           T.DIST(x, df, TRUE)
'   K_STATS_StudentT_Survival             T.DIST.RT(x, df)          [x >= 0 only in Excel]
'   K_STATS_StudentT_InverseCumulative    T.INV(p, df)
'   K_STATS_ChiSquare_Density             CHISQ.DIST(x, df, FALSE)
'   K_STATS_ChiSquare_Cumulative          CHISQ.DIST(x, df, TRUE)
'   K_STATS_ChiSquare_Survival            CHISQ.DIST.RT(x, df)
'   K_STATS_ChiSquare_InverseCumulative   CHISQ.INV(p, df)
'   K_STATS_F_Density                     F.DIST(x, df1, df2, FALSE)
'   K_STATS_F_Cumulative                  F.DIST(x, df1, df2, TRUE)
'   K_STATS_F_Survival                    F.DIST.RT(x, df1, df2)
'   K_STATS_F_InverseCumulative           F.INV(p, df1, df2)
'
' ALGORITHM PROVENANCE
'   - Student t density:
'       Closed-form log-density. The log-gamma difference is taken through
'       PROB_LogGammaHalfDiff rather than as a literal subtraction of two
'       log-gammas, which cancels for large df (a 4.3E-10 relative error at
'       df = 1E+6, 1.1E-7 at df = 1E+8).
'   - Student t tail:
'       Exact closed forms at df = 1 (arctangent) and df = 2 (algebraic, in the
'       rationalized arrangement that does not cancel), otherwise the standard
'       regularized incomplete beta transformation with both beta arguments passed
'       explicitly so that neither is recovered by subtraction.
'   - Student t quantile:
'       Cornish-Fisher normal expansion in 1/df as a seed, refined by Newton's
'       method on the survival function using the analytic density as derivative,
'       safeguarded by a bisection bracket. Exact closed forms at df = 1 and 2.
'   - Chi-square:
'       Regularized incomplete gamma, P(df/2, x/2) and Q(df/2, x/2). Quantile via
'       PROB_TryGammaInvP (Wilson-Hilferty seed, safeguarded Newton).
'   - F:
'       Regularized incomplete beta transformation with y = r/(1+r) and its
'       complement 1/(1+r), where r = (df1/df2)*x. Quantile via
'       PROB_TryBetaInvRegularized, recovering x = (df2/df1) * (X / Y) from the
'       returned pair so that the upper tail keeps full relative precision.
'   Nothing here is a newly-invented algorithm; the packaging, the validation
'   policy, the survival surface and the cancellation-free argument pairs are the
'   local contribution.
'
' DESIGN PRINCIPLES
'   - Public worksheet-facing functions return Variant so they can return CVErr.
'   - Private numerical kernels return Double or a Boolean Try contract, and avoid
'     worksheet-facing overhead.
'   - Invalid domains fail explicitly; they are not silently repaired.
'   - Non-convergence of any inner iteration is a failure, never a partial answer.
'   - Overflow fails explicitly and returns CVErr(xlErrNum) rather than a clamped
'     sentinel value. Underflow of a density is a valid zero.
'   - CDF functions return mathematically meaningful support-edge values, for
'     example ChiSquare_Cumulative(x <= 0) = 0.
'   - No MsgBox is raised by public worksheet-facing functions.
'
' ERROR POLICY
'   - Invalid numeric domains return CVErr(xlErrNum).
'   - Non-convergence or overflow returns CVErr(xlErrNum).
'   - Unexpected runtime errors return CVErr(xlErrValue).
'   - Detailed diagnostic messages are written to Status.
'   - Application.StatusBar is not written by default.
'
' ACCURACY (relative error, measured against 40 to 60 digit arithmetic)
'   Student t density                 <= 8.4E-15  (all df up to 1E+8)
'   Student t cumulative / survival   <= 1.3E-12  (worst at df = 1000; <= 4E-16 for df <= 30)
'   Student t quantile                <= 3E-12    (worst at df = 1E+5)
'   Chi-square cumulative / survival  <= 2.6E-10  (worst at df = 1E+6; <= 3E-13 for df <= 1000)
'   Chi-square quantile               <= 4.7E-12
'   F cumulative / survival           <= 1.1E-10  (worst at df1 = df2 = 1E+5)
'   F quantile                        <= 5.9E-13
'
' PERFORMANCE
'   The quantile routines use a seeded, safeguarded Newton iteration rather than
'   fixed-count bisection. The Student t quantile converges in a mean of 5.7
'   survival evaluations across the tested grid, and in zero for df = 1 and df = 2
'   where the closed form applies. The chi-square quantile averages 11 iterations
'   and the F quantile 17.
'
' NOTES
'   - Degrees of freedom are accepted as positive real numbers. This is more
'     general than some worksheet contexts, which treat them as integer-valued.
'   - Densities are unbounded at the left edge of the support for some parameter
'     values: chi-square at x = 0 with df < 2, and F at x = 0 with df1 < 2. Those
'     cases return CVErr(xlErrNum), matching Excel.
'   - Student t survival accepts negative x, unlike Excel's T.DIST.RT.
'   - F arguments are assembled through a log-ratio logistic pair; extreme
'     degree ratios saturate to the mathematically correct beta boundary instead
'     of raising an intermediate overflow.
'
' DEPENDENCIES
'   - M_STATS_PROBDIST_CORE
'   - M_STATS_PROBDIST_SPECIALFUNCS
'
' UPDATED
'   2026-07-09
'==============================================================================

'==============================================================================
' PRIVATE CONSTANTS
'==============================================================================

Private Const PROB_T_INV_MAX_ITER      As Long = 500      'Safeguarded Newton iterations, Student t quantile
Private Const PROB_T_INV_SEARCH_MAX    As Double = 1E+150 'Safeguarded iterative-branch ceiling


'==============================================================================
' PUBLIC - STUDENT T
'==============================================================================

Public Function K_STATS_StudentT_Density( _
    ByVal X As Double, _
    ByVal DegreesFreedom As Double, _
    Optional ByRef Status As String = "") _
    As Variant
'
'==============================================================================
' K_STATS_StudentT_Density
'------------------------------------------------------------------------------
' PURPOSE
'   Returns the Student t probability density function at X.
'
' WHY THIS EXISTS
'   Student t densities are used in regression diagnostics, likelihood examples,
'   robust residual analysis and inference with an estimated variance.
'
' WORKSHEET EQUIVALENT
'   T.DIST(X, DegreesFreedom, FALSE)
'
' INPUTS
'   X
'     Evaluation point.
'
'   DegreesFreedom
'     Student t degrees of freedom.
'     Must be strictly positive.
'
'   Status
'     Optional ByRef diagnostic message.
'     Empty on success.
'     Populated on failure.
'
' RETURNS
'   Variant
'     Success => Double density value.
'     Failure => CVErr(xlErrNum) or CVErr(xlErrValue).
'
' BEHAVIOR
'   - Validates X and DegreesFreedom.
'   - Uses the log-density form for numerical stability.
'   - Underflow of the far tail is a valid zero.
'
' ERROR POLICY
'   - Invalid numeric domains return CVErr(xlErrNum).
'   - Unexpected runtime errors return CVErr(xlErrValue).
'   - Detailed diagnostic messages are written to Status.
'   - No MsgBox is raised.
'
' DEPENDENCIES
'   - PROB_ValidateXAndDF
'   - PROB_TryStudentTPDF
'   - PROB_SetStatus
'
' UPDATED
'   2026-07-09
'==============================================================================
'
'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Dim Density             As Double          'Computed density
    Dim FailMsg             As String          'Detailed failure message

'------------------------------------------------------------------------------
' INITIALIZE
'------------------------------------------------------------------------------
    'Route unexpected runtime errors to the error handler
        On Error GoTo Err_Handler
    'Clear diagnostic status
        PROB_SetStatus Status, vbNullString
    'Initialize the failure message buffer
        FailMsg = vbNullString

'------------------------------------------------------------------------------
' VALIDATE INPUTS
'------------------------------------------------------------------------------
    'Validate Student t inputs
        If Not PROB_ValidateXAndDF( _
            X, _
            DegreesFreedom, _
            FailMsg, _
            "DegreesFreedom") Then GoTo Fail_Num

'------------------------------------------------------------------------------
' COMPUTE DENSITY
'------------------------------------------------------------------------------
    'Evaluate the density kernel
        If Not PROB_TryStudentTPDF(X, DegreesFreedom, Density, FailMsg) Then GoTo Fail_Num

    'Return the density
        K_STATS_StudentT_Density = Density

'------------------------------------------------------------------------------
' RETURN SUCCESS
'------------------------------------------------------------------------------
    'Clear diagnostic status
        PROB_SetStatus Status, vbNullString
    'Exit before failure and error-handler blocks
        Exit Function

'------------------------------------------------------------------------------
' FAIL - NUMERIC
'------------------------------------------------------------------------------
Fail_Num:
    'Write diagnostics
        PROB_SetStatus Status, FailMsg
    'Return worksheet numeric error
        K_STATS_StudentT_Density = CVErr(xlErrNum)
    'Exit before the error handler
        Exit Function

'------------------------------------------------------------------------------
' ERROR HANDLER
'------------------------------------------------------------------------------
Err_Handler:
    'Write unexpected runtime errors to diagnostics
        PROB_SetStatus Status, "Unexpected error in K_STATS_StudentT_Density: " & Err.Description
    'Return worksheet value error
        K_STATS_StudentT_Density = CVErr(xlErrValue)
End Function


Public Function K_STATS_StudentT_Cumulative( _
    ByVal X As Double, _
    ByVal DegreesFreedom As Double, _
    Optional ByRef Status As String = "") _
    As Variant
'
'==============================================================================
' K_STATS_StudentT_Cumulative
'------------------------------------------------------------------------------
' PURPOSE
'   Returns the left-tail Student t cumulative distribution function at X.
'
' WHY THIS EXISTS
'   Student t cumulative probabilities are used for regression t-statistics,
'   confidence intervals, residual diagnostics and p-value calculations.
'
' WORKSHEET EQUIVALENT
'   T.DIST(X, DegreesFreedom, TRUE)
'
' INPUTS
'   X
'     Evaluation point.
'
'   DegreesFreedom
'     Student t degrees of freedom.
'     Must be strictly positive.
'
'   Status
'     Optional ByRef diagnostic message.
'
' RETURNS
'   Variant
'     Success => Double cumulative probability P(T <= X).
'     Failure => CVErr(xlErrNum) or CVErr(xlErrValue).
'
' BEHAVIOR
'   - Computes the left tail directly for X < 0, and as 1 minus the right tail
'     for X >= 0. Both arrangements keep full precision on the side that carries
'     the information.
'   - For a small-magnitude right-tail probability, prefer
'     K_STATS_StudentT_Survival.
'
' ERROR POLICY
'   - Invalid numeric domains return CVErr(xlErrNum).
'   - Non-convergence returns CVErr(xlErrNum).
'   - Unexpected runtime errors return CVErr(xlErrValue).
'   - Detailed diagnostic messages are written to Status.
'   - No MsgBox is raised.
'
' DEPENDENCIES
'   - PROB_ValidateXAndDF
'   - PROB_TryStudentTTail
'   - PROB_SetStatus
'
' UPDATED
'   2026-07-09
'==============================================================================
'
'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Dim Tail                As Double          'Right-tail probability at Abs(X)
    Dim FailMsg             As String          'Detailed failure message

'------------------------------------------------------------------------------
' INITIALIZE
'------------------------------------------------------------------------------
    'Route unexpected runtime errors to the error handler
        On Error GoTo Err_Handler
    'Clear diagnostic status
        PROB_SetStatus Status, vbNullString
    'Initialize the failure message buffer
        FailMsg = vbNullString

'------------------------------------------------------------------------------
' VALIDATE INPUTS
'------------------------------------------------------------------------------
    'Validate Student t inputs
        If Not PROB_ValidateXAndDF( _
            X, _
            DegreesFreedom, _
            FailMsg, _
            "DegreesFreedom") Then GoTo Fail_Num

'------------------------------------------------------------------------------
' COMPUTE CUMULATIVE PROBABILITY
'------------------------------------------------------------------------------
    'Evaluate the right tail at Abs(X)
        If Not PROB_TryStudentTTail(Abs(X), DegreesFreedom, Tail, FailMsg) Then GoTo Fail_Num

    'Map the tail into the left-tail CDF by symmetry
        If X < 0# Then
            K_STATS_StudentT_Cumulative = Tail
        Else
            K_STATS_StudentT_Cumulative = 1# - Tail
        End If

'------------------------------------------------------------------------------
' RETURN SUCCESS
'------------------------------------------------------------------------------
    'Clear diagnostic status
        PROB_SetStatus Status, vbNullString
    'Exit before failure and error-handler blocks
        Exit Function

'------------------------------------------------------------------------------
' FAIL - NUMERIC
'------------------------------------------------------------------------------
Fail_Num:
    'Write diagnostics
        PROB_SetStatus Status, FailMsg
    'Return worksheet numeric error
        K_STATS_StudentT_Cumulative = CVErr(xlErrNum)
    'Exit before the error handler
        Exit Function

'------------------------------------------------------------------------------
' ERROR HANDLER
'------------------------------------------------------------------------------
Err_Handler:
    'Write unexpected runtime errors to diagnostics
        PROB_SetStatus Status, "Unexpected error in K_STATS_StudentT_Cumulative: " & Err.Description
    'Return worksheet value error
        K_STATS_StudentT_Cumulative = CVErr(xlErrValue)
End Function


Public Function K_STATS_StudentT_Survival( _
    ByVal X As Double, _
    ByVal DegreesFreedom As Double, _
    Optional ByRef Status As String = "") _
    As Variant
'
'==============================================================================
' K_STATS_StudentT_Survival
'------------------------------------------------------------------------------
' PURPOSE
'   Returns the right-tail Student t probability P(T > X).
'
' WHY THIS EXISTS
'   This is the one-sided p-value of a t-statistic. It cannot be obtained from
'   the CDF: 1 - K_STATS_StudentT_Cumulative(20, 30) evaluates to exactly zero,
'   while the true right tail is 3.3745418329E-19. Every right-tail probability
'   below about 1E-16 is invisible to a CDF-based calculation.
'
' WORKSHEET EQUIVALENT
'   T.DIST.RT(X, DegreesFreedom), which Excel restricts to X >= 0. This function
'   accepts any finite X.
'
' INPUTS
'   X
'     Evaluation point.
'
'   DegreesFreedom
'     Student t degrees of freedom.
'     Must be strictly positive.
'
'   Status
'     Optional ByRef diagnostic message.
'
' RETURNS
'   Variant
'     Success => Double right-tail probability P(T > X).
'     Failure => CVErr(xlErrNum) or CVErr(xlErrValue).
'
' ERROR POLICY
'   - Invalid numeric domains return CVErr(xlErrNum).
'   - Non-convergence returns CVErr(xlErrNum).
'   - Unexpected runtime errors return CVErr(xlErrValue).
'   - Detailed diagnostic messages are written to Status.
'   - No MsgBox is raised.
'
' DEPENDENCIES
'   - PROB_ValidateXAndDF
'   - PROB_TryStudentTTail
'   - PROB_SetStatus
'
' UPDATED
'   2026-07-09
'==============================================================================
'
'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Dim Tail                As Double          'Right-tail probability at Abs(X)
    Dim FailMsg             As String          'Detailed failure message

'------------------------------------------------------------------------------
' INITIALIZE
'------------------------------------------------------------------------------
    'Route unexpected runtime errors to the error handler
        On Error GoTo Err_Handler
    'Clear diagnostic status
        PROB_SetStatus Status, vbNullString
    'Initialize the failure message buffer
        FailMsg = vbNullString

'------------------------------------------------------------------------------
' VALIDATE INPUTS
'------------------------------------------------------------------------------
    'Validate Student t inputs
        If Not PROB_ValidateXAndDF( _
            X, _
            DegreesFreedom, _
            FailMsg, _
            "DegreesFreedom") Then GoTo Fail_Num

'------------------------------------------------------------------------------
' COMPUTE SURVIVAL PROBABILITY
'------------------------------------------------------------------------------
    'Evaluate the right tail at Abs(X)
        If Not PROB_TryStudentTTail(Abs(X), DegreesFreedom, Tail, FailMsg) Then GoTo Fail_Num

    'Map the tail into the survival function by symmetry
        If X > 0# Then
            K_STATS_StudentT_Survival = Tail
        Else
            K_STATS_StudentT_Survival = 1# - Tail
        End If

'------------------------------------------------------------------------------
' RETURN SUCCESS
'------------------------------------------------------------------------------
    'Clear diagnostic status
        PROB_SetStatus Status, vbNullString
    'Exit before failure and error-handler blocks
        Exit Function

'------------------------------------------------------------------------------
' FAIL - NUMERIC
'------------------------------------------------------------------------------
Fail_Num:
    'Write diagnostics
        PROB_SetStatus Status, FailMsg
    'Return worksheet numeric error
        K_STATS_StudentT_Survival = CVErr(xlErrNum)
    'Exit before the error handler
        Exit Function

'------------------------------------------------------------------------------
' ERROR HANDLER
'------------------------------------------------------------------------------
Err_Handler:
    'Write unexpected runtime errors to diagnostics
        PROB_SetStatus Status, "Unexpected error in K_STATS_StudentT_Survival: " & Err.Description
    'Return worksheet value error
        K_STATS_StudentT_Survival = CVErr(xlErrValue)
End Function


Public Function K_STATS_StudentT_InverseCumulative( _
    ByVal Probability As Double, _
    ByVal DegreesFreedom As Double, _
    Optional ByRef Status As String = "") _
    As Variant
'
'==============================================================================
' K_STATS_StudentT_InverseCumulative
'------------------------------------------------------------------------------
' PURPOSE
'   Returns the inverse left-tail Student t cumulative distribution function.
'
' WHY THIS EXISTS
'   Student t quantiles are required for confidence intervals, hypothesis tests
'   and regression coefficient inference.
'
' WORKSHEET EQUIVALENT
'   T.INV(Probability, DegreesFreedom)
'
' INPUTS
'   Probability
'     Left-tail probability.
'     Must be strictly between 0 and 1.
'
'   DegreesFreedom
'     Student t degrees of freedom.
'     Must be strictly positive.
'
'   Status
'     Optional ByRef diagnostic message.
'
' RETURNS
'   Variant
'     Success => Double quantile x such that P(T <= x) = Probability.
'     Failure => CVErr(xlErrNum) or CVErr(xlErrValue).
'
' BEHAVIOR
'   - Reduces by symmetry to the right tail, taking the complement as 1 -
'     Probability, which is exact for Probability >= 0.5.
'   - Inverts the survival function rather than the CDF. Inverting the CDF in the
'     right tail is ill-conditioned regardless of the root finder used: at
'     Probability = 1 - 1E-12 the CDF is known only to about 1E-4 relative, so the
'     quantile inherits that error.
'   - Uses exact closed forms at DegreesFreedom = 1 and 2, otherwise a
'     Cornish-Fisher seed refined by safeguarded Newton.
'   - Imposes no artificial bracket cap. T.INV(1E-14, 1) is -3.19E+13 and is
'     returned; a 1E+12 search cap would refuse it.
'
' ERROR POLICY
'   - Invalid numeric domains return CVErr(xlErrNum).
'   - Non-convergence returns CVErr(xlErrNum).
'   - Unexpected runtime errors return CVErr(xlErrValue).
'   - Detailed diagnostic messages are written to Status.
'   - No MsgBox is raised.
'
' DEPENDENCIES
'   - PROB_IsValidProbabilityOpen, PROB_IsPositiveFinite
'   - PROB_TryStudentTInvTail
'   - PROB_SetStatus
'
' UPDATED
'   2026-07-09
'==============================================================================
'
'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Dim Quantile            As Double          'Non-negative quantile magnitude
    Dim FailMsg             As String          'Detailed failure message

'------------------------------------------------------------------------------
' INITIALIZE
'------------------------------------------------------------------------------
    'Route unexpected runtime errors to the error handler
        On Error GoTo Err_Handler
    'Clear diagnostic status
        PROB_SetStatus Status, vbNullString
    'Initialize the failure message buffer
        FailMsg = vbNullString

'------------------------------------------------------------------------------
' VALIDATE INPUTS
'------------------------------------------------------------------------------
    'Validate probability domain
        If Not PROB_IsValidProbabilityOpen(Probability) Then
            FailMsg = "Probability must be strictly between 0 and 1"
            GoTo Fail_Num
        End If
    'Validate degrees of freedom
        If Not PROB_IsPositiveWithinSupportedMagnitude(DegreesFreedom) Then
            FailMsg = "DegreesFreedom must be a finite strictly positive number"
            GoTo Fail_Num
        End If
        If 0.5 * DegreesFreedom <= 0# Then
            FailMsg = "DegreesFreedom is too small for the half-degree special-function parameter"
            GoTo Fail_Num
        End If

'------------------------------------------------------------------------------
' COMPUTE QUANTILE
'------------------------------------------------------------------------------
    'Return zero at the median
        If Probability = 0.5 Then
            K_STATS_StudentT_InverseCumulative = 0#
            GoTo Return_Success
        End If

    'Invert the right tail; 1 - Probability is exact for Probability >= 0.5
        If Probability > 0.5 Then
            If Not PROB_TryStudentTInvTail( _
                1# - Probability, DegreesFreedom, Quantile, FailMsg) Then GoTo Fail_Num
            K_STATS_StudentT_InverseCumulative = Quantile
        Else
            If Not PROB_TryStudentTInvTail( _
                Probability, DegreesFreedom, Quantile, FailMsg) Then GoTo Fail_Num
            K_STATS_StudentT_InverseCumulative = -Quantile
        End If

'------------------------------------------------------------------------------
' RETURN SUCCESS
'------------------------------------------------------------------------------
Return_Success:
    'Clear diagnostic status
        PROB_SetStatus Status, vbNullString
    'Exit before failure and error-handler blocks
        Exit Function

'------------------------------------------------------------------------------
' FAIL - NUMERIC
'------------------------------------------------------------------------------
Fail_Num:
    'Populate a fallback failure message if required
        If Len(FailMsg) = 0 Then FailMsg = "Invalid inputs for Student t inverse cumulative distribution"
    'Write diagnostics
        PROB_SetStatus Status, FailMsg
    'Return worksheet numeric error
        K_STATS_StudentT_InverseCumulative = CVErr(xlErrNum)
    'Exit before the error handler
        Exit Function

'------------------------------------------------------------------------------
' ERROR HANDLER
'------------------------------------------------------------------------------
Err_Handler:
    'Write unexpected runtime errors to diagnostics
        PROB_SetStatus Status, "Unexpected error in K_STATS_StudentT_InverseCumulative: " & Err.Description
    'Return worksheet value error
        K_STATS_StudentT_InverseCumulative = CVErr(xlErrValue)
End Function


'==============================================================================
' PUBLIC - CHI-SQUARE
'==============================================================================

Public Function K_STATS_ChiSquare_Density( _
    ByVal X As Double, _
    ByVal DegreesFreedom As Double, _
    Optional ByRef Status As String = "") _
    As Variant
'
'==============================================================================
' K_STATS_ChiSquare_Density
'------------------------------------------------------------------------------
' PURPOSE
'   Returns the chi-square probability density function at X.
'
' WHY THIS EXISTS
'   Chi-square densities appear in likelihood surfaces, Bayesian variance priors
'   and simulation acceptance ratios.
'
' WORKSHEET EQUIVALENT
'   CHISQ.DIST(X, DegreesFreedom, FALSE)
'
' INPUTS
'   X
'     Evaluation point.
'     The density is zero for X < 0.
'
'   DegreesFreedom
'     Chi-square degrees of freedom.
'     Must be strictly positive.
'
'   Status
'     Optional ByRef diagnostic message.
'
' RETURNS
'   Variant
'     Success => Double density value.
'     Failure => CVErr(xlErrNum) or CVErr(xlErrValue).
'
' BEHAVIOR
'   - Returns 0 for X < 0.
'   - At X = 0 the density is 0 for df > 2, 0.5 for df = 2, and unbounded for
'     df < 2. The unbounded case returns CVErr(xlErrNum), matching Excel.
'
' ERROR POLICY
'   - Invalid numeric domains return CVErr(xlErrNum).
'   - Overflow returns CVErr(xlErrNum).
'   - Unexpected runtime errors return CVErr(xlErrValue).
'   - Detailed diagnostic messages are written to Status.
'   - No MsgBox is raised.
'
' DEPENDENCIES
'   - PROB_ValidateXAndDF
'   - PROB_LogGamma, PROB_TryExp
'   - PROB_SetStatus
'
' UPDATED
'   2026-07-09
'==============================================================================
'
'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Dim HalfDF              As Double          'DegreesFreedom / 2
    Dim LogDensity          As Double          'Log-density value
    Dim Density             As Double          'Density value
    Dim FailMsg             As String          'Detailed failure message

'------------------------------------------------------------------------------
' INITIALIZE
'------------------------------------------------------------------------------
    'Route unexpected runtime errors to the error handler
        On Error GoTo Err_Handler
    'Clear diagnostic status
        PROB_SetStatus Status, vbNullString
    'Initialize the failure message buffer
        FailMsg = vbNullString

'------------------------------------------------------------------------------
' VALIDATE INPUTS
'------------------------------------------------------------------------------
    'Validate chi-square inputs
        If Not PROB_ValidateXAndDF( _
            X, _
            DegreesFreedom, _
            FailMsg, _
            "DegreesFreedom") Then GoTo Fail_Num

'------------------------------------------------------------------------------
' HANDLE THE SUPPORT EDGE
'------------------------------------------------------------------------------
    'Return zero outside the support
        If X < 0# Then
            K_STATS_ChiSquare_Density = 0#
            GoTo Return_Success
        End If

    'Handle the origin, where the density is 0, 0.5 or unbounded
        If X = 0# Then
            If DegreesFreedom < 2# Then
                FailMsg = "Chi-square density is unbounded at X = 0 when DegreesFreedom < 2"
                GoTo Fail_Num
            ElseIf DegreesFreedom = 2# Then
                K_STATS_ChiSquare_Density = 0.5
            Else
                K_STATS_ChiSquare_Density = 0#
            End If
            GoTo Return_Success
        End If

'------------------------------------------------------------------------------
' COMPUTE DENSITY
'------------------------------------------------------------------------------
    'Compute half degrees of freedom
        HalfDF = 0.5 * DegreesFreedom

    'Compute the log-density
        LogDensity = _
            (HalfDF - 1#) * Log(X) - _
            0.5 * X - _
            HalfDF * Log(2#) - _
            PROB_LogGamma(HalfDF)

    'Exponentiate; underflow to zero is a valid result
        If Not PROB_TryExp(LogDensity, Density) Then
            FailMsg = "Chi-square density overflowed a Double"
            GoTo Fail_Num
        End If

    'Return the density
        K_STATS_ChiSquare_Density = Density

'------------------------------------------------------------------------------
' RETURN SUCCESS
'------------------------------------------------------------------------------
Return_Success:
    'Clear diagnostic status
        PROB_SetStatus Status, vbNullString
    'Exit before failure and error-handler blocks
        Exit Function

'------------------------------------------------------------------------------
' FAIL - NUMERIC
'------------------------------------------------------------------------------
Fail_Num:
    'Write diagnostics
        PROB_SetStatus Status, FailMsg
    'Return worksheet numeric error
        K_STATS_ChiSquare_Density = CVErr(xlErrNum)
    'Exit before the error handler
        Exit Function

'------------------------------------------------------------------------------
' ERROR HANDLER
'------------------------------------------------------------------------------
Err_Handler:
    'Write unexpected runtime errors to diagnostics
        PROB_SetStatus Status, "Unexpected error in K_STATS_ChiSquare_Density: " & Err.Description
    'Return worksheet value error
        K_STATS_ChiSquare_Density = CVErr(xlErrValue)
End Function


Public Function K_STATS_ChiSquare_Cumulative( _
    ByVal X As Double, _
    ByVal DegreesFreedom As Double, _
    Optional ByRef Status As String = "") _
    As Variant
'
'==============================================================================
' K_STATS_ChiSquare_Cumulative
'------------------------------------------------------------------------------
' PURPOSE
'   Returns the left-tail chi-square cumulative distribution function at X.
'
' WHY THIS EXISTS
'   Chi-square cumulative probabilities are used in goodness-of-fit tests,
'   likelihood-ratio tests, variance tests and model-validation diagnostics.
'
' WORKSHEET EQUIVALENT
'   CHISQ.DIST(X, DegreesFreedom, TRUE)
'
' INPUTS
'   X
'     Evaluation point.
'     For X <= 0, the cumulative probability is 0.
'
'   DegreesFreedom
'     Chi-square degrees of freedom.
'     Must be strictly positive.
'
'   Status
'     Optional ByRef diagnostic message.
'
' RETURNS
'   Variant
'     Success => Double cumulative probability P(ChiSq <= X).
'     Failure => CVErr(xlErrNum) or CVErr(xlErrValue).
'
' BEHAVIOR
'   - Returns 0 for X <= 0.
'   - Computes P(DegreesFreedom / 2, X / 2), the regularized lower incomplete
'     gamma function.
'   - Non-convergence of the underlying series or continued fraction returns
'     CVErr(xlErrNum). A partial sum is never returned.
'
' ERROR POLICY
'   - Invalid numeric domains return CVErr(xlErrNum).
'   - Non-convergence returns CVErr(xlErrNum).
'   - Unexpected runtime errors return CVErr(xlErrValue).
'   - Detailed diagnostic messages are written to Status.
'   - No MsgBox is raised.
'
' DEPENDENCIES
'   - PROB_ValidateXAndDF
'   - PROB_TryGammaRegularizedP
'   - PROB_SetStatus
'
' UPDATED
'   2026-07-09
'==============================================================================
'
'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Dim Value               As Double          'Cumulative probability
    Dim FailMsg             As String          'Detailed failure message

'------------------------------------------------------------------------------
' INITIALIZE
'------------------------------------------------------------------------------
    'Route unexpected runtime errors to the error handler
        On Error GoTo Err_Handler
    'Clear diagnostic status
        PROB_SetStatus Status, vbNullString
    'Initialize the failure message buffer
        FailMsg = vbNullString

'------------------------------------------------------------------------------
' VALIDATE INPUTS
'------------------------------------------------------------------------------
    'Validate chi-square inputs
        If Not PROB_ValidateXAndDF( _
            X, _
            DegreesFreedom, _
            FailMsg, _
            "DegreesFreedom") Then GoTo Fail_Num

'------------------------------------------------------------------------------
' COMPUTE CUMULATIVE PROBABILITY
'------------------------------------------------------------------------------
    'Return zero for values outside the positive support
        If X <= 0# Then
            K_STATS_ChiSquare_Cumulative = 0#
            GoTo Return_Success
        End If

    'Evaluate the regularized lower incomplete gamma function
        If Not PROB_TryGammaRegularizedP( _
            0.5 * DegreesFreedom, _
            0.5 * X, _
            Value, _
            FailMsg) Then GoTo Fail_Num

    'Return the cumulative probability
        K_STATS_ChiSquare_Cumulative = Value

'------------------------------------------------------------------------------
' RETURN SUCCESS
'------------------------------------------------------------------------------
Return_Success:
    'Clear diagnostic status
        PROB_SetStatus Status, vbNullString
    'Exit before failure and error-handler blocks
        Exit Function

'------------------------------------------------------------------------------
' FAIL - NUMERIC
'------------------------------------------------------------------------------
Fail_Num:
    'Write diagnostics
        PROB_SetStatus Status, FailMsg
    'Return worksheet numeric error
        K_STATS_ChiSquare_Cumulative = CVErr(xlErrNum)
    'Exit before the error handler
        Exit Function

'------------------------------------------------------------------------------
' ERROR HANDLER
'------------------------------------------------------------------------------
Err_Handler:
    'Write unexpected runtime errors to diagnostics
        PROB_SetStatus Status, "Unexpected error in K_STATS_ChiSquare_Cumulative: " & Err.Description
    'Return worksheet value error
        K_STATS_ChiSquare_Cumulative = CVErr(xlErrValue)
End Function


Public Function K_STATS_ChiSquare_Survival( _
    ByVal X As Double, _
    ByVal DegreesFreedom As Double, _
    Optional ByRef Status As String = "") _
    As Variant
'
'==============================================================================
' K_STATS_ChiSquare_Survival
'------------------------------------------------------------------------------
' PURPOSE
'   Returns the right-tail chi-square probability P(ChiSq > X).
'
' WHY THIS EXISTS
'   This is the p-value of a goodness-of-fit or likelihood-ratio statistic. It
'   cannot be recovered from the CDF: with 10 degrees of freedom at X = 200 the
'   right tail is 1.6139305337E-37, while 1 - CDF evaluates to exactly zero.
'
' WORKSHEET EQUIVALENT
'   CHISQ.DIST.RT(X, DegreesFreedom)
'
' INPUTS
'   X
'     Evaluation point.
'     For X <= 0, the survival probability is 1.
'
'   DegreesFreedom
'     Chi-square degrees of freedom.
'     Must be strictly positive.
'
'   Status
'     Optional ByRef diagnostic message.
'
' RETURNS
'   Variant
'     Success => Double right-tail probability P(ChiSq > X).
'     Failure => CVErr(xlErrNum) or CVErr(xlErrValue).
'
' ERROR POLICY
'   - Invalid numeric domains return CVErr(xlErrNum).
'   - Non-convergence returns CVErr(xlErrNum).
'   - Unexpected runtime errors return CVErr(xlErrValue).
'   - Detailed diagnostic messages are written to Status.
'   - No MsgBox is raised.
'
' DEPENDENCIES
'   - PROB_ValidateXAndDF
'   - PROB_TryGammaRegularizedQ
'   - PROB_SetStatus
'
' UPDATED
'   2026-07-09
'==============================================================================
'
'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Dim Value               As Double          'Survival probability
    Dim FailMsg             As String          'Detailed failure message

'------------------------------------------------------------------------------
' INITIALIZE
'------------------------------------------------------------------------------
    'Route unexpected runtime errors to the error handler
        On Error GoTo Err_Handler
    'Clear diagnostic status
        PROB_SetStatus Status, vbNullString
    'Initialize the failure message buffer
        FailMsg = vbNullString

'------------------------------------------------------------------------------
' VALIDATE INPUTS
'------------------------------------------------------------------------------
    'Validate chi-square inputs
        If Not PROB_ValidateXAndDF( _
            X, _
            DegreesFreedom, _
            FailMsg, _
            "DegreesFreedom") Then GoTo Fail_Num

'------------------------------------------------------------------------------
' COMPUTE SURVIVAL PROBABILITY
'------------------------------------------------------------------------------
    'Return one for values outside the positive support
        If X <= 0# Then
            K_STATS_ChiSquare_Survival = 1#
            GoTo Return_Success
        End If

    'Evaluate the regularized upper incomplete gamma function
        If Not PROB_TryGammaRegularizedQ( _
            0.5 * DegreesFreedom, _
            0.5 * X, _
            Value, _
            FailMsg) Then GoTo Fail_Num

    'Return the survival probability
        K_STATS_ChiSquare_Survival = Value

'------------------------------------------------------------------------------
' RETURN SUCCESS
'------------------------------------------------------------------------------
Return_Success:
    'Clear diagnostic status
        PROB_SetStatus Status, vbNullString
    'Exit before failure and error-handler blocks
        Exit Function

'------------------------------------------------------------------------------
' FAIL - NUMERIC
'------------------------------------------------------------------------------
Fail_Num:
    'Write diagnostics
        PROB_SetStatus Status, FailMsg
    'Return worksheet numeric error
        K_STATS_ChiSquare_Survival = CVErr(xlErrNum)
    'Exit before the error handler
        Exit Function

'------------------------------------------------------------------------------
' ERROR HANDLER
'------------------------------------------------------------------------------
Err_Handler:
    'Write unexpected runtime errors to diagnostics
        PROB_SetStatus Status, "Unexpected error in K_STATS_ChiSquare_Survival: " & Err.Description
    'Return worksheet value error
        K_STATS_ChiSquare_Survival = CVErr(xlErrValue)
End Function


Public Function K_STATS_ChiSquare_InverseCumulative( _
    ByVal Probability As Double, _
    ByVal DegreesFreedom As Double, _
    Optional ByRef Status As String = "") _
    As Variant
'
'==============================================================================
' K_STATS_ChiSquare_InverseCumulative
'------------------------------------------------------------------------------
' PURPOSE
'   Returns the inverse left-tail chi-square cumulative distribution function.
'
' WHY THIS EXISTS
'   Chi-square quantiles give critical values for goodness-of-fit and
'   likelihood-ratio tests and the endpoints of variance confidence intervals.
'
' WORKSHEET EQUIVALENT
'   CHISQ.INV(Probability, DegreesFreedom)
'
' INPUTS
'   Probability
'     Left-tail probability.
'     Must be strictly between 0 and 1.
'
'   DegreesFreedom
'     Chi-square degrees of freedom.
'     Must be strictly positive.
'
'   Status
'     Optional ByRef diagnostic message.
'
' RETURNS
'   Variant
'     Success => Double quantile x such that P(ChiSq <= x) = Probability.
'     Failure => CVErr(xlErrNum) or CVErr(xlErrValue).
'
' BEHAVIOR
'   - Returns 2 * G, where G solves P(DegreesFreedom / 2, G) = Probability.
'   - The solver drives whichever of P and Q is the smaller onto its target, so
'     the upper tail does not dissolve into cancellation.
'
' ERROR POLICY
'   - Invalid numeric domains return CVErr(xlErrNum).
'   - Non-convergence returns CVErr(xlErrNum).
'   - Unexpected runtime errors return CVErr(xlErrValue).
'   - Detailed diagnostic messages are written to Status.
'   - No MsgBox is raised.
'
' DEPENDENCIES
'   - PROB_IsValidProbabilityOpen, PROB_IsPositiveFinite
'   - PROB_TryGammaInvP
'   - PROB_SetStatus
'
' UPDATED
'   2026-07-09
'==============================================================================
'
'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Dim GammaQuantile       As Double          'Quantile of the unit-scale gamma
    Dim ChiSquareQuantile   As Double          'Rescaled chi-square quantile
    Dim FailMsg             As String          'Detailed failure message

'------------------------------------------------------------------------------
' INITIALIZE
'------------------------------------------------------------------------------
    'Route unexpected runtime errors to the error handler
        On Error GoTo Err_Handler
    'Clear diagnostic status
        PROB_SetStatus Status, vbNullString
    'Initialize the failure message buffer
        FailMsg = vbNullString

'------------------------------------------------------------------------------
' VALIDATE INPUTS
'------------------------------------------------------------------------------
    'Validate probability domain
        If Not PROB_IsValidProbabilityOpen(Probability) Then
            FailMsg = "Probability must be strictly between 0 and 1"
            GoTo Fail_Num
        End If
    'Validate degrees of freedom
        If Not PROB_IsPositiveWithinSupportedMagnitude(DegreesFreedom) Then
            FailMsg = "DegreesFreedom must be a finite strictly positive number"
            GoTo Fail_Num
        End If
        If 0.5 * DegreesFreedom <= 0# Then
            FailMsg = "DegreesFreedom is too small for the half-degree special-function parameter"
            GoTo Fail_Num
        End If

'------------------------------------------------------------------------------
' COMPUTE QUANTILE
'------------------------------------------------------------------------------
    'Invert the regularized lower incomplete gamma function
        If Not PROB_TryGammaInvP( _
            Probability, _
            1# - Probability, _
            0.5 * DegreesFreedom, _
            GammaQuantile, _
            FailMsg) Then GoTo Fail_Num

    'Rescale from the unit-scale gamma to the chi-square with an explicit
    'overflow classification.
        If Not PROB_TryMultiply(2#, GammaQuantile, ChiSquareQuantile) Then
            FailMsg = "Chi-square quantile overflowed a Double"
            GoTo Fail_Num
        End If

        K_STATS_ChiSquare_InverseCumulative = ChiSquareQuantile

'------------------------------------------------------------------------------
' RETURN SUCCESS
'------------------------------------------------------------------------------
    'Clear diagnostic status
        PROB_SetStatus Status, vbNullString
    'Exit before failure and error-handler blocks
        Exit Function

'------------------------------------------------------------------------------
' FAIL - NUMERIC
'------------------------------------------------------------------------------
Fail_Num:
    'Write diagnostics
        PROB_SetStatus Status, FailMsg
    'Return worksheet numeric error
        K_STATS_ChiSquare_InverseCumulative = CVErr(xlErrNum)
    'Exit before the error handler
        Exit Function

'------------------------------------------------------------------------------
' ERROR HANDLER
'------------------------------------------------------------------------------
Err_Handler:
    'Write unexpected runtime errors to diagnostics
        PROB_SetStatus Status, "Unexpected error in K_STATS_ChiSquare_InverseCumulative: " & Err.Description
    'Return worksheet value error
        K_STATS_ChiSquare_InverseCumulative = CVErr(xlErrValue)
End Function


'==============================================================================
' PUBLIC - F
'==============================================================================

Public Function K_STATS_F_Density( _
    ByVal X As Double, _
    ByVal DegreesFreedom1 As Double, _
    ByVal DegreesFreedom2 As Double, _
    Optional ByRef Status As String = "") _
    As Variant
'
    Dim HalfDF1            As Double
    Dim HalfDF2            As Double
    Dim LogRatio           As Double
    Dim LogOnePlusRatio    As Double
    Dim ExpNegLogRatio     As Double
    Dim LogDensity         As Double
    Dim Density            As Double
    Dim FailMsg            As String

        On Error GoTo Err_Handler
        PROB_SetStatus Status, vbNullString
        FailMsg = vbNullString

        If Not PROB_ValidateXAndTwoDF( _
            X, DegreesFreedom1, DegreesFreedom2, FailMsg) Then GoTo Fail_Num

        If X < 0# Then
            K_STATS_F_Density = 0#
            GoTo Return_Success
        End If

        If X = 0# Then
            If DegreesFreedom1 < 2# Then
                FailMsg = "F density is unbounded at X = 0 when DegreesFreedom1 < 2"
                GoTo Fail_Num
            ElseIf DegreesFreedom1 = 2# Then
                K_STATS_F_Density = 1#
            Else
                K_STATS_F_Density = 0#
            End If
            GoTo Return_Success
        End If

        HalfDF1 = 0.5 * DegreesFreedom1
        HalfDF2 = 0.5 * DegreesFreedom2
        LogRatio = Log(X) + Log(DegreesFreedom1) - Log(DegreesFreedom2)

        'For a positive log-ratio, expand Log(1+r) as Log(r)+Log(1+1/r).
        'This prevents cancellation between two terms of order df1*Log(r).
        If LogRatio >= 0# Then
            If Not PROB_TryExp(-LogRatio, ExpNegLogRatio) Then
                ExpNegLogRatio = 0#
            End If

            LogOnePlusRatio = PROB_Log1p(ExpNegLogRatio)
            LogDensity = _
                -HalfDF2 * LogRatio - _
                (HalfDF1 + HalfDF2) * LogOnePlusRatio - _
                Log(X) - PROB_LogBeta(HalfDF1, HalfDF2)
        Else
            LogOnePlusRatio = PROB_TF_LogOnePlusExp(LogRatio)
            LogDensity = _
                HalfDF1 * LogRatio - _
                (HalfDF1 + HalfDF2) * LogOnePlusRatio - _
                Log(X) - PROB_LogBeta(HalfDF1, HalfDF2)
        End If

        If Not PROB_TryExp(LogDensity, Density) Then
            FailMsg = "F density overflowed a Double"
            GoTo Fail_Num
        End If

        K_STATS_F_Density = Density

Return_Success:
        PROB_SetStatus Status, vbNullString
        Exit Function
Fail_Num:
        PROB_SetStatus Status, FailMsg
        K_STATS_F_Density = CVErr(xlErrNum)
        Exit Function
Err_Handler:
        PROB_SetStatus Status, "Unexpected error in K_STATS_F_Density: " & Err.Description
        K_STATS_F_Density = CVErr(xlErrValue)
End Function


Public Function K_STATS_F_Cumulative( _
    ByVal X As Double, _
    ByVal DegreesFreedom1 As Double, _
    ByVal DegreesFreedom2 As Double, _
    Optional ByRef Status As String = "") _
    As Variant
'
    Dim BetaX              As Double
    Dim BetaY              As Double
    Dim Value              As Double
    Dim FailMsg            As String

        On Error GoTo Err_Handler
        PROB_SetStatus Status, vbNullString
        FailMsg = vbNullString

        If Not PROB_ValidateXAndTwoDF( _
            X, DegreesFreedom1, DegreesFreedom2, FailMsg) Then GoTo Fail_Num

        If X <= 0# Then
            K_STATS_F_Cumulative = 0#
            GoTo Return_Success
        End If

        PROB_TF_LogisticPair _
            Log(X) + Log(DegreesFreedom1) - Log(DegreesFreedom2), BetaX, BetaY

        If Not PROB_TryBetaRegularized( _
            BetaX, BetaY, 0.5 * DegreesFreedom1, 0.5 * DegreesFreedom2, _
            Value, FailMsg) Then GoTo Fail_Num

        K_STATS_F_Cumulative = Value

Return_Success:
        PROB_SetStatus Status, vbNullString
        Exit Function
Fail_Num:
        PROB_SetStatus Status, FailMsg
        K_STATS_F_Cumulative = CVErr(xlErrNum)
        Exit Function
Err_Handler:
        PROB_SetStatus Status, "Unexpected error in K_STATS_F_Cumulative: " & Err.Description
        K_STATS_F_Cumulative = CVErr(xlErrValue)
End Function


Public Function K_STATS_F_Survival( _
    ByVal X As Double, _
    ByVal DegreesFreedom1 As Double, _
    ByVal DegreesFreedom2 As Double, _
    Optional ByRef Status As String = "") _
    As Variant
'
    Dim BetaX              As Double
    Dim BetaY              As Double
    Dim Value              As Double
    Dim FailMsg            As String

        On Error GoTo Err_Handler
        PROB_SetStatus Status, vbNullString
        FailMsg = vbNullString

        If Not PROB_ValidateXAndTwoDF( _
            X, DegreesFreedom1, DegreesFreedom2, FailMsg) Then GoTo Fail_Num

        If X <= 0# Then
            K_STATS_F_Survival = 1#
            GoTo Return_Success
        End If

        PROB_TF_LogisticPair _
            Log(X) + Log(DegreesFreedom1) - Log(DegreesFreedom2), BetaX, BetaY

        If Not PROB_TryBetaRegularized( _
            BetaY, BetaX, 0.5 * DegreesFreedom2, 0.5 * DegreesFreedom1, _
            Value, FailMsg) Then GoTo Fail_Num

        K_STATS_F_Survival = Value

Return_Success:
        PROB_SetStatus Status, vbNullString
        Exit Function
Fail_Num:
        PROB_SetStatus Status, FailMsg
        K_STATS_F_Survival = CVErr(xlErrNum)
        Exit Function
Err_Handler:
        PROB_SetStatus Status, "Unexpected error in K_STATS_F_Survival: " & Err.Description
        K_STATS_F_Survival = CVErr(xlErrValue)
End Function


Public Function K_STATS_F_InverseCumulative( _
    ByVal Probability As Double, _
    ByVal DegreesFreedom1 As Double, _
    ByVal DegreesFreedom2 As Double, _
    Optional ByRef Status As String = "") _
    As Variant
'
'==============================================================================
' K_STATS_F_InverseCumulative
'------------------------------------------------------------------------------
' PURPOSE
'   Returns the inverse left-tail F cumulative distribution function.
'
' WHY THIS EXISTS
'   F quantiles give the critical values of ANOVA and nested-model tests.
'
' WORKSHEET EQUIVALENT
'   F.INV(Probability, DegreesFreedom1, DegreesFreedom2)
'
' INPUTS
'   Probability
'     Left-tail probability.
'     Must be strictly between 0 and 1.
'
'   DegreesFreedom1
'     Numerator degrees of freedom. Must be strictly positive.
'
'   DegreesFreedom2
'     Denominator degrees of freedom. Must be strictly positive.
'
'   Status
'     Optional ByRef diagnostic message.
'
' RETURNS
'   Variant
'     Success => Double quantile x such that P(F <= x) = Probability.
'     Failure => CVErr(xlErrNum) or CVErr(xlErrValue).
'
' BEHAVIOR
'   - Inverts the incomplete beta, then maps back with x = (df2/df1) * (Y / Z),
'     where Y and Z are the beta root and its complement as returned by the
'     solver. Because the solver returns both, the small one is never recovered
'     by subtraction, and the far upper tail stays accurate: with df1 = df2 = 0.5
'     and Probability = 1 - 1E-9 the quantile is 8.46E+34 to 14 digits.
'
' ERROR POLICY
'   - Invalid numeric domains return CVErr(xlErrNum).
'   - Non-convergence returns CVErr(xlErrNum).
'   - A quantile that overflows a Double returns CVErr(xlErrNum).
'   - Unexpected runtime errors return CVErr(xlErrValue).
'   - Detailed diagnostic messages are written to Status.
'   - No MsgBox is raised.
'
' DEPENDENCIES
'   - PROB_IsValidProbabilityOpen, PROB_IsPositiveFinite
'   - PROB_TryBetaInvRegularized
'   - PROB_SetStatus
'
' UPDATED
'   2026-07-09
'==============================================================================
'
'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Dim BetaX               As Double          'Beta root
    Dim BetaY               As Double          'Complement of the beta root
    Dim LogQuantile         As Double          'Log of the F quantile
    Dim Quantile            As Double          'F quantile
    Dim FailMsg             As String          'Detailed failure message

'------------------------------------------------------------------------------
' INITIALIZE
'------------------------------------------------------------------------------
    'Route unexpected runtime errors to the error handler
        On Error GoTo Err_Handler
    'Clear diagnostic status
        PROB_SetStatus Status, vbNullString
    'Initialize the failure message buffer
        FailMsg = vbNullString

'------------------------------------------------------------------------------
' VALIDATE INPUTS
'------------------------------------------------------------------------------
    'Validate probability domain
        If Not PROB_IsValidProbabilityOpen(Probability) Then
            FailMsg = "Probability must be strictly between 0 and 1"
            GoTo Fail_Num
        End If
    'Validate numerator degrees of freedom
        If Not PROB_IsPositiveWithinSupportedMagnitude(DegreesFreedom1) Then
            FailMsg = "DegreesFreedom1 must be a finite strictly positive number"
            GoTo Fail_Num
        End If
    'Validate denominator degrees of freedom
        If Not PROB_IsPositiveWithinSupportedMagnitude(DegreesFreedom2) Then
            FailMsg = "DegreesFreedom2 must be a finite strictly positive number"
            GoTo Fail_Num
        End If
        If 0.5 * DegreesFreedom1 <= 0# Then
            FailMsg = "DegreesFreedom1 is too small for the half-degree special-function parameter"
            GoTo Fail_Num
        End If
        If 0.5 * DegreesFreedom2 <= 0# Then
            FailMsg = "DegreesFreedom2 is too small for the half-degree special-function parameter"
            GoTo Fail_Num
        End If

'------------------------------------------------------------------------------
' COMPUTE QUANTILE
'------------------------------------------------------------------------------
    'Invert the regularized incomplete beta function
        If Not PROB_TryBetaInvRegularized( _
            Probability, _
            1# - Probability, _
            0.5 * DegreesFreedom1, _
            0.5 * DegreesFreedom2, _
            BetaX, _
            BetaY, _
            FailMsg) Then GoTo Fail_Num

    'Reject a quantile at the upper edge of the support
        If BetaY <= 0# Then
            FailMsg = "F quantile overflowed a Double at Probability = " & Probability
            GoTo Fail_Num
        End If

    'Map the beta root back in the log domain, so neither the degrees-of-
    'freedom ratio nor the beta-root ratio can overflow as an intermediate.
        If BetaX <= 0# Then
            K_STATS_F_InverseCumulative = 0#
            GoTo Return_Success
        End If

        LogQuantile = _
            Log(DegreesFreedom2) - Log(DegreesFreedom1) + _
            Log(BetaX) - Log(BetaY)

        If Not PROB_TryExp(LogQuantile, Quantile) Then
            FailMsg = "F quantile overflowed a Double at Probability = " & Probability
            GoTo Fail_Num
        End If

        K_STATS_F_InverseCumulative = Quantile

'------------------------------------------------------------------------------
' RETURN SUCCESS
'------------------------------------------------------------------------------
Return_Success:
    'Clear diagnostic status
        PROB_SetStatus Status, vbNullString
    'Exit before failure and error-handler blocks
        Exit Function

'------------------------------------------------------------------------------
' FAIL - NUMERIC
'------------------------------------------------------------------------------
Fail_Num:
    'Write diagnostics
        PROB_SetStatus Status, FailMsg
    'Return worksheet numeric error
        K_STATS_F_InverseCumulative = CVErr(xlErrNum)
    'Exit before the error handler
        Exit Function

'------------------------------------------------------------------------------
' ERROR HANDLER
'------------------------------------------------------------------------------
Err_Handler:
    'Write unexpected runtime errors to diagnostics
        PROB_SetStatus Status, "Unexpected error in K_STATS_F_InverseCumulative: " & Err.Description
    'Return worksheet value error
        K_STATS_F_InverseCumulative = CVErr(xlErrValue)
End Function


'==============================================================================
' PRIVATE STUDENT T KERNELS
'==============================================================================

Private Function PROB_TryStudentTPDF( _
    ByVal X As Double, _
    ByVal DegreesFreedom As Double, _
    ByRef Result As Double, _
    ByRef FailMsg As String) _
    As Boolean
'
    Dim LogOnePlusRatio    As Double
    Dim LogDensity         As Double

        If Not PROB_TF_TryLogOnePlusSquareRatio( _
            Abs(X), DegreesFreedom, LogOnePlusRatio) Then
            FailMsg = "Student t square ratio could not be represented"
            Exit Function
        End If

        LogDensity = _
            PROB_LogGammaHalfDiff(0.5 * DegreesFreedom) - _
            0.5 * (Log(DegreesFreedom) + Log(PROB_PI)) - _
            0.5 * (DegreesFreedom + 1#) * LogOnePlusRatio

        If Not PROB_TryExp(LogDensity, Result) Then
            FailMsg = "Student t density overflowed a Double"
            Exit Function
        End If

        PROB_TryStudentTPDF = True
End Function


Private Function PROB_TryStudentTTail( _
    ByVal AbsX As Double, _
    ByVal DegreesFreedom As Double, _
    ByRef Result As Double, _
    ByRef FailMsg As String) _
    As Boolean
'
    Dim S                   As Double
    Dim InverseX            As Double
    Dim BetaX               As Double
    Dim BetaY               As Double
    Dim Ibeta               As Double
    Dim CentralMass         As Double

        If AbsX = 0# Then
            Result = 0.5
            PROB_TryStudentTTail = True
            Exit Function
        End If

        'For tiny arguments use the local odd-power integral. This preserves the
        'first representable displacement from 0.5 even when AbsX squared would
        'underflow or the beta complement would round to zero.
        If PROB_TF_TryStudentTCentralMass( _
            AbsX, DegreesFreedom, CentralMass, FailMsg) Then
            If CentralMass >= 0# Then
                Result = 0.5 - CentralMass
                If Result < 0# Then Result = 0#
                PROB_TryStudentTTail = True
                Exit Function
            End If
        ElseIf Len(FailMsg) > 0 Then
            Exit Function
        End If

        'Cauchy: use the central form for AbsX <= 1 to avoid 1 / AbsX overflow,
        'and the reciprocal form in the far tail to avoid subtraction.
        If DegreesFreedom = 1# Then
            If AbsX <= 1# Then
                Result = 0.5 - Atn(AbsX) / PROB_PI
            Else
                If Not PROB_TryDivide(1#, AbsX, InverseX) Then
                    Result = 0#
                Else
                    Result = Atn(InverseX) / PROB_PI
                End If
            End If
            PROB_TryStudentTTail = True
            Exit Function
        End If

        'Two degrees of freedom. Scale by AbsX in the far tail so AbsX squared
        'is never formed when it would overflow.
        If DegreesFreedom = 2# Then
            If AbsX <= 1# Then
                S = Sqr(2# + AbsX * AbsX)
                Result = 1# / (S * (S + AbsX))
            Else
                InverseX = 1# / AbsX
                S = Sqr(1# + 2# * InverseX * InverseX)
                Result = InverseX * InverseX / (S * (S + 1#))
            End If
            PROB_TryStudentTTail = True
            Exit Function
        End If

        If Not PROB_TF_TrySquareRatioPair( _
            AbsX, DegreesFreedom, BetaX, BetaY) Then
            FailMsg = "Student t beta arguments could not be represented"
            Exit Function
        End If

        If Not PROB_TryBetaRegularized( _
            BetaX, BetaY, 0.5 * DegreesFreedom, 0.5, Ibeta, FailMsg) Then Exit Function

        Result = 0.5 * Ibeta
        PROB_TryStudentTTail = True
End Function


Private Function PROB_TryStudentTInvTail( _
    ByVal Tail As Double, _
    ByVal DegreesFreedom As Double, _
    ByRef Result As Double, _
    ByRef FailMsg As String) _
    As Boolean
'
'==============================================================================
' PROB_TryStudentTInvTail
'------------------------------------------------------------------------------
' PURPOSE
'   Solves P(T > X) = Tail for X >= 0.
'
' NUMERICAL CONTRACT
'   - Exact df=1 and df=2 branches are evaluated through guarded reciprocal or
'     log-domain formulas.
'   - df < 0.5 uses direct incomplete-beta inversion, avoiding singular powers
'     of 1 / df in the Cornish-Fisher seed.
'   - The general safeguarded Newton branch guards division, addition and bracket
'     expansion so predictable overflow is returned to the public API as #NUM!.
'==============================================================================
'
    Dim Z                   As Double
    Dim Z2                  As Double
    Dim X                   As Double
    Dim XNew                As Double
    Dim Low                 As Double
    Dim High                As Double
    Dim HasHigh             As Boolean
    Dim TailAtX             As Double
    Dim Residual            As Double
    Dim Density             As Double
    Dim NewtonStep          As Double
    Dim TanValue            As Double
    Dim Numerator           As Double
    Dim LogQuantile         As Double
    Dim Converged           As Boolean
    Dim IterIdx             As Long

        If Tail >= 0.5 Then
            Result = 0#
            PROB_TryStudentTInvTail = True
            Exit Function
        End If

    'Cauchy closed form.
        If DegreesFreedom = 1# Then
            If Tail > 0.25 Then
                Result = Tan(PROB_PI * (0.5 - Tail))
            Else
                TanValue = Tan(PROB_PI * Tail)

                If Not PROB_TryDivide(1#, TanValue, Result) Then
                    FailMsg = "Student t quantile overflowed a Double for DegreesFreedom = 1"
                    Exit Function
                End If
            End If

            PROB_TryStudentTInvTail = True
            Exit Function
        End If

    'Two-degrees-of-freedom closed form in the log domain.
        If DegreesFreedom = 2# Then
            Numerator = 1# - 2# * Tail

            If Numerator <= 0# Then
                Result = 0#
                PROB_TryStudentTInvTail = True
                Exit Function
            End If

            LogQuantile = _
                Log(Numerator) - _
                0.5 * (Log(2#) + Log(Tail) + PROB_Log1p(-Tail))

            If Not PROB_TryExp(LogQuantile, Result) Then
                FailMsg = "Student t quantile overflowed a Double for DegreesFreedom = 2"
                Exit Function
            End If

            PROB_TryStudentTInvTail = True
            Exit Function
        End If

    'Very small degrees of freedom use the beta transformation directly.
        If DegreesFreedom < 0.5 Then
            If Not PROB_TF_TryStudentTInvTailSmallDF( _
                Tail, DegreesFreedom, Result, FailMsg) Then Exit Function

            PROB_TryStudentTInvTail = True
            Exit Function
        End If

    'Cornish-Fisher seed for the regular branch.
        Z = -PROB_NormalInvCDFRaw(Tail)
        Z2 = Z * Z
        X = Z + _
            (Z2 * Z + Z) / (4# * DegreesFreedom) + _
            (5# * Z2 * Z2 * Z + 16# * Z2 * Z + 3# * Z) / _
                (96# * DegreesFreedom * DegreesFreedom) + _
            (3# * Z2 * Z2 * Z2 * Z + 19# * Z2 * Z2 * Z + _
                17# * Z2 * Z - 15# * Z) / _
                (384# * DegreesFreedom * DegreesFreedom * DegreesFreedom)

        If X <= 0# Or Not PROB_IsFinite(X) Then X = 1#

        Low = 0#
        High = 0#
        HasHigh = False

        For IterIdx = 1 To PROB_T_INV_MAX_ITER
            If Not PROB_TryStudentTTail( _
                X, DegreesFreedom, TailAtX, FailMsg) Then Exit Function

            Residual = TailAtX - Tail

            If Residual > 0# Then
                If X > Low Then Low = X
            Else
                If (Not HasHigh) Or X < High Then High = X
                HasHigh = True
            End If

            If Not PROB_TryStudentTPDF( _
                X, DegreesFreedom, Density, FailMsg) Then Exit Function

            If Density <= 0# Then
                If HasHigh Then
                    XNew = 0.5 * (Low + High)
                ElseIf Not PROB_TryMultiply(2#, X, XNew) Then
                    FailMsg = "Student t inverse bracket expansion overflowed a Double"
                    Exit Function
                End If
            ElseIf Not PROB_TryDivide(Residual, Density, NewtonStep) Then
                If HasHigh Then
                    XNew = 0.5 * (Low + High)
                ElseIf Not PROB_TryMultiply(2#, X, XNew) Then
                    FailMsg = "Student t inverse Newton step overflowed a Double"
                    Exit Function
                End If
            ElseIf Not PROB_TryAdd(X, NewtonStep, XNew) Then
                If HasHigh Then
                    XNew = 0.5 * (Low + High)
                ElseIf Not PROB_TryMultiply(2#, X, XNew) Then
                    FailMsg = "Student t inverse iterate overflowed a Double"
                    Exit Function
                End If
            Else
                If HasHigh Then
                    If XNew <= Low Or XNew >= High Then
                        XNew = 0.5 * (Low + High)
                    End If
                ElseIf XNew <= Low Then
                    If Not PROB_TryMultiply(2#, X, XNew) Then
                        FailMsg = "Student t inverse bracket expansion overflowed a Double"
                        Exit Function
                    End If
                End If
            End If

            If XNew > PROB_T_INV_SEARCH_MAX Then
                FailMsg = "Student t quantile exceeds the supported iterative range " & _
                          "for DegreesFreedom = " & DegreesFreedom
                Exit Function
            End If

            If Abs(XNew - X) <= PROB_EPS * Abs(XNew) Or XNew = X Then
                X = XNew
                Converged = True
                Exit For
            End If

            X = XNew
        Next IterIdx

        If Not Converged Then
            FailMsg = "Student t inverse failed to converge in " & _
                      PROB_T_INV_MAX_ITER & " iterations for DegreesFreedom = " & _
                      DegreesFreedom
            Exit Function
        End If

        Result = X
        PROB_TryStudentTInvTail = True
End Function


Private Function PROB_TF_LogOnePlusExp( _
    ByVal LogRatio As Double) _
    As Double
'
'==============================================================================
' PURPOSE
'   Returns Log(1 + Exp(LogRatio)) without overflow or cancellation.
'==============================================================================
'
    Dim SmallTerm As Double

        If LogRatio > 0# Then
            If Not PROB_TryExp(-LogRatio, SmallTerm) Then SmallTerm = 0#
            PROB_TF_LogOnePlusExp = LogRatio + PROB_Log1p(SmallTerm)
        Else
            If Not PROB_TryExp(LogRatio, SmallTerm) Then SmallTerm = 0#
            PROB_TF_LogOnePlusExp = PROB_Log1p(SmallTerm)
        End If
End Function


Private Sub PROB_TF_LogisticPair( _
    ByVal LogRatio As Double, _
    ByRef LeftValue As Double, _
    ByRef RightValue As Double)
'
'==============================================================================
' PURPOSE
'   Forms r/(1+r) and 1/(1+r) directly from Log(r), without ever forming r
'   when it lies outside the Double range.
'==============================================================================
'
    Dim SmallTerm As Double
    Dim Denominator As Double

        If LogRatio >= 0# Then
            If Not PROB_TryExp(-LogRatio, SmallTerm) Then SmallTerm = 0#
            Denominator = 1# + SmallTerm
            LeftValue = 1# / Denominator
            RightValue = SmallTerm / Denominator
        Else
            If Not PROB_TryExp(LogRatio, SmallTerm) Then SmallTerm = 0#
            Denominator = 1# + SmallTerm
            LeftValue = SmallTerm / Denominator
            RightValue = 1# / Denominator
        End If
End Sub


Private Function PROB_TF_TryLogOnePlusSquareRatio( _
    ByVal AbsX As Double, _
    ByVal DegreesFreedom As Double, _
    ByRef Result As Double) _
    As Boolean
'
'==============================================================================
' PURPOSE
'   Returns Log(1 + AbsX^2 / DegreesFreedom) without squaring AbsX.
'==============================================================================
'
    Dim LogRatio As Double

        If AbsX = 0# Then
            Result = 0#
            PROB_TF_TryLogOnePlusSquareRatio = True
            Exit Function
        End If

        LogRatio = 2# * Log(AbsX) - Log(DegreesFreedom)
        Result = PROB_TF_LogOnePlusExp(LogRatio)
        PROB_TF_TryLogOnePlusSquareRatio = PROB_IsFinite(Result)
End Function


Private Function PROB_TF_TrySquareRatioPair( _
    ByVal AbsX As Double, _
    ByVal DegreesFreedom As Double, _
    ByRef BetaX As Double, _
    ByRef BetaY As Double) _
    As Boolean
'
'==============================================================================
' PURPOSE
'   Forms df/(df+x^2) and x^2/(df+x^2) through a log-ratio logistic pair.
'==============================================================================
'
    Dim LogRatio As Double

        If AbsX = 0# Then
            BetaX = 1#
            BetaY = 0#
            PROB_TF_TrySquareRatioPair = True
            Exit Function
        End If

        LogRatio = 2# * Log(AbsX) - Log(DegreesFreedom)
        PROB_TF_LogisticPair LogRatio, BetaY, BetaX
        PROB_TF_TrySquareRatioPair = True
End Function


Private Function PROB_TF_TryStudentTCentralMass( _
    ByVal AbsX As Double, _
    ByVal DegreesFreedom As Double, _
    ByRef Result As Double, _
    ByRef FailMsg As String) _
    As Boolean
'
'==============================================================================
' PURPOSE
'   Returns the probability mass from zero to AbsX when the local series is
'   appropriate. A negative Result is a sentinel meaning "use the general path".
'==============================================================================
'
    Dim LogCurvature       As Double
    Dim Curvature          As Double
    Dim DensityZero        As Double
    Dim Correction         As Double

        Result = -1#

        LogCurvature = _
            2# * Log(AbsX) + PROB_Log1p(DegreesFreedom) - Log(DegreesFreedom)

        If LogCurvature > Log(0.00000001) Then
            PROB_TF_TryStudentTCentralMass = True
            Exit Function
        End If

        If Not PROB_TryExp(LogCurvature, Curvature) Then Curvature = 0#

        If Not PROB_TryStudentTPDF( _
            0#, DegreesFreedom, DensityZero, FailMsg) Then Exit Function

        Correction = _
            1# - Curvature / 6# + _
            Curvature * Curvature * _
                (DegreesFreedom + 3#) / (40# * (DegreesFreedom + 1#))

        If Not PROB_TryMultiply(DensityZero, AbsX, Result) Then
            FailMsg = "Student t central probability overflowed"
            Exit Function
        End If

        Result = Result * Correction
        PROB_TF_TryStudentTCentralMass = True
End Function


Private Function PROB_TF_TryStudentTInvTailSmallDF( _
    ByVal Tail As Double, _
    ByVal DegreesFreedom As Double, _
    ByRef Result As Double, _
    ByRef FailMsg As String) _
    As Boolean
'
'==============================================================================
' PURPOSE
'   Inverts the Student t tail through the incomplete-beta root for small df,
'   avoiding the singular Cornish-Fisher powers of 1 / df.
'==============================================================================
'
    Dim BetaProbability    As Double
    Dim BetaComplement     As Double
    Dim BetaX              As Double
    Dim BetaY              As Double
    Dim LogQuantile        As Double

        BetaProbability = 2# * Tail
        BetaComplement = 2# * (0.5 - Tail)

        If Not PROB_TryBetaInvRegularized( _
            BetaProbability, BetaComplement, _
            0.5 * DegreesFreedom, 0.5, _
            BetaX, BetaY, FailMsg) Then Exit Function

        If BetaX <= 0# Then
            FailMsg = "Student t quantile overflowed for very small DegreesFreedom"
            Exit Function
        End If

        If BetaY <= 0# Then
            Result = 0#
            PROB_TF_TryStudentTInvTailSmallDF = True
            Exit Function
        End If

        LogQuantile = _
            0.5 * (Log(DegreesFreedom) + Log(BetaY) - Log(BetaX))

        If Not PROB_TryExp(LogQuantile, Result) Then
            FailMsg = "Student t quantile overflowed for very small DegreesFreedom"
            Exit Function
        End If

        PROB_TF_TryStudentTInvTailSmallDF = True
End Function


'==============================================================================
' PRIVATE VALIDATION HELPERS
'==============================================================================

Private Function PROB_ValidateXAndDF( _
    ByVal X As Double, _
    ByVal DegreesFreedom As Double, _
    ByRef FailMsg As String, _
    ByVal DFName As String) _
    As Boolean
'
'==============================================================================
' PURPOSE
'   Validates a finite evaluation point and a positive degree parameter inside
'   the special-function kernel's supported magnitude domain.
'==============================================================================
'
        If Not PROB_IsFinite(X) Then
            FailMsg = "X must be a finite number"
            Exit Function
        End If

        If Not PROB_IsPositiveWithinSupportedMagnitude(DegreesFreedom) Then
            FailMsg = DFName & " must be a supported finite strictly positive number"
            Exit Function
        End If

        If 0.5 * DegreesFreedom <= 0# Then
            FailMsg = DFName & " is too small for the half-degree special-function parameter"
            Exit Function
        End If

        PROB_ValidateXAndDF = True
End Function


Private Function PROB_ValidateXAndTwoDF( _
    ByVal X As Double, _
    ByVal DegreesFreedom1 As Double, _
    ByVal DegreesFreedom2 As Double, _
    ByRef FailMsg As String) _
    As Boolean
'
'==============================================================================
' PURPOSE
'   Validates a finite evaluation point and two positive degree parameters inside
'   the special-function kernel's supported magnitude domain.
'==============================================================================
'
        If Not PROB_IsFinite(X) Then
            FailMsg = "X must be a finite number"
            Exit Function
        End If

        If Not PROB_IsPositiveWithinSupportedMagnitude(DegreesFreedom1) Then
            FailMsg = "DegreesFreedom1 must be a supported finite strictly positive number"
            Exit Function
        End If

        If Not PROB_IsPositiveWithinSupportedMagnitude(DegreesFreedom2) Then
            FailMsg = "DegreesFreedom2 must be a supported finite strictly positive number"
            Exit Function
        End If

        If 0.5 * DegreesFreedom1 <= 0# Then
            FailMsg = "DegreesFreedom1 is too small for the half-degree special-function parameter"
            Exit Function
        End If

        If 0.5 * DegreesFreedom2 <= 0# Then
            FailMsg = "DegreesFreedom2 is too small for the half-degree special-function parameter"
            Exit Function
        End If

        PROB_ValidateXAndTwoDF = True
End Function




