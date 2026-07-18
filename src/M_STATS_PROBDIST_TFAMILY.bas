Attribute VB_Name = "M_STATS_PROBDIST_TFAMILY"
Option Explicit

'==============================================================================
' M_STATS_PROBDIST_TFAMILY
'------------------------------------------------------------------------------
' PURPOSE
'   Provides worksheet-facing functions for the Student t, chi-square and F
'   distributions: density, left-tail cumulative probability, right-tail
'   survival probability and inverse cumulative probability.
'
' WHY THIS EXISTS
'   These distributions form the classical test-statistic family:
'
'     - Student t models a standardized estimate with an estimated variance.
'     - Chi-square models sums of squared standardized normal variables.
'     - F models a ratio of scaled chi-square variables.
'
'   Their cumulative and survival functions reduce to the regularized
'   incomplete beta or incomplete gamma functions. Keeping the public wrappers
'   in one module makes validation, tail orientation, error handling and
'   worksheet behavior consistent across the family.
'
' PUBLIC SURFACE (12 UDFs)
'   Student t
'     K_STATS_StudentT_Density
'     K_STATS_StudentT_Cumulative
'     K_STATS_StudentT_Survival
'     K_STATS_StudentT_InverseCumulative
'
'   Chi-square
'     K_STATS_ChiSquare_Density
'     K_STATS_ChiSquare_Cumulative
'     K_STATS_ChiSquare_Survival
'     K_STATS_ChiSquare_InverseCumulative
'
'   F
'     K_STATS_F_Density
'     K_STATS_F_Cumulative
'     K_STATS_F_Survival
'     K_STATS_F_InverseCumulative
'
' WORKSHEET EQUIVALENTS
'   K_STATS_StudentT_Density              T.DIST(X, DF, FALSE)
'   K_STATS_StudentT_Cumulative           T.DIST(X, DF, TRUE)
'   K_STATS_StudentT_Survival             T.DIST.RT(X, DF)
'   K_STATS_StudentT_InverseCumulative    T.INV(P, DF)
'
'   K_STATS_ChiSquare_Density             CHISQ.DIST(X, DF, FALSE)
'   K_STATS_ChiSquare_Cumulative          CHISQ.DIST(X, DF, TRUE)
'   K_STATS_ChiSquare_Survival            CHISQ.DIST.RT(X, DF)
'   K_STATS_ChiSquare_InverseCumulative   CHISQ.INV(P, DF)
'
'   K_STATS_F_Density                     F.DIST(X, DF1, DF2, FALSE)
'   K_STATS_F_Cumulative                  F.DIST(X, DF1, DF2, TRUE)
'   K_STATS_F_Survival                    F.DIST.RT(X, DF1, DF2)
'   K_STATS_F_InverseCumulative           F.INV(P, DF1, DF2)
'
' PARAMETERIZATION
'   - Degrees of freedom are accepted as positive real numbers.
'   - Degree parameters are restricted to the numerical range supported by the
'     shared special-function kernels.
'   - Student t survival accepts negative X, unlike Excel T.DIST.RT.
'
' NUMERICAL DESIGN
'   Student t density
'     Uses a logarithmic density and PROB_LogGammaHalfDiff so large degrees of
'     freedom do not subtract two nearly equal log-gamma values.
'
'   Student t cumulative and survival
'     Use exact closed forms for one and two degrees of freedom. Other cases use
'     the incomplete-beta transformation with both complementary beta arguments
'     formed directly. Tiny arguments use a local central-mass expansion.
'
'   Student t inverse
'     Uses exact closed forms for one and two degrees of freedom. Very small
'     degrees of freedom use direct beta inversion. The general branch uses a
'     Cornish-Fisher seed followed by safeguarded Newton iteration and bisection.
'     No artificial quantile ceiling is imposed; actual Double overflow is
'     detected through the shared arithmetic Try-contract.
'
'   Chi-square
'     Uses the regularized incomplete gamma functions P(DF / 2, X / 2) and
'     Q(DF / 2, X / 2). Quantiles use PROB_TryGammaInvP and a guarded rescale.
'
'   F
'     Uses a log-ratio logistic pair to form r / (1 + r) and 1 / (1 + r), where
'     r = X * DF1 / DF2. Neither r nor the degree ratio is formed directly.
'     The quantile is reconstructed in the logarithmic domain from the beta root
'     and its explicitly returned complement.
'
' DESIGN PRINCIPLES
'   - Public worksheet functions return Variant so failures can return CVErr.
'   - Invalid domains are rejected explicitly and are never silently repaired.
'   - Predictable numerical failure returns CVErr(xlErrNum).
'   - Unexpected runtime failure returns CVErr(xlErrValue).
'   - Mathematically valid underflow returns zero.
'   - Right-tail probabilities are evaluated directly, not as 1 minus the CDF.
'   - No public function raises a MsgBox.
'
' ERROR POLICY
'   - Invalid parameters, density poles, non-convergence and predictable
'     overflow return CVErr(xlErrNum).
'   - Unexpected runtime errors return CVErr(xlErrValue).
'   - Detailed diagnostics are written to the optional Status argument.
'   - Application.StatusBar is not written by this module.
'
' ACCURACY REFERENCE
'   The current regression grid targets approximately:
'
'     Student t density                 <= 2E-14 relative error (tested range)
'     Student t cumulative / survival   <= 1.3E-12 relative error
'     Student t quantile                <= 3.0E-12 relative error
'     Chi-square cumulative / survival  <= 2.6E-10 relative error
'     Chi-square quantile               <= 4.7E-12 relative error
'     F cumulative / survival           <= 1.1E-10 relative error
'     F quantile                        <= 5.9E-13 relative error
'
' DEPENDENCIES
'   - M_STATS_PROBDIST_CORE
'       PROB_EPS
'       PROB_PI
'       PROB_IsFinite
'       PROB_IsPositiveWithinSupportedMagnitude
'       PROB_IsValidProbabilityOpen
'       PROB_TryAdd
'       PROB_TryMultiply
'       PROB_TryDivide
'       PROB_TryExp
'       PROB_Log1p
'       PROB_SetStatus
'       PROB_NormalInvCDFRaw
'
'   - M_STATS_PROBDIST_SPECIALFUNCS
'       PROB_LogGamma
'       PROB_LogGammaHalfDiff
'       PROB_LogBeta
'       PROB_TryGammaRegularizedP
'       PROB_TryGammaRegularizedQ
'       PROB_TryGammaInvP
'       PROB_TryBetaRegularized
'       PROB_TryBetaInvRegularized
'
' NOTES
'   - Chi-square density is unbounded at X = 0 when DF < 2.
'   - F density is unbounded at X = 0 when DF1 < 2.
'   - Those density poles return CVErr(xlErrNum).
'   - The survival functions should be used for small right-tail probabilities;
'     subtracting a CDF from one loses the tail once the CDF rounds to one.
'
' UPDATED
'   2026-07-11 - House-style rewrite and numerical-contract hardening.
'==============================================================================


'==============================================================================
' PRIVATE CONSTANTS
'==============================================================================

'Maximum safeguarded Newton iterations for the Student t quantile
Private Const PROB_T_INV_MAX_ITER As Long = 500

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
'   - PROB_TF_ValidateXAndDF
'   - PROB_TryStudentTPDF
'   - PROB_SetStatus
'
' UPDATED
'   2026-07-11 - House-style rewrite and numerical-contract hardening.
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
        If Not PROB_TF_ValidateXAndDF( _
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
'   - PROB_TF_ValidateXAndDF
'   - PROB_TryStudentTTail
'   - PROB_SetStatus
'
' UPDATED
'   2026-07-11 - House-style rewrite and numerical-contract hardening.
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
        If Not PROB_TF_ValidateXAndDF( _
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
'   - PROB_TF_ValidateXAndDF
'   - PROB_TryStudentTTail
'   - PROB_SetStatus
'
' UPDATED
'   2026-07-11 - House-style rewrite and numerical-contract hardening.
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
        If Not PROB_TF_ValidateXAndDF( _
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
'   - Imposes no artificial quantile cap. Guarded arithmetic detects the
'     actual Double boundary.
'
' ERROR POLICY
'   - Invalid numeric domains return CVErr(xlErrNum).
'   - Non-convergence returns CVErr(xlErrNum).
'   - Unexpected runtime errors return CVErr(xlErrValue).
'   - Detailed diagnostic messages are written to Status.
'   - No MsgBox is raised.
'
' DEPENDENCIES
'   - PROB_IsValidProbabilityOpen
'   - PROB_TF_ValidateDF
'   - PROB_TryStudentTInvTail
'   - PROB_SetStatus
'
' UPDATED
'   2026-07-11 - House-style rewrite and numerical-contract hardening.
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

    'Validate the degree parameter and its half-degree kernel argument
        If Not PROB_TF_ValidateDF( _
            DegreesFreedom, "DegreesFreedom", FailMsg) Then
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
'   - PROB_TF_ValidateXAndDF
'   - PROB_LogGamma, PROB_TryExp
'   - PROB_SetStatus
'
' UPDATED
'   2026-07-11 - House-style rewrite and numerical-contract hardening.
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
        If Not PROB_TF_ValidateXAndDF( _
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
'   - PROB_TF_ValidateXAndDF
'   - PROB_TryGammaRegularizedP
'   - PROB_SetStatus
'
' UPDATED
'   2026-07-11 - House-style rewrite and numerical-contract hardening.
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
        If Not PROB_TF_ValidateXAndDF( _
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
'   - PROB_TF_ValidateXAndDF
'   - PROB_TryGammaRegularizedQ
'   - PROB_SetStatus
'
' UPDATED
'   2026-07-11 - House-style rewrite and numerical-contract hardening.
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
        If Not PROB_TF_ValidateXAndDF( _
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
'   - PROB_IsValidProbabilityOpen
'   - PROB_TF_ValidateDF
'   - PROB_TryGammaInvP
'   - PROB_TryMultiply
'   - PROB_SetStatus
'
' UPDATED
'   2026-07-11 - House-style rewrite and numerical-contract hardening.
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

    'Validate the degree parameter and its half-degree kernel argument
        If Not PROB_TF_ValidateDF( _
            DegreesFreedom, "DegreesFreedom", FailMsg) Then
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
'==============================================================================
' K_STATS_F_Density
'------------------------------------------------------------------------------
' PURPOSE
'   Returns the F probability density at X.
'
' WHY THIS EXISTS
'   The F density is used in variance-ratio likelihoods, ANOVA reference plots
'   and nested-model diagnostics. The logarithmic implementation supports
'   extreme degree ratios without forming X * DegreesFreedom1 / DegreesFreedom2.
'
' WORKSHEET EQUIVALENT
'   F.DIST(X, DegreesFreedom1, DegreesFreedom2, FALSE)
'
' INPUTS
'   X                 Evaluation point. Values below zero have density zero.
'   DegreesFreedom1   Positive numerator degrees of freedom.
'   DegreesFreedom2   Positive denominator degrees of freedom.
'   Status            Optional ByRef diagnostic message.
'
' RETURNS
'   Variant
'     Success => Double density.
'     Failure => CVErr(xlErrNum) or CVErr(xlErrValue).
'
' BEHAVIOR
'   - Returns zero below the support.
'   - At X = 0, returns #NUM! for DegreesFreedom1 < 2, one for
'     DegreesFreedom1 = 2, and zero for DegreesFreedom1 > 2.
'   - Forms the scaled ratio only through its logarithm.
'   - Uses separate positive- and negative-log-ratio forms to prevent
'     cancellation between large logarithmic terms.
'   - Mathematically valid exponential underflow returns zero.
'
' ERROR POLICY
'   - Invalid parameters, a density pole or density overflow return #NUM!.
'   - Unexpected runtime errors return #VALUE!.
'   - Diagnostics are written to Status.
'
' DEPENDENCIES
'   - PROB_TF_ValidateXAndTwoDF
'   - PROB_TF_LogOnePlusExp
'   - PROB_LogBeta
'   - PROB_TryExp
'   - PROB_SetStatus
'
' CALLED FROM
'   - Worksheet formulas
'   - M_STATS_PROBDIST_TEST
'
' UPDATED
'   2026-07-11 - House-style rewrite; log-ratio hardening retained.
'==============================================================================
'
'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Dim HalfDF1             As Double          'DegreesFreedom1 divided by two
    Dim HalfDF2             As Double          'DegreesFreedom2 divided by two
    Dim LogRatio            As Double          'Log(X * DF1 / DF2)
    Dim LogOnePlusRatio     As Double          'Stable residual Log(1 + ratio)
    Dim ExpNegLogRatio      As Double          'Exp(-LogRatio)
    Dim LogDensity          As Double          'Logarithm of the density
    Dim Density             As Double          'Returned density
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
    'Validate X and both degree parameters
        If Not PROB_TF_ValidateXAndTwoDF( _
            X, DegreesFreedom1, DegreesFreedom2, FailMsg) Then
            GoTo Fail_Num
        End If

'------------------------------------------------------------------------------
' HANDLE SUPPORT EDGES
'------------------------------------------------------------------------------
    'Return zero below the positive support
        If X < 0# Then
            K_STATS_F_Density = 0#
            GoTo Return_Success
        End If

    'Handle the origin, where the density is zero, one or unbounded
        If X = 0# Then
            If DegreesFreedom1 < 2# Then
                FailMsg = _
                    "F density is unbounded at X = 0 when " & _
                    "DegreesFreedom1 < 2"
                GoTo Fail_Num
            ElseIf DegreesFreedom1 = 2# Then
                K_STATS_F_Density = 1#
            Else
                K_STATS_F_Density = 0#
            End If

            GoTo Return_Success
        End If

'------------------------------------------------------------------------------
' COMPUTE LOG-DENSITY
'------------------------------------------------------------------------------
    'Form the half-degree parameters
        HalfDF1 = 0.5 * DegreesFreedom1
        HalfDF2 = 0.5 * DegreesFreedom2

    'Form the logarithm of the scaled F ratio without direct multiplication
        LogRatio = _
            Log(X) + _
            Log(DegreesFreedom1) - _
            Log(DegreesFreedom2)

    'Use the reciprocal ratio when LogRatio is non-negative
        If LogRatio >= 0# Then
            If Not PROB_TryExp(-LogRatio, ExpNegLogRatio) Then
                ExpNegLogRatio = 0#
            End If

            LogOnePlusRatio = PROB_Log1p(ExpNegLogRatio)

            LogDensity = _
                -HalfDF2 * LogRatio - _
                (HalfDF1 + HalfDF2) * LogOnePlusRatio - _
                Log(X) - _
                PROB_LogBeta(HalfDF1, HalfDF2)

    'Use the direct softplus form when LogRatio is negative
        Else
            LogOnePlusRatio = PROB_TF_LogOnePlusExp(LogRatio)

            LogDensity = _
                HalfDF1 * LogRatio - _
                (HalfDF1 + HalfDF2) * LogOnePlusRatio - _
                Log(X) - _
                PROB_LogBeta(HalfDF1, HalfDF2)
        End If

'------------------------------------------------------------------------------
' EXPONENTIATE
'------------------------------------------------------------------------------
    'Exponentiate the log-density; far-tail underflow is a valid zero
        If Not PROB_TryExp(LogDensity, Density) Then
            FailMsg = "F density overflowed a Double"
            GoTo Fail_Num
        End If

    'Return the density
        K_STATS_F_Density = Density

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
        K_STATS_F_Density = CVErr(xlErrNum)
    'Exit before the error handler
        Exit Function

'------------------------------------------------------------------------------
' ERROR HANDLER
'------------------------------------------------------------------------------
Err_Handler:
    'Write unexpected runtime errors to diagnostics
        PROB_SetStatus Status, "Unexpected error in K_STATS_F_Density: " & Err.Description
    'Return worksheet value error
        K_STATS_F_Density = CVErr(xlErrValue)
End Function

Public Function K_STATS_F_Cumulative( _
    ByVal X As Double, _
    ByVal DegreesFreedom1 As Double, _
    ByVal DegreesFreedom2 As Double, _
    Optional ByRef Status As String = "") _
    As Variant
'
'==============================================================================
' K_STATS_F_Cumulative
'------------------------------------------------------------------------------
' PURPOSE
'   Returns the left-tail F cumulative probability P(F <= X).
'
' WHY THIS EXISTS
'   F cumulative probabilities support ANOVA, nested-model comparison and
'   variance-ratio testing. Both incomplete-beta arguments are formed directly
'   so extreme degree ratios do not overflow and the complement is not recovered
'   through subtraction.
'
' WORKSHEET EQUIVALENT
'   F.DIST(X, DegreesFreedom1, DegreesFreedom2, TRUE)
'
' INPUTS
'   X                 Evaluation point. Values at or below zero return zero.
'   DegreesFreedom1   Positive numerator degrees of freedom.
'   DegreesFreedom2   Positive denominator degrees of freedom.
'   Status            Optional ByRef diagnostic message.
'
' RETURNS
'   Variant
'     Success => Double cumulative probability.
'     Failure => CVErr(xlErrNum) or CVErr(xlErrValue).
'
' BEHAVIOR
'   - Returns zero for X <= 0.
'   - Forms the beta pair from Log(X * DegreesFreedom1 / DegreesFreedom2).
'   - Evaluates I_BetaX(DF1 / 2, DF2 / 2).
'   - For a small upper tail, K_STATS_F_Survival should be used directly.
'
' ERROR POLICY
'   - Invalid parameters or non-convergence return #NUM!.
'   - Unexpected runtime errors return #VALUE!.
'   - Diagnostics are written to Status.
'
' DEPENDENCIES
'   - PROB_TF_ValidateXAndTwoDF
'   - PROB_TF_LogisticPair
'   - PROB_TryBetaRegularized
'   - PROB_SetStatus
'
' CALLED FROM
'   - Worksheet formulas
'   - M_STATS_PROBDIST_TEST
'
' UPDATED
'   2026-07-11 - House-style rewrite; log-ratio hardening retained.
'==============================================================================
'
'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Dim LogRatio            As Double          'Log(X * DF1 / DF2)
    Dim BetaX               As Double          'Ratio divided by one plus ratio
    Dim BetaY               As Double          'Complementary beta argument
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
    'Validate X and both degree parameters
        If Not PROB_TF_ValidateXAndTwoDF( _
            X, DegreesFreedom1, DegreesFreedom2, FailMsg) Then
            GoTo Fail_Num
        End If

'------------------------------------------------------------------------------
' HANDLE SUPPORT EDGE
'------------------------------------------------------------------------------
    'Return zero outside the positive support
        If X <= 0# Then
            K_STATS_F_Cumulative = 0#
            GoTo Return_Success
        End If

'------------------------------------------------------------------------------
' FORM BETA ARGUMENTS
'------------------------------------------------------------------------------
    'Form the log-ratio without direct multiplication or division
        LogRatio = _
            Log(X) + _
            Log(DegreesFreedom1) - _
            Log(DegreesFreedom2)

    'Form both complementary beta arguments from the same log-ratio
        PROB_TF_LogisticPair LogRatio, BetaX, BetaY

'------------------------------------------------------------------------------
' COMPUTE CUMULATIVE PROBABILITY
'------------------------------------------------------------------------------
    'Evaluate the regularized incomplete beta function
        If Not PROB_TryBetaRegularized( _
            BetaX, _
            BetaY, _
            0.5 * DegreesFreedom1, _
            0.5 * DegreesFreedom2, _
            Value, _
            FailMsg) Then
            GoTo Fail_Num
        End If

    'Return the cumulative probability
        K_STATS_F_Cumulative = Value

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
        K_STATS_F_Cumulative = CVErr(xlErrNum)
    'Exit before the error handler
        Exit Function

'------------------------------------------------------------------------------
' ERROR HANDLER
'------------------------------------------------------------------------------
Err_Handler:
    'Write unexpected runtime errors to diagnostics
        PROB_SetStatus Status, "Unexpected error in K_STATS_F_Cumulative: " & Err.Description
    'Return worksheet value error
        K_STATS_F_Cumulative = CVErr(xlErrValue)
End Function

Public Function K_STATS_F_Survival( _
    ByVal X As Double, _
    ByVal DegreesFreedom1 As Double, _
    ByVal DegreesFreedom2 As Double, _
    Optional ByRef Status As String = "") _
    As Variant
'
'==============================================================================
' K_STATS_F_Survival
'------------------------------------------------------------------------------
' PURPOSE
'   Returns the right-tail F probability P(F > X).
'
' WHY THIS EXISTS
'   This is the one-sided p-value of an ANOVA or nested-model F statistic. It is
'   evaluated as the reflected incomplete beta and therefore retains a small
'   right tail that would be lost in the subtraction 1 - CDF.
'
' WORKSHEET EQUIVALENT
'   F.DIST.RT(X, DegreesFreedom1, DegreesFreedom2)
'
' INPUTS
'   X                 Evaluation point. Values at or below zero return one.
'   DegreesFreedom1   Positive numerator degrees of freedom.
'   DegreesFreedom2   Positive denominator degrees of freedom.
'   Status            Optional ByRef diagnostic message.
'
' RETURNS
'   Variant
'     Success => Double survival probability.
'     Failure => CVErr(xlErrNum) or CVErr(xlErrValue).
'
' BEHAVIOR
'   - Returns one for X <= 0.
'   - Forms the beta pair from Log(X * DegreesFreedom1 / DegreesFreedom2).
'   - Evaluates the reflected beta I_BetaY(DF2 / 2, DF1 / 2).
'
' ERROR POLICY
'   - Invalid parameters or non-convergence return #NUM!.
'   - Unexpected runtime errors return #VALUE!.
'   - Diagnostics are written to Status.
'
' DEPENDENCIES
'   - PROB_TF_ValidateXAndTwoDF
'   - PROB_TF_LogisticPair
'   - PROB_TryBetaRegularized
'   - PROB_SetStatus
'
' CALLED FROM
'   - Worksheet formulas
'   - M_STATS_PROBDIST_TEST
'
' UPDATED
'   2026-07-11 - House-style rewrite; reflected-tail hardening retained.
'==============================================================================
'
'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Dim LogRatio            As Double          'Log(X * DF1 / DF2)
    Dim BetaX               As Double          'Ratio divided by one plus ratio
    Dim BetaY               As Double          'Complementary beta argument
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
    'Validate X and both degree parameters
        If Not PROB_TF_ValidateXAndTwoDF( _
            X, DegreesFreedom1, DegreesFreedom2, FailMsg) Then
            GoTo Fail_Num
        End If

'------------------------------------------------------------------------------
' HANDLE SUPPORT EDGE
'------------------------------------------------------------------------------
    'Return one outside the positive support
        If X <= 0# Then
            K_STATS_F_Survival = 1#
            GoTo Return_Success
        End If

'------------------------------------------------------------------------------
' FORM BETA ARGUMENTS
'------------------------------------------------------------------------------
    'Form the log-ratio without direct multiplication or division
        LogRatio = _
            Log(X) + _
            Log(DegreesFreedom1) - _
            Log(DegreesFreedom2)

    'Form both complementary beta arguments from the same log-ratio
        PROB_TF_LogisticPair LogRatio, BetaX, BetaY

'------------------------------------------------------------------------------
' COMPUTE SURVIVAL PROBABILITY
'------------------------------------------------------------------------------
    'Evaluate the reflected regularized incomplete beta function
        If Not PROB_TryBetaRegularized( _
            BetaY, _
            BetaX, _
            0.5 * DegreesFreedom2, _
            0.5 * DegreesFreedom1, _
            Value, _
            FailMsg) Then
            GoTo Fail_Num
        End If

    'Return the survival probability
        K_STATS_F_Survival = Value

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
        K_STATS_F_Survival = CVErr(xlErrNum)
    'Exit before the error handler
        Exit Function

'------------------------------------------------------------------------------
' ERROR HANDLER
'------------------------------------------------------------------------------
Err_Handler:
    'Write unexpected runtime errors to diagnostics
        PROB_SetStatus Status, "Unexpected error in K_STATS_F_Survival: " & Err.Description
    'Return worksheet value error
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
'   - PROB_IsValidProbabilityOpen
'   - PROB_TF_ValidateTwoDF
'   - PROB_TryBetaInvRegularized
'   - PROB_TryExp
'   - PROB_SetStatus
'
' UPDATED
'   2026-07-11 - House-style rewrite and numerical-contract hardening.
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

    'Validate both degree parameters and their half-degree kernel arguments
        If Not PROB_TF_ValidateTwoDF( _
            DegreesFreedom1, DegreesFreedom2, FailMsg) Then
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
'==============================================================================
' PROB_TryStudentTPDF
'------------------------------------------------------------------------------
' PURPOSE
'   Evaluates the Student t density kernel.
'
' INPUTS
'   X                Finite evaluation point.
'   DegreesFreedom   Supported positive degrees of freedom.
'
' OUTPUTS
'   Result    Density value on success.
'   FailMsg   Detailed numerical failure message on failure.
'
' RETURNS
'   Boolean
'     True  => Result contains a valid density.
'     False => The logarithmic ratio or final exponential could not be resolved.
'
' NUMERICAL METHOD
'   Uses the logarithmic density:
'
'     LogGamma((DF + 1) / 2) - LogGamma(DF / 2)
'       - 0.5 Log(DF * Pi)
'       - 0.5 (DF + 1) Log(1 + X ^ 2 / DF).
'
'   The log-gamma difference and square ratio are evaluated by dedicated stable
'   helpers. The routine never squares X directly.
'
' ERROR POLICY
'   - Predictable numerical failure is reported through False and FailMsg.
'   - Valid density underflow returns zero and True.
'
' DEPENDENCIES
'   - PROB_TF_TryLogOnePlusSquareRatio
'   - PROB_LogGammaHalfDiff
'   - PROB_TryExp
'   - PROB_PI
'
' CALLED FROM
'   - K_STATS_StudentT_Density
'   - PROB_TryStudentTInvTail
'   - PROB_TF_TryStudentTCentralMass
'
' UPDATED
'   2026-07-11 - House-style rewrite; numerical method unchanged.
'==============================================================================
'
'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Dim LogOnePlusRatio     As Double          'Log(1 + X squared / DF)
    Dim LogDensity          As Double          'Logarithm of the density

'------------------------------------------------------------------------------
' INITIALIZE OUTPUTS
'------------------------------------------------------------------------------
    'Clear outputs before attempting the calculation
        Result = 0#
        FailMsg = vbNullString

'------------------------------------------------------------------------------
' FORM LOG-DENSITY
'------------------------------------------------------------------------------
    'Evaluate the stable logarithmic square-ratio term
        If Not PROB_TF_TryLogOnePlusSquareRatio( _
            Abs(X), DegreesFreedom, LogOnePlusRatio) Then
            FailMsg = "Student t square ratio could not be represented"
            Exit Function
        End If

    'Assemble the logarithmic density
        LogDensity = _
            PROB_LogGammaHalfDiff(0.5 * DegreesFreedom) - _
            0.5 * (Log(DegreesFreedom) + Log(PROB_PI)) - _
            0.5 * (DegreesFreedom + 1#) * LogOnePlusRatio

'------------------------------------------------------------------------------
' EXPONENTIATE
'------------------------------------------------------------------------------
    'Exponentiate; valid far-tail underflow returns zero
        If Not PROB_TryExp(LogDensity, Result) Then
            FailMsg = "Student t density overflowed a Double"
            Exit Function
        End If

'------------------------------------------------------------------------------
' RETURN SUCCESS
'------------------------------------------------------------------------------
    'Report a valid density
        PROB_TryStudentTPDF = True
End Function


Private Function PROB_TryStudentTTail( _
    ByVal AbsX As Double, _
    ByVal DegreesFreedom As Double, _
    ByRef Result As Double, _
    ByRef FailMsg As String) _
    As Boolean
'
'==============================================================================
' PROB_TryStudentTTail
'------------------------------------------------------------------------------
' PURPOSE
'   Returns the Student t right-tail probability at a non-negative argument.
'
' INPUTS
'   AbsX              Absolute evaluation point; must be non-negative.
'   DegreesFreedom    Supported positive degrees of freedom.
'
' OUTPUTS
'   Result    Right-tail probability P(T > AbsX).
'   FailMsg   Detailed numerical failure message on failure.
'
' RETURNS
'   Boolean
'     True  => Result contains a valid probability.
'     False => A beta argument or special-function evaluation failed.
'
' NUMERICAL METHOD
'   - AbsX = 0: returns one half exactly.
'   - Tiny AbsX: uses a local central-mass expansion.
'   - DF = 1: uses stable central and reciprocal arctangent forms.
'   - DF = 2: uses algebraic forms that avoid an overflowing square.
'   - General DF: evaluates one half of the regularized incomplete beta.
'
' ERROR POLICY
'   - Predictable numerical failure is reported through False and FailMsg.
'   - A mathematically valid far-tail underflow returns zero and True.
'
' DEPENDENCIES
'   - PROB_TF_TryStudentTCentralMass
'   - PROB_TF_TrySquareRatioPair
'   - PROB_TryStudentTPDF
'   - PROB_TryBetaRegularized
'   - PROB_TryDivide
'   - PROB_PI
'
' CALLED FROM
'   - K_STATS_StudentT_Cumulative
'   - K_STATS_StudentT_Survival
'   - PROB_TryStudentTInvTail
'
' UPDATED
'   2026-07-11 - House-style rewrite; tiny-argument branches retained.
'==============================================================================
'
'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Dim S                   As Double          'Auxiliary square-root term
    Dim InverseX            As Double          'Reciprocal absolute argument
    Dim BetaX               As Double          'Incomplete-beta left argument
    Dim BetaY               As Double          'Incomplete-beta complement
    Dim Ibeta               As Double          'Regularized incomplete beta
    Dim CentralMass         As Double          'Probability mass from zero to AbsX

'------------------------------------------------------------------------------
' INITIALIZE OUTPUTS
'------------------------------------------------------------------------------
    'Clear outputs before attempting the calculation
        Result = 0#
        FailMsg = vbNullString

'------------------------------------------------------------------------------
' HANDLE THE CENTER
'------------------------------------------------------------------------------
    'Return the exact symmetric tail at zero
        If AbsX = 0# Then
            Result = 0.5
            PROB_TryStudentTTail = True
            Exit Function
        End If

    'Use the local expansion when it is accurate enough
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

'------------------------------------------------------------------------------
' CAUCHY CLOSED FORM
'------------------------------------------------------------------------------
    'Use the central arctangent form near zero
        If DegreesFreedom = 1# Then
            If AbsX <= 1# Then
                Result = 0.5 - Atn(AbsX) / PROB_PI

    'Use the reciprocal arctangent form in the far tail
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

'------------------------------------------------------------------------------
' TWO-DEGREE CLOSED FORM
'------------------------------------------------------------------------------
    'Use the direct form while the square is harmless
        If DegreesFreedom = 2# Then
            If AbsX <= 1# Then
                S = Sqr(2# + AbsX * AbsX)
                Result = 1# / (S * (S + AbsX))

    'Scale by the reciprocal in the far tail so AbsX squared is never formed
            Else
                InverseX = 1# / AbsX
                S = Sqr(1# + 2# * InverseX * InverseX)
                Result = InverseX * InverseX / (S * (S + 1#))
            End If

            PROB_TryStudentTTail = True
            Exit Function
        End If

'------------------------------------------------------------------------------
' GENERAL INCOMPLETE-BETA BRANCH
'------------------------------------------------------------------------------
    'Form DF / (DF + X squared) and its complement without squaring X
        If Not PROB_TF_TrySquareRatioPair( _
            AbsX, DegreesFreedom, BetaX, BetaY) Then
            FailMsg = "Student t beta arguments could not be represented"
            Exit Function
        End If

    'Evaluate the regularized incomplete beta transformation
        If Not PROB_TryBetaRegularized( _
            BetaX, _
            BetaY, _
            0.5 * DegreesFreedom, _
            0.5, _
            Ibeta, _
            FailMsg) Then
            Exit Function
        End If

    'Convert the beta result into the one-sided Student t tail
        Result = 0.5 * Ibeta

'------------------------------------------------------------------------------
' RETURN SUCCESS
'------------------------------------------------------------------------------
    'Report a valid probability
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
'   Solves P(T > X) = Tail for the non-negative Student t quantile X.
'
' INPUTS
'   Tail              Right-tail probability in the interval (0, 0.5].
'   DegreesFreedom    Supported positive degrees of freedom.
'
' OUTPUTS
'   Result    Non-negative quantile on success.
'   FailMsg   Detailed numerical failure message on failure.
'
' RETURNS
'   Boolean
'     True  => Result contains a valid quantile.
'     False => The quantile overflowed, a kernel failed or iteration did not
'              converge.
'
' NUMERICAL METHOD
'   - DF = 1: exact Cauchy inverse with a guarded reciprocal.
'   - DF = 2: exact logarithmic closed form.
'   - DF < 0.5: direct incomplete-beta inversion.
'   - General DF: Cornish-Fisher seed followed by safeguarded Newton iteration.
'   - The bracket midpoint uses Low + 0.5 * (High - Low), avoiding overflow in
'     Low + High.
'   - No artificial quantile cap is imposed. Guarded arithmetic detects the
'     actual Double boundary.
'
' ERROR POLICY
'   - Predictable numerical failure is reported through False and FailMsg.
'   - Non-convergence never returns a partial answer.
'
' DEPENDENCIES
'   - PROB_NormalInvCDFRaw
'   - PROB_TryStudentTTail
'   - PROB_TryStudentTPDF
'   - PROB_TF_TryStudentTInvTailSmallDF
'   - PROB_TryMultiply
'   - PROB_TryDivide
'   - PROB_TryAdd
'   - PROB_TryExp
'   - PROB_Log1p
'   - PROB_IsFinite
'   - PROB_EPS
'   - PROB_PI
'
' CALLED FROM
'   - K_STATS_StudentT_InverseCumulative
'
' UPDATED
'   2026-07-11 - Removed artificial search ceiling; stable midpoint added.
'==============================================================================
'
'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Dim Z                   As Double          'Upper-tail standard-normal seed
    Dim Z2                  As Double          'Squared normal seed
    Dim X                   As Double          'Current iterate
    Dim XNew                As Double          'Candidate next iterate
    Dim Low                 As Double          'Lower quantile bracket
    Dim High                As Double          'Upper quantile bracket
    Dim HasHigh             As Boolean         'Whether an upper bracket exists
    Dim TailAtX             As Double          'Computed tail at current iterate
    Dim Residual            As Double          'Computed tail minus target tail
    Dim Density             As Double          'Student t density at current iterate
    Dim NewtonStep          As Double          'Residual divided by density
    Dim TanValue            As Double          'Cauchy tangent denominator
    Dim Numerator           As Double          'DF = 2 closed-form numerator
    Dim LogQuantile         As Double          'Logarithm of a closed-form quantile
    Dim Converged           As Boolean         'Iteration convergence flag
    Dim IterIdx             As Long            'Iteration counter

'------------------------------------------------------------------------------
' INITIALIZE OUTPUTS
'------------------------------------------------------------------------------
    'Clear outputs before attempting the calculation
        Result = 0#
        FailMsg = vbNullString

'------------------------------------------------------------------------------
' HANDLE THE MEDIAN
'------------------------------------------------------------------------------
    'Return zero for a tail at or above one half
        If Tail >= 0.5 Then
            Result = 0#
            PROB_TryStudentTInvTail = True
            Exit Function
        End If

'------------------------------------------------------------------------------
' CAUCHY CLOSED FORM
'------------------------------------------------------------------------------
    'Use the central tangent form when it is well conditioned
        If DegreesFreedom = 1# Then
            If Tail > 0.25 Then
                Result = Tan(PROB_PI * (0.5 - Tail))

    'Use the reciprocal tangent form in the far tail
            Else
                TanValue = Tan(PROB_PI * Tail)

                If Not PROB_TryDivide(1#, TanValue, Result) Then
                    FailMsg = _
                        "Student t quantile overflowed a Double for " & _
                        "DegreesFreedom = 1"
                    Exit Function
                End If
            End If

            PROB_TryStudentTInvTail = True
            Exit Function
        End If

'------------------------------------------------------------------------------
' TWO-DEGREE CLOSED FORM
'------------------------------------------------------------------------------
    'Evaluate the exact two-degree inverse in the logarithmic domain
        If DegreesFreedom = 2# Then
            Numerator = 1# - 2# * Tail

            If Numerator <= 0# Then
                Result = 0#
                PROB_TryStudentTInvTail = True
                Exit Function
            End If

            LogQuantile = _
                Log(Numerator) - _
                0.5 * ( _
                    Log(2#) + _
                    Log(Tail) + _
                    PROB_Log1p(-Tail))

            If Not PROB_TryExp(LogQuantile, Result) Then
                FailMsg = _
                    "Student t quantile overflowed a Double for " & _
                    "DegreesFreedom = 2"
                Exit Function
            End If

            PROB_TryStudentTInvTail = True
            Exit Function
        End If

'------------------------------------------------------------------------------
' SMALL-DEGREE BETA-INVERSE BRANCH
'------------------------------------------------------------------------------
    'Avoid singular Cornish-Fisher powers when DegreesFreedom is below one half
        If DegreesFreedom < 0.5 Then
            If Not PROB_TF_TryStudentTInvTailSmallDF( _
                Tail, DegreesFreedom, Result, FailMsg) Then
                Exit Function
            End If

            PROB_TryStudentTInvTail = True
            Exit Function
        End If

'------------------------------------------------------------------------------
' CORNISH-FISHER SEED
'------------------------------------------------------------------------------
    'Map the right-tail probability into a positive normal quantile
        Z = -PROB_NormalInvCDFRaw(Tail)
        Z2 = Z * Z

    'Build the third-order Cornish-Fisher approximation
        X = _
            Z + _
            (Z2 * Z + Z) / (4# * DegreesFreedom) + _
            ( _
                5# * Z2 * Z2 * Z + _
                16# * Z2 * Z + _
                3# * Z _
            ) / (96# * DegreesFreedom * DegreesFreedom) + _
            ( _
                3# * Z2 * Z2 * Z2 * Z + _
                19# * Z2 * Z2 * Z + _
                17# * Z2 * Z - _
                15# * Z _
            ) / ( _
                384# * _
                DegreesFreedom * _
                DegreesFreedom * _
                DegreesFreedom)

    'Use a safe positive fallback if the approximation is unusable
        If X <= 0# Or Not PROB_IsFinite(X) Then X = 1#

'------------------------------------------------------------------------------
' INITIALIZE BRACKET
'------------------------------------------------------------------------------
    'Start with the support lower bound and no known upper bracket
        Low = 0#
        High = 0#
        HasHigh = False
        Converged = False

'------------------------------------------------------------------------------
' SAFEGUARDED NEWTON ITERATION
'------------------------------------------------------------------------------
        For IterIdx = 1 To PROB_T_INV_MAX_ITER
            'Evaluate the tail at the current iterate
                If Not PROB_TryStudentTTail( _
                    X, DegreesFreedom, TailAtX, FailMsg) Then
                    Exit Function
                End If

            'Form the tail residual
                Residual = TailAtX - Tail

            'Update the monotone bracket
                If Residual > 0# Then
                    If X > Low Then Low = X
                Else
                    If (Not HasHigh) Or X < High Then High = X
                    HasHigh = True
                End If

            'Evaluate the analytic derivative magnitude
                If Not PROB_TryStudentTPDF( _
                    X, DegreesFreedom, Density, FailMsg) Then
                    Exit Function
                End If

            'Use bisection or bracket expansion when the density is unavailable
                If Density <= 0# Then
                    If HasHigh Then
                        XNew = Low + 0.5 * (High - Low)
                    ElseIf Not PROB_TryMultiply(2#, X, XNew) Then
                        FailMsg = _
                            "Student t inverse bracket expansion " & _
                            "overflowed a Double"
                        Exit Function
                    End If

            'Use bisection or expansion when the Newton division overflows
                ElseIf Not PROB_TryDivide( _
                    Residual, Density, NewtonStep) Then

                    If HasHigh Then
                        XNew = Low + 0.5 * (High - Low)
                    ElseIf Not PROB_TryMultiply(2#, X, XNew) Then
                        FailMsg = _
                            "Student t inverse Newton step overflowed a Double"
                        Exit Function
                    End If

            'Use bisection or expansion when the Newton addition overflows
                ElseIf Not PROB_TryAdd(X, NewtonStep, XNew) Then
                    If HasHigh Then
                        XNew = Low + 0.5 * (High - Low)
                    ElseIf Not PROB_TryMultiply(2#, X, XNew) Then
                        FailMsg = _
                            "Student t inverse iterate overflowed a Double"
                        Exit Function
                    End If

            'Safeguard a valid Newton candidate against the current bracket
                Else
                    If HasHigh Then
                        If XNew <= Low Or XNew >= High Then
                            XNew = Low + 0.5 * (High - Low)
                        End If
                    ElseIf XNew <= Low Then
                        If Not PROB_TryMultiply(2#, X, XNew) Then
                            FailMsg = _
                                "Student t inverse bracket expansion " & _
                                "overflowed a Double"
                            Exit Function
                        End If
                    End If
                End If

            'Accept a step that cannot move at the current Double resolution
                If Abs(XNew - X) <= PROB_EPS * Abs(XNew) Or XNew = X Then
                    X = XNew
                    Converged = True
                    Exit For
                End If

            'Advance to the next iterate
                X = XNew
        Next IterIdx

'------------------------------------------------------------------------------
' CHECK CONVERGENCE
'------------------------------------------------------------------------------
    'Reject a solver that exhausted the iteration budget
        If Not Converged Then
            FailMsg = _
                "Student t inverse failed to converge in " & _
                PROB_T_INV_MAX_ITER & _
                " iterations for DegreesFreedom = " & _
                DegreesFreedom
            Exit Function
        End If

'------------------------------------------------------------------------------
' RETURN SUCCESS
'------------------------------------------------------------------------------
    'Return the converged non-negative quantile
        Result = X
        PROB_TryStudentTInvTail = True
End Function


'==============================================================================
' PRIVATE LOGARITHMIC AND RATIO HELPERS
'==============================================================================

Private Function PROB_TF_LogOnePlusExp( _
    ByVal LogRatio As Double) _
    As Double
'
'==============================================================================
' PROB_TF_LogOnePlusExp
'------------------------------------------------------------------------------
' PURPOSE
'   Returns Log(1 + Exp(LogRatio)) without overflow or cancellation.
'
' INPUTS
'   LogRatio  Finite logarithmic ratio.
'
' RETURNS
'   Double
'     Stable softplus value.
'
' NUMERICAL METHOD
'   - Positive LogRatio:
'       LogRatio + Log1p(Exp(-LogRatio)).
'   - Non-positive LogRatio:
'       Log1p(Exp(LogRatio)).
'
' DEPENDENCIES
'   - PROB_TryExp
'   - PROB_Log1p
'
' CALLED FROM
'   - K_STATS_F_Density
'   - PROB_TF_TryLogOnePlusSquareRatio
'
' UPDATED
'   2026-07-11 - House-style rewrite; numerical method unchanged.
'==============================================================================
'
'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Dim SmallTerm           As Double          'Exponentiated small-side term

'------------------------------------------------------------------------------
' COMPUTE SOFTPLUS
'------------------------------------------------------------------------------
    'Use the complemented form for a positive logarithmic ratio
        If LogRatio > 0# Then
            If Not PROB_TryExp(-LogRatio, SmallTerm) Then
                SmallTerm = 0#
            End If

            PROB_TF_LogOnePlusExp = _
                LogRatio + PROB_Log1p(SmallTerm)

    'Use the direct form for a non-positive logarithmic ratio
        Else
            If Not PROB_TryExp(LogRatio, SmallTerm) Then
                SmallTerm = 0#
            End If

            PROB_TF_LogOnePlusExp = PROB_Log1p(SmallTerm)
        End If
End Function


Private Sub PROB_TF_LogisticPair( _
    ByVal LogRatio As Double, _
    ByRef LeftValue As Double, _
    ByRef RightValue As Double)
'
'==============================================================================
' PROB_TF_LogisticPair
'------------------------------------------------------------------------------
' PURPOSE
'   Forms r / (1 + r) and 1 / (1 + r) directly from Log(r).
'
' INPUTS
'   LogRatio  Finite logarithm of a positive ratio r.
'
' OUTPUTS
'   LeftValue   r / (1 + r).
'   RightValue  1 / (1 + r).
'
' NUMERICAL METHOD
'   Exponentiates only the non-positive side of the ratio. The routine therefore
'   never forms r when r lies outside the representable Double range.
'
' DEPENDENCIES
'   - PROB_TryExp
'
' CALLED FROM
'   - K_STATS_F_Cumulative
'   - K_STATS_F_Survival
'   - PROB_TF_TrySquareRatioPair
'
' UPDATED
'   2026-07-11 - House-style rewrite; numerical method unchanged.
'==============================================================================
'
'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Dim SmallTerm           As Double          'Exp of the non-positive log-ratio
    Dim Denominator         As Double          'One plus SmallTerm

'------------------------------------------------------------------------------
' COMPUTE COMPLEMENTARY PAIR
'------------------------------------------------------------------------------
    'Use Exp(-LogRatio) when the original ratio is at least one
        If LogRatio >= 0# Then
            If Not PROB_TryExp(-LogRatio, SmallTerm) Then
                SmallTerm = 0#
            End If

            Denominator = 1# + SmallTerm
            LeftValue = 1# / Denominator
            RightValue = SmallTerm / Denominator

    'Use Exp(LogRatio) when the original ratio is below one
        Else
            If Not PROB_TryExp(LogRatio, SmallTerm) Then
                SmallTerm = 0#
            End If

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
' PROB_TF_TryLogOnePlusSquareRatio
'------------------------------------------------------------------------------
' PURPOSE
'   Returns Log(1 + AbsX ^ 2 / DegreesFreedom) without squaring AbsX.
'
' INPUTS
'   AbsX              Non-negative absolute evaluation point.
'   DegreesFreedom    Supported positive degrees of freedom.
'
' OUTPUTS
'   Result  Stable logarithmic ratio term.
'
' RETURNS
'   Boolean
'     True when Result is finite; otherwise False.
'
' NUMERICAL METHOD
'   Forms LogRatio = 2 Log(AbsX) - Log(DegreesFreedom), then evaluates the
'   softplus through PROB_TF_LogOnePlusExp.
'
' DEPENDENCIES
'   - PROB_TF_LogOnePlusExp
'   - PROB_IsFinite
'
' CALLED FROM
'   - PROB_TryStudentTPDF
'
' UPDATED
'   2026-07-11 - House-style rewrite; direct squaring remains eliminated.
'==============================================================================
'
'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Dim LogRatio            As Double          'Log(AbsX squared / DF)

'------------------------------------------------------------------------------
' HANDLE ZERO
'------------------------------------------------------------------------------
    'Return Log(1) exactly at the center
        If AbsX = 0# Then
            Result = 0#
            PROB_TF_TryLogOnePlusSquareRatio = True
            Exit Function
        End If

'------------------------------------------------------------------------------
' COMPUTE LOGARITHMIC RATIO
'------------------------------------------------------------------------------
    'Form the logarithmic square ratio without squaring AbsX
        LogRatio = 2# * Log(AbsX) - Log(DegreesFreedom)

    'Evaluate the stable softplus
        Result = PROB_TF_LogOnePlusExp(LogRatio)

'------------------------------------------------------------------------------
' RETURN STATUS
'------------------------------------------------------------------------------
    'Report whether the result is finite
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
' PROB_TF_TrySquareRatioPair
'------------------------------------------------------------------------------
' PURPOSE
'   Forms DF / (DF + AbsX ^ 2) and AbsX ^ 2 / (DF + AbsX ^ 2) without
'   squaring AbsX.
'
' INPUTS
'   AbsX              Non-negative absolute evaluation point.
'   DegreesFreedom    Supported positive degrees of freedom.
'
' OUTPUTS
'   BetaX  DF / (DF + AbsX ^ 2).
'   BetaY  AbsX ^ 2 / (DF + AbsX ^ 2).
'
' RETURNS
'   Boolean
'     True after the complementary pair has been formed.
'
' DEPENDENCIES
'   - PROB_TF_LogisticPair
'
' CALLED FROM
'   - PROB_TryStudentTTail
'
' UPDATED
'   2026-07-11 - House-style rewrite; direct squaring remains eliminated.
'==============================================================================
'
'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Dim LogRatio            As Double          'Log(AbsX squared / DF)

'------------------------------------------------------------------------------
' HANDLE ZERO
'------------------------------------------------------------------------------
    'Return the exact complementary pair at the center
        If AbsX = 0# Then
            BetaX = 1#
            BetaY = 0#
            PROB_TF_TrySquareRatioPair = True
            Exit Function
        End If

'------------------------------------------------------------------------------
' COMPUTE COMPLEMENTARY PAIR
'------------------------------------------------------------------------------
    'Form the logarithmic square ratio without squaring AbsX
        LogRatio = 2# * Log(AbsX) - Log(DegreesFreedom)

    'Map the square ratio into the complementary beta arguments
        PROB_TF_LogisticPair LogRatio, BetaY, BetaX

'------------------------------------------------------------------------------
' RETURN SUCCESS
'------------------------------------------------------------------------------
    'Report a valid complementary pair
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
' PROB_TF_TryStudentTCentralMass
'------------------------------------------------------------------------------
' PURPOSE
'   Evaluates the Student t probability mass from zero to AbsX when a local
'   odd-power expansion is appropriate.
'
' INPUTS
'   AbsX              Strictly positive absolute evaluation point.
'   DegreesFreedom    Supported positive degrees of freedom.
'
' OUTPUTS
'   Result
'     Non-negative central mass when the local expansion is used.
'     Negative sentinel when the caller should use the general beta path.
'
'   FailMsg
'     Detailed numerical failure message on failure.
'
' RETURNS
'   Boolean
'     True  => Result contains either the central mass or the negative sentinel.
'     False => The density or final product could not be resolved.
'
' NUMERICAL METHOD
'   Uses Density(0) * AbsX multiplied by the first correction terms of the local
'   odd-power integral. The expansion is selected only when its curvature proxy
'   is no greater than 1E-8.
'
' DEPENDENCIES
'   - PROB_TryStudentTPDF
'   - PROB_TryMultiply
'   - PROB_TryExp
'   - PROB_Log1p
'
' CALLED FROM
'   - PROB_TryStudentTTail
'
' NOTES
'   A negative Result is a private control sentinel, not a probability and not a
'   public error result.
'
' UPDATED
'   2026-07-11 - House-style rewrite; local expansion unchanged.
'==============================================================================
'
'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Dim LogCurvature        As Double          'Logarithm of the curvature proxy
    Dim Curvature           As Double          'Curvature proxy
    Dim DensityZero         As Double          'Student t density at zero
    Dim Correction          As Double          'Local integral correction factor

'------------------------------------------------------------------------------
' INITIALIZE OUTPUTS
'------------------------------------------------------------------------------
    'Use a negative sentinel until the local expansion is selected
        Result = -1#
        FailMsg = vbNullString

'------------------------------------------------------------------------------
' SELECT LOCAL EXPANSION
'------------------------------------------------------------------------------
    'Form the logarithm of X squared times (DF + 1) / DF
        LogCurvature = _
            2# * Log(AbsX) + _
            PROB_Log1p(DegreesFreedom) - _
            Log(DegreesFreedom)

    'Return the sentinel when the local expansion would not be accurate enough
        If LogCurvature > Log(0.00000001) Then
            PROB_TF_TryStudentTCentralMass = True
            Exit Function
        End If

'------------------------------------------------------------------------------
' COMPUTE CENTRAL MASS
'------------------------------------------------------------------------------
    'Recover the small curvature proxy
        If Not PROB_TryExp(LogCurvature, Curvature) Then
            Curvature = 0#
        End If

    'Evaluate the density at the center
        If Not PROB_TryStudentTPDF( _
            0#, DegreesFreedom, DensityZero, FailMsg) Then
            Exit Function
        End If

    'Build the local integral correction
        Correction = _
            1# - _
            Curvature / 6# + _
            Curvature * Curvature * _
                (DegreesFreedom + 3#) / _
                (40# * (DegreesFreedom + 1#))

    'Form the leading density-times-width term under the arithmetic contract
        If Not PROB_TryMultiply(DensityZero, AbsX, Result) Then
            FailMsg = "Student t central probability overflowed"
            Exit Function
        End If

    'Apply the bounded correction factor
        Result = Result * Correction

'------------------------------------------------------------------------------
' RETURN SUCCESS
'------------------------------------------------------------------------------
    'Report either the local result or the negative sentinel
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
' PROB_TF_TryStudentTInvTailSmallDF
'------------------------------------------------------------------------------
' PURPOSE
'   Inverts the Student t right tail for very small degrees of freedom through
'   the incomplete-beta transformation.
'
' INPUTS
'   Tail              Right-tail probability in (0, 0.5).
'   DegreesFreedom    Positive degrees of freedom below 0.5.
'
' OUTPUTS
'   Result    Non-negative Student t quantile.
'   FailMsg   Detailed numerical failure message on failure.
'
' RETURNS
'   Boolean
'     True  => Result contains a valid quantile.
'     False => Beta inversion or final exponentiation failed.
'
' WHY THIS EXISTS
'   The general Cornish-Fisher seed contains powers of 1 / DegreesFreedom and is
'   unsuitable for very small degrees of freedom. Direct beta inversion avoids
'   those singular intermediate expressions.
'
' DEPENDENCIES
'   - PROB_TryBetaInvRegularized
'   - PROB_TryExp
'
' CALLED FROM
'   - PROB_TryStudentTInvTail
'
' UPDATED
'   2026-07-11 - House-style rewrite; small-DF branch unchanged.
'==============================================================================
'
'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Dim BetaProbability     As Double          'Incomplete-beta target
    Dim BetaComplement      As Double          'Complementary beta target
    Dim BetaX               As Double          'Incomplete-beta root
    Dim BetaY               As Double          'Complementary root
    Dim LogQuantile         As Double          'Logarithm of the Student t quantile

'------------------------------------------------------------------------------
' INITIALIZE OUTPUTS
'------------------------------------------------------------------------------
    'Clear outputs before attempting the calculation
        Result = 0#
        FailMsg = vbNullString

'------------------------------------------------------------------------------
' INVERT BETA TRANSFORMATION
'------------------------------------------------------------------------------
    'Map the one-sided t tail into the regularized-beta target pair
        BetaProbability = 2# * Tail
        BetaComplement = 2# * (0.5 - Tail)

    'Invert the regularized incomplete beta function
        If Not PROB_TryBetaInvRegularized( _
            BetaProbability, _
            BetaComplement, _
            0.5 * DegreesFreedom, _
            0.5, _
            BetaX, _
            BetaY, _
            FailMsg) Then
            Exit Function
        End If

'------------------------------------------------------------------------------
' MAP ROOT BACK TO STUDENT T
'------------------------------------------------------------------------------
    'A zero beta root implies a quantile beyond the Double range
        If BetaX <= 0# Then
            FailMsg = _
                "Student t quantile overflowed for very small " & _
                "DegreesFreedom"
            Exit Function
        End If

    'A zero complement corresponds to the center
        If BetaY <= 0# Then
            Result = 0#
            PROB_TF_TryStudentTInvTailSmallDF = True
            Exit Function
        End If

    'Assemble the Student t quantile in the logarithmic domain
        LogQuantile = _
            0.5 * ( _
                Log(DegreesFreedom) + _
                Log(BetaY) - _
                Log(BetaX))

    'Exponentiate under the shared numerical contract
        If Not PROB_TryExp(LogQuantile, Result) Then
            FailMsg = _
                "Student t quantile overflowed for very small " & _
                "DegreesFreedom"
            Exit Function
        End If

'------------------------------------------------------------------------------
' RETURN SUCCESS
'------------------------------------------------------------------------------
    'Report a valid quantile
        PROB_TF_TryStudentTInvTailSmallDF = True
End Function


'==============================================================================
' PRIVATE VALIDATION HELPERS
'==============================================================================

Private Function PROB_TF_ValidateDF( _
    ByVal DegreesFreedom As Double, _
    ByVal DFName As String, _
    ByRef FailMsg As String) _
    As Boolean
'
'==============================================================================
' PROB_TF_ValidateDF
'------------------------------------------------------------------------------
' PURPOSE
'   Validates one degree-of-freedom parameter and its half-degree argument.
'
' INPUTS
'   DegreesFreedom  Degree parameter to validate.
'   DFName          Parameter name used in diagnostics.
'
' OUTPUTS
'   FailMsg  Detailed validation failure message.
'
' RETURNS
'   Boolean
'     True when DegreesFreedom is positive, finite, within the supported
'     special-function magnitude and large enough that DegreesFreedom / 2 does
'     not underflow to zero.
'
' DEPENDENCIES
'   - PROB_IsPositiveWithinSupportedMagnitude
'
' CALLED FROM
'   - Public Student t, chi-square and F wrappers
'   - PROB_TF_ValidateTwoDF
'   - PROB_TF_ValidateXAndDF
'
' UPDATED
'   2026-07-11 - Centralized degree validation and diagnostics.
'==============================================================================
'
'------------------------------------------------------------------------------
' INITIALIZE OUTPUT
'------------------------------------------------------------------------------
    'Clear the failure message before validation
        FailMsg = vbNullString

'------------------------------------------------------------------------------
' VALIDATE DEGREE PARAMETER
'------------------------------------------------------------------------------
    'Require a supported positive finite value
        If Not PROB_IsPositiveWithinSupportedMagnitude(DegreesFreedom) Then
            FailMsg = _
                DFName & _
                " must be a supported finite strictly positive number"
            Exit Function
        End If

    'Require a positive representable half-degree kernel parameter
        If 0.5 * DegreesFreedom <= 0# Then
            FailMsg = _
                DFName & _
                " is too small for the half-degree special-function parameter"
            Exit Function
        End If

'------------------------------------------------------------------------------
' RETURN SUCCESS
'------------------------------------------------------------------------------
    'Report a valid degree parameter
        PROB_TF_ValidateDF = True
End Function


Private Function PROB_TF_ValidateTwoDF( _
    ByVal DegreesFreedom1 As Double, _
    ByVal DegreesFreedom2 As Double, _
    ByRef FailMsg As String) _
    As Boolean
'
'==============================================================================
' PROB_TF_ValidateTwoDF
'------------------------------------------------------------------------------
' PURPOSE
'   Validates the numerator and denominator degree parameters of the F family.
'
' INPUTS
'   DegreesFreedom1  Numerator degrees of freedom.
'   DegreesFreedom2  Denominator degrees of freedom.
'
' OUTPUTS
'   FailMsg  Detailed validation failure message.
'
' RETURNS
'   Boolean
'     True when both degree parameters pass PROB_TF_ValidateDF.
'
' DEPENDENCIES
'   - PROB_TF_ValidateDF
'
' CALLED FROM
'   - K_STATS_F_InverseCumulative
'   - PROB_TF_ValidateXAndTwoDF
'
' UPDATED
'   2026-07-11 - Added shared two-degree validator.
'==============================================================================
'
'------------------------------------------------------------------------------
' VALIDATE DEGREE PARAMETERS
'------------------------------------------------------------------------------
    'Validate numerator degrees of freedom
        If Not PROB_TF_ValidateDF( _
            DegreesFreedom1, "DegreesFreedom1", FailMsg) Then
            Exit Function
        End If

    'Validate denominator degrees of freedom
        If Not PROB_TF_ValidateDF( _
            DegreesFreedom2, "DegreesFreedom2", FailMsg) Then
            Exit Function
        End If

'------------------------------------------------------------------------------
' RETURN SUCCESS
'------------------------------------------------------------------------------
    'Report two valid degree parameters
        PROB_TF_ValidateTwoDF = True
End Function


Private Function PROB_TF_ValidateXAndDF( _
    ByVal X As Double, _
    ByVal DegreesFreedom As Double, _
    ByRef FailMsg As String, _
    ByVal DFName As String) _
    As Boolean
'
'==============================================================================
' PROB_TF_ValidateXAndDF
'------------------------------------------------------------------------------
' PURPOSE
'   Validates a finite evaluation point and one degree parameter.
'
' INPUTS
'   X                 Evaluation point.
'   DegreesFreedom    Degree parameter.
'   DFName            Parameter name used in diagnostics.
'
' OUTPUTS
'   FailMsg  Detailed validation failure message.
'
' RETURNS
'   Boolean
'     True when X is finite and DegreesFreedom passes PROB_TF_ValidateDF.
'
' DEPENDENCIES
'   - PROB_IsFinite
'   - PROB_TF_ValidateDF
'
' CALLED FROM
'   - Student t density, cumulative and survival wrappers
'   - Chi-square density, cumulative and survival wrappers
'
' UPDATED
'   2026-07-11 - House-style rewrite and validator centralization.
'==============================================================================
'
'------------------------------------------------------------------------------
' INITIALIZE OUTPUT
'------------------------------------------------------------------------------
    'Clear the failure message before validation
        FailMsg = vbNullString

'------------------------------------------------------------------------------
' VALIDATE INPUTS
'------------------------------------------------------------------------------
    'Require a finite evaluation point
        If Not PROB_IsFinite(X) Then
            FailMsg = "X must be a finite number"
            Exit Function
        End If

    'Validate the degree parameter
        If Not PROB_TF_ValidateDF( _
            DegreesFreedom, DFName, FailMsg) Then
            Exit Function
        End If

'------------------------------------------------------------------------------
' RETURN SUCCESS
'------------------------------------------------------------------------------
    'Report valid inputs
        PROB_TF_ValidateXAndDF = True
End Function


Private Function PROB_TF_ValidateXAndTwoDF( _
    ByVal X As Double, _
    ByVal DegreesFreedom1 As Double, _
    ByVal DegreesFreedom2 As Double, _
    ByRef FailMsg As String) _
    As Boolean
'
'==============================================================================
' PROB_TF_ValidateXAndTwoDF
'------------------------------------------------------------------------------
' PURPOSE
'   Validates a finite F evaluation point and two degree parameters.
'
' INPUTS
'   X                 Evaluation point.
'   DegreesFreedom1   Numerator degrees of freedom.
'   DegreesFreedom2   Denominator degrees of freedom.
'
' OUTPUTS
'   FailMsg  Detailed validation failure message.
'
' RETURNS
'   Boolean
'     True when X is finite and both degree parameters pass validation.
'
' DEPENDENCIES
'   - PROB_IsFinite
'   - PROB_TF_ValidateTwoDF
'
' CALLED FROM
'   - K_STATS_F_Density
'   - K_STATS_F_Cumulative
'   - K_STATS_F_Survival
'
' UPDATED
'   2026-07-11 - House-style rewrite and validator centralization.
'==============================================================================
'
'------------------------------------------------------------------------------
' INITIALIZE OUTPUT
'------------------------------------------------------------------------------
    'Clear the failure message before validation
        FailMsg = vbNullString

'------------------------------------------------------------------------------
' VALIDATE INPUTS
'------------------------------------------------------------------------------
    'Require a finite evaluation point
        If Not PROB_IsFinite(X) Then
            FailMsg = "X must be a finite number"
            Exit Function
        End If

    'Validate both degree parameters
        If Not PROB_TF_ValidateTwoDF( _
            DegreesFreedom1, DegreesFreedom2, FailMsg) Then
            Exit Function
        End If

'------------------------------------------------------------------------------
' RETURN SUCCESS
'------------------------------------------------------------------------------
    'Report valid inputs
        PROB_TF_ValidateXAndTwoDF = True
End Function



