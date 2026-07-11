Attribute VB_Name = "M_STATS_PROBDIST_CONTINUOUS"

Option Explicit

'==============================================================================
' M_STATS_PROBDIST_CONTINUOUS
'------------------------------------------------------------------------------
' PURPOSE
'   Worksheet-facing distribution functions for the five core continuous
'   distributions that are not in the normal family: Gamma, Beta, Exponential,
'   Weibull and the continuous Uniform. For each: density, cumulative, survival
'   and inverse-cumulative, plus arithmetic moments (Mean, Variance, StdDev) for
'   the three whose moments are not one-liners (Gamma, Beta, Weibull).
'
' WHY THIS EXISTS
'   These five, together with the normal and t families already written, cover
'   essentially every continuous distribution an actuarial, reliability or
'   risk model reaches for. Each one here is a thin wrapper over a kernel that
'   already exists and is already tested in M_STATS_PROBDIST_SPECIALFUNCS:
'     - Gamma and Chi-square share the regularized incomplete gamma P(a, x).
'     - Beta and F share the regularized incomplete beta I_x(a, b).
'     - Exponential and Weibull are closed forms built on Exp, PROB_Expm1 and
'       PROB_Log1p, with no iteration at all.
'   Writing them as wrappers means the hard numerics are neither duplicated nor
'   re-verified; only the parameter marshalling is new, and that is exactly what
'   the cross-family identities in the test harness exist to check.
'
' PARAMETERISATION
'   Every distribution matches its Excel worksheet counterpart argument-for-
'   argument, even where that makes the library internally inconsistent:
'     - Gamma(X, Shape, ScaleParam)          -> GAMMA.DIST(X, Shape, ScaleParam, .)
'     - Beta(X, Alpha, Beta)            -> BETA.DIST(X, Alpha, Beta, .)
'     - Exponential(X, Lambda)          -> EXPON.DIST(X, Lambda, .)   [Lambda = RATE]
'     - Weibull(X, Shape, ScaleParam)        -> WEIBULL.DIST(X, Shape, ScaleParam, .)
'     - Uniform(X, LowerBound, UpperBound)
'   Note the deliberate inconsistency: Gamma and Weibull take a SCALE, whereas
'   Exponential takes a RATE. This is why the Exponential-to-Gamma identity
'   carries a reciprocal: Exponential(Lambda) is Gamma(Shape = 1, ScaleParam = 1 /
'   Lambda), NOT Gamma(1, Lambda). The reciprocal is written out explicitly
'   wherever it appears so it can never be mistaken for a bug. Agreeing with the
'   worksheet was judged more important than agreeing with ourselves.
'
' ALGORITHM PROVENANCE
'   - Gamma CDF / survival:
'       Regularized incomplete gamma, P(Shape, X / ScaleParam) and Q(Shape, X / ScaleParam),
'       through PROB_TryGammaRegularizedP / Q. Quantile via PROB_TryGammaInvP
'       (Wilson-Hilferty seed, safeguarded Newton), rescaled by ScaleParam.
'   - Beta CDF / survival:
'       Regularized incomplete beta. The survival is computed as the swapped
'       incomplete beta I_(1-X)(B, A) rather than 1 - I_X(A, B), so the upper
'       tail never loses precision to a subtraction from one. Quantile via
'       PROB_TryBetaInvRegularized.
'   - Exponential CDF:
'       1 - Exp(-Lambda * X) evaluated as -PROB_Expm1(-Lambda * X) so the left
'       tail keeps full relative precision. Quantile -PROB_Log1p(-P) / Lambda,
'       the mirror trick on the P -> 0 side. No iteration.
'   - Weibull CDF:
'       1 - Exp(-(X / ScaleParam) ^ Shape) as -PROB_Expm1(-(X / ScaleParam) ^ Shape).
'       Quantile ScaleParam * (-PROB_Log1p(-P)) ^ (1 / Shape). No iteration.
'   - Uniform:
'       Exact closed forms throughout.
'   - Weibull moments:
'       Mean = ScaleParam * Gamma(1 + 1 / Shape), Variance = ScaleParam^2 * (Gamma(1 + 2 /
'       Shape) - Gamma(1 + 1 / Shape)^2), through Exp(PROB_LogGamma(.)) with an
'       explicit overflow guard on the Gamma-function evaluation.
'   The Try contract, the (X, Y) argument pairing and the PROB_Expm1 / PROB_Log1p
'   left-tail treatment are the local contribution; the underlying kernels are
'   the same published algorithms used by the t family.
'
' DESIGN PRINCIPLES
'   - Public worksheet-facing functions return Variant so they can return CVErr.
'   - Private numerical work goes through the shared kernels; no continued
'     fraction or series is re-implemented here.
'   - Invalid domains fail explicitly; they are not silently repaired.
'   - A density pole is not a representable number. Where a density diverges to
'     positive infinity (Gamma or Weibull at X = 0 with Shape < 1, Beta at an
'     endpoint with the corresponding shape < 1) the function returns
'     CVErr(xlErrNum), the same contract overflow uses. Underflow of a density to
'     zero remains a valid zero.
'   - CDF functions return mathematically meaningful support-edge values:
'     Cumulative below the support is 0, above it is 1.
'   - Non-convergence or overflow returns CVErr(xlErrNum).
'   - No MsgBox is raised by any public worksheet-facing function.
'
' ERROR POLICY
'   - Invalid numeric domains return CVErr(xlErrNum).
'   - Density poles, non-convergence and overflow return CVErr(xlErrNum).
'   - Unexpected runtime errors return CVErr(xlErrValue).
'   - Detailed diagnostic messages are written to Status.
'   - Application.StatusBar is not written by default.
'
' DEPENDENCIES
'   - M_STATS_PROBDIST_CORE
'       PROB_IsFinite, PROB_IsPositiveFinite, PROB_IsValidProbabilityOpen,
'       PROB_TryExp, PROB_Log1p, PROB_Expm1, PROB_SetStatus,
'       PROB_MAX_EXP, PROB_DOUBLE_MAX
'   - M_STATS_PROBDIST_SPECIALFUNCS
'       PROB_LogGamma, PROB_TryGammaRegularizedP, PROB_TryGammaRegularizedQ,
'       PROB_TryGammaInvP, PROB_LogBeta, PROB_TryBetaRegularized,
'       PROB_TryBetaInvRegularized
'
' PUBLIC SURFACE (29 UDFs)
'   Gamma        Density Cumulative Survival InverseCumulative Mean Variance StdDev
'   Beta         Density Cumulative Survival InverseCumulative Mean Variance StdDev
'   Exponential  Density Cumulative Survival InverseCumulative
'   Weibull      Density Cumulative Survival InverseCumulative Mean Variance StdDev
'   Uniform      Density Cumulative Survival InverseCumulative
'
' UPDATED
'   2026-07-11
'==============================================================================


'==============================================================================
' GAMMA DISTRIBUTION
'==============================================================================

Public Function K_STATS_Gamma_Density( _
    ByVal X As Double, _
    ByVal Shape As Double, _
    ByVal ScaleParam As Double, _
    Optional ByRef Status As String = "") _
    As Variant
'
'==============================================================================
' K_STATS_Gamma_Density
'------------------------------------------------------------------------------
' PURPOSE
'   Returns the Gamma probability density at X for shape k and scale theta.
'
' WHY THIS EXISTS
'   The Gamma density is the workhorse of waiting-time, insurance-loss and
'   Bayesian-conjugate models. Evaluating it through the log-density and a single
'   guarded exponential keeps it finite deep into the tail, where the naive
'   product of a large power and a small exponential would overflow or underflow
'   prematurely.
'
' WORKSHEET EQUIVALENT
'   GAMMA.DIST(X, Shape, ScaleParam, FALSE)
'
' INPUTS
'   X       Evaluation point. For X < 0 the density is 0.
'   Shape   Shape parameter k. Must be strictly positive.
'   ScaleParam   ScaleParam parameter theta. Must be strictly positive.
'   Status  Optional ByRef diagnostic message.
'
' RETURNS
'   Variant
'     Success => Double density value.
'     Failure => CVErr(xlErrNum) or CVErr(xlErrValue).
'
' BEHAVIOR
'   - Returns 0 for X < 0.
'   - At X = 0 the density is unbounded when Shape < 1: returns CVErr(xlErrNum).
'     When Shape = 1 it equals 1 / ScaleParam; when Shape > 1 it equals 0.
'   - Otherwise Exp((k - 1)*Log(X) - X/theta - k*Log(theta) - LogGamma(k)).
'   - Underflow to 0 in the tail is a valid zero; overflow returns CVErr(xlErrNum).
'
' ERROR POLICY
'   - Invalid numeric domains return CVErr(xlErrNum).
'   - Density pole or overflow returns CVErr(xlErrNum).
'   - Unexpected runtime errors return CVErr(xlErrValue).
'   - Detailed diagnostic messages are written to Status.
'   - No MsgBox is raised.
'
' DEPENDENCIES
'   - PROB_CN_ValidateXAndTwoPositive
'   - PROB_LogGamma, PROB_TryExp, PROB_SetStatus
'
' UPDATED
'   2026-07-11
'==============================================================================
'
'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Dim LogDensity          As Double          'Log of the density
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
    'Validate the evaluation point and both positive parameters
        If Not PROB_CN_ValidateXAndTwoPositive( _
            X, Shape, ScaleParam, FailMsg, "Shape", "ScaleParam") Then GoTo Fail_Num

'------------------------------------------------------------------------------
' HANDLE THE SUPPORT EDGE
'------------------------------------------------------------------------------
    'Return zero below the support
        If X < 0# Then
            K_STATS_Gamma_Density = 0#
            GoTo Return_Success
        End If

    'Handle the origin, where the density is 0, 1 / ScaleParam or unbounded
        If X = 0# Then
            If Shape < 1# Then
                FailMsg = "Gamma density is unbounded at X = 0 when Shape < 1"
                GoTo Fail_Num
            ElseIf Shape = 1# Then
                K_STATS_Gamma_Density = 1# / ScaleParam
            Else
                K_STATS_Gamma_Density = 0#
            End If
            GoTo Return_Success
        End If

'------------------------------------------------------------------------------
' COMPUTE DENSITY
'------------------------------------------------------------------------------
    'Compute the log-density
        LogDensity = _
            (Shape - 1#) * Log(X) - _
            X / ScaleParam - _
            Shape * Log(ScaleParam) - _
            PROB_LogGamma(Shape)

    'Exponentiate; underflow to zero is a valid result
        If Not PROB_TryExp(LogDensity, Density) Then
            FailMsg = "Gamma density overflowed a Double"
            GoTo Fail_Num
        End If

    'Return the density
        K_STATS_Gamma_Density = Density

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
        K_STATS_Gamma_Density = CVErr(xlErrNum)
    'Exit before the error handler
        Exit Function

'------------------------------------------------------------------------------
' ERROR HANDLER
'------------------------------------------------------------------------------
Err_Handler:
    'Write unexpected runtime errors to diagnostics
        PROB_SetStatus Status, "Unexpected error in K_STATS_Gamma_Density: " & Err.Description
    'Return worksheet value error
        K_STATS_Gamma_Density = CVErr(xlErrValue)
End Function


Public Function K_STATS_Gamma_Cumulative( _
    ByVal X As Double, _
    ByVal Shape As Double, _
    ByVal ScaleParam As Double, _
    Optional ByRef Status As String = "") _
    As Variant
'
'==============================================================================
' K_STATS_Gamma_Cumulative
'------------------------------------------------------------------------------
' PURPOSE
'   Returns the left-tail Gamma cumulative distribution function at X.
'
' WHY THIS EXISTS
'   Gamma cumulative probabilities give waiting-time and aggregate-loss
'   exceedance levels directly, and are the bridge that turns a chi-square into
'   a Gamma(df/2, 2) for cross-checking.
'
' WORKSHEET EQUIVALENT
'   GAMMA.DIST(X, Shape, ScaleParam, TRUE)
'
' INPUTS
'   X       Evaluation point. For X <= 0 the cumulative probability is 0.
'   Shape   Shape parameter k. Must be strictly positive.
'   ScaleParam   ScaleParam parameter theta. Must be strictly positive.
'   Status  Optional ByRef diagnostic message.
'
' RETURNS
'   Variant
'     Success => Double cumulative probability P(Gamma <= X).
'     Failure => CVErr(xlErrNum) or CVErr(xlErrValue).
'
' BEHAVIOR
'   - Returns 0 for X <= 0.
'   - Computes P(Shape, X / ScaleParam), the regularized lower incomplete gamma.
'   - Non-convergence returns CVErr(xlErrNum); a partial sum is never returned.
'
' ERROR POLICY
'   - Invalid numeric domains return CVErr(xlErrNum).
'   - Non-convergence returns CVErr(xlErrNum).
'   - Unexpected runtime errors return CVErr(xlErrValue).
'   - Detailed diagnostic messages are written to Status.
'   - No MsgBox is raised.
'
' DEPENDENCIES
'   - PROB_CN_ValidateXAndTwoPositive
'   - PROB_TryGammaRegularizedP, PROB_SetStatus
'
' UPDATED
'   2026-07-11
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
    'Validate the evaluation point and both positive parameters
        If Not PROB_CN_ValidateXAndTwoPositive( _
            X, Shape, ScaleParam, FailMsg, "Shape", "ScaleParam") Then GoTo Fail_Num

'------------------------------------------------------------------------------
' COMPUTE CUMULATIVE PROBABILITY
'------------------------------------------------------------------------------
    'Return zero for values outside the positive support
        If X <= 0# Then
            K_STATS_Gamma_Cumulative = 0#
            GoTo Return_Success
        End If

    'Evaluate the regularized lower incomplete gamma function
        If Not PROB_TryGammaRegularizedP( _
            Shape, X / ScaleParam, Value, FailMsg) Then GoTo Fail_Num

    'Return the cumulative probability
        K_STATS_Gamma_Cumulative = Value

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
        K_STATS_Gamma_Cumulative = CVErr(xlErrNum)
    'Exit before the error handler
        Exit Function

'------------------------------------------------------------------------------
' ERROR HANDLER
'------------------------------------------------------------------------------
Err_Handler:
    'Write unexpected runtime errors to diagnostics
        PROB_SetStatus Status, "Unexpected error in K_STATS_Gamma_Cumulative: " & Err.Description
    'Return worksheet value error
        K_STATS_Gamma_Cumulative = CVErr(xlErrValue)
End Function


Public Function K_STATS_Gamma_Survival( _
    ByVal X As Double, _
    ByVal Shape As Double, _
    ByVal ScaleParam As Double, _
    Optional ByRef Status As String = "") _
    As Variant
'
'==============================================================================
' K_STATS_Gamma_Survival
'------------------------------------------------------------------------------
' PURPOSE
'   Returns the right-tail Gamma survival function Q = 1 - CDF at X.
'
' WHY THIS EXISTS
'   The survival function is the exceedance probability a reliability or
'   solvency model actually asks for, and evaluating it through the upper
'   incomplete gamma Q keeps the small tail accurate instead of computing it as
'   1 minus a CDF that has already rounded to one.
'
' WORKSHEET EQUIVALENT
'   1 - GAMMA.DIST(X, Shape, ScaleParam, TRUE)
'
' INPUTS
'   X       Evaluation point. For X <= 0 the survival probability is 1.
'   Shape   Shape parameter k. Must be strictly positive.
'   ScaleParam   ScaleParam parameter theta. Must be strictly positive.
'   Status  Optional ByRef diagnostic message.
'
' RETURNS
'   Variant
'     Success => Double survival probability P(Gamma > X).
'     Failure => CVErr(xlErrNum) or CVErr(xlErrValue).
'
' BEHAVIOR
'   - Returns 1 for X <= 0.
'   - Computes Q(Shape, X / ScaleParam), the regularized upper incomplete gamma.
'   - Non-convergence returns CVErr(xlErrNum).
'
' ERROR POLICY
'   - Invalid numeric domains return CVErr(xlErrNum).
'   - Non-convergence returns CVErr(xlErrNum).
'   - Unexpected runtime errors return CVErr(xlErrValue).
'   - Detailed diagnostic messages are written to Status.
'   - No MsgBox is raised.
'
' DEPENDENCIES
'   - PROB_CN_ValidateXAndTwoPositive
'   - PROB_TryGammaRegularizedQ, PROB_SetStatus
'
' UPDATED
'   2026-07-11
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
    'Validate the evaluation point and both positive parameters
        If Not PROB_CN_ValidateXAndTwoPositive( _
            X, Shape, ScaleParam, FailMsg, "Shape", "ScaleParam") Then GoTo Fail_Num

'------------------------------------------------------------------------------
' COMPUTE SURVIVAL PROBABILITY
'------------------------------------------------------------------------------
    'Return one for values outside the positive support
        If X <= 0# Then
            K_STATS_Gamma_Survival = 1#
            GoTo Return_Success
        End If

    'Evaluate the regularized upper incomplete gamma function
        If Not PROB_TryGammaRegularizedQ( _
            Shape, X / ScaleParam, Value, FailMsg) Then GoTo Fail_Num

    'Return the survival probability
        K_STATS_Gamma_Survival = Value

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
        K_STATS_Gamma_Survival = CVErr(xlErrNum)
    'Exit before the error handler
        Exit Function

'------------------------------------------------------------------------------
' ERROR HANDLER
'------------------------------------------------------------------------------
Err_Handler:
    'Write unexpected runtime errors to diagnostics
        PROB_SetStatus Status, "Unexpected error in K_STATS_Gamma_Survival: " & Err.Description
    'Return worksheet value error
        K_STATS_Gamma_Survival = CVErr(xlErrValue)
End Function


Public Function K_STATS_Gamma_InverseCumulative( _
    ByVal Probability As Double, _
    ByVal Shape As Double, _
    ByVal ScaleParam As Double, _
    Optional ByRef Status As String = "") _
    As Variant
'
'==============================================================================
' K_STATS_Gamma_InverseCumulative
'------------------------------------------------------------------------------
' PURPOSE
'   Returns the Gamma quantile: the X for which P(Gamma <= X) = Probability.
'
' WHY THIS EXISTS
'   Quantiles set reserve levels, capital requirements and simulation cut-offs.
'   The inverse is solved once, in the unit-scale gamma, then rescaled by ScaleParam.
'
' WORKSHEET EQUIVALENT
'   GAMMA.INV(Probability, Shape, ScaleParam)
'
' INPUTS
'   Probability  Target cumulative probability, strictly between 0 and 1.
'   Shape        Shape parameter k. Must be strictly positive.
'   ScaleParam        ScaleParam parameter theta. Must be strictly positive.
'   Status       Optional ByRef diagnostic message.
'
' RETURNS
'   Variant
'     Success => Double quantile X.
'     Failure => CVErr(xlErrNum) or CVErr(xlErrValue).
'
' BEHAVIOR
'   - Inverts P(Shape, .) via PROB_TryGammaInvP on the unit-scale gamma, then
'     multiplies the result by ScaleParam.
'   - Non-convergence returns CVErr(xlErrNum).
'
' ERROR POLICY
'   - Probability outside (0, 1) or invalid parameters return CVErr(xlErrNum).
'   - Non-convergence returns CVErr(xlErrNum).
'   - Unexpected runtime errors return CVErr(xlErrValue).
'   - Detailed diagnostic messages are written to Status.
'   - No MsgBox is raised.
'
' DEPENDENCIES
'   - PROB_IsValidProbabilityOpen, PROB_IsPositiveFinite
'   - PROB_TryGammaInvP, PROB_SetStatus
'
' UPDATED
'   2026-07-11
'==============================================================================
'
'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Dim GammaQuantile       As Double          'Unit-scale gamma quantile
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
    'Validate shape
        If Not PROB_IsPositiveFinite(Shape) Then
            FailMsg = "Shape must be a finite strictly positive number"
            GoTo Fail_Num
        End If
    'Validate scale
        If Not PROB_IsPositiveFinite(ScaleParam) Then
            FailMsg = "ScaleParam must be a finite strictly positive number"
            GoTo Fail_Num
        End If

'------------------------------------------------------------------------------
' COMPUTE QUANTILE
'------------------------------------------------------------------------------
    'Invert the regularized lower incomplete gamma function on the unit scale
        If Not PROB_TryGammaInvP( _
            Probability, 1# - Probability, Shape, GammaQuantile, FailMsg) Then GoTo Fail_Num

    'Rescale from the unit-scale gamma to the requested scale
        K_STATS_Gamma_InverseCumulative = ScaleParam * GammaQuantile

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
        K_STATS_Gamma_InverseCumulative = CVErr(xlErrNum)
    'Exit before the error handler
        Exit Function

'------------------------------------------------------------------------------
' ERROR HANDLER
'------------------------------------------------------------------------------
Err_Handler:
    'Write unexpected runtime errors to diagnostics
        PROB_SetStatus Status, "Unexpected error in K_STATS_Gamma_InverseCumulative: " & Err.Description
    'Return worksheet value error
        K_STATS_Gamma_InverseCumulative = CVErr(xlErrValue)
End Function


Public Function K_STATS_Gamma_Mean( _
    ByVal Shape As Double, _
    ByVal ScaleParam As Double, _
    Optional ByRef Status As String = "") _
    As Variant
'
'==============================================================================
' K_STATS_Gamma_Mean
'------------------------------------------------------------------------------
' PURPOSE
'   Returns the mean of the Gamma distribution, Shape * ScaleParam.
'
' WORKSHEET EQUIVALENT
'   (none; Shape * ScaleParam)
'
' INPUTS
'   Shape   Shape parameter k. Must be strictly positive.
'   ScaleParam   ScaleParam parameter theta. Must be strictly positive.
'   Status  Optional ByRef diagnostic message.
'
' RETURNS
'   Variant
'     Success => Double mean Shape * ScaleParam.
'     Failure => CVErr(xlErrNum) or CVErr(xlErrValue).
'
' BEHAVIOR
'   - Guards the product against Double overflow before forming it.
'
' ERROR POLICY
'   - Invalid parameters or overflow return CVErr(xlErrNum).
'   - Unexpected runtime errors return CVErr(xlErrValue).
'
' DEPENDENCIES
'   - PROB_IsPositiveFinite, PROB_DOUBLE_MAX, PROB_SetStatus
'
' UPDATED
'   2026-07-11
'==============================================================================
'
'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
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
    'Validate shape
        If Not PROB_IsPositiveFinite(Shape) Then
            FailMsg = "Shape must be a finite strictly positive number"
            GoTo Fail_Num
        End If
    'Validate scale
        If Not PROB_IsPositiveFinite(ScaleParam) Then
            FailMsg = "ScaleParam must be a finite strictly positive number"
            GoTo Fail_Num
        End If

'------------------------------------------------------------------------------
' COMPUTE MEAN
'------------------------------------------------------------------------------
    'Guard the product against overflow, then form it
        If Shape > PROB_DOUBLE_MAX / ScaleParam Then
            FailMsg = "Gamma mean overflows Double range"
            GoTo Fail_Num
        End If

    'Return the mean
        K_STATS_Gamma_Mean = Shape * ScaleParam

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
        K_STATS_Gamma_Mean = CVErr(xlErrNum)
    'Exit before the error handler
        Exit Function

'------------------------------------------------------------------------------
' ERROR HANDLER
'------------------------------------------------------------------------------
Err_Handler:
    'Write unexpected runtime errors to diagnostics
        PROB_SetStatus Status, "Unexpected error in K_STATS_Gamma_Mean: " & Err.Description
    'Return worksheet value error
        K_STATS_Gamma_Mean = CVErr(xlErrValue)
End Function


Public Function K_STATS_Gamma_Variance( _
    ByVal Shape As Double, _
    ByVal ScaleParam As Double, _
    Optional ByRef Status As String = "") _
    As Variant
'
'==============================================================================
' K_STATS_Gamma_Variance
'------------------------------------------------------------------------------
' PURPOSE
'   Returns the variance of the Gamma distribution, Shape * ScaleParam ^ 2.
'
' WORKSHEET EQUIVALENT
'   (none; Shape * ScaleParam ^ 2)
'
' INPUTS
'   Shape   Shape parameter k. Must be strictly positive.
'   ScaleParam   ScaleParam parameter theta. Must be strictly positive.
'   Status  Optional ByRef diagnostic message.
'
' RETURNS
'   Variant
'     Success => Double variance Shape * ScaleParam ^ 2.
'     Failure => CVErr(xlErrNum) or CVErr(xlErrValue).
'
' BEHAVIOR
'   - Guards each multiplication against Double overflow, nested so the guard is
'     never itself a division by zero.
'
' ERROR POLICY
'   - Invalid parameters or overflow return CVErr(xlErrNum).
'   - Unexpected runtime errors return CVErr(xlErrValue).
'
' DEPENDENCIES
'   - PROB_IsPositiveFinite, PROB_DOUBLE_MAX, PROB_SetStatus
'
' UPDATED
'   2026-07-11
'==============================================================================
'
'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Dim Partial             As Double          'Shape * ScaleParam, held for the second multiply
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
    'Validate shape
        If Not PROB_IsPositiveFinite(Shape) Then
            FailMsg = "Shape must be a finite strictly positive number"
            GoTo Fail_Num
        End If
    'Validate scale
        If Not PROB_IsPositiveFinite(ScaleParam) Then
            FailMsg = "ScaleParam must be a finite strictly positive number"
            GoTo Fail_Num
        End If

'------------------------------------------------------------------------------
' COMPUTE VARIANCE
'------------------------------------------------------------------------------
    'Guard and form Shape * ScaleParam
        If Shape > PROB_DOUBLE_MAX / ScaleParam Then
            FailMsg = "Gamma variance overflows Double range"
            GoTo Fail_Num
        End If
        Partial = Shape * ScaleParam

    'Guard and form (Shape * ScaleParam) * ScaleParam
        If Partial > PROB_DOUBLE_MAX / ScaleParam Then
            FailMsg = "Gamma variance overflows Double range"
            GoTo Fail_Num
        End If

    'Return the variance
        K_STATS_Gamma_Variance = Partial * ScaleParam

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
        K_STATS_Gamma_Variance = CVErr(xlErrNum)
    'Exit before the error handler
        Exit Function

'------------------------------------------------------------------------------
' ERROR HANDLER
'------------------------------------------------------------------------------
Err_Handler:
    'Write unexpected runtime errors to diagnostics
        PROB_SetStatus Status, "Unexpected error in K_STATS_Gamma_Variance: " & Err.Description
    'Return worksheet value error
        K_STATS_Gamma_Variance = CVErr(xlErrValue)
End Function


Public Function K_STATS_Gamma_StdDev( _
    ByVal Shape As Double, _
    ByVal ScaleParam As Double, _
    Optional ByRef Status As String = "") _
    As Variant
'
'==============================================================================
' K_STATS_Gamma_StdDev
'------------------------------------------------------------------------------
' PURPOSE
'   Returns the standard deviation of the Gamma distribution, ScaleParam * Sqr(Shape).
'
' WORKSHEET EQUIVALENT
'   (none; ScaleParam * SQRT(Shape))
'
' INPUTS
'   Shape   Shape parameter k. Must be strictly positive.
'   ScaleParam   ScaleParam parameter theta. Must be strictly positive.
'   Status  Optional ByRef diagnostic message.
'
' RETURNS
'   Variant
'     Success => Double standard deviation ScaleParam * Sqr(Shape).
'     Failure => CVErr(xlErrNum) or CVErr(xlErrValue).
'
' BEHAVIOR
'   - Uses the reduced-magnitude form ScaleParam * Sqr(Shape) rather than
'     Sqr(Variance), so it stays finite for parameters whose variance would
'     overflow.
'
' ERROR POLICY
'   - Invalid parameters or overflow return CVErr(xlErrNum).
'   - Unexpected runtime errors return CVErr(xlErrValue).
'
' DEPENDENCIES
'   - PROB_IsPositiveFinite, PROB_DOUBLE_MAX, PROB_SetStatus
'
' UPDATED
'   2026-07-11
'==============================================================================
'
'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Dim RootShape           As Double          'Sqr(Shape)
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
    'Validate shape
        If Not PROB_IsPositiveFinite(Shape) Then
            FailMsg = "Shape must be a finite strictly positive number"
            GoTo Fail_Num
        End If
    'Validate scale
        If Not PROB_IsPositiveFinite(ScaleParam) Then
            FailMsg = "ScaleParam must be a finite strictly positive number"
            GoTo Fail_Num
        End If

'------------------------------------------------------------------------------
' COMPUTE STANDARD DEVIATION
'------------------------------------------------------------------------------
    'Take the root of the shape
        RootShape = Sqr(Shape)

    'Guard the product against overflow, then form it
        If ScaleParam > PROB_DOUBLE_MAX / RootShape Then
            FailMsg = "Gamma standard deviation overflows Double range"
            GoTo Fail_Num
        End If

    'Return the standard deviation
        K_STATS_Gamma_StdDev = ScaleParam * RootShape

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
        K_STATS_Gamma_StdDev = CVErr(xlErrNum)
    'Exit before the error handler
        Exit Function

'------------------------------------------------------------------------------
' ERROR HANDLER
'------------------------------------------------------------------------------
Err_Handler:
    'Write unexpected runtime errors to diagnostics
        PROB_SetStatus Status, "Unexpected error in K_STATS_Gamma_StdDev: " & Err.Description
    'Return worksheet value error
        K_STATS_Gamma_StdDev = CVErr(xlErrValue)
End Function


'==============================================================================
' BETA DISTRIBUTION
'==============================================================================

Public Function K_STATS_Beta_Density( _
    ByVal X As Double, _
    ByVal Alpha As Double, _
    ByVal Beta As Double, _
    Optional ByRef Status As String = "") _
    As Variant
'
'==============================================================================
' K_STATS_Beta_Density
'------------------------------------------------------------------------------
' PURPOSE
'   Returns the Beta probability density at X on [0, 1] for shapes Alpha, Beta.
'
' WHY THIS EXISTS
'   The Beta density is the conjugate prior for a proportion and the reference
'   shape for anything bounded on [0, 1]. Evaluating it through the log-density
'   and PROB_LogBeta avoids the overflow that a direct ratio of gamma functions
'   would hit for large shapes.
'
' WORKSHEET EQUIVALENT
'   BETA.DIST(X, Alpha, Beta, FALSE)
'
' INPUTS
'   X       Evaluation point. Outside [0, 1] the density is 0.
'   Alpha   First shape parameter. Must be strictly positive.
'   Beta    Second shape parameter. Must be strictly positive.
'   Status  Optional ByRef diagnostic message.
'
' RETURNS
'   Variant
'     Success => Double density value.
'     Failure => CVErr(xlErrNum) or CVErr(xlErrValue).
'
' BEHAVIOR
'   - Returns 0 for X < 0 or X > 1.
'   - At X = 0 the density is unbounded when Alpha < 1: returns CVErr(xlErrNum).
'     When Alpha = 1 it equals Beta; when Alpha > 1 it equals 0.
'   - At X = 1 the density is unbounded when Beta < 1: returns CVErr(xlErrNum).
'     When Beta = 1 it equals Alpha; when Beta > 1 it equals 0.
'   - Otherwise Exp((Alpha-1)*Log(X) + (Beta-1)*Log1p(-X) - LogBeta(Alpha,Beta)).
'     The (1 - X) logarithm is taken through PROB_Log1p(-X) so the density keeps
'     full relative precision as X approaches 1.
'
' ERROR POLICY
'   - Invalid numeric domains return CVErr(xlErrNum).
'   - Density pole or overflow returns CVErr(xlErrNum).
'   - Unexpected runtime errors return CVErr(xlErrValue).
'   - Detailed diagnostic messages are written to Status.
'   - No MsgBox is raised.
'
' DEPENDENCIES
'   - PROB_CN_ValidateXAndTwoPositive
'   - PROB_Log1p, PROB_LogBeta, PROB_TryExp, PROB_SetStatus
'
' UPDATED
'   2026-07-11
'==============================================================================
'
'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Dim LogDensity          As Double          'Log of the density
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
    'Validate the evaluation point and both positive parameters
        If Not PROB_CN_ValidateXAndTwoPositive( _
            X, Alpha, Beta, FailMsg, "Alpha", "Beta") Then GoTo Fail_Num

'------------------------------------------------------------------------------
' HANDLE THE SUPPORT EDGE
'------------------------------------------------------------------------------
    'Return zero outside the closed unit interval
        If X < 0# Or X > 1# Then
            K_STATS_Beta_Density = 0#
            GoTo Return_Success
        End If

    'Handle the left endpoint, where the density is 0, Beta or unbounded
        If X = 0# Then
            If Alpha < 1# Then
                FailMsg = "Beta density is unbounded at X = 0 when Alpha < 1"
                GoTo Fail_Num
            ElseIf Alpha = 1# Then
                K_STATS_Beta_Density = Beta
            Else
                K_STATS_Beta_Density = 0#
            End If
            GoTo Return_Success
        End If

    'Handle the right endpoint, where the density is 0, Alpha or unbounded
        If X = 1# Then
            If Beta < 1# Then
                FailMsg = "Beta density is unbounded at X = 1 when Beta < 1"
                GoTo Fail_Num
            ElseIf Beta = 1# Then
                K_STATS_Beta_Density = Alpha
            Else
                K_STATS_Beta_Density = 0#
            End If
            GoTo Return_Success
        End If

'------------------------------------------------------------------------------
' COMPUTE DENSITY
'------------------------------------------------------------------------------
    'Compute the log-density, taking Log(1 - X) through PROB_Log1p(-X)
        LogDensity = _
            (Alpha - 1#) * Log(X) + _
            (Beta - 1#) * PROB_Log1p(-X) - _
            PROB_LogBeta(Alpha, Beta)

    'Exponentiate; underflow to zero is a valid result
        If Not PROB_TryExp(LogDensity, Density) Then
            FailMsg = "Beta density overflowed a Double"
            GoTo Fail_Num
        End If

    'Return the density
        K_STATS_Beta_Density = Density

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
        K_STATS_Beta_Density = CVErr(xlErrNum)
    'Exit before the error handler
        Exit Function

'------------------------------------------------------------------------------
' ERROR HANDLER
'------------------------------------------------------------------------------
Err_Handler:
    'Write unexpected runtime errors to diagnostics
        PROB_SetStatus Status, "Unexpected error in K_STATS_Beta_Density: " & Err.Description
    'Return worksheet value error
        K_STATS_Beta_Density = CVErr(xlErrValue)
End Function


Public Function K_STATS_Beta_Cumulative( _
    ByVal X As Double, _
    ByVal Alpha As Double, _
    ByVal Beta As Double, _
    Optional ByRef Status As String = "") _
    As Variant
'
'==============================================================================
' K_STATS_Beta_Cumulative
'------------------------------------------------------------------------------
' PURPOSE
'   Returns the left-tail Beta cumulative distribution function at X.
'
' WHY THIS EXISTS
'   The Beta CDF is the regularized incomplete beta itself, and is the bridge
'   that turns an F variate into a Beta for cross-checking.
'
' WORKSHEET EQUIVALENT
'   BETA.DIST(X, Alpha, Beta, TRUE)
'
' INPUTS
'   X       Evaluation point. For X <= 0 the CDF is 0; for X >= 1 it is 1.
'   Alpha   First shape parameter. Must be strictly positive.
'   Beta    Second shape parameter. Must be strictly positive.
'   Status  Optional ByRef diagnostic message.
'
' RETURNS
'   Variant
'     Success => Double cumulative probability I_X(Alpha, Beta).
'     Failure => CVErr(xlErrNum) or CVErr(xlErrValue).
'
' BEHAVIOR
'   - Returns 0 for X <= 0 and 1 for X >= 1.
'   - Otherwise the regularized incomplete beta I_X(Alpha, Beta), computed with
'     both X and 1 - X passed to the kernel.
'
' ERROR POLICY
'   - Invalid numeric domains return CVErr(xlErrNum).
'   - Non-convergence returns CVErr(xlErrNum).
'   - Unexpected runtime errors return CVErr(xlErrValue).
'   - Detailed diagnostic messages are written to Status.
'   - No MsgBox is raised.
'
' DEPENDENCIES
'   - PROB_CN_ValidateXAndTwoPositive
'   - PROB_TryBetaRegularized, PROB_SetStatus
'
' UPDATED
'   2026-07-11
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
    'Validate the evaluation point and both positive parameters
        If Not PROB_CN_ValidateXAndTwoPositive( _
            X, Alpha, Beta, FailMsg, "Alpha", "Beta") Then GoTo Fail_Num

'------------------------------------------------------------------------------
' COMPUTE CUMULATIVE PROBABILITY
'------------------------------------------------------------------------------
    'Return the support edges exactly
        If X <= 0# Then
            K_STATS_Beta_Cumulative = 0#
            GoTo Return_Success
        End If
        If X >= 1# Then
            K_STATS_Beta_Cumulative = 1#
            GoTo Return_Success
        End If

    'Evaluate the regularized incomplete beta I_X(Alpha, Beta)
        If Not PROB_TryBetaRegularized( _
            X, 1# - X, Alpha, Beta, Value, FailMsg) Then GoTo Fail_Num

    'Return the cumulative probability
        K_STATS_Beta_Cumulative = Value

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
        K_STATS_Beta_Cumulative = CVErr(xlErrNum)
    'Exit before the error handler
        Exit Function

'------------------------------------------------------------------------------
' ERROR HANDLER
'------------------------------------------------------------------------------
Err_Handler:
    'Write unexpected runtime errors to diagnostics
        PROB_SetStatus Status, "Unexpected error in K_STATS_Beta_Cumulative: " & Err.Description
    'Return worksheet value error
        K_STATS_Beta_Cumulative = CVErr(xlErrValue)
End Function


Public Function K_STATS_Beta_Survival( _
    ByVal X As Double, _
    ByVal Alpha As Double, _
    ByVal Beta As Double, _
    Optional ByRef Status As String = "") _
    As Variant
'
'==============================================================================
' K_STATS_Beta_Survival
'------------------------------------------------------------------------------
' PURPOSE
'   Returns the right-tail Beta survival function 1 - CDF at X.
'
' WHY THIS EXISTS
'   Computing the upper tail as the swapped incomplete beta I_(1-X)(Beta, Alpha)
'   rather than 1 - I_X(Alpha, Beta) keeps the small tail accurate: the naive
'   subtraction loses every digit once the CDF has rounded to one.
'
' WORKSHEET EQUIVALENT
'   1 - BETA.DIST(X, Alpha, Beta, TRUE)
'
' INPUTS
'   X       Evaluation point. For X <= 0 the survival is 1; for X >= 1 it is 0.
'   Alpha   First shape parameter. Must be strictly positive.
'   Beta    Second shape parameter. Must be strictly positive.
'   Status  Optional ByRef diagnostic message.
'
' RETURNS
'   Variant
'     Success => Double survival probability P(Beta > X) = I_(1-X)(Beta, Alpha).
'     Failure => CVErr(xlErrNum) or CVErr(xlErrValue).
'
' BEHAVIOR
'   - Returns 1 for X <= 0 and 0 for X >= 1.
'   - Otherwise the swapped regularized incomplete beta I_(1-X)(Beta, Alpha).
'
' ERROR POLICY
'   - Invalid numeric domains return CVErr(xlErrNum).
'   - Non-convergence returns CVErr(xlErrNum).
'   - Unexpected runtime errors return CVErr(xlErrValue).
'   - Detailed diagnostic messages are written to Status.
'   - No MsgBox is raised.
'
' DEPENDENCIES
'   - PROB_CN_ValidateXAndTwoPositive
'   - PROB_TryBetaRegularized, PROB_SetStatus
'
' UPDATED
'   2026-07-11
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
    'Validate the evaluation point and both positive parameters
        If Not PROB_CN_ValidateXAndTwoPositive( _
            X, Alpha, Beta, FailMsg, "Alpha", "Beta") Then GoTo Fail_Num

'------------------------------------------------------------------------------
' COMPUTE SURVIVAL PROBABILITY
'------------------------------------------------------------------------------
    'Return the support edges exactly
        If X <= 0# Then
            K_STATS_Beta_Survival = 1#
            GoTo Return_Success
        End If
        If X >= 1# Then
            K_STATS_Beta_Survival = 0#
            GoTo Return_Success
        End If

    'Evaluate the swapped regularized incomplete beta I_(1-X)(Beta, Alpha)
        If Not PROB_TryBetaRegularized( _
            1# - X, X, Beta, Alpha, Value, FailMsg) Then GoTo Fail_Num

    'Return the survival probability
        K_STATS_Beta_Survival = Value

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
        K_STATS_Beta_Survival = CVErr(xlErrNum)
    'Exit before the error handler
        Exit Function

'------------------------------------------------------------------------------
' ERROR HANDLER
'------------------------------------------------------------------------------
Err_Handler:
    'Write unexpected runtime errors to diagnostics
        PROB_SetStatus Status, "Unexpected error in K_STATS_Beta_Survival: " & Err.Description
    'Return worksheet value error
        K_STATS_Beta_Survival = CVErr(xlErrValue)
End Function


Public Function K_STATS_Beta_InverseCumulative( _
    ByVal Probability As Double, _
    ByVal Alpha As Double, _
    ByVal Beta As Double, _
    Optional ByRef Status As String = "") _
    As Variant
'
'==============================================================================
' K_STATS_Beta_InverseCumulative
'------------------------------------------------------------------------------
' PURPOSE
'   Returns the Beta quantile: the X for which I_X(Alpha, Beta) = Probability.
'
' WORKSHEET EQUIVALENT
'   BETA.INV(Probability, Alpha, Beta)
'
' INPUTS
'   Probability  Target cumulative probability, strictly between 0 and 1.
'   Alpha        First shape parameter. Must be strictly positive.
'   Beta         Second shape parameter. Must be strictly positive.
'   Status       Optional ByRef diagnostic message.
'
' RETURNS
'   Variant
'     Success => Double quantile X in (0, 1).
'     Failure => CVErr(xlErrNum) or CVErr(xlErrValue).
'
' BEHAVIOR
'   - Inverts the regularized incomplete beta via PROB_TryBetaInvRegularized,
'     which returns the quantile and its complement as a cancellation-free pair;
'     this function returns the quantile.
'
' ERROR POLICY
'   - Probability outside (0, 1) or invalid parameters return CVErr(xlErrNum).
'   - Non-convergence returns CVErr(xlErrNum).
'   - Unexpected runtime errors return CVErr(xlErrValue).
'   - Detailed diagnostic messages are written to Status.
'   - No MsgBox is raised.
'
' DEPENDENCIES
'   - PROB_IsValidProbabilityOpen, PROB_IsPositiveFinite
'   - PROB_TryBetaInvRegularized, PROB_SetStatus
'
' UPDATED
'   2026-07-11
'==============================================================================
'
'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Dim ResultX             As Double          'Quantile
    Dim ResultY             As Double          'Complementary quantile 1 - X
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
    'Validate first shape
        If Not PROB_IsPositiveFinite(Alpha) Then
            FailMsg = "Alpha must be a finite strictly positive number"
            GoTo Fail_Num
        End If
    'Validate second shape
        If Not PROB_IsPositiveFinite(Beta) Then
            FailMsg = "Beta must be a finite strictly positive number"
            GoTo Fail_Num
        End If

'------------------------------------------------------------------------------
' COMPUTE QUANTILE
'------------------------------------------------------------------------------
    'Invert the regularized incomplete beta function
        If Not PROB_TryBetaInvRegularized( _
            Probability, 1# - Probability, Alpha, Beta, ResultX, ResultY, FailMsg) Then GoTo Fail_Num

    'Return the quantile
        K_STATS_Beta_InverseCumulative = ResultX

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
        K_STATS_Beta_InverseCumulative = CVErr(xlErrNum)
    'Exit before the error handler
        Exit Function

'------------------------------------------------------------------------------
' ERROR HANDLER
'------------------------------------------------------------------------------
Err_Handler:
    'Write unexpected runtime errors to diagnostics
        PROB_SetStatus Status, "Unexpected error in K_STATS_Beta_InverseCumulative: " & Err.Description
    'Return worksheet value error
        K_STATS_Beta_InverseCumulative = CVErr(xlErrValue)
End Function


Public Function K_STATS_Beta_Mean( _
    ByVal Alpha As Double, _
    ByVal Beta As Double, _
    Optional ByRef Status As String = "") _
    As Variant
'
'==============================================================================
' K_STATS_Beta_Mean
'------------------------------------------------------------------------------
' PURPOSE
'   Returns the mean of the Beta distribution, Alpha / (Alpha + Beta).
'
' WORKSHEET EQUIVALENT
'   (none; Alpha / (Alpha + Beta))
'
' INPUTS
'   Alpha   First shape parameter. Must be strictly positive.
'   Beta    Second shape parameter. Must be strictly positive.
'   Status  Optional ByRef diagnostic message.
'
' RETURNS
'   Variant
'     Success => Double mean in (0, 1).
'     Failure => CVErr(xlErrNum) or CVErr(xlErrValue).
'
' BEHAVIOR
'   - Rejects the case where Alpha + Beta overflows to a non-finite Double.
'
' ERROR POLICY
'   - Invalid parameters or a non-finite Alpha + Beta return CVErr(xlErrNum).
'   - Unexpected runtime errors return CVErr(xlErrValue).
'
' DEPENDENCIES
'   - PROB_IsPositiveFinite, PROB_IsFinite, PROB_SetStatus
'
' UPDATED
'   2026-07-11
'==============================================================================
'
'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Dim Sum                 As Double          'Alpha + Beta
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
    'Validate first shape
        If Not PROB_IsPositiveFinite(Alpha) Then
            FailMsg = "Alpha must be a finite strictly positive number"
            GoTo Fail_Num
        End If
    'Validate second shape
        If Not PROB_IsPositiveFinite(Beta) Then
            FailMsg = "Beta must be a finite strictly positive number"
            GoTo Fail_Num
        End If

'------------------------------------------------------------------------------
' COMPUTE MEAN
'------------------------------------------------------------------------------
    'Form the shape sum and reject a non-finite result
        Sum = Alpha + Beta
        If Not PROB_IsFinite(Sum) Then
            FailMsg = "Beta mean overflows Double range in Alpha + Beta"
            GoTo Fail_Num
        End If

    'Return the mean
        K_STATS_Beta_Mean = Alpha / Sum

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
        K_STATS_Beta_Mean = CVErr(xlErrNum)
    'Exit before the error handler
        Exit Function

'------------------------------------------------------------------------------
' ERROR HANDLER
'------------------------------------------------------------------------------
Err_Handler:
    'Write unexpected runtime errors to diagnostics
        PROB_SetStatus Status, "Unexpected error in K_STATS_Beta_Mean: " & Err.Description
    'Return worksheet value error
        K_STATS_Beta_Mean = CVErr(xlErrValue)
End Function


Public Function K_STATS_Beta_Variance( _
    ByVal Alpha As Double, _
    ByVal Beta As Double, _
    Optional ByRef Status As String = "") _
    As Variant
'
'==============================================================================
' K_STATS_Beta_Variance
'------------------------------------------------------------------------------
' PURPOSE
'   Returns the variance of the Beta distribution,
'   Alpha * Beta / ((Alpha + Beta) ^ 2 * (Alpha + Beta + 1)).
'
' WORKSHEET EQUIVALENT
'   (none)
'
' INPUTS
'   Alpha   First shape parameter. Must be strictly positive.
'   Beta    Second shape parameter. Must be strictly positive.
'   Status  Optional ByRef diagnostic message.
'
' RETURNS
'   Variant
'     Success => Double variance in (0, 0.25].
'     Failure => CVErr(xlErrNum) or CVErr(xlErrValue).
'
' BEHAVIOR
'   - Uses the algebraically equal form Mean * (1 - Mean) / (Alpha + Beta + 1),
'     which never forms the product Alpha * Beta and so cannot overflow for
'     large shapes.
'
' ERROR POLICY
'   - Invalid parameters or a non-finite Alpha + Beta return CVErr(xlErrNum).
'   - Unexpected runtime errors return CVErr(xlErrValue).
'
' DEPENDENCIES
'   - PROB_IsPositiveFinite, PROB_IsFinite, PROB_SetStatus
'
' UPDATED
'   2026-07-11
'==============================================================================
'
'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Dim Sum                 As Double          'Alpha + Beta
    Dim MeanValue           As Double          'Alpha / Sum
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
    'Validate first shape
        If Not PROB_IsPositiveFinite(Alpha) Then
            FailMsg = "Alpha must be a finite strictly positive number"
            GoTo Fail_Num
        End If
    'Validate second shape
        If Not PROB_IsPositiveFinite(Beta) Then
            FailMsg = "Beta must be a finite strictly positive number"
            GoTo Fail_Num
        End If

'------------------------------------------------------------------------------
' COMPUTE VARIANCE
'------------------------------------------------------------------------------
    'Form the shape sum and reject a non-finite result
        Sum = Alpha + Beta
        If Not PROB_IsFinite(Sum) Then
            FailMsg = "Beta variance overflows Double range in Alpha + Beta"
            GoTo Fail_Num
        End If

    'Return the variance in the non-overflowing form
        MeanValue = Alpha / Sum
        K_STATS_Beta_Variance = MeanValue * (1# - MeanValue) / (Sum + 1#)

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
        K_STATS_Beta_Variance = CVErr(xlErrNum)
    'Exit before the error handler
        Exit Function

'------------------------------------------------------------------------------
' ERROR HANDLER
'------------------------------------------------------------------------------
Err_Handler:
    'Write unexpected runtime errors to diagnostics
        PROB_SetStatus Status, "Unexpected error in K_STATS_Beta_Variance: " & Err.Description
    'Return worksheet value error
        K_STATS_Beta_Variance = CVErr(xlErrValue)
End Function


Public Function K_STATS_Beta_StdDev( _
    ByVal Alpha As Double, _
    ByVal Beta As Double, _
    Optional ByRef Status As String = "") _
    As Variant
'
'==============================================================================
' K_STATS_Beta_StdDev
'------------------------------------------------------------------------------
' PURPOSE
'   Returns the standard deviation of the Beta distribution, Sqr(Variance).
'
' WORKSHEET EQUIVALENT
'   (none)
'
' INPUTS
'   Alpha   First shape parameter. Must be strictly positive.
'   Beta    Second shape parameter. Must be strictly positive.
'   Status  Optional ByRef diagnostic message.
'
' RETURNS
'   Variant
'     Success => Double standard deviation.
'     Failure => CVErr(xlErrNum) or CVErr(xlErrValue).
'
' BEHAVIOR
'   - The Beta variance is bounded above by 0.25, so Sqr(Variance) is safe and
'     needs no separate overflow guard.
'
' ERROR POLICY
'   - Invalid parameters or a non-finite Alpha + Beta return CVErr(xlErrNum).
'   - Unexpected runtime errors return CVErr(xlErrValue).
'
' DEPENDENCIES
'   - PROB_IsPositiveFinite, PROB_IsFinite, PROB_SetStatus
'
' UPDATED
'   2026-07-11
'==============================================================================
'
'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Dim Sum                 As Double          'Alpha + Beta
    Dim MeanValue           As Double          'Alpha / Sum
    Dim Variance            As Double          'Beta variance
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
    'Validate first shape
        If Not PROB_IsPositiveFinite(Alpha) Then
            FailMsg = "Alpha must be a finite strictly positive number"
            GoTo Fail_Num
        End If
    'Validate second shape
        If Not PROB_IsPositiveFinite(Beta) Then
            FailMsg = "Beta must be a finite strictly positive number"
            GoTo Fail_Num
        End If

'------------------------------------------------------------------------------
' COMPUTE STANDARD DEVIATION
'------------------------------------------------------------------------------
    'Form the shape sum and reject a non-finite result
        Sum = Alpha + Beta
        If Not PROB_IsFinite(Sum) Then
            FailMsg = "Beta standard deviation overflows Double range in Alpha + Beta"
            GoTo Fail_Num
        End If

    'Form the variance in the non-overflowing form, then take its root
        MeanValue = Alpha / Sum
        Variance = MeanValue * (1# - MeanValue) / (Sum + 1#)
        K_STATS_Beta_StdDev = Sqr(Variance)

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
        K_STATS_Beta_StdDev = CVErr(xlErrNum)
    'Exit before the error handler
        Exit Function

'------------------------------------------------------------------------------
' ERROR HANDLER
'------------------------------------------------------------------------------
Err_Handler:
    'Write unexpected runtime errors to diagnostics
        PROB_SetStatus Status, "Unexpected error in K_STATS_Beta_StdDev: " & Err.Description
    'Return worksheet value error
        K_STATS_Beta_StdDev = CVErr(xlErrValue)
End Function


'==============================================================================
' EXPONENTIAL DISTRIBUTION  (Lambda = RATE, matching EXPON.DIST)
'==============================================================================

Public Function K_STATS_Exponential_Density( _
    ByVal X As Double, _
    ByVal Lambda As Double, _
    Optional ByRef Status As String = "") _
    As Variant
'
'==============================================================================
' K_STATS_Exponential_Density
'------------------------------------------------------------------------------
' PURPOSE
'   Returns the Exponential probability density at X for rate Lambda.
'
' WHY THIS EXISTS
'   The Exponential is the memoryless waiting time between Poisson events. Its
'   density is a closed form; it is written through the log-density only so that
'   an enormous rate cannot overflow the leading factor before the exponential
'   damps it.
'
' WORKSHEET EQUIVALENT
'   EXPON.DIST(X, Lambda, FALSE)
'
' INPUTS
'   X       Evaluation point. For X < 0 the density is 0.
'   Lambda  Rate parameter (NOT a scale). Must be strictly positive. The mean is
'           1 / Lambda.
'   Status  Optional ByRef diagnostic message.
'
' RETURNS
'   Variant
'     Success => Double density value.
'     Failure => CVErr(xlErrNum) or CVErr(xlErrValue).
'
' BEHAVIOR
'   - Returns 0 for X < 0.
'   - At X = 0 the density is Lambda; it is finite for every valid Lambda, so
'     there is no pole here (unlike Gamma or Weibull with shape < 1).
'   - Otherwise Exp(Log(Lambda) - Lambda * X).
'
' ERROR POLICY
'   - Invalid numeric domains return CVErr(xlErrNum).
'   - Overflow returns CVErr(xlErrNum).
'   - Unexpected runtime errors return CVErr(xlErrValue).
'   - Detailed diagnostic messages are written to Status.
'   - No MsgBox is raised.
'
' DEPENDENCIES
'   - PROB_CN_ValidateXLambda
'   - PROB_TryExp, PROB_SetStatus
'
' UPDATED
'   2026-07-11
'==============================================================================
'
'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
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
    'Validate the evaluation point and the positive rate
        If Not PROB_CN_ValidateXLambda(X, Lambda, FailMsg) Then GoTo Fail_Num

'------------------------------------------------------------------------------
' HANDLE THE SUPPORT EDGE
'------------------------------------------------------------------------------
    'Return zero below the support
        If X < 0# Then
            K_STATS_Exponential_Density = 0#
            GoTo Return_Success
        End If

'------------------------------------------------------------------------------
' COMPUTE DENSITY
'------------------------------------------------------------------------------
    'Exponentiate the log-density; underflow to zero is a valid result
        If Not PROB_TryExp(Log(Lambda) - Lambda * X, Density) Then
            FailMsg = "Exponential density overflowed a Double"
            GoTo Fail_Num
        End If

    'Return the density
        K_STATS_Exponential_Density = Density

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
        K_STATS_Exponential_Density = CVErr(xlErrNum)
    'Exit before the error handler
        Exit Function

'------------------------------------------------------------------------------
' ERROR HANDLER
'------------------------------------------------------------------------------
Err_Handler:
    'Write unexpected runtime errors to diagnostics
        PROB_SetStatus Status, "Unexpected error in K_STATS_Exponential_Density: " & Err.Description
    'Return worksheet value error
        K_STATS_Exponential_Density = CVErr(xlErrValue)
End Function


Public Function K_STATS_Exponential_Cumulative( _
    ByVal X As Double, _
    ByVal Lambda As Double, _
    Optional ByRef Status As String = "") _
    As Variant
'
'==============================================================================
' K_STATS_Exponential_Cumulative
'------------------------------------------------------------------------------
' PURPOSE
'   Returns the left-tail Exponential cumulative distribution function at X.
'
' WHY THIS EXISTS
'   The Exponential CDF is 1 - Exp(-Lambda * X). Evaluated naively that is
'   exactly 0 for small X once Exp(-Lambda * X) rounds to 1; computing it as
'   -PROB_Expm1(-Lambda * X) keeps full relative precision all the way down to
'   X near zero, which matters for tiny survival complements and for the
'   Chi-square(2)-to-Exponential identity check.
'
' WORKSHEET EQUIVALENT
'   EXPON.DIST(X, Lambda, TRUE)
'
' INPUTS
'   X       Evaluation point. For X <= 0 the CDF is 0.
'   Lambda  Rate parameter. Must be strictly positive.
'   Status  Optional ByRef diagnostic message.
'
' RETURNS
'   Variant
'     Success => Double cumulative probability 1 - Exp(-Lambda * X).
'     Failure => CVErr(xlErrNum) or CVErr(xlErrValue).
'
' BEHAVIOR
'   - Returns 0 for X <= 0.
'   - Otherwise -PROB_Expm1(-Lambda * X).
'
' ERROR POLICY
'   - Invalid numeric domains return CVErr(xlErrNum).
'   - Unexpected runtime errors return CVErr(xlErrValue).
'   - Detailed diagnostic messages are written to Status.
'   - No MsgBox is raised.
'
' DEPENDENCIES
'   - PROB_CN_ValidateXLambda
'   - PROB_Expm1, PROB_SetStatus
'
' UPDATED
'   2026-07-11
'==============================================================================
'
'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
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
    'Validate the evaluation point and the positive rate
        If Not PROB_CN_ValidateXLambda(X, Lambda, FailMsg) Then GoTo Fail_Num

'------------------------------------------------------------------------------
' COMPUTE CUMULATIVE PROBABILITY
'------------------------------------------------------------------------------
    'Return zero for values outside the positive support
        If X <= 0# Then
            K_STATS_Exponential_Cumulative = 0#
            GoTo Return_Success
        End If

    'Compute 1 - Exp(-Lambda * X) without cancellation
        K_STATS_Exponential_Cumulative = -PROB_Expm1(-Lambda * X)

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
        K_STATS_Exponential_Cumulative = CVErr(xlErrNum)
    'Exit before the error handler
        Exit Function

'------------------------------------------------------------------------------
' ERROR HANDLER
'------------------------------------------------------------------------------
Err_Handler:
    'Write unexpected runtime errors to diagnostics
        PROB_SetStatus Status, "Unexpected error in K_STATS_Exponential_Cumulative: " & Err.Description
    'Return worksheet value error
        K_STATS_Exponential_Cumulative = CVErr(xlErrValue)
End Function


Public Function K_STATS_Exponential_Survival( _
    ByVal X As Double, _
    ByVal Lambda As Double, _
    Optional ByRef Status As String = "") _
    As Variant
'
'==============================================================================
' K_STATS_Exponential_Survival
'------------------------------------------------------------------------------
' PURPOSE
'   Returns the right-tail Exponential survival function Exp(-Lambda * X).
'
' WHY THIS EXISTS
'   The survival is the exact memoryless reliability function and is the natural
'   quantity for exceedance and time-to-failure work. It is a bare exponential,
'   so it needs no cancellation-avoiding trick: the argument is never positive.
'
' WORKSHEET EQUIVALENT
'   1 - EXPON.DIST(X, Lambda, TRUE)
'
' INPUTS
'   X       Evaluation point. For X <= 0 the survival is 1.
'   Lambda  Rate parameter. Must be strictly positive.
'   Status  Optional ByRef diagnostic message.
'
' RETURNS
'   Variant
'     Success => Double survival probability Exp(-Lambda * X).
'     Failure => CVErr(xlErrNum) or CVErr(xlErrValue).
'
' BEHAVIOR
'   - Returns 1 for X <= 0.
'   - Otherwise Exp(-Lambda * X); underflow to 0 in the far tail is a valid zero.
'
' ERROR POLICY
'   - Invalid numeric domains return CVErr(xlErrNum).
'   - Unexpected runtime errors return CVErr(xlErrValue).
'   - Detailed diagnostic messages are written to Status.
'   - No MsgBox is raised.
'
' DEPENDENCIES
'   - PROB_CN_ValidateXLambda
'   - PROB_SetStatus
'
' UPDATED
'   2026-07-11
'==============================================================================
'
'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
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
    'Validate the evaluation point and the positive rate
        If Not PROB_CN_ValidateXLambda(X, Lambda, FailMsg) Then GoTo Fail_Num

'------------------------------------------------------------------------------
' COMPUTE SURVIVAL PROBABILITY
'------------------------------------------------------------------------------
    'Return one for values outside the positive support
        If X <= 0# Then
            K_STATS_Exponential_Survival = 1#
            GoTo Return_Success
        End If

    'Compute Exp(-Lambda * X); the argument is non-positive, so no overflow
        K_STATS_Exponential_Survival = Exp(-Lambda * X)

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
        K_STATS_Exponential_Survival = CVErr(xlErrNum)
    'Exit before the error handler
        Exit Function

'------------------------------------------------------------------------------
' ERROR HANDLER
'------------------------------------------------------------------------------
Err_Handler:
    'Write unexpected runtime errors to diagnostics
        PROB_SetStatus Status, "Unexpected error in K_STATS_Exponential_Survival: " & Err.Description
    'Return worksheet value error
        K_STATS_Exponential_Survival = CVErr(xlErrValue)
End Function


Public Function K_STATS_Exponential_InverseCumulative( _
    ByVal Probability As Double, _
    ByVal Lambda As Double, _
    Optional ByRef Status As String = "") _
    As Variant
'
'==============================================================================
' K_STATS_Exponential_InverseCumulative
'------------------------------------------------------------------------------
' PURPOSE
'   Returns the Exponential quantile: the X for which the CDF equals Probability.
'
' WHY THIS EXISTS
'   The quantile is -Log(1 - Probability) / Lambda. Evaluated as
'   -PROB_Log1p(-Probability) / Lambda it keeps full relative precision as
'   Probability approaches 0, the mirror of the PROB_Expm1 treatment in the CDF.
'
' WORKSHEET EQUIVALENT
'   (none; -LN(1 - Probability) / Lambda)
'
' INPUTS
'   Probability  Target cumulative probability, strictly between 0 and 1.
'   Lambda       Rate parameter. Must be strictly positive.
'   Status       Optional ByRef diagnostic message.
'
' RETURNS
'   Variant
'     Success => Double quantile X >= 0.
'     Failure => CVErr(xlErrNum) or CVErr(xlErrValue).
'
' BEHAVIOR
'   - Returns -PROB_Log1p(-Probability) / Lambda.
'
' ERROR POLICY
'   - Probability outside (0, 1) or invalid Lambda return CVErr(xlErrNum).
'   - Unexpected runtime errors return CVErr(xlErrValue).
'   - Detailed diagnostic messages are written to Status.
'   - No MsgBox is raised.
'
' DEPENDENCIES
'   - PROB_IsValidProbabilityOpen, PROB_IsPositiveFinite
'   - PROB_Log1p, PROB_SetStatus
'
' UPDATED
'   2026-07-11
'==============================================================================
'
'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
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
    'Validate rate
        If Not PROB_IsPositiveFinite(Lambda) Then
            FailMsg = "Lambda must be a finite strictly positive number"
            GoTo Fail_Num
        End If

'------------------------------------------------------------------------------
' COMPUTE QUANTILE
'------------------------------------------------------------------------------
    'Compute -Log(1 - Probability) / Lambda without cancellation
        K_STATS_Exponential_InverseCumulative = -PROB_Log1p(-Probability) / Lambda

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
        K_STATS_Exponential_InverseCumulative = CVErr(xlErrNum)
    'Exit before the error handler
        Exit Function

'------------------------------------------------------------------------------
' ERROR HANDLER
'------------------------------------------------------------------------------
Err_Handler:
    'Write unexpected runtime errors to diagnostics
        PROB_SetStatus Status, "Unexpected error in K_STATS_Exponential_InverseCumulative: " & Err.Description
    'Return worksheet value error
        K_STATS_Exponential_InverseCumulative = CVErr(xlErrValue)
End Function


'==============================================================================
' WEIBULL DISTRIBUTION  (Shape, ScaleParam = Excel alpha, beta)
'==============================================================================

Public Function K_STATS_Weibull_Density( _
    ByVal X As Double, _
    ByVal Shape As Double, _
    ByVal ScaleParam As Double, _
    Optional ByRef Status As String = "") _
    As Variant
'
'==============================================================================
' K_STATS_Weibull_Density
'------------------------------------------------------------------------------
' PURPOSE
'   Returns the Weibull probability density at X for shape k and scale lambda.
'
' WHY THIS EXISTS
'   The Weibull is the default life-data and time-to-failure model. The reduced
'   variable z = (X / ScaleParam) ^ Shape is formed through PROB_TryExp so that a huge
'   z degrades gracefully to a zero density instead of raising an overflow, and
'   the density itself is taken from the log-density for the same reason.
'
' WORKSHEET EQUIVALENT
'   WEIBULL.DIST(X, Shape, ScaleParam, FALSE)
'
' INPUTS
'   X       Evaluation point. For X < 0 the density is 0.
'   Shape   Shape parameter k (Excel Alpha). Must be strictly positive.
'   ScaleParam   ScaleParam parameter lambda (Excel Beta). Must be strictly positive.
'   Status  Optional ByRef diagnostic message.
'
' RETURNS
'   Variant
'     Success => Double density value.
'     Failure => CVErr(xlErrNum) or CVErr(xlErrValue).
'
' BEHAVIOR
'   - Returns 0 for X < 0.
'   - At X = 0 the density is unbounded when Shape < 1: returns CVErr(xlErrNum).
'     When Shape = 1 it equals 1 / ScaleParam; when Shape > 1 it equals 0.
'   - When z = (X / ScaleParam) ^ Shape overflows, the density underflows to 0.
'   - Otherwise Exp(Log(k) - Log(lambda) + (k-1)*(Log(X)-Log(lambda)) - z).
'
' ERROR POLICY
'   - Invalid numeric domains return CVErr(xlErrNum).
'   - Density pole or overflow returns CVErr(xlErrNum).
'   - Unexpected runtime errors return CVErr(xlErrValue).
'   - Detailed diagnostic messages are written to Status.
'   - No MsgBox is raised.
'
' DEPENDENCIES
'   - PROB_CN_ValidateXAndTwoPositive
'   - PROB_TryExp, PROB_SetStatus
'
' UPDATED
'   2026-07-11
'==============================================================================
'
'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Dim LogRatio            As Double          'Log(X) - Log(ScaleParam)
    Dim Z                   As Double          '(X / ScaleParam) ^ Shape
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
    'Validate the evaluation point and both positive parameters
        If Not PROB_CN_ValidateXAndTwoPositive( _
            X, Shape, ScaleParam, FailMsg, "Shape", "ScaleParam") Then GoTo Fail_Num

'------------------------------------------------------------------------------
' HANDLE THE SUPPORT EDGE
'------------------------------------------------------------------------------
    'Return zero below the support
        If X < 0# Then
            K_STATS_Weibull_Density = 0#
            GoTo Return_Success
        End If

    'Handle the origin, where the density is 0, 1 / ScaleParam or unbounded
        If X = 0# Then
            If Shape < 1# Then
                FailMsg = "Weibull density is unbounded at X = 0 when Shape < 1"
                GoTo Fail_Num
            ElseIf Shape = 1# Then
                K_STATS_Weibull_Density = 1# / ScaleParam
            Else
                K_STATS_Weibull_Density = 0#
            End If
            GoTo Return_Success
        End If

'------------------------------------------------------------------------------
' COMPUTE DENSITY
'------------------------------------------------------------------------------
    'Form the log of the scaled variable
        LogRatio = Log(X) - Log(ScaleParam)

    'Form z = (X / ScaleParam) ^ Shape; overflow means the density underflows to 0
        If Not PROB_TryExp(Shape * LogRatio, Z) Then
            K_STATS_Weibull_Density = 0#
            GoTo Return_Success
        End If

    'Exponentiate the log-density; underflow to zero is a valid result
        If Not PROB_TryExp( _
            Log(Shape) - Log(ScaleParam) + (Shape - 1#) * LogRatio - Z, _
            Density) Then
            FailMsg = "Weibull density overflowed a Double"
            GoTo Fail_Num
        End If

    'Return the density
        K_STATS_Weibull_Density = Density

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
        K_STATS_Weibull_Density = CVErr(xlErrNum)
    'Exit before the error handler
        Exit Function

'------------------------------------------------------------------------------
' ERROR HANDLER
'------------------------------------------------------------------------------
Err_Handler:
    'Write unexpected runtime errors to diagnostics
        PROB_SetStatus Status, "Unexpected error in K_STATS_Weibull_Density: " & Err.Description
    'Return worksheet value error
        K_STATS_Weibull_Density = CVErr(xlErrValue)
End Function


Public Function K_STATS_Weibull_Cumulative( _
    ByVal X As Double, _
    ByVal Shape As Double, _
    ByVal ScaleParam As Double, _
    Optional ByRef Status As String = "") _
    As Variant
'
'==============================================================================
' K_STATS_Weibull_Cumulative
'------------------------------------------------------------------------------
' PURPOSE
'   Returns the left-tail Weibull cumulative distribution function at X.
'
' WHY THIS EXISTS
'   The Weibull CDF is 1 - Exp(-(X / ScaleParam) ^ Shape). Computing it as
'   -PROB_Expm1(-(X / ScaleParam) ^ Shape) keeps the small-X tail correct to full
'   relative precision instead of rounding it to exactly zero.
'
' WORKSHEET EQUIVALENT
'   WEIBULL.DIST(X, Shape, ScaleParam, TRUE)
'
' INPUTS
'   X       Evaluation point. For X <= 0 the CDF is 0.
'   Shape   Shape parameter k. Must be strictly positive.
'   ScaleParam   ScaleParam parameter lambda. Must be strictly positive.
'   Status  Optional ByRef diagnostic message.
'
' RETURNS
'   Variant
'     Success => Double cumulative probability.
'     Failure => CVErr(xlErrNum) or CVErr(xlErrValue).
'
' BEHAVIOR
'   - Returns 0 for X <= 0.
'   - When z = (X / ScaleParam) ^ Shape overflows, the CDF is 1.
'   - Otherwise -PROB_Expm1(-z).
'
' ERROR POLICY
'   - Invalid numeric domains return CVErr(xlErrNum).
'   - Unexpected runtime errors return CVErr(xlErrValue).
'   - Detailed diagnostic messages are written to Status.
'   - No MsgBox is raised.
'
' DEPENDENCIES
'   - PROB_CN_ValidateXAndTwoPositive
'   - PROB_TryExp, PROB_Expm1, PROB_SetStatus
'
' UPDATED
'   2026-07-11
'==============================================================================
'
'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Dim Z                   As Double          '(X / ScaleParam) ^ Shape
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
    'Validate the evaluation point and both positive parameters
        If Not PROB_CN_ValidateXAndTwoPositive( _
            X, Shape, ScaleParam, FailMsg, "Shape", "ScaleParam") Then GoTo Fail_Num

'------------------------------------------------------------------------------
' COMPUTE CUMULATIVE PROBABILITY
'------------------------------------------------------------------------------
    'Return zero for values outside the positive support
        If X <= 0# Then
            K_STATS_Weibull_Cumulative = 0#
            GoTo Return_Success
        End If

    'Form z = (X / ScaleParam) ^ Shape; overflow saturates the CDF at 1
        If Not PROB_TryExp(Shape * (Log(X) - Log(ScaleParam)), Z) Then
            K_STATS_Weibull_Cumulative = 1#
            GoTo Return_Success
        End If

    'Compute 1 - Exp(-z) without cancellation
        K_STATS_Weibull_Cumulative = -PROB_Expm1(-Z)

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
        K_STATS_Weibull_Cumulative = CVErr(xlErrNum)
    'Exit before the error handler
        Exit Function

'------------------------------------------------------------------------------
' ERROR HANDLER
'------------------------------------------------------------------------------
Err_Handler:
    'Write unexpected runtime errors to diagnostics
        PROB_SetStatus Status, "Unexpected error in K_STATS_Weibull_Cumulative: " & Err.Description
    'Return worksheet value error
        K_STATS_Weibull_Cumulative = CVErr(xlErrValue)
End Function


Public Function K_STATS_Weibull_Survival( _
    ByVal X As Double, _
    ByVal Shape As Double, _
    ByVal ScaleParam As Double, _
    Optional ByRef Status As String = "") _
    As Variant
'
'==============================================================================
' K_STATS_Weibull_Survival
'------------------------------------------------------------------------------
' PURPOSE
'   Returns the right-tail Weibull survival function Exp(-(X / ScaleParam) ^ Shape).
'
' WHY THIS EXISTS
'   The survival function is the Weibull reliability function directly, and it is
'   a bare exponential of a non-positive argument, so it needs no cancellation
'   trick and cannot overflow.
'
' WORKSHEET EQUIVALENT
'   1 - WEIBULL.DIST(X, Shape, ScaleParam, TRUE)
'
' INPUTS
'   X       Evaluation point. For X <= 0 the survival is 1.
'   Shape   Shape parameter k. Must be strictly positive.
'   ScaleParam   ScaleParam parameter lambda. Must be strictly positive.
'   Status  Optional ByRef diagnostic message.
'
' RETURNS
'   Variant
'     Success => Double survival probability Exp(-(X / ScaleParam) ^ Shape).
'     Failure => CVErr(xlErrNum) or CVErr(xlErrValue).
'
' BEHAVIOR
'   - Returns 1 for X <= 0.
'   - When z = (X / ScaleParam) ^ Shape overflows, the survival is 0.
'   - Otherwise Exp(-z); underflow to 0 in the far tail is a valid zero.
'
' ERROR POLICY
'   - Invalid numeric domains return CVErr(xlErrNum).
'   - Unexpected runtime errors return CVErr(xlErrValue).
'   - Detailed diagnostic messages are written to Status.
'   - No MsgBox is raised.
'
' DEPENDENCIES
'   - PROB_CN_ValidateXAndTwoPositive
'   - PROB_TryExp, PROB_SetStatus
'
' UPDATED
'   2026-07-11
'==============================================================================
'
'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Dim Z                   As Double          '(X / ScaleParam) ^ Shape
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
    'Validate the evaluation point and both positive parameters
        If Not PROB_CN_ValidateXAndTwoPositive( _
            X, Shape, ScaleParam, FailMsg, "Shape", "ScaleParam") Then GoTo Fail_Num

'------------------------------------------------------------------------------
' COMPUTE SURVIVAL PROBABILITY
'------------------------------------------------------------------------------
    'Return one for values outside the positive support
        If X <= 0# Then
            K_STATS_Weibull_Survival = 1#
            GoTo Return_Success
        End If

    'Form z = (X / ScaleParam) ^ Shape; overflow drives the survival to 0
        If Not PROB_TryExp(Shape * (Log(X) - Log(ScaleParam)), Z) Then
            K_STATS_Weibull_Survival = 0#
            GoTo Return_Success
        End If

    'Compute Exp(-z); the argument is non-positive, so no overflow
        K_STATS_Weibull_Survival = Exp(-Z)

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
        K_STATS_Weibull_Survival = CVErr(xlErrNum)
    'Exit before the error handler
        Exit Function

'------------------------------------------------------------------------------
' ERROR HANDLER
'------------------------------------------------------------------------------
Err_Handler:
    'Write unexpected runtime errors to diagnostics
        PROB_SetStatus Status, "Unexpected error in K_STATS_Weibull_Survival: " & Err.Description
    'Return worksheet value error
        K_STATS_Weibull_Survival = CVErr(xlErrValue)
End Function


Public Function K_STATS_Weibull_InverseCumulative( _
    ByVal Probability As Double, _
    ByVal Shape As Double, _
    ByVal ScaleParam As Double, _
    Optional ByRef Status As String = "") _
    As Variant
'
'==============================================================================
' K_STATS_Weibull_InverseCumulative
'------------------------------------------------------------------------------
' PURPOSE
'   Returns the Weibull quantile: the X for which the CDF equals Probability.
'
' WHY THIS EXISTS
'   The quantile is ScaleParam * (-Log(1 - Probability)) ^ (1 / Shape). Taking the
'   inner logarithm through -PROB_Log1p(-Probability) keeps full precision as
'   Probability approaches 0.
'
' WORKSHEET EQUIVALENT
'   (none; ScaleParam * (-LN(1 - Probability)) ^ (1 / Shape))
'
' INPUTS
'   Probability  Target cumulative probability, strictly between 0 and 1.
'   Shape        Shape parameter k. Must be strictly positive.
'   ScaleParam        ScaleParam parameter lambda. Must be strictly positive.
'   Status       Optional ByRef diagnostic message.
'
' RETURNS
'   Variant
'     Success => Double quantile X >= 0.
'     Failure => CVErr(xlErrNum) or CVErr(xlErrValue).
'
' BEHAVIOR
'   - Returns ScaleParam * (-PROB_Log1p(-Probability)) ^ (1 / Shape).
'
' ERROR POLICY
'   - Probability outside (0, 1) or invalid parameters return CVErr(xlErrNum).
'   - Unexpected runtime errors return CVErr(xlErrValue).
'   - Detailed diagnostic messages are written to Status.
'   - No MsgBox is raised.
'
' DEPENDENCIES
'   - PROB_IsValidProbabilityOpen, PROB_IsPositiveFinite
'   - PROB_Log1p, PROB_SetStatus
'
' UPDATED
'   2026-07-11
'==============================================================================
'
'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Dim NegLogComplement    As Double          '-Log(1 - Probability)
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
    'Validate shape
        If Not PROB_IsPositiveFinite(Shape) Then
            FailMsg = "Shape must be a finite strictly positive number"
            GoTo Fail_Num
        End If
    'Validate scale
        If Not PROB_IsPositiveFinite(ScaleParam) Then
            FailMsg = "ScaleParam must be a finite strictly positive number"
            GoTo Fail_Num
        End If

'------------------------------------------------------------------------------
' COMPUTE QUANTILE
'------------------------------------------------------------------------------
    'Form -Log(1 - Probability) without cancellation
        NegLogComplement = -PROB_Log1p(-Probability)

    'Return ScaleParam * (-Log(1 - Probability)) ^ (1 / Shape)
        K_STATS_Weibull_InverseCumulative = ScaleParam * NegLogComplement ^ (1# / Shape)

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
        K_STATS_Weibull_InverseCumulative = CVErr(xlErrNum)
    'Exit before the error handler
        Exit Function

'------------------------------------------------------------------------------
' ERROR HANDLER
'------------------------------------------------------------------------------
Err_Handler:
    'Write unexpected runtime errors to diagnostics
        PROB_SetStatus Status, "Unexpected error in K_STATS_Weibull_InverseCumulative: " & Err.Description
    'Return worksheet value error
        K_STATS_Weibull_InverseCumulative = CVErr(xlErrValue)
End Function


Public Function K_STATS_Weibull_Mean( _
    ByVal Shape As Double, _
    ByVal ScaleParam As Double, _
    Optional ByRef Status As String = "") _
    As Variant
'
'==============================================================================
' K_STATS_Weibull_Mean
'------------------------------------------------------------------------------
' PURPOSE
'   Returns the mean of the Weibull distribution, ScaleParam * Gamma(1 + 1 / Shape).
'
' WORKSHEET EQUIVALENT
'   (none; ScaleParam * EXP(GAMMALN(1 + 1 / Shape)))
'
' INPUTS
'   Shape   Shape parameter k. Must be strictly positive.
'   ScaleParam   ScaleParam parameter lambda. Must be strictly positive.
'   Status  Optional ByRef diagnostic message.
'
' RETURNS
'   Variant
'     Success => Double mean.
'     Failure => CVErr(xlErrNum) or CVErr(xlErrValue).
'
' BEHAVIOR
'   - Evaluates the Gamma function through Exp(PROB_LogGamma(.)) with an overflow
'     guard, then guards the product with ScaleParam before forming it.
'
' ERROR POLICY
'   - Invalid parameters or overflow return CVErr(xlErrNum).
'   - Unexpected runtime errors return CVErr(xlErrValue).
'
' DEPENDENCIES
'   - PROB_IsPositiveFinite, PROB_LogGamma, PROB_TryExp, PROB_DOUBLE_MAX,
'     PROB_SetStatus
'
' UPDATED
'   2026-07-11
'==============================================================================
'
'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Dim G1                  As Double          'Gamma(1 + 1 / Shape)
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
    'Validate shape
        If Not PROB_IsPositiveFinite(Shape) Then
            FailMsg = "Shape must be a finite strictly positive number"
            GoTo Fail_Num
        End If
    'Validate scale
        If Not PROB_IsPositiveFinite(ScaleParam) Then
            FailMsg = "ScaleParam must be a finite strictly positive number"
            GoTo Fail_Num
        End If

'------------------------------------------------------------------------------
' COMPUTE MEAN
'------------------------------------------------------------------------------
    'Evaluate Gamma(1 + 1 / Shape) through its logarithm
        If Not PROB_TryExp(PROB_LogGamma(1# + 1# / Shape), G1) Then
            FailMsg = "Weibull mean overflows Double range in Gamma(1 + 1 / Shape)"
            GoTo Fail_Num
        End If

    'Guard the product against overflow, then form it
        If ScaleParam > PROB_DOUBLE_MAX / G1 Then
            FailMsg = "Weibull mean overflows Double range"
            GoTo Fail_Num
        End If

    'Return the mean
        K_STATS_Weibull_Mean = ScaleParam * G1

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
        K_STATS_Weibull_Mean = CVErr(xlErrNum)
    'Exit before the error handler
        Exit Function

'------------------------------------------------------------------------------
' ERROR HANDLER
'------------------------------------------------------------------------------
Err_Handler:
    'Write unexpected runtime errors to diagnostics
        PROB_SetStatus Status, "Unexpected error in K_STATS_Weibull_Mean: " & Err.Description
    'Return worksheet value error
        K_STATS_Weibull_Mean = CVErr(xlErrValue)
End Function


Public Function K_STATS_Weibull_Variance( _
    ByVal Shape As Double, _
    ByVal ScaleParam As Double, _
    Optional ByRef Status As String = "") _
    As Variant
'
'==============================================================================
' K_STATS_Weibull_Variance
'------------------------------------------------------------------------------
' PURPOSE
'   Returns the variance of the Weibull distribution,
'   ScaleParam ^ 2 * (Gamma(1 + 2 / Shape) - Gamma(1 + 1 / Shape) ^ 2).
'
' WORKSHEET EQUIVALENT
'   (none)
'
' INPUTS
'   Shape   Shape parameter k. Must be strictly positive.
'   ScaleParam   ScaleParam parameter lambda. Must be strictly positive.
'   Status  Optional ByRef diagnostic message.
'
' RETURNS
'   Variant
'     Success => Double variance.
'     Failure => CVErr(xlErrNum) or CVErr(xlErrValue).
'
' BEHAVIOR
'   - Both Gamma values are taken through Exp(PROB_LogGamma(.)) with overflow
'     guards. The variance factor Gamma(1+2/k) - Gamma(1+1/k)^2 is non-negative
'     in exact arithmetic; a tiny negative rounding result is clamped to 0.
'   - The ScaleParam ^ 2 multiplication is guarded, nested, against overflow.
'
' ERROR POLICY
'   - Invalid parameters or overflow return CVErr(xlErrNum).
'   - Unexpected runtime errors return CVErr(xlErrValue).
'
' DEPENDENCIES
'   - PROB_IsPositiveFinite, PROB_LogGamma, PROB_TryExp, PROB_DOUBLE_MAX,
'     PROB_SetStatus
'
' UPDATED
'   2026-07-11
'==============================================================================
'
'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Dim G1                  As Double          'Gamma(1 + 1 / Shape)
    Dim G2                  As Double          'Gamma(1 + 2 / Shape)
    Dim Factor              As Double          'G2 - G1 ^ 2, the shape variance
    Dim Partial             As Double          'ScaleParam * ScaleParam, held for the guard
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
    'Validate shape
        If Not PROB_IsPositiveFinite(Shape) Then
            FailMsg = "Shape must be a finite strictly positive number"
            GoTo Fail_Num
        End If
    'Validate scale
        If Not PROB_IsPositiveFinite(ScaleParam) Then
            FailMsg = "ScaleParam must be a finite strictly positive number"
            GoTo Fail_Num
        End If

'------------------------------------------------------------------------------
' COMPUTE VARIANCE
'------------------------------------------------------------------------------
    'Evaluate Gamma(1 + 1 / Shape) and Gamma(1 + 2 / Shape) through logarithms
        If Not PROB_TryExp(PROB_LogGamma(1# + 1# / Shape), G1) Then
            FailMsg = "Weibull variance overflows Double range in Gamma(1 + 1 / Shape)"
            GoTo Fail_Num
        End If
        If Not PROB_TryExp(PROB_LogGamma(1# + 2# / Shape), G2) Then
            FailMsg = "Weibull variance overflows Double range in Gamma(1 + 2 / Shape)"
            GoTo Fail_Num
        End If

    'Form the shape variance factor, clamping a tiny negative rounding to zero
        Factor = G2 - G1 * G1
        If Factor < 0# Then Factor = 0#

    'Guard and form ScaleParam * ScaleParam
        If ScaleParam > PROB_DOUBLE_MAX / ScaleParam Then
            FailMsg = "Weibull variance overflows Double range"
            GoTo Fail_Num
        End If
        Partial = ScaleParam * ScaleParam

    'Guard and form (ScaleParam * ScaleParam) * Factor
        If Factor > 0# And Partial > PROB_DOUBLE_MAX / Factor Then
            FailMsg = "Weibull variance overflows Double range"
            GoTo Fail_Num
        End If

    'Return the variance
        K_STATS_Weibull_Variance = Partial * Factor

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
        K_STATS_Weibull_Variance = CVErr(xlErrNum)
    'Exit before the error handler
        Exit Function

'------------------------------------------------------------------------------
' ERROR HANDLER
'------------------------------------------------------------------------------
Err_Handler:
    'Write unexpected runtime errors to diagnostics
        PROB_SetStatus Status, "Unexpected error in K_STATS_Weibull_Variance: " & Err.Description
    'Return worksheet value error
        K_STATS_Weibull_Variance = CVErr(xlErrValue)
End Function


Public Function K_STATS_Weibull_StdDev( _
    ByVal Shape As Double, _
    ByVal ScaleParam As Double, _
    Optional ByRef Status As String = "") _
    As Variant
'
'==============================================================================
' K_STATS_Weibull_StdDev
'------------------------------------------------------------------------------
' PURPOSE
'   Returns the standard deviation of the Weibull distribution,
'   ScaleParam * Sqr(Gamma(1 + 2 / Shape) - Gamma(1 + 1 / Shape) ^ 2).
'
' WORKSHEET EQUIVALENT
'   (none)
'
' INPUTS
'   Shape   Shape parameter k. Must be strictly positive.
'   ScaleParam   ScaleParam parameter lambda. Must be strictly positive.
'   Status  Optional ByRef diagnostic message.
'
' RETURNS
'   Variant
'     Success => Double standard deviation.
'     Failure => CVErr(xlErrNum) or CVErr(xlErrValue).
'
' BEHAVIOR
'   - Uses the reduced-magnitude form ScaleParam * Sqr(shape variance factor) rather
'     than Sqr(Variance), so it stays finite where the variance itself would
'     overflow.
'
' ERROR POLICY
'   - Invalid parameters or overflow return CVErr(xlErrNum).
'   - Unexpected runtime errors return CVErr(xlErrValue).
'
' DEPENDENCIES
'   - PROB_IsPositiveFinite, PROB_LogGamma, PROB_TryExp, PROB_DOUBLE_MAX,
'     PROB_SetStatus
'
' UPDATED
'   2026-07-11
'==============================================================================
'
'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Dim G1                  As Double          'Gamma(1 + 1 / Shape)
    Dim G2                  As Double          'Gamma(1 + 2 / Shape)
    Dim RootFactor          As Double          'Sqr(G2 - G1 ^ 2)
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
    'Validate shape
        If Not PROB_IsPositiveFinite(Shape) Then
            FailMsg = "Shape must be a finite strictly positive number"
            GoTo Fail_Num
        End If
    'Validate scale
        If Not PROB_IsPositiveFinite(ScaleParam) Then
            FailMsg = "ScaleParam must be a finite strictly positive number"
            GoTo Fail_Num
        End If

'------------------------------------------------------------------------------
' COMPUTE STANDARD DEVIATION
'------------------------------------------------------------------------------
    'Evaluate Gamma(1 + 1 / Shape) and Gamma(1 + 2 / Shape) through logarithms
        If Not PROB_TryExp(PROB_LogGamma(1# + 1# / Shape), G1) Then
            FailMsg = "Weibull standard deviation overflows Double range in Gamma(1 + 1 / Shape)"
            GoTo Fail_Num
        End If
        If Not PROB_TryExp(PROB_LogGamma(1# + 2# / Shape), G2) Then
            FailMsg = "Weibull standard deviation overflows Double range in Gamma(1 + 2 / Shape)"
            GoTo Fail_Num
        End If

    'Root of the shape variance factor, clamping a tiny negative rounding to zero
        RootFactor = G2 - G1 * G1
        If RootFactor < 0# Then RootFactor = 0#
        RootFactor = Sqr(RootFactor)

    'Guard the product against overflow, then form it
        If RootFactor > 0# And ScaleParam > PROB_DOUBLE_MAX / RootFactor Then
            FailMsg = "Weibull standard deviation overflows Double range"
            GoTo Fail_Num
        End If

    'Return the standard deviation
        K_STATS_Weibull_StdDev = ScaleParam * RootFactor

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
        K_STATS_Weibull_StdDev = CVErr(xlErrNum)
    'Exit before the error handler
        Exit Function

'------------------------------------------------------------------------------
' ERROR HANDLER
'------------------------------------------------------------------------------
Err_Handler:
    'Write unexpected runtime errors to diagnostics
        PROB_SetStatus Status, "Unexpected error in K_STATS_Weibull_StdDev: " & Err.Description
    'Return worksheet value error
        K_STATS_Weibull_StdDev = CVErr(xlErrValue)
End Function


'==============================================================================
' CONTINUOUS UNIFORM DISTRIBUTION  (LowerBound < UpperBound)
'==============================================================================

Public Function K_STATS_Uniform_Density( _
    ByVal X As Double, _
    ByVal LowerBound As Double, _
    ByVal UpperBound As Double, _
    Optional ByRef Status As String = "") _
    As Variant
'
'==============================================================================
' K_STATS_Uniform_Density
'------------------------------------------------------------------------------
' PURPOSE
'   Returns the continuous Uniform density at X on [LowerBound, UpperBound].
'
' WHY THIS EXISTS
'   The continuous Uniform is the flat reference distribution and the target of
'   the inverse-transform sampling identity. Its density is the constant
'   1 / (UpperBound - LowerBound) on the support and 0 outside it.
'
' WORKSHEET EQUIVALENT
'   (none)
'
' INPUTS
'   X            Evaluation point. Outside [LowerBound, UpperBound] density is 0.
'   LowerBound   Lower support bound.
'   UpperBound   Upper support bound. Must be strictly greater than LowerBound.
'   Status       Optional ByRef diagnostic message.
'
' RETURNS
'   Variant
'     Success => Double density value.
'     Failure => CVErr(xlErrNum) or CVErr(xlErrValue).
'
' BEHAVIOR
'   - Returns 0 for X < LowerBound or X > UpperBound.
'   - Otherwise 1 / (UpperBound - LowerBound).
'
' ERROR POLICY
'   - Non-finite bounds or UpperBound <= LowerBound return CVErr(xlErrNum).
'   - Unexpected runtime errors return CVErr(xlErrValue).
'   - Detailed diagnostic messages are written to Status.
'   - No MsgBox is raised.
'
' DEPENDENCIES
'   - PROB_CN_ValidateXBounds
'   - PROB_SetStatus
'
' UPDATED
'   2026-07-11
'==============================================================================
'
'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
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
    'Validate the evaluation point and the ordered bounds
        If Not PROB_CN_ValidateXBounds( _
            X, LowerBound, UpperBound, FailMsg) Then GoTo Fail_Num

'------------------------------------------------------------------------------
' COMPUTE DENSITY
'------------------------------------------------------------------------------
    'Return zero outside the closed support
        If X < LowerBound Or X > UpperBound Then
            K_STATS_Uniform_Density = 0#
            GoTo Return_Success
        End If

    'Return the constant density on the support
        K_STATS_Uniform_Density = 1# / (UpperBound - LowerBound)

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
        K_STATS_Uniform_Density = CVErr(xlErrNum)
    'Exit before the error handler
        Exit Function

'------------------------------------------------------------------------------
' ERROR HANDLER
'------------------------------------------------------------------------------
Err_Handler:
    'Write unexpected runtime errors to diagnostics
        PROB_SetStatus Status, "Unexpected error in K_STATS_Uniform_Density: " & Err.Description
    'Return worksheet value error
        K_STATS_Uniform_Density = CVErr(xlErrValue)
End Function


Public Function K_STATS_Uniform_Cumulative( _
    ByVal X As Double, _
    ByVal LowerBound As Double, _
    ByVal UpperBound As Double, _
    Optional ByRef Status As String = "") _
    As Variant
'
'==============================================================================
' K_STATS_Uniform_Cumulative
'------------------------------------------------------------------------------
' PURPOSE
'   Returns the continuous Uniform cumulative distribution function at X.
'
' WORKSHEET EQUIVALENT
'   (none)
'
' INPUTS
'   X            Evaluation point.
'   LowerBound   Lower support bound.
'   UpperBound   Upper support bound. Must be strictly greater than LowerBound.
'   Status       Optional ByRef diagnostic message.
'
' RETURNS
'   Variant
'     Success => Double cumulative probability in [0, 1].
'     Failure => CVErr(xlErrNum) or CVErr(xlErrValue).
'
' BEHAVIOR
'   - Returns 0 for X < LowerBound and 1 for X > UpperBound.
'   - Otherwise (X - LowerBound) / (UpperBound - LowerBound).
'
' ERROR POLICY
'   - Non-finite bounds or UpperBound <= LowerBound return CVErr(xlErrNum).
'   - Unexpected runtime errors return CVErr(xlErrValue).
'   - Detailed diagnostic messages are written to Status.
'   - No MsgBox is raised.
'
' DEPENDENCIES
'   - PROB_CN_ValidateXBounds
'   - PROB_SetStatus
'
' UPDATED
'   2026-07-11
'==============================================================================
'
'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
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
    'Validate the evaluation point and the ordered bounds
        If Not PROB_CN_ValidateXBounds( _
            X, LowerBound, UpperBound, FailMsg) Then GoTo Fail_Num

'------------------------------------------------------------------------------
' COMPUTE CUMULATIVE PROBABILITY
'------------------------------------------------------------------------------
    'Return the support edges exactly
        If X < LowerBound Then
            K_STATS_Uniform_Cumulative = 0#
            GoTo Return_Success
        End If
        If X > UpperBound Then
            K_STATS_Uniform_Cumulative = 1#
            GoTo Return_Success
        End If

    'Return the linear cumulative probability on the support
        K_STATS_Uniform_Cumulative = (X - LowerBound) / (UpperBound - LowerBound)

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
        K_STATS_Uniform_Cumulative = CVErr(xlErrNum)
    'Exit before the error handler
        Exit Function

'------------------------------------------------------------------------------
' ERROR HANDLER
'------------------------------------------------------------------------------
Err_Handler:
    'Write unexpected runtime errors to diagnostics
        PROB_SetStatus Status, "Unexpected error in K_STATS_Uniform_Cumulative: " & Err.Description
    'Return worksheet value error
        K_STATS_Uniform_Cumulative = CVErr(xlErrValue)
End Function


Public Function K_STATS_Uniform_Survival( _
    ByVal X As Double, _
    ByVal LowerBound As Double, _
    ByVal UpperBound As Double, _
    Optional ByRef Status As String = "") _
    As Variant
'
'==============================================================================
' K_STATS_Uniform_Survival
'------------------------------------------------------------------------------
' PURPOSE
'   Returns the continuous Uniform survival function 1 - CDF at X.
'
' WHY THIS EXISTS
'   The survival is computed directly as (UpperBound - X) / (UpperBound -
'   LowerBound) rather than 1 minus the CDF, so it is exact at both ends of the
'   support.
'
' WORKSHEET EQUIVALENT
'   (none)
'
' INPUTS
'   X            Evaluation point.
'   LowerBound   Lower support bound.
'   UpperBound   Upper support bound. Must be strictly greater than LowerBound.
'   Status       Optional ByRef diagnostic message.
'
' RETURNS
'   Variant
'     Success => Double survival probability in [0, 1].
'     Failure => CVErr(xlErrNum) or CVErr(xlErrValue).
'
' BEHAVIOR
'   - Returns 1 for X < LowerBound and 0 for X > UpperBound.
'   - Otherwise (UpperBound - X) / (UpperBound - LowerBound).
'
' ERROR POLICY
'   - Non-finite bounds or UpperBound <= LowerBound return CVErr(xlErrNum).
'   - Unexpected runtime errors return CVErr(xlErrValue).
'   - Detailed diagnostic messages are written to Status.
'   - No MsgBox is raised.
'
' DEPENDENCIES
'   - PROB_CN_ValidateXBounds
'   - PROB_SetStatus
'
' UPDATED
'   2026-07-11
'==============================================================================
'
'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
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
    'Validate the evaluation point and the ordered bounds
        If Not PROB_CN_ValidateXBounds( _
            X, LowerBound, UpperBound, FailMsg) Then GoTo Fail_Num

'------------------------------------------------------------------------------
' COMPUTE SURVIVAL PROBABILITY
'------------------------------------------------------------------------------
    'Return the support edges exactly
        If X < LowerBound Then
            K_STATS_Uniform_Survival = 1#
            GoTo Return_Success
        End If
        If X > UpperBound Then
            K_STATS_Uniform_Survival = 0#
            GoTo Return_Success
        End If

    'Return the linear survival probability on the support
        K_STATS_Uniform_Survival = (UpperBound - X) / (UpperBound - LowerBound)

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
        K_STATS_Uniform_Survival = CVErr(xlErrNum)
    'Exit before the error handler
        Exit Function

'------------------------------------------------------------------------------
' ERROR HANDLER
'------------------------------------------------------------------------------
Err_Handler:
    'Write unexpected runtime errors to diagnostics
        PROB_SetStatus Status, "Unexpected error in K_STATS_Uniform_Survival: " & Err.Description
    'Return worksheet value error
        K_STATS_Uniform_Survival = CVErr(xlErrValue)
End Function


Public Function K_STATS_Uniform_InverseCumulative( _
    ByVal Probability As Double, _
    ByVal LowerBound As Double, _
    ByVal UpperBound As Double, _
    Optional ByRef Status As String = "") _
    As Variant
'
'==============================================================================
' K_STATS_Uniform_InverseCumulative
'------------------------------------------------------------------------------
' PURPOSE
'   Returns the continuous Uniform quantile,
'   LowerBound + Probability * (UpperBound - LowerBound).
'
' WHY THIS EXISTS
'   This is the inverse-transform map that turns a Uniform(0, 1) draw into a draw
'   on any interval, and is the reason the Uniform is the seed of Monte Carlo.
'
' WORKSHEET EQUIVALENT
'   (none)
'
' INPUTS
'   Probability  Target cumulative probability, strictly between 0 and 1.
'   LowerBound   Lower support bound.
'   UpperBound   Upper support bound. Must be strictly greater than LowerBound.
'   Status       Optional ByRef diagnostic message.
'
' RETURNS
'   Variant
'     Success => Double quantile in (LowerBound, UpperBound).
'     Failure => CVErr(xlErrNum) or CVErr(xlErrValue).
'
' BEHAVIOR
'   - Returns LowerBound + Probability * (UpperBound - LowerBound).
'
' ERROR POLICY
'   - Probability outside (0, 1), non-finite bounds or UpperBound <= LowerBound
'     return CVErr(xlErrNum).
'   - Unexpected runtime errors return CVErr(xlErrValue).
'   - Detailed diagnostic messages are written to Status.
'   - No MsgBox is raised.
'
' DEPENDENCIES
'   - PROB_IsValidProbabilityOpen, PROB_CN_ValidateBounds
'   - PROB_SetStatus
'
' UPDATED
'   2026-07-11
'==============================================================================
'
'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
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
    'Validate the ordered bounds
        If Not PROB_CN_ValidateBounds( _
            LowerBound, UpperBound, FailMsg) Then GoTo Fail_Num

'------------------------------------------------------------------------------
' COMPUTE QUANTILE
'------------------------------------------------------------------------------
    'Return the linearly interpolated quantile
        K_STATS_Uniform_InverseCumulative = _
            LowerBound + Probability * (UpperBound - LowerBound)

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
        K_STATS_Uniform_InverseCumulative = CVErr(xlErrNum)
    'Exit before the error handler
        Exit Function

'------------------------------------------------------------------------------
' ERROR HANDLER
'------------------------------------------------------------------------------
Err_Handler:
    'Write unexpected runtime errors to diagnostics
        PROB_SetStatus Status, "Unexpected error in K_STATS_Uniform_InverseCumulative: " & Err.Description
    'Return worksheet value error
        K_STATS_Uniform_InverseCumulative = CVErr(xlErrValue)
End Function


'==============================================================================
' PRIVATE VALIDATORS
'==============================================================================

Private Function PROB_CN_ValidateXAndTwoPositive( _
    ByVal X As Double, _
    ByVal Param1 As Double, _
    ByVal Param2 As Double, _
    ByRef FailMsg As String, _
    ByVal Name1 As String, _
    ByVal Name2 As String) _
    As Boolean
'
'==============================================================================
' PROB_CN_ValidateXAndTwoPositive
'------------------------------------------------------------------------------
' PURPOSE
'   Validates a finite evaluation point and two strictly positive finite
'   parameters. Shared by the Gamma, Beta and Weibull families, whose two
'   parameters differ only in name (Shape/ScaleParam, Alpha/Beta, Shape/ScaleParam).
'
' RETURNS
'   True when X is finite and both parameters are finite and strictly positive;
'   otherwise False with FailMsg set to the first violation found.
'==============================================================================
'
'------------------------------------------------------------------------------
' VALIDATE
'------------------------------------------------------------------------------
    'Validate the evaluation point
        If Not PROB_IsFinite(X) Then
            FailMsg = "X must be a finite number"
            Exit Function
        End If

    'Validate the first parameter
        If Not PROB_IsPositiveFinite(Param1) Then
            FailMsg = Name1 & " must be a finite strictly positive number"
            Exit Function
        End If

    'Validate the second parameter
        If Not PROB_IsPositiveFinite(Param2) Then
            FailMsg = Name2 & " must be a finite strictly positive number"
            Exit Function
        End If

'------------------------------------------------------------------------------
' RETURN SUCCESS
'------------------------------------------------------------------------------
    'Report success
        PROB_CN_ValidateXAndTwoPositive = True
End Function


Private Function PROB_CN_ValidateXLambda( _
    ByVal X As Double, _
    ByVal Lambda As Double, _
    ByRef FailMsg As String) _
    As Boolean
'
'==============================================================================
' PROB_CN_ValidateXLambda
'------------------------------------------------------------------------------
' PURPOSE
'   Validates a finite evaluation point and a strictly positive finite rate for
'   the Exponential family.
'
' RETURNS
'   True when X is finite and Lambda is finite and strictly positive; otherwise
'   False with FailMsg set.
'==============================================================================
'
'------------------------------------------------------------------------------
' VALIDATE
'------------------------------------------------------------------------------
    'Validate the evaluation point
        If Not PROB_IsFinite(X) Then
            FailMsg = "X must be a finite number"
            Exit Function
        End If

    'Validate the rate
        If Not PROB_IsPositiveFinite(Lambda) Then
            FailMsg = "Lambda must be a finite strictly positive number"
            Exit Function
        End If

'------------------------------------------------------------------------------
' RETURN SUCCESS
'------------------------------------------------------------------------------
    'Report success
        PROB_CN_ValidateXLambda = True
End Function


Private Function PROB_CN_ValidateBounds( _
    ByVal LowerBound As Double, _
    ByVal UpperBound As Double, _
    ByRef FailMsg As String) _
    As Boolean
'
'==============================================================================
' PROB_CN_ValidateBounds
'------------------------------------------------------------------------------
' PURPOSE
'   Validates two finite bounds with UpperBound strictly greater than
'   LowerBound. Used by the Uniform family; the X-taking members reach it
'   through PROB_CN_ValidateXBounds, the inverse reaches it directly.
'
' RETURNS
'   True when both bounds are finite and UpperBound > LowerBound; otherwise
'   False with FailMsg set.
'==============================================================================
'
'------------------------------------------------------------------------------
' VALIDATE
'------------------------------------------------------------------------------
    'Validate the lower bound
        If Not PROB_IsFinite(LowerBound) Then
            FailMsg = "LowerBound must be a finite number"
            Exit Function
        End If

    'Validate the upper bound
        If Not PROB_IsFinite(UpperBound) Then
            FailMsg = "UpperBound must be a finite number"
            Exit Function
        End If

    'Require a non-degenerate, correctly ordered support
        If Not (UpperBound > LowerBound) Then
            FailMsg = "UpperBound must be strictly greater than LowerBound"
            Exit Function
        End If

'------------------------------------------------------------------------------
' RETURN SUCCESS
'------------------------------------------------------------------------------
    'Report success
        PROB_CN_ValidateBounds = True
End Function


Private Function PROB_CN_ValidateXBounds( _
    ByVal X As Double, _
    ByVal LowerBound As Double, _
    ByVal UpperBound As Double, _
    ByRef FailMsg As String) _
    As Boolean
'
'==============================================================================
' PROB_CN_ValidateXBounds
'------------------------------------------------------------------------------
' PURPOSE
'   Validates a finite evaluation point and two finite, correctly ordered bounds
'   for the Uniform family.
'
' RETURNS
'   True when X is finite and the bounds pass PROB_CN_ValidateBounds; otherwise
'   False with FailMsg set.
'==============================================================================
'
'------------------------------------------------------------------------------
' VALIDATE
'------------------------------------------------------------------------------
    'Validate the evaluation point
        If Not PROB_IsFinite(X) Then
            FailMsg = "X must be a finite number"
            Exit Function
        End If

    'Validate the ordered bounds
        If Not PROB_CN_ValidateBounds(LowerBound, UpperBound, FailMsg) Then Exit Function

'------------------------------------------------------------------------------
' RETURN SUCCESS
'------------------------------------------------------------------------------
    'Report success
        PROB_CN_ValidateXBounds = True
End Function


