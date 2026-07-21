Attribute VB_Name = "M_STATS_PROBDIST_CONTINUOUS"
Option Explicit

'==============================================================================
' M_STATS_PROBDIST_CONTINUOUS
'------------------------------------------------------------------------------
' PURPOSE
'   Provides worksheet-facing functions for the Gamma, Beta, Exponential,
'   Weibull and continuous Uniform distributions.
'
' WHY THIS EXISTS
'   These distributions complete the principal continuous-distribution layer
'   of the probability library outside the normal and Student-t families. The
'   module exposes a consistent worksheet API while delegating special-function
'   work to the shared numerical kernels.
'
' PUBLIC API
'   Gamma
'     K_STATS_Gamma_Density
'     K_STATS_Gamma_Cumulative
'     K_STATS_Gamma_Survival
'     K_STATS_Gamma_InverseCumulative
'     K_STATS_Gamma_Mean
'     K_STATS_Gamma_Variance
'     K_STATS_Gamma_StdDev
'
'   Beta
'     K_STATS_Beta_Density
'     K_STATS_Beta_Cumulative
'     K_STATS_Beta_Survival
'     K_STATS_Beta_InverseCumulative
'     K_STATS_Beta_Mean
'     K_STATS_Beta_Variance
'     K_STATS_Beta_StdDev
'
'   Exponential
'     K_STATS_Exponential_Density
'     K_STATS_Exponential_Cumulative
'     K_STATS_Exponential_Survival
'     K_STATS_Exponential_InverseCumulative
'
'   Weibull
'     K_STATS_Weibull_Density
'     K_STATS_Weibull_Cumulative
'     K_STATS_Weibull_Survival
'     K_STATS_Weibull_InverseCumulative
'     K_STATS_Weibull_Mean
'     K_STATS_Weibull_Variance
'     K_STATS_Weibull_StdDev
'
'   Continuous Uniform
'     K_STATS_Uniform_Density
'     K_STATS_Uniform_Cumulative
'     K_STATS_Uniform_Survival
'     K_STATS_Uniform_InverseCumulative
'
' PARAMETERIZATION
'   The public signatures follow the corresponding Excel worksheet conventions:
'
'     Gamma(X, Shape, ScaleParam)
'       Shape is the Gamma shape parameter.
'       ScaleParam is the Gamma scale parameter.
'
'     Beta(X, Alpha, Beta)
'       Alpha and Beta are positive shape parameters.
'
'     Exponential(X, Lambda)
'       Lambda is the rate, not the scale.
'
'     Weibull(X, Shape, ScaleParam)
'       Shape is the Weibull shape parameter.
'       ScaleParam is the Weibull scale parameter.
'
'     Uniform(X, LowerBound, UpperBound)
'       LowerBound and UpperBound define the finite support.
'
' NUMERICAL DESIGN
'   - Gamma CDF and survival use the regularized incomplete gamma functions.
'   - Beta CDF and survival use paired incomplete-beta arguments so the smaller
'     tail is evaluated directly rather than by subtraction from one.
'   - Exponential and Weibull CDFs use PROB_Expm1 in their left tails.
'   - Exponential and Weibull quantiles are assembled in the logarithmic domain.
'   - Weibull large-shape moments use a cancellation-free asymptotic expansion.
'   - Gamma, Exponential and Weibull intermediate arithmetic is guarded through
'     the shared Try-contract.
'   - Uniform calculations use scaled or convex-combination forms so the full
'     finite Double range can be accepted without overflowing interval widths.
'
' BETA ACCURACY REGIMES
'   Beta accuracy is regime-specific and is governed by the machine-readable
'   contract in benchmark/accuracy_contracts.csv (rendered in benchmark/README).
'   The regime is set by the shape ratio min(Alpha, Beta) / max(Alpha, Beta):
'   - Balanced shapes (ratio >= 0.1) retain the tight contract: density and
'     survival to 5E-15, CDF to 2E-14, inverse quantile to 5E-15 (relative).
'   - Strongly unbalanced shapes (ratio < 0.1) carry SEPARATE MEASURED thresholds,
'     validated on an independent holdout and frozen: density 4E-12, CDF 1E-10,
'     survival 2E-10, inverse quantile 1E-10, and inverse forward-tail residual
'     1E-9 (relative).
'   - In the unbalanced regime PROB_LogBeta forms Log(Beta) from a stable
'     log-gamma difference, so the catastrophic cancellation of the naive
'     three-log-gamma identity is removed.
'   - With that cancellation gone, residual unbalanced accuracy is dominated by
'     the downstream incomplete-beta evaluation (CDF, survival, inverse), not by
'     LogBeta normalization.
'
' DESIGN PRINCIPLES
'   - Public worksheet functions return Variant so failures can be represented
'     by worksheet error values.
'   - Invalid domains are rejected explicitly and are never silently repaired.
'   - Predictable numerical failure returns CVErr(xlErrNum).
'   - Unexpected runtime failure returns CVErr(xlErrValue).
'   - Mathematically valid underflow returns zero.
'   - Support-edge values are returned exactly where they are representable.
'   - No public function raises a MsgBox.
'
' ERROR POLICY
'   - Invalid parameters, density poles, non-convergence and predictable
'     overflow return CVErr(xlErrNum).
'   - Unexpected runtime errors return CVErr(xlErrValue).
'   - Detailed diagnostics are written to the optional Status argument.
'   - Application.StatusBar is not written by this module.
'
' DEPENDENCIES
'   - M_STATS_PROBDIST_CORE
'       PROB_IsFinite
'       PROB_IsPositiveFinite
'       PROB_IsPositiveWithinSupportedMagnitude
'       PROB_IsValidProbabilityOpen
'       PROB_TryAdd
'       PROB_TryMultiply
'       PROB_TryDivide
'       PROB_TryExp
'       PROB_Log1p
'       PROB_Expm1
'       PROB_SetStatus
'
'   - M_STATS_PROBDIST_SPECIALFUNCS
'       PROB_LogGamma
'       PROB_LogBeta
'       PROB_TryGammaRegularizedP
'       PROB_TryGammaRegularizedQ
'       PROB_TryGammaInvP
'       PROB_TryBetaRegularized
'       PROB_TryBetaInvRegularized
'
' NOTES
'   - Shape parameters passed to iterative or asymptotic kernels are constrained
'     by PROB_IsPositiveWithinSupportedMagnitude. That 1E100 bound is
'     representational, not a convergence guarantee: the incomplete-gamma and
'     incomplete-beta kernels converge over a smaller range (roughly 1E9 and 1E7;
'     see M_STATS_PROBDIST_SPECIALFUNCS). A shape between that range and 1E100 is
'     accepted, attempted, and returns a clean non-convergence error, not a wrong
'     answer.
'   - Evaluation points, rates, scales and Uniform bounds may use the full finite
'     Double range where the implemented formula remains numerically meaningful.
'
' UPDATED
'   2026-07-21
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
'   Returns the Gamma probability density at X.
'
' WHY THIS EXISTS
'   Gamma densities are used for positive skewed quantities, waiting times and
'   severity models. The calculation is performed in the logarithmic domain so
'   large or small parameter combinations do not form unstable powers directly.
'
' WORKSHEET EQUIVALENT
'   GAMMA.DIST(X, Shape, ScaleParam, FALSE)
'
' INPUTS
'   X           Evaluation point. Values below zero have density zero.
'   Shape       Positive Gamma shape parameter.
'   ScaleParam  Positive Gamma scale parameter.
'   Status      Optional ByRef diagnostic message.
'
' RETURNS
'   Variant
'     Success => Double density.
'     Failure => CVErr(xlErrNum) or CVErr(xlErrValue).
'
' BEHAVIOR
'   - Returns zero below the support.
'   - At X = 0, returns #NUM! for Shape < 1, 1 / ScaleParam for Shape = 1,
'     and zero for Shape > 1.
'   - If X / ScaleParam overflows, the exponential tail makes the density zero.
'   - Mathematically valid exponential underflow returns zero.
'
' ERROR POLICY
'   - Invalid parameters, a density pole or density overflow return #NUM!.
'   - Unexpected runtime errors return #VALUE!.
'   - Diagnostics are written to Status.
'
' DEPENDENCIES
'   - PROB_CN_ValidateXShapeScale
'   - PROB_TryDivide
'   - PROB_TryExp
'   - PROB_LogGamma
'   - PROB_SetStatus
'
' CALLED FROM
'   - Worksheet formulas
'   - M_STATS_PROBDIST_TEST
'
' UPDATED
'   2026-07-21
'==============================================================================
'
'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Dim StandardX           As Double          'X divided by ScaleParam
    Dim LogRatio            As Double          'Log(X / ScaleParam), formed safely
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
    'Validate X, Shape and ScaleParam under their distinct numerical contracts
        If Not PROB_CN_ValidateXShapeScale( _
            X, Shape, ScaleParam, FailMsg, "Shape", "ScaleParam") Then
            GoTo Fail_Num
        End If

'------------------------------------------------------------------------------
' HANDLE SUPPORT EDGES
'------------------------------------------------------------------------------
    'Return zero below the positive support
        If X < 0# Then
            K_STATS_Gamma_Density = 0#
            GoTo Return_Success
        End If

    'Handle the origin according to the Gamma shape
        If X = 0# Then
            If Shape < 1# Then
                FailMsg = _
                    "Gamma density is unbounded at X = 0 when Shape < 1"
                GoTo Fail_Num

            ElseIf Shape = 1# Then
                If Not PROB_TryDivide(1#, ScaleParam, Density) Then
                    FailMsg = "Gamma density overflows Double at X = 0"
                    GoTo Fail_Num
                End If

                K_STATS_Gamma_Density = Density

            Else
                K_STATS_Gamma_Density = 0#
            End If

            GoTo Return_Success
        End If

'------------------------------------------------------------------------------
' COMPUTE DENSITY
'------------------------------------------------------------------------------
    'Form the standardized Gamma variate with explicit overflow handling
        If Not PROB_TryDivide(X, ScaleParam, StandardX) Then
            K_STATS_Gamma_Density = 0#
            GoTo Return_Success
        End If

    'Form Log(X / ScaleParam) without dividing potentially extreme operands
        LogRatio = Log(X) - Log(ScaleParam)

    'Use the scale-separated log-density to reduce cancellation
        LogDensity = _
            (Shape - 1#) * LogRatio - _
            StandardX - _
            Log(ScaleParam) - _
            PROB_LogGamma(Shape)

    'Exponentiate under the shared overflow and underflow contract
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
'   Returns the Gamma left-tail cumulative probability at X.
'
' WORKSHEET EQUIVALENT
'   GAMMA.DIST(X, Shape, ScaleParam, TRUE)
'
' INPUTS
'   X           Evaluation point.
'   Shape       Positive Gamma shape parameter.
'   ScaleParam  Positive Gamma scale parameter.
'   Status      Optional ByRef diagnostic message.
'
' RETURNS
'   Variant
'     Success => Double cumulative probability in [0, 1].
'     Failure => CVErr(xlErrNum) or CVErr(xlErrValue).
'
' BEHAVIOR
'   - Returns zero for X <= 0.
'   - Evaluates P(Shape, X / ScaleParam) for positive X.
'   - If X / ScaleParam overflows, returns the limiting value one.
'
' ERROR POLICY
'   - Invalid parameters or kernel non-convergence return #NUM!.
'   - Unexpected runtime errors return #VALUE!.
'
' DEPENDENCIES
'   - PROB_CN_ValidateXShapeScale
'   - PROB_TryDivide
'   - PROB_TryGammaRegularizedP
'   - PROB_SetStatus
'
' CALLED FROM
'   - Worksheet formulas
'   - M_STATS_PROBDIST_TEST
'
' UPDATED
'   2026-07-21
'==============================================================================
'
'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Dim StandardX           As Double          'X divided by ScaleParam
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
    'Validate X, Shape and ScaleParam
        If Not PROB_CN_ValidateXShapeScale( _
            X, Shape, ScaleParam, FailMsg, "Shape", "ScaleParam") Then
            GoTo Fail_Num
        End If

'------------------------------------------------------------------------------
' HANDLE SUPPORT EDGE
'------------------------------------------------------------------------------
    'Return zero at and below the support origin
        If X <= 0# Then
            K_STATS_Gamma_Cumulative = 0#
            GoTo Return_Success
        End If

'------------------------------------------------------------------------------
' COMPUTE CUMULATIVE PROBABILITY
'------------------------------------------------------------------------------
    'Treat a positive standardized-ratio overflow as the +infinity limit
        If Not PROB_TryDivide(X, ScaleParam, StandardX) Then
            K_STATS_Gamma_Cumulative = 1#
            GoTo Return_Success
        End If

    'Evaluate the regularized lower incomplete gamma function
        If Not PROB_TryGammaRegularizedP( _
            Shape, StandardX, Value, FailMsg) Then
            GoTo Fail_Num
        End If

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
'   Returns the Gamma right-tail survival probability at X.
'
' WHY THIS EXISTS
'   The upper tail is evaluated directly through the regularized incomplete
'   gamma Q function so small probabilities are not lost to 1 minus CDF.
'
' WORKSHEET EQUIVALENT
'   1 - GAMMA.DIST(X, Shape, ScaleParam, TRUE)
'
' INPUTS
'   X           Evaluation point.
'   Shape       Positive Gamma shape parameter.
'   ScaleParam  Positive Gamma scale parameter.
'   Status      Optional ByRef diagnostic message.
'
' RETURNS
'   Variant
'     Success => Double survival probability in [0, 1].
'     Failure => CVErr(xlErrNum) or CVErr(xlErrValue).
'
' BEHAVIOR
'   - Returns one for X <= 0.
'   - Evaluates Q(Shape, X / ScaleParam) for positive X.
'   - If X / ScaleParam overflows, returns the limiting value zero.
'
' ERROR POLICY
'   - Invalid parameters or kernel non-convergence return #NUM!.
'   - Unexpected runtime errors return #VALUE!.
'
' DEPENDENCIES
'   - PROB_CN_ValidateXShapeScale
'   - PROB_TryDivide
'   - PROB_TryGammaRegularizedQ
'   - PROB_SetStatus
'
' CALLED FROM
'   - Worksheet formulas
'   - M_STATS_PROBDIST_TEST
'
' UPDATED
'   2026-07-21
'==============================================================================
'
'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Dim StandardX           As Double          'X divided by ScaleParam
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
    'Validate X, Shape and ScaleParam
        If Not PROB_CN_ValidateXShapeScale( _
            X, Shape, ScaleParam, FailMsg, "Shape", "ScaleParam") Then
            GoTo Fail_Num
        End If

'------------------------------------------------------------------------------
' HANDLE SUPPORT EDGE
'------------------------------------------------------------------------------
    'Return one at and below the support origin
        If X <= 0# Then
            K_STATS_Gamma_Survival = 1#
            GoTo Return_Success
        End If

'------------------------------------------------------------------------------
' COMPUTE SURVIVAL PROBABILITY
'------------------------------------------------------------------------------
    'Treat a positive standardized-ratio overflow as the +infinity limit
        If Not PROB_TryDivide(X, ScaleParam, StandardX) Then
            K_STATS_Gamma_Survival = 0#
            GoTo Return_Success
        End If

    'Evaluate the regularized upper incomplete gamma function
        If Not PROB_TryGammaRegularizedQ( _
            Shape, StandardX, Value, FailMsg) Then
            GoTo Fail_Num
        End If

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
'   Returns the Gamma quantile for Probability, Shape and ScaleParam.
'
' WHY THIS EXISTS
'   The numerical inversion is performed once on the unit-scale Gamma and the
'   result is rescaled under the shared multiplication Try-contract.
'
' WORKSHEET EQUIVALENT
'   GAMMA.INV(Probability, Shape, ScaleParam)
'
' INPUTS
'   Probability  Target cumulative probability in the open unit interval.
'   Shape        Positive Gamma shape parameter.
'   ScaleParam   Positive Gamma scale parameter.
'   Status       Optional ByRef diagnostic message.
'
' RETURNS
'   Variant
'     Success => Double quantile.
'     Failure => CVErr(xlErrNum) or CVErr(xlErrValue).
'
' BEHAVIOR
'   - Inverts the regularized lower incomplete gamma function on the unit scale.
'   - Rescales the unit quantile with explicit overflow classification.
'
' ERROR POLICY
'   - Invalid inputs, non-convergence or rescaling overflow return #NUM!.
'   - Unexpected runtime errors return #VALUE!.
'
' DEPENDENCIES
'   - PROB_IsValidProbabilityOpen
'   - PROB_IsPositiveWithinSupportedMagnitude
'   - PROB_IsPositiveFinite
'   - PROB_TryGammaInvP
'   - PROB_TryMultiply
'   - PROB_SetStatus
'
' CALLED FROM
'   - Worksheet formulas
'   - M_STATS_PROBDIST_TEST
'
' UPDATED
'   2026-07-21
'==============================================================================
'
'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Dim GammaQuantile       As Double          'Unit-scale gamma quantile
    Dim RescaledQuantile    As Double          'Quantile after applying ScaleParam
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
        If Not PROB_IsPositiveWithinSupportedMagnitude(Shape) Then
            FailMsg = "Shape must be a finite strictly positive number within the parameter-magnitude guard (< 1E100)"
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

    'Rescale from the unit-scale gamma with explicit overflow classification
        If Not PROB_TryMultiply(ScaleParam, GammaQuantile, RescaledQuantile) Then
            FailMsg = "Gamma quantile overflowed a Double"
            GoTo Fail_Num
        End If

        K_STATS_Gamma_InverseCumulative = RescaledQuantile

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
'   Returns the Gamma mean, Shape * ScaleParam.
'
' INPUTS
'   Shape       Positive Gamma shape parameter.
'   ScaleParam  Positive Gamma scale parameter.
'   Status      Optional ByRef diagnostic message.
'
' RETURNS
'   Variant
'     Success => Double mean.
'     Failure => CVErr(xlErrNum) or CVErr(xlErrValue).
'
' BEHAVIOR
'   - Forms the product through PROB_TryMultiply.
'
' ERROR POLICY
'   - Invalid inputs or product overflow return #NUM!.
'   - Unexpected runtime errors return #VALUE!.
'
' DEPENDENCIES
'   - PROB_IsPositiveWithinSupportedMagnitude
'   - PROB_IsPositiveFinite
'   - PROB_TryMultiply
'   - PROB_SetStatus
'
' CALLED FROM
'   - Worksheet formulas
'   - M_STATS_PROBDIST_TEST
'
' UPDATED
'   2026-07-21
'==============================================================================
'
'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Dim MeanValue           As Double          'Computed mean
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
        If Not PROB_IsPositiveWithinSupportedMagnitude(Shape) Then
            FailMsg = "Shape must be a finite strictly positive number within the parameter-magnitude guard (< 1E100)"
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
    'Guard the product through the shared arithmetic contract
        If Not PROB_TryMultiply(Shape, ScaleParam, MeanValue) Then
            FailMsg = "Gamma mean overflows Double range"
            GoTo Fail_Num
        End If

        K_STATS_Gamma_Mean = MeanValue

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
'   Returns the Gamma variance, Shape * ScaleParam squared.
'
' INPUTS
'   Shape       Positive Gamma shape parameter.
'   ScaleParam  Positive Gamma scale parameter.
'   Status      Optional ByRef diagnostic message.
'
' RETURNS
'   Variant
'     Success => Double variance.
'     Failure => CVErr(xlErrNum) or CVErr(xlErrValue).
'
' BEHAVIOR
'   - Guards each multiplication separately so no intermediate overflow escapes
'     into the unexpected-error handler.
'
' ERROR POLICY
'   - Invalid inputs or product overflow return #NUM!.
'   - Unexpected runtime errors return #VALUE!.
'
' DEPENDENCIES
'   - PROB_IsPositiveWithinSupportedMagnitude
'   - PROB_IsPositiveFinite
'   - PROB_TryMultiply
'   - PROB_SetStatus
'
' CALLED FROM
'   - Worksheet formulas
'   - M_STATS_PROBDIST_TEST
'
' UPDATED
'   2026-07-21
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
        If Not PROB_IsPositiveWithinSupportedMagnitude(Shape) Then
            FailMsg = "Shape must be a finite strictly positive number within the parameter-magnitude guard (< 1E100)"
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
    'Guard both multiplications through the shared arithmetic contract
        If Not PROB_TryMultiply(Shape, ScaleParam, Partial) Then
            FailMsg = "Gamma variance overflows Double range"
            GoTo Fail_Num
        End If

        If Not PROB_TryMultiply(Partial, ScaleParam, Partial) Then
            FailMsg = "Gamma variance overflows Double range"
            GoTo Fail_Num
        End If

        K_STATS_Gamma_Variance = Partial

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
'   Returns the Gamma standard deviation, ScaleParam * Sqr(Shape).
'
' INPUTS
'   Shape       Positive Gamma shape parameter.
'   ScaleParam  Positive Gamma scale parameter.
'   Status      Optional ByRef diagnostic message.
'
' RETURNS
'   Variant
'     Success => Double standard deviation.
'     Failure => CVErr(xlErrNum) or CVErr(xlErrValue).
'
' BEHAVIOR
'   - Uses the reduced-magnitude closed form rather than taking the square root
'     of a potentially overflowing variance.
'   - Guards the final multiplication explicitly.
'
' ERROR POLICY
'   - Invalid inputs or product overflow return #NUM!.
'   - Unexpected runtime errors return #VALUE!.
'
' DEPENDENCIES
'   - PROB_IsPositiveWithinSupportedMagnitude
'   - PROB_IsPositiveFinite
'   - PROB_TryMultiply
'   - PROB_SetStatus
'
' CALLED FROM
'   - Worksheet formulas
'   - M_STATS_PROBDIST_TEST
'
' UPDATED
'   2026-07-21
'==============================================================================
'
'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Dim RootShape           As Double          'Sqr(Shape)
    Dim StdDevValue         As Double          'Computed standard deviation
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
        If Not PROB_IsPositiveWithinSupportedMagnitude(Shape) Then
            FailMsg = "Shape must be a finite strictly positive number within the parameter-magnitude guard (< 1E100)"
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

    'Guard the product through the shared arithmetic contract
        If Not PROB_TryMultiply(ScaleParam, RootShape, StdDevValue) Then
            FailMsg = "Gamma standard deviation overflows Double range"
            GoTo Fail_Num
        End If

        K_STATS_Gamma_StdDev = StdDevValue

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
'   Returns the Beta probability density at X for shapes Alpha and Beta.
'
' WHY THIS EXISTS
'   The density is evaluated in the logarithmic domain through PROB_LogBeta so
'   large shape parameters do not form unstable Gamma-function ratios directly.
'
' WORKSHEET EQUIVALENT
'   BETA.DIST(X, Alpha, Beta, FALSE)
'
' INPUTS
'   X       Evaluation point. Values outside [0, 1] have density zero.
'   Alpha   Positive first Beta shape parameter.
'   Beta    Positive second Beta shape parameter.
'   Status  Optional ByRef diagnostic message.
'
' RETURNS
'   Variant
'     Success => Double density.
'     Failure => CVErr(xlErrNum) or CVErr(xlErrValue).
'
' BEHAVIOR
'   - Handles both endpoint poles explicitly.
'   - Uses PROB_Log1p(-X) near the right endpoint.
'   - Mathematically valid underflow returns zero.
'
' ERROR POLICY
'   - Invalid inputs, endpoint poles or density overflow return #NUM!.
'   - Unexpected runtime errors return #VALUE!.
'
' DEPENDENCIES
'   - PROB_CN_ValidateXTwoShapes
'   - PROB_Log1p
'   - PROB_LogBeta
'   - PROB_TryExp
'   - PROB_SetStatus
'
' CALLED FROM
'   - Worksheet formulas
'   - M_STATS_PROBDIST_TEST
'
' UPDATED
'   2026-07-21
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
        If Not PROB_CN_ValidateXTwoShapes( _
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
'   Returns the Beta left-tail cumulative probability at X.
'
' WORKSHEET EQUIVALENT
'   BETA.DIST(X, Alpha, Beta, TRUE)
'
' INPUTS
'   X       Evaluation point.
'   Alpha   Positive first Beta shape parameter.
'   Beta    Positive second Beta shape parameter.
'   Status  Optional ByRef diagnostic message.
'
' RETURNS
'   Variant
'     Success => Double cumulative probability in [0, 1].
'     Failure => CVErr(xlErrNum) or CVErr(xlErrValue).
'
' BEHAVIOR
'   - Returns zero for X <= 0 and one for X >= 1.
'   - Passes X and 1 - X as a paired argument to the incomplete-beta kernel.
'
' ERROR POLICY
'   - Invalid inputs or kernel non-convergence return #NUM!.
'   - Unexpected runtime errors return #VALUE!.
'
' DEPENDENCIES
'   - PROB_CN_ValidateXTwoShapes
'   - PROB_TryBetaRegularized
'   - PROB_SetStatus
'
' CALLED FROM
'   - Worksheet formulas
'   - M_STATS_PROBDIST_TEST
'
' UPDATED
'   2026-07-21
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
        If Not PROB_CN_ValidateXTwoShapes( _
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
'   Returns the Beta right-tail survival probability at X.
'
' WHY THIS EXISTS
'   The upper tail is evaluated as the swapped incomplete beta
'   I_(1-X)(Beta, Alpha), avoiding loss of significance from 1 minus CDF.
'
' WORKSHEET EQUIVALENT
'   1 - BETA.DIST(X, Alpha, Beta, TRUE)
'
' INPUTS
'   X       Evaluation point.
'   Alpha   Positive first Beta shape parameter.
'   Beta    Positive second Beta shape parameter.
'   Status  Optional ByRef diagnostic message.
'
' RETURNS
'   Variant
'     Success => Double survival probability in [0, 1].
'     Failure => CVErr(xlErrNum) or CVErr(xlErrValue).
'
' BEHAVIOR
'   - Returns one for X <= 0 and zero for X >= 1.
'   - Evaluates the smaller upper tail directly.
'
' ERROR POLICY
'   - Invalid inputs or kernel non-convergence return #NUM!.
'   - Unexpected runtime errors return #VALUE!.
'
' DEPENDENCIES
'   - PROB_CN_ValidateXTwoShapes
'   - PROB_TryBetaRegularized
'   - PROB_SetStatus
'
' CALLED FROM
'   - Worksheet formulas
'   - M_STATS_PROBDIST_TEST
'
' UPDATED
'   2026-07-21
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
        If Not PROB_CN_ValidateXTwoShapes( _
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
'   Returns the Beta quantile for Probability, Alpha and Beta.
'
' WORKSHEET EQUIVALENT
'   BETA.INV(Probability, Alpha, Beta)
'
' INPUTS
'   Probability  Target cumulative probability in the open unit interval.
'   Alpha        Positive first Beta shape parameter.
'   Beta         Positive second Beta shape parameter.
'   Status       Optional ByRef diagnostic message.
'
' RETURNS
'   Variant
'     Success => Double quantile in (0, 1).
'     Failure => CVErr(xlErrNum) or CVErr(xlErrValue).
'
' BEHAVIOR
'   - Uses the paired beta-inverse kernel, which returns the quantile and its
'     complement without reconstructing the smaller value by subtraction.
'
' ERROR POLICY
'   - Invalid inputs or non-convergence return #NUM!.
'   - Unexpected runtime errors return #VALUE!.
'
' DEPENDENCIES
'   - PROB_IsValidProbabilityOpen
'   - PROB_IsPositiveWithinSupportedMagnitude
'   - PROB_TryBetaInvRegularized
'   - PROB_SetStatus
'
' CALLED FROM
'   - Worksheet formulas
'   - M_STATS_PROBDIST_TEST
'
' UPDATED
'   2026-07-21
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
        If Not PROB_IsPositiveWithinSupportedMagnitude(Alpha) Then
            FailMsg = "Alpha must be a finite strictly positive number within the parameter-magnitude guard (< 1E100)"
            GoTo Fail_Num
        End If
    'Validate second shape
        If Not PROB_IsPositiveWithinSupportedMagnitude(Beta) Then
            FailMsg = "Beta must be a finite strictly positive number within the parameter-magnitude guard (< 1E100)"
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
'   Returns the Beta mean, Alpha / (Alpha + Beta).
'
' WHY THIS EXISTS
'   The implementation uses a ratio-scaled form rather than forming Alpha plus
'   Beta first. This preserves the mean for highly unbalanced shapes and avoids
'   an unnecessary intermediate-sum restriction.
'
' INPUTS
'   Alpha   Positive first Beta shape parameter.
'   Beta    Positive second Beta shape parameter.
'   Status  Optional ByRef diagnostic message.
'
' RETURNS
'   Variant
'     Success => Double mean in (0, 1).
'     Failure => CVErr(xlErrNum) or CVErr(xlErrValue).
'
' BEHAVIOR
'   - When Alpha >= Beta, evaluates 1 / (1 + Beta / Alpha).
'   - Otherwise evaluates (Alpha / Beta) / (1 + Alpha / Beta).
'
' ERROR POLICY
'   - Invalid shape parameters return #NUM!.
'   - Unexpected runtime errors return #VALUE!.
'
' DEPENDENCIES
'   - PROB_IsPositiveWithinSupportedMagnitude
'   - PROB_SetStatus
'
' CALLED FROM
'   - Worksheet formulas
'   - M_STATS_PROBDIST_TEST
'
' UPDATED
'   2026-07-21
'==============================================================================
'
'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Dim Ratio               As Double          'Smaller shape divided by larger shape
    Dim MeanValue           As Double          'Returned mean
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
    'Validate the first shape parameter
        If Not PROB_IsPositiveWithinSupportedMagnitude(Alpha) Then
            FailMsg = _
                "Alpha must be a finite strictly positive number within the parameter-magnitude guard (< 1E100)"
            GoTo Fail_Num
        End If

    'Validate the second shape parameter
        If Not PROB_IsPositiveWithinSupportedMagnitude(Beta) Then
            FailMsg = _
                "Beta must be a finite strictly positive number within the parameter-magnitude guard (< 1E100)"
            GoTo Fail_Num
        End If

'------------------------------------------------------------------------------
' COMPUTE MEAN
'------------------------------------------------------------------------------
    'Scale by the larger shape so the intermediate ratio is at most one
        If Alpha >= Beta Then
            Ratio = Beta / Alpha
            MeanValue = 1# / (1# + Ratio)
        Else
            Ratio = Alpha / Beta
            MeanValue = Ratio / (1# + Ratio)
        End If

    'Return the mean
        K_STATS_Beta_Mean = MeanValue

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
'   Returns the Beta variance.
'
' NUMERICAL METHOD
'   Uses Mean * (1 - Mean) / (Alpha + Beta + 1). The mean is formed through a
'   ratio-scaled expression and the denominator is assembled through guarded
'   addition. The direct product Alpha * Beta is never formed.
'
' INPUTS
'   Alpha   Positive first Beta shape parameter.
'   Beta    Positive second Beta shape parameter.
'   Status  Optional ByRef diagnostic message.
'
' RETURNS
'   Variant
'     Success => Double variance in (0, 0.25].
'     Failure => CVErr(xlErrNum) or CVErr(xlErrValue).
'
' ERROR POLICY
'   - Invalid shape parameters or denominator overflow return #NUM!.
'   - Mathematically valid underflow returns zero.
'   - Unexpected runtime errors return #VALUE!.
'
' DEPENDENCIES
'   - PROB_IsPositiveWithinSupportedMagnitude
'   - PROB_TryAdd
'   - PROB_SetStatus
'
' CALLED FROM
'   - Worksheet formulas
'   - M_STATS_PROBDIST_TEST
'
' UPDATED
'   2026-07-21
'==============================================================================
'
'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Dim ShapeSum            As Double          'Alpha plus Beta
    Dim Denominator         As Double          'Alpha plus Beta plus one
    Dim Ratio               As Double          'Smaller shape divided by larger shape
    Dim MeanValue           As Double          'Beta mean
    Dim VarianceValue       As Double          'Returned variance
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
    'Validate the first shape parameter
        If Not PROB_IsPositiveWithinSupportedMagnitude(Alpha) Then
            FailMsg = _
                "Alpha must be a finite strictly positive number within the parameter-magnitude guard (< 1E100)"
            GoTo Fail_Num
        End If

    'Validate the second shape parameter
        If Not PROB_IsPositiveWithinSupportedMagnitude(Beta) Then
            FailMsg = _
                "Beta must be a finite strictly positive number within the parameter-magnitude guard (< 1E100)"
            GoTo Fail_Num
        End If

'------------------------------------------------------------------------------
' COMPUTE VARIANCE
'------------------------------------------------------------------------------
    'Form the mean through a ratio no greater than one
        If Alpha >= Beta Then
            Ratio = Beta / Alpha
            MeanValue = 1# / (1# + Ratio)
        Else
            Ratio = Alpha / Beta
            MeanValue = Ratio / (1# + Ratio)
        End If

    'Form Alpha + Beta under the shared arithmetic contract
        If Not PROB_TryAdd(Alpha, Beta, ShapeSum) Then
            FailMsg = "Beta variance overflowed in Alpha + Beta"
            GoTo Fail_Num
        End If

    'Form Alpha + Beta + 1 under the shared arithmetic contract
        If Not PROB_TryAdd(ShapeSum, 1#, Denominator) Then
            FailMsg = "Beta variance denominator overflowed a Double"
            GoTo Fail_Num
        End If

    'Evaluate the bounded variance expression
        VarianceValue = _
            MeanValue * (1# - MeanValue) / Denominator

    'Return the variance
        K_STATS_Beta_Variance = VarianceValue

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
'   Returns the Beta standard deviation.
'
' NUMERICAL METHOD
'   Forms the stable Beta variance through a ratio-scaled mean and guarded
'   denominator, then takes its square root. The variance is bounded above by
'   one quarter, so the square root cannot overflow.
'
' INPUTS
'   Alpha   Positive first Beta shape parameter.
'   Beta    Positive second Beta shape parameter.
'   Status  Optional ByRef diagnostic message.
'
' RETURNS
'   Variant
'     Success => Double standard deviation.
'     Failure => CVErr(xlErrNum) or CVErr(xlErrValue).
'
' ERROR POLICY
'   - Invalid shape parameters or denominator overflow return #NUM!.
'   - Mathematically valid underflow returns zero.
'   - Unexpected runtime errors return #VALUE!.
'
' DEPENDENCIES
'   - PROB_IsPositiveWithinSupportedMagnitude
'   - PROB_TryAdd
'   - PROB_SetStatus
'
' CALLED FROM
'   - Worksheet formulas
'   - M_STATS_PROBDIST_TEST
'
' UPDATED
'   2026-07-21
'==============================================================================
'
'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Dim ShapeSum            As Double          'Alpha plus Beta
    Dim Denominator         As Double          'Alpha plus Beta plus one
    Dim Ratio               As Double          'Smaller shape divided by larger shape
    Dim MeanValue           As Double          'Beta mean
    Dim VarianceValue       As Double          'Beta variance
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
    'Validate the first shape parameter
        If Not PROB_IsPositiveWithinSupportedMagnitude(Alpha) Then
            FailMsg = _
                "Alpha must be a finite strictly positive number within the parameter-magnitude guard (< 1E100)"
            GoTo Fail_Num
        End If

    'Validate the second shape parameter
        If Not PROB_IsPositiveWithinSupportedMagnitude(Beta) Then
            FailMsg = _
                "Beta must be a finite strictly positive number within the parameter-magnitude guard (< 1E100)"
            GoTo Fail_Num
        End If

'------------------------------------------------------------------------------
' COMPUTE STANDARD DEVIATION
'------------------------------------------------------------------------------
    'Form the mean through a ratio no greater than one
        If Alpha >= Beta Then
            Ratio = Beta / Alpha
            MeanValue = 1# / (1# + Ratio)
        Else
            Ratio = Alpha / Beta
            MeanValue = Ratio / (1# + Ratio)
        End If

    'Form Alpha + Beta under the shared arithmetic contract
        If Not PROB_TryAdd(Alpha, Beta, ShapeSum) Then
            FailMsg = "Beta standard deviation overflowed in Alpha + Beta"
            GoTo Fail_Num
        End If

    'Form Alpha + Beta + 1 under the shared arithmetic contract
        If Not PROB_TryAdd(ShapeSum, 1#, Denominator) Then
            FailMsg = _
                "Beta standard deviation denominator overflowed a Double"
            GoTo Fail_Num
        End If

    'Evaluate the variance and take its square root
        VarianceValue = _
            MeanValue * (1# - MeanValue) / Denominator

        K_STATS_Beta_StdDev = Sqr(VarianceValue)

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
' WORKSHEET EQUIVALENT
'   EXPON.DIST(X, Lambda, FALSE)
'
' INPUTS
'   X       Evaluation point. Values below zero have density zero.
'   Lambda  Positive rate parameter.
'   Status  Optional ByRef diagnostic message.
'
' RETURNS
'   Variant
'     Success => Double density.
'     Failure => CVErr(xlErrNum) or CVErr(xlErrValue).
'
' BEHAVIOR
'   - Returns zero below the support.
'   - Returns Lambda at X = 0.
'   - If Lambda * X overflows, the exponential damping makes the density zero.
'   - Mathematically valid underflow returns zero.
'
' ERROR POLICY
'   - Invalid parameters or density overflow return #NUM!.
'   - Unexpected runtime errors return #VALUE!.
'
' DEPENDENCIES
'   - PROB_CN_ValidateXLambda
'   - PROB_TryMultiply
'   - PROB_TryExp
'   - PROB_SetStatus
'
' CALLED FROM
'   - Worksheet formulas
'   - M_STATS_PROBDIST_TEST
'
' UPDATED
'   2026-07-21
'==============================================================================
'
'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Dim LambdaX             As Double          'Lambda multiplied by X
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
    'Validate the evaluation point and rate
        If Not PROB_CN_ValidateXLambda(X, Lambda, FailMsg) Then
            GoTo Fail_Num
        End If

'------------------------------------------------------------------------------
' HANDLE SUPPORT EDGES
'------------------------------------------------------------------------------
    'Return zero below the support
        If X < 0# Then
            K_STATS_Exponential_Density = 0#
            GoTo Return_Success
        End If

    'Return the rate at the support origin
        If X = 0# Then
            K_STATS_Exponential_Density = Lambda
            GoTo Return_Success
        End If

'------------------------------------------------------------------------------
' COMPUTE DENSITY
'------------------------------------------------------------------------------
    'Treat positive product overflow as complete exponential damping
        If Not PROB_TryMultiply(Lambda, X, LambdaX) Then
            K_STATS_Exponential_Density = 0#
            GoTo Return_Success
        End If

    'Evaluate Exp(Log(Lambda) - Lambda * X) under the shared contract
        If Not PROB_TryExp(Log(Lambda) - LambdaX, Density) Then
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
'   Returns the Exponential left-tail cumulative probability at X.
'
' WORKSHEET EQUIVALENT
'   EXPON.DIST(X, Lambda, TRUE)
'
' INPUTS
'   X       Evaluation point.
'   Lambda  Positive rate parameter.
'   Status  Optional ByRef diagnostic message.
'
' RETURNS
'   Variant
'     Success => Double cumulative probability in [0, 1].
'     Failure => CVErr(xlErrNum) or CVErr(xlErrValue).
'
' BEHAVIOR
'   - Returns zero for X <= 0.
'   - Evaluates 1 - Exp(-Lambda * X) through -PROB_Expm1(-Lambda * X).
'   - If Lambda * X overflows, returns the limiting value one.
'
' ERROR POLICY
'   - Invalid parameters return #NUM!.
'   - Unexpected runtime errors return #VALUE!.
'
' DEPENDENCIES
'   - PROB_CN_ValidateXLambda
'   - PROB_TryMultiply
'   - PROB_Expm1
'   - PROB_SetStatus
'
' CALLED FROM
'   - Worksheet formulas
'   - M_STATS_PROBDIST_TEST
'
' UPDATED
'   2026-07-21
'==============================================================================
'
'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Dim LambdaX             As Double          'Lambda multiplied by X
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
    'Validate the evaluation point and rate
        If Not PROB_CN_ValidateXLambda(X, Lambda, FailMsg) Then
            GoTo Fail_Num
        End If

'------------------------------------------------------------------------------
' HANDLE SUPPORT EDGE
'------------------------------------------------------------------------------
    'Return zero at and below the support origin
        If X <= 0# Then
            K_STATS_Exponential_Cumulative = 0#
            GoTo Return_Success
        End If

'------------------------------------------------------------------------------
' COMPUTE CUMULATIVE PROBABILITY
'------------------------------------------------------------------------------
    'Treat positive product overflow as the +infinity limit
        If Not PROB_TryMultiply(Lambda, X, LambdaX) Then
            K_STATS_Exponential_Cumulative = 1#
            GoTo Return_Success
        End If

    'Evaluate the CDF without cancellation in the left tail
        K_STATS_Exponential_Cumulative = -PROB_Expm1(-LambdaX)

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
'   Returns the Exponential right-tail survival probability at X.
'
' WORKSHEET EQUIVALENT
'   1 - EXPON.DIST(X, Lambda, TRUE)
'
' INPUTS
'   X       Evaluation point.
'   Lambda  Positive rate parameter.
'   Status  Optional ByRef diagnostic message.
'
' RETURNS
'   Variant
'     Success => Double survival probability in [0, 1].
'     Failure => CVErr(xlErrNum) or CVErr(xlErrValue).
'
' BEHAVIOR
'   - Returns one for X <= 0.
'   - Evaluates Exp(-Lambda * X) directly.
'   - If Lambda * X overflows, returns the limiting value zero.
'
' ERROR POLICY
'   - Invalid parameters return #NUM!.
'   - Unexpected runtime errors return #VALUE!.
'
' DEPENDENCIES
'   - PROB_CN_ValidateXLambda
'   - PROB_TryMultiply
'   - PROB_TryExp
'   - PROB_SetStatus
'
' CALLED FROM
'   - Worksheet formulas
'   - M_STATS_PROBDIST_TEST
'
' UPDATED
'   2026-07-21
'==============================================================================
'
'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Dim LambdaX             As Double          'Lambda multiplied by X
    Dim Survival            As Double          'Returned survival probability
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
    'Validate the evaluation point and rate
        If Not PROB_CN_ValidateXLambda(X, Lambda, FailMsg) Then
            GoTo Fail_Num
        End If

'------------------------------------------------------------------------------
' HANDLE SUPPORT EDGE
'------------------------------------------------------------------------------
    'Return one at and below the support origin
        If X <= 0# Then
            K_STATS_Exponential_Survival = 1#
            GoTo Return_Success
        End If

'------------------------------------------------------------------------------
' COMPUTE SURVIVAL PROBABILITY
'------------------------------------------------------------------------------
    'Treat positive product overflow as the +infinity limit
        If Not PROB_TryMultiply(Lambda, X, LambdaX) Then
            K_STATS_Exponential_Survival = 0#
            GoTo Return_Success
        End If

    'Evaluate the direct survival exponential
        If Not PROB_TryExp(-LambdaX, Survival) Then
            FailMsg = _
                "Unexpected positive exponential argument in survival kernel"
            GoTo Fail_Num
        End If

    'Return the survival probability
        K_STATS_Exponential_Survival = Survival

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
'   Returns the Exponential quantile for Probability and rate Lambda.
'
' WHY THIS EXISTS
'   The quantile is evaluated in the logarithmic domain so very small rates and
'   extreme probabilities are classified before a final Double overflow occurs.
'
' WORKSHEET EQUIVALENT
'   EXPON.INV(Probability, Lambda)
'
' INPUTS
'   Probability  Target cumulative probability in the open unit interval.
'   Lambda       Positive rate parameter.
'   Status       Optional ByRef diagnostic message.
'
' RETURNS
'   Variant
'     Success => Double quantile.
'     Failure => CVErr(xlErrNum) or CVErr(xlErrValue).
'
' BEHAVIOR
'   - Computes -Log(1 - Probability) through PROB_Log1p.
'   - Forms Log(Quantile) before the final guarded exponential.
'   - A final quantile outside Double range returns #NUM!.
'
' ERROR POLICY
'   - Invalid inputs or quantile overflow return #NUM!.
'   - Unexpected runtime errors return #VALUE!.
'
' DEPENDENCIES
'   - PROB_IsValidProbabilityOpen
'   - PROB_IsPositiveFinite
'   - PROB_Log1p
'   - PROB_TryExp
'   - PROB_SetStatus
'
' CALLED FROM
'   - Worksheet formulas
'   - M_STATS_PROBDIST_TEST
'
' UPDATED
'   2026-07-21
'==============================================================================
'
'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Dim NegLogComplement    As Double          '-Log(1 - Probability)
    Dim LogQuantile         As Double          'Logarithm of the quantile
    Dim Quantile            As Double          'Returned quantile
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
    'Validate the target probability
        If Not PROB_IsValidProbabilityOpen(Probability) Then
            FailMsg = "Probability must be strictly between 0 and 1"
            GoTo Fail_Num
        End If

    'Validate the rate parameter
        If Not PROB_IsPositiveFinite(Lambda) Then
            FailMsg = "Lambda must be a finite strictly positive number"
            GoTo Fail_Num
        End If

'------------------------------------------------------------------------------
' COMPUTE QUANTILE
'------------------------------------------------------------------------------
    'Evaluate -Log(1 - Probability) without cancellation
        NegLogComplement = -PROB_Log1p(-Probability)

    'Assemble the quantile in the logarithmic domain
        LogQuantile = Log(NegLogComplement) - Log(Lambda)

    'Exponentiate under the shared overflow and underflow contract
        If Not PROB_TryExp(LogQuantile, Quantile) Then
            FailMsg = "Exponential quantile overflowed a Double"
            GoTo Fail_Num
        End If

    'Return the quantile
        K_STATS_Exponential_InverseCumulative = Quantile

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
'   Returns the Weibull probability density at X.
'
' WHY THIS EXISTS
'   Weibull densities are central to reliability and time-to-failure modelling.
'   The power term and the density are evaluated through guarded logarithmic
'   expressions so extreme parameters are classified predictably.
'
' WORKSHEET EQUIVALENT
'   WEIBULL.DIST(X, Shape, ScaleParam, FALSE)
'
' INPUTS
'   X           Evaluation point. Values below zero have density zero.
'   Shape       Positive Weibull shape parameter.
'   ScaleParam  Positive Weibull scale parameter.
'   Status      Optional ByRef diagnostic message.
'
' RETURNS
'   Variant
'     Success => Double density.
'     Failure => CVErr(xlErrNum) or CVErr(xlErrValue).
'
' BEHAVIOR
'   - Returns zero below the support.
'   - At X = 0, returns #NUM! for Shape < 1, 1 / ScaleParam for Shape = 1,
'     and zero for Shape > 1.
'   - A power term tending to positive infinity drives the density to zero.
'   - Mathematically valid underflow returns zero.
'
' ERROR POLICY
'   - Invalid parameters, a density pole or density overflow return #NUM!.
'   - Unexpected runtime errors return #VALUE!.
'
' DEPENDENCIES
'   - PROB_CN_ValidateXShapeScale
'   - PROB_CN_TryWeibullPower
'   - PROB_TryMultiply
'   - PROB_TryAdd
'   - PROB_TryDivide
'   - PROB_TryExp
'   - PROB_SetStatus
'
' CALLED FROM
'   - Worksheet formulas
'   - M_STATS_PROBDIST_TEST
'
' UPDATED
'   2026-07-21
'==============================================================================
'
'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Dim LogRatio            As Double          'Log(X / ScaleParam)
    Dim ShapeLogRatio       As Double          '(Shape - 1) * LogRatio
    Dim LogDensity          As Double          'Logarithm of the density
    Dim LogDensityPartial   As Double          'Intermediate log-density sum
    Dim PowerValue          As Double          '(X / ScaleParam) ^ Shape
    Dim PowerIsInfinite     As Boolean         'True when the power exceeds Double
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
    'Validate X, Shape and ScaleParam
        If Not PROB_CN_ValidateXShapeScale( _
            X, Shape, ScaleParam, FailMsg, "Shape", "ScaleParam") Then
            GoTo Fail_Num
        End If

'------------------------------------------------------------------------------
' HANDLE SUPPORT EDGES
'------------------------------------------------------------------------------
    'Return zero below the positive support
        If X < 0# Then
            K_STATS_Weibull_Density = 0#
            GoTo Return_Success
        End If

    'Handle the origin according to the Weibull shape
        If X = 0# Then
            If Shape < 1# Then
                FailMsg = _
                    "Weibull density is unbounded at X = 0 when Shape < 1"
                GoTo Fail_Num

            ElseIf Shape = 1# Then
                If Not PROB_TryDivide(1#, ScaleParam, Density) Then
                    FailMsg = "Weibull density overflows Double at X = 0"
                    GoTo Fail_Num
                End If

                K_STATS_Weibull_Density = Density

            Else
                K_STATS_Weibull_Density = 0#
            End If

            GoTo Return_Success
        End If

'------------------------------------------------------------------------------
' COMPUTE DENSITY
'------------------------------------------------------------------------------
    'Form the logarithm of the scaled evaluation point
        LogRatio = Log(X) - Log(ScaleParam)

    'Evaluate the Weibull power under the shared arithmetic contract
        If Not PROB_CN_TryWeibullPower( _
            LogRatio, Shape, PowerValue, PowerIsInfinite, FailMsg) Then
            GoTo Fail_Num
        End If

    'An infinite positive power drives the density to zero
        If PowerIsInfinite Then
            K_STATS_Weibull_Density = 0#
            GoTo Return_Success
        End If

    'Form the shape-adjusted logarithmic term with explicit overflow handling
        If Not PROB_TryMultiply(Shape - 1#, LogRatio, ShapeLogRatio) Then
            If (Shape - 1#) * Sgn(LogRatio) < 0# Then
                K_STATS_Weibull_Density = 0#
                GoTo Return_Success
            End If

            FailMsg = "Weibull log-density overflowed a Double"
            GoTo Fail_Num
        End If

    'Combine the variable and power terms without an unclassified overflow
        If Not PROB_TryAdd(ShapeLogRatio, -PowerValue, LogDensityPartial) Then
            If ShapeLogRatio < 0# Then
                K_STATS_Weibull_Density = 0#
                GoTo Return_Success
            End If

            FailMsg = "Weibull log-density overflowed a Double"
            GoTo Fail_Num
        End If

    'Add the shape and scale normalization terms
        If Not PROB_TryAdd( _
            Log(Shape) - Log(ScaleParam), _
            LogDensityPartial, _
            LogDensity) Then

            If LogDensityPartial < 0# Then
                K_STATS_Weibull_Density = 0#
                GoTo Return_Success
            End If

            FailMsg = "Weibull log-density overflowed a Double"
            GoTo Fail_Num
        End If

    'Exponentiate the assembled log-density
        If Not PROB_TryExp(LogDensity, Density) Then
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
'   Returns the Weibull left-tail cumulative probability at X.
'
' WORKSHEET EQUIVALENT
'   WEIBULL.DIST(X, Shape, ScaleParam, TRUE)
'
' INPUTS
'   X           Evaluation point.
'   Shape       Positive Weibull shape parameter.
'   ScaleParam  Positive Weibull scale parameter.
'   Status      Optional ByRef diagnostic message.
'
' RETURNS
'   Variant
'     Success => Double cumulative probability in [0, 1].
'     Failure => CVErr(xlErrNum) or CVErr(xlErrValue).
'
' BEHAVIOR
'   - Returns zero for X <= 0.
'   - Evaluates 1 - Exp(-(X / ScaleParam) ^ Shape) through PROB_Expm1.
'   - A power term tending to positive infinity returns the limiting value one.
'
' ERROR POLICY
'   - Invalid parameters return #NUM!.
'   - Unexpected runtime errors return #VALUE!.
'
' DEPENDENCIES
'   - PROB_CN_ValidateXShapeScale
'   - PROB_CN_TryWeibullPower
'   - PROB_Expm1
'   - PROB_SetStatus
'
' CALLED FROM
'   - Worksheet formulas
'   - M_STATS_PROBDIST_TEST
'
' UPDATED
'   2026-07-21
'==============================================================================
'
'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Dim LogRatio            As Double          'Log(X / ScaleParam)
    Dim PowerValue          As Double          '(X / ScaleParam) ^ Shape
    Dim PowerIsInfinite     As Boolean         'True when the power exceeds Double
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
    'Validate X, Shape and ScaleParam
        If Not PROB_CN_ValidateXShapeScale( _
            X, Shape, ScaleParam, FailMsg, "Shape", "ScaleParam") Then
            GoTo Fail_Num
        End If

'------------------------------------------------------------------------------
' HANDLE SUPPORT EDGE
'------------------------------------------------------------------------------
    'Return zero at and below the support origin
        If X <= 0# Then
            K_STATS_Weibull_Cumulative = 0#
            GoTo Return_Success
        End If

'------------------------------------------------------------------------------
' COMPUTE CUMULATIVE PROBABILITY
'------------------------------------------------------------------------------
    'Form the logarithm of the scaled evaluation point
        LogRatio = Log(X) - Log(ScaleParam)

    'Evaluate the Weibull power under the shared arithmetic contract
        If Not PROB_CN_TryWeibullPower( _
            LogRatio, Shape, PowerValue, PowerIsInfinite, FailMsg) Then
            GoTo Fail_Num
        End If

    'An infinite positive power saturates the CDF at one
        If PowerIsInfinite Then
            K_STATS_Weibull_Cumulative = 1#
            GoTo Return_Success
        End If

    'Evaluate one minus the survival without left-tail cancellation
        K_STATS_Weibull_Cumulative = -PROB_Expm1(-PowerValue)

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
'   Returns the Weibull right-tail survival probability at X.
'
' WORKSHEET EQUIVALENT
'   1 - WEIBULL.DIST(X, Shape, ScaleParam, TRUE)
'
' INPUTS
'   X           Evaluation point.
'   Shape       Positive Weibull shape parameter.
'   ScaleParam  Positive Weibull scale parameter.
'   Status      Optional ByRef diagnostic message.
'
' RETURNS
'   Variant
'     Success => Double survival probability in [0, 1].
'     Failure => CVErr(xlErrNum) or CVErr(xlErrValue).
'
' BEHAVIOR
'   - Returns one for X <= 0.
'   - Evaluates Exp(-(X / ScaleParam) ^ Shape) directly.
'   - A power term tending to positive infinity returns the limiting value zero.
'
' ERROR POLICY
'   - Invalid parameters return #NUM!.
'   - Unexpected runtime errors return #VALUE!.
'
' DEPENDENCIES
'   - PROB_CN_ValidateXShapeScale
'   - PROB_CN_TryWeibullPower
'   - PROB_TryExp
'   - PROB_SetStatus
'
' CALLED FROM
'   - Worksheet formulas
'   - M_STATS_PROBDIST_TEST
'
' UPDATED
'   2026-07-21
'==============================================================================
'
'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Dim LogRatio            As Double          'Log(X / ScaleParam)
    Dim PowerValue          As Double          '(X / ScaleParam) ^ Shape
    Dim PowerIsInfinite     As Boolean         'True when the power exceeds Double
    Dim Survival            As Double          'Returned survival probability
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
    'Validate X, Shape and ScaleParam
        If Not PROB_CN_ValidateXShapeScale( _
            X, Shape, ScaleParam, FailMsg, "Shape", "ScaleParam") Then
            GoTo Fail_Num
        End If

'------------------------------------------------------------------------------
' HANDLE SUPPORT EDGE
'------------------------------------------------------------------------------
    'Return one at and below the support origin
        If X <= 0# Then
            K_STATS_Weibull_Survival = 1#
            GoTo Return_Success
        End If

'------------------------------------------------------------------------------
' COMPUTE SURVIVAL PROBABILITY
'------------------------------------------------------------------------------
    'Form the logarithm of the scaled evaluation point
        LogRatio = Log(X) - Log(ScaleParam)

    'Evaluate the Weibull power under the shared arithmetic contract
        If Not PROB_CN_TryWeibullPower( _
            LogRatio, Shape, PowerValue, PowerIsInfinite, FailMsg) Then
            GoTo Fail_Num
        End If

    'An infinite positive power drives the survival probability to zero
        If PowerIsInfinite Then
            K_STATS_Weibull_Survival = 0#
            GoTo Return_Success
        End If

    'Evaluate the survival exponential under the shared contract
        If Not PROB_TryExp(-PowerValue, Survival) Then
            FailMsg = _
                "Unexpected positive exponential argument in Weibull survival"
            GoTo Fail_Num
        End If

    'Return the survival probability
        K_STATS_Weibull_Survival = Survival

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
'   Returns the Weibull quantile for Probability, Shape and ScaleParam.
'
' WHY THIS EXISTS
'   The quantile is assembled in the logarithmic domain so tiny shapes, extreme
'   probabilities and large scales are classified before a runtime overflow can
'   escape into the worksheet error contract.
'
' WORKSHEET EQUIVALENT
'   WEIBULL.INV(Probability, Shape, ScaleParam)
'
' INPUTS
'   Probability  Target cumulative probability in the open unit interval.
'   Shape        Positive Weibull shape parameter.
'   ScaleParam   Positive Weibull scale parameter.
'   Status       Optional ByRef diagnostic message.
'
' RETURNS
'   Variant
'     Success => Double quantile.
'     Failure => CVErr(xlErrNum) or CVErr(xlErrValue).
'
' BEHAVIOR
'   - Computes Log(Quantile) as:
'         Log(ScaleParam) + Log(-Log(1 - Probability)) / Shape.
'   - Guards the division, addition and final exponential separately.
'   - A final quantile outside Double range returns #NUM!.
'
' ERROR POLICY
'   - Invalid inputs or predictable overflow return #NUM!.
'   - Unexpected runtime errors return #VALUE!.
'
' DEPENDENCIES
'   - PROB_IsValidProbabilityOpen
'   - PROB_IsPositiveWithinSupportedMagnitude
'   - PROB_IsPositiveFinite
'   - PROB_Log1p
'   - PROB_TryDivide
'   - PROB_TryAdd
'   - PROB_TryExp
'   - PROB_SetStatus
'
' CALLED FROM
'   - Worksheet formulas
'   - M_STATS_PROBDIST_TEST
'
' UPDATED
'   2026-07-21
'==============================================================================
'
'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Dim NegLogComplement    As Double          '-Log(1 - Probability)
    Dim ShapeTerm           As Double          'Log(NegLogComplement) / Shape
    Dim LogQuantile         As Double          'Logarithm of the quantile
    Dim Quantile            As Double          'Returned quantile
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
    'Validate the target probability
        If Not PROB_IsValidProbabilityOpen(Probability) Then
            FailMsg = "Probability must be strictly between 0 and 1"
            GoTo Fail_Num
        End If

    'Validate the shape under the supported-kernel contract
        If Not PROB_IsPositiveWithinSupportedMagnitude(Shape) Then
            FailMsg = _
                "Shape must be a finite strictly positive number within the parameter-magnitude guard (< 1E100)"
            GoTo Fail_Num
        End If

    'Validate the scale over the full finite Double range
        If Not PROB_IsPositiveFinite(ScaleParam) Then
            FailMsg = "ScaleParam must be a finite strictly positive number"
            GoTo Fail_Num
        End If

'------------------------------------------------------------------------------
' COMPUTE QUANTILE
'------------------------------------------------------------------------------
    'Evaluate -Log(1 - Probability) without cancellation
        NegLogComplement = -PROB_Log1p(-Probability)

    'Divide the logarithmic probability term by Shape under the Try-contract
        If Not PROB_TryDivide( _
            Log(NegLogComplement), Shape, ShapeTerm) Then

            FailMsg = "Weibull quantile exponent overflowed a Double"
            GoTo Fail_Num
        End If

    'Add the logarithm of the scale under the Try-contract
        If Not PROB_TryAdd( _
            Log(ScaleParam), ShapeTerm, LogQuantile) Then

            FailMsg = "Weibull log-quantile overflowed a Double"
            GoTo Fail_Num
        End If

    'Exponentiate the assembled log-quantile
        If Not PROB_TryExp(LogQuantile, Quantile) Then
            FailMsg = "Weibull quantile overflowed a Double"
            GoTo Fail_Num
        End If

    'Return the quantile
        K_STATS_Weibull_InverseCumulative = Quantile

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
'   Returns the Weibull mean, ScaleParam * Gamma(1 + 1 / Shape).
'
' WHY THIS EXISTS
'   The Gamma factor is kept in logarithmic form until the final exponential.
'   This permits a large Gamma factor to combine with a small scale whenever the
'   final mean remains representable.
'
' INPUTS
'   Shape       Positive Weibull shape parameter.
'   ScaleParam  Positive Weibull scale parameter.
'   Status      Optional ByRef diagnostic message.
'
' RETURNS
'   Variant
'     Success => Double mean.
'     Failure => CVErr(xlErrNum) or CVErr(xlErrValue).
'
' BEHAVIOR
'   - Guards the reciprocal shape.
'   - Rejects reciprocal shapes beyond the supported log-gamma range used here.
'   - Performs the scale adjustment and final result in the logarithmic domain.
'   - Mathematically valid underflow returns zero.
'
' ERROR POLICY
'   - Invalid inputs or final-result overflow return #NUM!.
'   - Unexpected runtime errors return #VALUE!.
'
' DEPENDENCIES
'   - PROB_IsPositiveWithinSupportedMagnitude
'   - PROB_IsPositiveFinite
'   - PROB_TryDivide
'   - PROB_TryAdd
'   - PROB_TryExp
'   - PROB_LogGamma
'   - PROB_SetStatus
'
' CALLED FROM
'   - Worksheet formulas
'   - M_STATS_PROBDIST_TEST
'
' NOTES
'   MAX_SAFE_EPSILON is a conservative implementation boundary. Above it, the
'   Gamma factor already exceeds what the smallest positive finite scale can
'   offset into the representable Double range.
'
' UPDATED
'   2026-07-21
'==============================================================================
'
'------------------------------------------------------------------------------
' CONSTANTS
'------------------------------------------------------------------------------
    Const MAX_SAFE_EPSILON As Double = 1000#

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Dim Epsilon             As Double          'Reciprocal shape
    Dim LogMean             As Double          'Logarithm of the mean
    Dim MeanValue           As Double          'Returned mean
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
    'Validate the shape under the supported-kernel contract
        If Not PROB_IsPositiveWithinSupportedMagnitude(Shape) Then
            FailMsg = _
                "Shape must be a finite strictly positive number within the parameter-magnitude guard (< 1E100)"
            GoTo Fail_Num
        End If

    'Validate the scale over the full finite Double range
        If Not PROB_IsPositiveFinite(ScaleParam) Then
            FailMsg = "ScaleParam must be a finite strictly positive number"
            GoTo Fail_Num
        End If

'------------------------------------------------------------------------------
' COMPUTE MEAN
'------------------------------------------------------------------------------
    'Form the reciprocal shape under the Try-contract
        If Not PROB_TryDivide(1#, Shape, Epsilon) Then
            FailMsg = "Weibull reciprocal shape overflowed a Double"
            GoTo Fail_Num
        End If

    'Reject a reciprocal shape beyond the supported compensable range
        If Epsilon > MAX_SAFE_EPSILON Then
            FailMsg = _
                "Weibull mean exceeds Double range for the supplied Shape"
            GoTo Fail_Num
        End If

    'Add the logarithmic scale and Gamma factor under the Try-contract
        If Not PROB_TryAdd( _
            Log(ScaleParam), _
            PROB_LogGamma(1# + Epsilon), _
            LogMean) Then

            FailMsg = "Weibull log-mean overflowed a Double"
            GoTo Fail_Num
        End If

    'Exponentiate the assembled logarithmic mean
        If Not PROB_TryExp(LogMean, MeanValue) Then
            FailMsg = "Weibull mean overflows Double range"
            GoTo Fail_Num
        End If

    'Return the mean
        K_STATS_Weibull_Mean = MeanValue

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
'   Returns the Weibull variance.
'
' WHY THIS EXISTS
'   Direct subtraction of Gamma(1 + 2 / Shape) and Gamma(1 + 1 / Shape) squared
'   loses all precision for large Shape. The private helper returns the variance
'   factor in logarithmic form and switches to an asymptotic expansion where
'   necessary.
'
' INPUTS
'   Shape       Positive Weibull shape parameter.
'   ScaleParam  Positive Weibull scale parameter.
'   Status      Optional ByRef diagnostic message.
'
' RETURNS
'   Variant
'     Success => Double variance.
'     Failure => CVErr(xlErrNum) or CVErr(xlErrValue).
'
' BEHAVIOR
'   - Computes the shape factor without large-shape cancellation.
'   - Applies ScaleParam squared in the logarithmic domain.
'   - Mathematically valid underflow returns zero.
'
' ERROR POLICY
'   - Invalid inputs, unresolved shape factors or final overflow return #NUM!.
'   - Unexpected runtime errors return #VALUE!.
'
' DEPENDENCIES
'   - PROB_IsPositiveWithinSupportedMagnitude
'   - PROB_IsPositiveFinite
'   - PROB_CN_TryWeibullLogVarianceFactor
'   - PROB_TryAdd
'   - PROB_TryExp
'   - PROB_SetStatus
'
' CALLED FROM
'   - Worksheet formulas
'   - M_STATS_PROBDIST_TEST
'
' UPDATED
'   2026-07-21
'==============================================================================
'
'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Dim LogScaleSquared     As Double          'Two times Log(ScaleParam)
    Dim LogShapeFactor      As Double          'Logarithm of the variance factor
    Dim LogVariance         As Double          'Logarithm of the final variance
    Dim VarianceValue       As Double          'Returned variance
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
    'Validate the shape under the supported-kernel contract
        If Not PROB_IsPositiveWithinSupportedMagnitude(Shape) Then
            FailMsg = _
                "Shape must be a finite strictly positive number within the parameter-magnitude guard (< 1E100)"
            GoTo Fail_Num
        End If

    'Validate the scale over the full finite Double range
        If Not PROB_IsPositiveFinite(ScaleParam) Then
            FailMsg = "ScaleParam must be a finite strictly positive number"
            GoTo Fail_Num
        End If

'------------------------------------------------------------------------------
' COMPUTE VARIANCE
'------------------------------------------------------------------------------
    'Resolve the logarithm of the shape-dependent variance factor
        If Not PROB_CN_TryWeibullLogVarianceFactor( _
            Shape, LogShapeFactor, FailMsg) Then
            GoTo Fail_Num
        End If

    'Form twice the logarithm of the scale
        LogScaleSquared = 2# * Log(ScaleParam)

    'Assemble the final log-variance under the Try-contract
        If Not PROB_TryAdd( _
            LogScaleSquared, LogShapeFactor, LogVariance) Then

            FailMsg = "Weibull log-variance overflowed a Double"
            GoTo Fail_Num
        End If

    'Exponentiate the final log-variance
        If Not PROB_TryExp(LogVariance, VarianceValue) Then
            FailMsg = "Weibull variance overflows Double range"
            GoTo Fail_Num
        End If

    'Return the variance
        K_STATS_Weibull_Variance = VarianceValue

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
'   Returns the Weibull standard deviation.
'
' WHY THIS EXISTS
'   The standard deviation is assembled directly from one half of the stable
'   logarithmic variance factor. It therefore remains representable in cases
'   where an intermediate variance would overflow.
'
' INPUTS
'   Shape       Positive Weibull shape parameter.
'   ScaleParam  Positive Weibull scale parameter.
'   Status      Optional ByRef diagnostic message.
'
' RETURNS
'   Variant
'     Success => Double standard deviation.
'     Failure => CVErr(xlErrNum) or CVErr(xlErrValue).
'
' BEHAVIOR
'   - Computes the shape factor without large-shape cancellation.
'   - Applies the scale in the logarithmic domain.
'   - Mathematically valid underflow returns zero.
'
' ERROR POLICY
'   - Invalid inputs, unresolved shape factors or final overflow return #NUM!.
'   - Unexpected runtime errors return #VALUE!.
'
' DEPENDENCIES
'   - PROB_IsPositiveWithinSupportedMagnitude
'   - PROB_IsPositiveFinite
'   - PROB_CN_TryWeibullLogVarianceFactor
'   - PROB_TryAdd
'   - PROB_TryExp
'   - PROB_SetStatus
'
' CALLED FROM
'   - Worksheet formulas
'   - M_STATS_PROBDIST_TEST
'
' UPDATED
'   2026-07-21
'==============================================================================
'
'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Dim HalfLogShapeFactor  As Double          'Half of the log variance factor
    Dim LogShapeFactor      As Double          'Logarithm of the variance factor
    Dim LogStdDev           As Double          'Logarithm of the standard deviation
    Dim StdDevValue         As Double          'Returned standard deviation
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
    'Validate the shape under the supported-kernel contract
        If Not PROB_IsPositiveWithinSupportedMagnitude(Shape) Then
            FailMsg = _
                "Shape must be a finite strictly positive number within the parameter-magnitude guard (< 1E100)"
            GoTo Fail_Num
        End If

    'Validate the scale over the full finite Double range
        If Not PROB_IsPositiveFinite(ScaleParam) Then
            FailMsg = "ScaleParam must be a finite strictly positive number"
            GoTo Fail_Num
        End If

'------------------------------------------------------------------------------
' COMPUTE STANDARD DEVIATION
'------------------------------------------------------------------------------
    'Resolve the logarithm of the shape-dependent variance factor
        If Not PROB_CN_TryWeibullLogVarianceFactor( _
            Shape, LogShapeFactor, FailMsg) Then
            GoTo Fail_Num
        End If

    'Take one half of the logarithmic variance factor
        HalfLogShapeFactor = 0.5 * LogShapeFactor

    'Apply the logarithmic scale under the Try-contract
        If Not PROB_TryAdd( _
            Log(ScaleParam), HalfLogShapeFactor, LogStdDev) Then

            FailMsg = "Weibull log-standard-deviation overflowed a Double"
            GoTo Fail_Num
        End If

    'Exponentiate the assembled log-standard-deviation
        If Not PROB_TryExp(LogStdDev, StdDevValue) Then
            FailMsg = "Weibull standard deviation overflows Double range"
            GoTo Fail_Num
        End If

    'Return the standard deviation
        K_STATS_Weibull_StdDev = StdDevValue

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
'   Returns the continuous Uniform density on [LowerBound, UpperBound].
'
' INPUTS
'   X            Evaluation point.
'   LowerBound   Finite lower support bound.
'   UpperBound   Finite upper support bound, strictly greater than LowerBound.
'   Status       Optional ByRef diagnostic message.
'
' RETURNS
'   Variant
'     Success => Double density.
'     Failure => CVErr(xlErrNum) or CVErr(xlErrValue).
'
' BEHAVIOR
'   - Returns zero outside the closed support.
'   - Uses direct width arithmetic when representable.
'   - Uses a scaled reciprocal when opposite-sign bounds make the mathematical
'     width exceed the largest finite Double.
'
' ERROR POLICY
'   - Invalid bounds or a density beyond Double range return #NUM!.
'   - Mathematically valid underflow returns zero.
'   - Unexpected runtime errors return #VALUE!.
'
' DEPENDENCIES
'   - PROB_CN_ValidateXBounds
'   - PROB_TryAdd
'   - PROB_TryDivide
'   - PROB_SetStatus
'
' CALLED FROM
'   - Worksheet formulas
'   - M_STATS_PROBDIST_TEST
'
' UPDATED
'   2026-07-21
'==============================================================================
'
'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Dim ScaleValue          As Double          'Largest absolute support bound
    Dim Width               As Double          'Direct support width
    Dim WidthScaled         As Double          'Support width after scaling
    Dim ReciprocalScale     As Double          'One divided by ScaleValue
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
    'Validate X and the ordered finite bounds
        If Not PROB_CN_ValidateXBounds( _
            X, LowerBound, UpperBound, FailMsg) Then
            GoTo Fail_Num
        End If

'------------------------------------------------------------------------------
' HANDLE SUPPORT
'------------------------------------------------------------------------------
    'Return zero outside the closed support
        If X < LowerBound Or X > UpperBound Then
            K_STATS_Uniform_Density = 0#
            GoTo Return_Success
        End If

'------------------------------------------------------------------------------
' COMPUTE DENSITY
'------------------------------------------------------------------------------
    'Attempt to form the support width directly
        If PROB_TryAdd(UpperBound, -LowerBound, Width) Then
            If Width <= 0# Then
                FailMsg = "Uniform support width must be strictly positive"
                GoTo Fail_Num
            End If

            If Not PROB_TryDivide(1#, Width, Density) Then
                FailMsg = "Uniform density overflows Double range"
                GoTo Fail_Num
            End If

        Else
            'Scale opposite-sign extreme bounds before forming their width
                ScaleValue = Abs(LowerBound)

                If Abs(UpperBound) > ScaleValue Then
                    ScaleValue = Abs(UpperBound)
                End If

                WidthScaled = _
                    UpperBound / ScaleValue - _
                    LowerBound / ScaleValue

            'Form one divided by the mathematical width without forming width
                If Not PROB_TryDivide(1#, ScaleValue, ReciprocalScale) Then
                    FailMsg = "Uniform density overflows Double range"
                    GoTo Fail_Num
                End If

                Density = ReciprocalScale / WidthScaled
        End If

    'Return the density
        K_STATS_Uniform_Density = Density

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
'   Returns the continuous Uniform cumulative probability at X.
'
' INPUTS
'   X            Evaluation point.
'   LowerBound   Finite lower support bound.
'   UpperBound   Finite upper support bound, strictly greater than LowerBound.
'   Status       Optional ByRef diagnostic message.
'
' RETURNS
'   Variant
'     Success => Double cumulative probability in [0, 1].
'     Failure => CVErr(xlErrNum) or CVErr(xlErrValue).
'
' BEHAVIOR
'   - Returns zero below the support and one above it.
'   - Uses direct differences when the support width is representable.
'   - Uses scaled coordinates when the mathematical support width exceeds the
'     largest finite Double.
'
' ERROR POLICY
'   - Invalid bounds return #NUM!.
'   - Unexpected runtime errors return #VALUE!.
'
' DEPENDENCIES
'   - PROB_CN_ValidateXBounds
'   - PROB_TryAdd
'   - PROB_SetStatus
'
' CALLED FROM
'   - Worksheet formulas
'   - M_STATS_PROBDIST_TEST
'
' UPDATED
'   2026-07-21
'==============================================================================
'
'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Dim ScaleValue          As Double          'Largest absolute support bound
    Dim Width               As Double          'Direct support width
    Dim Numerator           As Double          'Direct distance from LowerBound
    Dim WidthScaled         As Double          'Scaled support width
    Dim XScaled             As Double          'Scaled evaluation point
    Dim LowerScaled         As Double          'Scaled lower bound
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
    'Validate X and the ordered finite bounds
        If Not PROB_CN_ValidateXBounds( _
            X, LowerBound, UpperBound, FailMsg) Then
            GoTo Fail_Num
        End If

'------------------------------------------------------------------------------
' HANDLE SUPPORT EDGES
'------------------------------------------------------------------------------
    'Return zero at and below the lower support edge
        If X <= LowerBound Then
            K_STATS_Uniform_Cumulative = 0#
            GoTo Return_Success
        End If

    'Return one at and above the upper support edge
        If X >= UpperBound Then
            K_STATS_Uniform_Cumulative = 1#
            GoTo Return_Success
        End If

'------------------------------------------------------------------------------
' COMPUTE CUMULATIVE PROBABILITY
'------------------------------------------------------------------------------
    'Use direct differences when both are representable
        If PROB_TryAdd(UpperBound, -LowerBound, Width) And _
           PROB_TryAdd(X, -LowerBound, Numerator) Then

            K_STATS_Uniform_Cumulative = Numerator / Width

        Else
            'Scale all coordinates by the largest support magnitude
                ScaleValue = Abs(LowerBound)

                If Abs(UpperBound) > ScaleValue Then
                    ScaleValue = Abs(UpperBound)
                End If

                LowerScaled = LowerBound / ScaleValue
                XScaled = X / ScaleValue
                WidthScaled = _
                    UpperBound / ScaleValue - LowerScaled

            'Return the scaled linear position
                K_STATS_Uniform_Cumulative = _
                    (XScaled - LowerScaled) / WidthScaled
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
'   Returns the continuous Uniform survival probability at X.
'
' INPUTS
'   X            Evaluation point.
'   LowerBound   Finite lower support bound.
'   UpperBound   Finite upper support bound, strictly greater than LowerBound.
'   Status       Optional ByRef diagnostic message.
'
' RETURNS
'   Variant
'     Success => Double survival probability in [0, 1].
'     Failure => CVErr(xlErrNum) or CVErr(xlErrValue).
'
' BEHAVIOR
'   - Returns one below the support and zero above it.
'   - Computes the right tail directly rather than subtracting the CDF from one.
'   - Uses scaled coordinates when the mathematical support width exceeds the
'     largest finite Double.
'
' ERROR POLICY
'   - Invalid bounds return #NUM!.
'   - Unexpected runtime errors return #VALUE!.
'
' DEPENDENCIES
'   - PROB_CN_ValidateXBounds
'   - PROB_TryAdd
'   - PROB_SetStatus
'
' CALLED FROM
'   - Worksheet formulas
'   - M_STATS_PROBDIST_TEST
'
' UPDATED
'   2026-07-21
'==============================================================================
'
'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Dim ScaleValue          As Double          'Largest absolute support bound
    Dim Width               As Double          'Direct support width
    Dim Numerator           As Double          'Direct distance to UpperBound
    Dim WidthScaled         As Double          'Scaled support width
    Dim XScaled             As Double          'Scaled evaluation point
    Dim UpperScaled         As Double          'Scaled upper bound
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
    'Validate X and the ordered finite bounds
        If Not PROB_CN_ValidateXBounds( _
            X, LowerBound, UpperBound, FailMsg) Then
            GoTo Fail_Num
        End If

'------------------------------------------------------------------------------
' HANDLE SUPPORT EDGES
'------------------------------------------------------------------------------
    'Return one at and below the lower support edge
        If X <= LowerBound Then
            K_STATS_Uniform_Survival = 1#
            GoTo Return_Success
        End If

    'Return zero at and above the upper support edge
        If X >= UpperBound Then
            K_STATS_Uniform_Survival = 0#
            GoTo Return_Success
        End If

'------------------------------------------------------------------------------
' COMPUTE SURVIVAL PROBABILITY
'------------------------------------------------------------------------------
    'Use direct differences when both are representable
        If PROB_TryAdd(UpperBound, -LowerBound, Width) And _
           PROB_TryAdd(UpperBound, -X, Numerator) Then

            K_STATS_Uniform_Survival = Numerator / Width

        Else
            'Scale all coordinates by the largest support magnitude
                ScaleValue = Abs(LowerBound)

                If Abs(UpperBound) > ScaleValue Then
                    ScaleValue = Abs(UpperBound)
                End If

                UpperScaled = UpperBound / ScaleValue
                XScaled = X / ScaleValue
                WidthScaled = _
                    UpperScaled - LowerBound / ScaleValue

            'Return the scaled direct right tail
                K_STATS_Uniform_Survival = _
                    (UpperScaled - XScaled) / WidthScaled
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
'   Returns the continuous Uniform quantile.
'
' NUMERICAL METHOD
'   Evaluates the quantile as the convex combination:
'       (1 - Probability) * LowerBound + Probability * UpperBound.
'   This avoids forming UpperBound - LowerBound when opposite-sign finite bounds
'   have a mathematical width exceeding the largest finite Double.
'
' INPUTS
'   Probability  Target cumulative probability in the open unit interval.
'   LowerBound   Finite lower support bound.
'   UpperBound   Finite upper support bound, strictly greater than LowerBound.
'   Status       Optional ByRef diagnostic message.
'
' RETURNS
'   Variant
'     Success => Double quantile strictly inside the support, subject to Double
'                rounding at extremely narrow relative intervals.
'     Failure => CVErr(xlErrNum) or CVErr(xlErrValue).
'
' ERROR POLICY
'   - Invalid probability or bounds return #NUM!.
'   - Predictable arithmetic overflow returns #NUM!.
'   - Unexpected runtime errors return #VALUE!.
'
' DEPENDENCIES
'   - PROB_IsValidProbabilityOpen
'   - PROB_CN_ValidateBounds
'   - PROB_TryMultiply
'   - PROB_TryAdd
'   - PROB_SetStatus
'
' CALLED FROM
'   - Worksheet formulas
'   - M_STATS_PROBDIST_TEST
'
' UPDATED
'   2026-07-21
'==============================================================================
'
'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Dim LowerWeight         As Double          'One minus Probability
    Dim LowerTerm           As Double          'Weighted lower bound
    Dim UpperTerm           As Double          'Weighted upper bound
    Dim Quantile            As Double          'Returned quantile
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
    'Validate the target probability
        If Not PROB_IsValidProbabilityOpen(Probability) Then
            FailMsg = "Probability must be strictly between 0 and 1"
            GoTo Fail_Num
        End If

    'Validate the ordered finite bounds
        If Not PROB_CN_ValidateBounds( _
            LowerBound, UpperBound, FailMsg) Then
            GoTo Fail_Num
        End If

'------------------------------------------------------------------------------
' COMPUTE QUANTILE
'------------------------------------------------------------------------------
    'Form the complementary probability weight
        LowerWeight = 1# - Probability

    'Weight the lower support bound
        If Not PROB_TryMultiply( _
            LowerWeight, LowerBound, LowerTerm) Then

            FailMsg = "Uniform lower quantile term overflowed a Double"
            GoTo Fail_Num
        End If

    'Weight the upper support bound
        If Not PROB_TryMultiply( _
            Probability, UpperBound, UpperTerm) Then

            FailMsg = "Uniform upper quantile term overflowed a Double"
            GoTo Fail_Num
        End If

    'Add the two convex-combination terms
        If Not PROB_TryAdd(LowerTerm, UpperTerm, Quantile) Then
            FailMsg = "Uniform quantile overflowed a Double"
            GoTo Fail_Num
        End If

    'Return the quantile
        K_STATS_Uniform_InverseCumulative = Quantile

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
' PRIVATE NUMERICAL HELPERS
'==============================================================================

Private Function PROB_CN_TryWeibullPower( _
    ByVal LogRatio As Double, _
    ByVal Shape As Double, _
    ByRef PowerValue As Double, _
    ByRef PowerIsInfinite As Boolean, _
    ByRef FailMsg As String) _
    As Boolean
'
'==============================================================================
' PROB_CN_TryWeibullPower
'------------------------------------------------------------------------------
' PURPOSE
'   Evaluates Exp(Shape * LogRatio), the Weibull power term, without allowing an
'   intermediate multiplication or exponential overflow to escape the numerical
'   Try-contract.
'
' INPUTS
'   LogRatio  Log(X / ScaleParam), already finite.
'   Shape     Supported positive Weibull shape parameter.
'
' OUTPUTS
'   PowerValue       Finite power value, including valid underflow to zero.
'   PowerIsInfinite  True when the mathematical power exceeds Double range.
'   FailMsg          Detailed failure message when the helper returns False.
'
' RETURNS
'   Boolean
'     True  => The power was resolved either as a finite Double or as a known
'              positive-infinity limit identified by PowerIsInfinite.
'     False => An unexpected sign or arithmetic state prevented classification.
'
' BEHAVIOR
'   - Positive log-power overflow is classified as a positive-infinity limit.
'   - Negative log-power overflow is classified as valid underflow to zero.
'   - A finite log-power is exponentiated through PROB_TryExp.
'
' DEPENDENCIES
'   - PROB_TryMultiply
'   - PROB_TryExp
'
' CALLED FROM
'   - K_STATS_Weibull_Density
'   - K_STATS_Weibull_Cumulative
'   - K_STATS_Weibull_Survival
'
' UPDATED
'   2026-07-21
'==============================================================================
'
'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Dim LogPower            As Double          'Shape multiplied by LogRatio

'------------------------------------------------------------------------------
' INITIALIZE OUTPUTS
'------------------------------------------------------------------------------
    'Clear all outputs before attempting the calculation
        PowerValue = 0#
        PowerIsInfinite = False
        FailMsg = vbNullString

'------------------------------------------------------------------------------
' FORM LOG-POWER
'------------------------------------------------------------------------------
    'Multiply Shape and LogRatio under the shared arithmetic contract
        If Not PROB_TryMultiply(Shape, LogRatio, LogPower) Then
            If LogRatio > 0# Then
                PowerIsInfinite = True
                PROB_CN_TryWeibullPower = True
                Exit Function
            End If

            If LogRatio < 0# Then
                PowerValue = 0#
                PROB_CN_TryWeibullPower = True
                Exit Function
            End If

            FailMsg = "Weibull log-power could not be classified"
            Exit Function
        End If

'------------------------------------------------------------------------------
' EXPONENTIATE
'------------------------------------------------------------------------------
    'Exponentiate the finite log-power
        If Not PROB_TryExp(LogPower, PowerValue) Then
            If LogPower > 0# Then
                PowerIsInfinite = True
                PROB_CN_TryWeibullPower = True
                Exit Function
            End If

            FailMsg = "Weibull power could not be evaluated"
            Exit Function
        End If

'------------------------------------------------------------------------------
' RETURN SUCCESS
'------------------------------------------------------------------------------
    'Report the finite power result
        PROB_CN_TryWeibullPower = True
End Function


Private Function PROB_CN_TryWeibullLogVarianceFactor( _
    ByVal Shape As Double, _
    ByRef LogFactor As Double, _
    ByRef FailMsg As String) _
    As Boolean
'
'==============================================================================
' PROB_CN_TryWeibullLogVarianceFactor
'------------------------------------------------------------------------------
' PURPOSE
'   Returns the logarithm of the Weibull variance factor:
'
'       Gamma(1 + 2 / Shape) - Gamma(1 + 1 / Shape) ^ 2.
'
' WHY THIS EXISTS
'   For large Shape, both Gamma terms round close to one and their direct
'   subtraction loses the complete variance factor. For moderate Shape, the
'   factor is formed through LogGamma and Expm1. For large Shape, a dedicated
'   asymptotic polynomial resolves the small positive difference directly.
'
' INPUTS
'   Shape  Supported positive Weibull shape parameter.
'
' OUTPUTS
'   LogFactor  Logarithm of the positive variance factor.
'   FailMsg    Detailed failure message when the helper returns False.
'
' RETURNS
'   Boolean
'     True  => LogFactor contains a valid result.
'     False => The reciprocal shape, variance factor or supported numerical
'              range could not be resolved.
'
' NUMERICAL METHOD
'   - Shape >= LARGE_SHAPE:
'       Epsilon = 1 / Shape.
'       Factor = Epsilon ^ 2 times a sixth-order polynomial.
'   - Shape < LARGE_SHAPE:
'       Delta = LogGamma(1 + 2Epsilon) - 2 LogGamma(1 + Epsilon).
'       LogFactor = 2 LogGamma(1 + Epsilon) + Log(Exp(Delta) - 1).
'       The final term uses PROB_Expm1 for small Delta and a complemented
'       logarithmic expression for larger Delta.
'
' ERROR POLICY
'   - Predictable numerical failure is reported through False and FailMsg.
'   - The helper does not raise worksheet errors directly.
'
' DEPENDENCIES
'   - PROB_TryDivide
'   - PROB_TryExp
'   - PROB_LogGamma
'   - PROB_Expm1
'   - PROB_Log1p
'
' CALLED FROM
'   - K_STATS_Weibull_Variance
'   - K_STATS_Weibull_StdDev
'
' UPDATED
'   2026-07-21
'==============================================================================
'
'------------------------------------------------------------------------------
' CONSTANTS
'------------------------------------------------------------------------------
    Const LARGE_SHAPE As Double = 100#
    Const MAX_SAFE_EPSILON As Double = 1000#

    Const C2 As Double = 1.64493406684823
    Const C3 As Double = -4.30307722854915
    Const C4 As Double = 11.7183391772189
    Const C5 As Double = -26.5314191646401
    Const C6 As Double = 57.6761128596097
    Const C7 As Double = -120.625407747693
    Const C8 As Double = 247.658400419811

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Dim Epsilon             As Double          'Reciprocal shape
    Dim Polynomial          As Double          'Large-shape asymptotic polynomial
    Dim Factor              As Double          'Positive variance factor
    Dim LogGamma1           As Double          'LogGamma(1 + Epsilon)
    Dim LogGamma2           As Double          'LogGamma(1 + 2 * Epsilon)
    Dim Delta               As Double          'Difference of logarithmic Gamma terms
    Dim ExpMinusDelta       As Double          'Exp(-Delta)
    Dim Expm1Delta          As Double          'Exp(Delta) - 1
    Dim LogExpm1Delta       As Double          'Log(Exp(Delta) - 1)

'------------------------------------------------------------------------------
' INITIALIZE OUTPUTS
'------------------------------------------------------------------------------
    'Clear output values before attempting the calculation
        LogFactor = 0#
        FailMsg = vbNullString

'------------------------------------------------------------------------------
' FORM RECIPROCAL SHAPE
'------------------------------------------------------------------------------
    'Form Epsilon under the shared arithmetic contract
        If Not PROB_TryDivide(1#, Shape, Epsilon) Then
            FailMsg = "Weibull reciprocal shape overflowed a Double"
            Exit Function
        End If

    'Reject a reciprocal shape beyond the supported compensable range
        If Epsilon > MAX_SAFE_EPSILON Then
            FailMsg = _
                "Weibull variance exceeds Double range for the supplied Shape"
            Exit Function
        End If

'------------------------------------------------------------------------------
' LARGE-SHAPE ASYMPTOTIC BRANCH
'------------------------------------------------------------------------------
    'Use the cancellation-free asymptotic expansion for large Shape
        If Shape >= LARGE_SHAPE Then
            Polynomial = C8
            Polynomial = C7 + Epsilon * Polynomial
            Polynomial = C6 + Epsilon * Polynomial
            Polynomial = C5 + Epsilon * Polynomial
            Polynomial = C4 + Epsilon * Polynomial
            Polynomial = C3 + Epsilon * Polynomial
            Polynomial = C2 + Epsilon * Polynomial

            Factor = Epsilon * Epsilon * Polynomial

            If Factor <= 0# Then
                FailMsg = "Weibull variance factor is not positive"
                Exit Function
            End If

            LogFactor = Log(Factor)
            PROB_CN_TryWeibullLogVarianceFactor = True
            Exit Function
        End If

'------------------------------------------------------------------------------
' MODERATE-SHAPE LOG-GAMMA BRANCH
'------------------------------------------------------------------------------
    'Evaluate the two logarithmic Gamma terms
        LogGamma1 = PROB_LogGamma(1# + Epsilon)
        LogGamma2 = PROB_LogGamma(1# + 2# * Epsilon)

    'Form the logarithmic Gamma difference
        Delta = LogGamma2 - 2# * LogGamma1

    'The exact variance factor is strictly positive
        If Delta <= 0# Then
            FailMsg = "Weibull variance factor lost positivity"
            Exit Function
        End If

    'Use Expm1 while Delta is small
        If Delta < 0.5 Then
            Expm1Delta = PROB_Expm1(Delta)

            If Expm1Delta <= 0# Then
                FailMsg = "Weibull variance factor could not be resolved"
                Exit Function
            End If

            LogExpm1Delta = Log(Expm1Delta)

        Else
            'Use Delta + Log(1 - Exp(-Delta)) for larger Delta
                If Not PROB_TryExp(-Delta, ExpMinusDelta) Then
                    ExpMinusDelta = 0#
                End If

                LogExpm1Delta = _
                    Delta + PROB_Log1p(-ExpMinusDelta)
        End If

    'Assemble the logarithm of the final variance factor
        LogFactor = 2# * LogGamma1 + LogExpm1Delta

'------------------------------------------------------------------------------
' RETURN SUCCESS
'------------------------------------------------------------------------------
    'Report the resolved logarithmic variance factor
        PROB_CN_TryWeibullLogVarianceFactor = True
End Function


'==============================================================================
' PRIVATE VALIDATORS
'==============================================================================

Private Function PROB_CN_ValidateXShapeScale( _
    ByVal X As Double, _
    ByVal Shape As Double, _
    ByVal ScaleParam As Double, _
    ByRef FailMsg As String, _
    ByVal ShapeName As String, _
    ByVal ScaleName As String) _
    As Boolean
'
'==============================================================================
' PROB_CN_ValidateXShapeScale
'------------------------------------------------------------------------------
' PURPOSE
'   Validates an evaluation point, a shape parameter constrained by the shared
'   supported-magnitude policy, and a scale parameter allowed over the full
'   finite positive Double range.
'
' INPUTS
'   X           Evaluation point.
'   Shape       Algorithmic shape parameter.
'   ScaleParam  Positive scale parameter.
'   ShapeName   Name used in diagnostic messages.
'   ScaleName   Name used in diagnostic messages.
'
' OUTPUTS
'   FailMsg  Empty on success; detailed validation message on failure.
'
' RETURNS
'   Boolean
'     True  => All inputs satisfy their distinct contracts.
'     False => At least one input is invalid.
'
' CALLED FROM
'   - Gamma density, CDF and survival
'   - Weibull density, CDF and survival
'
' UPDATED
'   2026-07-21
'==============================================================================
'
'------------------------------------------------------------------------------
' VALIDATE INPUTS
'------------------------------------------------------------------------------
    'Validate the evaluation point over the full finite Double range
        If Not PROB_IsFinite(X) Then
            FailMsg = "X must be a finite number"
            Exit Function
        End If

    'Validate the shape under the supported-kernel contract
        If Not PROB_IsPositiveWithinSupportedMagnitude(Shape) Then
            FailMsg = _
                ShapeName & _
                " must be a finite strictly positive number within the parameter-magnitude guard (< 1E100)"
            Exit Function
        End If

    'Validate the scale over the full finite positive Double range
        If Not PROB_IsPositiveFinite(ScaleParam) Then
            FailMsg = _
                ScaleName & " must be a finite strictly positive number"
            Exit Function
        End If

'------------------------------------------------------------------------------
' RETURN SUCCESS
'------------------------------------------------------------------------------
    'Report successful validation
        PROB_CN_ValidateXShapeScale = True
End Function


Private Function PROB_CN_ValidateXTwoShapes( _
    ByVal X As Double, _
    ByVal Shape1 As Double, _
    ByVal Shape2 As Double, _
    ByRef FailMsg As String, _
    ByVal Shape1Name As String, _
    ByVal Shape2Name As String) _
    As Boolean
'
'==============================================================================
' PROB_CN_ValidateXTwoShapes
'------------------------------------------------------------------------------
' PURPOSE
'   Validates an evaluation point and two positive shape parameters constrained
'   by the shared supported-magnitude policy.
'
' WHY THIS EXISTS
'   Beta functions require both Alpha and Beta to satisfy the numerical-kernel
'   shape contract. Treating the second parameter as an unrestricted scale would
'   allow unsupported values into the incomplete-beta and log-beta kernels.
'
' INPUTS
'   X           Evaluation point.
'   Shape1      First algorithmic shape parameter.
'   Shape2      Second algorithmic shape parameter.
'   Shape1Name  Name used in diagnostic messages.
'   Shape2Name  Name used in diagnostic messages.
'
' OUTPUTS
'   FailMsg  Empty on success; detailed validation message on failure.
'
' RETURNS
'   Boolean
'     True  => X and both shapes are valid.
'     False => At least one input is invalid.
'
' CALLED FROM
'   - K_STATS_Beta_Density
'   - K_STATS_Beta_Cumulative
'   - K_STATS_Beta_Survival
'
' UPDATED
'   2026-07-21
'==============================================================================
'
'------------------------------------------------------------------------------
' VALIDATE INPUTS
'------------------------------------------------------------------------------
    'Validate the evaluation point over the full finite Double range
        If Not PROB_IsFinite(X) Then
            FailMsg = "X must be a finite number"
            Exit Function
        End If

    'Validate the first shape parameter
        If Not PROB_IsPositiveWithinSupportedMagnitude(Shape1) Then
            FailMsg = _
                Shape1Name & _
                " must be a finite strictly positive number within the parameter-magnitude guard (< 1E100)"
            Exit Function
        End If

    'Validate the second shape parameter
        If Not PROB_IsPositiveWithinSupportedMagnitude(Shape2) Then
            FailMsg = _
                Shape2Name & _
                " must be a finite strictly positive number within the parameter-magnitude guard (< 1E100)"
            Exit Function
        End If

'------------------------------------------------------------------------------
' RETURN SUCCESS
'------------------------------------------------------------------------------
    'Report successful validation
        PROB_CN_ValidateXTwoShapes = True
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
'   Validates a finite evaluation point and a positive finite Exponential rate.
'
' INPUTS
'   X       Evaluation point.
'   Lambda  Exponential rate parameter.
'
' OUTPUTS
'   FailMsg  Empty on success; detailed validation message on failure.
'
' RETURNS
'   Boolean
'     True  => X and Lambda are valid.
'     False => At least one input is invalid.
'
' CALLED FROM
'   - K_STATS_Exponential_Density
'   - K_STATS_Exponential_Cumulative
'   - K_STATS_Exponential_Survival
'
' UPDATED
'   2026-07-21
'==============================================================================
'
'------------------------------------------------------------------------------
' VALIDATE INPUTS
'------------------------------------------------------------------------------
    'Validate the evaluation point over the full finite Double range
        If Not PROB_IsFinite(X) Then
            FailMsg = "X must be a finite number"
            Exit Function
        End If

    'Validate the rate over the full finite positive Double range
        If Not PROB_IsPositiveFinite(Lambda) Then
            FailMsg = "Lambda must be a finite strictly positive number"
            Exit Function
        End If

'------------------------------------------------------------------------------
' RETURN SUCCESS
'------------------------------------------------------------------------------
    'Report successful validation
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
'   Validates two finite continuous-Uniform bounds with UpperBound strictly
'   greater than LowerBound.
'
' INPUTS
'   LowerBound  Lower support bound.
'   UpperBound  Upper support bound.
'
' OUTPUTS
'   FailMsg  Empty on success; detailed validation message on failure.
'
' RETURNS
'   Boolean
'     True  => Both bounds are finite and correctly ordered.
'     False => At least one bound is invalid or the support is degenerate.
'
' BEHAVIOR
'   - Uses true finiteness rather than the algorithmic supported-magnitude cap.
'   - Does not form UpperBound - LowerBound during validation.
'
' CALLED FROM
'   - PROB_CN_ValidateXBounds
'   - K_STATS_Uniform_InverseCumulative
'
' UPDATED
'   2026-07-21
'==============================================================================
'
'------------------------------------------------------------------------------
' VALIDATE INPUTS
'------------------------------------------------------------------------------
    'Validate the lower support bound
        If Not PROB_IsFinite(LowerBound) Then
            FailMsg = "LowerBound must be a finite number"
            Exit Function
        End If

    'Validate the upper support bound
        If Not PROB_IsFinite(UpperBound) Then
            FailMsg = "UpperBound must be a finite number"
            Exit Function
        End If

    'Require a non-degenerate, correctly ordered support
        If Not (UpperBound > LowerBound) Then
            FailMsg = _
                "UpperBound must be strictly greater than LowerBound"
            Exit Function
        End If

'------------------------------------------------------------------------------
' RETURN SUCCESS
'------------------------------------------------------------------------------
    'Report successful validation
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
'   Validates a finite evaluation point and two finite, correctly ordered
'   continuous-Uniform support bounds.
'
' INPUTS
'   X           Evaluation point.
'   LowerBound  Lower support bound.
'   UpperBound  Upper support bound.
'
' OUTPUTS
'   FailMsg  Empty on success; detailed validation message on failure.
'
' RETURNS
'   Boolean
'     True  => X and both bounds are valid.
'     False => At least one input is invalid.
'
' DEPENDENCIES
'   - PROB_CN_ValidateBounds
'   - PROB_IsFinite
'
' CALLED FROM
'   - K_STATS_Uniform_Density
'   - K_STATS_Uniform_Cumulative
'   - K_STATS_Uniform_Survival
'
' UPDATED
'   2026-07-21
'==============================================================================
'
'------------------------------------------------------------------------------
' VALIDATE INPUTS
'------------------------------------------------------------------------------
    'Validate the evaluation point over the full finite Double range
        If Not PROB_IsFinite(X) Then
            FailMsg = "X must be a finite number"
            Exit Function
        End If

    'Validate the ordered support bounds
        If Not PROB_CN_ValidateBounds( _
            LowerBound, UpperBound, FailMsg) Then
            Exit Function
        End If

'------------------------------------------------------------------------------
' RETURN SUCCESS
'------------------------------------------------------------------------------
    'Report successful validation
        PROB_CN_ValidateXBounds = True
End Function


