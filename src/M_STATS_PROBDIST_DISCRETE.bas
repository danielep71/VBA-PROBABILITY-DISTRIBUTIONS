Attribute VB_Name = "M_STATS_PROBDIST_DISCRETE"
Option Explicit

'==============================================================================
' M_STATS_PROBDIST_DISCRETE
'------------------------------------------------------------------------------
' PURPOSE
'   Provides worksheet-facing functions for the Binomial, Poisson and
'   Geometric discrete distributions.
'
' WHY THIS EXISTS
'   Extends the library beyond continuous distributions into the discrete
'   domain. Mass functions and tail probabilities are assembled in log space
'   so that values Excel collapses to zero remain accurate deep into the
'   tails. This is Batch 1 of the discrete module; the negative binomial,
'   hypergeometric and discrete uniform arrive in Batch 2.
'
' PUBLIC API
'   Binomial
'     K_STATS_Binomial_PMF
'     K_STATS_Binomial_Cumulative
'     K_STATS_Binomial_Survival
'     K_STATS_Binomial_InverseCumulative
'     K_STATS_Binomial_Mean
'     K_STATS_Binomial_Variance
'     K_STATS_Binomial_StdDev
'
'   Poisson
'     K_STATS_Poisson_PMF
'     K_STATS_Poisson_Cumulative
'     K_STATS_Poisson_Survival
'     K_STATS_Poisson_InverseCumulative
'     K_STATS_Poisson_Mean
'     K_STATS_Poisson_Variance
'     K_STATS_Poisson_StdDev
'
'   Geometric (number of failures before the first success)
'     K_STATS_Geometric_PMF
'     K_STATS_Geometric_Cumulative
'     K_STATS_Geometric_Survival
'     K_STATS_Geometric_InverseCumulative
'     K_STATS_Geometric_Mean
'     K_STATS_Geometric_Variance
'     K_STATS_Geometric_StdDev
'
' PARAMETERIZATION
'   Binomial(NumberSuccesses, Trials, ProbSuccess)
'     Matches BINOM.DIST / BINOM.INV. ProbSuccess in [0, 1].
'   Poisson(NumberEvents, Mean)
'     Matches POISSON.DIST. Mean >= 0.
'   Geometric(NumberFailures, ProbSuccess)
'     Support k = 0, 1, 2, ... counts failures before the first success, so
'     PMF = p(1-p)^k and the distribution is the r = 1 negative binomial.
'     ProbSuccess in (0, 1]. Excel has no native geometric function.
'
' COUNT ARGUMENTS
'   Count arguments (NumberSuccesses, Trials, NumberEvents, NumberFailures)
'   are truncated toward zero to an integer, as Excel does. A negative or
'   non-finite count is a #NUM! error. For the Binomial, NumberSuccesses
'   must satisfy 0 <= NumberSuccesses <= Trials.
'
' ROUTING
'   Binomial CDF  P(X<=k) = I(1-p; N-k, k+1)   via PROB_TryBetaRegularized
'   Binomial SF   P(X>k)  = I(p; k+1, N-k)     via PROB_TryBetaRegularized
'   Poisson  CDF  P(X<=k) = Q(k+1, lambda)     via PROB_TryGammaRegularizedQ
'   Poisson  SF   P(X>k)  = P(k+1, lambda)     via PROB_TryGammaRegularizedP
'   Geometric CDF/SF via PROB_Expm1 / PROB_Log1p for left- and right-tail
'   precision. Mass functions use PROB_LogChoose / PROB_LogGamma in log space.
'
' DEPENDENCIES
'   - M_STATS_PROBDIST_CORE
'   - M_STATS_PROBDIST_SPECIALFUNCS
'
' UPDATED
'   2026-07-19 - Batch 1: Binomial, Poisson, Geometric.
'==============================================================================

'Upper bound for the unbounded (Poisson) quantile search; a result beyond
'this indicates non-convergence rather than a plausible integer quantile.
Private Const PROB_DS_INVERSE_CEILING As Double = 1E+15


'
'==============================================================================
' K_STATS_Binomial_PMF
'------------------------------------------------------------------------------
' PURPOSE
'   Returns the Binomial probability mass P(X = NumberSuccesses).
'
' WORKSHEET EQUIVALENT
'   BINOM.DIST(NumberSuccesses, Trials, ProbSuccess, FALSE)
'
' INPUTS
'   NumberSuccessesSuccesses k; truncated to an integer in [0, Trials].
'   Trials      Number of trials n; truncated to a non-negative integer.
'   ProbSuccess Success probability p in [0, 1].
'   Status      Optional ByRef diagnostic message.
'
' RETURNS
'   Variant
'     Success => Double mass in [0, 1].
'     Failure => CVErr(xlErrNum) or CVErr(xlErrValue).
'
' BEHAVIOR
'   - Assembles Log(C(n,k)) + k*Log(p) + (n-k)*Log(1-p) and exponentiates.
'   - Exponential underflow in the deep tail is a valid zero.
'   - Handles the degenerate p = 0 and p = 1 boundaries exactly.
'
' ERROR POLICY
'   - Invalid parameters or k outside [0, n] return #NUM!.
'   - Unexpected runtime errors return #VALUE!.
'
' DEPENDENCIES
'   - PROB_DS_ValidateBinomialCount
'   - PROB_LogChoose
'   - PROB_Log1p
'   - PROB_TryExp
'   - PROB_SetStatus
'
' CALLED FROM
'   - Worksheet formulas
'   - M_STATS_PROBDIST_TEST
'
' UPDATED
'   2026-07-19
'==============================================================================
'
Public Function K_STATS_Binomial_PMF( _
    ByVal NumberSuccesses As Double, _
    ByVal Trials As Double, _
    ByVal ProbSuccess As Double, _
    Optional ByRef Status As String = "") _
    As Variant
'
'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Dim K                   As Double          'Successes as a truncated integer
    Dim N                   As Double          'Trials as a truncated integer
    Dim LogMass             As Double          'Natural log of the mass
    Dim Value               As Double          'Computed mass
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
    'Validate and truncate the two counts and the success probability
        If Not PROB_DS_ValidateBinomialCount( _
            NumberSuccesses, Trials, ProbSuccess, K, N, FailMsg) Then GoTo Fail_Num

'------------------------------------------------------------------------------
' COMPUTE MASS
'------------------------------------------------------------------------------
    'Degenerate success probabilities place all mass at one point
        If ProbSuccess <= 0# Then
            If K = 0# Then Value = 1# Else Value = 0#
            K_STATS_Binomial_PMF = Value
            GoTo Return_Success
        End If
        If ProbSuccess >= 1# Then
            If K = N Then Value = 1# Else Value = 0#
            K_STATS_Binomial_PMF = Value
            GoTo Return_Success
        End If

    'Assemble the log mass, then exponentiate (deep-tail underflow is zero)
        LogMass = PROB_LogChoose(N, K) _
            + K * Log(ProbSuccess) _
            + (N - K) * PROB_Log1p(-ProbSuccess)
        If Not PROB_TryExp(LogMass, Value) Then
            FailMsg = "Binomial mass overflowed a Double"
            GoTo Fail_Num
        End If

        K_STATS_Binomial_PMF = Value

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
        K_STATS_Binomial_PMF = CVErr(xlErrNum)
    'Exit before the error handler
        Exit Function

'------------------------------------------------------------------------------
' ERROR HANDLER
'------------------------------------------------------------------------------
Err_Handler:
    'Write unexpected runtime errors to diagnostics
        PROB_SetStatus Status, "Unexpected error in K_STATS_Binomial_PMF: " & Err.Description
    'Return worksheet value error
        K_STATS_Binomial_PMF = CVErr(xlErrValue)
End Function


'
'==============================================================================
' K_STATS_Binomial_Cumulative
'------------------------------------------------------------------------------
' PURPOSE
'   Returns the Binomial left-tail probability P(X <= NumberSuccesses).
'
' WORKSHEET EQUIVALENT
'   BINOM.DIST(NumberSuccesses, Trials, ProbSuccess, TRUE)
'
' INPUTS
'   NumberSuccessesSuccesses k; truncated to an integer in [0, Trials].
'   Trials      Number of trials n; truncated to a non-negative integer.
'   ProbSuccess Success probability p in [0, 1].
'   Status      Optional ByRef diagnostic message.
'
' RETURNS
'   Variant
'     Success => Double probability in [0, 1].
'     Failure => CVErr(xlErrNum) or CVErr(xlErrValue).
'
' BEHAVIOR
'   - Routes to the regularized incomplete beta function.
'   - Returns one when k reaches the last trial.
'
' ERROR POLICY
'   - Invalid parameters or k outside [0, n] return #NUM!.
'   - Kernel non-convergence returns #NUM!.
'   - Unexpected runtime errors return #VALUE!.
'
' DEPENDENCIES
'   - PROB_DS_ValidateBinomialCount
'   - PROB_DS_TryBinomialCDF
'   - PROB_SetStatus
'
' CALLED FROM
'   - Worksheet formulas
'   - M_STATS_PROBDIST_TEST
'
' UPDATED
'   2026-07-19
'==============================================================================
'
Public Function K_STATS_Binomial_Cumulative( _
    ByVal NumberSuccesses As Double, _
    ByVal Trials As Double, _
    ByVal ProbSuccess As Double, _
    Optional ByRef Status As String = "") _
    As Variant
'
'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Dim K                   As Double          'Successes as a truncated integer
    Dim N                   As Double          'Trials as a truncated integer
    Dim Value               As Double          'Computed probability
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
    'Validate and truncate the two counts and the success probability
        If Not PROB_DS_ValidateBinomialCount( _
            NumberSuccesses, Trials, ProbSuccess, K, N, FailMsg) Then GoTo Fail_Num

'------------------------------------------------------------------------------
' COMPUTE PROBABILITY
'------------------------------------------------------------------------------
    'P(X<=k) = I(1-p; N-k, k+1), evaluated through the shared beta kernel
        If Not PROB_DS_TryBinomialCDF(K, N, ProbSuccess, Value, FailMsg) Then GoTo Fail_Num

        K_STATS_Binomial_Cumulative = Value

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
        K_STATS_Binomial_Cumulative = CVErr(xlErrNum)
    'Exit before the error handler
        Exit Function

'------------------------------------------------------------------------------
' ERROR HANDLER
'------------------------------------------------------------------------------
Err_Handler:
    'Write unexpected runtime errors to diagnostics
        PROB_SetStatus Status, "Unexpected error in K_STATS_Binomial_Cumulative: " & Err.Description
    'Return worksheet value error
        K_STATS_Binomial_Cumulative = CVErr(xlErrValue)
End Function


'
'==============================================================================
' K_STATS_Binomial_Survival
'------------------------------------------------------------------------------
' PURPOSE
'   Returns the Binomial right-tail probability P(X > NumberSuccesses).
'
' WORKSHEET EQUIVALENT
'   1 - BINOM.DIST(NumberSuccesses, Trials, ProbSuccess, TRUE)
'
' INPUTS
'   NumberSuccessesSuccesses k; truncated to an integer in [0, Trials].
'   Trials      Number of trials n; truncated to a non-negative integer.
'   ProbSuccess Success probability p in [0, 1].
'   Status      Optional ByRef diagnostic message.
'
' RETURNS
'   Variant
'     Success => Double probability in [0, 1].
'     Failure => CVErr(xlErrNum) or CVErr(xlErrValue).
'
' BEHAVIOR
'   - Computes the upper tail directly rather than as 1 - CDF.
'   - Returns zero when k reaches the last trial.
'
' ERROR POLICY
'   - Invalid parameters or k outside [0, n] return #NUM!.
'   - Kernel non-convergence returns #NUM!.
'   - Unexpected runtime errors return #VALUE!.
'
' DEPENDENCIES
'   - PROB_DS_ValidateBinomialCount
'   - PROB_DS_TryBinomialSF
'   - PROB_SetStatus
'
' CALLED FROM
'   - Worksheet formulas
'   - M_STATS_PROBDIST_TEST
'
' UPDATED
'   2026-07-19
'==============================================================================
'
Public Function K_STATS_Binomial_Survival( _
    ByVal NumberSuccesses As Double, _
    ByVal Trials As Double, _
    ByVal ProbSuccess As Double, _
    Optional ByRef Status As String = "") _
    As Variant
'
'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Dim K                   As Double          'Successes as a truncated integer
    Dim N                   As Double          'Trials as a truncated integer
    Dim Value               As Double          'Computed probability
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
    'Validate and truncate the two counts and the success probability
        If Not PROB_DS_ValidateBinomialCount( _
            NumberSuccesses, Trials, ProbSuccess, K, N, FailMsg) Then GoTo Fail_Num

'------------------------------------------------------------------------------
' COMPUTE PROBABILITY
'------------------------------------------------------------------------------
    'P(X>k) = I(p; k+1, N-k), computed directly for right-tail accuracy
        If Not PROB_DS_TryBinomialSF(K, N, ProbSuccess, Value, FailMsg) Then GoTo Fail_Num

        K_STATS_Binomial_Survival = Value

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
        K_STATS_Binomial_Survival = CVErr(xlErrNum)
    'Exit before the error handler
        Exit Function

'------------------------------------------------------------------------------
' ERROR HANDLER
'------------------------------------------------------------------------------
Err_Handler:
    'Write unexpected runtime errors to diagnostics
        PROB_SetStatus Status, "Unexpected error in K_STATS_Binomial_Survival: " & Err.Description
    'Return worksheet value error
        K_STATS_Binomial_Survival = CVErr(xlErrValue)
End Function


'
'==============================================================================
' K_STATS_Binomial_InverseCumulative
'------------------------------------------------------------------------------
' PURPOSE
'   Returns the smallest NumberSuccesses whose cumulative probability is at
'   least Probability.
'
' WHY THIS EXISTS
'   The cumulative Binomial is a step function, so the inverse is found by
'   an integer lower-bound bisection over [0, Trials] on the shared CDF.
'
' WORKSHEET EQUIVALENT
'   BINOM.INV(Trials, ProbSuccess, Probability)
'
' INPUTS
'   Probability Target cumulative probability in the open (0, 1).
'   Trials      Number of trials n; truncated to a non-negative integer.
'   ProbSuccess Success probability p in [0, 1].
'   Status      Optional ByRef diagnostic message.
'
' RETURNS
'   Variant
'     Success => Double integer quantile in [0, Trials].
'     Failure => CVErr(xlErrNum) or CVErr(xlErrValue).
'
' BEHAVIOR
'   - Returns the least k with P(X <= k) >= Probability.
'
' ERROR POLICY
'   - Probability outside (0, 1) or invalid parameters return #NUM!.
'   - Kernel non-convergence returns #NUM!.
'   - Unexpected runtime errors return #VALUE!.
'
' DEPENDENCIES
'   - PROB_IsValidProbabilityOpen
'   - PROB_DS_ValidateTrials
'   - PROB_DS_ValidateProbClosed
'   - PROB_DS_TryBinomialInverse
'   - PROB_SetStatus
'
' CALLED FROM
'   - Worksheet formulas
'   - M_STATS_PROBDIST_TEST
'
' UPDATED
'   2026-07-19
'==============================================================================
'
Public Function K_STATS_Binomial_InverseCumulative( _
    ByVal Probability As Double, _
    ByVal Trials As Double, _
    ByVal ProbSuccess As Double, _
    Optional ByRef Status As String = "") _
    As Variant
'
'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Dim N                   As Double          'Trials as a truncated integer
    Dim Quantile            As Double          'Computed integer quantile
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
    'Validate the probability domain
        If Not PROB_IsValidProbabilityOpen(Probability) Then
            FailMsg = "Probability must be strictly between 0 and 1"
            GoTo Fail_Num
        End If
    'Validate and truncate Trials
        If Not PROB_DS_ValidateTrials(Trials, N, FailMsg) Then GoTo Fail_Num
    'Validate the success probability
        If Not PROB_DS_ValidateProbClosed(ProbSuccess, FailMsg) Then GoTo Fail_Num

'------------------------------------------------------------------------------
' COMPUTE QUANTILE
'------------------------------------------------------------------------------
    'Bisect the cumulative step function for the least qualifying k
        If Not PROB_DS_TryBinomialInverse( _
            Probability, N, ProbSuccess, Quantile, FailMsg) Then GoTo Fail_Num

        K_STATS_Binomial_InverseCumulative = Quantile

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
        K_STATS_Binomial_InverseCumulative = CVErr(xlErrNum)
    'Exit before the error handler
        Exit Function

'------------------------------------------------------------------------------
' ERROR HANDLER
'------------------------------------------------------------------------------
Err_Handler:
    'Write unexpected runtime errors to diagnostics
        PROB_SetStatus Status, "Unexpected error in K_STATS_Binomial_InverseCumulative: " & Err.Description
    'Return worksheet value error
        K_STATS_Binomial_InverseCumulative = CVErr(xlErrValue)
End Function


'
'==============================================================================
' K_STATS_Binomial_Mean
'------------------------------------------------------------------------------
' PURPOSE
'   Returns the Binomial mean, Trials * ProbSuccess.
'
' INPUTS
'   Trials      Number of trials n; truncated to a non-negative integer.
'   ProbSuccess Success probability p in [0, 1].
'   Status      Optional ByRef diagnostic message.
'
' RETURNS
'   Variant
'     Success => Double value.
'     Failure => CVErr(xlErrNum) or CVErr(xlErrValue).
'
' ERROR POLICY
'   - Invalid inputs or overflow return #NUM!.
'   - Unexpected runtime errors return #VALUE!.
'
' DEPENDENCIES
'   - PROB_DS_ValidateTrials
'   - PROB_DS_ValidateProbClosed
'   - PROB_TryMultiply
'   - PROB_SetStatus
'
' CALLED FROM
'   - Worksheet formulas
'   - M_STATS_PROBDIST_TEST
'
' UPDATED
'   2026-07-19
'==============================================================================
'
Public Function K_STATS_Binomial_Mean( _
    ByVal Trials As Double, _
    ByVal ProbSuccess As Double, _
    Optional ByRef Status As String = "") _
    As Variant
'
'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Dim N                   As Double          'Trials as a truncated integer
    Dim Value               As Double          'Computed mean
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
    'Validate and truncate Trials
        If Not PROB_DS_ValidateTrials(Trials, N, FailMsg) Then GoTo Fail_Num
    'Validate the success probability
        If Not PROB_DS_ValidateProbClosed(ProbSuccess, FailMsg) Then GoTo Fail_Num

'------------------------------------------------------------------------------
' COMPUTE
'------------------------------------------------------------------------------
    'Form n*p through the shared multiplication contract
        If Not PROB_TryMultiply(N, ProbSuccess, Value) Then
            FailMsg = "Binomial mean overflows Double range"
            GoTo Fail_Num
        End If

        K_STATS_Binomial_Mean = Value

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
        K_STATS_Binomial_Mean = CVErr(xlErrNum)
    'Exit before the error handler
        Exit Function

'------------------------------------------------------------------------------
' ERROR HANDLER
'------------------------------------------------------------------------------
Err_Handler:
    'Write unexpected runtime errors to diagnostics
        PROB_SetStatus Status, "Unexpected error in K_STATS_Binomial_Mean: " & Err.Description
    'Return worksheet value error
        K_STATS_Binomial_Mean = CVErr(xlErrValue)
End Function


'
'==============================================================================
' K_STATS_Binomial_Variance
'------------------------------------------------------------------------------
' PURPOSE
'   Returns the Binomial variance, Trials * ProbSuccess * (1 - ProbSuccess).
'
' INPUTS
'   Trials      Number of trials n; truncated to a non-negative integer.
'   ProbSuccess Success probability p in [0, 1].
'   Status      Optional ByRef diagnostic message.
'
' RETURNS
'   Variant
'     Success => Double value.
'     Failure => CVErr(xlErrNum) or CVErr(xlErrValue).
'
' ERROR POLICY
'   - Invalid inputs or overflow return #NUM!.
'   - Unexpected runtime errors return #VALUE!.
'
' DEPENDENCIES
'   - PROB_DS_ValidateTrials
'   - PROB_DS_ValidateProbClosed
'   - PROB_TryMultiply
'   - PROB_SetStatus
'
' CALLED FROM
'   - Worksheet formulas
'   - M_STATS_PROBDIST_TEST
'
' UPDATED
'   2026-07-19
'==============================================================================
'
Public Function K_STATS_Binomial_Variance( _
    ByVal Trials As Double, _
    ByVal ProbSuccess As Double, _
    Optional ByRef Status As String = "") _
    As Variant
'
'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Dim N                   As Double          'Trials as a truncated integer
    Dim Np                  As Double          'Intermediate n*p
    Dim Value               As Double          'Computed variance
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
    'Validate and truncate Trials
        If Not PROB_DS_ValidateTrials(Trials, N, FailMsg) Then GoTo Fail_Num
    'Validate the success probability
        If Not PROB_DS_ValidateProbClosed(ProbSuccess, FailMsg) Then GoTo Fail_Num

'------------------------------------------------------------------------------
' COMPUTE
'------------------------------------------------------------------------------
    'Form n*p, then multiply by the failure probability
        If Not PROB_TryMultiply(N, ProbSuccess, Np) Then
            FailMsg = "Binomial variance overflows Double range"
            GoTo Fail_Num
        End If
        If Not PROB_TryMultiply(Np, 1# - ProbSuccess, Value) Then
            FailMsg = "Binomial variance overflows Double range"
            GoTo Fail_Num
        End If

        K_STATS_Binomial_Variance = Value

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
        K_STATS_Binomial_Variance = CVErr(xlErrNum)
    'Exit before the error handler
        Exit Function

'------------------------------------------------------------------------------
' ERROR HANDLER
'------------------------------------------------------------------------------
Err_Handler:
    'Write unexpected runtime errors to diagnostics
        PROB_SetStatus Status, "Unexpected error in K_STATS_Binomial_Variance: " & Err.Description
    'Return worksheet value error
        K_STATS_Binomial_Variance = CVErr(xlErrValue)
End Function


'
'==============================================================================
' K_STATS_Binomial_StdDev
'------------------------------------------------------------------------------
' PURPOSE
'   Returns the Binomial standard deviation.
'
' INPUTS
'   Trials      Number of trials n; truncated to a non-negative integer.
'   ProbSuccess Success probability p in [0, 1].
'   Status      Optional ByRef diagnostic message.
'
' RETURNS
'   Variant
'     Success => Double value.
'     Failure => CVErr(xlErrNum) or CVErr(xlErrValue).
'
' ERROR POLICY
'   - Invalid inputs or overflow return #NUM!.
'   - Unexpected runtime errors return #VALUE!.
'
' DEPENDENCIES
'   - PROB_DS_ValidateTrials
'   - PROB_DS_ValidateProbClosed
'   - PROB_TryMultiply
'   - PROB_SetStatus
'
' CALLED FROM
'   - Worksheet formulas
'   - M_STATS_PROBDIST_TEST
'
' UPDATED
'   2026-07-19
'==============================================================================
'
Public Function K_STATS_Binomial_StdDev( _
    ByVal Trials As Double, _
    ByVal ProbSuccess As Double, _
    Optional ByRef Status As String = "") _
    As Variant
'
'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Dim N                   As Double          'Trials as a truncated integer
    Dim Np                  As Double          'Intermediate n*p
    Dim Variance            As Double          'Computed variance
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
    'Validate and truncate Trials
        If Not PROB_DS_ValidateTrials(Trials, N, FailMsg) Then GoTo Fail_Num
    'Validate the success probability
        If Not PROB_DS_ValidateProbClosed(ProbSuccess, FailMsg) Then GoTo Fail_Num

'------------------------------------------------------------------------------
' COMPUTE
'------------------------------------------------------------------------------
    'Form the variance, then take its non-negative square root
        If Not PROB_TryMultiply(N, ProbSuccess, Np) Then
            FailMsg = "Binomial standard deviation overflows Double range"
            GoTo Fail_Num
        End If
        If Not PROB_TryMultiply(Np, 1# - ProbSuccess, Variance) Then
            FailMsg = "Binomial standard deviation overflows Double range"
            GoTo Fail_Num
        End If

        K_STATS_Binomial_StdDev = Sqr(Variance)

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
        K_STATS_Binomial_StdDev = CVErr(xlErrNum)
    'Exit before the error handler
        Exit Function

'------------------------------------------------------------------------------
' ERROR HANDLER
'------------------------------------------------------------------------------
Err_Handler:
    'Write unexpected runtime errors to diagnostics
        PROB_SetStatus Status, "Unexpected error in K_STATS_Binomial_StdDev: " & Err.Description
    'Return worksheet value error
        K_STATS_Binomial_StdDev = CVErr(xlErrValue)
End Function


'
'==============================================================================
' K_STATS_Poisson_PMF
'------------------------------------------------------------------------------
' PURPOSE
'   Returns the Poisson probability mass P(X = NumberEvents).
'
' WORKSHEET EQUIVALENT
'   POISSON.DIST(NumberEvents, Mean, FALSE)
'
' INPUTS
'   NumberEventsEvent count k; truncated to a non-negative integer.
'   Mean        Poisson mean lambda >= 0.
'   Status      Optional ByRef diagnostic message.
'
' RETURNS
'   Variant
'     Success => Double mass in [0, 1].
'     Failure => CVErr(xlErrNum) or CVErr(xlErrValue).
'
' BEHAVIOR
'   - Assembles k*Log(lambda) - lambda - LogGamma(k+1) and exponentiates.
'   - Exponential underflow in the deep tail is a valid zero.
'   - Mean = 0 places all mass at k = 0.
'
' ERROR POLICY
'   - Invalid parameters return #NUM!.
'   - Unexpected runtime errors return #VALUE!.
'
' DEPENDENCIES
'   - PROB_DS_ValidateCount
'   - PROB_IsFinite
'   - PROB_LogGamma
'   - PROB_TryExp
'   - PROB_SetStatus
'
' CALLED FROM
'   - Worksheet formulas
'   - M_STATS_PROBDIST_TEST
'
' UPDATED
'   2026-07-19
'==============================================================================
'
Public Function K_STATS_Poisson_PMF( _
    ByVal NumberEvents As Double, _
    ByVal Mean As Double, _
    Optional ByRef Status As String = "") _
    As Variant
'
'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Dim K                   As Double          'Events as a truncated integer
    Dim LogMass             As Double          'Natural log of the mass
    Dim Value               As Double          'Computed mass
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
    'Validate and truncate the event count
        If Not PROB_DS_ValidateCount(NumberEvents, K, FailMsg, "NumberEvents") Then GoTo Fail_Num
    'Validate the mean over the non-negative finite range
        If Not PROB_IsFinite(Mean) Or Mean < 0# Then
            FailMsg = "Mean must be a finite non-negative number"
            GoTo Fail_Num
        End If

'------------------------------------------------------------------------------
' COMPUTE MASS
'------------------------------------------------------------------------------
    'A zero mean is a point mass at k = 0
        If Mean <= 0# Then
            If K = 0# Then Value = 1# Else Value = 0#
            K_STATS_Poisson_PMF = Value
            GoTo Return_Success
        End If

    'Assemble the log mass, then exponentiate (deep-tail underflow is zero)
        LogMass = K * Log(Mean) - Mean - PROB_LogGamma(K + 1#)
        If Not PROB_TryExp(LogMass, Value) Then
            FailMsg = "Poisson mass overflowed a Double"
            GoTo Fail_Num
        End If

        K_STATS_Poisson_PMF = Value

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
        K_STATS_Poisson_PMF = CVErr(xlErrNum)
    'Exit before the error handler
        Exit Function

'------------------------------------------------------------------------------
' ERROR HANDLER
'------------------------------------------------------------------------------
Err_Handler:
    'Write unexpected runtime errors to diagnostics
        PROB_SetStatus Status, "Unexpected error in K_STATS_Poisson_PMF: " & Err.Description
    'Return worksheet value error
        K_STATS_Poisson_PMF = CVErr(xlErrValue)
End Function


'
'==============================================================================
' K_STATS_Poisson_Cumulative
'------------------------------------------------------------------------------
' PURPOSE
'   Returns the Poisson left-tail probability P(X <= NumberEvents).
'
' WORKSHEET EQUIVALENT
'   POISSON.DIST(NumberEvents, Mean, TRUE)
'
' INPUTS
'   NumberEventsEvent count k; truncated to a non-negative integer.
'   Mean        Poisson mean lambda >= 0.
'   Status      Optional ByRef diagnostic message.
'
' RETURNS
'   Variant
'     Success => Double probability in [0, 1].
'     Failure => CVErr(xlErrNum) or CVErr(xlErrValue).
'
' BEHAVIOR
'   - Routes to the regularized upper incomplete gamma function.
'   - Mean = 0 returns one.
'
' ERROR POLICY
'   - Invalid parameters or kernel non-convergence return #NUM!.
'   - Unexpected runtime errors return #VALUE!.
'
' DEPENDENCIES
'   - PROB_DS_ValidateCount
'   - PROB_IsFinite
'   - PROB_DS_TryPoissonCDF
'   - PROB_SetStatus
'
' CALLED FROM
'   - Worksheet formulas
'   - M_STATS_PROBDIST_TEST
'
' UPDATED
'   2026-07-19
'==============================================================================
'
Public Function K_STATS_Poisson_Cumulative( _
    ByVal NumberEvents As Double, _
    ByVal Mean As Double, _
    Optional ByRef Status As String = "") _
    As Variant
'
'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Dim K                   As Double          'Events as a truncated integer
    Dim Value               As Double          'Computed probability
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
    'Validate and truncate the event count
        If Not PROB_DS_ValidateCount(NumberEvents, K, FailMsg, "NumberEvents") Then GoTo Fail_Num
    'Validate the mean over the non-negative finite range
        If Not PROB_IsFinite(Mean) Or Mean < 0# Then
            FailMsg = "Mean must be a finite non-negative number"
            GoTo Fail_Num
        End If

'------------------------------------------------------------------------------
' COMPUTE PROBABILITY
'------------------------------------------------------------------------------
    'P(X<=k) = Q(k+1, lambda), the regularized upper incomplete gamma
        If Not PROB_DS_TryPoissonCDF(K, Mean, Value, FailMsg) Then GoTo Fail_Num

        K_STATS_Poisson_Cumulative = Value

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
        K_STATS_Poisson_Cumulative = CVErr(xlErrNum)
    'Exit before the error handler
        Exit Function

'------------------------------------------------------------------------------
' ERROR HANDLER
'------------------------------------------------------------------------------
Err_Handler:
    'Write unexpected runtime errors to diagnostics
        PROB_SetStatus Status, "Unexpected error in K_STATS_Poisson_Cumulative: " & Err.Description
    'Return worksheet value error
        K_STATS_Poisson_Cumulative = CVErr(xlErrValue)
End Function


'
'==============================================================================
' K_STATS_Poisson_Survival
'------------------------------------------------------------------------------
' PURPOSE
'   Returns the Poisson right-tail probability P(X > NumberEvents).
'
' WORKSHEET EQUIVALENT
'   1 - POISSON.DIST(NumberEvents, Mean, TRUE)
'
' INPUTS
'   NumberEventsEvent count k; truncated to a non-negative integer.
'   Mean        Poisson mean lambda >= 0.
'   Status      Optional ByRef diagnostic message.
'
' RETURNS
'   Variant
'     Success => Double probability in [0, 1].
'     Failure => CVErr(xlErrNum) or CVErr(xlErrValue).
'
' BEHAVIOR
'   - Computes the upper tail directly rather than as 1 - CDF.
'   - Mean = 0 returns zero.
'
' ERROR POLICY
'   - Invalid parameters or kernel non-convergence return #NUM!.
'   - Unexpected runtime errors return #VALUE!.
'
' DEPENDENCIES
'   - PROB_DS_ValidateCount
'   - PROB_IsFinite
'   - PROB_DS_TryPoissonSF
'   - PROB_SetStatus
'
' CALLED FROM
'   - Worksheet formulas
'   - M_STATS_PROBDIST_TEST
'
' UPDATED
'   2026-07-19
'==============================================================================
'
Public Function K_STATS_Poisson_Survival( _
    ByVal NumberEvents As Double, _
    ByVal Mean As Double, _
    Optional ByRef Status As String = "") _
    As Variant
'
'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Dim K                   As Double          'Events as a truncated integer
    Dim Value               As Double          'Computed probability
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
    'Validate and truncate the event count
        If Not PROB_DS_ValidateCount(NumberEvents, K, FailMsg, "NumberEvents") Then GoTo Fail_Num
    'Validate the mean over the non-negative finite range
        If Not PROB_IsFinite(Mean) Or Mean < 0# Then
            FailMsg = "Mean must be a finite non-negative number"
            GoTo Fail_Num
        End If

'------------------------------------------------------------------------------
' COMPUTE PROBABILITY
'------------------------------------------------------------------------------
    'P(X>k) = P(k+1, lambda), the regularized lower incomplete gamma
        If Not PROB_DS_TryPoissonSF(K, Mean, Value, FailMsg) Then GoTo Fail_Num

        K_STATS_Poisson_Survival = Value

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
        K_STATS_Poisson_Survival = CVErr(xlErrNum)
    'Exit before the error handler
        Exit Function

'------------------------------------------------------------------------------
' ERROR HANDLER
'------------------------------------------------------------------------------
Err_Handler:
    'Write unexpected runtime errors to diagnostics
        PROB_SetStatus Status, "Unexpected error in K_STATS_Poisson_Survival: " & Err.Description
    'Return worksheet value error
        K_STATS_Poisson_Survival = CVErr(xlErrValue)
End Function


'
'==============================================================================
' K_STATS_Poisson_InverseCumulative
'------------------------------------------------------------------------------
' PURPOSE
'   Returns the smallest NumberEvents whose cumulative probability is at
'   least Probability.
'
' WHY THIS EXISTS
'   The support is unbounded above, so the quantile is bracketed by an
'   exponential search and then found by integer bisection.
'
' INPUTS
'   Probability Target cumulative probability in the open (0, 1).
'   Mean        Poisson mean lambda >= 0.
'   Status      Optional ByRef diagnostic message.
'
' RETURNS
'   Variant
'     Success => Double integer quantile.
'     Failure => CVErr(xlErrNum) or CVErr(xlErrValue).
'
' BEHAVIOR
'   - Returns the least k with P(X <= k) >= Probability.
'
' ERROR POLICY
'   - Probability outside (0, 1) or invalid Mean return #NUM!.
'   - A quantile beyond the search ceiling returns #NUM!.
'   - Unexpected runtime errors return #VALUE!.
'
' DEPENDENCIES
'   - PROB_IsValidProbabilityOpen
'   - PROB_IsFinite
'   - PROB_DS_TryPoissonInverse
'   - PROB_SetStatus
'
' CALLED FROM
'   - Worksheet formulas
'   - M_STATS_PROBDIST_TEST
'
' UPDATED
'   2026-07-19
'==============================================================================
'
Public Function K_STATS_Poisson_InverseCumulative( _
    ByVal Probability As Double, _
    ByVal Mean As Double, _
    Optional ByRef Status As String = "") _
    As Variant
'
'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Dim Quantile            As Double          'Computed integer quantile
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
    'Validate the probability domain
        If Not PROB_IsValidProbabilityOpen(Probability) Then
            FailMsg = "Probability must be strictly between 0 and 1"
            GoTo Fail_Num
        End If
    'Validate the mean over the non-negative finite range
        If Not PROB_IsFinite(Mean) Or Mean < 0# Then
            FailMsg = "Mean must be a finite non-negative number"
            GoTo Fail_Num
        End If

'------------------------------------------------------------------------------
' COMPUTE QUANTILE
'------------------------------------------------------------------------------
    'Bracket, then bisect the cumulative step function
        If Not PROB_DS_TryPoissonInverse( _
            Probability, Mean, Quantile, FailMsg) Then GoTo Fail_Num

        K_STATS_Poisson_InverseCumulative = Quantile

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
        K_STATS_Poisson_InverseCumulative = CVErr(xlErrNum)
    'Exit before the error handler
        Exit Function

'------------------------------------------------------------------------------
' ERROR HANDLER
'------------------------------------------------------------------------------
Err_Handler:
    'Write unexpected runtime errors to diagnostics
        PROB_SetStatus Status, "Unexpected error in K_STATS_Poisson_InverseCumulative: " & Err.Description
    'Return worksheet value error
        K_STATS_Poisson_InverseCumulative = CVErr(xlErrValue)
End Function


'
'==============================================================================
' K_STATS_Poisson_Mean
'------------------------------------------------------------------------------
' PURPOSE
'   Returns the Poisson mean, which equals lambda.
'
' INPUTS
'   Mean        Poisson mean lambda >= 0.
'   Status      Optional ByRef diagnostic message.
'
' RETURNS
'   Variant
'     Success => Double value.
'     Failure => CVErr(xlErrNum) or CVErr(xlErrValue).
'
' ERROR POLICY
'   - Invalid inputs or overflow return #NUM!.
'   - Unexpected runtime errors return #VALUE!.
'
' DEPENDENCIES
'   - PROB_IsFinite
'   - PROB_SetStatus
'
' CALLED FROM
'   - Worksheet formulas
'   - M_STATS_PROBDIST_TEST
'
' UPDATED
'   2026-07-19
'==============================================================================
'
Public Function K_STATS_Poisson_Mean( _
    ByVal Mean As Double, _
    Optional ByRef Status As String = "") _
    As Variant
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
    'Validate the mean over the non-negative finite range
        If Not PROB_IsFinite(Mean) Or Mean < 0# Then
            FailMsg = "Mean must be a finite non-negative number"
            GoTo Fail_Num
        End If

'------------------------------------------------------------------------------
' COMPUTE
'------------------------------------------------------------------------------
    'The Poisson mean equals its parameter
        K_STATS_Poisson_Mean = Mean

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
        K_STATS_Poisson_Mean = CVErr(xlErrNum)
    'Exit before the error handler
        Exit Function

'------------------------------------------------------------------------------
' ERROR HANDLER
'------------------------------------------------------------------------------
Err_Handler:
    'Write unexpected runtime errors to diagnostics
        PROB_SetStatus Status, "Unexpected error in K_STATS_Poisson_Mean: " & Err.Description
    'Return worksheet value error
        K_STATS_Poisson_Mean = CVErr(xlErrValue)
End Function


'
'==============================================================================
' K_STATS_Poisson_Variance
'------------------------------------------------------------------------------
' PURPOSE
'   Returns the Poisson variance, which equals lambda.
'
' INPUTS
'   Mean        Poisson mean lambda >= 0.
'   Status      Optional ByRef diagnostic message.
'
' RETURNS
'   Variant
'     Success => Double value.
'     Failure => CVErr(xlErrNum) or CVErr(xlErrValue).
'
' ERROR POLICY
'   - Invalid inputs or overflow return #NUM!.
'   - Unexpected runtime errors return #VALUE!.
'
' DEPENDENCIES
'   - PROB_IsFinite
'   - PROB_SetStatus
'
' CALLED FROM
'   - Worksheet formulas
'   - M_STATS_PROBDIST_TEST
'
' UPDATED
'   2026-07-19
'==============================================================================
'
Public Function K_STATS_Poisson_Variance( _
    ByVal Mean As Double, _
    Optional ByRef Status As String = "") _
    As Variant
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
    'Validate the mean over the non-negative finite range
        If Not PROB_IsFinite(Mean) Or Mean < 0# Then
            FailMsg = "Mean must be a finite non-negative number"
            GoTo Fail_Num
        End If

'------------------------------------------------------------------------------
' COMPUTE
'------------------------------------------------------------------------------
    'The Poisson variance equals its parameter
        K_STATS_Poisson_Variance = Mean

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
        K_STATS_Poisson_Variance = CVErr(xlErrNum)
    'Exit before the error handler
        Exit Function

'------------------------------------------------------------------------------
' ERROR HANDLER
'------------------------------------------------------------------------------
Err_Handler:
    'Write unexpected runtime errors to diagnostics
        PROB_SetStatus Status, "Unexpected error in K_STATS_Poisson_Variance: " & Err.Description
    'Return worksheet value error
        K_STATS_Poisson_Variance = CVErr(xlErrValue)
End Function


'
'==============================================================================
' K_STATS_Poisson_StdDev
'------------------------------------------------------------------------------
' PURPOSE
'   Returns the Poisson standard deviation, Sqr(lambda).
'
' INPUTS
'   Mean        Poisson mean lambda >= 0.
'   Status      Optional ByRef diagnostic message.
'
' RETURNS
'   Variant
'     Success => Double value.
'     Failure => CVErr(xlErrNum) or CVErr(xlErrValue).
'
' ERROR POLICY
'   - Invalid inputs or overflow return #NUM!.
'   - Unexpected runtime errors return #VALUE!.
'
' DEPENDENCIES
'   - PROB_IsFinite
'   - PROB_SetStatus
'
' CALLED FROM
'   - Worksheet formulas
'   - M_STATS_PROBDIST_TEST
'
' UPDATED
'   2026-07-19
'==============================================================================
'
Public Function K_STATS_Poisson_StdDev( _
    ByVal Mean As Double, _
    Optional ByRef Status As String = "") _
    As Variant
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
    'Validate the mean over the non-negative finite range
        If Not PROB_IsFinite(Mean) Or Mean < 0# Then
            FailMsg = "Mean must be a finite non-negative number"
            GoTo Fail_Num
        End If

'------------------------------------------------------------------------------
' COMPUTE
'------------------------------------------------------------------------------
    'The Poisson standard deviation is the square root of the mean
        K_STATS_Poisson_StdDev = Sqr(Mean)

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
        K_STATS_Poisson_StdDev = CVErr(xlErrNum)
    'Exit before the error handler
        Exit Function

'------------------------------------------------------------------------------
' ERROR HANDLER
'------------------------------------------------------------------------------
Err_Handler:
    'Write unexpected runtime errors to diagnostics
        PROB_SetStatus Status, "Unexpected error in K_STATS_Poisson_StdDev: " & Err.Description
    'Return worksheet value error
        K_STATS_Poisson_StdDev = CVErr(xlErrValue)
End Function


'
'==============================================================================
' K_STATS_Geometric_PMF
'------------------------------------------------------------------------------
' PURPOSE
'   Returns the Geometric probability mass P(X = NumberFailures), where X
'   counts failures before the first success.
'
' INPUTS
'   NumberFailuresFailure count k; truncated to a non-negative integer.
'   ProbSuccess Success probability p in (0, 1].
'   Status      Optional ByRef diagnostic message.
'
' RETURNS
'   Variant
'     Success => Double mass in [0, 1].
'     Failure => CVErr(xlErrNum) or CVErr(xlErrValue).
'
' BEHAVIOR
'   - Assembles Log(p) + k*Log(1-p) and exponentiates.
'   - Exponential underflow in the deep tail is a valid zero.
'   - p = 1 places all mass at k = 0.
'
' ERROR POLICY
'   - Invalid parameters return #NUM!.
'   - Unexpected runtime errors return #VALUE!.
'
' DEPENDENCIES
'   - PROB_DS_ValidateCount
'   - PROB_DS_ValidateProbHalfOpen
'   - PROB_Log1p
'   - PROB_TryExp
'   - PROB_SetStatus
'
' CALLED FROM
'   - Worksheet formulas
'   - M_STATS_PROBDIST_TEST
'
' UPDATED
'   2026-07-19
'==============================================================================
'
Public Function K_STATS_Geometric_PMF( _
    ByVal NumberFailures As Double, _
    ByVal ProbSuccess As Double, _
    Optional ByRef Status As String = "") _
    As Variant
'
'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Dim K                   As Double          'Failures as a truncated integer
    Dim LogMass             As Double          'Natural log of the mass
    Dim Value               As Double          'Computed mass
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
    'Validate and truncate the failure count
        If Not PROB_DS_ValidateCount(NumberFailures, K, FailMsg, "NumberFailures") Then GoTo Fail_Num
    'Validate the success probability over (0, 1]
        If Not PROB_DS_ValidateProbHalfOpen(ProbSuccess, FailMsg) Then GoTo Fail_Num

'------------------------------------------------------------------------------
' COMPUTE MASS
'------------------------------------------------------------------------------
    'Certain success places all mass at k = 0
        If ProbSuccess >= 1# Then
            If K = 0# Then Value = 1# Else Value = 0#
            K_STATS_Geometric_PMF = Value
            GoTo Return_Success
        End If

    'Assemble the log mass, then exponentiate (deep-tail underflow is zero)
        LogMass = Log(ProbSuccess) + K * PROB_Log1p(-ProbSuccess)
        If Not PROB_TryExp(LogMass, Value) Then
            FailMsg = "Geometric mass overflowed a Double"
            GoTo Fail_Num
        End If

        K_STATS_Geometric_PMF = Value

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
        K_STATS_Geometric_PMF = CVErr(xlErrNum)
    'Exit before the error handler
        Exit Function

'------------------------------------------------------------------------------
' ERROR HANDLER
'------------------------------------------------------------------------------
Err_Handler:
    'Write unexpected runtime errors to diagnostics
        PROB_SetStatus Status, "Unexpected error in K_STATS_Geometric_PMF: " & Err.Description
    'Return worksheet value error
        K_STATS_Geometric_PMF = CVErr(xlErrValue)
End Function


'
'==============================================================================
' K_STATS_Geometric_Cumulative
'------------------------------------------------------------------------------
' PURPOSE
'   Returns the Geometric left-tail probability P(X <= NumberFailures).
'
' INPUTS
'   NumberFailuresFailure count k; truncated to a non-negative integer.
'   ProbSuccess Success probability p in (0, 1].
'   Status      Optional ByRef diagnostic message.
'
' RETURNS
'   Variant
'     Success => Double probability in [0, 1].
'     Failure => CVErr(xlErrNum) or CVErr(xlErrValue).
'
' BEHAVIOR
'   - CDF = 1 - (1-p)^(k+1), formed through Expm1 for small-probability accuracy.
'   - p = 1 returns one.
'
' ERROR POLICY
'   - Invalid parameters return #NUM!.
'   - Unexpected runtime errors return #VALUE!.
'
' DEPENDENCIES
'   - PROB_DS_ValidateCount
'   - PROB_DS_ValidateProbHalfOpen
'   - PROB_Expm1
'   - PROB_Log1p
'   - PROB_SetStatus
'
' CALLED FROM
'   - Worksheet formulas
'   - M_STATS_PROBDIST_TEST
'
' UPDATED
'   2026-07-19
'==============================================================================
'
Public Function K_STATS_Geometric_Cumulative( _
    ByVal NumberFailures As Double, _
    ByVal ProbSuccess As Double, _
    Optional ByRef Status As String = "") _
    As Variant
'
'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Dim K                   As Double          'Failures as a truncated integer
    Dim Value               As Double          'Computed probability
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
    'Validate and truncate the failure count
        If Not PROB_DS_ValidateCount(NumberFailures, K, FailMsg, "NumberFailures") Then GoTo Fail_Num
    'Validate the success probability over (0, 1]
        If Not PROB_DS_ValidateProbHalfOpen(ProbSuccess, FailMsg) Then GoTo Fail_Num

'------------------------------------------------------------------------------
' COMPUTE PROBABILITY
'------------------------------------------------------------------------------
    'Certain success gives a cumulative probability of one
        If ProbSuccess >= 1# Then
            Value = 1#
        Else
    'CDF = 1 - (1-p)^(k+1) = -Expm1((k+1) * Log1p(-p))
            Value = -PROB_Expm1((K + 1#) * PROB_Log1p(-ProbSuccess))
        End If

        K_STATS_Geometric_Cumulative = Value

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
        K_STATS_Geometric_Cumulative = CVErr(xlErrNum)
    'Exit before the error handler
        Exit Function

'------------------------------------------------------------------------------
' ERROR HANDLER
'------------------------------------------------------------------------------
Err_Handler:
    'Write unexpected runtime errors to diagnostics
        PROB_SetStatus Status, "Unexpected error in K_STATS_Geometric_Cumulative: " & Err.Description
    'Return worksheet value error
        K_STATS_Geometric_Cumulative = CVErr(xlErrValue)
End Function


'
'==============================================================================
' K_STATS_Geometric_Survival
'------------------------------------------------------------------------------
' PURPOSE
'   Returns the Geometric right-tail probability P(X > NumberFailures).
'
' INPUTS
'   NumberFailuresFailure count k; truncated to a non-negative integer.
'   ProbSuccess Success probability p in (0, 1].
'   Status      Optional ByRef diagnostic message.
'
' RETURNS
'   Variant
'     Success => Double probability in [0, 1].
'     Failure => CVErr(xlErrNum) or CVErr(xlErrValue).
'
' BEHAVIOR
'   - SF = (1-p)^(k+1), computed directly for right-tail accuracy.
'   - p = 1 returns zero.
'
' ERROR POLICY
'   - Invalid parameters return #NUM!.
'   - Unexpected runtime errors return #VALUE!.
'
' DEPENDENCIES
'   - PROB_DS_ValidateCount
'   - PROB_DS_ValidateProbHalfOpen
'   - PROB_TryExp
'   - PROB_Log1p
'   - PROB_SetStatus
'
' CALLED FROM
'   - Worksheet formulas
'   - M_STATS_PROBDIST_TEST
'
' UPDATED
'   2026-07-19
'==============================================================================
'
Public Function K_STATS_Geometric_Survival( _
    ByVal NumberFailures As Double, _
    ByVal ProbSuccess As Double, _
    Optional ByRef Status As String = "") _
    As Variant
'
'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Dim K                   As Double          'Failures as a truncated integer
    Dim Value               As Double          'Computed probability
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
    'Validate and truncate the failure count
        If Not PROB_DS_ValidateCount(NumberFailures, K, FailMsg, "NumberFailures") Then GoTo Fail_Num
    'Validate the success probability over (0, 1]
        If Not PROB_DS_ValidateProbHalfOpen(ProbSuccess, FailMsg) Then GoTo Fail_Num

'------------------------------------------------------------------------------
' COMPUTE PROBABILITY
'------------------------------------------------------------------------------
    'Certain success gives a survival probability of zero
        If ProbSuccess >= 1# Then
            Value = 0#
        Else
    'SF = (1-p)^(k+1) = Exp((k+1) * Log1p(-p)); underflow is a valid zero
            If Not PROB_TryExp((K + 1#) * PROB_Log1p(-ProbSuccess), Value) Then
                FailMsg = "Geometric survival overflowed a Double"
                GoTo Fail_Num
            End If
        End If

        K_STATS_Geometric_Survival = Value

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
        K_STATS_Geometric_Survival = CVErr(xlErrNum)
    'Exit before the error handler
        Exit Function

'------------------------------------------------------------------------------
' ERROR HANDLER
'------------------------------------------------------------------------------
Err_Handler:
    'Write unexpected runtime errors to diagnostics
        PROB_SetStatus Status, "Unexpected error in K_STATS_Geometric_Survival: " & Err.Description
    'Return worksheet value error
        K_STATS_Geometric_Survival = CVErr(xlErrValue)
End Function


'
'==============================================================================
' K_STATS_Geometric_InverseCumulative
'------------------------------------------------------------------------------
' PURPOSE
'   Returns the smallest NumberFailures whose cumulative probability is at
'   least Probability.
'
' WHY THIS EXISTS
'   The Geometric quantile has the closed form ceil(Log(1-P)/Log(1-p)) - 1,
'   seeded from that expression and corrected by at most a step or two to
'   absorb floating-point boundary error.
'
' INPUTS
'   Probability Target cumulative probability in the open (0, 1).
'   ProbSuccess Success probability p in (0, 1].
'   Status      Optional ByRef diagnostic message.
'
' RETURNS
'   Variant
'     Success => Double integer quantile.
'     Failure => CVErr(xlErrNum) or CVErr(xlErrValue).
'
' BEHAVIOR
'   - Returns the least k with P(X <= k) >= Probability.
'   - p = 1 returns zero.
'
' ERROR POLICY
'   - Probability outside (0, 1) or invalid p return #NUM!.
'   - Unexpected runtime errors return #VALUE!.
'
' DEPENDENCIES
'   - PROB_IsValidProbabilityOpen
'   - PROB_DS_ValidateProbHalfOpen
'   - PROB_DS_TryGeometricInverse
'   - PROB_SetStatus
'
' CALLED FROM
'   - Worksheet formulas
'   - M_STATS_PROBDIST_TEST
'
' UPDATED
'   2026-07-19
'==============================================================================
'
Public Function K_STATS_Geometric_InverseCumulative( _
    ByVal Probability As Double, _
    ByVal ProbSuccess As Double, _
    Optional ByRef Status As String = "") _
    As Variant
'
'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Dim Quantile            As Double          'Computed integer quantile
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
    'Validate the probability domain
        If Not PROB_IsValidProbabilityOpen(Probability) Then
            FailMsg = "Probability must be strictly between 0 and 1"
            GoTo Fail_Num
        End If
    'Validate the success probability over (0, 1]
        If Not PROB_DS_ValidateProbHalfOpen(ProbSuccess, FailMsg) Then GoTo Fail_Num

'------------------------------------------------------------------------------
' COMPUTE QUANTILE
'------------------------------------------------------------------------------
    'Certain success puts the entire mass at k = 0
        If ProbSuccess >= 1# Then
            K_STATS_Geometric_InverseCumulative = 0#
            GoTo Return_Success
        End If

    'Closed-form seed with a short monotone correction
        If Not PROB_DS_TryGeometricInverse( _
            Probability, ProbSuccess, Quantile, FailMsg) Then GoTo Fail_Num

        K_STATS_Geometric_InverseCumulative = Quantile

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
        K_STATS_Geometric_InverseCumulative = CVErr(xlErrNum)
    'Exit before the error handler
        Exit Function

'------------------------------------------------------------------------------
' ERROR HANDLER
'------------------------------------------------------------------------------
Err_Handler:
    'Write unexpected runtime errors to diagnostics
        PROB_SetStatus Status, "Unexpected error in K_STATS_Geometric_InverseCumulative: " & Err.Description
    'Return worksheet value error
        K_STATS_Geometric_InverseCumulative = CVErr(xlErrValue)
End Function


'
'==============================================================================
' K_STATS_Geometric_Mean
'------------------------------------------------------------------------------
' PURPOSE
'   Returns the Geometric mean, (1 - p) / p.
'
' INPUTS
'   ProbSuccess Success probability p in (0, 1].
'   Status      Optional ByRef diagnostic message.
'
' RETURNS
'   Variant
'     Success => Double value.
'     Failure => CVErr(xlErrNum) or CVErr(xlErrValue).
'
' ERROR POLICY
'   - Invalid inputs or overflow return #NUM!.
'   - Unexpected runtime errors return #VALUE!.
'
' DEPENDENCIES
'   - PROB_DS_ValidateProbHalfOpen
'   - PROB_TryDivide
'   - PROB_SetStatus
'
' CALLED FROM
'   - Worksheet formulas
'   - M_STATS_PROBDIST_TEST
'
' UPDATED
'   2026-07-19
'==============================================================================
'
Public Function K_STATS_Geometric_Mean( _
    ByVal ProbSuccess As Double, _
    Optional ByRef Status As String = "") _
    As Variant
'
'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Dim Value               As Double          'Computed mean
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
    'Validate the success probability over (0, 1]
        If Not PROB_DS_ValidateProbHalfOpen(ProbSuccess, FailMsg) Then GoTo Fail_Num

'------------------------------------------------------------------------------
' COMPUTE
'------------------------------------------------------------------------------
    'Mean = (1 - p) / p through the shared division contract
        If Not PROB_TryDivide(1# - ProbSuccess, ProbSuccess, Value) Then
            FailMsg = "Geometric mean overflows Double range"
            GoTo Fail_Num
        End If

        K_STATS_Geometric_Mean = Value

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
        K_STATS_Geometric_Mean = CVErr(xlErrNum)
    'Exit before the error handler
        Exit Function

'------------------------------------------------------------------------------
' ERROR HANDLER
'------------------------------------------------------------------------------
Err_Handler:
    'Write unexpected runtime errors to diagnostics
        PROB_SetStatus Status, "Unexpected error in K_STATS_Geometric_Mean: " & Err.Description
    'Return worksheet value error
        K_STATS_Geometric_Mean = CVErr(xlErrValue)
End Function


'
'==============================================================================
' K_STATS_Geometric_Variance
'------------------------------------------------------------------------------
' PURPOSE
'   Returns the Geometric variance, (1 - p) / p^2.
'
' INPUTS
'   ProbSuccess Success probability p in (0, 1].
'   Status      Optional ByRef diagnostic message.
'
' RETURNS
'   Variant
'     Success => Double value.
'     Failure => CVErr(xlErrNum) or CVErr(xlErrValue).
'
' ERROR POLICY
'   - Invalid inputs or overflow return #NUM!.
'   - Unexpected runtime errors return #VALUE!.
'
' DEPENDENCIES
'   - PROB_DS_ValidateProbHalfOpen
'   - PROB_TryMultiply
'   - PROB_TryDivide
'   - PROB_SetStatus
'
' CALLED FROM
'   - Worksheet formulas
'   - M_STATS_PROBDIST_TEST
'
' UPDATED
'   2026-07-19
'==============================================================================
'
Public Function K_STATS_Geometric_Variance( _
    ByVal ProbSuccess As Double, _
    Optional ByRef Status As String = "") _
    As Variant
'
'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Dim PP                  As Double          'Intermediate p squared
    Dim Value               As Double          'Computed variance
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
    'Validate the success probability over (0, 1]
        If Not PROB_DS_ValidateProbHalfOpen(ProbSuccess, FailMsg) Then GoTo Fail_Num

'------------------------------------------------------------------------------
' COMPUTE
'------------------------------------------------------------------------------
    'Variance = (1 - p) / p^2 through the shared arithmetic contracts
        If Not PROB_TryMultiply(ProbSuccess, ProbSuccess, PP) Then
            FailMsg = "Geometric variance overflows Double range"
            GoTo Fail_Num
        End If
        If Not PROB_TryDivide(1# - ProbSuccess, PP, Value) Then
            FailMsg = "Geometric variance overflows Double range"
            GoTo Fail_Num
        End If

        K_STATS_Geometric_Variance = Value

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
        K_STATS_Geometric_Variance = CVErr(xlErrNum)
    'Exit before the error handler
        Exit Function

'------------------------------------------------------------------------------
' ERROR HANDLER
'------------------------------------------------------------------------------
Err_Handler:
    'Write unexpected runtime errors to diagnostics
        PROB_SetStatus Status, "Unexpected error in K_STATS_Geometric_Variance: " & Err.Description
    'Return worksheet value error
        K_STATS_Geometric_Variance = CVErr(xlErrValue)
End Function


'
'==============================================================================
' K_STATS_Geometric_StdDev
'------------------------------------------------------------------------------
' PURPOSE
'   Returns the Geometric standard deviation, Sqr((1 - p) / p^2).
'
' INPUTS
'   ProbSuccess Success probability p in (0, 1].
'   Status      Optional ByRef diagnostic message.
'
' RETURNS
'   Variant
'     Success => Double value.
'     Failure => CVErr(xlErrNum) or CVErr(xlErrValue).
'
' ERROR POLICY
'   - Invalid inputs or overflow return #NUM!.
'   - Unexpected runtime errors return #VALUE!.
'
' DEPENDENCIES
'   - PROB_DS_ValidateProbHalfOpen
'   - PROB_TryMultiply
'   - PROB_TryDivide
'   - PROB_SetStatus
'
' CALLED FROM
'   - Worksheet formulas
'   - M_STATS_PROBDIST_TEST
'
' UPDATED
'   2026-07-19
'==============================================================================
'
Public Function K_STATS_Geometric_StdDev( _
    ByVal ProbSuccess As Double, _
    Optional ByRef Status As String = "") _
    As Variant
'
'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Dim PP                  As Double          'Intermediate p squared
    Dim Variance            As Double          'Computed variance
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
    'Validate the success probability over (0, 1]
        If Not PROB_DS_ValidateProbHalfOpen(ProbSuccess, FailMsg) Then GoTo Fail_Num

'------------------------------------------------------------------------------
' COMPUTE
'------------------------------------------------------------------------------
    'Form the variance, then take its non-negative square root
        If Not PROB_TryMultiply(ProbSuccess, ProbSuccess, PP) Then
            FailMsg = "Geometric standard deviation overflows Double range"
            GoTo Fail_Num
        End If
        If Not PROB_TryDivide(1# - ProbSuccess, PP, Variance) Then
            FailMsg = "Geometric standard deviation overflows Double range"
            GoTo Fail_Num
        End If

        K_STATS_Geometric_StdDev = Sqr(Variance)

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
        K_STATS_Geometric_StdDev = CVErr(xlErrNum)
    'Exit before the error handler
        Exit Function

'------------------------------------------------------------------------------
' ERROR HANDLER
'------------------------------------------------------------------------------
Err_Handler:
    'Write unexpected runtime errors to diagnostics
        PROB_SetStatus Status, "Unexpected error in K_STATS_Geometric_StdDev: " & Err.Description
    'Return worksheet value error
        K_STATS_Geometric_StdDev = CVErr(xlErrValue)
End Function


'==============================================================================
' PRIVATE VALIDATION AND COMPUTE KERNELS
'==============================================================================


Private Function PROB_DS_ValidateCount( _
    ByVal Raw As Double, _
    ByRef CountOut As Double, _
    ByRef FailMsg As String, _
    ByVal ArgName As String) _
    As Boolean
'
'==============================================================================
' PROB_DS_ValidateCount
'------------------------------------------------------------------------------
' PURPOSE
'   Validates a non-negative finite count and truncates it toward zero to an
'   integer, matching the Excel treatment of discrete count arguments.
'
' UPDATED
'   2026-07-19
'==============================================================================
'
    'Reject a non-finite count
        If Not PROB_IsFinite(Raw) Then
            FailMsg = ArgName & " must be a finite number"
            Exit Function
        End If
    'Reject a negative count
        If Raw < 0# Then
            FailMsg = ArgName & " must not be negative"
            Exit Function
        End If
    'Truncate toward zero (Raw >= 0, so Int matches Fix)
        CountOut = Int(Raw)
        PROB_DS_ValidateCount = True
End Function


Private Function PROB_DS_ValidateTrials( _
    ByVal Trials As Double, _
    ByRef NOut As Double, _
    ByRef FailMsg As String) _
    As Boolean
'
'==============================================================================
' PROB_DS_ValidateTrials
'------------------------------------------------------------------------------
' PURPOSE
'   Validates and truncates the Binomial trial count.
'
' UPDATED
'   2026-07-19
'==============================================================================
'
        PROB_DS_ValidateTrials = PROB_DS_ValidateCount(Trials, NOut, FailMsg, "Trials")
End Function


Private Function PROB_DS_ValidateProbClosed( _
    ByVal P As Double, _
    ByRef FailMsg As String) _
    As Boolean
'
'==============================================================================
' PROB_DS_ValidateProbClosed
'------------------------------------------------------------------------------
' PURPOSE
'   Validates a success probability over the closed interval [0, 1].
'
' UPDATED
'   2026-07-19
'==============================================================================
'
        If Not PROB_IsFinite(P) Or P < 0# Or P > 1# Then
            FailMsg = "ProbSuccess must be between 0 and 1"
            Exit Function
        End If
        PROB_DS_ValidateProbClosed = True
End Function


Private Function PROB_DS_ValidateProbHalfOpen( _
    ByVal P As Double, _
    ByRef FailMsg As String) _
    As Boolean
'
'==============================================================================
' PROB_DS_ValidateProbHalfOpen
'------------------------------------------------------------------------------
' PURPOSE
'   Validates a success probability over the half-open interval (0, 1].
'
' UPDATED
'   2026-07-19
'==============================================================================
'
        If Not PROB_IsFinite(P) Or P <= 0# Or P > 1# Then
            FailMsg = "ProbSuccess must be greater than 0 and at most 1"
            Exit Function
        End If
        PROB_DS_ValidateProbHalfOpen = True
End Function


Private Function PROB_DS_ValidateBinomialCount( _
    ByVal RawK As Double, _
    ByVal RawN As Double, _
    ByVal P As Double, _
    ByRef KOut As Double, _
    ByRef NOut As Double, _
    ByRef FailMsg As String) _
    As Boolean
'
'==============================================================================
' PROB_DS_ValidateBinomialCount
'------------------------------------------------------------------------------
' PURPOSE
'   Validates and truncates both Binomial counts, enforces 0 <= k <= n, and
'   validates the success probability.
'
' UPDATED
'   2026-07-19
'==============================================================================
'
        If Not PROB_DS_ValidateCount(RawN, NOut, FailMsg, "Trials") Then Exit Function
        If Not PROB_DS_ValidateCount(RawK, KOut, FailMsg, "NumberSuccesses") Then Exit Function
        If KOut > NOut Then
            FailMsg = "NumberSuccesses must not exceed Trials"
            Exit Function
        End If
        If Not PROB_DS_ValidateProbClosed(P, FailMsg) Then Exit Function
        PROB_DS_ValidateBinomialCount = True
End Function


Private Function PROB_DS_TryBinomialCDF( _
    ByVal K As Double, _
    ByVal N As Double, _
    ByVal P As Double, _
    ByRef Result As Double, _
    ByRef FailMsg As String) _
    As Boolean
'
'==============================================================================
' PROB_DS_TryBinomialCDF
'------------------------------------------------------------------------------
' PURPOSE
'   Computes P(X <= K) = I(1-p; N-K, K+1), handling degenerate probabilities
'   and the upper support edge exactly.
'
' UPDATED
'   2026-07-19
'==============================================================================
'
    'All mass at zero
        If P <= 0# Then
            Result = 1#
            PROB_DS_TryBinomialCDF = True
            Exit Function
        End If
    'All mass at N
        If P >= 1# Then
            If K >= N Then Result = 1# Else Result = 0#
            PROB_DS_TryBinomialCDF = True
            Exit Function
        End If
    'Cumulative reaches one at the last trial
        If K >= N Then
            Result = 1#
            PROB_DS_TryBinomialCDF = True
            Exit Function
        End If
    'Interior: regularized incomplete beta
        PROB_DS_TryBinomialCDF = _
            PROB_TryBetaRegularized(1# - P, P, N - K, K + 1#, Result, FailMsg)
End Function


Private Function PROB_DS_TryBinomialSF( _
    ByVal K As Double, _
    ByVal N As Double, _
    ByVal P As Double, _
    ByRef Result As Double, _
    ByRef FailMsg As String) _
    As Boolean
'
'==============================================================================
' PROB_DS_TryBinomialSF
'------------------------------------------------------------------------------
' PURPOSE
'   Computes P(X > K) = I(p; K+1, N-K) directly for right-tail accuracy,
'   handling degenerate probabilities and the upper support edge exactly.
'
' UPDATED
'   2026-07-19
'==============================================================================
'
    'All mass at zero
        If P <= 0# Then
            Result = 0#
            PROB_DS_TryBinomialSF = True
            Exit Function
        End If
    'All mass at N
        If P >= 1# Then
            If K < N Then Result = 1# Else Result = 0#
            PROB_DS_TryBinomialSF = True
            Exit Function
        End If
    'Survival is zero at and beyond the last trial
        If K >= N Then
            Result = 0#
            PROB_DS_TryBinomialSF = True
            Exit Function
        End If
    'Interior: regularized incomplete beta
        PROB_DS_TryBinomialSF = _
            PROB_TryBetaRegularized(P, 1# - P, K + 1#, N - K, Result, FailMsg)
End Function


Private Function PROB_DS_TryPoissonCDF( _
    ByVal K As Double, _
    ByVal Lambda As Double, _
    ByRef Result As Double, _
    ByRef FailMsg As String) _
    As Boolean
'
'==============================================================================
' PROB_DS_TryPoissonCDF
'------------------------------------------------------------------------------
' PURPOSE
'   Computes P(X <= K) = Q(K+1, lambda) via the regularized upper incomplete
'   gamma function. Lambda = 0 is a point mass at zero.
'
' UPDATED
'   2026-07-19
'==============================================================================
'
        If Lambda <= 0# Then
            Result = 1#
            PROB_DS_TryPoissonCDF = True
            Exit Function
        End If
        PROB_DS_TryPoissonCDF = _
            PROB_TryGammaRegularizedQ(K + 1#, Lambda, Result, FailMsg)
End Function


Private Function PROB_DS_TryPoissonSF( _
    ByVal K As Double, _
    ByVal Lambda As Double, _
    ByRef Result As Double, _
    ByRef FailMsg As String) _
    As Boolean
'
'==============================================================================
' PROB_DS_TryPoissonSF
'------------------------------------------------------------------------------
' PURPOSE
'   Computes P(X > K) = P(K+1, lambda) via the regularized lower incomplete
'   gamma function. Lambda = 0 is a point mass at zero.
'
' UPDATED
'   2026-07-19
'==============================================================================
'
        If Lambda <= 0# Then
            Result = 0#
            PROB_DS_TryPoissonSF = True
            Exit Function
        End If
        PROB_DS_TryPoissonSF = _
            PROB_TryGammaRegularizedP(K + 1#, Lambda, Result, FailMsg)
End Function


Private Function PROB_DS_GeometricCDF( _
    ByVal K As Double, _
    ByVal P As Double) _
    As Double
'
'==============================================================================
' PROB_DS_GeometricCDF
'------------------------------------------------------------------------------
' PURPOSE
'   Computes P(X <= K) = 1 - (1-p)^(K+1) via Expm1. The caller guarantees
'   0 < p < 1 and K >= 0.
'
' UPDATED
'   2026-07-19
'==============================================================================
'
        PROB_DS_GeometricCDF = -PROB_Expm1((K + 1#) * PROB_Log1p(-P))
End Function


Private Function PROB_DS_TryBinomialInverse( _
    ByVal Probability As Double, _
    ByVal N As Double, _
    ByVal P As Double, _
    ByRef Result As Double, _
    ByRef FailMsg As String) _
    As Boolean
'
'==============================================================================
' PROB_DS_TryBinomialInverse
'------------------------------------------------------------------------------
' PURPOSE
'   Returns the least integer k in [0, N] with P(X <= k) >= Probability by
'   lower-bound bisection over the cumulative step function.
'
' UPDATED
'   2026-07-19
'==============================================================================
'
    'Bracket is [0, N]; CDF(N) = 1 >= Probability for Probability < 1
        Dim Lo As Double, Hi As Double, MidPoint As Double, CdfMid As Double
        Lo = 0#
        Hi = N
        Do While Lo < Hi
            MidPoint = Int((Lo + Hi) / 2#)
            If Not PROB_DS_TryBinomialCDF(MidPoint, N, P, CdfMid, FailMsg) Then Exit Function
            If CdfMid >= Probability Then
                Hi = MidPoint
            Else
                Lo = MidPoint + 1#
            End If
        Loop
        Result = Lo
        PROB_DS_TryBinomialInverse = True
End Function


Private Function PROB_DS_TryPoissonInverse( _
    ByVal Probability As Double, _
    ByVal Lambda As Double, _
    ByRef Result As Double, _
    ByRef FailMsg As String) _
    As Boolean
'
'==============================================================================
' PROB_DS_TryPoissonInverse
'------------------------------------------------------------------------------
' PURPOSE
'   Returns the least integer k with P(X <= k) >= Probability. The unbounded
'   support is bracketed by an exponential search, then bisected.
'
' UPDATED
'   2026-07-19
'==============================================================================
'
        Dim Lo As Double, Hi As Double, MidPoint As Double
        Dim CdfHi As Double, CdfMid As Double
    'A zero mean is a point mass at zero
        If Lambda <= 0# Then
            Result = 0#
            PROB_DS_TryPoissonInverse = True
            Exit Function
        End If
    'Exponential search for an upper bracket
        Lo = 0#
        Hi = 1#
        Do
            If Not PROB_DS_TryPoissonCDF(Hi, Lambda, CdfHi, FailMsg) Then Exit Function
            If CdfHi >= Probability Then Exit Do
            Lo = Hi
            Hi = Hi * 2#
            If Hi > PROB_DS_INVERSE_CEILING Then
                FailMsg = "Poisson quantile exceeded the inverse search ceiling"
                Exit Function
            End If
        Loop
    'Integer bisection over [Lo, Hi]
        Do While Lo < Hi
            MidPoint = Int((Lo + Hi) / 2#)
            If Not PROB_DS_TryPoissonCDF(MidPoint, Lambda, CdfMid, FailMsg) Then Exit Function
            If CdfMid >= Probability Then
                Hi = MidPoint
            Else
                Lo = MidPoint + 1#
            End If
        Loop
        Result = Lo
        PROB_DS_TryPoissonInverse = True
End Function


Private Function PROB_DS_TryGeometricInverse( _
    ByVal Probability As Double, _
    ByVal P As Double, _
    ByRef Result As Double, _
    ByRef FailMsg As String) _
    As Boolean
'
'==============================================================================
' PROB_DS_TryGeometricInverse
'------------------------------------------------------------------------------
' PURPOSE
'   Returns the least integer k with 1-(1-p)^(k+1) >= Probability. Seeds the
'   closed form ceil(Log1p(-Probability)/Log1p(-p)) - 1 and corrects by at
'   most a step or two to absorb floating-point boundary error.
'
' UPDATED
'   2026-07-19
'==============================================================================
'
        Dim Seed As Double
    'Closed-form real solution for (k+1); VBA ceil(x) = -Int(-x)
        Seed = PROB_Log1p(-Probability) / PROB_Log1p(-P)
        Seed = -Int(-Seed) - 1#
        If Seed < 0# Then Seed = 0#
    'Correct downward while a smaller k still qualifies
        Do While Seed > 0#
            If PROB_DS_GeometricCDF(Seed - 1#, P) >= Probability Then
                Seed = Seed - 1#
            Else
                Exit Do
            End If
        Loop
    'Correct upward until k qualifies
        Do While PROB_DS_GeometricCDF(Seed, P) < Probability
            Seed = Seed + 1#
        Loop
        Result = Seed
        PROB_DS_TryGeometricInverse = True
End Function


