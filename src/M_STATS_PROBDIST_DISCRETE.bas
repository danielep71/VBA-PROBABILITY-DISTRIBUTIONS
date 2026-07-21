Attribute VB_Name = "M_STATS_PROBDIST_DISCRETE"
Option Explicit

'==============================================================================
' M_STATS_PROBDIST_DISCRETE
'------------------------------------------------------------------------------
' PURPOSE
'   Provides worksheet-facing probability functions for the Binomial, Poisson
'   and Geometric discrete distributions.
'
' WHY
'   This module supplies the first production-hardened discrete layer of the
'   probability-distribution library. It keeps worksheet wrappers separate from
'   validated private kernels, computes both tails directly, uses Loader-style
'   deviance/Stirling arrangements for central mass accuracy, and places hard
'   limits around the parameter ranges supported by the current incomplete-beta
'   and incomplete-gamma implementations.
'
' PUBLIC API
'   Binomial
'     K_STATS_Binomial_PMF
'     K_STATS_Binomial_LogPMF
'     K_STATS_Binomial_Cumulative
'     K_STATS_Binomial_Survival
'     K_STATS_Binomial_InverseCumulative
'     K_STATS_Binomial_Mean
'     K_STATS_Binomial_Variance
'     K_STATS_Binomial_StdDev
'
'   Poisson
'     K_STATS_Poisson_PMF
'     K_STATS_Poisson_LogPMF
'     K_STATS_Poisson_Cumulative
'     K_STATS_Poisson_Survival
'     K_STATS_Poisson_InverseCumulative
'     K_STATS_Poisson_Mean
'     K_STATS_Poisson_Variance
'     K_STATS_Poisson_StdDev
'
'   Geometric (failures before the first success)
'     K_STATS_Geometric_PMF
'     K_STATS_Geometric_LogPMF
'     K_STATS_Geometric_Cumulative
'     K_STATS_Geometric_Survival
'     K_STATS_Geometric_InverseCumulative
'     K_STATS_Geometric_Mean
'     K_STATS_Geometric_Variance
'     K_STATS_Geometric_StdDev
'
'   NegativeBinomial
'     K_STATS_NegativeBinomial_PMF
'     K_STATS_NegativeBinomial_LogPMF
'     K_STATS_NegativeBinomial_Cumulative
'     K_STATS_NegativeBinomial_Survival
'     K_STATS_NegativeBinomial_InverseCumulative
'     K_STATS_NegativeBinomial_Mean
'     K_STATS_NegativeBinomial_Variance
'     K_STATS_NegativeBinomial_StdDev
'
'   Hypergeometric
'     K_STATS_Hypergeometric_PMF
'     K_STATS_Hypergeometric_LogPMF
'     K_STATS_Hypergeometric_Cumulative
'     K_STATS_Hypergeometric_Survival
'     K_STATS_Hypergeometric_InverseCumulative
'     K_STATS_Hypergeometric_Mean
'     K_STATS_Hypergeometric_Variance
'     K_STATS_Hypergeometric_StdDev
'
' PARAMETERIZATION
'   Binomial(NumberSuccesses, Trials, ProbSuccess)
'     - PMF/CDF/SF and moments accept ProbSuccess in [0, 1].
'     - InverseCumulative follows BINOM.INV and requires ProbSuccess in (0, 1).
'
'   Poisson(NumberEvents, Mean)
'     - Mean is the Poisson intensity lambda and must be non-negative.
'
'   Geometric(NumberFailures, ProbSuccess)
'     - NumberFailures counts failures before the first success.
'     - Support is k = 0, 1, 2, ...
'     - PMF = p * (1-p)^k and ProbSuccess is in (0, 1].
'
' COUNT POLICY
'   - Worksheet count inputs are truncated toward zero before validation.
'   - Every stored count is limited to the largest consecutively representable
'     integer in IEEE-754 Double: 2^53 - 1.
'   - Kernel-backed CDF/SF/inverse functions apply tighter limits aligned to the
'     iteration budgets in M_STATS_PROBDIST_SPECIALFUNCS.
'
' SUPPORTED NUMERICAL DOMAIN
'   - Binomial PMF and moments:
'       Trials <= 2^53 - 1.
'   - Binomial CDF, SF and inverse:
'       Trials <= 10,000,000.
'   - Poisson PMF:
'       NumberEvents and Mean <= 2^53 - 1.
'   - Poisson CDF and SF:
'       NumberEvents <= 20,000,000 and Mean <= 10,000,000.
'   - Poisson inverse:
'       Mean <= 10,000,000; the searched quantile is capped at
'       20,000,000.
'   - Geometric counts and returned quantiles:
'       <= 2^53 - 1.
'
' NUMERICAL METHODS
'   - Binomial PMF:
'       Catherine Loader's Stirling-error/deviance arrangement.
'   - Poisson PMF:
'       Catherine Loader's Stirling-error/deviance arrangement.
'   - Binomial CDF/SF:
'       Direct regularized incomplete-beta identities.
'   - Poisson CDF/SF:
'       Direct regularized incomplete-gamma identities.
'   - Binomial/Poisson inverses:
'       Integer lower-bound searches driven by the smaller tail.
'   - Geometric CDF/SF:
'       Log1p/Expm1 closed forms with guarded exponent products.
'
' ERROR POLICY
'   - Invalid domains, unsupported magnitudes, arithmetic overflow and kernel
'     non-convergence return CVErr(xlErrNum).
'   - Unexpected runtime errors return CVErr(xlErrValue).
'   - Mathematically valid exponential underflow returns zero.
'   - Detailed diagnostics are written to the optional ByRef Status argument.
'   - No MsgBox is raised.
'
'
' LOG-SPACE MASS
'   The _LogPMF functions return the natural log of the mass. They stay finite
'   and accurate where the ordinary _PMF underflows a Double to zero (for
'   example ln P well below -700), and return #NUM! only when the outcome has
'   probability exactly zero.
'
' NEGATIVE BINOMIAL AND HYPERGEOMETRIC
'   Negative binomial counts failures before the r-th success (r = 1 is the
'   Geometric); its mass reuses the Loader Binomial kernel and its CDF/SF use
'   the two-argument regularized incomplete beta. Hypergeometric mass is three
'   shared Binomial log-masses (p = n/N); its CDF/SF sum the near tail by an
'   exact successive-ratio recurrence, so only a few standard deviations of
'   terms are evaluated. Cumulative and inverse counts are capped for the
'   summation and bracketing paths; the single-point masses accept the full
'   exact-integer domain.
' DEPENDENCIES
'   - M_STATS_PROBDIST_CORE
'       PROB_HALF_LOG_TWO_PI
'       PROB_MACH_EPS
'       PROB_IsFinite
'       PROB_IsValidProbabilityOpen
'       PROB_TryAdd
'       PROB_TryMultiply
'       PROB_TryDivide
'       PROB_TryExp
'       PROB_Log1p
'       PROB_Expm1
'       PROB_NormalInvCDFRaw
'       PROB_SetStatus
'
'   - M_STATS_PROBDIST_SPECIALFUNCS
'       PROB_StirlingError
'       PROB_TryBetaRegularized
'       PROB_TryGammaRegularizedP
'       PROB_TryGammaRegularizedQ
'
' NOTES
'   - This module is complete for the stated Binomial, Poisson and Geometric
'     surface. Negative Binomial, Hypergeometric and Discrete Uniform remain a
'     separate future batch.
'   - Direct survival functions should be used for small right-tail
'     probabilities; subtracting the CDF from one loses those tails.
'
' UPDATED
'   2026-07-19 - Full production-hardening rewrite.
'==============================================================================


'==============================================================================
' PRIVATE CONSTANTS
'==============================================================================

'Largest integer for which every consecutive integer is exactly representable
'in IEEE-754 Double.
Private Const PROB_DS_MAX_EXACT_INTEGER As Double = 9.00719925474099E+15

'Maximum Binomial trial count passed to the current incomplete-beta kernel.
Private Const PROB_DS_MAX_BINOMIAL_KERNEL_N As Double = 10000000#

'Maximum Poisson mean passed to the current incomplete-gamma kernel.
Private Const PROB_DS_MAX_POISSON_KERNEL_MEAN As Double = 10000000#

'Maximum Poisson count/quantile passed to the current incomplete-gamma kernel.
Private Const PROB_DS_MAX_POISSON_KERNEL_COUNT As Double = 20000000#

'Iteration guards for finite integer searches and Loader's deviance series.
Private Const PROB_DS_MAX_INVERSE_ITER As Long = 128
Private Const PROB_DS_MAX_BRACKET_ITER As Long = 64
Private Const PROB_DS_MAX_GEOMETRIC_CORRECTIONS As Long = 8
Private Const PROB_DS_BD0_MAX_ITER As Long = 1000
Private Const PROB_DS_MAX_NEGBINOM_KERNEL_COUNT As Double = 20000000#
Private Const PROB_DS_MAX_HYPERGEOMETRIC_POP As Double = 100000000#
Private Const PROB_DS_MAX_HYPERGEOMETRIC_SUM_ITER As Long = 200000


'==============================================================================
' K_STATS_Binomial_PMF
'------------------------------------------------------------------------------
' PURPOSE
'   Returns the Binomial probability mass P(X = NumberSuccesses).
'
' WHY
'   The Loader arrangement avoids cancellation near the mode when Trials is large.
'
' WORKSHEET EQUIVALENT
'   BINOM.DIST(NumberSuccesses, Trials, ProbSuccess, FALSE)
'
' INPUTS
'   NumberSuccesses  Success count k; truncated toward zero.
'   Trials           Number of trials n; truncated toward zero.
'   ProbSuccess      Success probability p in [0, 1].
'   Status           Optional ByRef diagnostic message.
'
' RETURNS
'   Variant
'     Success => Double probability in [0, 1].
'     Failure => CVErr(xlErrNum) or CVErr(xlErrValue).
'
' BEHAVIOR
'   Supports exact integer counts through 2^53 - 1.
'   Handles p = 0 and p = 1 exactly.
'   Deep-tail exponential underflow is a valid zero.
'
' ERROR POLICY
'   Invalid domains, unsupported magnitudes or numerical failure return #NUM!.
'   Unexpected runtime errors return #VALUE!.
'   Diagnostics are written to Status.
'
' DEPENDENCIES
'   - PROB_DS_ValidateBinomialMassInputs
'   - PROB_DS_TryBinomialPMF
'   - PROB_SetStatus
'
' CALLED FROM
'   - Worksheet formulas
'   - M_STATS_PROBDIST_TEST
'
' UPDATED
'   2026-07-19 - Full production-hardening rewrite.
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
    Dim n                   As Double          'Trials as a truncated integer
    Dim Value               As Double          'Computed probability mass
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
    'Validate counts and probability for the Loader mass kernel
        If Not PROB_DS_ValidateBinomialMassInputs( _
            NumberSuccesses, Trials, ProbSuccess, K, n, FailMsg) Then GoTo Fail_Num

'------------------------------------------------------------------------------
' COMPUTE
'------------------------------------------------------------------------------
    'Compute the mass through Loader's stable deviance arrangement
        If Not PROB_DS_TryBinomialPMF( _
            K, n, ProbSuccess, Value, FailMsg) Then GoTo Fail_Num

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
        PROB_SetStatus Status, _
            "Unexpected error in K_STATS_Binomial_PMF: " & Err.Description
    'Return worksheet value error
        K_STATS_Binomial_PMF = CVErr(xlErrValue)
End Function


Public Function K_STATS_Binomial_LogPMF( _
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
    Dim n                   As Double          'Trials as a truncated integer
    Dim LogMass             As Double          'Natural logarithm of the mass
    Dim IsCertainZero       As Boolean         'Outcome is impossible or below Double range
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
    'Validate counts and probability for the Loader mass kernel
        If Not PROB_DS_ValidateBinomialMassInputs( _
            NumberSuccesses, Trials, ProbSuccess, K, n, FailMsg) Then GoTo Fail_Num

'------------------------------------------------------------------------------
' COMPUTE
'------------------------------------------------------------------------------
    'Assemble the log mass through Loader's stable deviance arrangement
        If Not PROB_DS_TryBinomialLogMass( _
            K, n, ProbSuccess, LogMass, IsCertainZero, FailMsg) Then GoTo Fail_Num
    'A zero-probability outcome has no defined log-mass
        If IsCertainZero Then
            FailMsg = "Log-mass is undefined: the outcome has probability zero"
            GoTo Fail_Num
        End If

        K_STATS_Binomial_LogPMF = LogMass

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
        K_STATS_Binomial_LogPMF = CVErr(xlErrNum)
    'Exit before the error handler
        Exit Function

'------------------------------------------------------------------------------
' ERROR HANDLER
'------------------------------------------------------------------------------
Err_Handler:
    'Write unexpected runtime errors to diagnostics
        PROB_SetStatus Status, _
            "Unexpected error in K_STATS_Binomial_LogPMF: " & Err.Description
    'Return worksheet value error
        K_STATS_Binomial_LogPMF = CVErr(xlErrValue)
End Function

'==============================================================================
' K_STATS_Binomial_Cumulative
'------------------------------------------------------------------------------
' PURPOSE
'   Returns the Binomial left-tail probability P(X <= NumberSuccesses).
'
' WHY
'   The probability is evaluated directly from the regularized incomplete beta.
'
' WORKSHEET EQUIVALENT
'   BINOM.DIST(NumberSuccesses, Trials, ProbSuccess, TRUE)
'
' INPUTS
'   NumberSuccesses  Success count k; truncated toward zero.
'   Trials           Number of trials n; truncated toward zero.
'   ProbSuccess      Success probability p in [0, 1].
'   Status           Optional ByRef diagnostic message.
'
' RETURNS
'   Variant
'     Success => Double probability in [0, 1].
'     Failure => CVErr(xlErrNum) or CVErr(xlErrValue).
'
' BEHAVIOR
'   Uses P(X<=k) = I(1-p; n-k, k+1).
'   Applies the tested incomplete-beta limit Trials <= 10,000,000.
'   Handles degenerate probabilities and support boundaries exactly.
'
' ERROR POLICY
'   Invalid domains, unsupported magnitudes or numerical failure return #NUM!.
'   Unexpected runtime errors return #VALUE!.
'   Diagnostics are written to Status.
'
' DEPENDENCIES
'   - PROB_DS_ValidateBinomialKernelInputs
'   - PROB_DS_TryBinomialCDF
'   - PROB_SetStatus
'
' CALLED FROM
'   - Worksheet formulas
'   - M_STATS_PROBDIST_TEST
'
' UPDATED
'   2026-07-19 - Full production-hardening rewrite.
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
    Dim n                   As Double          'Trials as a truncated integer
    Dim Value               As Double          'Computed cumulative probability
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
    'Validate against the supported incomplete-beta domain
        If Not PROB_DS_ValidateBinomialKernelInputs( _
            NumberSuccesses, Trials, ProbSuccess, K, n, FailMsg) Then GoTo Fail_Num

'------------------------------------------------------------------------------
' COMPUTE
'------------------------------------------------------------------------------
    'Evaluate the left tail directly
        If Not PROB_DS_TryBinomialCDF( _
            K, n, ProbSuccess, Value, FailMsg) Then GoTo Fail_Num

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
        PROB_SetStatus Status, _
            "Unexpected error in K_STATS_Binomial_Cumulative: " & Err.Description
    'Return worksheet value error
        K_STATS_Binomial_Cumulative = CVErr(xlErrValue)
End Function

'==============================================================================
' K_STATS_Binomial_Survival
'------------------------------------------------------------------------------
' PURPOSE
'   Returns the Binomial right-tail probability P(X > NumberSuccesses).
'
' WHY
'   A direct upper-tail identity avoids the precision loss of 1 - CDF.
'
' WORKSHEET EQUIVALENT
'   No single Excel function; numerically preferable to
'   1 - BINOM.DIST(NumberSuccesses, Trials, ProbSuccess, TRUE).
'
' INPUTS
'   NumberSuccesses  Success count k; truncated toward zero.
'   Trials           Number of trials n; truncated toward zero.
'   ProbSuccess      Success probability p in [0, 1].
'   Status           Optional ByRef diagnostic message.
'
' RETURNS
'   Variant
'     Success => Double probability in [0, 1].
'     Failure => CVErr(xlErrNum) or CVErr(xlErrValue).
'
' BEHAVIOR
'   Uses P(X>k) = I(p; k+1, n-k).
'   Applies the tested incomplete-beta limit Trials <= 10,000,000.
'   Handles degenerate probabilities and support boundaries exactly.
'
' ERROR POLICY
'   Invalid domains, unsupported magnitudes or numerical failure return #NUM!.
'   Unexpected runtime errors return #VALUE!.
'   Diagnostics are written to Status.
'
' DEPENDENCIES
'   - PROB_DS_ValidateBinomialKernelInputs
'   - PROB_DS_TryBinomialSF
'   - PROB_SetStatus
'
' CALLED FROM
'   - Worksheet formulas
'   - M_STATS_PROBDIST_TEST
'
' UPDATED
'   2026-07-19 - Full production-hardening rewrite.
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
    Dim n                   As Double          'Trials as a truncated integer
    Dim Value               As Double          'Computed survival probability
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
    'Validate against the supported incomplete-beta domain
        If Not PROB_DS_ValidateBinomialKernelInputs( _
            NumberSuccesses, Trials, ProbSuccess, K, n, FailMsg) Then GoTo Fail_Num

'------------------------------------------------------------------------------
' COMPUTE
'------------------------------------------------------------------------------
    'Evaluate the right tail directly
        If Not PROB_DS_TryBinomialSF( _
            K, n, ProbSuccess, Value, FailMsg) Then GoTo Fail_Num

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
        PROB_SetStatus Status, _
            "Unexpected error in K_STATS_Binomial_Survival: " & Err.Description
    'Return worksheet value error
        K_STATS_Binomial_Survival = CVErr(xlErrValue)
End Function

'==============================================================================
' K_STATS_Binomial_InverseCumulative
'------------------------------------------------------------------------------
' PURPOSE
'   Returns the least integer k satisfying P(X <= k) >= Probability.
'
' WHY
'   The lower-bound bisection is driven by the smaller of the CDF and SF tails.
'
' WORKSHEET EQUIVALENT
'   BINOM.INV(Trials, ProbSuccess, Probability)
'
' INPUTS
'   Probability      Target cumulative probability in (0, 1).
'   Trials           Number of trials n; truncated toward zero.
'   ProbSuccess      Success probability p in (0, 1), matching BINOM.INV.
'   Status           Optional ByRef diagnostic message.
'
' RETURNS
'   Variant
'     Success => Double containing an exactly representable integer quantile.
'     Failure => CVErr(xlErrNum) or CVErr(xlErrValue).
'
' BEHAVIOR
'   Uses an overflow-safe integer midpoint and a finite iteration guard.
'   Uses the direct SF when Probability > 0.5.
'   Applies the tested incomplete-beta limit Trials <= 10,000,000.
'
' ERROR POLICY
'   Invalid domains, unsupported magnitudes or numerical failure return #NUM!.
'   Unexpected runtime errors return #VALUE!.
'   Diagnostics are written to Status.
'
' DEPENDENCIES
'   - PROB_IsValidProbabilityOpen
'   - PROB_DS_ValidateTrialsKernel
'   - PROB_DS_ValidateProbOpen
'   - PROB_DS_TryBinomialInverse
'   - PROB_SetStatus
'
' CALLED FROM
'   - Worksheet formulas
'   - M_STATS_PROBDIST_TEST
'
' UPDATED
'   2026-07-19 - Full production-hardening rewrite.
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
    Dim n                   As Double          'Trials as a truncated integer
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
    'Validate the target cumulative probability
        If Not PROB_IsValidProbabilityOpen(Probability) Then
            FailMsg = "Probability must be strictly between 0 and 1"
            GoTo Fail_Num
        End If

    'Validate Trials against the incomplete-beta domain
        If Not PROB_DS_ValidateTrialsKernel(Trials, n, FailMsg) Then GoTo Fail_Num

    'Match the open success-probability contract of BINOM.INV
        If Not PROB_DS_ValidateProbOpen(ProbSuccess, FailMsg) Then GoTo Fail_Num

'------------------------------------------------------------------------------
' COMPUTE
'------------------------------------------------------------------------------
    'Find the least qualifying integer using the smaller tail
        If Not PROB_DS_TryBinomialInverse( _
            Probability, n, ProbSuccess, Quantile, FailMsg) Then GoTo Fail_Num

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
        PROB_SetStatus Status, _
            "Unexpected error in K_STATS_Binomial_InverseCumulative: " & Err.Description
    'Return worksheet value error
        K_STATS_Binomial_InverseCumulative = CVErr(xlErrValue)
End Function

'==============================================================================
' K_STATS_Binomial_Mean
'------------------------------------------------------------------------------
' PURPOSE
'   Returns the Binomial mean n * p.
'
' WHY
'   The moment is evaluated directly after exact-count validation.
'
' INPUTS
'   Trials       Number of trials n; truncated toward zero.
'   ProbSuccess  Success probability p in [0, 1].
'   Status       Optional ByRef diagnostic message.
'
' RETURNS
'   Variant
'     Success => Double value.
'     Failure => CVErr(xlErrNum) or CVErr(xlErrValue).
'
' BEHAVIOR
'   Uses the closed-form expression n * p.
'   Supports Trials through 2^53 - 1.
'
' ERROR POLICY
'   Invalid domains, unsupported magnitudes or numerical failure return #NUM!.
'   Unexpected runtime errors return #VALUE!.
'   Diagnostics are written to Status.
'
' DEPENDENCIES
'   - PROB_DS_ValidateTrialsExact
'   - PROB_DS_ValidateProbClosed
'   - PROB_TryMultiply
'   - PROB_SetStatus
'
' CALLED FROM
'   - Worksheet formulas
'   - M_STATS_PROBDIST_TEST
'
' UPDATED
'   2026-07-19 - Full production-hardening rewrite.
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
    Dim n                   As Double          'Trials as a truncated integer
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
    'Validate Trials in the exact-integer count domain
        If Not PROB_DS_ValidateTrialsExact(Trials, n, FailMsg) Then GoTo Fail_Num

    'Validate the closed success-probability domain
        If Not PROB_DS_ValidateProbClosed(ProbSuccess, FailMsg) Then GoTo Fail_Num

'------------------------------------------------------------------------------
' COMPUTE
'------------------------------------------------------------------------------
    'Form n * p through the shared overflow contract
        If Not PROB_TryMultiply(n, ProbSuccess, Value) Then
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
        PROB_SetStatus Status, _
            "Unexpected error in K_STATS_Binomial_Mean: " & Err.Description
    'Return worksheet value error
        K_STATS_Binomial_Mean = CVErr(xlErrValue)
End Function

'==============================================================================
' K_STATS_Binomial_Variance
'------------------------------------------------------------------------------
' PURPOSE
'   Returns the Binomial variance n * p * (1-p).
'
' WHY
'   The moment is evaluated directly after exact-count validation.
'
' INPUTS
'   Trials       Number of trials n; truncated toward zero.
'   ProbSuccess  Success probability p in [0, 1].
'   Status       Optional ByRef diagnostic message.
'
' RETURNS
'   Variant
'     Success => Double value.
'     Failure => CVErr(xlErrNum) or CVErr(xlErrValue).
'
' BEHAVIOR
'   Uses the closed-form expression n * p * (1-p).
'   Supports Trials through 2^53 - 1.
'
' ERROR POLICY
'   Invalid domains, unsupported magnitudes or numerical failure return #NUM!.
'   Unexpected runtime errors return #VALUE!.
'   Diagnostics are written to Status.
'
' DEPENDENCIES
'   - PROB_DS_ValidateTrialsExact
'   - PROB_DS_ValidateProbClosed
'   - PROB_TryMultiply
'   - PROB_SetStatus
'
' CALLED FROM
'   - Worksheet formulas
'   - M_STATS_PROBDIST_TEST
'
' UPDATED
'   2026-07-19 - Full production-hardening rewrite.
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
    Dim n                   As Double          'Trials as a truncated integer
    Dim Np                  As Double          'Intermediate n * p
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
    'Validate Trials in the exact-integer count domain
        If Not PROB_DS_ValidateTrialsExact(Trials, n, FailMsg) Then GoTo Fail_Num

    'Validate the closed success-probability domain
        If Not PROB_DS_ValidateProbClosed(ProbSuccess, FailMsg) Then GoTo Fail_Num

'------------------------------------------------------------------------------
' COMPUTE
'------------------------------------------------------------------------------
    'Form n * p, then apply the complementary probability
        If Not PROB_TryMultiply(n, ProbSuccess, Np) Then
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
        PROB_SetStatus Status, _
            "Unexpected error in K_STATS_Binomial_Variance: " & Err.Description
    'Return worksheet value error
        K_STATS_Binomial_Variance = CVErr(xlErrValue)
End Function

'==============================================================================
' K_STATS_Binomial_StdDev
'------------------------------------------------------------------------------
' PURPOSE
'   Returns the Binomial standard deviation Sqr(n * p * (1-p)).
'
' WHY
'   The moment is evaluated directly after exact-count validation.
'
' INPUTS
'   Trials       Number of trials n; truncated toward zero.
'   ProbSuccess  Success probability p in [0, 1].
'   Status       Optional ByRef diagnostic message.
'
' RETURNS
'   Variant
'     Success => Double value.
'     Failure => CVErr(xlErrNum) or CVErr(xlErrValue).
'
' BEHAVIOR
'   Uses the closed-form expression Sqr(n * p * (1-p)).
'   Supports Trials through 2^53 - 1.
'
' ERROR POLICY
'   Invalid domains, unsupported magnitudes or numerical failure return #NUM!.
'   Unexpected runtime errors return #VALUE!.
'   Diagnostics are written to Status.
'
' DEPENDENCIES
'   - PROB_DS_ValidateTrialsExact
'   - PROB_DS_ValidateProbClosed
'   - PROB_TryMultiply
'   - PROB_SetStatus
'
' CALLED FROM
'   - Worksheet formulas
'   - M_STATS_PROBDIST_TEST
'
' UPDATED
'   2026-07-19 - Full production-hardening rewrite.
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
    Dim n                   As Double          'Trials as a truncated integer
    Dim Np                  As Double          'Intermediate n * p
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
    'Validate Trials in the exact-integer count domain
        If Not PROB_DS_ValidateTrialsExact(Trials, n, FailMsg) Then GoTo Fail_Num

    'Validate the closed success-probability domain
        If Not PROB_DS_ValidateProbClosed(ProbSuccess, FailMsg) Then GoTo Fail_Num

'------------------------------------------------------------------------------
' COMPUTE
'------------------------------------------------------------------------------
    'Form the variance through guarded multiplication
        If Not PROB_TryMultiply(n, ProbSuccess, Np) Then
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
        PROB_SetStatus Status, _
            "Unexpected error in K_STATS_Binomial_StdDev: " & Err.Description
    'Return worksheet value error
        K_STATS_Binomial_StdDev = CVErr(xlErrValue)
End Function

'==============================================================================
' K_STATS_Poisson_PMF
'------------------------------------------------------------------------------
' PURPOSE
'   Returns the Poisson probability mass P(X = NumberEvents).
'
' WHY
'   Loader's deviance arrangement avoids the large cancellation in
'   k*Log(lambda) - lambda - LogGamma(k+1) near k = lambda.
'
' WORKSHEET EQUIVALENT
'   POISSON.DIST(NumberEvents, Mean, FALSE)
'
' INPUTS
'   NumberEvents  Event count k; truncated toward zero.
'   Mean          Poisson intensity lambda >= 0.
'   Status        Optional ByRef diagnostic message.
'
' RETURNS
'   Variant
'     Success => Double probability in [0, 1].
'     Failure => CVErr(xlErrNum) or CVErr(xlErrValue).
'
' BEHAVIOR
'   Supports NumberEvents and Mean through 2^53 - 1.
'   Mean = 0 is handled as a point mass at zero.
'   Deep-tail exponential underflow is a valid zero.
'
' ERROR POLICY
'   Invalid domains, unsupported magnitudes or numerical failure return #NUM!.
'   Unexpected runtime errors return #VALUE!.
'   Diagnostics are written to Status.
'
' DEPENDENCIES
'   - PROB_DS_ValidatePoissonPMFInputs
'   - PROB_DS_TryPoissonPMF
'   - PROB_SetStatus
'
' CALLED FROM
'   - Worksheet formulas
'   - M_STATS_PROBDIST_TEST
'
' UPDATED
'   2026-07-19 - Full production-hardening rewrite.
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
    Dim Value               As Double          'Computed probability mass
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
    'Validate the Loader mass domain
        If Not PROB_DS_ValidatePoissonPMFInputs( _
            NumberEvents, Mean, K, FailMsg) Then GoTo Fail_Num

'------------------------------------------------------------------------------
' COMPUTE
'------------------------------------------------------------------------------
    'Compute the mass through Loader's stable deviance arrangement
        If Not PROB_DS_TryPoissonPMF( _
            K, Mean, Value, FailMsg) Then GoTo Fail_Num

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
        PROB_SetStatus Status, _
            "Unexpected error in K_STATS_Poisson_PMF: " & Err.Description
    'Return worksheet value error
        K_STATS_Poisson_PMF = CVErr(xlErrValue)
End Function


Public Function K_STATS_Poisson_LogPMF( _
    ByVal NumberEvents As Double, _
    ByVal Mean As Double, _
    Optional ByRef Status As String = "") _
    As Variant
'
'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Dim K                   As Double          'Events as a truncated integer
    Dim LogMass             As Double          'Natural logarithm of the mass
    Dim IsCertainZero       As Boolean         'Outcome is impossible under a zero mean
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
    'Validate the Loader mass domain
        If Not PROB_DS_ValidatePoissonPMFInputs( _
            NumberEvents, Mean, K, FailMsg) Then GoTo Fail_Num

'------------------------------------------------------------------------------
' COMPUTE
'------------------------------------------------------------------------------
    'Assemble the log mass through Loader's stable deviance arrangement
        If Not PROB_DS_TryPoissonLogMass( _
            K, Mean, LogMass, IsCertainZero, FailMsg) Then GoTo Fail_Num
    'A zero-probability outcome has no defined log-mass
        If IsCertainZero Then
            FailMsg = "Log-mass is undefined: the outcome has probability zero"
            GoTo Fail_Num
        End If

        K_STATS_Poisson_LogPMF = LogMass

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
        K_STATS_Poisson_LogPMF = CVErr(xlErrNum)
    'Exit before the error handler
        Exit Function

'------------------------------------------------------------------------------
' ERROR HANDLER
'------------------------------------------------------------------------------
Err_Handler:
    'Write unexpected runtime errors to diagnostics
        PROB_SetStatus Status, _
            "Unexpected error in K_STATS_Poisson_LogPMF: " & Err.Description
    'Return worksheet value error
        K_STATS_Poisson_LogPMF = CVErr(xlErrValue)
End Function

'==============================================================================
' K_STATS_Poisson_Cumulative
'------------------------------------------------------------------------------
' PURPOSE
'   Returns the Poisson left-tail probability P(X <= NumberEvents).
'
' WHY
'   The probability is evaluated directly from the regularized upper incomplete gamma.
'
' WORKSHEET EQUIVALENT
'   POISSON.DIST(NumberEvents, Mean, TRUE)
'
' INPUTS
'   NumberEvents  Event count k; truncated toward zero.
'   Mean          Poisson intensity lambda >= 0.
'   Status        Optional ByRef diagnostic message.
'
' RETURNS
'   Variant
'     Success => Double probability in [0, 1].
'     Failure => CVErr(xlErrNum) or CVErr(xlErrValue).
'
' BEHAVIOR
'   Uses P(X<=k) = Q(k+1, lambda).
'   Applies the tested incomplete-gamma count and mean limits.
'   Mean = 0 returns one.
'
' ERROR POLICY
'   Invalid domains, unsupported magnitudes or numerical failure return #NUM!.
'   Unexpected runtime errors return #VALUE!.
'   Diagnostics are written to Status.
'
' DEPENDENCIES
'   - PROB_DS_ValidatePoissonKernelInputs
'   - PROB_DS_TryPoissonCDF
'   - PROB_SetStatus
'
' CALLED FROM
'   - Worksheet formulas
'   - M_STATS_PROBDIST_TEST
'
' UPDATED
'   2026-07-19 - Full production-hardening rewrite.
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
    Dim Value               As Double          'Computed cumulative probability
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
    'Validate against the supported incomplete-gamma domain
        If Not PROB_DS_ValidatePoissonKernelInputs( _
            NumberEvents, Mean, K, FailMsg) Then GoTo Fail_Num

'------------------------------------------------------------------------------
' COMPUTE
'------------------------------------------------------------------------------
    'Evaluate the left tail directly
        If Not PROB_DS_TryPoissonCDF( _
            K, Mean, Value, FailMsg) Then GoTo Fail_Num

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
        PROB_SetStatus Status, _
            "Unexpected error in K_STATS_Poisson_Cumulative: " & Err.Description
    'Return worksheet value error
        K_STATS_Poisson_Cumulative = CVErr(xlErrValue)
End Function

'==============================================================================
' K_STATS_Poisson_Survival
'------------------------------------------------------------------------------
' PURPOSE
'   Returns the Poisson right-tail probability P(X > NumberEvents).
'
' WHY
'   A direct lower incomplete-gamma identity avoids the precision loss of 1 - CDF.
'
' WORKSHEET EQUIVALENT
'   No single Excel function; numerically preferable to
'   1 - POISSON.DIST(NumberEvents, Mean, TRUE).
'
' INPUTS
'   NumberEvents  Event count k; truncated toward zero.
'   Mean          Poisson intensity lambda >= 0.
'   Status        Optional ByRef diagnostic message.
'
' RETURNS
'   Variant
'     Success => Double probability in [0, 1].
'     Failure => CVErr(xlErrNum) or CVErr(xlErrValue).
'
' BEHAVIOR
'   Uses P(X>k) = P(k+1, lambda).
'   Applies the tested incomplete-gamma count and mean limits.
'   Mean = 0 returns zero.
'
' ERROR POLICY
'   Invalid domains, unsupported magnitudes or numerical failure return #NUM!.
'   Unexpected runtime errors return #VALUE!.
'   Diagnostics are written to Status.
'
' DEPENDENCIES
'   - PROB_DS_ValidatePoissonKernelInputs
'   - PROB_DS_TryPoissonSF
'   - PROB_SetStatus
'
' CALLED FROM
'   - Worksheet formulas
'   - M_STATS_PROBDIST_TEST
'
' UPDATED
'   2026-07-19 - Full production-hardening rewrite.
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
    Dim Value               As Double          'Computed survival probability
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
    'Validate against the supported incomplete-gamma domain
        If Not PROB_DS_ValidatePoissonKernelInputs( _
            NumberEvents, Mean, K, FailMsg) Then GoTo Fail_Num

'------------------------------------------------------------------------------
' COMPUTE
'------------------------------------------------------------------------------
    'Evaluate the right tail directly
        If Not PROB_DS_TryPoissonSF( _
            K, Mean, Value, FailMsg) Then GoTo Fail_Num

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
        PROB_SetStatus Status, _
            "Unexpected error in K_STATS_Poisson_Survival: " & Err.Description
    'Return worksheet value error
        K_STATS_Poisson_Survival = CVErr(xlErrValue)
End Function

'==============================================================================
' K_STATS_Poisson_InverseCumulative
'------------------------------------------------------------------------------
' PURPOSE
'   Returns the least integer k satisfying P(X <= k) >= Probability.
'
' WHY
'   A Cornish-Fisher seed is bracketed and then refined by smaller-tail
'   integer bisection, avoiding a search that always starts at one.
'
' INPUTS
'   Probability  Target cumulative probability in (0, 1).
'   Mean         Poisson intensity lambda in the supported kernel domain.
'   Status       Optional ByRef diagnostic message.
'
' RETURNS
'   Variant
'     Success => Double containing an exactly representable integer quantile.
'     Failure => CVErr(xlErrNum) or CVErr(xlErrValue).
'
' BEHAVIOR
'   Uses the direct SF when Probability > 0.5.
'   Evaluates the configured quantile ceiling before reporting failure.
'   Every bracket and bisection loop has an explicit iteration guard.
'
' ERROR POLICY
'   Invalid domains, unsupported magnitudes or numerical failure return #NUM!.
'   Unexpected runtime errors return #VALUE!.
'   Diagnostics are written to Status.
'
' DEPENDENCIES
'   - PROB_IsValidProbabilityOpen
'   - PROB_DS_ValidatePoissonMeanKernel
'   - PROB_DS_TryPoissonInverse
'   - PROB_NormalInvCDFRaw
'   - PROB_SetStatus
'
' CALLED FROM
'   - Worksheet formulas
'   - M_STATS_PROBDIST_TEST
'
' UPDATED
'   2026-07-19 - Full production-hardening rewrite.
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
    'Validate the target cumulative probability
        If Not PROB_IsValidProbabilityOpen(Probability) Then
            FailMsg = "Probability must be strictly between 0 and 1"
            GoTo Fail_Num
        End If

    'Validate the mean against the incomplete-gamma domain
        If Not PROB_DS_ValidatePoissonMeanKernel(Mean, FailMsg) Then GoTo Fail_Num

'------------------------------------------------------------------------------
' COMPUTE
'------------------------------------------------------------------------------
    'Bracket and refine the least qualifying integer
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
        PROB_SetStatus Status, _
            "Unexpected error in K_STATS_Poisson_InverseCumulative: " & Err.Description
    'Return worksheet value error
        K_STATS_Poisson_InverseCumulative = CVErr(xlErrValue)
End Function

'==============================================================================
' K_STATS_Poisson_Mean
'------------------------------------------------------------------------------
' PURPOSE
'   Returns the Poisson mean, which equals lambda.
'
' WHY
'   The moment is exact in closed form and does not call an iterative kernel.
'
' INPUTS
'   Mean    Poisson intensity lambda >= 0.
'   Status  Optional ByRef diagnostic message.
'
' RETURNS
'   Variant
'     Success => Double value.
'     Failure => CVErr(xlErrNum) or CVErr(xlErrValue).
'
' BEHAVIOR
'   Accepts the full finite non-negative Double range.
'
' ERROR POLICY
'   Invalid domains, unsupported magnitudes or numerical failure return #NUM!.
'   Unexpected runtime errors return #VALUE!.
'   Diagnostics are written to Status.
'
' DEPENDENCIES
'   - PROB_DS_ValidateNonnegativeFinite
'   - PROB_SetStatus
'
' CALLED FROM
'   - Worksheet formulas
'   - M_STATS_PROBDIST_TEST
'
' UPDATED
'   2026-07-19 - Full production-hardening rewrite.
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
    'Validate the non-negative finite parameter
        If Not PROB_DS_ValidateNonnegativeFinite( _
            Mean, "Mean", FailMsg) Then GoTo Fail_Num

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
        PROB_SetStatus Status, _
            "Unexpected error in K_STATS_Poisson_Mean: " & Err.Description
    'Return worksheet value error
        K_STATS_Poisson_Mean = CVErr(xlErrValue)
End Function

'==============================================================================
' K_STATS_Poisson_Variance
'------------------------------------------------------------------------------
' PURPOSE
'   Returns the Poisson variance, which equals lambda.
'
' WHY
'   The moment is exact in closed form and does not call an iterative kernel.
'
' INPUTS
'   Mean    Poisson intensity lambda >= 0.
'   Status  Optional ByRef diagnostic message.
'
' RETURNS
'   Variant
'     Success => Double value.
'     Failure => CVErr(xlErrNum) or CVErr(xlErrValue).
'
' BEHAVIOR
'   Accepts the full finite non-negative Double range.
'
' ERROR POLICY
'   Invalid domains, unsupported magnitudes or numerical failure return #NUM!.
'   Unexpected runtime errors return #VALUE!.
'   Diagnostics are written to Status.
'
' DEPENDENCIES
'   - PROB_DS_ValidateNonnegativeFinite
'   - PROB_SetStatus
'
' CALLED FROM
'   - Worksheet formulas
'   - M_STATS_PROBDIST_TEST
'
' UPDATED
'   2026-07-19 - Full production-hardening rewrite.
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
    'Validate the non-negative finite parameter
        If Not PROB_DS_ValidateNonnegativeFinite( _
            Mean, "Mean", FailMsg) Then GoTo Fail_Num

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
        PROB_SetStatus Status, _
            "Unexpected error in K_STATS_Poisson_Variance: " & Err.Description
    'Return worksheet value error
        K_STATS_Poisson_Variance = CVErr(xlErrValue)
End Function

'==============================================================================
' K_STATS_Poisson_StdDev
'------------------------------------------------------------------------------
' PURPOSE
'   Returns the Poisson standard deviation Sqr(lambda).
'
' WHY
'   The moment is exact in closed form and does not call an iterative kernel.
'
' INPUTS
'   Mean    Poisson intensity lambda >= 0.
'   Status  Optional ByRef diagnostic message.
'
' RETURNS
'   Variant
'     Success => Double value.
'     Failure => CVErr(xlErrNum) or CVErr(xlErrValue).
'
' BEHAVIOR
'   Accepts the full finite non-negative Double range.
'
' ERROR POLICY
'   Invalid domains, unsupported magnitudes or numerical failure return #NUM!.
'   Unexpected runtime errors return #VALUE!.
'   Diagnostics are written to Status.
'
' DEPENDENCIES
'   - PROB_DS_ValidateNonnegativeFinite
'   - PROB_SetStatus
'
' CALLED FROM
'   - Worksheet formulas
'   - M_STATS_PROBDIST_TEST
'
' UPDATED
'   2026-07-19 - Full production-hardening rewrite.
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
    'Validate the non-negative finite parameter
        If Not PROB_DS_ValidateNonnegativeFinite( _
            Mean, "Mean", FailMsg) Then GoTo Fail_Num

'------------------------------------------------------------------------------
' COMPUTE
'------------------------------------------------------------------------------
    'Take the non-negative square root of the validated mean
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
        PROB_SetStatus Status, _
            "Unexpected error in K_STATS_Poisson_StdDev: " & Err.Description
    'Return worksheet value error
        K_STATS_Poisson_StdDev = CVErr(xlErrValue)
End Function

'==============================================================================
' K_STATS_Geometric_PMF
'------------------------------------------------------------------------------
' PURPOSE
'   Returns the Geometric mass P(X = NumberFailures), where X counts failures
'   before the first success.
'
' WHY
'   The logarithmic power is formed through guarded multiplication.
'
' INPUTS
'   NumberFailures  Failure count k; truncated toward zero.
'   ProbSuccess     Success probability p in (0, 1].
'   Status          Optional ByRef diagnostic message.
'
' RETURNS
'   Variant
'     Success => Double probability in [0, 1].
'     Failure => CVErr(xlErrNum) or CVErr(xlErrValue).
'
' BEHAVIOR
'   Uses p * (1-p)^k in the logarithmic domain.
'   Supports exact integer counts through 2^53 - 1.
'   p = 1 is handled as a point mass at zero.
'
' ERROR POLICY
'   Invalid domains, unsupported magnitudes or numerical failure return #NUM!.
'   Unexpected runtime errors return #VALUE!.
'   Diagnostics are written to Status.
'
' DEPENDENCIES
'   - PROB_DS_ValidateGeometricInputs
'   - PROB_DS_TryGeometricPMF
'   - PROB_SetStatus
'
' CALLED FROM
'   - Worksheet formulas
'   - M_STATS_PROBDIST_TEST
'
' UPDATED
'   2026-07-19 - Full production-hardening rewrite.
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
    Dim Value               As Double          'Computed probability mass
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
    'Validate the failure count and success probability
        If Not PROB_DS_ValidateGeometricInputs( _
            NumberFailures, ProbSuccess, K, FailMsg) Then GoTo Fail_Num

'------------------------------------------------------------------------------
' COMPUTE
'------------------------------------------------------------------------------
    'Evaluate the mass in the logarithmic domain
        If Not PROB_DS_TryGeometricPMF( _
            K, ProbSuccess, Value, FailMsg) Then GoTo Fail_Num

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
        PROB_SetStatus Status, _
            "Unexpected error in K_STATS_Geometric_PMF: " & Err.Description
    'Return worksheet value error
        K_STATS_Geometric_PMF = CVErr(xlErrValue)
End Function


Public Function K_STATS_Geometric_LogPMF( _
    ByVal NumberFailures As Double, _
    ByVal ProbSuccess As Double, _
    Optional ByRef Status As String = "") _
    As Variant
'
'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Dim K                   As Double          'Failures as a truncated integer
    Dim LogMass             As Double          'Natural logarithm of the mass
    Dim IsCertainZero       As Boolean         'Outcome is impossible or below Double range
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
    'Validate the failure count and success probability
        If Not PROB_DS_ValidateGeometricInputs( _
            NumberFailures, ProbSuccess, K, FailMsg) Then GoTo Fail_Num

'------------------------------------------------------------------------------
' COMPUTE
'------------------------------------------------------------------------------
    'Assemble the log mass in the logarithmic domain
        If Not PROB_DS_TryGeometricLogMass( _
            K, ProbSuccess, LogMass, IsCertainZero, FailMsg) Then GoTo Fail_Num
    'A zero-probability outcome has no defined log-mass
        If IsCertainZero Then
            FailMsg = "Log-mass is undefined: the outcome has probability zero"
            GoTo Fail_Num
        End If

        K_STATS_Geometric_LogPMF = LogMass

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
        K_STATS_Geometric_LogPMF = CVErr(xlErrNum)
    'Exit before the error handler
        Exit Function

'------------------------------------------------------------------------------
' ERROR HANDLER
'------------------------------------------------------------------------------
Err_Handler:
    'Write unexpected runtime errors to diagnostics
        PROB_SetStatus Status, _
            "Unexpected error in K_STATS_Geometric_LogPMF: " & Err.Description
    'Return worksheet value error
        K_STATS_Geometric_LogPMF = CVErr(xlErrValue)
End Function

'==============================================================================
' K_STATS_Geometric_Cumulative
'------------------------------------------------------------------------------
' PURPOSE
'   Returns the Geometric left-tail probability P(X <= NumberFailures).
'
' WHY
'   Expm1 preserves the small left tail when p or k is small.
'
' INPUTS
'   NumberFailures  Failure count k; truncated toward zero.
'   ProbSuccess     Success probability p in (0, 1].
'   Status          Optional ByRef diagnostic message.
'
' RETURNS
'   Variant
'     Success => Double probability in [0, 1].
'     Failure => CVErr(xlErrNum) or CVErr(xlErrValue).
'
' BEHAVIOR
'   Uses 1 - (1-p)^(k+1) as -Expm1((k+1)*Log1p(-p)).
'   A negative exponent overflow is translated to the exact limiting value one.
'
' ERROR POLICY
'   Invalid domains, unsupported magnitudes or numerical failure return #NUM!.
'   Unexpected runtime errors return #VALUE!.
'   Diagnostics are written to Status.
'
' DEPENDENCIES
'   - PROB_DS_ValidateGeometricInputs
'   - PROB_DS_TryGeometricCDF
'   - PROB_SetStatus
'
' CALLED FROM
'   - Worksheet formulas
'   - M_STATS_PROBDIST_TEST
'
' UPDATED
'   2026-07-19 - Full production-hardening rewrite.
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
    Dim Value               As Double          'Computed cumulative probability
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
    'Validate the failure count and success probability
        If Not PROB_DS_ValidateGeometricInputs( _
            NumberFailures, ProbSuccess, K, FailMsg) Then GoTo Fail_Num

'------------------------------------------------------------------------------
' COMPUTE
'------------------------------------------------------------------------------
    'Evaluate the left tail through the guarded closed form
        If Not PROB_DS_TryGeometricCDF( _
            K, ProbSuccess, Value, FailMsg) Then GoTo Fail_Num

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
        PROB_SetStatus Status, _
            "Unexpected error in K_STATS_Geometric_Cumulative: " & Err.Description
    'Return worksheet value error
        K_STATS_Geometric_Cumulative = CVErr(xlErrValue)
End Function

'==============================================================================
' K_STATS_Geometric_Survival
'------------------------------------------------------------------------------
' PURPOSE
'   Returns the Geometric right-tail probability P(X > NumberFailures).
'
' WHY
'   The survival probability is evaluated directly rather than as 1 - CDF.
'
' INPUTS
'   NumberFailures  Failure count k; truncated toward zero.
'   ProbSuccess     Success probability p in (0, 1].
'   Status          Optional ByRef diagnostic message.
'
' RETURNS
'   Variant
'     Success => Double probability in [0, 1].
'     Failure => CVErr(xlErrNum) or CVErr(xlErrValue).
'
' BEHAVIOR
'   Uses (1-p)^(k+1) in the logarithmic domain.
'   A negative exponent overflow is translated to the exact limiting value zero.
'
' ERROR POLICY
'   Invalid domains, unsupported magnitudes or numerical failure return #NUM!.
'   Unexpected runtime errors return #VALUE!.
'   Diagnostics are written to Status.
'
' DEPENDENCIES
'   - PROB_DS_ValidateGeometricInputs
'   - PROB_DS_TryGeometricSF
'   - PROB_SetStatus
'
' CALLED FROM
'   - Worksheet formulas
'   - M_STATS_PROBDIST_TEST
'
' UPDATED
'   2026-07-19 - Full production-hardening rewrite.
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
    Dim Value               As Double          'Computed survival probability
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
    'Validate the failure count and success probability
        If Not PROB_DS_ValidateGeometricInputs( _
            NumberFailures, ProbSuccess, K, FailMsg) Then GoTo Fail_Num

'------------------------------------------------------------------------------
' COMPUTE
'------------------------------------------------------------------------------
    'Evaluate the right tail directly
        If Not PROB_DS_TryGeometricSF( _
            K, ProbSuccess, Value, FailMsg) Then GoTo Fail_Num

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
        PROB_SetStatus Status, _
            "Unexpected error in K_STATS_Geometric_Survival: " & Err.Description
    'Return worksheet value error
        K_STATS_Geometric_Survival = CVErr(xlErrValue)
End Function

'==============================================================================
' K_STATS_Geometric_InverseCumulative
'------------------------------------------------------------------------------
' PURPOSE
'   Returns the least integer k satisfying P(X <= k) >= Probability.
'
' WHY
'   The closed-form seed is corrected against the smaller direct tail with hard
'   bounds on both the quantile and the number of correction steps.
'
' INPUTS
'   Probability  Target cumulative probability in (0, 1).
'   ProbSuccess  Success probability p in (0, 1].
'   Status       Optional ByRef diagnostic message.
'
' RETURNS
'   Variant
'     Success => Double containing an exactly representable integer quantile.
'     Failure => CVErr(xlErrNum) or CVErr(xlErrValue).
'
' BEHAVIOR
'   Uses ceil(Log1p(-Probability)/Log1p(-p)) - 1 as the seed.
'   Rejects a quantile above 2^53 - 1.
'   Uses the direct SF when Probability > 0.5.
'
' ERROR POLICY
'   Invalid domains, unsupported magnitudes or numerical failure return #NUM!.
'   Unexpected runtime errors return #VALUE!.
'   Diagnostics are written to Status.
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
'   2026-07-19 - Full production-hardening rewrite.
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
    'Validate the target cumulative probability
        If Not PROB_IsValidProbabilityOpen(Probability) Then
            FailMsg = "Probability must be strictly between 0 and 1"
            GoTo Fail_Num
        End If

    'Validate the half-open success-probability domain
        If Not PROB_DS_ValidateProbHalfOpen(ProbSuccess, FailMsg) Then GoTo Fail_Num

'------------------------------------------------------------------------------
' COMPUTE
'------------------------------------------------------------------------------
    'Certain success places every quantile at zero
        If ProbSuccess >= 1# Then
            K_STATS_Geometric_InverseCumulative = 0#
            GoTo Return_Success
        End If

    'Seed and correct against the smaller direct tail
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
        PROB_SetStatus Status, _
            "Unexpected error in K_STATS_Geometric_InverseCumulative: " & Err.Description
    'Return worksheet value error
        K_STATS_Geometric_InverseCumulative = CVErr(xlErrValue)
End Function

'==============================================================================
' K_STATS_Geometric_Mean
'------------------------------------------------------------------------------
' PURPOSE
'   Returns the Geometric mean (1-p)/p for the failures-before-success convention.
'
' WHY
'   The quotient is evaluated through the shared overflow contract.
'
' INPUTS
'   ProbSuccess  Success probability p in (0, 1].
'   Status       Optional ByRef diagnostic message.
'
' RETURNS
'   Variant
'     Success => Double value.
'     Failure => CVErr(xlErrNum) or CVErr(xlErrValue).
'
' BEHAVIOR
'   Returns zero at p = 1.
'   Reports #NUM! when the true mean exceeds Double range.
'
' ERROR POLICY
'   Invalid domains, unsupported magnitudes or numerical failure return #NUM!.
'   Unexpected runtime errors return #VALUE!.
'   Diagnostics are written to Status.
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
'   2026-07-19 - Full production-hardening rewrite.
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
    'Validate the half-open success-probability domain
        If Not PROB_DS_ValidateProbHalfOpen(ProbSuccess, FailMsg) Then GoTo Fail_Num

'------------------------------------------------------------------------------
' COMPUTE
'------------------------------------------------------------------------------
    'Mean = (1-p)/p through guarded division
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
        PROB_SetStatus Status, _
            "Unexpected error in K_STATS_Geometric_Mean: " & Err.Description
    'Return worksheet value error
        K_STATS_Geometric_Mean = CVErr(xlErrValue)
End Function

'==============================================================================
' K_STATS_Geometric_Variance
'------------------------------------------------------------------------------
' PURPOSE
'   Returns the Geometric variance (1-p)/p^2.
'
' WHY
'   The intermediate p^2 is guarded; underflow implies that the true variance
'   is already outside the representable Double range.
'
' INPUTS
'   ProbSuccess  Success probability p in (0, 1].
'   Status       Optional ByRef diagnostic message.
'
' RETURNS
'   Variant
'     Success => Double value.
'     Failure => CVErr(xlErrNum) or CVErr(xlErrValue).
'
' BEHAVIOR
'   Returns zero at p = 1.
'   Reports #NUM! when the true variance exceeds Double range.
'
' ERROR POLICY
'   Invalid domains, unsupported magnitudes or numerical failure return #NUM!.
'   Unexpected runtime errors return #VALUE!.
'   Diagnostics are written to Status.
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
'   2026-07-19 - Full production-hardening rewrite.
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
    Dim PP                  As Double          'Success probability squared
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
    'Validate the half-open success-probability domain
        If Not PROB_DS_ValidateProbHalfOpen(ProbSuccess, FailMsg) Then GoTo Fail_Num

'------------------------------------------------------------------------------
' COMPUTE
'------------------------------------------------------------------------------
    'Form p^2; underflow to zero is handled by the guarded division
        If Not PROB_TryMultiply(ProbSuccess, ProbSuccess, PP) Then
            FailMsg = "Geometric variance overflows Double range"
            GoTo Fail_Num
        End If

    'Variance = (1-p)/p^2
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
        PROB_SetStatus Status, _
            "Unexpected error in K_STATS_Geometric_Variance: " & Err.Description
    'Return worksheet value error
        K_STATS_Geometric_Variance = CVErr(xlErrValue)
End Function

'==============================================================================
' K_STATS_Geometric_StdDev
'------------------------------------------------------------------------------
' PURPOSE
'   Returns the Geometric standard deviation Sqr(1-p)/p.
'
' WHY
'   The direct standard-deviation formula avoids constructing an overflowing
'   variance when the standard deviation itself is still representable.
'
' INPUTS
'   ProbSuccess  Success probability p in (0, 1].
'   Status       Optional ByRef diagnostic message.
'
' RETURNS
'   Variant
'     Success => Double value.
'     Failure => CVErr(xlErrNum) or CVErr(xlErrValue).
'
' BEHAVIOR
'   Returns zero at p = 1.
'   Can remain finite even when the corresponding variance is not representable.
'
' ERROR POLICY
'   Invalid domains, unsupported magnitudes or numerical failure return #NUM!.
'   Unexpected runtime errors return #VALUE!.
'   Diagnostics are written to Status.
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
'   2026-07-19 - Full production-hardening rewrite.
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
    Dim Numerator           As Double          'Square root of 1-p
    Dim Value               As Double          'Computed standard deviation
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
    'Validate the half-open success-probability domain
        If Not PROB_DS_ValidateProbHalfOpen(ProbSuccess, FailMsg) Then GoTo Fail_Num

'------------------------------------------------------------------------------
' COMPUTE
'------------------------------------------------------------------------------
    'Use the direct stable formula Sqr(1-p)/p
        Numerator = Sqr(1# - ProbSuccess)

        If Not PROB_TryDivide(Numerator, ProbSuccess, Value) Then
            FailMsg = "Geometric standard deviation overflows Double range"
            GoTo Fail_Num
        End If

        K_STATS_Geometric_StdDev = Value

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
        PROB_SetStatus Status, _
            "Unexpected error in K_STATS_Geometric_StdDev: " & Err.Description
    'Return worksheet value error
        K_STATS_Geometric_StdDev = CVErr(xlErrValue)
End Function


Public Function K_STATS_NegativeBinomial_PMF( _
    ByVal NumberFailures As Double, _
    ByVal NumberSuccesses As Double, _
    ByVal ProbSuccess As Double, _
    Optional ByRef Status As String = "") _
    As Variant
'
'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Dim X                   As Double          'Failures as a truncated integer
    Dim R                   As Double          'Successes as a truncated integer
    Dim Value               As Double          'Computed probability mass
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
    'Validate counts and probability for the negative binomial mass
        If Not PROB_DS_ValidateNegBinomInputs( _
            NumberFailures, NumberSuccesses, ProbSuccess, _
            PROB_DS_MAX_EXACT_INTEGER, PROB_DS_MAX_EXACT_INTEGER, X, R, FailMsg) Then GoTo Fail_Num

'------------------------------------------------------------------------------
' COMPUTE
'------------------------------------------------------------------------------
    'Compute the mass by exponentiating the shared log-mass kernel
        If Not PROB_DS_TryNegativeBinomialPMF( _
            X, R, ProbSuccess, Value, FailMsg) Then GoTo Fail_Num

        K_STATS_NegativeBinomial_PMF = Value

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
        K_STATS_NegativeBinomial_PMF = CVErr(xlErrNum)
    'Exit before the error handler
        Exit Function

'------------------------------------------------------------------------------
' ERROR HANDLER
'------------------------------------------------------------------------------
Err_Handler:
    'Write unexpected runtime errors to diagnostics
        PROB_SetStatus Status, _
            "Unexpected error in K_STATS_NegativeBinomial_PMF: " & Err.Description
    'Return worksheet value error
        K_STATS_NegativeBinomial_PMF = CVErr(xlErrValue)
End Function


Public Function K_STATS_NegativeBinomial_LogPMF( _
    ByVal NumberFailures As Double, _
    ByVal NumberSuccesses As Double, _
    ByVal ProbSuccess As Double, _
    Optional ByRef Status As String = "") _
    As Variant
'
'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Dim X                   As Double          'Failures as a truncated integer
    Dim R                   As Double          'Successes as a truncated integer
    Dim LogMass             As Double          'Natural logarithm of the mass
    Dim IsCertainZero       As Boolean         'Outcome is impossible or below Double range
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
    'Validate counts and probability for the negative binomial mass
        If Not PROB_DS_ValidateNegBinomInputs( _
            NumberFailures, NumberSuccesses, ProbSuccess, _
            PROB_DS_MAX_EXACT_INTEGER, PROB_DS_MAX_EXACT_INTEGER, X, R, FailMsg) Then GoTo Fail_Num

'------------------------------------------------------------------------------
' COMPUTE
'------------------------------------------------------------------------------
    'Assemble the log mass through the shared Loader arrangement
        If Not PROB_DS_TryNegativeBinomialLogMass( _
            X, R, ProbSuccess, LogMass, IsCertainZero, FailMsg) Then GoTo Fail_Num
    'A zero-probability outcome has no defined log-mass
        If IsCertainZero Then
            FailMsg = "Log-mass is undefined: the outcome has probability zero"
            GoTo Fail_Num
        End If

        K_STATS_NegativeBinomial_LogPMF = LogMass

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
        K_STATS_NegativeBinomial_LogPMF = CVErr(xlErrNum)
    'Exit before the error handler
        Exit Function

'------------------------------------------------------------------------------
' ERROR HANDLER
'------------------------------------------------------------------------------
Err_Handler:
    'Write unexpected runtime errors to diagnostics
        PROB_SetStatus Status, _
            "Unexpected error in K_STATS_NegativeBinomial_LogPMF: " & Err.Description
    'Return worksheet value error
        K_STATS_NegativeBinomial_LogPMF = CVErr(xlErrValue)
End Function


Public Function K_STATS_NegativeBinomial_Cumulative( _
    ByVal NumberFailures As Double, _
    ByVal NumberSuccesses As Double, _
    ByVal ProbSuccess As Double, _
    Optional ByRef Status As String = "") _
    As Variant
'
'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Dim X                   As Double          'Failures as a truncated integer
    Dim R                   As Double          'Successes as a truncated integer
    Dim Value               As Double          'Computed cumulative probability
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
    'Validate counts and probability for the negative binomial mass
        If Not PROB_DS_ValidateNegBinomInputs( _
            NumberFailures, NumberSuccesses, ProbSuccess, _
            PROB_DS_MAX_NEGBINOM_KERNEL_COUNT, PROB_DS_MAX_BINOMIAL_KERNEL_N, X, R, FailMsg) Then GoTo Fail_Num

'------------------------------------------------------------------------------
' COMPUTE
'------------------------------------------------------------------------------
    'P(X <= x) = I_p(r, x + 1) through the no-cancellation incomplete beta
        If Not PROB_DS_TryNegativeBinomialCDF( _
            X, R, ProbSuccess, Value, FailMsg) Then GoTo Fail_Num

        K_STATS_NegativeBinomial_Cumulative = Value

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
        K_STATS_NegativeBinomial_Cumulative = CVErr(xlErrNum)
    'Exit before the error handler
        Exit Function

'------------------------------------------------------------------------------
' ERROR HANDLER
'------------------------------------------------------------------------------
Err_Handler:
    'Write unexpected runtime errors to diagnostics
        PROB_SetStatus Status, _
            "Unexpected error in K_STATS_NegativeBinomial_Cumulative: " & Err.Description
    'Return worksheet value error
        K_STATS_NegativeBinomial_Cumulative = CVErr(xlErrValue)
End Function


Public Function K_STATS_NegativeBinomial_Survival( _
    ByVal NumberFailures As Double, _
    ByVal NumberSuccesses As Double, _
    ByVal ProbSuccess As Double, _
    Optional ByRef Status As String = "") _
    As Variant
'
'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Dim X                   As Double          'Failures as a truncated integer
    Dim R                   As Double          'Successes as a truncated integer
    Dim Value               As Double          'Computed survival probability
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
    'Validate counts and probability for the negative binomial mass
        If Not PROB_DS_ValidateNegBinomInputs( _
            NumberFailures, NumberSuccesses, ProbSuccess, _
            PROB_DS_MAX_NEGBINOM_KERNEL_COUNT, PROB_DS_MAX_BINOMIAL_KERNEL_N, X, R, FailMsg) Then GoTo Fail_Num

'------------------------------------------------------------------------------
' COMPUTE
'------------------------------------------------------------------------------
    'P(X > x) = I_(1-p)(x + 1, r) computed directly as the small upper tail
        If Not PROB_DS_TryNegativeBinomialSF( _
            X, R, ProbSuccess, Value, FailMsg) Then GoTo Fail_Num

        K_STATS_NegativeBinomial_Survival = Value

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
        K_STATS_NegativeBinomial_Survival = CVErr(xlErrNum)
    'Exit before the error handler
        Exit Function

'------------------------------------------------------------------------------
' ERROR HANDLER
'------------------------------------------------------------------------------
Err_Handler:
    'Write unexpected runtime errors to diagnostics
        PROB_SetStatus Status, _
            "Unexpected error in K_STATS_NegativeBinomial_Survival: " & Err.Description
    'Return worksheet value error
        K_STATS_NegativeBinomial_Survival = CVErr(xlErrValue)
End Function


Public Function K_STATS_NegativeBinomial_InverseCumulative( _
    ByVal Probability As Double, _
    ByVal NumberSuccesses As Double, _
    ByVal ProbSuccess As Double, _
    Optional ByRef Status As String = "") _
    As Variant
'
'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Dim R                   As Double          'Successes as a truncated integer
    Dim Quantile            As Double          'Least qualifying failure count
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
    'Require a strictly interior cumulative probability
        If Not PROB_DS_ValidateProbOpen(Probability, FailMsg) Then GoTo Fail_Num
    'Validate the success count and success probability
        If Not PROB_DS_ValidateNegBinomSuccessProb( _
            NumberSuccesses, ProbSuccess, PROB_DS_MAX_BINOMIAL_KERNEL_N, R, FailMsg) Then GoTo Fail_Num

'------------------------------------------------------------------------------
' COMPUTE
'------------------------------------------------------------------------------
    'Seeded, bracketed integer search on the smaller direct tail
        If Not PROB_DS_TryNegativeBinomialInverse( _
            Probability, R, ProbSuccess, Quantile, FailMsg) Then GoTo Fail_Num

        K_STATS_NegativeBinomial_InverseCumulative = Quantile

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
        K_STATS_NegativeBinomial_InverseCumulative = CVErr(xlErrNum)
    'Exit before the error handler
        Exit Function

'------------------------------------------------------------------------------
' ERROR HANDLER
'------------------------------------------------------------------------------
Err_Handler:
    'Write unexpected runtime errors to diagnostics
        PROB_SetStatus Status, _
            "Unexpected error in K_STATS_NegativeBinomial_InverseCumulative: " & Err.Description
    'Return worksheet value error
        K_STATS_NegativeBinomial_InverseCumulative = CVErr(xlErrValue)
End Function


Public Function K_STATS_NegativeBinomial_Mean( _
    ByVal NumberSuccesses As Double, _
    ByVal ProbSuccess As Double, _
    Optional ByRef Status As String = "") _
    As Variant
'
'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Dim R                   As Double          'Successes as a truncated integer
    Dim Q                   As Double          'Failure probability 1-p
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
    'Validate the success count and success probability
        If Not PROB_DS_ValidateNegBinomSuccessProb( _
            NumberSuccesses, ProbSuccess, PROB_DS_MAX_EXACT_INTEGER, R, FailMsg) Then GoTo Fail_Num

'------------------------------------------------------------------------------
' COMPUTE
'------------------------------------------------------------------------------
        Q = 1# - ProbSuccess
        Value = R * Q / ProbSuccess
        If Not PROB_IsFinite(Value) Then
            FailMsg = "Negative binomial mean overflows Double range"
            GoTo Fail_Num
        End If

        K_STATS_NegativeBinomial_Mean = Value

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
        K_STATS_NegativeBinomial_Mean = CVErr(xlErrNum)
    'Exit before the error handler
        Exit Function

'------------------------------------------------------------------------------
' ERROR HANDLER
'------------------------------------------------------------------------------
Err_Handler:
    'Write unexpected runtime errors to diagnostics
        PROB_SetStatus Status, _
            "Unexpected error in K_STATS_NegativeBinomial_Mean: " & Err.Description
    'Return worksheet value error
        K_STATS_NegativeBinomial_Mean = CVErr(xlErrValue)
End Function


Public Function K_STATS_NegativeBinomial_Variance( _
    ByVal NumberSuccesses As Double, _
    ByVal ProbSuccess As Double, _
    Optional ByRef Status As String = "") _
    As Variant
'
'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Dim R                   As Double          'Successes as a truncated integer
    Dim Q                   As Double          'Failure probability 1-p
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
    'Validate the success count and success probability
        If Not PROB_DS_ValidateNegBinomSuccessProb( _
            NumberSuccesses, ProbSuccess, PROB_DS_MAX_EXACT_INTEGER, R, FailMsg) Then GoTo Fail_Num

'------------------------------------------------------------------------------
' COMPUTE
'------------------------------------------------------------------------------
        Q = 1# - ProbSuccess
        Value = R * Q / (ProbSuccess * ProbSuccess)
        If Not PROB_IsFinite(Value) Then
            FailMsg = "Negative binomial variance overflows Double range"
            GoTo Fail_Num
        End If

        K_STATS_NegativeBinomial_Variance = Value

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
        K_STATS_NegativeBinomial_Variance = CVErr(xlErrNum)
    'Exit before the error handler
        Exit Function

'------------------------------------------------------------------------------
' ERROR HANDLER
'------------------------------------------------------------------------------
Err_Handler:
    'Write unexpected runtime errors to diagnostics
        PROB_SetStatus Status, _
            "Unexpected error in K_STATS_NegativeBinomial_Variance: " & Err.Description
    'Return worksheet value error
        K_STATS_NegativeBinomial_Variance = CVErr(xlErrValue)
End Function


Public Function K_STATS_NegativeBinomial_StdDev( _
    ByVal NumberSuccesses As Double, _
    ByVal ProbSuccess As Double, _
    Optional ByRef Status As String = "") _
    As Variant
'
'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Dim R                   As Double          'Successes as a truncated integer
    Dim Q                   As Double          'Failure probability 1-p
    Dim Value               As Double          'Computed standard deviation
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
    'Validate the success count and success probability
        If Not PROB_DS_ValidateNegBinomSuccessProb( _
            NumberSuccesses, ProbSuccess, PROB_DS_MAX_EXACT_INTEGER, R, FailMsg) Then GoTo Fail_Num

'------------------------------------------------------------------------------
' COMPUTE
'------------------------------------------------------------------------------
        Q = 1# - ProbSuccess
        Value = Sqr(R * Q) / ProbSuccess
        If Not PROB_IsFinite(Value) Then
            FailMsg = "Negative binomial standard deviation overflows Double range"
            GoTo Fail_Num
        End If

        K_STATS_NegativeBinomial_StdDev = Value

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
        K_STATS_NegativeBinomial_StdDev = CVErr(xlErrNum)
    'Exit before the error handler
        Exit Function

'------------------------------------------------------------------------------
' ERROR HANDLER
'------------------------------------------------------------------------------
Err_Handler:
    'Write unexpected runtime errors to diagnostics
        PROB_SetStatus Status, _
            "Unexpected error in K_STATS_NegativeBinomial_StdDev: " & Err.Description
    'Return worksheet value error
        K_STATS_NegativeBinomial_StdDev = CVErr(xlErrValue)
End Function


Public Function K_STATS_Hypergeometric_PMF( _
    ByVal SampleSuccesses As Double, _
    ByVal SampleSize As Double, _
    ByVal PopulationSuccesses As Double, _
    ByVal PopulationSize As Double, _
    Optional ByRef Status As String = "") _
    As Variant
'
'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Dim K                   As Double          'Sample successes as a truncated integer
    Dim n                   As Double          'Sample size as a truncated integer
    Dim Kp                  As Double          'Population successes as a truncated integer
    Dim Np                  As Double          'Population size as a truncated integer
    Dim Value               As Double          'Computed probability mass
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
    'Validate sample and population counts against the support
        If Not PROB_DS_ValidateHypergeometricInputs( _
            SampleSuccesses, SampleSize, PopulationSuccesses, PopulationSize, True, _
            K, n, Kp, Np, FailMsg) Then GoTo Fail_Num

'------------------------------------------------------------------------------
' COMPUTE
'------------------------------------------------------------------------------
    'Compute the mass by exponentiating the shared log-mass kernel
        If Not PROB_DS_TryHypergeometricPMF( _
            K, n, Kp, Np, Value, FailMsg) Then GoTo Fail_Num

        K_STATS_Hypergeometric_PMF = Value

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
        K_STATS_Hypergeometric_PMF = CVErr(xlErrNum)
    'Exit before the error handler
        Exit Function

'------------------------------------------------------------------------------
' ERROR HANDLER
'------------------------------------------------------------------------------
Err_Handler:
    'Write unexpected runtime errors to diagnostics
        PROB_SetStatus Status, _
            "Unexpected error in K_STATS_Hypergeometric_PMF: " & Err.Description
    'Return worksheet value error
        K_STATS_Hypergeometric_PMF = CVErr(xlErrValue)
End Function


Public Function K_STATS_Hypergeometric_LogPMF( _
    ByVal SampleSuccesses As Double, _
    ByVal SampleSize As Double, _
    ByVal PopulationSuccesses As Double, _
    ByVal PopulationSize As Double, _
    Optional ByRef Status As String = "") _
    As Variant
'
'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Dim K                   As Double          'Sample successes as a truncated integer
    Dim n                   As Double          'Sample size as a truncated integer
    Dim Kp                  As Double          'Population successes as a truncated integer
    Dim Np                  As Double          'Population size as a truncated integer
    Dim LogMass             As Double          'Natural logarithm of the mass
    Dim IsCertainZero       As Boolean         'Outcome is impossible or below Double range
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
    'Validate sample and population counts against the support
        If Not PROB_DS_ValidateHypergeometricInputs( _
            SampleSuccesses, SampleSize, PopulationSuccesses, PopulationSize, True, _
            K, n, Kp, Np, FailMsg) Then GoTo Fail_Num

'------------------------------------------------------------------------------
' COMPUTE
'------------------------------------------------------------------------------
    'Assemble the log mass from three shared Binomial log-masses
        If Not PROB_DS_TryHypergeometricLogMass( _
            K, n, Kp, Np, LogMass, IsCertainZero, FailMsg) Then GoTo Fail_Num
    'A zero-probability outcome has no defined log-mass
        If IsCertainZero Then
            FailMsg = "Log-mass is undefined: the outcome has probability zero"
            GoTo Fail_Num
        End If

        K_STATS_Hypergeometric_LogPMF = LogMass

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
        K_STATS_Hypergeometric_LogPMF = CVErr(xlErrNum)
    'Exit before the error handler
        Exit Function

'------------------------------------------------------------------------------
' ERROR HANDLER
'------------------------------------------------------------------------------
Err_Handler:
    'Write unexpected runtime errors to diagnostics
        PROB_SetStatus Status, _
            "Unexpected error in K_STATS_Hypergeometric_LogPMF: " & Err.Description
    'Return worksheet value error
        K_STATS_Hypergeometric_LogPMF = CVErr(xlErrValue)
End Function


Public Function K_STATS_Hypergeometric_Cumulative( _
    ByVal SampleSuccesses As Double, _
    ByVal SampleSize As Double, _
    ByVal PopulationSuccesses As Double, _
    ByVal PopulationSize As Double, _
    Optional ByRef Status As String = "") _
    As Variant
'
'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Dim K                   As Double          'Sample successes as a truncated integer
    Dim n                   As Double          'Sample size as a truncated integer
    Dim Kp                  As Double          'Population successes as a truncated integer
    Dim Np                  As Double          'Population size as a truncated integer
    Dim Value               As Double          'Computed cumulative probability
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
    'Validate sample and population counts against the support
        If Not PROB_DS_ValidateHypergeometricInputs( _
            SampleSuccesses, SampleSize, PopulationSuccesses, PopulationSize, False, _
            K, n, Kp, Np, FailMsg) Then GoTo Fail_Num

'------------------------------------------------------------------------------
' COMPUTE
'------------------------------------------------------------------------------
    'P(X <= k) by directional ratio-series summation of the near tail
        If Not PROB_DS_TryHypergeometricCDF( _
            K, n, Kp, Np, Value, FailMsg) Then GoTo Fail_Num

        K_STATS_Hypergeometric_Cumulative = Value

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
        K_STATS_Hypergeometric_Cumulative = CVErr(xlErrNum)
    'Exit before the error handler
        Exit Function

'------------------------------------------------------------------------------
' ERROR HANDLER
'------------------------------------------------------------------------------
Err_Handler:
    'Write unexpected runtime errors to diagnostics
        PROB_SetStatus Status, _
            "Unexpected error in K_STATS_Hypergeometric_Cumulative: " & Err.Description
    'Return worksheet value error
        K_STATS_Hypergeometric_Cumulative = CVErr(xlErrValue)
End Function


Public Function K_STATS_Hypergeometric_Survival( _
    ByVal SampleSuccesses As Double, _
    ByVal SampleSize As Double, _
    ByVal PopulationSuccesses As Double, _
    ByVal PopulationSize As Double, _
    Optional ByRef Status As String = "") _
    As Variant
'
'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Dim K                   As Double          'Sample successes as a truncated integer
    Dim n                   As Double          'Sample size as a truncated integer
    Dim Kp                  As Double          'Population successes as a truncated integer
    Dim Np                  As Double          'Population size as a truncated integer
    Dim Value               As Double          'Computed survival probability
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
    'Validate sample and population counts against the support
        If Not PROB_DS_ValidateHypergeometricInputs( _
            SampleSuccesses, SampleSize, PopulationSuccesses, PopulationSize, False, _
            K, n, Kp, Np, FailMsg) Then GoTo Fail_Num

'------------------------------------------------------------------------------
' COMPUTE
'------------------------------------------------------------------------------
    'P(X > k) as the small upper tail, summed directly with no cancellation
        If Not PROB_DS_TryHypergeometricSF( _
            K, n, Kp, Np, Value, FailMsg) Then GoTo Fail_Num

        K_STATS_Hypergeometric_Survival = Value

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
        K_STATS_Hypergeometric_Survival = CVErr(xlErrNum)
    'Exit before the error handler
        Exit Function

'------------------------------------------------------------------------------
' ERROR HANDLER
'------------------------------------------------------------------------------
Err_Handler:
    'Write unexpected runtime errors to diagnostics
        PROB_SetStatus Status, _
            "Unexpected error in K_STATS_Hypergeometric_Survival: " & Err.Description
    'Return worksheet value error
        K_STATS_Hypergeometric_Survival = CVErr(xlErrValue)
End Function


Public Function K_STATS_Hypergeometric_InverseCumulative( _
    ByVal Probability As Double, _
    ByVal SampleSize As Double, _
    ByVal PopulationSuccesses As Double, _
    ByVal PopulationSize As Double, _
    Optional ByRef Status As String = "") _
    As Variant
'
'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Dim n                   As Double          'Sample size as a truncated integer
    Dim Kp                  As Double          'Population successes as a truncated integer
    Dim Np                  As Double          'Population size as a truncated integer
    Dim Quantile            As Double          'Least qualifying sample-success count
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
    'Require a strictly interior cumulative probability
        If Not PROB_DS_ValidateProbOpen(Probability, FailMsg) Then GoTo Fail_Num
    'Validate the population and sample sizes
        If Not PROB_DS_ValidateHypergeometricPopulation( _
            PopulationSize, PopulationSuccesses, SampleSize, Np, Kp, n, FailMsg) Then GoTo Fail_Num

'------------------------------------------------------------------------------
' COMPUTE
'------------------------------------------------------------------------------
    'Seeded, bracketed integer search across the support
        If Not PROB_DS_TryHypergeometricInverse( _
            Probability, n, Kp, Np, Quantile, FailMsg) Then GoTo Fail_Num

        K_STATS_Hypergeometric_InverseCumulative = Quantile

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
        K_STATS_Hypergeometric_InverseCumulative = CVErr(xlErrNum)
    'Exit before the error handler
        Exit Function

'------------------------------------------------------------------------------
' ERROR HANDLER
'------------------------------------------------------------------------------
Err_Handler:
    'Write unexpected runtime errors to diagnostics
        PROB_SetStatus Status, _
            "Unexpected error in K_STATS_Hypergeometric_InverseCumulative: " & Err.Description
    'Return worksheet value error
        K_STATS_Hypergeometric_InverseCumulative = CVErr(xlErrValue)
End Function


Public Function K_STATS_Hypergeometric_Mean( _
    ByVal SampleSize As Double, _
    ByVal PopulationSuccesses As Double, _
    ByVal PopulationSize As Double, _
    Optional ByRef Status As String = "") _
    As Variant
'
'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Dim n                   As Double          'Sample size as a truncated integer
    Dim Kp                  As Double          'Population successes as a truncated integer
    Dim Np                  As Double          'Population size as a truncated integer
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
    'Validate the population and sample sizes
        If Not PROB_DS_ValidateHypergeometricPopulation( _
            PopulationSize, PopulationSuccesses, SampleSize, Np, Kp, n, FailMsg) Then GoTo Fail_Num

'------------------------------------------------------------------------------
' COMPUTE
'------------------------------------------------------------------------------
    'Mean = n * K / N
        Value = n * Kp / Np
        If Not PROB_IsFinite(Value) Then
            FailMsg = "Hypergeometric mean overflows Double range"
            GoTo Fail_Num
        End If

        K_STATS_Hypergeometric_Mean = Value

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
        K_STATS_Hypergeometric_Mean = CVErr(xlErrNum)
    'Exit before the error handler
        Exit Function

'------------------------------------------------------------------------------
' ERROR HANDLER
'------------------------------------------------------------------------------
Err_Handler:
    'Write unexpected runtime errors to diagnostics
        PROB_SetStatus Status, _
            "Unexpected error in K_STATS_Hypergeometric_Mean: " & Err.Description
    'Return worksheet value error
        K_STATS_Hypergeometric_Mean = CVErr(xlErrValue)
End Function


Public Function K_STATS_Hypergeometric_Variance( _
    ByVal SampleSize As Double, _
    ByVal PopulationSuccesses As Double, _
    ByVal PopulationSize As Double, _
    Optional ByRef Status As String = "") _
    As Variant
'
'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Dim n                   As Double          'Sample size as a truncated integer
    Dim Kp                  As Double          'Population successes as a truncated integer
    Dim Np                  As Double          'Population size as a truncated integer
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
    'Validate the population and sample sizes
        If Not PROB_DS_ValidateHypergeometricPopulation( _
            PopulationSize, PopulationSuccesses, SampleSize, Np, Kp, n, FailMsg) Then GoTo Fail_Num

'------------------------------------------------------------------------------
' COMPUTE
'------------------------------------------------------------------------------
    'Variance = n (K/N) ((N-K)/N) ((N-n)/(N-1))
        If Np <= 1# Then
            Value = 0#
        Else
            Value = n * (Kp / Np) * ((Np - Kp) / Np) * ((Np - n) / (Np - 1#))
        End If
        If Not PROB_IsFinite(Value) Then
            FailMsg = "Hypergeometric variance overflows Double range"
            GoTo Fail_Num
        End If

        K_STATS_Hypergeometric_Variance = Value

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
        K_STATS_Hypergeometric_Variance = CVErr(xlErrNum)
    'Exit before the error handler
        Exit Function

'------------------------------------------------------------------------------
' ERROR HANDLER
'------------------------------------------------------------------------------
Err_Handler:
    'Write unexpected runtime errors to diagnostics
        PROB_SetStatus Status, _
            "Unexpected error in K_STATS_Hypergeometric_Variance: " & Err.Description
    'Return worksheet value error
        K_STATS_Hypergeometric_Variance = CVErr(xlErrValue)
End Function


Public Function K_STATS_Hypergeometric_StdDev( _
    ByVal SampleSize As Double, _
    ByVal PopulationSuccesses As Double, _
    ByVal PopulationSize As Double, _
    Optional ByRef Status As String = "") _
    As Variant
'
'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Dim n                   As Double          'Sample size as a truncated integer
    Dim Kp                  As Double          'Population successes as a truncated integer
    Dim Np                  As Double          'Population size as a truncated integer
    Dim Value               As Double          'Computed standard deviation
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
    'Validate the population and sample sizes
        If Not PROB_DS_ValidateHypergeometricPopulation( _
            PopulationSize, PopulationSuccesses, SampleSize, Np, Kp, n, FailMsg) Then GoTo Fail_Num

'------------------------------------------------------------------------------
' COMPUTE
'------------------------------------------------------------------------------
    'Standard deviation is the square root of the variance
        If Np <= 1# Then
            Value = 0#
        Else
            Value = Sqr(n * (Kp / Np) * ((Np - Kp) / Np) * ((Np - n) / (Np - 1#)))
        End If
        If Not PROB_IsFinite(Value) Then
            FailMsg = "Hypergeometric standard deviation overflows Double range"
            GoTo Fail_Num
        End If

        K_STATS_Hypergeometric_StdDev = Value

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
        K_STATS_Hypergeometric_StdDev = CVErr(xlErrNum)
    'Exit before the error handler
        Exit Function

'------------------------------------------------------------------------------
' ERROR HANDLER
'------------------------------------------------------------------------------
Err_Handler:
    'Write unexpected runtime errors to diagnostics
        PROB_SetStatus Status, _
            "Unexpected error in K_STATS_Hypergeometric_StdDev: " & Err.Description
    'Return worksheet value error
        K_STATS_Hypergeometric_StdDev = CVErr(xlErrValue)
End Function


'==============================================================================
' PRIVATE VALIDATION KERNELS
'==============================================================================


Private Function PROB_DS_ValidateCount( _
    ByVal RawValue As Double, _
    ByVal MaxSupported As Double, _
    ByVal ArgName As String, _
    ByVal Context As String, _
    ByRef CountOut As Double, _
    ByRef FailMsg As String) _
    As Boolean
'
'==============================================================================
' PROB_DS_ValidateCount
'------------------------------------------------------------------------------
' PURPOSE
'   Validates a non-negative finite worksheet count, truncates it toward zero
'   and enforces a documented maximum supported integer.
'
' INPUTS
'   RawValue      Unvalidated worksheet value.
'   MaxSupported  Largest accepted truncated count.
'   ArgName       User-facing argument name.
'   Context       Optional user-facing numerical-domain description.
'
' RETURNS
'   Boolean
'     TRUE  => CountOut contains the validated truncated integer.
'     FALSE => FailMsg contains a diagnostic.
'
' ERROR POLICY
'   - Negative, non-finite or unsupported counts fail explicitly.
'
' UPDATED
'   2026-07-19 - Exact-integer and supported-domain hardening.
'==============================================================================
'
'------------------------------------------------------------------------------
' VALIDATE
'------------------------------------------------------------------------------
    'Reject a non-finite count
        If Not PROB_IsFinite(RawValue) Then
            FailMsg = ArgName & " must be a finite number"
            Exit Function
        End If

    'Reject a negative count before truncation
        If RawValue < 0# Then
            FailMsg = ArgName & " must not be negative"
            Exit Function
        End If

'------------------------------------------------------------------------------
' TRUNCATE
'------------------------------------------------------------------------------
    'RawValue is non-negative, so Int truncates toward zero
        CountOut = Int(RawValue)

'------------------------------------------------------------------------------
' ENFORCE SUPPORTED RANGE
'------------------------------------------------------------------------------
    'Reject counts that cannot be handled by the selected numerical kernel
        If CountOut > MaxSupported Then
            FailMsg = ArgName & " exceeds the supported maximum of " & _
                      CStr(MaxSupported)

            If Len(Context) > 0 Then
                FailMsg = FailMsg & " for " & Context
            End If

            Exit Function
        End If

'------------------------------------------------------------------------------
' RETURN SUCCESS
'------------------------------------------------------------------------------
        PROB_DS_ValidateCount = True
End Function


Private Function PROB_DS_ValidateTrialsExact( _
    ByVal Trials As Double, _
    ByRef NOut As Double, _
    ByRef FailMsg As String) _
    As Boolean
'
'==============================================================================
' PROB_DS_ValidateTrialsExact
'------------------------------------------------------------------------------
' PURPOSE
'   Validates Binomial Trials in the exact-integer Double domain.
'
' UPDATED
'   2026-07-19
'==============================================================================
'
        PROB_DS_ValidateTrialsExact = PROB_DS_ValidateCount( _
            Trials, _
            PROB_DS_MAX_EXACT_INTEGER, _
            "Trials", _
            "exact-integer Binomial calculations", _
            NOut, _
            FailMsg)
End Function


Private Function PROB_DS_ValidateTrialsKernel( _
    ByVal Trials As Double, _
    ByRef NOut As Double, _
    ByRef FailMsg As String) _
    As Boolean
'
'==============================================================================
' PROB_DS_ValidateTrialsKernel
'------------------------------------------------------------------------------
' PURPOSE
'   Validates Binomial Trials against the tested incomplete-beta domain.
'
' UPDATED
'   2026-07-19
'==============================================================================
'
        PROB_DS_ValidateTrialsKernel = PROB_DS_ValidateCount( _
            Trials, _
            PROB_DS_MAX_BINOMIAL_KERNEL_N, _
            "Trials", _
            "Binomial cumulative, survival and inverse calculations", _
            NOut, _
            FailMsg)
End Function


Private Function PROB_DS_ValidateNonnegativeFinite( _
    ByVal Value As Double, _
    ByVal ArgName As String, _
    ByRef FailMsg As String) _
    As Boolean
'
'==============================================================================
' PROB_DS_ValidateNonnegativeFinite
'------------------------------------------------------------------------------
' PURPOSE
'   Validates a finite non-negative real parameter.
'
' UPDATED
'   2026-07-19
'==============================================================================
'
    'Reject non-finite values
        If Not PROB_IsFinite(Value) Then
            FailMsg = ArgName & " must be a finite non-negative number"
            Exit Function
        End If

    'Reject negative values
        If Value < 0# Then
            FailMsg = ArgName & " must be a finite non-negative number"
            Exit Function
        End If

        PROB_DS_ValidateNonnegativeFinite = True
End Function


Private Function PROB_DS_ValidatePoissonMeanPMF( _
    ByVal Mean As Double, _
    ByRef FailMsg As String) _
    As Boolean
'
'==============================================================================
' PROB_DS_ValidatePoissonMeanPMF
'------------------------------------------------------------------------------
' PURPOSE
'   Validates a Poisson mean for Loader's mass calculation.
'
' UPDATED
'   2026-07-19
'==============================================================================
'
        If Not PROB_DS_ValidateNonnegativeFinite(Mean, "Mean", FailMsg) Then Exit Function

        If Mean > PROB_DS_MAX_EXACT_INTEGER Then
            FailMsg = "Mean exceeds the supported maximum of " & _
                      CStr(PROB_DS_MAX_EXACT_INTEGER) & _
                      " for the Poisson mass calculation"
            Exit Function
        End If

        PROB_DS_ValidatePoissonMeanPMF = True
End Function


Private Function PROB_DS_ValidatePoissonMeanKernel( _
    ByVal Mean As Double, _
    ByRef FailMsg As String) _
    As Boolean
'
'==============================================================================
' PROB_DS_ValidatePoissonMeanKernel
'------------------------------------------------------------------------------
' PURPOSE
'   Validates a Poisson mean against the tested incomplete-gamma domain.
'
' UPDATED
'   2026-07-19
'==============================================================================
'
        If Not PROB_DS_ValidateNonnegativeFinite(Mean, "Mean", FailMsg) Then Exit Function

        If Mean > PROB_DS_MAX_POISSON_KERNEL_MEAN Then
            FailMsg = "Mean exceeds the supported maximum of " & _
                      CStr(PROB_DS_MAX_POISSON_KERNEL_MEAN) & _
                      " for Poisson cumulative, survival and inverse calculations"
            Exit Function
        End If

        PROB_DS_ValidatePoissonMeanKernel = True
End Function


Private Function PROB_DS_ValidateProbClosed( _
    ByVal Probability As Double, _
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
    'Reject non-finite probabilities
        If Not PROB_IsFinite(Probability) Then
            FailMsg = "ProbSuccess must be between 0 and 1"
            Exit Function
        End If

    'Reject values outside the closed unit interval
        If Probability < 0# Or Probability > 1# Then
            FailMsg = "ProbSuccess must be between 0 and 1"
            Exit Function
        End If

        PROB_DS_ValidateProbClosed = True
End Function


Private Function PROB_DS_ValidateProbOpen( _
    ByVal Probability As Double, _
    ByRef FailMsg As String) _
    As Boolean
'
'==============================================================================
' PROB_DS_ValidateProbOpen
'------------------------------------------------------------------------------
' PURPOSE
'   Validates a success probability over the open interval (0, 1).
'
' UPDATED
'   2026-07-19
'==============================================================================
'
    'Reject non-finite probabilities
        If Not PROB_IsFinite(Probability) Then
            FailMsg = "ProbSuccess must be strictly between 0 and 1"
            Exit Function
        End If

    'Reject endpoints and values outside the unit interval
        If Probability <= 0# Or Probability >= 1# Then
            FailMsg = "ProbSuccess must be strictly between 0 and 1"
            Exit Function
        End If

        PROB_DS_ValidateProbOpen = True
End Function


Private Function PROB_DS_ValidateProbHalfOpen( _
    ByVal Probability As Double, _
    ByRef FailMsg As String) _
    As Boolean
'
'==============================================================================
' PROB_DS_ValidateProbHalfOpen
'------------------------------------------------------------------------------
' PURPOSE
'   Validates a success probability over the interval (0, 1].
'
' UPDATED
'   2026-07-19
'==============================================================================
'
    'Reject non-finite probabilities
        If Not PROB_IsFinite(Probability) Then
            FailMsg = "ProbSuccess must be greater than 0 and at most 1"
            Exit Function
        End If

    'Reject zero, negative values and values above one
        If Probability <= 0# Or Probability > 1# Then
            FailMsg = "ProbSuccess must be greater than 0 and at most 1"
            Exit Function
        End If

        PROB_DS_ValidateProbHalfOpen = True
End Function


Private Function PROB_DS_ValidateBinomialMassInputs( _
    ByVal RawK As Double, _
    ByVal RawN As Double, _
    ByVal ProbSuccess As Double, _
    ByRef KOut As Double, _
    ByRef NOut As Double, _
    ByRef FailMsg As String) _
    As Boolean
'
'==============================================================================
' PROB_DS_ValidateBinomialMassInputs
'------------------------------------------------------------------------------
' PURPOSE
'   Validates Binomial mass inputs in the exact-integer Loader domain.
'
' UPDATED
'   2026-07-19
'==============================================================================
'
        If Not PROB_DS_ValidateTrialsExact(RawN, NOut, FailMsg) Then Exit Function

        If Not PROB_DS_ValidateCount( _
            RawK, _
            PROB_DS_MAX_EXACT_INTEGER, _
            "NumberSuccesses", _
            "the Binomial mass calculation", _
            KOut, _
            FailMsg) Then Exit Function

        If KOut > NOut Then
            FailMsg = "NumberSuccesses must not exceed Trials"
            Exit Function
        End If

        If Not PROB_DS_ValidateProbClosed(ProbSuccess, FailMsg) Then Exit Function

        PROB_DS_ValidateBinomialMassInputs = True
End Function


Private Function PROB_DS_ValidateBinomialKernelInputs( _
    ByVal RawK As Double, _
    ByVal RawN As Double, _
    ByVal ProbSuccess As Double, _
    ByRef KOut As Double, _
    ByRef NOut As Double, _
    ByRef FailMsg As String) _
    As Boolean
'
'==============================================================================
' PROB_DS_ValidateBinomialKernelInputs
'------------------------------------------------------------------------------
' PURPOSE
'   Validates Binomial CDF/SF inputs against the incomplete-beta domain.
'
' UPDATED
'   2026-07-19
'==============================================================================
'
        If Not PROB_DS_ValidateTrialsKernel(RawN, NOut, FailMsg) Then Exit Function

        If Not PROB_DS_ValidateCount( _
            RawK, _
            PROB_DS_MAX_BINOMIAL_KERNEL_N, _
            "NumberSuccesses", _
            "Binomial cumulative and survival calculations", _
            KOut, _
            FailMsg) Then Exit Function

        If KOut > NOut Then
            FailMsg = "NumberSuccesses must not exceed Trials"
            Exit Function
        End If

        If Not PROB_DS_ValidateProbClosed(ProbSuccess, FailMsg) Then Exit Function

        PROB_DS_ValidateBinomialKernelInputs = True
End Function


Private Function PROB_DS_ValidatePoissonPMFInputs( _
    ByVal RawK As Double, _
    ByVal Mean As Double, _
    ByRef KOut As Double, _
    ByRef FailMsg As String) _
    As Boolean
'
'==============================================================================
' PROB_DS_ValidatePoissonPMFInputs
'------------------------------------------------------------------------------
' PURPOSE
'   Validates Poisson PMF inputs in the Loader domain.
'
' UPDATED
'   2026-07-19
'==============================================================================
'
        If Not PROB_DS_ValidateCount( _
            RawK, _
            PROB_DS_MAX_EXACT_INTEGER, _
            "NumberEvents", _
            "the Poisson mass calculation", _
            KOut, _
            FailMsg) Then Exit Function

        If Not PROB_DS_ValidatePoissonMeanPMF(Mean, FailMsg) Then Exit Function

        PROB_DS_ValidatePoissonPMFInputs = True
End Function


Private Function PROB_DS_ValidatePoissonKernelInputs( _
    ByVal RawK As Double, _
    ByVal Mean As Double, _
    ByRef KOut As Double, _
    ByRef FailMsg As String) _
    As Boolean
'
'==============================================================================
' PROB_DS_ValidatePoissonKernelInputs
'------------------------------------------------------------------------------
' PURPOSE
'   Validates Poisson CDF/SF inputs against the incomplete-gamma domain.
'
' UPDATED
'   2026-07-19
'==============================================================================
'
        If Not PROB_DS_ValidateCount( _
            RawK, _
            PROB_DS_MAX_POISSON_KERNEL_COUNT, _
            "NumberEvents", _
            "Poisson cumulative and survival calculations", _
            KOut, _
            FailMsg) Then Exit Function

        If Not PROB_DS_ValidatePoissonMeanKernel(Mean, FailMsg) Then Exit Function

        PROB_DS_ValidatePoissonKernelInputs = True
End Function


Private Function PROB_DS_ValidateGeometricInputs( _
    ByVal RawK As Double, _
    ByVal ProbSuccess As Double, _
    ByRef KOut As Double, _
    ByRef FailMsg As String) _
    As Boolean
'
'==============================================================================
' PROB_DS_ValidateGeometricInputs
'------------------------------------------------------------------------------
' PURPOSE
'   Validates a Geometric failure count and success probability.
'
' UPDATED
'   2026-07-19
'==============================================================================
'
        If Not PROB_DS_ValidateCount( _
            RawK, _
            PROB_DS_MAX_EXACT_INTEGER, _
            "NumberFailures", _
            "Geometric calculations", _
            KOut, _
            FailMsg) Then Exit Function

        If Not PROB_DS_ValidateProbHalfOpen(ProbSuccess, FailMsg) Then Exit Function

        PROB_DS_ValidateGeometricInputs = True
End Function


Private Function PROB_DS_ValidateNegBinomSuccessProb( _
    ByVal RawSuccesses As Double, _
    ByVal ProbSuccess As Double, _
    ByVal MaxSuccesses As Double, _
    ByRef SuccessesOut As Double, _
    ByRef FailMsg As String) _
    As Boolean
'
'==============================================================================
' PROB_DS_ValidateNegBinomSuccessProb
'------------------------------------------------------------------------------
' PURPOSE
'   Validates NumberSuccesses (a count >= 1) and the half-open success
'   probability domain (0, 1] shared by every negative binomial entry point.
'
' UPDATED
'   2026-07-21
'==============================================================================
'
        If Not PROB_DS_ValidateCount(RawSuccesses, MaxSuccesses, _
            "NumberSuccesses", "the negative binomial", SuccessesOut, FailMsg) Then Exit Function

        If SuccessesOut < 1# Then
            FailMsg = "NumberSuccesses must be at least 1"
            Exit Function
        End If

        If Not PROB_DS_ValidateProbHalfOpen(ProbSuccess, FailMsg) Then Exit Function

        PROB_DS_ValidateNegBinomSuccessProb = True
End Function


Private Function PROB_DS_ValidateNegBinomInputs( _
    ByVal RawFailures As Double, _
    ByVal RawSuccesses As Double, _
    ByVal ProbSuccess As Double, _
    ByVal MaxFailures As Double, _
    ByVal MaxSuccesses As Double, _
    ByRef FailuresOut As Double, _
    ByRef SuccessesOut As Double, _
    ByRef FailMsg As String) _
    As Boolean
'
'==============================================================================
' PROB_DS_ValidateNegBinomInputs
'------------------------------------------------------------------------------
' PURPOSE
'   Validates the full (failures, successes, probability) triple and confirms
'   the implied Binomial trial count Failures + Successes stays exact.
'
' UPDATED
'   2026-07-21
'==============================================================================
'
        If Not PROB_DS_ValidateNegBinomSuccessProb(RawSuccesses, ProbSuccess, _
            MaxSuccesses, SuccessesOut, FailMsg) Then Exit Function

        If Not PROB_DS_ValidateCount(RawFailures, MaxFailures, _
            "NumberFailures", "the negative binomial", FailuresOut, FailMsg) Then Exit Function

    'The Loader mass is a Binomial over Failures + Successes trials
        If FailuresOut + SuccessesOut > PROB_DS_MAX_EXACT_INTEGER Then
            FailMsg = "NumberFailures + NumberSuccesses exceeds the exact-integer domain"
            Exit Function
        End If

        PROB_DS_ValidateNegBinomInputs = True
End Function


Private Function PROB_DS_ValidateHypergeometricPopulation( _
    ByVal RawPopulationSize As Double, _
    ByVal RawPopulationSuccesses As Double, _
    ByVal RawSampleSize As Double, _
    ByRef PopulationOut As Double, _
    ByRef SuccessesOut As Double, _
    ByRef SampleOut As Double, _
    ByRef FailMsg As String) _
    As Boolean
'
'==============================================================================
' PROB_DS_ValidateHypergeometricPopulation
'------------------------------------------------------------------------------
' PURPOSE
'   Validates the population triple: PopulationSize >= 1, PopulationSuccesses
'   and SampleSize each in [0, PopulationSize], within the supported ceiling.
'
' UPDATED
'   2026-07-21
'==============================================================================
'
        If Not PROB_DS_ValidateCount(RawPopulationSize, PROB_DS_MAX_HYPERGEOMETRIC_POP, _
            "PopulationSize", "the hypergeometric", PopulationOut, FailMsg) Then Exit Function

        If PopulationOut < 1# Then
            FailMsg = "PopulationSize must be at least 1"
            Exit Function
        End If

        If Not PROB_DS_ValidateCount(RawPopulationSuccesses, PopulationOut, _
            "PopulationSuccesses", "the hypergeometric", SuccessesOut, FailMsg) Then Exit Function

        If Not PROB_DS_ValidateCount(RawSampleSize, PopulationOut, _
            "SampleSize", "the hypergeometric", SampleOut, FailMsg) Then Exit Function

        PROB_DS_ValidateHypergeometricPopulation = True
End Function


Private Function PROB_DS_ValidateHypergeometricInputs( _
    ByVal RawSampleSuccesses As Double, _
    ByVal RawSampleSize As Double, _
    ByVal RawPopulationSuccesses As Double, _
    ByVal RawPopulationSize As Double, _
    ByVal RequireInSupport As Boolean, _
    ByRef SampleSuccessesOut As Double, _
    ByRef SampleOut As Double, _
    ByRef SuccessesOut As Double, _
    ByRef PopulationOut As Double, _
    ByRef FailMsg As String) _
    As Boolean
'
'==============================================================================
' PROB_DS_ValidateHypergeometricInputs
'------------------------------------------------------------------------------
' PURPOSE
'   Validates the full hypergeometric argument set. When RequireInSupport is
'   TRUE the sample-success count must lie in [max(0,n+K-N), min(n,K)], matching
'   the mass domain; the cumulative entry points pass FALSE and clamp instead.
'
' UPDATED
'   2026-07-21
'==============================================================================
'
    Dim KMin                As Double          'Lowest attainable sample success
    Dim KMax                As Double          'Highest attainable sample success

        If Not PROB_DS_ValidateHypergeometricPopulation( _
            RawPopulationSize, RawPopulationSuccesses, RawSampleSize, _
            PopulationOut, SuccessesOut, SampleOut, FailMsg) Then Exit Function

        If Not PROB_DS_ValidateCount(RawSampleSuccesses, SampleOut, _
            "SampleSuccesses", "the hypergeometric", SampleSuccessesOut, FailMsg) Then Exit Function

        If RequireInSupport Then
            KMin = SampleOut + SuccessesOut - PopulationOut
            If KMin < 0# Then KMin = 0#
            KMax = SampleOut
            If SuccessesOut < KMax Then KMax = SuccessesOut

            If SampleSuccessesOut < KMin Or SampleSuccessesOut > KMax Then
                FailMsg = "SampleSuccesses is outside the hypergeometric support"
                Exit Function
            End If
        End If

        PROB_DS_ValidateHypergeometricInputs = True
End Function


'==============================================================================
' PRIVATE MASS KERNELS
'==============================================================================


Private Function PROB_DS_TryDeviancePart( _
    ByVal X As Double, _
    ByVal MeanPart As Double, _
    ByRef Result As Double, _
    ByRef FailMsg As String) _
    As Boolean
'
'==============================================================================
' PROB_DS_TryDeviancePart
'------------------------------------------------------------------------------
' PURPOSE
'   Computes Loader's deviance component
'
'       bd0(X, MeanPart) = X * Log(X / MeanPart) + MeanPart - X
'
'   without cancellation when X is close to MeanPart.
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
            PROB_DS_TryDeviancePart = True
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

            For IterIdx = 1 To PROB_DS_BD0_MAX_ITER
                Ej = Ej * V2
                Term = Ej / (2# * CDbl(IterIdx) + 1#)
                NewSum = SumValue + Term

                If NewSum = SumValue Then
                    Result = NewSum
                    PROB_DS_TryDeviancePart = True
                    Exit Function
                End If

                ScaleValue = Abs(NewSum)
                If ScaleValue < 1# Then ScaleValue = 1#

                If Abs(Term) <= PROB_MACH_EPS * ScaleValue Then
                    Result = NewSum
                    PROB_DS_TryDeviancePart = True
                    Exit Function
                End If

                SumValue = NewSum
            Next IterIdx

            FailMsg = "Loader deviance series failed to converge in " & _
                      PROB_DS_BD0_MAX_ITER & " iterations"
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

        PROB_DS_TryDeviancePart = True
End Function


Private Function PROB_DS_TryBinomialLogMass( _
    ByVal K As Double, _
    ByVal n As Double, _
    ByVal ProbSuccess As Double, _
    ByRef LogMass As Double, _
    ByRef IsCertainZero As Boolean, _
    ByRef FailMsg As String) _
    As Boolean
'
'==============================================================================
' PROB_DS_TryBinomialLogMass
'------------------------------------------------------------------------------
' PURPOSE
'   Computes the natural log of the Binomial mass through Loader's stable
'   Stirling/deviance arrangement, without exponentiating.
'
'   IsCertainZero is returned True when the outcome has probability exactly
'   zero or a log-mass below the representable Double range; in both cases the
'   mass is zero and the log-mass is undefined.
'
'   PRECONDITION
'     0 <= K <= N, N <= 2^53 - 1 and ProbSuccess in [0, 1].
'
' UPDATED
'   2026-07-19 - Extracted shared log-mass; PMF now exponentiates it.
'==============================================================================
'
    Dim Q                   As Double          'Failure probability 1-p
    Dim j                   As Double          'Failure count N-K
    Dim Np                  As Double          'N * p
    Dim Nq                  As Double          'N * q
    Dim DevianceK           As Double          'bd0(K, Np)
    Dim DevianceJ           As Double          'bd0(J, Nq)
    Dim LogP                As Double          'Accurate logarithm of p

'------------------------------------------------------------------------------
' INITIALIZE
'------------------------------------------------------------------------------
        IsCertainZero = False

'------------------------------------------------------------------------------
' HANDLE DEGENERATE PROBABILITIES
'------------------------------------------------------------------------------
        If ProbSuccess <= 0# Then
            If K = 0# Then LogMass = 0# Else IsCertainZero = True
            PROB_DS_TryBinomialLogMass = True
            Exit Function
        End If

        If ProbSuccess >= 1# Then
            If K = n Then LogMass = 0# Else IsCertainZero = True
            PROB_DS_TryBinomialLogMass = True
            Exit Function
        End If

        Q = 1# - ProbSuccess

'------------------------------------------------------------------------------
' HANDLE SUPPORT EDGES
'------------------------------------------------------------------------------
    'At K = 0, Log(PMF) = N * Log(1-p); Log1p keeps small p accurate
        If K <= 0# Then
            If Not PROB_TryMultiply(n, PROB_Log1p(-ProbSuccess), LogMass) Then
                IsCertainZero = True
            End If
            PROB_DS_TryBinomialLogMass = True
            Exit Function
        End If

    'At K = N, use Log1p when p is near one
        If K >= n Then
            If ProbSuccess >= 0.5 Then
                LogP = PROB_Log1p(-Q)
            Else
                LogP = Log(ProbSuccess)
            End If

            If Not PROB_TryMultiply(n, LogP, LogMass) Then
                IsCertainZero = True
            End If
            PROB_DS_TryBinomialLogMass = True
            Exit Function
        End If

'------------------------------------------------------------------------------
' LOADER INTERIOR ARRANGEMENT
'------------------------------------------------------------------------------
        j = n - K

    'Form the two expected counts
        If Not PROB_TryMultiply(n, ProbSuccess, Np) Then
            FailMsg = "Binomial expected-success count overflowed"
            Exit Function
        End If

        If Not PROB_TryMultiply(n, Q, Nq) Then
            FailMsg = "Binomial expected-failure count overflowed"
            Exit Function
        End If

    'Compute the two deviance components
        If Not PROB_DS_TryDeviancePart(K, Np, DevianceK, FailMsg) Then Exit Function
        If Not PROB_DS_TryDeviancePart(j, Nq, DevianceJ, FailMsg) Then Exit Function

    'Assemble Loader's log mass without subtracting large log-gammas
        LogMass = _
            PROB_StirlingError(n) - _
            PROB_StirlingError(K) - _
            PROB_StirlingError(j) - _
            DevianceK - _
            DevianceJ - _
            PROB_HALF_LOG_TWO_PI - _
            0.5 * (Log(K) + Log(j) - Log(n))

        PROB_DS_TryBinomialLogMass = True
End Function


Private Function PROB_DS_TryBinomialPMF( _
    ByVal K As Double, _
    ByVal n As Double, _
    ByVal ProbSuccess As Double, _
    ByRef Result As Double, _
    ByRef FailMsg As String) _
    As Boolean
'
'==============================================================================
' PROB_DS_TryBinomialPMF
'------------------------------------------------------------------------------
' PURPOSE
'   Computes the Binomial mass by exponentiating the shared log-mass kernel.
'   A certain-zero outcome or deep-tail underflow is a valid zero.
'
' UPDATED
'   2026-07-19 - Now a thin exponentiation over PROB_DS_TryBinomialLogMass.
'==============================================================================
'
    Dim LogMass             As Double          'Natural logarithm of the mass
    Dim IsCertainZero       As Boolean         'Outcome is impossible or below Double range

        If Not PROB_DS_TryBinomialLogMass( _
            K, n, ProbSuccess, LogMass, IsCertainZero, FailMsg) Then Exit Function

        If IsCertainZero Then
            Result = 0#
            PROB_DS_TryBinomialPMF = True
            Exit Function
        End If

    'The true mass cannot overflow; negative underflow is a valid zero
        If Not PROB_TryExp(LogMass, Result) Then
            FailMsg = "Binomial mass exponentiation failed"
            Exit Function
        End If

    'Clamp harmless final round-off
        If Result < 0# Then Result = 0#
        If Result > 1# Then Result = 1#

        PROB_DS_TryBinomialPMF = True
End Function


Private Function PROB_DS_TryPoissonLogMass( _
    ByVal K As Double, _
    ByVal Mean As Double, _
    ByRef LogMass As Double, _
    ByRef IsCertainZero As Boolean, _
    ByRef FailMsg As String) _
    As Boolean
'
'==============================================================================
' PROB_DS_TryPoissonLogMass
'------------------------------------------------------------------------------
' PURPOSE
'   Computes the natural log of the Poisson mass through Loader's stable
'   Stirling/deviance arrangement, without exponentiating.
'
'   IsCertainZero is returned True only for the impossible outcome under a
'   zero mean (K > 0 with Mean = 0).
'
'   PRECONDITION
'     K is an exact non-negative integer and Mean is in [0, 2^53 - 1].
'
' UPDATED
'   2026-07-19 - Extracted shared log-mass; PMF now exponentiates it.
'==============================================================================
'
    Dim Deviance            As Double          'bd0(K, Mean)

'------------------------------------------------------------------------------
' INITIALIZE
'------------------------------------------------------------------------------
        IsCertainZero = False

'------------------------------------------------------------------------------
' HANDLE ZERO MEAN
'------------------------------------------------------------------------------
        If Mean <= 0# Then
            If K = 0# Then LogMass = 0# Else IsCertainZero = True
            PROB_DS_TryPoissonLogMass = True
            Exit Function
        End If

'------------------------------------------------------------------------------
' HANDLE ZERO COUNT
'------------------------------------------------------------------------------
    'At K = 0, Log(PMF) = -Mean
        If K <= 0# Then
            LogMass = -Mean
            PROB_DS_TryPoissonLogMass = True
            Exit Function
        End If

'------------------------------------------------------------------------------
' LOADER ARRANGEMENT
'------------------------------------------------------------------------------
        If Not PROB_DS_TryDeviancePart(K, Mean, Deviance, FailMsg) Then Exit Function

        LogMass = _
            -PROB_StirlingError(K) - _
            Deviance - _
            PROB_HALF_LOG_TWO_PI - _
            0.5 * Log(K)

        PROB_DS_TryPoissonLogMass = True
End Function


Private Function PROB_DS_TryPoissonPMF( _
    ByVal K As Double, _
    ByVal Mean As Double, _
    ByRef Result As Double, _
    ByRef FailMsg As String) _
    As Boolean
'
'==============================================================================
' PROB_DS_TryPoissonPMF
'------------------------------------------------------------------------------
' PURPOSE
'   Computes the Poisson mass by exponentiating the shared log-mass kernel.
'   A certain-zero outcome or deep-tail underflow is a valid zero.
'
' UPDATED
'   2026-07-19 - Now a thin exponentiation over PROB_DS_TryPoissonLogMass.
'==============================================================================
'
    Dim LogMass             As Double          'Natural logarithm of the mass
    Dim IsCertainZero       As Boolean         'Outcome is impossible under a zero mean

        If Not PROB_DS_TryPoissonLogMass( _
            K, Mean, LogMass, IsCertainZero, FailMsg) Then Exit Function

        If IsCertainZero Then
            Result = 0#
            PROB_DS_TryPoissonPMF = True
            Exit Function
        End If

        If Not PROB_TryExp(LogMass, Result) Then
            FailMsg = "Poisson mass exponentiation failed"
            Exit Function
        End If

        If Result < 0# Then Result = 0#
        If Result > 1# Then Result = 1#

        PROB_DS_TryPoissonPMF = True
End Function


Private Function PROB_DS_TryNegativeBinomialLogMass( _
    ByVal X As Double, _
    ByVal R As Double, _
    ByVal ProbSuccess As Double, _
    ByRef LogMass As Double, _
    ByRef IsCertainZero As Boolean, _
    ByRef FailMsg As String) _
    As Boolean
'
'==============================================================================
' PROB_DS_TryNegativeBinomialLogMass
'------------------------------------------------------------------------------
' PURPOSE
'   Natural log of the negative binomial mass P(X = x | r, p).
'   Uses the identity P(x) = [r/(r+x)] * Binom(r; r+x, p), so the log mass is
'   the shared Binomial log-mass plus Log(r) - Log(r+x). Reusing the frozen
'   Loader kernel keeps deep-tail accuracy where the ordinary mass underflows.
'
' UPDATED
'   2026-07-21
'==============================================================================
'
    Dim BinLogMass          As Double          'Binomial log-mass at (r; r+x)
    Dim BinZero             As Boolean         'Binomial certain-zero flag

        IsCertainZero = False

        If Not PROB_DS_TryBinomialLogMass( _
            R, R + X, ProbSuccess, BinLogMass, BinZero, FailMsg) Then Exit Function

        If BinZero Then
            IsCertainZero = True
            PROB_DS_TryNegativeBinomialLogMass = True
            Exit Function
        End If

        LogMass = BinLogMass + Log(R) - Log(R + X)

        PROB_DS_TryNegativeBinomialLogMass = True
End Function


Private Function PROB_DS_TryNegativeBinomialPMF( _
    ByVal X As Double, _
    ByVal R As Double, _
    ByVal ProbSuccess As Double, _
    ByRef Result As Double, _
    ByRef FailMsg As String) _
    As Boolean
'
'==============================================================================
' PROB_DS_TryNegativeBinomialPMF
'------------------------------------------------------------------------------
' PURPOSE
'   Negative binomial mass by exponentiating the shared log-mass kernel.
'   A certain-zero outcome or deep-tail underflow is a valid zero.
'
' UPDATED
'   2026-07-21
'==============================================================================
'
    Dim LogMass             As Double          'Natural logarithm of the mass
    Dim IsCertainZero       As Boolean         'Impossible or below Double range

        If Not PROB_DS_TryNegativeBinomialLogMass( _
            X, R, ProbSuccess, LogMass, IsCertainZero, FailMsg) Then Exit Function

        If IsCertainZero Then
            Result = 0#
            PROB_DS_TryNegativeBinomialPMF = True
            Exit Function
        End If

        If Not PROB_TryExp(LogMass, Result) Then
            FailMsg = "Negative binomial mass exponentiation failed"
            Exit Function
        End If

        If Result < 0# Then Result = 0#
        If Result > 1# Then Result = 1#

        PROB_DS_TryNegativeBinomialPMF = True
End Function


Private Function PROB_DS_TryHypergeometricLogMass( _
    ByVal K As Double, _
    ByVal n As Double, _
    ByVal Kp As Double, _
    ByVal Np As Double, _
    ByRef LogMass As Double, _
    ByRef IsCertainZero As Boolean, _
    ByRef FailMsg As String) _
    As Boolean
'
'==============================================================================
' PROB_DS_TryHypergeometricLogMass
'------------------------------------------------------------------------------
' PURPOSE
'   Natural log of the hypergeometric mass via Loader's factorisation
'     dhyper = dbinom(k;K,p) * dbinom(n-k;N-K,p) / dbinom(n;N,p),  p = n/N.
'   The (n/N) and complement powers cancel exactly, leaving the true log mass;
'   an impossible numerator term marks the outcome certain-zero.
'
' UPDATED
'   2026-07-21
'==============================================================================
'
    Dim P                   As Double          'Sampling fraction n/N
    Dim B1                  As Double          'Log dbinom(k; K, p)
    Dim B2                  As Double          'Log dbinom(n-k; N-K, p)
    Dim B3                  As Double          'Log dbinom(n; N, p)
    Dim Z1                  As Boolean         'Numerator term one certain-zero
    Dim Z2                  As Boolean         'Numerator term two certain-zero
    Dim Z3                  As Boolean         'Denominator certain-zero

        IsCertainZero = False
        P = n / Np

        If Not PROB_DS_TryBinomialLogMass(K, Kp, P, B1, Z1, FailMsg) Then Exit Function
        If Not PROB_DS_TryBinomialLogMass(n - K, Np - Kp, P, B2, Z2, FailMsg) Then Exit Function
        If Not PROB_DS_TryBinomialLogMass(n, Np, P, B3, Z3, FailMsg) Then Exit Function

    'An impossible numerator term means k is outside the support
        If Z1 Or Z2 Then
            IsCertainZero = True
            PROB_DS_TryHypergeometricLogMass = True
            Exit Function
        End If

        LogMass = B1 + B2 - B3

        PROB_DS_TryHypergeometricLogMass = True
End Function


Private Function PROB_DS_TryHypergeometricPMF( _
    ByVal K As Double, _
    ByVal n As Double, _
    ByVal Kp As Double, _
    ByVal Np As Double, _
    ByRef Result As Double, _
    ByRef FailMsg As String) _
    As Boolean
'
'==============================================================================
' PROB_DS_TryHypergeometricPMF
'------------------------------------------------------------------------------
' PURPOSE
'   Hypergeometric mass by exponentiating the shared log-mass kernel.
'   A certain-zero outcome or deep-tail underflow is a valid zero.
'
' UPDATED
'   2026-07-21
'==============================================================================
'
    Dim LogMass             As Double          'Natural logarithm of the mass
    Dim IsCertainZero       As Boolean         'Impossible or below Double range

        If Not PROB_DS_TryHypergeometricLogMass( _
            K, n, Kp, Np, LogMass, IsCertainZero, FailMsg) Then Exit Function

        If IsCertainZero Then
            Result = 0#
            PROB_DS_TryHypergeometricPMF = True
            Exit Function
        End If

        If Not PROB_TryExp(LogMass, Result) Then
            FailMsg = "Hypergeometric mass exponentiation failed"
            Exit Function
        End If

        If Result < 0# Then Result = 0#
        If Result > 1# Then Result = 1#

        PROB_DS_TryHypergeometricPMF = True
End Function


'==============================================================================
' PRIVATE CDF AND SURVIVAL KERNELS
'==============================================================================


Private Function PROB_DS_TryBinomialCDF( _
    ByVal K As Double, _
    ByVal n As Double, _
    ByVal ProbSuccess As Double, _
    ByRef Result As Double, _
    ByRef FailMsg As String) _
    As Boolean
'
'==============================================================================
' PROB_DS_TryBinomialCDF
'------------------------------------------------------------------------------
' PURPOSE
'   Computes P(X <= K) = I(1-p; N-K, K+1).
'
' PRECONDITION
'   Validated Binomial kernel inputs.
'
' UPDATED
'   2026-07-19
'==============================================================================
'
        If ProbSuccess <= 0# Then
            Result = 1#
            PROB_DS_TryBinomialCDF = True
            Exit Function
        End If

        If ProbSuccess >= 1# Then
            If K >= n Then Result = 1# Else Result = 0#
            PROB_DS_TryBinomialCDF = True
            Exit Function
        End If

        If K >= n Then
            Result = 1#
            PROB_DS_TryBinomialCDF = True
            Exit Function
        End If

        PROB_DS_TryBinomialCDF = PROB_TryBetaRegularized( _
            1# - ProbSuccess, _
            ProbSuccess, _
            n - K, _
            K + 1#, _
            Result, _
            FailMsg)
End Function


Private Function PROB_DS_TryBinomialSF( _
    ByVal K As Double, _
    ByVal n As Double, _
    ByVal ProbSuccess As Double, _
    ByRef Result As Double, _
    ByRef FailMsg As String) _
    As Boolean
'
'==============================================================================
' PROB_DS_TryBinomialSF
'------------------------------------------------------------------------------
' PURPOSE
'   Computes P(X > K) = I(p; K+1, N-K) directly.
'
' PRECONDITION
'   Validated Binomial kernel inputs.
'
' UPDATED
'   2026-07-19
'==============================================================================
'
        If ProbSuccess <= 0# Then
            Result = 0#
            PROB_DS_TryBinomialSF = True
            Exit Function
        End If

        If ProbSuccess >= 1# Then
            If K < n Then Result = 1# Else Result = 0#
            PROB_DS_TryBinomialSF = True
            Exit Function
        End If

        If K >= n Then
            Result = 0#
            PROB_DS_TryBinomialSF = True
            Exit Function
        End If

        PROB_DS_TryBinomialSF = PROB_TryBetaRegularized( _
            ProbSuccess, _
            1# - ProbSuccess, _
            K + 1#, _
            n - K, _
            Result, _
            FailMsg)
End Function


Private Function PROB_DS_TryPoissonCDF( _
    ByVal K As Double, _
    ByVal Mean As Double, _
    ByRef Result As Double, _
    ByRef FailMsg As String) _
    As Boolean
'
'==============================================================================
' PROB_DS_TryPoissonCDF
'------------------------------------------------------------------------------
' PURPOSE
'   Computes P(X <= K) = Q(K+1, Mean).
'
' UPDATED
'   2026-07-19
'==============================================================================
'
        If Mean <= 0# Then
            Result = 1#
            PROB_DS_TryPoissonCDF = True
            Exit Function
        End If

        PROB_DS_TryPoissonCDF = PROB_TryGammaRegularizedQ( _
            K + 1#, Mean, Result, FailMsg)
End Function


Private Function PROB_DS_TryPoissonSF( _
    ByVal K As Double, _
    ByVal Mean As Double, _
    ByRef Result As Double, _
    ByRef FailMsg As String) _
    As Boolean
'
'==============================================================================
' PROB_DS_TryPoissonSF
'------------------------------------------------------------------------------
' PURPOSE
'   Computes P(X > K) = P(K+1, Mean) directly.
'
' UPDATED
'   2026-07-19
'==============================================================================
'
        If Mean <= 0# Then
            Result = 0#
            PROB_DS_TryPoissonSF = True
            Exit Function
        End If

        PROB_DS_TryPoissonSF = PROB_TryGammaRegularizedP( _
            K + 1#, Mean, Result, FailMsg)
End Function


Private Function PROB_DS_TryGeometricLogMass( _
    ByVal K As Double, _
    ByVal ProbSuccess As Double, _
    ByRef LogMass As Double, _
    ByRef IsCertainZero As Boolean, _
    ByRef FailMsg As String) _
    As Boolean
'
'==============================================================================
' PROB_DS_TryGeometricLogMass
'------------------------------------------------------------------------------
' PURPOSE
'   Computes Log(p) + K * Log(1-p), the natural log of the Geometric mass,
'   with guarded logarithmic multiplication and addition.
'
'   IsCertainZero is returned True for the impossible outcome under certain
'   success (K > 0 with p = 1) or when the log-mass falls below Double range.
'
' UPDATED
'   2026-07-19 - Extracted shared log-mass; PMF now exponentiates it.
'==============================================================================
'
    Dim LogPower            As Double          'K * Log(1-p)

'------------------------------------------------------------------------------
' INITIALIZE
'------------------------------------------------------------------------------
        IsCertainZero = False

'------------------------------------------------------------------------------
' HANDLE CERTAIN SUCCESS
'------------------------------------------------------------------------------
        If ProbSuccess >= 1# Then
            If K = 0# Then LogMass = 0# Else IsCertainZero = True
            PROB_DS_TryGeometricLogMass = True
            Exit Function
        End If

'------------------------------------------------------------------------------
' ASSEMBLE LOG MASS
'------------------------------------------------------------------------------
    'A negative product overflow means the power is certainly zero
        If Not PROB_TryMultiply(K, PROB_Log1p(-ProbSuccess), LogPower) Then
            IsCertainZero = True
            PROB_DS_TryGeometricLogMass = True
            Exit Function
        End If

    'A negative addition overflow likewise implies a zero mass
        If Not PROB_TryAdd(Log(ProbSuccess), LogPower, LogMass) Then
            IsCertainZero = True
            PROB_DS_TryGeometricLogMass = True
            Exit Function
        End If

        PROB_DS_TryGeometricLogMass = True
End Function


Private Function PROB_DS_TryGeometricPMF( _
    ByVal K As Double, _
    ByVal ProbSuccess As Double, _
    ByRef Result As Double, _
    ByRef FailMsg As String) _
    As Boolean
'
'==============================================================================
' PROB_DS_TryGeometricPMF
'------------------------------------------------------------------------------
' PURPOSE
'   Computes p*(1-p)^K by exponentiating the shared log-mass kernel.
'   A certain-zero outcome or deep-tail underflow is a valid zero.
'
' UPDATED
'   2026-07-19 - Now a thin exponentiation over PROB_DS_TryGeometricLogMass.
'==============================================================================
'
    Dim LogMass             As Double          'Natural logarithm of the mass
    Dim IsCertainZero       As Boolean         'Outcome is impossible or below Double range

        If Not PROB_DS_TryGeometricLogMass( _
            K, ProbSuccess, LogMass, IsCertainZero, FailMsg) Then Exit Function

        If IsCertainZero Then
            Result = 0#
            PROB_DS_TryGeometricPMF = True
            Exit Function
        End If

        If Not PROB_TryExp(LogMass, Result) Then
            FailMsg = "Geometric mass exponentiation failed"
            Exit Function
        End If

        If Result < 0# Then Result = 0#
        If Result > 1# Then Result = 1#

        PROB_DS_TryGeometricPMF = True
End Function


Private Function PROB_DS_TryGeometricCDF( _
    ByVal K As Double, _
    ByVal ProbSuccess As Double, _
    ByRef Result As Double, _
    ByRef FailMsg As String) _
    As Boolean
'
'==============================================================================
' PROB_DS_TryGeometricCDF
'------------------------------------------------------------------------------
' PURPOSE
'   Computes 1-(1-p)^(K+1) through a guarded negative exponent and Expm1.
'
' UPDATED
'   2026-07-19
'==============================================================================
'
    Dim LogTail             As Double          '(K+1) * Log(1-p)

        If ProbSuccess >= 1# Then
            Result = 1#
            PROB_DS_TryGeometricCDF = True
            Exit Function
        End If

    'A negative product overflow means the CDF has reached one
        If Not PROB_TryMultiply(K + 1#, PROB_Log1p(-ProbSuccess), LogTail) Then
            Result = 1#
            PROB_DS_TryGeometricCDF = True
            Exit Function
        End If

        Result = -PROB_Expm1(LogTail)

        If Result < 0# Then Result = 0#
        If Result > 1# Then Result = 1#

        PROB_DS_TryGeometricCDF = True
End Function


Private Function PROB_DS_TryGeometricSF( _
    ByVal K As Double, _
    ByVal ProbSuccess As Double, _
    ByRef Result As Double, _
    ByRef FailMsg As String) _
    As Boolean
'
'==============================================================================
' PROB_DS_TryGeometricSF
'------------------------------------------------------------------------------
' PURPOSE
'   Computes (1-p)^(K+1) directly through a guarded negative exponent.
'
' UPDATED
'   2026-07-19
'==============================================================================
'
    Dim LogTail             As Double          '(K+1) * Log(1-p)

        If ProbSuccess >= 1# Then
            Result = 0#
            PROB_DS_TryGeometricSF = True
            Exit Function
        End If

    'A negative product overflow means the survival probability is zero
        If Not PROB_TryMultiply(K + 1#, PROB_Log1p(-ProbSuccess), LogTail) Then
            Result = 0#
            PROB_DS_TryGeometricSF = True
            Exit Function
        End If

        If Not PROB_TryExp(LogTail, Result) Then
            FailMsg = "Geometric survival exponentiation failed"
            Exit Function
        End If

        If Result < 0# Then Result = 0#
        If Result > 1# Then Result = 1#

        PROB_DS_TryGeometricSF = True
End Function


Private Function PROB_DS_TryNegativeBinomialCDF( _
    ByVal X As Double, _
    ByVal R As Double, _
    ByVal ProbSuccess As Double, _
    ByRef Result As Double, _
    ByRef FailMsg As String) _
    As Boolean
'
'==============================================================================
' PROB_DS_TryNegativeBinomialCDF
'------------------------------------------------------------------------------
' PURPOSE
'   P(X <= x) = I_p(r, x + 1) via the two-argument regularized incomplete beta.
'
' UPDATED
'   2026-07-21
'==============================================================================
'
    'Certain success places all mass at zero failures
        If ProbSuccess >= 1# Then
            Result = 1#
            PROB_DS_TryNegativeBinomialCDF = True
            Exit Function
        End If

        PROB_DS_TryNegativeBinomialCDF = PROB_TryBetaRegularized( _
            ProbSuccess, _
            1# - ProbSuccess, _
            R, _
            X + 1#, _
            Result, _
            FailMsg)
End Function


Private Function PROB_DS_TryNegativeBinomialSF( _
    ByVal X As Double, _
    ByVal R As Double, _
    ByVal ProbSuccess As Double, _
    ByRef Result As Double, _
    ByRef FailMsg As String) _
    As Boolean
'
'==============================================================================
' PROB_DS_TryNegativeBinomialSF
'------------------------------------------------------------------------------
' PURPOSE
'   P(X > x) = I_(1-p)(x + 1, r), the small upper tail computed with no cancellation.
'
' UPDATED
'   2026-07-21
'==============================================================================
'
    'Certain success leaves no mass above zero failures
        If ProbSuccess >= 1# Then
            Result = 0#
            PROB_DS_TryNegativeBinomialSF = True
            Exit Function
        End If

        PROB_DS_TryNegativeBinomialSF = PROB_TryBetaRegularized( _
            1# - ProbSuccess, _
            ProbSuccess, _
            X + 1#, _
            R, _
            Result, _
            FailMsg)
End Function


Private Function PROB_DS_TryHypergeometricLowerTail( _
    ByVal Anchor As Double, _
    ByVal KMin As Double, _
    ByVal n As Double, _
    ByVal Kp As Double, _
    ByVal Np As Double, _
    ByRef Result As Double, _
    ByRef FailMsg As String) _
    As Boolean
'
'==============================================================================
' PROB_DS_TryHypergeometricLowerTail
'------------------------------------------------------------------------------
' PURPOSE
'   P(X <= Anchor) summed downward from the anchor using the exact successive
'   mass ratio d(j-1)/d(j). Efficient and cancellation-free when Anchor is at
'   or below the mean: terms shrink monotonically into the lower tail.
'
' UPDATED
'   2026-07-21
'==============================================================================
'
    Dim Mass                As Double          'Mass at the anchor point
    Dim Term                As Double          'Current running term
    Dim Sum                 As Double          'Accumulated tail probability
    Dim Previous            As Double          'Sum before adding a term
    Dim Ratio               As Double          'd(j-1)/d(j)
    Dim j                   As Double          'Descending support index
    Dim IterIdx             As Long            'Iteration guard

        If Anchor < KMin Then
            Result = 0#
            PROB_DS_TryHypergeometricLowerTail = True
            Exit Function
        End If

        If Not PROB_DS_TryHypergeometricPMF(Anchor, n, Kp, Np, Mass, FailMsg) Then Exit Function

        Sum = Mass
        Term = Mass
        j = Anchor

        For IterIdx = 1 To PROB_DS_MAX_HYPERGEOMETRIC_SUM_ITER
            If j <= KMin Then Exit For

        'Ratio d(j-1)/d(j) = [ j (N-K-n+j) ] / [ (K-j+1)(n-j+1) ]
            Ratio = (j * (Np - Kp - n + j)) / ((Kp - j + 1#) * (n - j + 1#))
            Term = Term * Ratio
            j = j - 1#

            Previous = Sum
            Sum = Sum + Term
            If Sum = Previous Then Exit For
        Next IterIdx

        If Sum > 1# Then Sum = 1#
        Result = Sum
        PROB_DS_TryHypergeometricLowerTail = True
End Function


Private Function PROB_DS_TryHypergeometricUpperTail( _
    ByVal Anchor As Double, _
    ByVal KMax As Double, _
    ByVal n As Double, _
    ByVal Kp As Double, _
    ByVal Np As Double, _
    ByRef Result As Double, _
    ByRef FailMsg As String) _
    As Boolean
'
'==============================================================================
' PROB_DS_TryHypergeometricUpperTail
'------------------------------------------------------------------------------
' PURPOSE
'   P(X >= Anchor) summed upward from the anchor using the exact successive
'   mass ratio d(j+1)/d(j). Efficient and cancellation-free when Anchor is at
'   or above the mean: terms shrink monotonically into the upper tail.
'
' UPDATED
'   2026-07-21
'==============================================================================
'
    Dim Mass                As Double          'Mass at the anchor point
    Dim Term                As Double          'Current running term
    Dim Sum                 As Double          'Accumulated tail probability
    Dim Previous            As Double          'Sum before adding a term
    Dim Ratio               As Double          'd(j+1)/d(j)
    Dim j                   As Double          'Ascending support index
    Dim IterIdx             As Long            'Iteration guard

        If Anchor > KMax Then
            Result = 0#
            PROB_DS_TryHypergeometricUpperTail = True
            Exit Function
        End If

        If Not PROB_DS_TryHypergeometricPMF(Anchor, n, Kp, Np, Mass, FailMsg) Then Exit Function

        Sum = Mass
        Term = Mass
        j = Anchor

        For IterIdx = 1 To PROB_DS_MAX_HYPERGEOMETRIC_SUM_ITER
            If j >= KMax Then Exit For

        'Ratio d(j+1)/d(j) = [ (K-j)(n-j) ] / [ (j+1)(N-K-n+j+1) ]
            Ratio = ((Kp - j) * (n - j)) / ((j + 1#) * (Np - Kp - n + j + 1#))
            Term = Term * Ratio
            j = j + 1#

            Previous = Sum
            Sum = Sum + Term
            If Sum = Previous Then Exit For
        Next IterIdx

        If Sum > 1# Then Sum = 1#
        Result = Sum
        PROB_DS_TryHypergeometricUpperTail = True
End Function


Private Function PROB_DS_TryHypergeometricCDF( _
    ByVal K As Double, _
    ByVal n As Double, _
    ByVal Kp As Double, _
    ByVal Np As Double, _
    ByRef Result As Double, _
    ByRef FailMsg As String) _
    As Boolean
'
'==============================================================================
' PROB_DS_TryHypergeometricCDF
'------------------------------------------------------------------------------
' PURPOSE
'   P(X <= k). Sums the lower tail directly when k is at or below the mean,
'   otherwise returns 1 minus the small upper tail, so the near tail always
'   drives accuracy and only a few standard deviations of terms are needed.
'
' UPDATED
'   2026-07-21
'==============================================================================
'
    Dim KMin                As Double          'Lowest attainable count
    Dim KMax                As Double          'Highest attainable count
    Dim Mu                  As Double          'Mean n*K/N
    Dim UpperTail           As Double          'P(X >= k+1)

        KMin = n + Kp - Np
        If KMin < 0# Then KMin = 0#
        KMax = n
        If Kp < KMax Then KMax = Kp
        Mu = n * Kp / Np

        If K >= KMax Then
            Result = 1#
            PROB_DS_TryHypergeometricCDF = True
            Exit Function
        End If

        If K < KMin Then
            Result = 0#
            PROB_DS_TryHypergeometricCDF = True
            Exit Function
        End If

        If K <= Mu Then
            PROB_DS_TryHypergeometricCDF = PROB_DS_TryHypergeometricLowerTail( _
                K, KMin, n, Kp, Np, Result, FailMsg)
        Else
            If Not PROB_DS_TryHypergeometricUpperTail( _
                K + 1#, KMax, n, Kp, Np, UpperTail, FailMsg) Then Exit Function
            Result = 1# - UpperTail
            PROB_DS_TryHypergeometricCDF = True
        End If
End Function


Private Function PROB_DS_TryHypergeometricSF( _
    ByVal K As Double, _
    ByVal n As Double, _
    ByVal Kp As Double, _
    ByVal Np As Double, _
    ByRef Result As Double, _
    ByRef FailMsg As String) _
    As Boolean
'
'==============================================================================
' PROB_DS_TryHypergeometricSF
'------------------------------------------------------------------------------
' PURPOSE
'   P(X > k). Sums the upper tail directly when k is at or above the mean,
'   otherwise returns 1 minus the small lower tail.
'
' UPDATED
'   2026-07-21
'==============================================================================
'
    Dim KMin                As Double          'Lowest attainable count
    Dim KMax                As Double          'Highest attainable count
    Dim Mu                  As Double          'Mean n*K/N
    Dim LowerTail           As Double          'P(X <= k)

        KMin = n + Kp - Np
        If KMin < 0# Then KMin = 0#
        KMax = n
        If Kp < KMax Then KMax = Kp
        Mu = n * Kp / Np

        If K >= KMax Then
            Result = 0#
            PROB_DS_TryHypergeometricSF = True
            Exit Function
        End If

        If K < KMin Then
            Result = 1#
            PROB_DS_TryHypergeometricSF = True
            Exit Function
        End If

        If K >= Mu Then
            PROB_DS_TryHypergeometricSF = PROB_DS_TryHypergeometricUpperTail( _
                K + 1#, KMax, n, Kp, Np, Result, FailMsg)
        Else
            If Not PROB_DS_TryHypergeometricLowerTail( _
                K, KMin, n, Kp, Np, LowerTail, FailMsg) Then Exit Function
            Result = 1# - LowerTail
            PROB_DS_TryHypergeometricSF = True
        End If
End Function


'==============================================================================
' PRIVATE INVERSE KERNELS
'==============================================================================


Private Function PROB_DS_TryBinomialQualifies( _
    ByVal K As Double, _
    ByVal Probability As Double, _
    ByVal ComplementProbability As Double, _
    ByVal n As Double, _
    ByVal ProbSuccess As Double, _
    ByRef Qualifies As Boolean, _
    ByRef FailMsg As String) _
    As Boolean
'
'==============================================================================
' PROB_DS_TryBinomialQualifies
'------------------------------------------------------------------------------
' PURPOSE
'   Tests F(K) >= Probability using the smaller direct probability tail.
'
' UPDATED
'   2026-07-19
'==============================================================================
'
    Dim TailValue           As Double          'CDF or SF at K

        If Probability <= 0.5 Then
            If Not PROB_DS_TryBinomialCDF( _
                K, n, ProbSuccess, TailValue, FailMsg) Then Exit Function

            Qualifies = (TailValue >= Probability)
        Else
            If Not PROB_DS_TryBinomialSF( _
                K, n, ProbSuccess, TailValue, FailMsg) Then Exit Function

            Qualifies = (TailValue <= ComplementProbability)
        End If

        PROB_DS_TryBinomialQualifies = True
End Function


Private Function PROB_DS_TryPoissonQualifies( _
    ByVal K As Double, _
    ByVal Probability As Double, _
    ByVal ComplementProbability As Double, _
    ByVal Mean As Double, _
    ByRef Qualifies As Boolean, _
    ByRef FailMsg As String) _
    As Boolean
'
'==============================================================================
' PROB_DS_TryPoissonQualifies
'------------------------------------------------------------------------------
' PURPOSE
'   Tests F(K) >= Probability using the smaller direct probability tail.
'
' UPDATED
'   2026-07-19
'==============================================================================
'
    Dim TailValue           As Double          'CDF or SF at K

        If Probability <= 0.5 Then
            If Not PROB_DS_TryPoissonCDF( _
                K, Mean, TailValue, FailMsg) Then Exit Function

            Qualifies = (TailValue >= Probability)
        Else
            If Not PROB_DS_TryPoissonSF( _
                K, Mean, TailValue, FailMsg) Then Exit Function

            Qualifies = (TailValue <= ComplementProbability)
        End If

        PROB_DS_TryPoissonQualifies = True
End Function


Private Function PROB_DS_TryGeometricQualifies( _
    ByVal K As Double, _
    ByVal Probability As Double, _
    ByVal ComplementProbability As Double, _
    ByVal ProbSuccess As Double, _
    ByRef Qualifies As Boolean, _
    ByRef FailMsg As String) _
    As Boolean
'
'==============================================================================
' PROB_DS_TryGeometricQualifies
'------------------------------------------------------------------------------
' PURPOSE
'   Tests F(K) >= Probability using the smaller direct probability tail.
'
' UPDATED
'   2026-07-19
'==============================================================================
'
    Dim TailValue           As Double          'CDF or SF at K

        If Probability <= 0.5 Then
            If Not PROB_DS_TryGeometricCDF( _
                K, ProbSuccess, TailValue, FailMsg) Then Exit Function

            Qualifies = (TailValue >= Probability)
        Else
            If Not PROB_DS_TryGeometricSF( _
                K, ProbSuccess, TailValue, FailMsg) Then Exit Function

            Qualifies = (TailValue <= ComplementProbability)
        End If

        PROB_DS_TryGeometricQualifies = True
End Function


Private Function PROB_DS_TryBinomialInverse( _
    ByVal Probability As Double, _
    ByVal n As Double, _
    ByVal ProbSuccess As Double, _
    ByRef Result As Double, _
    ByRef FailMsg As String) _
    As Boolean
'
'==============================================================================
' PROB_DS_TryBinomialInverse
'------------------------------------------------------------------------------
' PURPOSE
'   Returns the least integer K in [0,N] satisfying F(K) >= Probability.
'
' METHOD
'   - Starts from a normal/Cornish-Fisher approximation using the Binomial
'     skewness (1-2p)/Sqr(npq).
'   - Expands downward or upward to a failing/qualifying bracket, capped at N.
'   - Refines with smaller-tail integer bisection.
'   The objective uses the smaller direct tail via PROB_DS_TryBinomialQualifies.
'
' UPDATED
'   2026-07-19 - Cornish-Fisher seed replaces the full-range bisection.
'==============================================================================
'
    Dim ComplementProbability   As Double      '1 - target probability
    Dim Z                   As Double          'Normal quantile seed
    Dim Np                  As Double          'Expected success count n*p
    Dim Npq                 As Double          'Variance n*p*(1-p)
    Dim SeedReal            As Double          'Cornish-Fisher real seed
    Dim Seed                As Double          'Truncated integer seed
    Dim StepSize            As Double          'Bracket expansion step
    Dim Candidate           As Double          'Candidate bracket point
    Dim Lo                  As Double          'Known failing integer or -1 sentinel
    Dim Hi                  As Double          'Known qualifying integer
    Dim MidPoint            As Double          'Integer bisection midpoint
    Dim Qualifies           As Boolean         'Objective result
    Dim BracketFound        As Boolean         'TRUE once Lo and Hi are established
    Dim IterIdx             As Long            'Iteration index

'------------------------------------------------------------------------------
' HANDLE DEGENERATE SUPPORT
'------------------------------------------------------------------------------
    'N = 0 is the point mass at zero
        If n <= 0# Then
            Result = 0#
            PROB_DS_TryBinomialInverse = True
            Exit Function
        End If

'------------------------------------------------------------------------------
' BUILD SEED
'------------------------------------------------------------------------------
        ComplementProbability = 1# - Probability
        Np = n * ProbSuccess
        Npq = Np * (1# - ProbSuccess)

        If Probability <= 0.5 Then
            Z = PROB_NormalInvCDFRaw(Probability)
        Else
            Z = -PROB_NormalInvCDFRaw(ComplementProbability)
        End If

    'Normal seed with the first Cornish-Fisher skewness correction
        SeedReal = Np + Sqr(Npq) * Z + (1# - 2# * ProbSuccess) / 6# * (Z * Z - 1#)

        If SeedReal <= 0# Then
            Seed = 0#
        ElseIf SeedReal >= n Then
            Seed = n
        Else
            Seed = Int(SeedReal)
        End If

'------------------------------------------------------------------------------
' EVALUATE SEED
'------------------------------------------------------------------------------
        If Not PROB_DS_TryBinomialQualifies( _
            Seed, Probability, ComplementProbability, n, ProbSuccess, _
            Qualifies, FailMsg) Then Exit Function

        StepSize = Int(Sqr(Npq)) + 1#
        If StepSize < 1# Then StepSize = 1#

'------------------------------------------------------------------------------
' FIND BRACKET
'------------------------------------------------------------------------------
        BracketFound = False

        If Qualifies Then
            Hi = Seed

            If Hi <= 0# Then
                Lo = -1#
                BracketFound = True
            Else
                For IterIdx = 1 To PROB_DS_MAX_BRACKET_ITER
                    Candidate = Hi - StepSize
                    If Candidate < 0# Then Candidate = 0#

                    If Not PROB_DS_TryBinomialQualifies( _
                        Candidate, Probability, ComplementProbability, n, ProbSuccess, _
                        Qualifies, FailMsg) Then Exit Function

                    If Qualifies Then
                        Hi = Candidate

                        If Hi <= 0# Then
                            Lo = -1#
                            BracketFound = True
                            Exit For
                        End If

                        StepSize = StepSize * 2#
                    Else
                        Lo = Candidate
                        BracketFound = True
                        Exit For
                    End If
                Next IterIdx

                If Not BracketFound Then
                    FailMsg = "Binomial inverse downward bracketing failed in " & _
                              PROB_DS_MAX_BRACKET_ITER & " iterations"
                    Exit Function
                End If
            End If
        Else
            Lo = Seed

            For IterIdx = 1 To PROB_DS_MAX_BRACKET_ITER
                Candidate = Lo + StepSize
                If Candidate >= n Then Candidate = n

                If Not PROB_DS_TryBinomialQualifies( _
                    Candidate, Probability, ComplementProbability, n, ProbSuccess, _
                    Qualifies, FailMsg) Then Exit Function

                If Qualifies Then
                    Hi = Candidate
                    BracketFound = True
                    Exit For
                End If

    'The full count N always qualifies for Probability < 1
                If Candidate >= n Then
                    FailMsg = "Binomial inverse failed to bracket at the full count"
                    Exit Function
                End If

                Lo = Candidate
                StepSize = StepSize * 2#
            Next IterIdx

            If Not BracketFound Then
                FailMsg = "Binomial inverse upward bracketing failed in " & _
                          PROB_DS_MAX_BRACKET_ITER & " iterations"
                Exit Function
            End If
        End If

'------------------------------------------------------------------------------
' BISECT BRACKET
'------------------------------------------------------------------------------
        For IterIdx = 1 To PROB_DS_MAX_INVERSE_ITER
            If Hi - Lo <= 1# Then
                Result = Hi
                PROB_DS_TryBinomialInverse = True
                Exit Function
            End If

            MidPoint = Lo + Int((Hi - Lo) / 2#)

            If Not PROB_DS_TryBinomialQualifies( _
                MidPoint, Probability, ComplementProbability, n, ProbSuccess, _
                Qualifies, FailMsg) Then Exit Function

            If Qualifies Then
                Hi = MidPoint
            Else
                Lo = MidPoint
            End If
        Next IterIdx

        FailMsg = "Binomial inverse failed to converge in " & _
                  PROB_DS_MAX_INVERSE_ITER & " integer iterations"
End Function


Private Function PROB_DS_TryPoissonInverse( _
    ByVal Probability As Double, _
    ByVal Mean As Double, _
    ByRef Result As Double, _
    ByRef FailMsg As String) _
    As Boolean
'
'==============================================================================
' PROB_DS_TryPoissonInverse
'------------------------------------------------------------------------------
' PURPOSE
'   Returns the least integer K satisfying F(K) >= Probability.
'
' METHOD
'   - Starts from a normal/Cornish-Fisher approximation.
'   - Expands downward or upward until a failing/qualifying bracket is known.
'   - Evaluates the configured ceiling before declaring failure.
'   - Refines with smaller-tail integer bisection.
'
' UPDATED
'   2026-07-19 - Replaces the unseeded powers-of-two search.
'==============================================================================
'
    Dim ComplementProbability As Double        '1 - target probability
    Dim Z                   As Double          'Normal quantile seed
    Dim SqrtMean            As Double          'Square root of Mean
    Dim SeedReal            As Double          'Cornish-Fisher real seed
    Dim Seed                As Double          'Truncated integer seed
    Dim StepSize            As Double          'Bracket expansion step
    Dim Candidate           As Double          'Candidate bracket point
    Dim Lo                  As Double          'Known failing integer or -1 sentinel
    Dim Hi                  As Double          'Known qualifying integer
    Dim MidPoint            As Double          'Integer bisection midpoint
    Dim Qualifies           As Boolean         'Objective result
    Dim BracketFound        As Boolean         'TRUE once Lo and Hi are established
    Dim IterIdx             As Long            'Iteration index

'------------------------------------------------------------------------------
' HANDLE ZERO MEAN
'------------------------------------------------------------------------------
        If Mean <= 0# Then
            Result = 0#
            PROB_DS_TryPoissonInverse = True
            Exit Function
        End If

'------------------------------------------------------------------------------
' BUILD SEED
'------------------------------------------------------------------------------
        ComplementProbability = 1# - Probability

        If Probability <= 0.5 Then
            Z = PROB_NormalInvCDFRaw(Probability)
        Else
            Z = -PROB_NormalInvCDFRaw(ComplementProbability)
        End If

        SqrtMean = Sqr(Mean)

    'First Cornish-Fisher correction for Poisson skewness
        SeedReal = Mean + SqrtMean * Z + (Z * Z - 1#) / 6#

        If SeedReal <= 0# Then
            Seed = 0#
        ElseIf SeedReal >= PROB_DS_MAX_POISSON_KERNEL_COUNT Then
            Seed = PROB_DS_MAX_POISSON_KERNEL_COUNT
        Else
            Seed = Int(SeedReal)
        End If

'------------------------------------------------------------------------------
' EVALUATE SEED
'------------------------------------------------------------------------------
        If Not PROB_DS_TryPoissonQualifies( _
            Seed, _
            Probability, _
            ComplementProbability, _
            Mean, _
            Qualifies, _
            FailMsg) Then Exit Function

        StepSize = Int(SqrtMean) + 1#
        If StepSize < 1# Then StepSize = 1#

'------------------------------------------------------------------------------
' FIND BRACKET
'------------------------------------------------------------------------------
        BracketFound = False

        If Qualifies Then
            Hi = Seed

            If Hi <= 0# Then
                Lo = -1#
                BracketFound = True
            Else
                For IterIdx = 1 To PROB_DS_MAX_BRACKET_ITER
                    Candidate = Hi - StepSize
                    If Candidate < 0# Then Candidate = 0#

                    If Not PROB_DS_TryPoissonQualifies( _
                        Candidate, _
                        Probability, _
                        ComplementProbability, _
                        Mean, _
                        Qualifies, _
                        FailMsg) Then Exit Function

                    If Qualifies Then
                        Hi = Candidate

                        If Hi <= 0# Then
                            Lo = -1#
                            BracketFound = True
                            Exit For
                        End If

                        StepSize = StepSize * 2#
                    Else
                        Lo = Candidate
                        BracketFound = True
                        Exit For
                    End If
                Next IterIdx

                If Not BracketFound Then
                    FailMsg = "Poisson inverse downward bracketing failed in " & _
                              PROB_DS_MAX_BRACKET_ITER & " iterations"
                    Exit Function
                End If
            End If
        Else
            Lo = Seed

            For IterIdx = 1 To PROB_DS_MAX_BRACKET_ITER
                Candidate = Lo + StepSize

                If Candidate >= PROB_DS_MAX_POISSON_KERNEL_COUNT Then
                    Candidate = PROB_DS_MAX_POISSON_KERNEL_COUNT
                End If

                If Not PROB_DS_TryPoissonQualifies( _
                    Candidate, _
                    Probability, _
                    ComplementProbability, _
                    Mean, _
                    Qualifies, _
                    FailMsg) Then Exit Function

                If Qualifies Then
                    Hi = Candidate
                    BracketFound = True
                    Exit For
                End If

                If Candidate >= PROB_DS_MAX_POISSON_KERNEL_COUNT Then
                    FailMsg = "Poisson quantile exceeds the supported ceiling of " & _
                              CStr(PROB_DS_MAX_POISSON_KERNEL_COUNT)
                    Exit Function
                End If

                Lo = Candidate
                StepSize = StepSize * 2#
            Next IterIdx

            If Not BracketFound Then
                FailMsg = "Poisson inverse upward bracketing failed in " & _
                          PROB_DS_MAX_BRACKET_ITER & " iterations"
                Exit Function
            End If
        End If

'------------------------------------------------------------------------------
' BISECT BRACKET
'------------------------------------------------------------------------------
        For IterIdx = 1 To PROB_DS_MAX_INVERSE_ITER
            If Hi - Lo <= 1# Then
                Result = Hi
                PROB_DS_TryPoissonInverse = True
                Exit Function
            End If

            MidPoint = Lo + Int((Hi - Lo) / 2#)

            If Not PROB_DS_TryPoissonQualifies( _
                MidPoint, _
                Probability, _
                ComplementProbability, _
                Mean, _
                Qualifies, _
                FailMsg) Then Exit Function

            If Qualifies Then
                Hi = MidPoint
            Else
                Lo = MidPoint
            End If
        Next IterIdx

        FailMsg = "Poisson inverse failed to converge in " & _
                  PROB_DS_MAX_INVERSE_ITER & " integer iterations"
End Function


Private Function PROB_DS_TryGeometricInverse( _
    ByVal Probability As Double, _
    ByVal ProbSuccess As Double, _
    ByRef Result As Double, _
    ByRef FailMsg As String) _
    As Boolean
'
'==============================================================================
' PROB_DS_TryGeometricInverse
'------------------------------------------------------------------------------
' PURPOSE
'   Returns the least integer K satisfying
'       1 - (1-p)^(K+1) >= Probability.
'
' METHOD
'   Seeds the closed form, rejects an unrepresentable integer quantile and then
'   performs a bounded monotone correction against the smaller direct tail.
'
' UPDATED
'   2026-07-19
'==============================================================================
'
    Dim ComplementProbability As Double        '1 - target probability
    Dim RatioValue          As Double          'Log-tail ratio for K+1
    Dim Seed                As Double          'Closed-form integer seed
    Dim Qualifies           As Boolean         'Objective result
    Dim IterIdx             As Long            'Correction iteration

        ComplementProbability = 1# - Probability

'------------------------------------------------------------------------------
' FORM CLOSED-FORM SEED
'------------------------------------------------------------------------------
    'A division overflow means the integer quantile is beyond Double support
        If Not PROB_TryDivide( _
            PROB_Log1p(-Probability), _
            PROB_Log1p(-ProbSuccess), _
            RatioValue) Then

            FailMsg = "Geometric quantile exceeds the supported exact-integer range"
            Exit Function
        End If

    'K = Ceil(RatioValue) - 1; VBA Ceil(X) = -Int(-X)
        Seed = -Int(-RatioValue) - 1#

        If Seed < 0# Then Seed = 0#

        If Seed > PROB_DS_MAX_EXACT_INTEGER Then
            FailMsg = "Geometric quantile exceeds the supported maximum of " & _
                      CStr(PROB_DS_MAX_EXACT_INTEGER)
            Exit Function
        End If

'------------------------------------------------------------------------------
' CORRECT DOWNWARD
'------------------------------------------------------------------------------
        For IterIdx = 1 To PROB_DS_MAX_GEOMETRIC_CORRECTIONS
            If Seed <= 0# Then Exit For

            If Not PROB_DS_TryGeometricQualifies( _
                Seed - 1#, _
                Probability, _
                ComplementProbability, _
                ProbSuccess, _
                Qualifies, _
                FailMsg) Then Exit Function

            If Qualifies Then
                Seed = Seed - 1#
            Else
                Exit For
            End If
        Next IterIdx

        If IterIdx > PROB_DS_MAX_GEOMETRIC_CORRECTIONS Then
            FailMsg = "Geometric inverse exceeded the downward correction budget"
            Exit Function
        End If

'------------------------------------------------------------------------------
' CORRECT UPWARD
'------------------------------------------------------------------------------
        For IterIdx = 1 To PROB_DS_MAX_GEOMETRIC_CORRECTIONS
            If Not PROB_DS_TryGeometricQualifies( _
                Seed, _
                Probability, _
                ComplementProbability, _
                ProbSuccess, _
                Qualifies, _
                FailMsg) Then Exit Function

            If Qualifies Then
                Result = Seed
                PROB_DS_TryGeometricInverse = True
                Exit Function
            End If

            If Seed >= PROB_DS_MAX_EXACT_INTEGER Then
                FailMsg = "Geometric quantile exceeds the supported maximum of " & _
                          CStr(PROB_DS_MAX_EXACT_INTEGER)
                Exit Function
            End If

            Seed = Seed + 1#
        Next IterIdx

        FailMsg = "Geometric inverse exceeded the upward correction budget"
End Function


Private Function PROB_DS_TryNegativeBinomialQualifies( _
    ByVal K As Double, _
    ByVal Probability As Double, _
    ByVal ComplementProbability As Double, _
    ByVal R As Double, _
    ByVal ProbSuccess As Double, _
    ByRef Qualifies As Boolean, _
    ByRef FailMsg As String) _
    As Boolean
'
'==============================================================================
' PROB_DS_TryNegativeBinomialQualifies
'------------------------------------------------------------------------------
' PURPOSE
'   Tests F(K) >= Probability using the smaller direct probability tail.
'
' UPDATED
'   2026-07-21
'==============================================================================
'
    Dim TailValue           As Double          'CDF or SF at K

        If Probability <= 0.5 Then
            If Not PROB_DS_TryNegativeBinomialCDF( _
                K, R, ProbSuccess, TailValue, FailMsg) Then Exit Function
            Qualifies = (TailValue >= Probability)
        Else
            If Not PROB_DS_TryNegativeBinomialSF( _
                K, R, ProbSuccess, TailValue, FailMsg) Then Exit Function
            Qualifies = (TailValue <= ComplementProbability)
        End If

        PROB_DS_TryNegativeBinomialQualifies = True
End Function


Private Function PROB_DS_TryNegativeBinomialInverse( _
    ByVal Probability As Double, _
    ByVal R As Double, _
    ByVal ProbSuccess As Double, _
    ByRef Result As Double, _
    ByRef FailMsg As String) _
    As Boolean
'
'==============================================================================
' PROB_DS_TryNegativeBinomialInverse
'------------------------------------------------------------------------------
' PURPOSE
'   Least integer failure count K with F(K) >= Probability.
'   Seeds from the normal/Cornish-Fisher approximation using the negative
'   binomial mean, spread and skewness, expands to a bracket capped at the
'   kernel ceiling, then refines by smaller-tail integer bisection.
'
' UPDATED
'   2026-07-21
'==============================================================================
'
    Dim ComplementProbability As Double        '1 - target probability
    Dim Q                   As Double          'Failure probability 1-p
    Dim Mean                As Double          'r(1-p)/p
    Dim SdDev               As Double          'Sqr(r(1-p))/p
    Dim Z                   As Double          'Normal quantile seed
    Dim SeedReal            As Double          'Cornish-Fisher real seed
    Dim Seed                As Double          'Truncated integer seed
    Dim StepSize            As Double          'Bracket expansion step
    Dim Candidate           As Double          'Candidate bracket point
    Dim Lo                  As Double          'Known failing integer or -1 sentinel
    Dim Hi                  As Double          'Known qualifying integer
    Dim MidPoint            As Double          'Integer bisection midpoint
    Dim Qualifies           As Boolean         'Objective result
    Dim BracketFound        As Boolean         'TRUE once Lo and Hi are established
    Dim IterIdx             As Long            'Iteration index

        ComplementProbability = 1# - Probability
        Q = 1# - ProbSuccess

    'Certain success collapses the distribution onto zero failures
        If Q <= 0# Then
            Result = 0#
            PROB_DS_TryNegativeBinomialInverse = True
            Exit Function
        End If

        Mean = R * Q / ProbSuccess
        SdDev = Sqr(R * Q) / ProbSuccess

        If Probability <= 0.5 Then
            Z = PROB_NormalInvCDFRaw(Probability)
        Else
            Z = -PROB_NormalInvCDFRaw(ComplementProbability)
        End If

    'Normal seed with the first Cornish-Fisher skewness correction
        SeedReal = Mean + SdDev * Z + (2# - ProbSuccess) / (6# * ProbSuccess) * (Z * Z - 1#)

        If SeedReal <= 0# Then
            Seed = 0#
        ElseIf SeedReal >= PROB_DS_MAX_NEGBINOM_KERNEL_COUNT Then
            Seed = PROB_DS_MAX_NEGBINOM_KERNEL_COUNT
        Else
            Seed = Int(SeedReal)
        End If

        If Not PROB_DS_TryNegativeBinomialQualifies( _
            Seed, Probability, ComplementProbability, R, ProbSuccess, _
            Qualifies, FailMsg) Then Exit Function

        StepSize = Int(SdDev) + 1#
        If StepSize < 1# Then StepSize = 1#

        BracketFound = False

        If Qualifies Then
            Hi = Seed

            If Hi <= 0# Then
                Lo = -1#
                BracketFound = True
            Else
                For IterIdx = 1 To PROB_DS_MAX_BRACKET_ITER
                    Candidate = Hi - StepSize
                    If Candidate < 0# Then Candidate = 0#

                    If Not PROB_DS_TryNegativeBinomialQualifies( _
                        Candidate, Probability, ComplementProbability, R, ProbSuccess, _
                        Qualifies, FailMsg) Then Exit Function

                    If Qualifies Then
                        Hi = Candidate
                        If Hi <= 0# Then
                            Lo = -1#
                            BracketFound = True
                            Exit For
                        End If
                        StepSize = StepSize * 2#
                    Else
                        Lo = Candidate
                        BracketFound = True
                        Exit For
                    End If
                Next IterIdx

                If Not BracketFound Then
                    FailMsg = "Negative binomial inverse downward bracketing failed"
                    Exit Function
                End If
            End If
        Else
            Lo = Seed

            For IterIdx = 1 To PROB_DS_MAX_BRACKET_ITER
                Candidate = Lo + StepSize

                If Candidate >= PROB_DS_MAX_NEGBINOM_KERNEL_COUNT Then
                    Candidate = PROB_DS_MAX_NEGBINOM_KERNEL_COUNT
                End If

                If Not PROB_DS_TryNegativeBinomialQualifies( _
                    Candidate, Probability, ComplementProbability, R, ProbSuccess, _
                    Qualifies, FailMsg) Then Exit Function

                If Qualifies Then
                    Hi = Candidate
                    BracketFound = True
                    Exit For
                End If

                If Candidate >= PROB_DS_MAX_NEGBINOM_KERNEL_COUNT Then
                    FailMsg = "Negative binomial quantile exceeds the supported ceiling of " & _
                              CStr(PROB_DS_MAX_NEGBINOM_KERNEL_COUNT)
                    Exit Function
                End If

                Lo = Candidate
                StepSize = StepSize * 2#
            Next IterIdx

            If Not BracketFound Then
                FailMsg = "Negative binomial inverse upward bracketing failed"
                Exit Function
            End If
        End If

        For IterIdx = 1 To PROB_DS_MAX_INVERSE_ITER
            If Hi - Lo <= 1# Then
                Result = Hi
                PROB_DS_TryNegativeBinomialInverse = True
                Exit Function
            End If

            MidPoint = Lo + Int((Hi - Lo) / 2#)

            If Not PROB_DS_TryNegativeBinomialQualifies( _
                MidPoint, Probability, ComplementProbability, R, ProbSuccess, _
                Qualifies, FailMsg) Then Exit Function

            If Qualifies Then
                Hi = MidPoint
            Else
                Lo = MidPoint
            End If
        Next IterIdx

        FailMsg = "Negative binomial inverse failed to converge in " & _
                  PROB_DS_MAX_INVERSE_ITER & " integer iterations"
End Function


Private Function PROB_DS_TryHypergeometricQualifies( _
    ByVal K As Double, _
    ByVal Probability As Double, _
    ByVal ComplementProbability As Double, _
    ByVal n As Double, _
    ByVal Kp As Double, _
    ByVal Np As Double, _
    ByRef Qualifies As Boolean, _
    ByRef FailMsg As String) _
    As Boolean
'
'==============================================================================
' PROB_DS_TryHypergeometricQualifies
'------------------------------------------------------------------------------
' PURPOSE
'   Tests F(K) >= Probability using the smaller direct probability tail.
'
' UPDATED
'   2026-07-21
'==============================================================================
'
    Dim TailValue           As Double          'CDF or SF at K

        If Probability <= 0.5 Then
            If Not PROB_DS_TryHypergeometricCDF( _
                K, n, Kp, Np, TailValue, FailMsg) Then Exit Function
            Qualifies = (TailValue >= Probability)
        Else
            If Not PROB_DS_TryHypergeometricSF( _
                K, n, Kp, Np, TailValue, FailMsg) Then Exit Function
            Qualifies = (TailValue <= ComplementProbability)
        End If

        PROB_DS_TryHypergeometricQualifies = True
End Function


Private Function PROB_DS_TryHypergeometricInverse( _
    ByVal Probability As Double, _
    ByVal n As Double, _
    ByVal Kp As Double, _
    ByVal Np As Double, _
    ByRef Result As Double, _
    ByRef FailMsg As String) _
    As Boolean
'
'==============================================================================
' PROB_DS_TryHypergeometricInverse
'------------------------------------------------------------------------------
' PURPOSE
'   Least integer sample-success count K in [KMin, KMax] with F(K) >= Probability.
'   Seeds from the normal approximation using the hypergeometric mean and
'   spread, brackets within the finite support, then refines by smaller-tail
'   integer bisection.  KMax always qualifies for Probability < 1.
'
' UPDATED
'   2026-07-21
'==============================================================================
'
    Dim ComplementProbability As Double        '1 - target probability
    Dim KMin                As Double          'Lowest attainable count
    Dim KMax                As Double          'Highest attainable count
    Dim Mu                  As Double          'Mean n*K/N
    Dim Variance            As Double          'Hypergeometric variance
    Dim SdDev               As Double          'Square root of the variance
    Dim Z                   As Double          'Normal quantile seed
    Dim SeedReal            As Double          'Real-valued seed
    Dim Seed                As Double          'Truncated integer seed
    Dim StepSize            As Double          'Bracket expansion step
    Dim Candidate           As Double          'Candidate bracket point
    Dim Lo                  As Double          'Known failing integer or KMin-1 sentinel
    Dim Hi                  As Double          'Known qualifying integer
    Dim MidPoint            As Double          'Integer bisection midpoint
    Dim Qualifies           As Boolean         'Objective result
    Dim BracketFound        As Boolean         'TRUE once Lo and Hi are established
    Dim IterIdx             As Long            'Iteration index

        ComplementProbability = 1# - Probability

        KMin = n + Kp - Np
        If KMin < 0# Then KMin = 0#
        KMax = n
        If Kp < KMax Then KMax = Kp

    'A single-point support is its own quantile
        If KMax <= KMin Then
            Result = KMin
            PROB_DS_TryHypergeometricInverse = True
            Exit Function
        End If

        Mu = n * Kp / Np
        If Np <= 1# Then
            Variance = 0#
        Else
            Variance = n * (Kp / Np) * ((Np - Kp) / Np) * ((Np - n) / (Np - 1#))
        End If
        SdDev = Sqr(Variance)

        If Probability <= 0.5 Then
            Z = PROB_NormalInvCDFRaw(Probability)
        Else
            Z = -PROB_NormalInvCDFRaw(ComplementProbability)
        End If

        SeedReal = Mu + SdDev * Z

        If SeedReal <= KMin Then
            Seed = KMin
        ElseIf SeedReal >= KMax Then
            Seed = KMax
        Else
            Seed = Int(SeedReal)
        End If

        If Not PROB_DS_TryHypergeometricQualifies( _
            Seed, Probability, ComplementProbability, n, Kp, Np, _
            Qualifies, FailMsg) Then Exit Function

        StepSize = Int(SdDev) + 1#
        If StepSize < 1# Then StepSize = 1#

        BracketFound = False

        If Qualifies Then
            Hi = Seed

            If Hi <= KMin Then
                Lo = KMin - 1#
                BracketFound = True
            Else
                For IterIdx = 1 To PROB_DS_MAX_BRACKET_ITER
                    Candidate = Hi - StepSize
                    If Candidate < KMin Then Candidate = KMin

                    If Not PROB_DS_TryHypergeometricQualifies( _
                        Candidate, Probability, ComplementProbability, n, Kp, Np, _
                        Qualifies, FailMsg) Then Exit Function

                    If Qualifies Then
                        Hi = Candidate
                        If Hi <= KMin Then
                            Lo = KMin - 1#
                            BracketFound = True
                            Exit For
                        End If
                        StepSize = StepSize * 2#
                    Else
                        Lo = Candidate
                        BracketFound = True
                        Exit For
                    End If
                Next IterIdx

                If Not BracketFound Then
                    FailMsg = "Hypergeometric inverse downward bracketing failed"
                    Exit Function
                End If
            End If
        Else
            Lo = Seed

            For IterIdx = 1 To PROB_DS_MAX_BRACKET_ITER
                Candidate = Lo + StepSize
                If Candidate >= KMax Then Candidate = KMax

                If Not PROB_DS_TryHypergeometricQualifies( _
                    Candidate, Probability, ComplementProbability, n, Kp, Np, _
                    Qualifies, FailMsg) Then Exit Function

                If Qualifies Then
                    Hi = Candidate
                    BracketFound = True
                    Exit For
                End If

                If Candidate >= KMax Then
                    FailMsg = "Hypergeometric inverse failed to bracket at the support ceiling"
                    Exit Function
                End If

                Lo = Candidate
                StepSize = StepSize * 2#
            Next IterIdx

            If Not BracketFound Then
                FailMsg = "Hypergeometric inverse upward bracketing failed"
                Exit Function
            End If
        End If

        For IterIdx = 1 To PROB_DS_MAX_INVERSE_ITER
            If Hi - Lo <= 1# Then
                Result = Hi
                PROB_DS_TryHypergeometricInverse = True
                Exit Function
            End If

            MidPoint = Lo + Int((Hi - Lo) / 2#)

            If Not PROB_DS_TryHypergeometricQualifies( _
                MidPoint, Probability, ComplementProbability, n, Kp, Np, _
                Qualifies, FailMsg) Then Exit Function

            If Qualifies Then
                Hi = MidPoint
            Else
                Lo = MidPoint
            End If
        Next IterIdx

        FailMsg = "Hypergeometric inverse failed to converge in " & _
                  PROB_DS_MAX_INVERSE_ITER & " integer iterations"
End Function


