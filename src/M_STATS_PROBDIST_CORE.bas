Attribute VB_Name = "M_STATS_PROBDIST_CORE"
Option Explicit
Option Private Module

'==============================================================================
' M_STATS_PROBDIST_CORE
'------------------------------------------------------------------------------
' PURPOSE
'   Provides the shared numeric constants, finiteness/domain predicates, safe
'   exponential and logarithm primitives, the raw inverse-normal kernel and the
'   diagnostic status writer used by every M_STATS_PROBDIST_* module.
'
' WHY THIS EXISTS
'   Before this module, PROB_IsFinite, PROB_IsValidProbabilityOpen, PROB_TryExp
'   and PROB_SetStatus each existed as separate copies inside NORMALFAMILY and
'   SPECIALFUNCS, and six constants were declared three times. Two copies of a
'   predicate are two chances to disagree, and a Private copy shadowing a Public
'   one of the same name compiles silently today but raises "Ambiguous name
'   detected" the moment anyone widens the scope of either. One definition, one
'   home.
'
'   Option Private Module keeps every name here invisible to the worksheet while
'   leaving it project-visible to the sibling modules, which is exactly the
'   scope a shared kernel layer wants.
'
' PUBLIC (PROJECT-SCOPED) SURFACE
'   Constants:
'     - PROB_PI, PROB_TWO_PI, PROB_HALF_LOG_TWO_PI, PROB_HALF_LOG_PI
'     - PROB_EPS, PROB_NUM_EPS, PROB_MACH_EPS
'     - PROB_MAX_EXP, PROB_MIN_EXP
'     - PROB_LARGE_NUMBER, PROB_DOUBLE_MAX, PROB_FPMIN
'     - PROB_WRITE_STATUS_BAR
'
'   Predicates:
'     - PROB_IsFinite
'     - PROB_IsWithinSupportedMagnitude
'     - PROB_IsPositiveFinite
'     - PROB_IsPositiveWithinSupportedMagnitude
'     - PROB_IsValidProbabilityOpen
'
'   Numeric primitives:
'     - PROB_TryExp
'     - PROB_TryAdd
'     - PROB_TryMultiply
'     - PROB_TryDivide
'     - PROB_Log1p
'     - PROB_Expm1
'     - PROB_NormalInvCDFRaw
'
'   Diagnostics:
'     - PROB_SetStatus
'
' ALGORITHM PROVENANCE
'   - PROB_Log1p:
'       Kahan's compensated form, Log(1+X) = Log(U) * X / (U - 1) with U = 1+X.
'       Exact to 1 ulp across the whole representable range of X, unlike a
'       Taylor-below-a-threshold arrangement, which is only as good as its
'       threshold. Public, published; not proprietary.
'   - PROB_Expm1:
'       Kahan's compensated form, Exp(X) - 1 = (U - 1) * X / Log(U) with
'       U = Exp(X). The mirror image of PROB_Log1p: it recovers the low bits of
'       Exp(X) - 1 that (U - 1) throws away when X is near zero, so the entire
'       left tail of the Exponential and Weibull CDFs keeps full relative
'       precision instead of collapsing to zero. Public, published.
'   - PROB_NormalInvCDFRaw:
'       Peter J. Acklam's rational approximation, released freely for any use by
'       the author. Raw accuracy is approximately 1.15E-9; it is used here as a
'       root-finder seed, not as a final answer. Public.
'
' DESIGN PRINCIPLES
'   - Nothing here knows about any distribution. This is a numerics layer.
'   - True finiteness and supported algorithm magnitude are separate contracts.
'     PROB_IsFinite tests the IEEE value; PROB_IsWithinSupportedMagnitude applies
'     the conservative 1E+100 policy only where a numerical algorithm needs it.
'   - Overflow fails explicitly: a computation that would exceed Double range
'     returns False rather than a clamped sentinel value.
'   - Underflow of an exponential is a valid zero, not an error.
'   - Kernels here never validate their callers' domains and never write Status.
'
' NOTES
'   - PROB_IsFinite is a true finiteness predicate. It does not impose the
'     project magnitude policy. Use PROB_IsWithinSupportedMagnitude explicitly
'     for shape, degree-of-freedom and other algorithmic parameters.
'   - The subtraction X - X distinguishes finite values from externally supplied
'     IEEE infinities while preserving the largest finite Double.
'   - ARCHITECTURE: this module owns shared constants and elementary numeric
'     helpers. M_STATS_PROBDIST_SPECIALFUNCS owns reusable special-function
'     kernels; distribution-family modules consume both layers without keeping
'     private duplicate copies; M_STATS_PROBDIST_TEST owns regression coverage.
'
' UPDATED
'   2026-07-11
'==============================================================================

'==============================================================================
' PUBLIC CONSTANTS
'==============================================================================

'The VBA editor canonicalizes long decimal literals to approximately
'15 significant decimal digits. Split constant expressions are therefore
'used where necessary to obtain the intended IEEE-754 Double value.
Public Const PROB_PI As Double = _
    3.14159265358979 + 3.10862446895044E-15

Public Const PROB_TWO_PI As Double = _
    2# * PROB_PI

Public Const PROB_HALF_LOG_TWO_PI As Double = _
    0.918938533204672 + 6.66133814775094E-16

Public Const PROB_HALF_LOG_PI As Double = _
    0.5723649429247 + 1.11022302462516E-16

Public Const PROB_EPS                  As Double = 0.000000000000001     '1E-15, relative convergence target
Public Const PROB_NUM_EPS              As Double = 0.00000000000003      '3E-14, continued-fraction / series stop
Public Const PROB_MACH_EPS             As Double = 2.22044604925031E-16  'Double epsilon

Public Const PROB_MAX_EXP              As Double = 709.782712893384      'Advisory Log(Double max)
Public Const PROB_MIN_EXP              As Double = -745.133219101941     'Advisory round-to-zero boundary

Public Const PROB_LARGE_NUMBER         As Double = 1E+100                'Supported algorithm magnitude bound
Public Const PROB_DOUBLE_MAX           As Double = 1.79769313486231E+308 'Approx largest finite Double
Public Const PROB_FPMIN                As Double = 1E-300                'Lentz denominator floor

Public Const PROB_WRITE_STATUS_BAR     As Boolean = False                'Master switch for Application.StatusBar writes


'==============================================================================
' PREDICATES
'==============================================================================

Public Function PROB_IsFinite( _
    ByVal X As Double) _
    As Boolean
'
'==============================================================================
' PROB_IsFinite
'------------------------------------------------------------------------------
' PURPOSE
'   Returns TRUE only when X is an IEEE-754 finite Double.
'
' RETURNS
'   Boolean
'     TRUE  => X is a finite Double.
'     FALSE => X is a NaN or an infinity supplied by external COM code.
'
' NOTES
'   This predicate deliberately does not apply PROB_LARGE_NUMBER. A separate
'   predicate owns that project-specific supported-magnitude policy. The
'   subtraction X - X is zero only for a finite X: it is NaN for an infinity and
'   raises at the overflow edge, both of which the handler catches.
'==============================================================================
'
'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Dim ZeroCheck           As Double          'X - X, zero only for finite X

'------------------------------------------------------------------------------
' EVALUATE
'------------------------------------------------------------------------------
    'Route any arithmetic fault on a non-finite input to the handler
        On Error GoTo Err_Handler

    'Reject a NaN, the only value not equal to itself
        If X <> X Then Exit Function

    'A finite X gives X - X = 0; an infinity gives NaN, which fails this test
        ZeroCheck = X - X
        PROB_IsFinite = (ZeroCheck = 0#)
        Exit Function

'------------------------------------------------------------------------------
' ERROR HANDLER
'------------------------------------------------------------------------------
Err_Handler:
    'Clear the fault and report non-finite
        Err.Clear
        PROB_IsFinite = False
End Function


Public Function PROB_IsWithinSupportedMagnitude( _
    ByVal X As Double) _
    As Boolean
'
'==============================================================================
' PROB_IsWithinSupportedMagnitude
'------------------------------------------------------------------------------
' PURPOSE
'   Returns TRUE when X is finite and lies inside the conservative numerical
'   domain Abs(X) < PROB_LARGE_NUMBER.
'
' RETURNS
'   Boolean
'     TRUE  => X is finite and Abs(X) < PROB_LARGE_NUMBER.
'     FALSE => X is non-finite or at or beyond the supported magnitude bound.
'
' USE
'   Apply this to dimensionless algorithmic parameters whose kernels are tested
'   only inside the project-supported range. Do not use it as a synonym for
'   mathematical finiteness.
'==============================================================================
'
'------------------------------------------------------------------------------
' VALIDATE
'------------------------------------------------------------------------------
    'A non-finite value is never within the supported magnitude
        If Not PROB_IsFinite(X) Then Exit Function

'------------------------------------------------------------------------------
' RETURN
'------------------------------------------------------------------------------
    'Return whether the magnitude sits inside the supported bound
        PROB_IsWithinSupportedMagnitude = (Abs(X) < PROB_LARGE_NUMBER)
End Function


Public Function PROB_IsPositiveFinite( _
    ByVal X As Double) _
    As Boolean
'
'==============================================================================
' PROB_IsPositiveFinite
'------------------------------------------------------------------------------
' PURPOSE
'   Returns TRUE when X is finite and strictly positive.
'
' RETURNS
'   Boolean
'     TRUE  => X is finite and X > 0.
'     FALSE => X is non-finite or non-positive.
'==============================================================================
'
'------------------------------------------------------------------------------
' VALIDATE
'------------------------------------------------------------------------------
    'A non-finite value cannot be positive-finite
        If Not PROB_IsFinite(X) Then Exit Function

'------------------------------------------------------------------------------
' RETURN
'------------------------------------------------------------------------------
    'Return whether X is strictly positive
        PROB_IsPositiveFinite = (X > 0#)
End Function


Public Function PROB_IsPositiveWithinSupportedMagnitude( _
    ByVal X As Double) _
    As Boolean
'
'==============================================================================
' PROB_IsPositiveWithinSupportedMagnitude
'------------------------------------------------------------------------------
' PURPOSE
'   Returns TRUE when X is strictly positive and lies inside the supported
'   algorithm magnitude domain.
'
' RETURNS
'   Boolean
'     TRUE  => X > 0 and Abs(X) < PROB_LARGE_NUMBER.
'     FALSE => X is non-positive or at or beyond the supported magnitude bound.
'==============================================================================
'
'------------------------------------------------------------------------------
' VALIDATE
'------------------------------------------------------------------------------
    'A value outside the supported magnitude fails regardless of sign
        If Not PROB_IsWithinSupportedMagnitude(X) Then Exit Function

'------------------------------------------------------------------------------
' RETURN
'------------------------------------------------------------------------------
    'Return whether X is strictly positive
        PROB_IsPositiveWithinSupportedMagnitude = (X > 0#)
End Function


Public Function PROB_IsValidProbabilityOpen( _
    ByVal Probability As Double) _
    As Boolean
'
'==============================================================================
' PROB_IsValidProbabilityOpen
'------------------------------------------------------------------------------
' PURPOSE
'   Returns TRUE when Probability is finite and strictly inside (0, 1).
'==============================================================================
'
'------------------------------------------------------------------------------
' VALIDATE
'------------------------------------------------------------------------------
    'Reject non-finite probabilities
        If Not PROB_IsFinite(Probability) Then Exit Function

    'Reject endpoints and values outside the unit interval
        If Probability <= 0# Or Probability >= 1# Then Exit Function

'------------------------------------------------------------------------------
' RETURN SUCCESS
'------------------------------------------------------------------------------
    'Return success
        PROB_IsValidProbabilityOpen = True
End Function


'==============================================================================
' NUMERIC PRIMITIVES
'==============================================================================

Public Function PROB_TryExp( _
    ByVal X As Double, _
    ByRef Result As Double) _
    As Boolean
'
'==============================================================================
' PROB_TryExp
'------------------------------------------------------------------------------
' PURPOSE
'   Attempts Exp(X) using the actual VBA floating-point boundary rather than a
'   prematurely rounded decimal cutoff.
'
' CONTRACT
'   - Finite representable result: returns TRUE and writes Result.
'   - Negative underflow:          returns TRUE and writes zero.
'   - Positive overflow:           returns FALSE; Result is not contractual.
'   - Non-finite X:                returns FALSE; Result is not contractual.
'==============================================================================
'
'------------------------------------------------------------------------------
' VALIDATE INPUT
'------------------------------------------------------------------------------
    'Reject a non-finite argument outright
        If Not PROB_IsFinite(X) Then Exit Function

'------------------------------------------------------------------------------
' COMPUTE
'------------------------------------------------------------------------------
    'Route an overflow or underflow fault to the handler
        On Error GoTo Err_Handler

    'Exponentiate and reject a non-finite (overflowed) result
        Result = Exp(X)
        If Not PROB_IsFinite(Result) Then GoTo Err_Handler

'------------------------------------------------------------------------------
' RETURN SUCCESS
'------------------------------------------------------------------------------
    'Report success
        PROB_TryExp = True
        Exit Function

'------------------------------------------------------------------------------
' ERROR HANDLER
'------------------------------------------------------------------------------
Err_Handler:
    'Clear the fault
        Err.Clear

    'A negative exponential can fail only by underflow, which is a valid zero
    'for probability densities and tail probabilities; positive overflow stays
    'a failure with Result left non-contractual
        If X < 0# Then
            Result = 0#
            PROB_TryExp = True
        End If
End Function


Public Function PROB_TryAdd( _
    ByVal A As Double, _
    ByVal B As Double, _
    ByRef Result As Double) _
    As Boolean
'
'==============================================================================
' PROB_TryAdd
'------------------------------------------------------------------------------
' PURPOSE
'   Attempts A + B and converts predictable Double overflow into a FALSE return.
'
' CONTRACT
'   - Finite representable result:  returns TRUE and writes Result.
'   - Overflow or non-finite input: returns FALSE; Result is not contractual.
'==============================================================================
'
'------------------------------------------------------------------------------
' VALIDATE INPUTS
'------------------------------------------------------------------------------
    'Reject non-finite operands
        If Not PROB_IsFinite(A) Then Exit Function
        If Not PROB_IsFinite(B) Then Exit Function

'------------------------------------------------------------------------------
' COMPUTE
'------------------------------------------------------------------------------
    'Route an overflow fault to the handler
        On Error GoTo Err_Handler

    'Add and reject a non-finite (overflowed) result
        Result = A + B
        If Not PROB_IsFinite(Result) Then Exit Function

'------------------------------------------------------------------------------
' RETURN SUCCESS
'------------------------------------------------------------------------------
    'Report success
        PROB_TryAdd = True
        Exit Function

'------------------------------------------------------------------------------
' ERROR HANDLER
'------------------------------------------------------------------------------
Err_Handler:
    'Clear the fault; the default FALSE return stands
        Err.Clear
End Function


Public Function PROB_TryMultiply( _
    ByVal A As Double, _
    ByVal B As Double, _
    ByRef Result As Double) _
    As Boolean
'
'==============================================================================
' PROB_TryMultiply
'------------------------------------------------------------------------------
' PURPOSE
'   Attempts A * B and converts predictable Double overflow into a FALSE return.
'   Underflow to zero is a valid successful result.
'
' CONTRACT
'   - Finite result (including underflow to zero): returns TRUE and writes Result.
'   - Overflow or non-finite input: returns FALSE; Result is not contractual.
'==============================================================================
'
'------------------------------------------------------------------------------
' VALIDATE INPUTS
'------------------------------------------------------------------------------
    'Reject non-finite operands
        If Not PROB_IsFinite(A) Then Exit Function
        If Not PROB_IsFinite(B) Then Exit Function

'------------------------------------------------------------------------------
' COMPUTE
'------------------------------------------------------------------------------
    'Route an overflow fault to the handler
        On Error GoTo Err_Handler

    'Multiply and reject a non-finite (overflowed) result
        Result = A * B
        If Not PROB_IsFinite(Result) Then Exit Function

'------------------------------------------------------------------------------
' RETURN SUCCESS
'------------------------------------------------------------------------------
    'Report success
        PROB_TryMultiply = True
        Exit Function

'------------------------------------------------------------------------------
' ERROR HANDLER
'------------------------------------------------------------------------------
Err_Handler:
    'Clear the fault; the default FALSE return stands
        Err.Clear
End Function


Public Function PROB_TryDivide( _
    ByVal Numerator As Double, _
    ByVal Denominator As Double, _
    ByRef Result As Double) _
    As Boolean
'
'==============================================================================
' PROB_TryDivide
'------------------------------------------------------------------------------
' PURPOSE
'   Attempts Numerator / Denominator and converts division by zero or Double
'   overflow into a FALSE return. Underflow to zero is a valid success.
'
' CONTRACT
'   - Finite result (including underflow to zero): returns TRUE and writes Result.
'   - Zero denominator, overflow or non-finite input: returns FALSE; Result is
'     not contractual.
'==============================================================================
'
'------------------------------------------------------------------------------
' VALIDATE INPUTS
'------------------------------------------------------------------------------
    'Reject non-finite operands and a zero denominator
        If Not PROB_IsFinite(Numerator) Then Exit Function
        If Not PROB_IsFinite(Denominator) Then Exit Function
        If Denominator = 0# Then Exit Function

'------------------------------------------------------------------------------
' COMPUTE
'------------------------------------------------------------------------------
    'Route an overflow fault to the handler
        On Error GoTo Err_Handler

    'Divide and reject a non-finite (overflowed) result
        Result = Numerator / Denominator
        If Not PROB_IsFinite(Result) Then Exit Function

'------------------------------------------------------------------------------
' RETURN SUCCESS
'------------------------------------------------------------------------------
    'Report success
        PROB_TryDivide = True
        Exit Function

'------------------------------------------------------------------------------
' ERROR HANDLER
'------------------------------------------------------------------------------
Err_Handler:
    'Clear the fault; the default FALSE return stands
        Err.Clear
End Function


Public Function PROB_Log1p( _
    ByVal X As Double) _
    As Double
'
'==============================================================================
' PROB_Log1p
'------------------------------------------------------------------------------
' PURPOSE
'   Returns Log(1 + X) accurately for every X > -1, including X near zero.
'
' PRECONDITION
'   X > -1.
'
' RATIONALE
'   The naive Log(1# + X) loses accuracy because 1# + X rounds away the low bits
'   of X: the absolute error of the sum is about 1.1E-16 regardless of X, so the
'   relative error of the logarithm is about 1.1E-16 / X. At X = 1E-8 that is a
'   relative error of 1E-8. Kahan's compensated form recovers the lost bits by
'   scaling Log(U) by the exactly-representable ratio X / (U - 1). Measured
'   relative error is at or below 2.1E-16 for X in [1E-12, 1E-2].
'==============================================================================
'
'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Dim U                   As Double          '1 + X, as actually rounded

'------------------------------------------------------------------------------
' COMPUTE
'------------------------------------------------------------------------------
    'Round the sum once and reuse it
        U = 1# + X

    'Return X exactly when the sum rounds back to one
        If U = 1# Then
            PROB_Log1p = X
    'Otherwise rescale by the exact ratio X / (U - 1)
        Else
            PROB_Log1p = Log(U) * X / (U - 1#)
        End If
End Function


Public Function PROB_Expm1( _
    ByVal X As Double) _
    As Double
'
'==============================================================================
' PROB_Expm1
'------------------------------------------------------------------------------
' PURPOSE
'   Returns Exp(X) - 1 accurately for every X, including X near zero.
'
' PRECONDITION
'   None on mathematical correctness. For sufficiently large positive X the
'   true value overflows a Double and VBA raises overflow error 6. Callers that
'   may pass positive X must first guard the exponential range with
'   PROB_TryExp or an equivalent explicit range check. Negative X is always safe.
'
' RATIONALE
'   The naive Exp(X) - 1# loses accuracy because Exp(X) rounds to a value near 1
'   and the subtraction then cancels the low bits: at X = 1E-10, Exp(X) rounds to
'   1 + 1E-10 with an absolute error of about 1.1E-16, so Exp(X) - 1 carries a
'   relative error of about 1E-6, and 1 - Exp(-(x/lambda)^k) collapses to exactly
'   0 across the whole left tail. Kahan's compensated form recovers the lost bits
'   by scaling U - 1 by the exactly-representable ratio X / Log(U). Measured
'   relative error is at or below 1.2E-16 for X in [-40, 0].
'
'   This is the exact mirror of PROB_Log1p and exists for the same reason: the
'   Exponential and Weibull cumulative distribution functions are 1 - Exp(-z),
'   and computing that as -PROB_Expm1(-z) is the only way to keep the small-z
'   result correct to full relative precision.
'==============================================================================
'
'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Dim U                   As Double          'Exp(X), as actually rounded

'------------------------------------------------------------------------------
' COMPUTE
'------------------------------------------------------------------------------
    'Round the exponential once and reuse it
        U = Exp(X)

    'Return X exactly when the exponential rounds back to one
        If U = 1# Then
            PROB_Expm1 = X
    'Return -1 exactly when the exponential underflows to zero
        ElseIf U = 0# Then
            PROB_Expm1 = -1#
    'Otherwise rescale by the exact ratio X / Log(U)
        Else
            PROB_Expm1 = (U - 1#) * X / Log(U)
        End If
End Function


Public Function PROB_NormalInvCDFRaw( _
    ByVal Probability As Double) _
    As Double
'
'==============================================================================
' PROB_NormalInvCDFRaw
'------------------------------------------------------------------------------
' PURPOSE
'   Returns the raw (unrefined) inverse standard normal CDF via Acklam's
'   rational approximation. Accurate to approximately 1.15E-9.
'
' PRECONDITION
'   0 < Probability < 1.
'
' WHY THIS EXISTS
'   Root finders in this project need a cheap starting point that is already
'   within a few units in the ninth digit. Refinement to machine precision is
'   the caller's job, either by a Halley step (NORMALFAMILY) or by the Newton
'   loop that is running anyway (TFAMILY quantiles).
'==============================================================================
'
'------------------------------------------------------------------------------
' DECLARE CONSTANTS
'------------------------------------------------------------------------------
    Const PLOW  As Double = 0.02425
    Const PHIGH As Double = 0.97575

    Const A1 As Double = -39.6968302866538
    Const A2 As Double = 220.946098424521
    Const A3 As Double = -275.928510446969
    Const A4 As Double = 138.357751867269
    Const A5 As Double = -30.6647980661472
    Const A6 As Double = 2.50662827745924

    Const B1 As Double = -54.4760987982241
    Const B2 As Double = 161.585836858041
    Const B3 As Double = -155.698979859887
    Const B4 As Double = 66.8013118877197
    Const B5 As Double = -13.2806815528857

    Const C1 As Double = -7.78489400243029E-03
    Const C2 As Double = -0.322396458041136
    Const C3 As Double = -2.40075827716184
    Const C4 As Double = -2.54973253934373
    Const C5 As Double = 4.37466414146497
    Const C6 As Double = 2.93816398269878

    Const D1 As Double = 7.78469570904146E-03
    Const D2 As Double = 0.32246712907004
    Const D3 As Double = 2.445134137143
    Const D4 As Double = 3.75440866190742

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Dim Q                   As Double          'Rational-approximation argument
    Dim R                   As Double          'Q squared, central branch

'------------------------------------------------------------------------------
' COMPUTE
'------------------------------------------------------------------------------
    'Lower tail
        If Probability < PLOW Then
            Q = Sqr(-2# * Log(Probability))
            PROB_NormalInvCDFRaw = _
                (((((C1 * Q + C2) * Q + C3) * Q + C4) * Q + C5) * Q + C6) / _
                ((((D1 * Q + D2) * Q + D3) * Q + D4) * Q + 1#)

    'Upper tail
        ElseIf Probability > PHIGH Then
            Q = Sqr(-2# * Log(1# - Probability))
            PROB_NormalInvCDFRaw = _
                -(((((C1 * Q + C2) * Q + C3) * Q + C4) * Q + C5) * Q + C6) / _
                ((((D1 * Q + D2) * Q + D3) * Q + D4) * Q + 1#)

    'Central region
        Else
            Q = Probability - 0.5
            R = Q * Q
            PROB_NormalInvCDFRaw = _
                (((((A1 * R + A2) * R + A3) * R + A4) * R + A5) * R + A6) * Q / _
                (((((B1 * R + B2) * R + B3) * R + B4) * R + B5) * R + 1#)
        End If
End Function


'==============================================================================
' DIAGNOSTICS
'==============================================================================

Public Sub PROB_SetStatus( _
    ByRef Status As String, _
    ByVal Message As String)
'
'==============================================================================
' PROB_SetStatus
'------------------------------------------------------------------------------
' PURPOSE
'   Writes a diagnostic message to the optional Status argument and, when
'   enabled, to the Excel status bar.
'
' NOTE
'   Status-bar writes are gated behind PROB_WRITE_STATUS_BAR (default False).
'   Such writes from a worksheet UDF are unreliable (Excel frequently blocks
'   object-model access from a function evaluation) and add churn when a function
'   is bulk-filled. The On Error Resume Next guard wraps only the status-bar
'   write, so a failure there cannot mask a failure of the ByRef assignment.
'==============================================================================
'
'------------------------------------------------------------------------------
' UPDATE STATUS
'------------------------------------------------------------------------------
    'Write the ByRef diagnostic message; this must never be silently swallowed
        Status = Message

    'Exit when status-bar writes are disabled
        If Not PROB_WRITE_STATUS_BAR Then Exit Sub

'------------------------------------------------------------------------------
' UPDATE STATUS BAR
'------------------------------------------------------------------------------
    'Suppress non-critical status-bar side effects
        On Error Resume Next

    'Update Excel status bar
        If Len(Message) = 0 Then
            Application.StatusBar = False
        Else
            Application.StatusBar = Message
        End If

    'Restore normal error propagation
        On Error GoTo 0
End Sub
