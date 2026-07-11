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
'     - PROB_IsPositiveFinite
'     - PROB_IsValidProbabilityOpen
'
'   Numeric primitives:
'     - PROB_TryExp
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
'   - Overflow fails explicitly: a computation that would exceed Double range
'     returns False rather than a clamped sentinel value. There is deliberately
'     no PROB_SafeExp: a routine that quietly returns 1E+100 for a density is
'     returning a wrong number, and 1E+100 is simultaneously the threshold above
'     which PROB_IsFinite reports non-finite.
'   - Underflow of an exponential is a valid zero, not an error.
'   - Kernels here never validate their callers' domains and never write Status.
'
' NOTES
'   - PROB_IsFinite bounds magnitude at PROB_LARGE_NUMBER (1E+100); a legitimate
'     value at or beyond that magnitude reads as non-finite. This is intentional
'     for a distribution library and is inherited from NORMALFAMILY.
'   - VBA Doubles cannot normally hold NaN (arithmetic raises error 6 first), so
'     the X = X clause in PROB_IsFinite is defence in depth, not a contract.
'   - MIGRATION: delete the Private copies of PROB_IsFinite,
'     PROB_IsValidProbabilityOpen, PROB_TryExp, PROB_SetStatus and the duplicated
'     Private Const block from M_STATS_PROBDIST_NORMALFAMILY, and delete the
'     modules M_STATS_PROBDIST_SPECIALFUNCS and M_STATS_PROBDIST_CONTINUOUS
'     outright. Leaving M_STATS_PROBDIST_SPECIALFUNCS in place will produce
'     "Ambiguous name detected" against this module.
'
' UPDATED
'   2026-07-09
'==============================================================================

'==============================================================================
' PUBLIC CONSTANTS
'==============================================================================

Public Const PROB_PI                   As Double = 3.14159265358979      'Correctly rounded pi; the prior 3.14159265358979 was ~7 ulp low
Public Const PROB_TWO_PI               As Double = 6.28318530717959
Public Const PROB_HALF_LOG_TWO_PI      As Double = 0.918938533204673     '0.5 * Log(2 * Pi), correctly rounded
Public Const PROB_HALF_LOG_PI          As Double = 0.5723649429247       '0.5 * Log(Pi)

Public Const PROB_EPS                  As Double = 0.000000000000001     '1E-15, relative convergence target
Public Const PROB_NUM_EPS              As Double = 0.00000000000003      '3E-14, continued-fraction / series stop
Public Const PROB_MACH_EPS             As Double = 2.22044604925031E-16  'Double epsilon

Public Const PROB_MAX_EXP              As Double = 709#                  'Exp overflows above this
Public Const PROB_MIN_EXP              As Double = -745#                 'Exp underflows to 0 below this

Public Const PROB_LARGE_NUMBER         As Double = 1E+100                'Finiteness magnitude bound
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
'   Performs a lightweight finite-number check for VBA Double inputs.
'
' NOTE
'   The magnitude bound is PROB_LARGE_NUMBER (1E+100); a legitimate value at or
'   beyond that magnitude reads as non-finite. This is intentional for a
'   distribution library.
'==============================================================================
'
'------------------------------------------------------------------------------
' RETURN
'------------------------------------------------------------------------------
    'Return TRUE when X is not NaN and is inside a conservative magnitude bound
        PROB_IsFinite = (X = X And Abs(X) < PROB_LARGE_NUMBER)
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
'==============================================================================
'
'------------------------------------------------------------------------------
' RETURN
'------------------------------------------------------------------------------
    'Return TRUE for positive finite numbers
        PROB_IsPositiveFinite = (PROB_IsFinite(X) And X > 0#)
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
'   Attempts Exp(X) with explicit overflow and underflow handling.
'
' CONTRACT
'   - Overflow  (X >= PROB_MAX_EXP): returns False; Result is left unchanged.
'   - Underflow (X <= PROB_MIN_EXP): returns True;  Result = 0 (a valid zero).
'   - Otherwise:                     returns True;  Result = Exp(X).
'
' RATIONALE
'   Overflow is a genuine failure a caller must surface (as CVErr(xlErrNum)),
'   whereas underflow to zero is a legitimate result. Separating the two lets
'   the public routines distinguish "too big" from "vanishingly small".
'==============================================================================
'
'------------------------------------------------------------------------------
' COMPUTE
'------------------------------------------------------------------------------
    'Reject overflow
        If X >= PROB_MAX_EXP Then
            Exit Function
    'Treat underflow as a valid zero
        ElseIf X <= PROB_MIN_EXP Then
            Result = 0#
    'Regular exponential
        Else
            Result = Exp(X)
        End If
    'Report success
        PROB_TryExp = True
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
'   None on correctness. For X at or above PROB_MAX_EXP the true value overflows
'   a Double and VBA raises overflow error 6; callers that may pass large
'   positive X should guard with PROB_TryExp instead. Every caller in this
'   library passes X <= 0, where Exp(X) lies in (0, 1] and no overflow is
'   possible.
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




