Attribute VB_Name = "M_STATS_PROBDIST_REGISTER"
'==============================================================================
' Function Wizard metadata for the complete public K_STATS_* worksheet API.
' Adapted from danielep71/K-PRICING, KPR_REGISTER_PUBLIC_UDFS (MPL-2.0).
' Register/clear are repeatable, best effort, workbook-qualified, and restore
' ThisWorkbook.Saved. No calculation, selection, alerts or event changes.
' Clear resets help/category; Excel cannot unregister a loaded VBA function.
' MacroOptions has no metadata getter: counts report accepted calls only.
' The manifest is the authoritative help text; tests bind it to PUBLIC_API.txt.
' Call PROB_Register_PublicUDFs from Workbook_Open with macros enabled.
'==============================================================================
Option Explicit
Option Private Module

Private Const PROB_UDF_CATEGORY As String = "Probability Distributions"
Private Const USER_DEFINED_CATEGORY As Long = 14
Private mCount As Long
Private mNames() As String
Private mArgNames() As String
Private mDescriptions() As String
Private mArgDescs() As Variant
Private mArgCounts() As Long
Private mLastReport As String
Private mLastCount As Long

Public Function PROB_Register_LastCount() As Long
    PROB_Register_LastCount = mLastCount
End Function


Public Function PROB_Register_PublicUDFs() As Long
'
'==============================================================================
'                           PROB_Register_PublicUDFs
'------------------------------------------------------------------------------
' PURPOSE
'   Registers the description, category and argument help of every supported
'   function with Application.MacroOptions.
'
' RETURNS
'   The number of functions Excel accepted: the manifest count when all succeed. Failures are
'   listed by PROB_Register_LastReport.
'
' BEHAVIOR
'   Idempotent: each run overwrites the same metadata by function name.
'   ThisWorkbook.Saved is restored to the value found on entry.
'
' UPDATED
'   2026-11-17
'==============================================================================
'
    PROB_Register_PublicUDFs = ApplyAll(False)

End Function


Public Function PROB_Register_ClearPublicUDFs() As Long
'
'==============================================================================
'                         PROB_Register_ClearPublicUDFs
'------------------------------------------------------------------------------
' PURPOSE
'   Best-effort removal of the Function Wizard metadata this module registers.
'
' RETURNS
'   The number of functions whose metadata Excel accepted a reset for,
'   equal to the manifest count when all succeed. Failures are listed by
'   PROB_Register_LastReport.
'
' BEHAVIOR
'   Blanks each description and argument description and moves the function
'   to Excel's built-in "User Defined" category. Excel exposes no symmetric
'   unregister operation: the VBA functions are not removed and stay callable
'   while this project is loaded. Repeating the clean-up is harmless.
'
' UPDATED
'   2026-11-17
'==============================================================================
'
    PROB_Register_ClearPublicUDFs = ApplyAll(True)

End Function


Public Function PROB_Register_LastReport() As String
'
' One line per function that Excel rejected during the last register or
' clean-up run; empty when every call succeeded.
'
    PROB_Register_LastReport = mLastReport

End Function


Public Function PROB_Register_ManifestCount() As Long
'
' Number of manifest records.
'
    LoadManifest
    PROB_Register_ManifestCount = mCount

End Function


Public Function PROB_Register_ManifestName( _
    ByVal Index As Long) _
    As String
'
' Function name of manifest record Index (1-based); empty outside the manifest.
'
    LoadManifest
    If Index >= 1 And Index <= mCount Then PROB_Register_ManifestName = mNames(Index)

End Function


Public Function PROB_Register_ManifestArgCount( _
    ByVal Index As Long) _
    As Long
'
' Argument count of manifest record Index (1-based); -1 outside the manifest.
'
    LoadManifest
    If Index >= 1 And Index <= mCount Then
        PROB_Register_ManifestArgCount = mArgCounts(Index)
    Else
        PROB_Register_ManifestArgCount = -1
    End If

End Function


Public Function PROB_Register_ManifestArgDescriptions( _
    ByVal Index As Long) _
    As Variant
'
' Copy of the 1-based argument-description array of manifest record Index;
' Empty for a function without arguments or outside the manifest.
'
    LoadManifest
    If Index >= 1 And Index <= mCount Then PROB_Register_ManifestArgDescriptions = mArgDescs(Index)

End Function


'
'------------------------------------------------------------------------------
'
'                                   LIFECYCLE
'
'------------------------------------------------------------------------------
'

Private Function ApplyAll( _
    ByVal Clear As Boolean) _
    As Long
'
' Registers (Clear = False) or resets (Clear = True) every manifest record and
' returns how many Excel accepted. ThisWorkbook.Saved is restored on exit.
'

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Dim WasSaved        As Boolean      'ThisWorkbook.Saved found on entry
    Dim HaveSaved       As Boolean      'TRUE once WasSaved was read
    Dim Accepted        As Long         'Records Excel accepted
    Dim I               As Long         'Record cursor
    Dim Why             As String       'Failure description of one record

'------------------------------------------------------------------------------
' INITIALIZE
'------------------------------------------------------------------------------
    On Error GoTo Apply_Error
    mLastReport = vbNullString
    LoadManifest
    WasSaved = ThisWorkbook.Saved
    HaveSaved = True

'------------------------------------------------------------------------------
' APPLY
'------------------------------------------------------------------------------
    'Each record independently, so one rejection never blocks the others
        For I = 1 To mCount
            If TryApply(I, Clear, Why) Then
                Accepted = Accepted + 1
            Else
                Note mNames(I) & ": " & Why
            End If
        Next I

'------------------------------------------------------------------------------
' RESTORE
'------------------------------------------------------------------------------
Apply_Exit:
    'MacroOptions marks the workbook as changed; put back what the caller had
        If HaveSaved Then
            On Error Resume Next
            If WasSaved And Not ThisWorkbook.Saved Then ThisWorkbook.Saved = True
            If Err.Number <> 0 Then Note "ThisWorkbook.Saved: error " & CStr(Err.Number) & ": " & Err.Description
            Err.Clear
            On Error GoTo 0
        End If
    mLastCount = Accepted
    ApplyAll = Accepted
    Exit Function

Apply_Error:
    Note "lifecycle: error " & CStr(Err.Number) & ": " & Err.Description
    Resume Apply_Exit

End Function


Private Function TryApply( _
    ByVal Index As Long, _
    ByVal Clear As Boolean, _
    ByRef Why As String) _
    As Boolean
'
' Applies one record. The workbook-qualified macro name is tried first so the
' call cannot resolve to a same-named macro in another open workbook; the bare
' name is the fallback Excel accepts when this project is the active one.
'

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Dim ErrQualified    As String       'Outcome of the workbook-qualified call
    Dim ErrBare         As String       'Outcome of the bare-name call

'------------------------------------------------------------------------------
' APPLY
'------------------------------------------------------------------------------
    Why = vbNullString
    ErrQualified = CallMacroOptions(QualifiedName(mNames(Index)), Index, Clear)
    If Len(ErrQualified) = 0 Then
        TryApply = True
        Exit Function
    End If
    'A bare name is safe only when this workbook owns the active project.
    If Not Application.ActiveWorkbook Is ThisWorkbook Then
        Why = ErrQualified
        Exit Function
    End If
    ErrBare = CallMacroOptions(mNames(Index), Index, Clear)
    If Len(ErrBare) = 0 Then
        TryApply = True
    Else
        Why = ErrQualified & " (bare name: " & ErrBare & ")"
    End If

End Function


Private Function CallMacroOptions( _
    ByVal Target As String, _
    ByVal Index As Long, _
    ByVal Clear As Boolean) _
    As String
'
' One Application.MacroOptions call; returns "" on success, otherwise the
' error Excel raised. ArgumentDescriptions is omitted for a function without
' arguments, as an empty array is rejected by some Excel builds.
'

'------------------------------------------------------------------------------
' CALL
'------------------------------------------------------------------------------
    On Error GoTo Call_Error
    If Clear Then
        If mArgCounts(Index) > 0 Then
            Application.MacroOptions Macro:=Target, Description:=vbNullString, _
                                     Category:=USER_DEFINED_CATEGORY, _
                                     ArgumentDescriptions:=BlankDescriptions(mArgCounts(Index))
        Else
            Application.MacroOptions Macro:=Target, Description:=vbNullString, _
                                     Category:=USER_DEFINED_CATEGORY
        End If
    Else
        If mArgCounts(Index) > 0 Then
            Application.MacroOptions Macro:=Target, Description:=mDescriptions(Index), _
                                     Category:=PROB_UDF_CATEGORY, _
                                     ArgumentDescriptions:=mArgDescs(Index)
        Else
            Application.MacroOptions Macro:=Target, Description:=mDescriptions(Index), _
                                     Category:=PROB_UDF_CATEGORY
        End If
    End If
    Exit Function

Call_Error:
    CallMacroOptions = "error " & CStr(Err.Number) & ": " & Err.Description
    Err.Clear

End Function


Private Function QualifiedName( _
    ByVal FunctionName As String) _
    As String
'
' 'Workbook name'!FunctionName, with any apostrophe in the name doubled.
'
    QualifiedName = "'" & Replace(ThisWorkbook.Name, "'", "''") & "'!" & FunctionName

End Function


Private Function BlankDescriptions( _
    ByVal Total As Long) _
    As Variant
'
' A 1-based array of Total empty argument descriptions for the clean-up.
'

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Dim Blank()         As Variant      'Result
    Dim I               As Long         'Element cursor

'------------------------------------------------------------------------------
' BUILD
'------------------------------------------------------------------------------
    ReDim Blank(1 To Total)
    For I = 1 To Total
        Blank(I) = vbNullString
    Next I
    BlankDescriptions = Blank

End Function


Private Sub Note( _
    ByVal Entry As String)
'
' Appends one line to the report of the current run.
'
    If Len(mLastReport) > 0 Then mLastReport = mLastReport & vbCrLf
    mLastReport = mLastReport & Entry

End Sub


'
'------------------------------------------------------------------------------
'
'                                   MANIFEST
'
'------------------------------------------------------------------------------
'

Private Sub AddRecord( _
    ByVal FunctionName As String, _
    ByVal ArgNames As String, _
    ByVal Description As String, _
    ParamArray ArgDescriptions() As Variant)
'
' Appends one manifest record. The argument descriptions arrive as a 0-based
' ParamArray and are stored as the 1-based array MacroOptions receives.
'

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Dim Descs()         As Variant      'Argument descriptions, 1-based
    Dim Total           As Long         'Number of argument descriptions
    Dim I               As Long         'Element cursor

'------------------------------------------------------------------------------
' STORE
'------------------------------------------------------------------------------
    mCount = mCount + 1
    ReDim Preserve mNames(1 To mCount)
    ReDim Preserve mArgNames(1 To mCount)
    ReDim Preserve mDescriptions(1 To mCount)
    ReDim Preserve mArgDescs(1 To mCount)
    ReDim Preserve mArgCounts(1 To mCount)
    mNames(mCount) = FunctionName
    mArgNames(mCount) = ArgNames
    mDescriptions(mCount) = Description
    Total = UBound(ArgDescriptions) - LBound(ArgDescriptions) + 1
    mArgCounts(mCount) = Total
    If Total > 0 Then
        ReDim Descs(1 To Total)
        For I = 1 To Total
            Descs(I) = CStr(ArgDescriptions(LBound(ArgDescriptions) + I - 1))
        Next I
        mArgDescs(mCount) = Descs
    Else
        mArgDescs(mCount) = Empty
    End If

End Sub


Private Sub LoadManifest()
    mCount = 0
    Erase mNames
    Erase mArgNames
    Erase mDescriptions
    Erase mArgDescs
    Erase mArgCounts

    AddRecord "K_STATS_Beta_Cumulative", "X|Alpha|Beta|Status", _
        "Beta: Lower-tail cumulative probability P(X <= x).", _
        "X: Evaluation point x.", _
        "Alpha: Positive first Beta shape parameter.", _
        "Beta: Positive second Beta shape parameter.", _
        "Status: (optional, default """"): Optional ByRef diagnostic string for VBA callers; omit in worksheet formulas. Empty on success."
    AddRecord "K_STATS_Beta_Density", "X|Alpha|Beta|Status", _
        "Beta: Probability density at x.", _
        "X: Evaluation point x.", _
        "Alpha: Positive first Beta shape parameter.", _
        "Beta: Positive second Beta shape parameter.", _
        "Status: (optional, default """"): Optional ByRef diagnostic string for VBA callers; omit in worksheet formulas. Empty on success."
    AddRecord "K_STATS_Beta_InverseCumulative", "Probability|Alpha|Beta|Status", _
        "Beta: Lower-tail quantile for the supplied probability.", _
        "Probability: Lower-tail probability. Endpoint handling follows the function contract.", _
        "Alpha: Positive first Beta shape parameter.", _
        "Beta: Positive second Beta shape parameter.", _
        "Status: (optional, default """"): Optional ByRef diagnostic string for VBA callers; omit in worksheet formulas. Empty on success."
    AddRecord "K_STATS_Beta_Mean", "Alpha|Beta|Status", _
        "Beta: Arithmetic mean of the distribution.", _
        "Alpha: Positive first Beta shape parameter.", _
        "Beta: Positive second Beta shape parameter.", _
        "Status: (optional, default """"): Optional ByRef diagnostic string for VBA callers; omit in worksheet formulas. Empty on success."
    AddRecord "K_STATS_Beta_StdDev", "Alpha|Beta|Status", _
        "Beta: Standard deviation of the distribution.", _
        "Alpha: Positive first Beta shape parameter.", _
        "Beta: Positive second Beta shape parameter.", _
        "Status: (optional, default """"): Optional ByRef diagnostic string for VBA callers; omit in worksheet formulas. Empty on success."
    AddRecord "K_STATS_Beta_Survival", "X|Alpha|Beta|Status", _
        "Beta: Upper-tail probability P(X > x), evaluated directly.", _
        "X: Evaluation point x.", _
        "Alpha: Positive first Beta shape parameter.", _
        "Beta: Positive second Beta shape parameter.", _
        "Status: (optional, default """"): Optional ByRef diagnostic string for VBA callers; omit in worksheet formulas. Empty on success."
    AddRecord "K_STATS_Beta_Variance", "Alpha|Beta|Status", _
        "Beta: Variance of the distribution.", _
        "Alpha: Positive first Beta shape parameter.", _
        "Beta: Positive second Beta shape parameter.", _
        "Status: (optional, default """"): Optional ByRef diagnostic string for VBA callers; omit in worksheet formulas. Empty on success."
    AddRecord "K_STATS_Exponential_Cumulative", "X|Lambda|Status", _
        "Exponential: Lower-tail cumulative probability P(X <= x).", _
        "X: Evaluation point x.", _
        "Lambda: Positive exponential rate (inverse x units), not scale.", _
        "Status: (optional, default """"): Optional ByRef diagnostic string for VBA callers; omit in worksheet formulas. Empty on success."
    AddRecord "K_STATS_Exponential_Density", "X|Lambda|Status", _
        "Exponential: Probability density at x.", _
        "X: Evaluation point x.", _
        "Lambda: Positive exponential rate (inverse x units), not scale.", _
        "Status: (optional, default """"): Optional ByRef diagnostic string for VBA callers; omit in worksheet formulas. Empty on success."
    AddRecord "K_STATS_Exponential_InverseCumulative", "Probability|Lambda|Status", _
        "Exponential: Lower-tail quantile for the supplied probability.", _
        "Probability: Lower-tail probability. Endpoint handling follows the function contract.", _
        "Lambda: Positive exponential rate (inverse x units), not scale.", _
        "Status: (optional, default """"): Optional ByRef diagnostic string for VBA callers; omit in worksheet formulas. Empty on success."
    AddRecord "K_STATS_Exponential_Survival", "X|Lambda|Status", _
        "Exponential: Upper-tail probability P(X > x), evaluated directly.", _
        "X: Evaluation point x.", _
        "Lambda: Positive exponential rate (inverse x units), not scale.", _
        "Status: (optional, default """"): Optional ByRef diagnostic string for VBA callers; omit in worksheet formulas. Empty on success."
    AddRecord "K_STATS_Gamma_Cumulative", "X|Shape|ScaleParam|Status", _
        "Gamma: Lower-tail cumulative probability P(X <= x).", _
        "X: Evaluation point x.", _
        "Shape: Positive shape parameter.", _
        "ScaleParam: Positive scale parameter in x units, not rate.", _
        "Status: (optional, default """"): Optional ByRef diagnostic string for VBA callers; omit in worksheet formulas. Empty on success."
    AddRecord "K_STATS_Gamma_Density", "X|Shape|ScaleParam|Status", _
        "Gamma: Probability density at x.", _
        "X: Evaluation point x.", _
        "Shape: Positive shape parameter.", _
        "ScaleParam: Positive scale parameter in x units, not rate.", _
        "Status: (optional, default """"): Optional ByRef diagnostic string for VBA callers; omit in worksheet formulas. Empty on success."
    AddRecord "K_STATS_Gamma_InverseCumulative", "Probability|Shape|ScaleParam|Status", _
        "Gamma: Lower-tail quantile for the supplied probability.", _
        "Probability: Lower-tail probability. Endpoint handling follows the function contract.", _
        "Shape: Positive shape parameter.", _
        "ScaleParam: Positive scale parameter in x units, not rate.", _
        "Status: (optional, default """"): Optional ByRef diagnostic string for VBA callers; omit in worksheet formulas. Empty on success."
    AddRecord "K_STATS_Gamma_Mean", "Shape|ScaleParam|Status", _
        "Gamma: Arithmetic mean of the distribution.", _
        "Shape: Positive shape parameter.", _
        "ScaleParam: Positive scale parameter in x units, not rate.", _
        "Status: (optional, default """"): Optional ByRef diagnostic string for VBA callers; omit in worksheet formulas. Empty on success."
    AddRecord "K_STATS_Gamma_StdDev", "Shape|ScaleParam|Status", _
        "Gamma: Standard deviation of the distribution.", _
        "Shape: Positive shape parameter.", _
        "ScaleParam: Positive scale parameter in x units, not rate.", _
        "Status: (optional, default """"): Optional ByRef diagnostic string for VBA callers; omit in worksheet formulas. Empty on success."
    AddRecord "K_STATS_Gamma_Survival", "X|Shape|ScaleParam|Status", _
        "Gamma: Upper-tail probability P(X > x), evaluated directly.", _
        "X: Evaluation point x.", _
        "Shape: Positive shape parameter.", _
        "ScaleParam: Positive scale parameter in x units, not rate.", _
        "Status: (optional, default """"): Optional ByRef diagnostic string for VBA callers; omit in worksheet formulas. Empty on success."
    AddRecord "K_STATS_Gamma_Variance", "Shape|ScaleParam|Status", _
        "Gamma: Variance of the distribution.", _
        "Shape: Positive shape parameter.", _
        "ScaleParam: Positive scale parameter in x units, not rate.", _
        "Status: (optional, default """"): Optional ByRef diagnostic string for VBA callers; omit in worksheet formulas. Empty on success."
    AddRecord "K_STATS_Uniform_Cumulative", "X|LowerBound|UpperBound|Status", _
        "Uniform: Lower-tail cumulative probability P(X <= x).", _
        "X: Evaluation point x.", _
        "LowerBound: Lower support bound; UpperBound must exceed LowerBound.", _
        "UpperBound: Upper support bound; UpperBound must exceed LowerBound.", _
        "Status: (optional, default """"): Optional ByRef diagnostic string for VBA callers; omit in worksheet formulas. Empty on success."
    AddRecord "K_STATS_Uniform_Density", "X|LowerBound|UpperBound|Status", _
        "Uniform: Probability density at x.", _
        "X: Evaluation point x.", _
        "LowerBound: Lower support bound; UpperBound must exceed LowerBound.", _
        "UpperBound: Upper support bound; UpperBound must exceed LowerBound.", _
        "Status: (optional, default """"): Optional ByRef diagnostic string for VBA callers; omit in worksheet formulas. Empty on success."
    AddRecord "K_STATS_Uniform_InverseCumulative", "Probability|LowerBound|UpperBound|Status", _
        "Uniform: Lower-tail quantile for the supplied probability.", _
        "Probability: Lower-tail probability. Endpoint handling follows the function contract.", _
        "LowerBound: Lower support bound; UpperBound must exceed LowerBound.", _
        "UpperBound: Upper support bound; UpperBound must exceed LowerBound.", _
        "Status: (optional, default """"): Optional ByRef diagnostic string for VBA callers; omit in worksheet formulas. Empty on success."
    AddRecord "K_STATS_Uniform_Survival", "X|LowerBound|UpperBound|Status", _
        "Uniform: Upper-tail probability P(X > x), evaluated directly.", _
        "X: Evaluation point x.", _
        "LowerBound: Lower support bound; UpperBound must exceed LowerBound.", _
        "UpperBound: Upper support bound; UpperBound must exceed LowerBound.", _
        "Status: (optional, default """"): Optional ByRef diagnostic string for VBA callers; omit in worksheet formulas. Empty on success."
    AddRecord "K_STATS_Weibull_Cumulative", "X|Shape|ScaleParam|Status", _
        "Weibull: Lower-tail cumulative probability P(X <= x).", _
        "X: Evaluation point x.", _
        "Shape: Positive shape parameter.", _
        "ScaleParam: Positive scale parameter in x units, not rate.", _
        "Status: (optional, default """"): Optional ByRef diagnostic string for VBA callers; omit in worksheet formulas. Empty on success."
    AddRecord "K_STATS_Weibull_Density", "X|Shape|ScaleParam|Status", _
        "Weibull: Probability density at x.", _
        "X: Evaluation point x.", _
        "Shape: Positive shape parameter.", _
        "ScaleParam: Positive scale parameter in x units, not rate.", _
        "Status: (optional, default """"): Optional ByRef diagnostic string for VBA callers; omit in worksheet formulas. Empty on success."
    AddRecord "K_STATS_Weibull_InverseCumulative", "Probability|Shape|ScaleParam|Status", _
        "Weibull: Lower-tail quantile for the supplied probability.", _
        "Probability: Lower-tail probability. Endpoint handling follows the function contract.", _
        "Shape: Positive shape parameter.", _
        "ScaleParam: Positive scale parameter in x units, not rate.", _
        "Status: (optional, default """"): Optional ByRef diagnostic string for VBA callers; omit in worksheet formulas. Empty on success."
    AddRecord "K_STATS_Weibull_Mean", "Shape|ScaleParam|Status", _
        "Weibull: Arithmetic mean of the distribution.", _
        "Shape: Positive shape parameter.", _
        "ScaleParam: Positive scale parameter in x units, not rate.", _
        "Status: (optional, default """"): Optional ByRef diagnostic string for VBA callers; omit in worksheet formulas. Empty on success."
    AddRecord "K_STATS_Weibull_StdDev", "Shape|ScaleParam|Status", _
        "Weibull: Standard deviation of the distribution.", _
        "Shape: Positive shape parameter.", _
        "ScaleParam: Positive scale parameter in x units, not rate.", _
        "Status: (optional, default """"): Optional ByRef diagnostic string for VBA callers; omit in worksheet formulas. Empty on success."
    AddRecord "K_STATS_Weibull_Survival", "X|Shape|ScaleParam|Status", _
        "Weibull: Upper-tail probability P(X > x), evaluated directly.", _
        "X: Evaluation point x.", _
        "Shape: Positive shape parameter.", _
        "ScaleParam: Positive scale parameter in x units, not rate.", _
        "Status: (optional, default """"): Optional ByRef diagnostic string for VBA callers; omit in worksheet formulas. Empty on success."
    AddRecord "K_STATS_Weibull_Variance", "Shape|ScaleParam|Status", _
        "Weibull: Variance of the distribution.", _
        "Shape: Positive shape parameter.", _
        "ScaleParam: Positive scale parameter in x units, not rate.", _
        "Status: (optional, default """"): Optional ByRef diagnostic string for VBA callers; omit in worksheet formulas. Empty on success."
    AddRecord "K_STATS_Binomial_Cumulative", "NumberSuccesses|Trials|ProbSuccess|Status", _
        "Binomial: Lower-tail cumulative probability P(X <= x).", _
        "NumberSuccesses: Success count; fractional values are truncated toward zero.", _
        "Trials: Trial count; fractional values are truncated toward zero.", _
        "ProbSuccess: Probability of success per trial; valid endpoints depend on the operation.", _
        "Status: (optional, default """"): Optional ByRef diagnostic string for VBA callers; omit in worksheet formulas. Empty on success."
    AddRecord "K_STATS_Binomial_InverseCumulative", "Probability|Trials|ProbSuccess|Status", _
        "Binomial: Lower-tail quantile for the supplied probability.", _
        "Probability: Lower-tail probability. Endpoint handling follows the function contract.", _
        "Trials: Trial count; fractional values are truncated toward zero.", _
        "ProbSuccess: Probability of success per trial; valid endpoints depend on the operation.", _
        "Status: (optional, default """"): Optional ByRef diagnostic string for VBA callers; omit in worksheet formulas. Empty on success."
    AddRecord "K_STATS_Binomial_LogPMF", "NumberSuccesses|Trials|ProbSuccess|Status", _
        "Binomial: Natural logarithm of the probability mass P(X = k).", _
        "NumberSuccesses: Success count; fractional values are truncated toward zero.", _
        "Trials: Trial count; fractional values are truncated toward zero.", _
        "ProbSuccess: Probability of success per trial; valid endpoints depend on the operation.", _
        "Status: (optional, default """"): Optional ByRef diagnostic string for VBA callers; omit in worksheet formulas. Empty on success."
    AddRecord "K_STATS_Binomial_Mean", "Trials|ProbSuccess|Status", _
        "Binomial: Arithmetic mean of the distribution.", _
        "Trials: Trial count; fractional values are truncated toward zero.", _
        "ProbSuccess: Probability of success per trial; valid endpoints depend on the operation.", _
        "Status: (optional, default """"): Optional ByRef diagnostic string for VBA callers; omit in worksheet formulas. Empty on success."
    AddRecord "K_STATS_Binomial_PMF", "NumberSuccesses|Trials|ProbSuccess|Status", _
        "Binomial: Probability mass P(X = k).", _
        "NumberSuccesses: Success count; fractional values are truncated toward zero.", _
        "Trials: Trial count; fractional values are truncated toward zero.", _
        "ProbSuccess: Probability of success per trial; valid endpoints depend on the operation.", _
        "Status: (optional, default """"): Optional ByRef diagnostic string for VBA callers; omit in worksheet formulas. Empty on success."
    AddRecord "K_STATS_Binomial_StdDev", "Trials|ProbSuccess|Status", _
        "Binomial: Standard deviation of the distribution.", _
        "Trials: Trial count; fractional values are truncated toward zero.", _
        "ProbSuccess: Probability of success per trial; valid endpoints depend on the operation.", _
        "Status: (optional, default """"): Optional ByRef diagnostic string for VBA callers; omit in worksheet formulas. Empty on success."
    AddRecord "K_STATS_Binomial_Survival", "NumberSuccesses|Trials|ProbSuccess|Status", _
        "Binomial: Upper-tail probability P(X > x), evaluated directly.", _
        "NumberSuccesses: Success count; fractional values are truncated toward zero.", _
        "Trials: Trial count; fractional values are truncated toward zero.", _
        "ProbSuccess: Probability of success per trial; valid endpoints depend on the operation.", _
        "Status: (optional, default """"): Optional ByRef diagnostic string for VBA callers; omit in worksheet formulas. Empty on success."
    AddRecord "K_STATS_Binomial_Variance", "Trials|ProbSuccess|Status", _
        "Binomial: Variance of the distribution.", _
        "Trials: Trial count; fractional values are truncated toward zero.", _
        "ProbSuccess: Probability of success per trial; valid endpoints depend on the operation.", _
        "Status: (optional, default """"): Optional ByRef diagnostic string for VBA callers; omit in worksheet formulas. Empty on success."
    AddRecord "K_STATS_DiscreteUniform_Cumulative", "X|LowerBound|UpperBound|Status", _
        "Discrete uniform: Lower-tail cumulative probability P(X <= x).", _
        "X: Evaluation count; fractional values are truncated toward zero.", _
        "LowerBound: Lower inclusive integer support bound; fractional values truncated toward zero.", _
        "UpperBound: Upper inclusive integer support bound; fractional values truncated toward zero.", _
        "Status: (optional, default """"): Optional ByRef diagnostic string for VBA callers; omit in worksheet formulas. Empty on success."
    AddRecord "K_STATS_DiscreteUniform_InverseCumulative", "Probability|LowerBound|UpperBound|Status", _
        "Discrete uniform: Lower-tail quantile for the supplied probability.", _
        "Probability: Lower-tail probability. Endpoint handling follows the function contract.", _
        "LowerBound: Lower inclusive integer support bound; fractional values truncated toward zero.", _
        "UpperBound: Upper inclusive integer support bound; fractional values truncated toward zero.", _
        "Status: (optional, default """"): Optional ByRef diagnostic string for VBA callers; omit in worksheet formulas. Empty on success."
    AddRecord "K_STATS_DiscreteUniform_LogPMF", "X|LowerBound|UpperBound|Status", _
        "Discrete uniform: Natural logarithm of the probability mass P(X = k).", _
        "X: Evaluation count; fractional values are truncated toward zero.", _
        "LowerBound: Lower inclusive integer support bound; fractional values truncated toward zero.", _
        "UpperBound: Upper inclusive integer support bound; fractional values truncated toward zero.", _
        "Status: (optional, default """"): Optional ByRef diagnostic string for VBA callers; omit in worksheet formulas. Empty on success."
    AddRecord "K_STATS_DiscreteUniform_Mean", "LowerBound|UpperBound|Status", _
        "Discrete uniform: Arithmetic mean of the distribution.", _
        "LowerBound: Lower inclusive integer support bound; fractional values truncated toward zero.", _
        "UpperBound: Upper inclusive integer support bound; fractional values truncated toward zero.", _
        "Status: (optional, default """"): Optional ByRef diagnostic string for VBA callers; omit in worksheet formulas. Empty on success."
    AddRecord "K_STATS_DiscreteUniform_PMF", "X|LowerBound|UpperBound|Status", _
        "Discrete uniform: Probability mass P(X = k).", _
        "X: Evaluation count; fractional values are truncated toward zero.", _
        "LowerBound: Lower inclusive integer support bound; fractional values truncated toward zero.", _
        "UpperBound: Upper inclusive integer support bound; fractional values truncated toward zero.", _
        "Status: (optional, default """"): Optional ByRef diagnostic string for VBA callers; omit in worksheet formulas. Empty on success."
    AddRecord "K_STATS_DiscreteUniform_StdDev", "LowerBound|UpperBound|Status", _
        "Discrete uniform: Standard deviation of the distribution.", _
        "LowerBound: Lower inclusive integer support bound; fractional values truncated toward zero.", _
        "UpperBound: Upper inclusive integer support bound; fractional values truncated toward zero.", _
        "Status: (optional, default """"): Optional ByRef diagnostic string for VBA callers; omit in worksheet formulas. Empty on success."
    AddRecord "K_STATS_DiscreteUniform_Survival", "X|LowerBound|UpperBound|Status", _
        "Discrete uniform: Upper-tail probability P(X > x), evaluated directly.", _
        "X: Evaluation count; fractional values are truncated toward zero.", _
        "LowerBound: Lower inclusive integer support bound; fractional values truncated toward zero.", _
        "UpperBound: Upper inclusive integer support bound; fractional values truncated toward zero.", _
        "Status: (optional, default """"): Optional ByRef diagnostic string for VBA callers; omit in worksheet formulas. Empty on success."
    AddRecord "K_STATS_DiscreteUniform_Variance", "LowerBound|UpperBound|Status", _
        "Discrete uniform: Variance of the distribution.", _
        "LowerBound: Lower inclusive integer support bound; fractional values truncated toward zero.", _
        "UpperBound: Upper inclusive integer support bound; fractional values truncated toward zero.", _
        "Status: (optional, default """"): Optional ByRef diagnostic string for VBA callers; omit in worksheet formulas. Empty on success."
    AddRecord "K_STATS_Geometric_Cumulative", "NumberFailures|ProbSuccess|Status", _
        "Geometric: Lower-tail cumulative probability P(X <= x). X counts failures before the first success.", _
        "NumberFailures: Failures before the required success; fractional values are truncated toward zero.", _
        "ProbSuccess: Probability of success per trial; valid endpoints depend on the operation.", _
        "Status: (optional, default """"): Optional ByRef diagnostic string for VBA callers; omit in worksheet formulas. Empty on success."
    AddRecord "K_STATS_Geometric_InverseCumulative", "Probability|ProbSuccess|Status", _
        "Geometric: Lower-tail quantile for the supplied probability. X counts failures before the first success.", _
        "Probability: Lower-tail probability. Endpoint handling follows the function contract.", _
        "ProbSuccess: Probability of success per trial; valid endpoints depend on the operation.", _
        "Status: (optional, default """"): Optional ByRef diagnostic string for VBA callers; omit in worksheet formulas. Empty on success."
    AddRecord "K_STATS_Geometric_LogPMF", "NumberFailures|ProbSuccess|Status", _
        "Geometric: Natural logarithm of the probability mass P(X = k). X counts failures before the first success.", _
        "NumberFailures: Failures before the required success; fractional values are truncated toward zero.", _
        "ProbSuccess: Probability of success per trial; valid endpoints depend on the operation.", _
        "Status: (optional, default """"): Optional ByRef diagnostic string for VBA callers; omit in worksheet formulas. Empty on success."
    AddRecord "K_STATS_Geometric_Mean", "ProbSuccess|Status", _
        "Geometric: Arithmetic mean of the distribution. X counts failures before the first success.", _
        "ProbSuccess: Probability of success per trial; valid endpoints depend on the operation.", _
        "Status: (optional, default """"): Optional ByRef diagnostic string for VBA callers; omit in worksheet formulas. Empty on success."
    AddRecord "K_STATS_Geometric_PMF", "NumberFailures|ProbSuccess|Status", _
        "Geometric: Probability mass P(X = k). X counts failures before the first success.", _
        "NumberFailures: Failures before the required success; fractional values are truncated toward zero.", _
        "ProbSuccess: Probability of success per trial; valid endpoints depend on the operation.", _
        "Status: (optional, default """"): Optional ByRef diagnostic string for VBA callers; omit in worksheet formulas. Empty on success."
    AddRecord "K_STATS_Geometric_StdDev", "ProbSuccess|Status", _
        "Geometric: Standard deviation of the distribution. X counts failures before the first success.", _
        "ProbSuccess: Probability of success per trial; valid endpoints depend on the operation.", _
        "Status: (optional, default """"): Optional ByRef diagnostic string for VBA callers; omit in worksheet formulas. Empty on success."
    AddRecord "K_STATS_Geometric_Survival", "NumberFailures|ProbSuccess|Status", _
        "Geometric: Upper-tail probability P(X > x), evaluated directly. X counts failures before the first success.", _
        "NumberFailures: Failures before the required success; fractional values are truncated toward zero.", _
        "ProbSuccess: Probability of success per trial; valid endpoints depend on the operation.", _
        "Status: (optional, default """"): Optional ByRef diagnostic string for VBA callers; omit in worksheet formulas. Empty on success."
    AddRecord "K_STATS_Geometric_Variance", "ProbSuccess|Status", _
        "Geometric: Variance of the distribution. X counts failures before the first success.", _
        "ProbSuccess: Probability of success per trial; valid endpoints depend on the operation.", _
        "Status: (optional, default """"): Optional ByRef diagnostic string for VBA callers; omit in worksheet formulas. Empty on success."
    AddRecord "K_STATS_Hypergeometric_Cumulative", "SampleSuccesses|SampleSize|PopulationSuccesses|PopulationSize|Status", _
        "Hypergeometric: Lower-tail cumulative probability P(X <= x).", _
        "SampleSuccesses: Successes in the sample; fractional counts are truncated toward zero.", _
        "SampleSize: Number drawn without replacement; fractional counts are truncated toward zero.", _
        "PopulationSuccesses: Successes in the population; fractional counts are truncated toward zero.", _
        "PopulationSize: Population size; fractional counts are truncated toward zero.", _
        "Status: (optional, default """"): Optional ByRef diagnostic string for VBA callers; omit in worksheet formulas. Empty on success."
    AddRecord "K_STATS_Hypergeometric_InverseCumulative", "Probability|SampleSize|PopulationSuccesses|PopulationSize|Status", _
        "Hypergeometric: Lower-tail quantile for the supplied probability.", _
        "Probability: Lower-tail probability. Endpoint handling follows the function contract.", _
        "SampleSize: Number drawn without replacement; fractional counts are truncated toward zero.", _
        "PopulationSuccesses: Successes in the population; fractional counts are truncated toward zero.", _
        "PopulationSize: Population size; fractional counts are truncated toward zero.", _
        "Status: (optional, default """"): Optional ByRef diagnostic string for VBA callers; omit in worksheet formulas. Empty on success."
    AddRecord "K_STATS_Hypergeometric_LogPMF", "SampleSuccesses|SampleSize|PopulationSuccesses|PopulationSize|Status", _
        "Hypergeometric: Natural logarithm of the probability mass P(X = k).", _
        "SampleSuccesses: Successes in the sample; fractional counts are truncated toward zero.", _
        "SampleSize: Number drawn without replacement; fractional counts are truncated toward zero.", _
        "PopulationSuccesses: Successes in the population; fractional counts are truncated toward zero.", _
        "PopulationSize: Population size; fractional counts are truncated toward zero.", _
        "Status: (optional, default """"): Optional ByRef diagnostic string for VBA callers; omit in worksheet formulas. Empty on success."
    AddRecord "K_STATS_Hypergeometric_Mean", "SampleSize|PopulationSuccesses|PopulationSize|Status", _
        "Hypergeometric: Arithmetic mean of the distribution.", _
        "SampleSize: Number drawn without replacement; fractional counts are truncated toward zero.", _
        "PopulationSuccesses: Successes in the population; fractional counts are truncated toward zero.", _
        "PopulationSize: Population size; fractional counts are truncated toward zero.", _
        "Status: (optional, default """"): Optional ByRef diagnostic string for VBA callers; omit in worksheet formulas. Empty on success."
    AddRecord "K_STATS_Hypergeometric_PMF", "SampleSuccesses|SampleSize|PopulationSuccesses|PopulationSize|Status", _
        "Hypergeometric: Probability mass P(X = k).", _
        "SampleSuccesses: Successes in the sample; fractional counts are truncated toward zero.", _
        "SampleSize: Number drawn without replacement; fractional counts are truncated toward zero.", _
        "PopulationSuccesses: Successes in the population; fractional counts are truncated toward zero.", _
        "PopulationSize: Population size; fractional counts are truncated toward zero.", _
        "Status: (optional, default """"): Optional ByRef diagnostic string for VBA callers; omit in worksheet formulas. Empty on success."
    AddRecord "K_STATS_Hypergeometric_StdDev", "SampleSize|PopulationSuccesses|PopulationSize|Status", _
        "Hypergeometric: Standard deviation of the distribution.", _
        "SampleSize: Number drawn without replacement; fractional counts are truncated toward zero.", _
        "PopulationSuccesses: Successes in the population; fractional counts are truncated toward zero.", _
        "PopulationSize: Population size; fractional counts are truncated toward zero.", _
        "Status: (optional, default """"): Optional ByRef diagnostic string for VBA callers; omit in worksheet formulas. Empty on success."
    AddRecord "K_STATS_Hypergeometric_Survival", "SampleSuccesses|SampleSize|PopulationSuccesses|PopulationSize|Status", _
        "Hypergeometric: Upper-tail probability P(X > x), evaluated directly.", _
        "SampleSuccesses: Successes in the sample; fractional counts are truncated toward zero.", _
        "SampleSize: Number drawn without replacement; fractional counts are truncated toward zero.", _
        "PopulationSuccesses: Successes in the population; fractional counts are truncated toward zero.", _
        "PopulationSize: Population size; fractional counts are truncated toward zero.", _
        "Status: (optional, default """"): Optional ByRef diagnostic string for VBA callers; omit in worksheet formulas. Empty on success."
    AddRecord "K_STATS_Hypergeometric_Variance", "SampleSize|PopulationSuccesses|PopulationSize|Status", _
        "Hypergeometric: Variance of the distribution.", _
        "SampleSize: Number drawn without replacement; fractional counts are truncated toward zero.", _
        "PopulationSuccesses: Successes in the population; fractional counts are truncated toward zero.", _
        "PopulationSize: Population size; fractional counts are truncated toward zero.", _
        "Status: (optional, default """"): Optional ByRef diagnostic string for VBA callers; omit in worksheet formulas. Empty on success."
    AddRecord "K_STATS_NegativeBinomial_Cumulative", "NumberFailures|NumberSuccesses|ProbSuccess|Status", _
        "Negative binomial: Lower-tail cumulative probability P(X <= x). X counts failures before NumberSuccesses successes.", _
        "NumberFailures: Failures before the required success; fractional values are truncated toward zero.", _
        "NumberSuccesses: Required success count, positive; fractional values are truncated toward zero.", _
        "ProbSuccess: Probability of success per trial; valid endpoints depend on the operation.", _
        "Status: (optional, default """"): Optional ByRef diagnostic string for VBA callers; omit in worksheet formulas. Empty on success."
    AddRecord "K_STATS_NegativeBinomial_InverseCumulative", "Probability|NumberSuccesses|ProbSuccess|Status", _
        "Negative binomial: Lower-tail quantile for the supplied probability. X counts failures before NumberSuccesses successes.", _
        "Probability: Lower-tail probability. Endpoint handling follows the function contract.", _
        "NumberSuccesses: Required success count, positive; fractional values are truncated toward zero.", _
        "ProbSuccess: Probability of success per trial; valid endpoints depend on the operation.", _
        "Status: (optional, default """"): Optional ByRef diagnostic string for VBA callers; omit in worksheet formulas. Empty on success."
    AddRecord "K_STATS_NegativeBinomial_LogPMF", "NumberFailures|NumberSuccesses|ProbSuccess|Status", _
        "Negative binomial: Natural logarithm of the probability mass P(X = k). X counts failures before NumberSuccesses successes.", _
        "NumberFailures: Failures before the required success; fractional values are truncated toward zero.", _
        "NumberSuccesses: Required success count, positive; fractional values are truncated toward zero.", _
        "ProbSuccess: Probability of success per trial; valid endpoints depend on the operation.", _
        "Status: (optional, default """"): Optional ByRef diagnostic string for VBA callers; omit in worksheet formulas. Empty on success."
    AddRecord "K_STATS_NegativeBinomial_Mean", "NumberSuccesses|ProbSuccess|Status", _
        "Negative binomial: Arithmetic mean of the distribution. X counts failures before NumberSuccesses successes.", _
        "NumberSuccesses: Required success count, positive; fractional values are truncated toward zero.", _
        "ProbSuccess: Probability of success per trial; valid endpoints depend on the operation.", _
        "Status: (optional, default """"): Optional ByRef diagnostic string for VBA callers; omit in worksheet formulas. Empty on success."
    AddRecord "K_STATS_NegativeBinomial_PMF", "NumberFailures|NumberSuccesses|ProbSuccess|Status", _
        "Negative binomial: Probability mass P(X = k). X counts failures before NumberSuccesses successes.", _
        "NumberFailures: Failures before the required success; fractional values are truncated toward zero.", _
        "NumberSuccesses: Required success count, positive; fractional values are truncated toward zero.", _
        "ProbSuccess: Probability of success per trial; valid endpoints depend on the operation.", _
        "Status: (optional, default """"): Optional ByRef diagnostic string for VBA callers; omit in worksheet formulas. Empty on success."
    AddRecord "K_STATS_NegativeBinomial_StdDev", "NumberSuccesses|ProbSuccess|Status", _
        "Negative binomial: Standard deviation of the distribution. X counts failures before NumberSuccesses successes.", _
        "NumberSuccesses: Required success count, positive; fractional values are truncated toward zero.", _
        "ProbSuccess: Probability of success per trial; valid endpoints depend on the operation.", _
        "Status: (optional, default """"): Optional ByRef diagnostic string for VBA callers; omit in worksheet formulas. Empty on success."
    AddRecord "K_STATS_NegativeBinomial_Survival", "NumberFailures|NumberSuccesses|ProbSuccess|Status", _
        "Negative binomial: Upper-tail probability P(X > x), evaluated directly. X counts failures before NumberSuccesses successes.", _
        "NumberFailures: Failures before the required success; fractional values are truncated toward zero.", _
        "NumberSuccesses: Required success count, positive; fractional values are truncated toward zero.", _
        "ProbSuccess: Probability of success per trial; valid endpoints depend on the operation.", _
        "Status: (optional, default """"): Optional ByRef diagnostic string for VBA callers; omit in worksheet formulas. Empty on success."
    AddRecord "K_STATS_NegativeBinomial_Variance", "NumberSuccesses|ProbSuccess|Status", _
        "Negative binomial: Variance of the distribution. X counts failures before NumberSuccesses successes.", _
        "NumberSuccesses: Required success count, positive; fractional values are truncated toward zero.", _
        "ProbSuccess: Probability of success per trial; valid endpoints depend on the operation.", _
        "Status: (optional, default """"): Optional ByRef diagnostic string for VBA callers; omit in worksheet formulas. Empty on success."
    AddRecord "K_STATS_Poisson_Cumulative", "NumberEvents|Mean|Status", _
        "Poisson: Lower-tail cumulative probability P(X <= x).", _
        "NumberEvents: Poisson event count; fractional values are truncated toward zero.", _
        "Mean: Non-negative Poisson intensity (expected event count).", _
        "Status: (optional, default """"): Optional ByRef diagnostic string for VBA callers; omit in worksheet formulas. Empty on success."
    AddRecord "K_STATS_Poisson_InverseCumulative", "Probability|Mean|Status", _
        "Poisson: Lower-tail quantile for the supplied probability.", _
        "Probability: Lower-tail probability. Endpoint handling follows the function contract.", _
        "Mean: Non-negative Poisson intensity (expected event count).", _
        "Status: (optional, default """"): Optional ByRef diagnostic string for VBA callers; omit in worksheet formulas. Empty on success."
    AddRecord "K_STATS_Poisson_LogPMF", "NumberEvents|Mean|Status", _
        "Poisson: Natural logarithm of the probability mass P(X = k).", _
        "NumberEvents: Poisson event count; fractional values are truncated toward zero.", _
        "Mean: Non-negative Poisson intensity (expected event count).", _
        "Status: (optional, default """"): Optional ByRef diagnostic string for VBA callers; omit in worksheet formulas. Empty on success."
    AddRecord "K_STATS_Poisson_Mean", "Mean|Status", _
        "Poisson: Arithmetic mean of the distribution.", _
        "Mean: Non-negative Poisson intensity (expected event count).", _
        "Status: (optional, default """"): Optional ByRef diagnostic string for VBA callers; omit in worksheet formulas. Empty on success."
    AddRecord "K_STATS_Poisson_PMF", "NumberEvents|Mean|Status", _
        "Poisson: Probability mass P(X = k).", _
        "NumberEvents: Poisson event count; fractional values are truncated toward zero.", _
        "Mean: Non-negative Poisson intensity (expected event count).", _
        "Status: (optional, default """"): Optional ByRef diagnostic string for VBA callers; omit in worksheet formulas. Empty on success."
    AddRecord "K_STATS_Poisson_StdDev", "Mean|Status", _
        "Poisson: Standard deviation of the distribution.", _
        "Mean: Non-negative Poisson intensity (expected event count).", _
        "Status: (optional, default """"): Optional ByRef diagnostic string for VBA callers; omit in worksheet formulas. Empty on success."
    AddRecord "K_STATS_Poisson_Survival", "NumberEvents|Mean|Status", _
        "Poisson: Upper-tail probability P(X > x), evaluated directly.", _
        "NumberEvents: Poisson event count; fractional values are truncated toward zero.", _
        "Mean: Non-negative Poisson intensity (expected event count).", _
        "Status: (optional, default """"): Optional ByRef diagnostic string for VBA callers; omit in worksheet formulas. Empty on success."
    AddRecord "K_STATS_Poisson_Variance", "Mean|Status", _
        "Poisson: Variance of the distribution.", _
        "Mean: Non-negative Poisson intensity (expected event count).", _
        "Status: (optional, default """"): Optional ByRef diagnostic string for VBA callers; omit in worksheet formulas. Empty on success."
    AddRecord "K_STATS_Lognormal_Cumulative", "X|MeanLog|StdDevLog|Status", _
        "Lognormal: Lower-tail cumulative probability P(X <= x).", _
        "X: Evaluation point x.", _
        "MeanLog: Mean of the natural logarithm of X, not the arithmetic mean.", _
        "StdDevLog: Positive standard deviation of the natural logarithm of X.", _
        "Status: (optional, default """"): Optional ByRef diagnostic string for VBA callers; omit in worksheet formulas. Empty on success."
    AddRecord "K_STATS_Lognormal_Density", "X|MeanLog|StdDevLog|Status", _
        "Lognormal: Probability density at x.", _
        "X: Evaluation point x.", _
        "MeanLog: Mean of the natural logarithm of X, not the arithmetic mean.", _
        "StdDevLog: Positive standard deviation of the natural logarithm of X.", _
        "Status: (optional, default """"): Optional ByRef diagnostic string for VBA callers; omit in worksheet formulas. Empty on success."
    AddRecord "K_STATS_Lognormal_InverseCumulative", "Probability|MeanLog|StdDevLog|Status", _
        "Lognormal: Lower-tail quantile for the supplied probability.", _
        "Probability: Lower-tail probability. Endpoint handling follows the function contract.", _
        "MeanLog: Mean of the natural logarithm of X, not the arithmetic mean.", _
        "StdDevLog: Positive standard deviation of the natural logarithm of X.", _
        "Status: (optional, default """"): Optional ByRef diagnostic string for VBA callers; omit in worksheet formulas. Empty on success."
    AddRecord "K_STATS_Lognormal_InverseSurvival", "Probability|MeanLog|StdDevLog|Status", _
        "Lognormal: Upper-tail quantile for the supplied survival probability.", _
        "Probability: Upper-tail probability P(X > x); must be strictly between 0 and 1.", _
        "MeanLog: Mean of the natural logarithm of X, not the arithmetic mean.", _
        "StdDevLog: Positive standard deviation of the natural logarithm of X.", _
        "Status: (optional, default """"): Optional ByRef diagnostic string for VBA callers; omit in worksheet formulas. Empty on success."
    AddRecord "K_STATS_Lognormal_Mean", "MeanLog|StdDevLog|Status", _
        "Lognormal: Arithmetic mean of the distribution.", _
        "MeanLog: Mean of the natural logarithm of X, not the arithmetic mean.", _
        "StdDevLog: Positive standard deviation of the natural logarithm of X.", _
        "Status: (optional, default """"): Optional ByRef diagnostic string for VBA callers; omit in worksheet formulas. Empty on success."
    AddRecord "K_STATS_Lognormal_ParametersFromMeanStdDev", "Mean|StdDev|Status", _
        "Lognormal: Converts arithmetic mean and standard deviation to MeanLog and StdDevLog; returns a 1 x 2 array.", _
        "Mean: Positive arithmetic mean of X, not MeanLog.", _
        "StdDev: Non-negative arithmetic standard deviation. Zero returns StdDevLog=0, which density/CDF/inverse functions reject.", _
        "Status: (optional, default """"): Optional ByRef diagnostic string for VBA callers; omit in worksheet formulas. Empty on success."
    AddRecord "K_STATS_Lognormal_StdDev", "MeanLog|StdDevLog|Status", _
        "Lognormal: Standard deviation of the distribution.", _
        "MeanLog: Mean of the natural logarithm of X, not the arithmetic mean.", _
        "StdDevLog: Positive standard deviation of the natural logarithm of X.", _
        "Status: (optional, default """"): Optional ByRef diagnostic string for VBA callers; omit in worksheet formulas. Empty on success."
    AddRecord "K_STATS_Lognormal_Survival", "X|MeanLog|StdDevLog|Status", _
        "Lognormal: Upper-tail probability P(X > x), evaluated directly.", _
        "X: Evaluation point x.", _
        "MeanLog: Mean of the natural logarithm of X, not the arithmetic mean.", _
        "StdDevLog: Positive standard deviation of the natural logarithm of X.", _
        "Status: (optional, default """"): Optional ByRef diagnostic string for VBA callers; omit in worksheet formulas. Empty on success."
    AddRecord "K_STATS_Lognormal_Variance", "MeanLog|StdDevLog|Status", _
        "Lognormal: Variance of the distribution.", _
        "MeanLog: Mean of the natural logarithm of X, not the arithmetic mean.", _
        "StdDevLog: Positive standard deviation of the natural logarithm of X.", _
        "Status: (optional, default """"): Optional ByRef diagnostic string for VBA callers; omit in worksheet formulas. Empty on success."
    AddRecord "K_STATS_Normal_Cumulative", "X|Mean|StdDev|Status", _
        "Normal: Lower-tail cumulative probability P(X <= x).", _
        "X: Evaluation point x.", _
        "Mean: Arithmetic mean.", _
        "StdDev: Positive arithmetic standard deviation.", _
        "Status: (optional, default """"): Optional ByRef diagnostic string for VBA callers; omit in worksheet formulas. Empty on success."
    AddRecord "K_STATS_Normal_Density", "X|Mean|StdDev|Status", _
        "Normal: Probability density at x.", _
        "X: Evaluation point x.", _
        "Mean: Arithmetic mean.", _
        "StdDev: Positive arithmetic standard deviation.", _
        "Status: (optional, default """"): Optional ByRef diagnostic string for VBA callers; omit in worksheet formulas. Empty on success."
    AddRecord "K_STATS_Normal_IntervalProbability", "LowerBound|UpperBound|Mean|StdDev|Status", _
        "Normal: Probability between the lower and upper endpoints.", _
        "LowerBound: Lower bound.", _
        "UpperBound: Upper bound.", _
        "Mean: Arithmetic mean.", _
        "StdDev: Positive arithmetic standard deviation.", _
        "Status: (optional, default """"): Optional ByRef diagnostic string for VBA callers; omit in worksheet formulas. Empty on success."
    AddRecord "K_STATS_Normal_InverseCumulative", "Probability|Mean|StdDev|Status", _
        "Normal: Lower-tail quantile for the supplied probability.", _
        "Probability: Lower-tail probability. Endpoint handling follows the function contract.", _
        "Mean: Arithmetic mean.", _
        "StdDev: Positive arithmetic standard deviation.", _
        "Status: (optional, default """"): Optional ByRef diagnostic string for VBA callers; omit in worksheet formulas. Empty on success."
    AddRecord "K_STATS_Normal_InverseSurvival", "Probability|Mean|StdDev|Status", _
        "Normal: Upper-tail quantile for the supplied survival probability.", _
        "Probability: Upper-tail probability P(X > x); must be strictly between 0 and 1.", _
        "Mean: Arithmetic mean.", _
        "StdDev: Positive arithmetic standard deviation.", _
        "Status: (optional, default """"): Optional ByRef diagnostic string for VBA callers; omit in worksheet formulas. Empty on success."
    AddRecord "K_STATS_Normal_Survival", "X|Mean|StdDev|Status", _
        "Normal: Upper-tail probability P(X > x), evaluated directly.", _
        "X: Evaluation point x.", _
        "Mean: Arithmetic mean.", _
        "StdDev: Positive arithmetic standard deviation.", _
        "Status: (optional, default """"): Optional ByRef diagnostic string for VBA callers; omit in worksheet formulas. Empty on success."
    AddRecord "K_STATS_Normal_ZScore", "X|Mean|StdDev|Status", _
        "Normal: Standardized value (x - mean) / standard deviation.", _
        "X: Evaluation point x.", _
        "Mean: Arithmetic mean.", _
        "StdDev: Positive arithmetic standard deviation.", _
        "Status: (optional, default """"): Optional ByRef diagnostic string for VBA callers; omit in worksheet formulas. Empty on success."
    AddRecord "K_STATS_NormalStandard_Cumulative", "Z|Status", _
        "Standard normal: Lower-tail cumulative probability P(X <= x).", _
        "Z: Standard normal evaluation point z.", _
        "Status: (optional, default """"): Optional ByRef diagnostic string for VBA callers; omit in worksheet formulas. Empty on success."
    AddRecord "K_STATS_NormalStandard_Density", "Z|Status", _
        "Standard normal: Probability density at x.", _
        "Z: Standard normal evaluation point z.", _
        "Status: (optional, default """"): Optional ByRef diagnostic string for VBA callers; omit in worksheet formulas. Empty on success."
    AddRecord "K_STATS_NormalStandard_IntervalProbability", "LowerZ|UpperZ|Status", _
        "Standard normal: Probability between the lower and upper endpoints.", _
        "LowerZ: Lower endpoint in standard normal units.", _
        "UpperZ: Upper endpoint in standard normal units.", _
        "Status: (optional, default """"): Optional ByRef diagnostic string for VBA callers; omit in worksheet formulas. Empty on success."
    AddRecord "K_STATS_NormalStandard_InverseCumulative", "Probability|Status", _
        "Standard normal: Lower-tail quantile for the supplied probability.", _
        "Probability: Lower-tail probability. Endpoint handling follows the function contract.", _
        "Status: (optional, default """"): Optional ByRef diagnostic string for VBA callers; omit in worksheet formulas. Empty on success."
    AddRecord "K_STATS_NormalStandard_InverseCumulativeFast", "Probability", _
        "Standard normal: Fast approximate standard normal lower-tail quantile.", _
        "Probability: Lower-tail probability. Endpoint handling follows the function contract."
    AddRecord "K_STATS_NormalStandard_InverseSurvival", "Probability|Status", _
        "Standard normal: Upper-tail quantile for the supplied survival probability.", _
        "Probability: Upper-tail probability P(X > x); must be strictly between 0 and 1.", _
        "Status: (optional, default """"): Optional ByRef diagnostic string for VBA callers; omit in worksheet formulas. Empty on success."
    AddRecord "K_STATS_NormalStandard_Survival", "Z|Status", _
        "Standard normal: Upper-tail probability P(X > x), evaluated directly.", _
        "Z: Standard normal evaluation point z.", _
        "Status: (optional, default """"): Optional ByRef diagnostic string for VBA callers; omit in worksheet formulas. Empty on success."
    AddRecord "K_STATS_ChiSquare_Cumulative", "X|DegreesFreedom|Status", _
        "Chi-square: Lower-tail cumulative probability P(X <= x).", _
        "X: Evaluation point x.", _
        "DegreesFreedom: Positive degrees of freedom.", _
        "Status: (optional, default """"): Optional ByRef diagnostic string for VBA callers; omit in worksheet formulas. Empty on success."
    AddRecord "K_STATS_ChiSquare_Density", "X|DegreesFreedom|Status", _
        "Chi-square: Probability density at x.", _
        "X: Evaluation point x.", _
        "DegreesFreedom: Positive degrees of freedom.", _
        "Status: (optional, default """"): Optional ByRef diagnostic string for VBA callers; omit in worksheet formulas. Empty on success."
    AddRecord "K_STATS_ChiSquare_InverseCumulative", "Probability|DegreesFreedom|Status", _
        "Chi-square: Lower-tail quantile for the supplied probability.", _
        "Probability: Lower-tail probability. Endpoint handling follows the function contract.", _
        "DegreesFreedom: Positive degrees of freedom.", _
        "Status: (optional, default """"): Optional ByRef diagnostic string for VBA callers; omit in worksheet formulas. Empty on success."
    AddRecord "K_STATS_ChiSquare_Survival", "X|DegreesFreedom|Status", _
        "Chi-square: Upper-tail probability P(X > x), evaluated directly.", _
        "X: Evaluation point x.", _
        "DegreesFreedom: Positive degrees of freedom.", _
        "Status: (optional, default """"): Optional ByRef diagnostic string for VBA callers; omit in worksheet formulas. Empty on success."
    AddRecord "K_STATS_F_Cumulative", "X|DegreesFreedom1|DegreesFreedom2|Status", _
        "F: Lower-tail cumulative probability P(X <= x).", _
        "X: Evaluation point x.", _
        "DegreesFreedom1: Positive numerator degrees of freedom.", _
        "DegreesFreedom2: Positive denominator degrees of freedom.", _
        "Status: (optional, default """"): Optional ByRef diagnostic string for VBA callers; omit in worksheet formulas. Empty on success."
    AddRecord "K_STATS_F_Density", "X|DegreesFreedom1|DegreesFreedom2|Status", _
        "F: Probability density at x.", _
        "X: Evaluation point x.", _
        "DegreesFreedom1: Positive numerator degrees of freedom.", _
        "DegreesFreedom2: Positive denominator degrees of freedom.", _
        "Status: (optional, default """"): Optional ByRef diagnostic string for VBA callers; omit in worksheet formulas. Empty on success."
    AddRecord "K_STATS_F_InverseCumulative", "Probability|DegreesFreedom1|DegreesFreedom2|Status", _
        "F: Lower-tail quantile for the supplied probability.", _
        "Probability: Lower-tail probability. Endpoint handling follows the function contract.", _
        "DegreesFreedom1: Positive numerator degrees of freedom.", _
        "DegreesFreedom2: Positive denominator degrees of freedom.", _
        "Status: (optional, default """"): Optional ByRef diagnostic string for VBA callers; omit in worksheet formulas. Empty on success."
    AddRecord "K_STATS_F_Survival", "X|DegreesFreedom1|DegreesFreedom2|Status", _
        "F: Upper-tail probability P(X > x), evaluated directly.", _
        "X: Evaluation point x.", _
        "DegreesFreedom1: Positive numerator degrees of freedom.", _
        "DegreesFreedom2: Positive denominator degrees of freedom.", _
        "Status: (optional, default """"): Optional ByRef diagnostic string for VBA callers; omit in worksheet formulas. Empty on success."
    AddRecord "K_STATS_StudentT_Cumulative", "X|DegreesFreedom|Status", _
        "Student t: Lower-tail cumulative probability P(X <= x).", _
        "X: Evaluation point x.", _
        "DegreesFreedom: Positive degrees of freedom.", _
        "Status: (optional, default """"): Optional ByRef diagnostic string for VBA callers; omit in worksheet formulas. Empty on success."
    AddRecord "K_STATS_StudentT_Density", "X|DegreesFreedom|Status", _
        "Student t: Probability density at x.", _
        "X: Evaluation point x.", _
        "DegreesFreedom: Positive degrees of freedom.", _
        "Status: (optional, default """"): Optional ByRef diagnostic string for VBA callers; omit in worksheet formulas. Empty on success."
    AddRecord "K_STATS_StudentT_InverseCumulative", "Probability|DegreesFreedom|Status", _
        "Student t: Lower-tail quantile for the supplied probability.", _
        "Probability: Lower-tail probability. Endpoint handling follows the function contract.", _
        "DegreesFreedom: Positive degrees of freedom.", _
        "Status: (optional, default """"): Optional ByRef diagnostic string for VBA callers; omit in worksheet formulas. Empty on success."
    AddRecord "K_STATS_StudentT_Survival", "X|DegreesFreedom|Status", _
        "Student t: Upper-tail probability P(X > x), evaluated directly.", _
        "X: Evaluation point x.", _
        "DegreesFreedom: Positive degrees of freedom.", _
        "Status: (optional, default """"): Optional ByRef diagnostic string for VBA callers; omit in worksheet formulas. Empty on success."
End Sub
