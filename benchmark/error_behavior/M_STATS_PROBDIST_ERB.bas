Attribute VB_Name = "M_STATS_PROBDIST_ERB"
Option Explicit

'One Double and its two 32-bit words, for exact transport through text
Private Type ERB_Words
    Low                     As Long
    High                    As Long
End Type

Private Type ERB_Number
    Value                   As Double
End Type

'An empty cell of the first worksheet, for the class empty
Private Const ERB_EMPTY_CELL            As String = "XFD1048576"

'==============================================================================
' M_STATS_PROBDIST_ERB
'------------------------------------------------------------------------------
' PURPOSE
'   Export macro for the survey of what every public function returns on
'   bad input (#331, benchmark/error_behavior/README.md). It makes each call
'   of the grid and records what came back. It asserts nothing.
'
' USAGE
'   Export_ErrorBehavior fills error_behavior_grid.csv. It finds the file
'   beside the workbook, or asks for it; ci/Export-ExcelGrids.ps1 runs it on
'   the self-hosted runner with the file staged beside the export workbook.
'
' GRID FORMAT (study grid; 6 columns)
'   point_id, function, class, argument, arguments, observed
'
'   The macro writes the last. arguments holds the call's arguments
'   separated by ";": a number, or NAN, PINF, NINF, TEXT, EMPTY.
'
' HOW A CALL IS MADE
'   A call with TEXT or EMPTY among its arguments is built as a worksheet
'   formula and evaluated as a cell would evaluate it, because only a cell
'   passes a text or an empty cell. Every other call is made directly, with
'   Double arguments, through DirectCall, which is how VBA code can pass a
'   value no cell can hold. A direct call lets an error the function raises
'   reach this module's handler; through Application.Run it would stop Excel
'   on a dialog.
'
' WHAT IS RECORDED
'   The 16 hexadecimal digits of a numeric result; the worksheet error
'   returned (#NUM!, #VALUE!, #DIV/0!, #N/A, or ERR and its number); RAISED
'   and the error number when the call raised a VBA error instead of
'   returning; TYPE and the VarType of any other result.
'
' UPDATED
'   2026-11-16
'==============================================================================


Public Sub Export_ErrorBehavior()
'
'==============================================================================
' Export_ErrorBehavior
'------------------------------------------------------------------------------
' PURPOSE
'   Fills the observation column of error_behavior_grid.csv.
'==============================================================================
'
    Dim Path                As String          'File written
    Dim Filled              As Long            'Rows written, or -1

    Filled = ExportFile("error_behavior_grid.csv", "Error behavior", Path)
    If Filled < 0 Then Exit Sub

    MsgBox "Error behavior complete: " & Filled & _
           " observation(s) written to" & vbCrLf & Path, _
           vbInformation, "Error behavior"
End Sub


Private Function ExportFile( _
    ByVal FileName As String, _
    ByVal Title As String, _
    ByRef Path As String) _
    As Long
'
'==============================================================================
' ExportFile
'------------------------------------------------------------------------------
' PURPOSE
'   Fills the observation column of the grid file, row by row. Returns the
'   number of rows written, or -1 when the file was not chosen or the export
'   failed; the failure is reported here, the success by the caller.
'
' BEHAVIOR
'   - Reads the whole file, makes the call of every data row, rewrites the
'     file.
'   - A row with fewer than 6 fields is left as it is and not counted.
'==============================================================================
'
'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Dim Lines()             As String          'File lines
    Dim Raw                 As String          'File contents
    Dim FileNum             As Integer         'Input file handle
    Dim OutNum              As Integer         'Output file handle
    Dim Cols                As Variant         'Split fields of one row
    Dim Filled              As Long            'Rows written
    Dim I                   As Long            'Row index
'------------------------------------------------------------------------------
' READ
'------------------------------------------------------------------------------
    On Error GoTo Err_Handler

        Path = ResolveGridPath(FileName, Title)
        ExportFile = -1
        If Len(Path) = 0 Then Exit Function     'User cancelled the picker

        Filled = 0

        FileNum = FreeFile
        Open Path For Input As #FileNum
        Raw = Input$(LOF(FileNum), FileNum)
        Close #FileNum

        Raw = Replace(Raw, vbCrLf, vbLf)
        Raw = Replace(Raw, vbCr, vbLf)
        Lines = Split(Raw, vbLf)
'------------------------------------------------------------------------------
' CALL
'------------------------------------------------------------------------------
    'Row 0 is the header; data starts at row 1
        For I = 1 To UBound(Lines)
            If Len(Trim$(Lines(I))) = 0 Then GoTo ContinueRow

            Cols = Split(Lines(I), ",")
            If UBound(Cols) < 5 Then GoTo ContinueRow

            Cols(5) = Observed(CStr(Cols(1)), CStr(Cols(4)))

            Lines(I) = Join(Cols, ",")
            Filled = Filled + 1
ContinueRow:
        Next I
'------------------------------------------------------------------------------
' WRITE
'------------------------------------------------------------------------------
        OutNum = FreeFile
        Open Path For Output As #OutNum
        For I = 0 To UBound(Lines)
            If I < UBound(Lines) Or Len(Lines(I)) > 0 Then Print #OutNum, Lines(I)
        Next I
        Close #OutNum

    ExportFile = Filled
    Exit Function

Err_Handler:
    On Error Resume Next
    Close #FileNum
    Close #OutNum
    MsgBox Title & " failed: " & Err.Description, _
           vbExclamation, Title
    ExportFile = -1
End Function


Private Function Observed(ByVal FunctionName As String, ByVal Stored As String) As String
'
'==============================================================================
' Observed
'------------------------------------------------------------------------------
' PURPOSE
'   Makes one call and returns what came back, as text.
'==============================================================================
'
    Dim Tokens()            As String          'The stored arguments
    Dim A(0 To 3)           As Double          'The arguments as Doubles
    Dim Formula             As String          'The call as a worksheet formula
    Dim AsCell              As Boolean         'TRUE when a text or an empty cell is passed
    Dim V                   As Variant         'Result of the call
    Dim I                   As Long            'Argument index

    On Error GoTo Raised

    Tokens = Split(Stored, ";")

    For I = 0 To UBound(Tokens)
        If Tokens(I) = "TEXT" Or Tokens(I) = "EMPTY" Then AsCell = True
    Next I

    If AsCell Then
        Formula = FunctionName & "("
        For I = 0 To UBound(Tokens)
            If I > 0 Then Formula = Formula & ","
            Select Case Tokens(I)
                Case "TEXT": Formula = Formula & """abc"""
                Case "EMPTY": Formula = Formula & ERB_EMPTY_CELL
                Case Else: Formula = Formula & Tokens(I)
            End Select
        Next I
        V = ThisWorkbook.Worksheets(1).Evaluate(Formula & ")")
    Else
        If UBound(Tokens) > 3 Then
            Observed = "UNSUPPORTED"
            Exit Function
        End If
        For I = 0 To UBound(Tokens)
            A(I) = Argument(Tokens(I))
        Next I
        V = DirectCall(FunctionName, A)
    End If

    Observed = Described(V)
    Exit Function

Raised:
    Observed = "RAISED" & CStr(Err.Number)
End Function


Private Function DirectCall(ByVal FunctionName As String, ByRef A() As Double) As Variant
'
'==============================================================================
' DirectCall
'------------------------------------------------------------------------------
' PURPOSE
'   Calls one public function by name with the first of the four Doubles of
'   A that it takes. An unknown name returns the text UNKNOWN.
'
' GENERATED
'   benchmark/error_behavior/build_grid.py --dispatch, from
'   docs/PUBLIC_API.txt. test_survey.py fails when the two differ.
'==============================================================================
'
    Select Case FunctionName
        Case "K_STATS_Beta_Cumulative"
            DirectCall = K_STATS_Beta_Cumulative(A(0), A(1), A(2))
        Case "K_STATS_Beta_Density"
            DirectCall = K_STATS_Beta_Density(A(0), A(1), A(2))
        Case "K_STATS_Beta_InverseCumulative"
            DirectCall = K_STATS_Beta_InverseCumulative(A(0), A(1), A(2))
        Case "K_STATS_Beta_Mean"
            DirectCall = K_STATS_Beta_Mean(A(0), A(1))
        Case "K_STATS_Beta_StdDev"
            DirectCall = K_STATS_Beta_StdDev(A(0), A(1))
        Case "K_STATS_Beta_Survival"
            DirectCall = K_STATS_Beta_Survival(A(0), A(1), A(2))
        Case "K_STATS_Beta_Variance"
            DirectCall = K_STATS_Beta_Variance(A(0), A(1))
        Case "K_STATS_Exponential_Cumulative"
            DirectCall = K_STATS_Exponential_Cumulative(A(0), A(1))
        Case "K_STATS_Exponential_Density"
            DirectCall = K_STATS_Exponential_Density(A(0), A(1))
        Case "K_STATS_Exponential_InverseCumulative"
            DirectCall = K_STATS_Exponential_InverseCumulative(A(0), A(1))
        Case "K_STATS_Exponential_Survival"
            DirectCall = K_STATS_Exponential_Survival(A(0), A(1))
        Case "K_STATS_Gamma_Cumulative"
            DirectCall = K_STATS_Gamma_Cumulative(A(0), A(1), A(2))
        Case "K_STATS_Gamma_Density"
            DirectCall = K_STATS_Gamma_Density(A(0), A(1), A(2))
        Case "K_STATS_Gamma_InverseCumulative"
            DirectCall = K_STATS_Gamma_InverseCumulative(A(0), A(1), A(2))
        Case "K_STATS_Gamma_Mean"
            DirectCall = K_STATS_Gamma_Mean(A(0), A(1))
        Case "K_STATS_Gamma_StdDev"
            DirectCall = K_STATS_Gamma_StdDev(A(0), A(1))
        Case "K_STATS_Gamma_Survival"
            DirectCall = K_STATS_Gamma_Survival(A(0), A(1), A(2))
        Case "K_STATS_Gamma_Variance"
            DirectCall = K_STATS_Gamma_Variance(A(0), A(1))
        Case "K_STATS_Uniform_Cumulative"
            DirectCall = K_STATS_Uniform_Cumulative(A(0), A(1), A(2))
        Case "K_STATS_Uniform_Density"
            DirectCall = K_STATS_Uniform_Density(A(0), A(1), A(2))
        Case "K_STATS_Uniform_InverseCumulative"
            DirectCall = K_STATS_Uniform_InverseCumulative(A(0), A(1), A(2))
        Case "K_STATS_Uniform_Survival"
            DirectCall = K_STATS_Uniform_Survival(A(0), A(1), A(2))
        Case "K_STATS_Weibull_Cumulative"
            DirectCall = K_STATS_Weibull_Cumulative(A(0), A(1), A(2))
        Case "K_STATS_Weibull_Density"
            DirectCall = K_STATS_Weibull_Density(A(0), A(1), A(2))
        Case "K_STATS_Weibull_InverseCumulative"
            DirectCall = K_STATS_Weibull_InverseCumulative(A(0), A(1), A(2))
        Case "K_STATS_Weibull_Mean"
            DirectCall = K_STATS_Weibull_Mean(A(0), A(1))
        Case "K_STATS_Weibull_StdDev"
            DirectCall = K_STATS_Weibull_StdDev(A(0), A(1))
        Case "K_STATS_Weibull_Survival"
            DirectCall = K_STATS_Weibull_Survival(A(0), A(1), A(2))
        Case "K_STATS_Weibull_Variance"
            DirectCall = K_STATS_Weibull_Variance(A(0), A(1))
        Case "K_STATS_Binomial_Cumulative"
            DirectCall = K_STATS_Binomial_Cumulative(A(0), A(1), A(2))
        Case "K_STATS_Binomial_InverseCumulative"
            DirectCall = K_STATS_Binomial_InverseCumulative(A(0), A(1), A(2))
        Case "K_STATS_Binomial_LogPMF"
            DirectCall = K_STATS_Binomial_LogPMF(A(0), A(1), A(2))
        Case "K_STATS_Binomial_Mean"
            DirectCall = K_STATS_Binomial_Mean(A(0), A(1))
        Case "K_STATS_Binomial_PMF"
            DirectCall = K_STATS_Binomial_PMF(A(0), A(1), A(2))
        Case "K_STATS_Binomial_StdDev"
            DirectCall = K_STATS_Binomial_StdDev(A(0), A(1))
        Case "K_STATS_Binomial_Survival"
            DirectCall = K_STATS_Binomial_Survival(A(0), A(1), A(2))
        Case "K_STATS_Binomial_Variance"
            DirectCall = K_STATS_Binomial_Variance(A(0), A(1))
        Case "K_STATS_DiscreteUniform_Cumulative"
            DirectCall = K_STATS_DiscreteUniform_Cumulative(A(0), A(1), A(2))
        Case "K_STATS_DiscreteUniform_InverseCumulative"
            DirectCall = K_STATS_DiscreteUniform_InverseCumulative(A(0), A(1), A(2))
        Case "K_STATS_DiscreteUniform_LogPMF"
            DirectCall = K_STATS_DiscreteUniform_LogPMF(A(0), A(1), A(2))
        Case "K_STATS_DiscreteUniform_Mean"
            DirectCall = K_STATS_DiscreteUniform_Mean(A(0), A(1))
        Case "K_STATS_DiscreteUniform_PMF"
            DirectCall = K_STATS_DiscreteUniform_PMF(A(0), A(1), A(2))
        Case "K_STATS_DiscreteUniform_StdDev"
            DirectCall = K_STATS_DiscreteUniform_StdDev(A(0), A(1))
        Case "K_STATS_DiscreteUniform_Survival"
            DirectCall = K_STATS_DiscreteUniform_Survival(A(0), A(1), A(2))
        Case "K_STATS_DiscreteUniform_Variance"
            DirectCall = K_STATS_DiscreteUniform_Variance(A(0), A(1))
        Case "K_STATS_Geometric_Cumulative"
            DirectCall = K_STATS_Geometric_Cumulative(A(0), A(1))
        Case "K_STATS_Geometric_InverseCumulative"
            DirectCall = K_STATS_Geometric_InverseCumulative(A(0), A(1))
        Case "K_STATS_Geometric_LogPMF"
            DirectCall = K_STATS_Geometric_LogPMF(A(0), A(1))
        Case "K_STATS_Geometric_Mean"
            DirectCall = K_STATS_Geometric_Mean(A(0))
        Case "K_STATS_Geometric_PMF"
            DirectCall = K_STATS_Geometric_PMF(A(0), A(1))
        Case "K_STATS_Geometric_StdDev"
            DirectCall = K_STATS_Geometric_StdDev(A(0))
        Case "K_STATS_Geometric_Survival"
            DirectCall = K_STATS_Geometric_Survival(A(0), A(1))
        Case "K_STATS_Geometric_Variance"
            DirectCall = K_STATS_Geometric_Variance(A(0))
        Case "K_STATS_Hypergeometric_Cumulative"
            DirectCall = K_STATS_Hypergeometric_Cumulative(A(0), A(1), A(2), A(3))
        Case "K_STATS_Hypergeometric_InverseCumulative"
            DirectCall = K_STATS_Hypergeometric_InverseCumulative(A(0), A(1), A(2), A(3))
        Case "K_STATS_Hypergeometric_LogPMF"
            DirectCall = K_STATS_Hypergeometric_LogPMF(A(0), A(1), A(2), A(3))
        Case "K_STATS_Hypergeometric_Mean"
            DirectCall = K_STATS_Hypergeometric_Mean(A(0), A(1), A(2))
        Case "K_STATS_Hypergeometric_PMF"
            DirectCall = K_STATS_Hypergeometric_PMF(A(0), A(1), A(2), A(3))
        Case "K_STATS_Hypergeometric_StdDev"
            DirectCall = K_STATS_Hypergeometric_StdDev(A(0), A(1), A(2))
        Case "K_STATS_Hypergeometric_Survival"
            DirectCall = K_STATS_Hypergeometric_Survival(A(0), A(1), A(2), A(3))
        Case "K_STATS_Hypergeometric_Variance"
            DirectCall = K_STATS_Hypergeometric_Variance(A(0), A(1), A(2))
        Case "K_STATS_NegativeBinomial_Cumulative"
            DirectCall = K_STATS_NegativeBinomial_Cumulative(A(0), A(1), A(2))
        Case "K_STATS_NegativeBinomial_InverseCumulative"
            DirectCall = K_STATS_NegativeBinomial_InverseCumulative(A(0), A(1), A(2))
        Case "K_STATS_NegativeBinomial_LogPMF"
            DirectCall = K_STATS_NegativeBinomial_LogPMF(A(0), A(1), A(2))
        Case "K_STATS_NegativeBinomial_Mean"
            DirectCall = K_STATS_NegativeBinomial_Mean(A(0), A(1))
        Case "K_STATS_NegativeBinomial_PMF"
            DirectCall = K_STATS_NegativeBinomial_PMF(A(0), A(1), A(2))
        Case "K_STATS_NegativeBinomial_StdDev"
            DirectCall = K_STATS_NegativeBinomial_StdDev(A(0), A(1))
        Case "K_STATS_NegativeBinomial_Survival"
            DirectCall = K_STATS_NegativeBinomial_Survival(A(0), A(1), A(2))
        Case "K_STATS_NegativeBinomial_Variance"
            DirectCall = K_STATS_NegativeBinomial_Variance(A(0), A(1))
        Case "K_STATS_Poisson_Cumulative"
            DirectCall = K_STATS_Poisson_Cumulative(A(0), A(1))
        Case "K_STATS_Poisson_InverseCumulative"
            DirectCall = K_STATS_Poisson_InverseCumulative(A(0), A(1))
        Case "K_STATS_Poisson_LogPMF"
            DirectCall = K_STATS_Poisson_LogPMF(A(0), A(1))
        Case "K_STATS_Poisson_Mean"
            DirectCall = K_STATS_Poisson_Mean(A(0))
        Case "K_STATS_Poisson_PMF"
            DirectCall = K_STATS_Poisson_PMF(A(0), A(1))
        Case "K_STATS_Poisson_StdDev"
            DirectCall = K_STATS_Poisson_StdDev(A(0))
        Case "K_STATS_Poisson_Survival"
            DirectCall = K_STATS_Poisson_Survival(A(0), A(1))
        Case "K_STATS_Poisson_Variance"
            DirectCall = K_STATS_Poisson_Variance(A(0))
        Case "K_STATS_Lognormal_Cumulative"
            DirectCall = K_STATS_Lognormal_Cumulative(A(0), A(1), A(2))
        Case "K_STATS_Lognormal_Density"
            DirectCall = K_STATS_Lognormal_Density(A(0), A(1), A(2))
        Case "K_STATS_Lognormal_InverseCumulative"
            DirectCall = K_STATS_Lognormal_InverseCumulative(A(0), A(1), A(2))
        Case "K_STATS_Lognormal_InverseSurvival"
            DirectCall = K_STATS_Lognormal_InverseSurvival(A(0), A(1), A(2))
        Case "K_STATS_Lognormal_Mean"
            DirectCall = K_STATS_Lognormal_Mean(A(0), A(1))
        Case "K_STATS_Lognormal_ParametersFromMeanStdDev"
            DirectCall = K_STATS_Lognormal_ParametersFromMeanStdDev(A(0), A(1))
        Case "K_STATS_Lognormal_StdDev"
            DirectCall = K_STATS_Lognormal_StdDev(A(0), A(1))
        Case "K_STATS_Lognormal_Survival"
            DirectCall = K_STATS_Lognormal_Survival(A(0), A(1), A(2))
        Case "K_STATS_Lognormal_Variance"
            DirectCall = K_STATS_Lognormal_Variance(A(0), A(1))
        Case "K_STATS_Normal_Cumulative"
            DirectCall = K_STATS_Normal_Cumulative(A(0), A(1), A(2))
        Case "K_STATS_Normal_Density"
            DirectCall = K_STATS_Normal_Density(A(0), A(1), A(2))
        Case "K_STATS_Normal_IntervalProbability"
            DirectCall = K_STATS_Normal_IntervalProbability(A(0), A(1), A(2), A(3))
        Case "K_STATS_Normal_InverseCumulative"
            DirectCall = K_STATS_Normal_InverseCumulative(A(0), A(1), A(2))
        Case "K_STATS_Normal_InverseSurvival"
            DirectCall = K_STATS_Normal_InverseSurvival(A(0), A(1), A(2))
        Case "K_STATS_Normal_Survival"
            DirectCall = K_STATS_Normal_Survival(A(0), A(1), A(2))
        Case "K_STATS_Normal_ZScore"
            DirectCall = K_STATS_Normal_ZScore(A(0), A(1), A(2))
        Case "K_STATS_NormalStandard_Cumulative"
            DirectCall = K_STATS_NormalStandard_Cumulative(A(0))
        Case "K_STATS_NormalStandard_Density"
            DirectCall = K_STATS_NormalStandard_Density(A(0))
        Case "K_STATS_NormalStandard_IntervalProbability"
            DirectCall = K_STATS_NormalStandard_IntervalProbability(A(0), A(1))
        Case "K_STATS_NormalStandard_InverseCumulative"
            DirectCall = K_STATS_NormalStandard_InverseCumulative(A(0))
        Case "K_STATS_NormalStandard_InverseCumulativeFast"
            DirectCall = K_STATS_NormalStandard_InverseCumulativeFast(A(0))
        Case "K_STATS_NormalStandard_InverseSurvival"
            DirectCall = K_STATS_NormalStandard_InverseSurvival(A(0))
        Case "K_STATS_NormalStandard_Survival"
            DirectCall = K_STATS_NormalStandard_Survival(A(0))
        Case "K_STATS_ChiSquare_Cumulative"
            DirectCall = K_STATS_ChiSquare_Cumulative(A(0), A(1))
        Case "K_STATS_ChiSquare_Density"
            DirectCall = K_STATS_ChiSquare_Density(A(0), A(1))
        Case "K_STATS_ChiSquare_InverseCumulative"
            DirectCall = K_STATS_ChiSquare_InverseCumulative(A(0), A(1))
        Case "K_STATS_ChiSquare_Survival"
            DirectCall = K_STATS_ChiSquare_Survival(A(0), A(1))
        Case "K_STATS_F_Cumulative"
            DirectCall = K_STATS_F_Cumulative(A(0), A(1), A(2))
        Case "K_STATS_F_Density"
            DirectCall = K_STATS_F_Density(A(0), A(1), A(2))
        Case "K_STATS_F_InverseCumulative"
            DirectCall = K_STATS_F_InverseCumulative(A(0), A(1), A(2))
        Case "K_STATS_F_Survival"
            DirectCall = K_STATS_F_Survival(A(0), A(1), A(2))
        Case "K_STATS_StudentT_Cumulative"
            DirectCall = K_STATS_StudentT_Cumulative(A(0), A(1))
        Case "K_STATS_StudentT_Density"
            DirectCall = K_STATS_StudentT_Density(A(0), A(1))
        Case "K_STATS_StudentT_InverseCumulative"
            DirectCall = K_STATS_StudentT_InverseCumulative(A(0), A(1))
        Case "K_STATS_StudentT_Survival"
            DirectCall = K_STATS_StudentT_Survival(A(0), A(1))
        Case Else
            DirectCall = "UNKNOWN"
    End Select
End Function


Private Function Argument(ByVal Token As String) As Double
'
'==============================================================================
' Argument
'------------------------------------------------------------------------------
' PURPOSE
'   The Double a stored argument names: a number, or one of the three values
'   no decimal text denotes, built from their bit patterns.
'==============================================================================
'
    Select Case Token
        Case "NAN": Argument = FromWords(&H7FF80000, 0)
        Case "PINF": Argument = FromWords(&H7FF00000, 0)
        Case "NINF": Argument = FromWords(&HFFF00000, 0)
        Case Else: Argument = Val(Token)
    End Select
End Function


Private Function FromWords(ByVal HighWord As Long, ByVal LowWord As Long) As Double
'
'==============================================================================
' FromWords
'------------------------------------------------------------------------------
' PURPOSE
'   The Double whose high and low 32 bits are the two words.
'==============================================================================
'
    Dim Words               As ERB_Words       'The two halves
    Dim Number              As ERB_Number      'The same eight bytes as a Double

    Words.High = HighWord
    Words.Low = LowWord
    LSet Number = Words
    FromWords = Number.Value
End Function


Private Function Described(ByVal V As Variant) As String
'
'==============================================================================
' Described
'------------------------------------------------------------------------------
' PURPOSE
'   A result as text: the bits of a number, the name of a worksheet error,
'   or TYPE and the VarType of anything else.
'==============================================================================
'
    If IsError(V) Then
        Select Case CLng(V)
            Case 2036: Described = "#NUM!"
            Case 2015: Described = "#VALUE!"
            Case 2007: Described = "#DIV/0!"
            Case 2042: Described = "#N/A"
            Case Else: Described = "ERR" & CStr(CLng(V))
        End Select
    ElseIf VarType(V) = vbDouble Or VarType(V) = vbInteger Or VarType(V) = vbLong Then
        Described = ToBits(CDbl(V))
    Else
        Described = "TYPE" & CStr(VarType(V))
    End If
End Function


Private Function ToBits(ByVal X As Double) As String
'
'==============================================================================
' ToBits
'------------------------------------------------------------------------------
' PURPOSE
'   The 64 bits of X as 16 hexadecimal digits, most significant first.
'==============================================================================
'
    Dim Words               As ERB_Words       'The two halves
    Dim Number              As ERB_Number      'The same eight bytes as a Double

    Number.Value = X
    LSet Words = Number
    ToBits = Right$("00000000" & Hex$(Words.High), 8) & _
             Right$("00000000" & Hex$(Words.Low), 8)
End Function


Private Function ResolveGridPath( _
    ByVal FileName As String, _
    ByVal Title As String) _
    As String
'
'==============================================================================
' ResolveGridPath
'------------------------------------------------------------------------------
' PURPOSE
'   Finds the grid file beside the workbook, or asks for it.
'==============================================================================
'
    Dim BookPath            As String          'Workbook folder
    Dim Candidate           As String          'Path next to the workbook
    Dim Picked              As Variant         'File-dialog result

    BookPath = ThisWorkbook.Path

    If Len(BookPath) > 0 And LCase$(Left$(BookPath, 4)) <> "http" Then
        Candidate = BookPath & Application.PathSeparator & FileName
        If Len(Dir$(Candidate)) > 0 Then
            ResolveGridPath = Candidate
            Exit Function
        End If
    End If

    MsgBox "Could not locate " & FileName & " automatically " & _
           "(the workbook may be on OneDrive/SharePoint). Please select it.", _
           vbInformation, "Locate " & Title

    Picked = Application.GetOpenFilename( _
        FileFilter:=Title & " (*.csv),*.csv", _
        Title:="Select " & FileName)

    If VarType(Picked) = vbBoolean Then
        ResolveGridPath = vbNullString
    Else
        ResolveGridPath = CStr(Picked)
    End If
End Function
