Attribute VB_Name = "M_STATS_PROBDIST_GIS2"
Option Explicit

'One Double and its two 32-bit words, for exact transport through text
Private Type GIS2_Words
    Low                     As Long
    High                    As Long
End Type

Private Type GIS2_Number
    Value                   As Double
End Type

'==============================================================================
' M_STATS_PROBDIST_GIS2
'------------------------------------------------------------------------------
' PURPOSE
'   Export macro for the second #14 holdout of the Gamma inverse at the final
'   scale (PREREGISTRATION_2.md).
'   Fills the observed side of gamma_inverse_scale_grid2.csv
'   (benchmark/gamma_inverse_scale_study/PREREGISTRATION_2.md): at each
'   (Probability, Shape, Scale) it calls the public inverses and the public
'   cumulative at the returned quantile, and records each result as the 16
'   hexadecimal digits of its Double.
'
' USAGE
'   Run Export_GammaInverseScaleHoldout2. It finds the grid beside the
'   workbook, or asks for it. ci/Export-ExcelGrids.ps1 runs it on the
'   self-hosted runner with the grid staged beside the export workbook (#225).
'
' GRID FORMAT (study grid; 10 columns)
'   point_id, kind, p_bits, a_bits, s_bits, ref_x, gamma_inv, chi_inv,
'   gamma_cdf, log_unit
'
'   This macro writes the last four columns. The inputs are rebuilt from
'   their bit patterns, so no decimal text is parsed. It never reads the
'   reference, so the observed side stays independent.
'
'       gamma_inv   K_STATS_Gamma_InverseCumulative(P, A, S)
'       chi_inv     K_STATS_ChiSquare_InverseCumulative(P, 2A) where S = 2;
'                   NA elsewhere
'       gamma_cdf   K_STATS_Gamma_Cumulative(gamma_inv, A, S) where gamma_inv
'                   is a positive number; NA elsewhere
'       log_unit    PROB_TryGammaInvLogUnit(P, A): the value the kernel
'                   itself classifies the call by
'
' ERROR POLICY
'   A public function that returns a worksheet error is recorded as CVERR, a
'   kernel that returns False as FAIL, and a runtime fault as ERROR. None is
'   ever scored as a value.
'
' DEPENDENCIES
'   - K_STATS_Gamma_InverseCumulative, _Cumulative  (M_STATS_PROBDIST_CONTINUOUS)
'   - K_STATS_ChiSquare_InverseCumulative           (M_STATS_PROBDIST_TFAMILY)
'   - PROB_TryGammaInvLogUnit                       (M_STATS_PROBDIST_SPECIALFUNCS)
'
' UPDATED
'   2026-11-14
'==============================================================================


Public Sub Export_GammaInverseScaleHoldout2()
'
'==============================================================================
' Export_GammaInverseScaleHoldout2
'------------------------------------------------------------------------------
' PURPOSE
'   Fills the four observed columns of gamma_inverse_scale_grid2.csv, row by row.
'
' BEHAVIOR
'   - Reads the whole grid, evaluates every data row, rewrites the file.
'   - A row with fewer than 10 fields, or with inputs that are not 16
'     hexadecimal digits, is left as it is and not counted.
'==============================================================================
'
'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Dim Path                As String          'Resolved grid path
    Dim Lines()             As String          'File lines
    Dim Raw                 As String          'File contents
    Dim FileNum             As Integer         'Input file handle
    Dim OutNum              As Integer         'Output file handle
    Dim Cols                As Variant         'Split fields of one row
    Dim P                   As Double          'Probability
    Dim A                   As Double          'Shape
    Dim S                   As Double          'Scale
    Dim Filled              As Long            'Rows written
    Dim I                   As Long            'Row index
'------------------------------------------------------------------------------
' READ
'------------------------------------------------------------------------------
    On Error GoTo Err_Handler

    'Resolve the grid path (robust to OneDrive / SharePoint, where
    'ThisWorkbook.Path returns an http URL that Open cannot use)
        Path = ResolveGridPath()
        If Len(Path) = 0 Then Exit Sub          'User cancelled the picker

        Filled = 0

        FileNum = FreeFile
        Open Path For Input As #FileNum
        Raw = Input$(LOF(FileNum), FileNum)
        Close #FileNum

        Raw = Replace(Raw, vbCrLf, vbLf)
        Raw = Replace(Raw, vbCr, vbLf)
        Lines = Split(Raw, vbLf)
'------------------------------------------------------------------------------
' EVALUATE
'------------------------------------------------------------------------------
    'Row 0 is the header; data starts at row 1
        For I = 1 To UBound(Lines)
            If Len(Trim$(Lines(I))) = 0 Then GoTo ContinueRow

            Cols = Split(Lines(I), ",")
            If UBound(Cols) < 9 Then GoTo ContinueRow
            If Not IsBits(CStr(Cols(2))) Then GoTo ContinueRow
            If Not IsBits(CStr(Cols(3))) Then GoTo ContinueRow
            If Not IsBits(CStr(Cols(4))) Then GoTo ContinueRow

            P = FromBits(CStr(Cols(2)))
            A = FromBits(CStr(Cols(3)))
            S = FromBits(CStr(Cols(4)))

            Cols(6) = GammaInverse(P, A, S)
            If S = 2# Then
                Cols(7) = ChiSquareInverse(P, A)
            Else
                Cols(7) = "NA"
            End If
            Cols(8) = RoundTrip(CStr(Cols(6)), A, S)
            Cols(9) = LogUnit(P, A)

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

    MsgBox "Gamma inverse second holdout complete: " & Filled & _
           " observation(s) written to" & vbCrLf & Path, _
           vbInformation, "Gamma inverse second holdout"
    Exit Sub

Err_Handler:
    On Error Resume Next
    Close #FileNum
    Close #OutNum
    MsgBox "Gamma inverse second holdout failed: " & Err.Description, _
           vbExclamation, "Gamma inverse second holdout"
End Sub


Private Function GammaInverse( _
    ByVal P As Double, _
    ByVal A As Double, _
    ByVal S As Double) _
    As String
    Dim Result              As Variant         'Public function result

    On Error GoTo Err_Handler

    Result = K_STATS_Gamma_InverseCumulative(P, A, S)
    If IsError(Result) Then
        GammaInverse = "CVERR"
    Else
        GammaInverse = ToBits(CDbl(Result))
    End If
    Exit Function

Err_Handler:
    GammaInverse = "ERROR"
End Function


Private Function ChiSquareInverse( _
    ByVal P As Double, _
    ByVal A As Double) _
    As String
    Dim Result              As Variant         'Public function result

    On Error GoTo Err_Handler

    Result = K_STATS_ChiSquare_InverseCumulative(P, 2# * A)
    If IsError(Result) Then
        ChiSquareInverse = "CVERR"
    Else
        ChiSquareInverse = ToBits(CDbl(Result))
    End If
    Exit Function

Err_Handler:
    ChiSquareInverse = "ERROR"
End Function


Private Function RoundTrip( _
    ByVal QuantileBits As String, _
    ByVal A As Double, _
    ByVal S As Double) _
    As String
    Dim X                   As Double          'Returned quantile
    Dim Result              As Variant         'Public function result

    On Error GoTo Err_Handler

    RoundTrip = "NA"
    If Not IsBits(QuantileBits) Then Exit Function
    X = FromBits(QuantileBits)
    If Not (X > 0#) Then Exit Function

    Result = K_STATS_Gamma_Cumulative(X, A, S)
    If IsError(Result) Then
        RoundTrip = "CVERR"
    Else
        RoundTrip = ToBits(CDbl(Result))
    End If
    Exit Function

Err_Handler:
    RoundTrip = "ERROR"
End Function


Private Function LogUnit( _
    ByVal P As Double, _
    ByVal A As Double) _
    As String
    Dim LogGamma1p          As Double          'Log(Gamma(1 + A))
    Dim Value               As Double          'The classifying value
    Dim Msg                 As String          'Kernel diagnostic

    On Error GoTo Err_Handler

    If PROB_TryGammaInvLogUnit(P, A, LogGamma1p, Value, Msg) Then
        LogUnit = ToBits(Value)
    Else
        LogUnit = "FAIL"
    End If
    Exit Function

Err_Handler:
    LogUnit = "ERROR"
End Function


Private Function IsBits(ByVal Text As String) As Boolean
'
'==============================================================================
' IsBits
'------------------------------------------------------------------------------
' PURPOSE
'   True when Text is exactly 16 hexadecimal digits.
'==============================================================================
'
    Dim I                   As Long            'Character index

    If Len(Text) <> 16 Then Exit Function
    For I = 1 To 16
        If InStr(1, "0123456789ABCDEFabcdef", Mid$(Text, I, 1), vbBinaryCompare) = 0 Then
            Exit Function
        End If
    Next I
    IsBits = True
End Function


Private Function FromBits(ByVal Text As String) As Double
'
'==============================================================================
' FromBits
'------------------------------------------------------------------------------
' PURPOSE
'   The Double whose 64 bits are the 16 hexadecimal digits of Text, most
'   significant first.
'==============================================================================
'
    Dim Words               As GIS2_Words       'The two halves
    Dim Number              As GIS2_Number      'The same eight bytes as a Double

    Words.High = CLng("&H" & Left$(Text, 8))
    Words.Low = CLng("&H" & Mid$(Text, 9, 8))
    LSet Number = Words
    FromBits = Number.Value
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
    Dim Words               As GIS2_Words       'The two halves
    Dim Number              As GIS2_Number      'The same eight bytes as a Double

    Number.Value = X
    LSet Words = Number
    ToBits = Right$("00000000" & Hex$(Words.High), 8) & _
             Right$("00000000" & Hex$(Words.Low), 8)
End Function


Private Function ResolveGridPath() As String
'
'==============================================================================
' ResolveGridPath
'------------------------------------------------------------------------------
' PURPOSE
'   Finds gamma_inverse_scale_grid2.csv beside the workbook, or asks for it.
'==============================================================================
'
    Dim BookPath            As String          'Workbook folder
    Dim Candidate           As String          'Path next to the workbook
    Dim Picked              As Variant         'File-dialog result

    BookPath = ThisWorkbook.Path

    If Len(BookPath) > 0 And LCase$(Left$(BookPath, 4)) <> "http" Then
        Candidate = BookPath & Application.PathSeparator & _
                    "gamma_inverse_scale_grid2.csv"
        If Len(Dir$(Candidate)) > 0 Then
            ResolveGridPath = Candidate
            Exit Function
        End If
    End If

    MsgBox "Could not locate gamma_inverse_scale_grid2.csv automatically " & _
           "(the workbook may be on OneDrive/SharePoint). Please select it.", _
           vbInformation, "Locate Gamma inverse second holdout"

    Picked = Application.GetOpenFilename( _
        FileFilter:="Gamma inverse second holdout (*.csv),*.csv", _
        Title:="Select gamma_inverse_scale_grid2.csv")

    If VarType(Picked) = vbBoolean Then
        ResolveGridPath = vbNullString
    Else
        ResolveGridPath = CStr(Picked)
    End If
End Function
