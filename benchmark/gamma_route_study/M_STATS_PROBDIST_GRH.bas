Attribute VB_Name = "M_STATS_PROBDIST_GRH"
Option Explicit

'One Double and its two 32-bit words, for exact transport through text
Private Type GRH_Words
    Low                     As Long
    High                    As Long
End Type

Private Type GRH_Number
    Value                   As Double
End Type

'==============================================================================
' M_STATS_PROBDIST_GRH
'------------------------------------------------------------------------------
' PURPOSE
'   Export macro for the #314 routing holdout. Fills the observed side of
'   holdout_route_grid.csv (FAST_ROUTE_PREREGISTRATION.md): at each (A, X) it
'   calls the four public routes and the two double-double entry points, and
'   records each result as the 16 hexadecimal digits of its Double.
'
' USAGE
'   Run Export_GammaRouteHoldout. It finds the grid beside the workbook, or
'   asks for it. ci/Export-ExcelGrids.ps1 runs it on the self-hosted runner
'   with the grid staged beside the export workbook (#225).
'
' GRID FORMAT (study grid; 12 columns)
'   point_id, kind, a_bits, x_bits, ref_p, ref_q, gamma_cdf, gamma_sf,
'   chi_cdf, chi_sf, dd_p, dd_q
'
'   This macro writes the last six columns. A and X are rebuilt from a_bits
'   and x_bits, so no decimal text is parsed. It never reads the references,
'   so the observed side stays independent.
'
'       gamma_cdf   K_STATS_Gamma_Cumulative(X, A, 1)
'       gamma_sf    K_STATS_Gamma_Survival(X, A, 1)
'       chi_cdf     K_STATS_ChiSquare_Cumulative(2X, 2A)
'       chi_sf      K_STATS_ChiSquare_Survival(2X, 2A)
'       dd_p        PROB_TryGammaRegularizedP(A, X)
'       dd_q        PROB_TryGammaRegularizedQ(A, X)
'
' ERROR POLICY
'   A public function that returns a worksheet error is recorded as CVERR, a
'   kernel that returns False as FAIL, and a runtime fault as ERROR. None is
'   ever scored as a value.
'
' DEPENDENCIES
'   - K_STATS_Gamma_Cumulative, _Survival      (M_STATS_PROBDIST_CONTINUOUS)
'   - K_STATS_ChiSquare_Cumulative, _Survival  (M_STATS_PROBDIST_TFAMILY)
'   - PROB_TryGammaRegularizedP, _Q            (M_STATS_PROBDIST_SPECIALFUNCS)
'
' UPDATED
'   2026-11-14
'==============================================================================


Public Sub Export_GammaRouteHoldout()
'
'==============================================================================
' Export_GammaRouteHoldout
'------------------------------------------------------------------------------
' PURPOSE
'   Fills the six observed columns of holdout_route_grid.csv, row by row.
'
' BEHAVIOR
'   - Reads the whole grid, evaluates every data row, rewrites the file.
'   - A row with fewer than 12 fields, or with a_bits or x_bits that are not
'     16 hexadecimal digits, is left as it is and not counted.
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
    Dim A                   As Double          'Shape
    Dim X                   As Double          'Argument
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
            If UBound(Cols) < 11 Then GoTo ContinueRow
            If Not IsBits(CStr(Cols(2))) Then GoTo ContinueRow
            If Not IsBits(CStr(Cols(3))) Then GoTo ContinueRow

            A = FromBits(CStr(Cols(2)))
            X = FromBits(CStr(Cols(3)))

            Cols(6) = PublicValue("gamma_cdf", A, X)
            Cols(7) = PublicValue("gamma_sf", A, X)
            Cols(8) = PublicValue("chi_cdf", A, X)
            Cols(9) = PublicValue("chi_sf", A, X)
            Cols(10) = KernelValue("p", A, X)
            Cols(11) = KernelValue("q", A, X)

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

    MsgBox "Gamma routing holdout complete: " & Filled & _
           " observation(s) written to" & vbCrLf & Path, _
           vbInformation, "Gamma routing holdout"
    Exit Sub

Err_Handler:
    On Error Resume Next
    Close #FileNum
    Close #OutNum
    MsgBox "Gamma routing holdout failed: " & Err.Description, _
           vbExclamation, "Gamma routing holdout"
End Sub


Private Function PublicValue( _
    ByVal Route As String, _
    ByVal A As Double, _
    ByVal X As Double) _
    As String
'
'==============================================================================
' PublicValue
'------------------------------------------------------------------------------
' PURPOSE
'   Calls one public route and returns its result as 16 hexadecimal digits,
'   CVERR for a worksheet error, or ERROR for a runtime fault.
'
' NOTES
'   The Chi-square routes take 2X and 2A. Doubling a Double is exact unless it
'   overflows, which no row of the design reaches.
'==============================================================================
'
    Dim Result              As Variant         'Public function result

    On Error GoTo Err_Handler

    Select Case Route
        Case "gamma_cdf"
            Result = K_STATS_Gamma_Cumulative(X, A, 1#)
        Case "gamma_sf"
            Result = K_STATS_Gamma_Survival(X, A, 1#)
        Case "chi_cdf"
            Result = K_STATS_ChiSquare_Cumulative(2# * X, 2# * A)
        Case Else
            Result = K_STATS_ChiSquare_Survival(2# * X, 2# * A)
    End Select

    If IsError(Result) Then
        PublicValue = "CVERR"
    Else
        PublicValue = ToBits(CDbl(Result))
    End If
    Exit Function

Err_Handler:
    PublicValue = "ERROR"
End Function


Private Function KernelValue( _
    ByVal Tail As String, _
    ByVal A As Double, _
    ByVal X As Double) _
    As String
'
'==============================================================================
' KernelValue
'------------------------------------------------------------------------------
' PURPOSE
'   Calls one double-double entry point and returns its result as 16
'   hexadecimal digits, FAIL when it returns False, or ERROR for a runtime
'   fault.
'==============================================================================
'
    Dim Value               As Double          'Kernel result
    Dim Msg                 As String          'Kernel diagnostic
    Dim OK                  As Boolean         'Kernel succeeded

    On Error GoTo Err_Handler

    If Tail = "p" Then
        OK = PROB_TryGammaRegularizedP(A, X, Value, Msg)
    Else
        OK = PROB_TryGammaRegularizedQ(A, X, Value, Msg)
    End If

    If OK Then
        KernelValue = ToBits(Value)
    Else
        KernelValue = "FAIL"
    End If
    Exit Function

Err_Handler:
    KernelValue = "ERROR"
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
    Dim Words               As GRH_Words       'The two halves
    Dim Number              As GRH_Number      'The same eight bytes as a Double

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
    Dim Words               As GRH_Words       'The two halves
    Dim Number              As GRH_Number      'The same eight bytes as a Double

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
'   Finds holdout_route_grid.csv beside the workbook, or asks for it.
'==============================================================================
'
    Dim BookPath            As String          'Workbook folder
    Dim Candidate           As String          'Path next to the workbook
    Dim Picked              As Variant         'File-dialog result

    BookPath = ThisWorkbook.Path

    If Len(BookPath) > 0 And LCase$(Left$(BookPath, 4)) <> "http" Then
        Candidate = BookPath & Application.PathSeparator & _
                    "holdout_route_grid.csv"
        If Len(Dir$(Candidate)) > 0 Then
            ResolveGridPath = Candidate
            Exit Function
        End If
    End If

    MsgBox "Could not locate holdout_route_grid.csv automatically " & _
           "(the workbook may be on OneDrive/SharePoint). Please select it.", _
           vbInformation, "Locate Gamma routing holdout"

    Picked = Application.GetOpenFilename( _
        FileFilter:="Gamma routing holdout (*.csv),*.csv", _
        Title:="Select holdout_route_grid.csv")

    If VarType(Picked) = vbBoolean Then
        ResolveGridPath = vbNullString
    Else
        ResolveGridPath = CStr(Picked)
    End If
End Function
