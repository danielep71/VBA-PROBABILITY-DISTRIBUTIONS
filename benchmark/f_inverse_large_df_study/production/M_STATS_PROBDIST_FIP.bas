Attribute VB_Name = "M_STATS_PROBDIST_FIP"
Option Explicit

'One Double and its two 32-bit words, for exact transport through text
Private Type FIP_Words
    Low                     As Long
    High                    As Long
End Type

Private Type FIP_Number
    Value                   As Double
End Type

'==============================================================================
' M_STATS_PROBDIST_FIP
'------------------------------------------------------------------------------
' PURPOSE
'   Export macros for the #35 production holdout of the beta inverse
'   (benchmark/f_inverse_large_df_study/production/PRODUCTION_PREREGISTRATION.md).
'   At each row they call the public inverse the row names and record its
'   result as the 16 hexadecimal digits of its Double.
'
' USAGE
'   Export_FInversePairBaseline fills f_inverse_pair_baseline.csv and is run
'   on the source before the change. Export_FInversePair fills
'   f_inverse_pair_grid.csv and is run on the candidate. The two files hold
'   the same rows. Each macro finds its file beside the workbook, or asks
'   for it; ci/Export-ExcelGrids.ps1 runs them on the self-hosted runner with
'   the file staged beside the export workbook (#225).
'
' GRID FORMAT (study grid; 8 columns)
'   point_id, kind, func, p_bits, a_bits, b_bits, ref, value
'
'   The macros write the last column. The arguments are rebuilt from their
'   bit patterns, so no decimal text is parsed. They never read the
'   reference, so the observed side stays independent.
'
'       func F      K_STATS_F_InverseCumulative(P, A, B)
'       func T      K_STATS_StudentT_InverseCumulative(P, A)
'       func Beta   K_STATS_Beta_InverseCumulative(P, A, B)
'
' ERROR POLICY
'   A public function that returns a worksheet error is recorded as CVERR and
'   a runtime fault as ERROR. Neither is ever scored as a value.
'
' DEPENDENCIES
'   - K_STATS_F_InverseCumulative, K_STATS_StudentT_InverseCumulative
'                                               (M_STATS_PROBDIST_TFAMILY)
'   - K_STATS_Beta_InverseCumulative            (M_STATS_PROBDIST_CONTINUOUS)
'
' UPDATED
'   2026-11-14
'==============================================================================


Public Sub Export_FInversePairBaseline()
'
'==============================================================================
' Export_FInversePairBaseline
'------------------------------------------------------------------------------
' PURPOSE
'   Fills the value column of f_inverse_pair_baseline.csv.
'==============================================================================
'
    Dim Path                As String          'File written
    Dim Filled              As Long            'Rows written, or -1

    Filled = ExportFile("f_inverse_pair_baseline.csv", "F inverse pair baseline", Path)
    If Filled < 0 Then Exit Sub

    MsgBox "F inverse pair baseline complete: " & Filled & _
           " observation(s) written to" & vbCrLf & Path, _
           vbInformation, "F inverse pair baseline"
End Sub


Public Sub Export_FInversePair()
'
'==============================================================================
' Export_FInversePair
'------------------------------------------------------------------------------
' PURPOSE
'   Fills the value column of f_inverse_pair_grid.csv.
'==============================================================================
'
    Dim Path                As String          'File written
    Dim Filled              As Long            'Rows written, or -1

    Filled = ExportFile("f_inverse_pair_grid.csv", "F inverse pair holdout", Path)
    If Filled < 0 Then Exit Sub

    MsgBox "F inverse pair holdout complete: " & Filled & _
           " observation(s) written to" & vbCrLf & Path, _
           vbInformation, "F inverse pair holdout"
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
'   Fills the value column of one grid file, row by row. Returns the number
'   of rows written, or -1 when the file was not chosen or the export
'   failed; the failure is reported here, the success by the caller.
'
' BEHAVIOR
'   - Reads the whole file, evaluates every data row, rewrites the file.
'   - A row with fewer than 8 fields, or with p_bits or a_bits that are not
'     16 hexadecimal digits, is left as it is and not counted.
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
    Dim P                   As Double          'Probability
    Dim A                   As Double          'First parameter
    Dim B                   As Double          'Second parameter, where there is one
    Dim Filled              As Long            'Rows written
    Dim I                   As Long            'Row index
'------------------------------------------------------------------------------
' READ
'------------------------------------------------------------------------------
    On Error GoTo Err_Handler

    'Resolve the grid path (robust to OneDrive / SharePoint, where
    'ThisWorkbook.Path returns an http URL that Open cannot use)
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
' EVALUATE
'------------------------------------------------------------------------------
    'Row 0 is the header; data starts at row 1
        For I = 1 To UBound(Lines)
            If Len(Trim$(Lines(I))) = 0 Then GoTo ContinueRow

            Cols = Split(Lines(I), ",")
            If UBound(Cols) < 7 Then GoTo ContinueRow
            If Not IsBits(CStr(Cols(3))) Then GoTo ContinueRow
            If Not IsBits(CStr(Cols(4))) Then GoTo ContinueRow

            P = FromBits(CStr(Cols(3)))
            A = FromBits(CStr(Cols(4)))
            B = 0#
            If IsBits(CStr(Cols(5))) Then B = FromBits(CStr(Cols(5)))

            Cols(7) = Quantile(CStr(Cols(2)), P, A, B)

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


Private Function Quantile( _
    ByVal Func As String, _
    ByVal P As Double, _
    ByVal A As Double, _
    ByVal B As Double) _
    As String
    Dim Result              As Variant         'Public function result

    On Error GoTo Err_Handler

    Select Case Func
        Case "F"
            Result = K_STATS_F_InverseCumulative(P, A, B)
        Case "T"
            Result = K_STATS_StudentT_InverseCumulative(P, A)
        Case "Beta"
            Result = K_STATS_Beta_InverseCumulative(P, A, B)
        Case Else
            Quantile = "ERROR"
            Exit Function
    End Select

    If IsError(Result) Then
        Quantile = "CVERR"
    Else
        Quantile = ToBits(CDbl(Result))
    End If
    Exit Function

Err_Handler:
    Quantile = "ERROR"
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
    Dim Words               As FIP_Words       'The two halves
    Dim Number              As FIP_Number      'The same eight bytes as a Double

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
    Dim Words               As FIP_Words       'The two halves
    Dim Number              As FIP_Number      'The same eight bytes as a Double

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
