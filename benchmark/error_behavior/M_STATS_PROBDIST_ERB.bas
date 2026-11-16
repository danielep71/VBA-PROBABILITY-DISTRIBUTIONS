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
'   passes a text or an empty cell. Every other call goes through
'   Application.Run with Double arguments, which is how VBA code can pass
'   a value no cell can hold.
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
    Dim Target              As String          'The function, qualified by this workbook
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
        Target = "'" & ThisWorkbook.Name & "'!" & FunctionName
        Select Case UBound(Tokens)
            Case 0
                V = Application.Run(Target, Argument(Tokens(0)))
            Case 1
                V = Application.Run(Target, Argument(Tokens(0)), Argument(Tokens(1)))
            Case 2
                V = Application.Run(Target, Argument(Tokens(0)), Argument(Tokens(1)), _
                                    Argument(Tokens(2)))
            Case 3
                V = Application.Run(Target, Argument(Tokens(0)), Argument(Tokens(1)), _
                                    Argument(Tokens(2)), Argument(Tokens(3)))
            Case Else
                Observed = "UNSUPPORTED"
                Exit Function
        End Select
    End If

    Observed = Described(V)
    Exit Function

Raised:
    Observed = "RAISED" & CStr(Err.Number)
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
