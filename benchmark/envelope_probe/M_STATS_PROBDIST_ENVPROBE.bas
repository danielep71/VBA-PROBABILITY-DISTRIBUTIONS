Attribute VB_Name = "M_STATS_PROBDIST_ENVPROBE"
Option Explicit
'
'==============================================================================
' M_STATS_PROBDIST_ENVPROBE
'------------------------------------------------------------------------------
' PURPOSE
'   Fills observed_vba in benchmark/envelope_probe/envelope_probe_grid.csv by
'   calling the PUBLIC incomplete-function kernels directly.
'
' WHY THIS EXISTS
'   The StudentT, ChiSquare and F accuracy envelopes reject degrees of freedom
'   above their caps, so the public UDFs cannot measure whether those caps are
'   still justified after the CR-P1-02 prefactor repair. The kernels beneath
'   them are Public and carry no envelope, so probing them directly answers the
'   question without editing the caps first. Each grid row supplies exactly the
'   arguments the corresponding public function would pass.
'
' GRID FORMAT (study grid; 12 columns, arg4 present)
'   function, vba_kernel, claim, metric, arg1, arg2, arg3, arg4,
'   reference, observed_vba, regime, evidence_set
'
' BEHAVIOR
'   Writes the two-part hi;lo representation used by the other study exports, or
'   the token ERROR when the kernel reports failure. A kernel that refuses is a
'   measurement: it marks the edge of the reachable domain.
'
' DEPENDENCIES
'   - PROB_TryGammaRegularizedP
'   - PROB_TryBetaRegularized
'
' UPDATED
'   2026-07-29
'==============================================================================
'
Private Const GRID_FILE As String = "envelope_probe_grid.csv"


Public Sub Export_EnvelopeProbe()
'
'==============================================================================
' Export_EnvelopeProbe
'------------------------------------------------------------------------------
' PURPOSE
'   Reads the probe grid, evaluates every row and writes observed_vba back.
'
' ERROR POLICY
'   Reports failures through a message box; file handles are always closed.
'==============================================================================
'
'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Dim Path                As String          'Resolved grid location
    Dim Lines()             As String          'Whole file, split on vbLf
    Dim Parts()             As String          'Current row, split on comma
    Dim Buffer              As String          'Raw file contents
    Dim RowText             As String          'Current line
    Dim Idx                 As Long            'Row index
    Dim FileNo              As Integer         'File handle
    Dim Evaluated           As Long            'Rows evaluated
    Dim Errored             As Long            'Rows that refused
'------------------------------------------------------------------------------
' INITIALIZE
'------------------------------------------------------------------------------
    On Error GoTo Err_Handler

    FileNo = 0
    Evaluated = 0
    Errored = 0

    Path = ResolveGridPath()
    If Len(Path) = 0 Then Exit Sub
'------------------------------------------------------------------------------
' READ
'------------------------------------------------------------------------------
    FileNo = FreeFile
    Open Path For Binary As #FileNo
    Buffer = Space$(LOF(FileNo))
    Get #FileNo, , Buffer
    Close #FileNo
    FileNo = 0

    Buffer = Replace$(Buffer, vbCrLf, vbLf)
    Lines = Split(Buffer, vbLf)
'------------------------------------------------------------------------------
' EVALUATE
'------------------------------------------------------------------------------
    For Idx = 1 To UBound(Lines)
        RowText = Lines(Idx)

        If Len(Trim$(RowText)) > 0 Then
            Parts = Split(RowText, ",")

            If UBound(Parts) >= 11 Then
                Parts(9) = EvaluateKernel(Parts(0), Parts(4), Parts(5), _
                                          Parts(6), Parts(7))
                If Parts(9) = "ERROR" Then
                    Errored = Errored + 1
                Else
                    Evaluated = Evaluated + 1
                End If
                Lines(Idx) = Join(Parts, ",")
            End If
        End If
    Next Idx
'------------------------------------------------------------------------------
' WRITE
'------------------------------------------------------------------------------
    FileNo = FreeFile
    Open Path For Output As #FileNo
    Print #FileNo, Join(Lines, vbLf);
    Close #FileNo
    FileNo = 0

    MsgBox "Envelope probe complete." & vbCrLf & _
           "Evaluated: " & Evaluated & vbCrLf & _
           "Refused (ERROR): " & Errored & vbCrLf & vbCrLf & _
           "Now run:  python analyze_envelope_probe.py", _
           vbInformation, "Envelope probe"
    Exit Sub

'------------------------------------------------------------------------------
' ERROR HANDLER
'------------------------------------------------------------------------------
Err_Handler:
    On Error Resume Next
    If FileNo <> 0 Then Close #FileNo
    MsgBox "Envelope probe failed: " & Err.Description, vbExclamation
End Sub


Private Function EvaluateKernel( _
    ByVal FunctionName As String, _
    ByVal S1 As String, _
    ByVal S2 As String, _
    ByVal S3 As String, _
    ByVal S4 As String) _
    As String
'
'==============================================================================
' EvaluateKernel
'------------------------------------------------------------------------------
' PURPOSE
'   Dispatches the probed kernel and returns a token: the hi;lo pair on success,
'   or ERROR when the kernel's Boolean contract reports failure.
'==============================================================================
'
'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Dim Result              As Double          'Kernel output
    Dim FailMsg             As String          'Kernel failure detail
    Dim Ok                  As Boolean         'Kernel success flag
'------------------------------------------------------------------------------
' INITIALIZE
'------------------------------------------------------------------------------
    On Error GoTo Err_Handler

    Ok = False
'------------------------------------------------------------------------------
' COMPUTE
'------------------------------------------------------------------------------
    Select Case Trim$(FunctionName)
        Case "PROB_TryGammaRegularizedP"
            Ok = PROB_TryGammaRegularizedP(ParseDouble(S1), ParseDouble(S2), _
                                           Result, FailMsg)

        Case "PROB_TryBetaRegularized"
            Ok = PROB_TryBetaRegularized(ParseDouble(S1), ParseDouble(S2), _
                                         ParseDouble(S3), ParseDouble(S4), _
                                         Result, FailMsg)

        Case Else
            EvaluateKernel = "ERROR"
            Exit Function
    End Select

    If Not Ok Then
        EvaluateKernel = "ERROR"
        Exit Function
    End If
'------------------------------------------------------------------------------
' RETURN SUCCESS
'------------------------------------------------------------------------------
    EvaluateKernel = FormatFullPrecision(Result)
    Exit Function

'------------------------------------------------------------------------------
' ERROR HANDLER
'------------------------------------------------------------------------------
Err_Handler:
    EvaluateKernel = "ERROR"
End Function


Private Function ResolveGridPath() As String
'
'==============================================================================
' ResolveGridPath
'------------------------------------------------------------------------------
' PURPOSE
'   Returns a usable LOCAL path to envelope_probe_grid.csv, or an empty string
'   if the user cancels.
'
' WHY THIS EXISTS
'   ThisWorkbook.Path returns an http(s) URL when the workbook lives on OneDrive
'   or SharePoint, and Open cannot read a URL. Dir$ raises on such a path rather
'   than returning empty, so the URL case must be excluded BEFORE Dir$ is called,
'   not detected by its result. This prefers a local workbook folder and
'   otherwise asks the user.
'==============================================================================
'
'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Dim BookPath            As String          'Workbook folder
    Dim Candidate           As String          'Path next to the workbook
    Dim Picked              As Variant         'File-dialog result
'------------------------------------------------------------------------------
' RESOLVE
'------------------------------------------------------------------------------
    BookPath = ThisWorkbook.Path
    If Len(BookPath) > 0 And LCase$(Left$(BookPath, 4)) <> "http" Then
        Candidate = BookPath & Application.PathSeparator & GRID_FILE
        If Len(Dir$(Candidate)) > 0 Then
            ResolveGridPath = Candidate
            Exit Function
        End If
    End If

    MsgBox "Could not locate " & GRID_FILE & " automatically " & _
           "(the workbook may be on OneDrive/SharePoint). Please select it.", _
           vbInformation, "Locate probe grid"
    Picked = Application.GetOpenFilename( _
        FileFilter:="Probe grid (*.csv),*.csv", _
        Title:="Select " & GRID_FILE)

    If VarType(Picked) = vbBoolean Then
        ResolveGridPath = vbNullString
    Else
        ResolveGridPath = CStr(Picked)
    End If
End Function


Private Function ParseDouble(ByVal Text As String) As Double
'
'==============================================================================
' ParseDouble
'------------------------------------------------------------------------------
' PURPOSE
'   Parses an invariant decimal string independently of the Excel locale.
'
' WHY Val AND NOT CDbl
'   CDbl honours the locale: under an Italian locale it reads "." as a THOUSANDS
'   separator, so "1.5865E-01" becomes 1.5865E+13. Val always reads "." as the
'   decimal point. Any stray comma is normalised first.
'==============================================================================
'
    Dim Clean               As String          'Normalised input

    Clean = Trim$(Text)
    If Len(Clean) = 0 Then
        ParseDouble = 0#
        Exit Function
    End If

    Clean = Replace$(Clean, ",", ".")
    ParseDouble = Val(Clean)
End Function


Private Function FormatFullPrecision(ByVal X As Double) As String
'
'==============================================================================
' FormatFullPrecision
'------------------------------------------------------------------------------
' PURPOSE
'   Writes X as a hi;lo pair so no accuracy is lost through 15-digit text.
'==============================================================================
'
    Dim Hi                  As Double          'Leading 15-digit part
    Dim Lo                  As Double          'Remainder

    If X = 0# Then
        FormatFullPrecision = "0"
        Exit Function
    End If

    Hi = Val(Fmt15(X))                     'Val is locale-independent; CDbl is not
    Lo = X - Hi

    If Lo = 0# Then
        FormatFullPrecision = Fmt15(Hi)
    Else
        FormatFullPrecision = Fmt15(Hi) & ";" & Fmt15(Lo)
    End If
End Function


Private Function Fmt15(ByVal X As Double) As String
'
'==============================================================================
' Fmt15
'------------------------------------------------------------------------------
' PURPOSE
'   Formats X with 15 significant digits using an invariant decimal point.
'==============================================================================
'
    Dim Text                As String          'Formatted value
    Dim Sep                 As String          'Locale decimal separator

    Text = Format$(X, "0.00000000000000E+00")
    Sep = Mid$(CStr(1.5), 2, 1)
    If Sep <> "." Then Text = Replace$(Text, Sep, ".")

    Fmt15 = Text
End Function


