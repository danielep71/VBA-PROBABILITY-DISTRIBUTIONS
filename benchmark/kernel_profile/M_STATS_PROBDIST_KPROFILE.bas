Attribute VB_Name = "M_STATS_PROBDIST_KPROFILE"
Option Explicit

Private Declare PtrSafe Function QueryPerformanceCounter Lib "kernel32" (ByRef Value As Currency) As Long
Private Declare PtrSafe Function QueryPerformanceFrequency Lib "kernel32" (ByRef Value As Currency) As Long
Private Type KP_Number
    Value As Double
End Type
Private Type KP_Words
    Low As Long
    High As Long
End Type
Private KP_Frequency As Currency
Private KP_Depth As Long
Private KP_Id(1 To 128) As Long
Private KP_Start(1 To 128) As Currency
Private KP_Child(1 To 128) As Currency
Private KP_Count(1 To 32) As Long
Private KP_Inclusive(1 To 32) As Currency
Private KP_Exclusive(1 To 32) As Currency

Public Sub KP_Enter(ByVal Id As Long)
    KP_Depth = KP_Depth + 1
    If KP_Depth > 128 Then Err.Raise 5, "KP", "stack overflow"
    KP_Id(KP_Depth) = Id
    KP_Child(KP_Depth) = 0@
    QueryPerformanceCounter KP_Start(KP_Depth)
End Sub


Public Sub KP_Leave(ByVal Id As Long)
    Dim Stamp As Currency, Duration As Currency
    QueryPerformanceCounter Stamp
    If KP_Depth < 1 Then Err.Raise 5, "KP", "stack underflow"
    If KP_Id(KP_Depth) <> Id Then Err.Raise 5, "KP", "unwound error or mismatched scope"
    Duration = Stamp - KP_Start(KP_Depth)
    KP_Inclusive(Id) = KP_Inclusive(Id) + Duration
    KP_Exclusive(Id) = KP_Exclusive(Id) + Duration - KP_Child(KP_Depth)
    KP_Count(Id) = KP_Count(Id) + 1
    KP_Depth = KP_Depth - 1
    If KP_Depth > 0 Then KP_Child(KP_Depth) = KP_Child(KP_Depth) + Duration
End Sub


Private Function KP_Bits(ByVal Value As Double) As String
    Dim N As KP_Number, W As KP_Words
    N.Value = Value
    LSet W = N
    KP_Bits = Right$("00000000" & Hex$(W.High), 8) & Right$("00000000" & Hex$(W.Low), 8)
End Function


Private Function KP_Decimal(ByVal Value As Double) As String
    KP_Decimal = Replace(Format$(Value, "0.000000"), ",", ".")
End Function


Private Function KP_Evaluate(ByVal A As Double, ByVal X As Double, ByVal Route As String) As Double
    Dim V As Double, Message As String, OK As Boolean, Result As Variant
    Select Case Route
        Case "p": OK = PROB_TryGammaRegularizedP(A, X, V, Message)
        Case "q": OK = PROB_TryGammaRegularizedQ(A, X, V, Message)
        Case "inverse"
            Result = K_STATS_Gamma_InverseCumulative(X, A, 1#)
            If IsError(Result) Then Err.Raise 5, "KP", "inverse failed"
            V = CDbl(Result): OK = True
        Case Else: Err.Raise 5, "KP", "unknown route"
    End Select
    If Not OK Then Err.Raise 5, "KP", Message
    KP_Evaluate = V
End Function


Public Function KP_Run(ByVal A As Double, ByVal X As Double, ByVal Route As String, ByVal Repeats As Long) As String
    Dim T0 As Currency, T1 As Currency, I As Long, V As Double, Expected As String
    Dim Out As String
    QueryPerformanceFrequency KP_Frequency
    If KP_Frequency <= 0@ Or Repeats < 1 Then Err.Raise 5, "KP", "invalid clock or repeats"
    KP_Depth = 0
    V = KP_Evaluate(A, X, Route)
    Expected = KP_Bits(V)
    If KP_Depth <> 0 Then Err.Raise 5, "KP", "unbalanced warmup"
    Erase KP_Count: Erase KP_Inclusive: Erase KP_Exclusive
    QueryPerformanceCounter T0
    For I = 1 To Repeats
        V = KP_Evaluate(A, X, Route)
        If KP_Bits(V) <> Expected Then Err.Raise 5, "KP", "unstable bits"
    Next I
    QueryPerformanceCounter T1
    If KP_Depth <> 0 Then Err.Raise 5, "KP", "unbalanced sample"
    Out = "TOTAL," & Expected & "," & CStr(Repeats) & "," & KP_Decimal((T1 - T0) / KP_Frequency * 1000000#) & vbLf
    For I = 1 To 32
        If KP_Count(I) > 0 Then
            Out = Out & "SCOPE," & CStr(I) & "," & CStr(KP_Count(I)) & "," & _
                KP_Decimal(KP_Inclusive(I) / KP_Frequency * 1000000#) & "," & _
                KP_Decimal(KP_Exclusive(I) / KP_Frequency * 1000000#) & vbLf
        End If
    Next I
    KP_Run = Out
End Function


Public Function KP_Calibrate() As String
    Dim I As Long, T0 As Currency, T1 As Currency
    QueryPerformanceFrequency KP_Frequency
    KP_Depth = 0
    QueryPerformanceCounter T0
    For I = 1 To 100000
        KP_Enter 32
        KP_Leave 32
    Next I
    QueryPerformanceCounter T1
    KP_Calibrate = KP_Decimal((T1 - T0) / KP_Frequency * 10#)
End Function
