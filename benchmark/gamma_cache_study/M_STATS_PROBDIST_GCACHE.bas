Attribute VB_Name = "M_STATS_PROBDIST_GCACHE"
Option Explicit
Private Declare PtrSafe Function QueryPerformanceCounter Lib "kernel32" (ByRef Value As Currency) As Long
Private Declare PtrSafe Function QueryPerformanceFrequency Lib "kernel32" (ByRef Value As Currency) As Long
Private Type GC_Number
    Value As Double
End Type
Private Type GC_Words
    Low As Long
    High As Long
End Type

Private Function GC_Bits(ByVal Value As Double) As String
    Dim N As GC_Number, W As GC_Words
    N.Value = Value
    LSet W = N
    GC_Bits = Right$("00000000" & Hex$(W.High), 8) & Right$("00000000" & Hex$(W.Low), 8)
End Function


Private Function GC_Eval(ByVal A As Double, ByVal X As Double, ByVal Route As String) As Double
    Dim V As Double, Message As String, OK As Boolean, Result As Variant
    Select Case Route
        Case "p": OK = PROB_TryGammaRegularizedP(A, X, V, Message)
        Case "q": OK = PROB_TryGammaRegularizedQ(A, X, V, Message)
        Case "inverse"
            Result = K_STATS_Gamma_InverseCumulative(X, A, 1#)
            If IsError(Result) Then Err.Raise 5, "GC", "inverse failed"
            V = CDbl(Result): OK = True
        Case Else: Err.Raise 5, "GC", "unknown route"
    End Select
    If Not OK Then Err.Raise 5, "GC", Message
    GC_Eval = V
End Function


Public Function GC_Sample(ByVal A As Double, ByVal X As Double, ByVal Route As String, ByVal Cold As Boolean) As String
    Dim T0 As Currency, T1 As Currency, Frequency As Currency, Elapsed As Currency
    Dim I As Long, N As Long, V As Double, Discard As Double, Expected As String
    N = 4
    If Route = "inverse" Then N = 1
    QueryPerformanceFrequency Frequency
    V = GC_Eval(A, X, Route)
    Expected = GC_Bits(V)
    For I = 1 To N
        If Cold Then Discard = GC_Eval(0.75, 1#, "p")
        QueryPerformanceCounter T0
        V = GC_Eval(A, X, Route)
        QueryPerformanceCounter T1
        Elapsed = Elapsed + T1 - T0
        If GC_Bits(V) <> Expected Then Err.Raise 5, "GC", "unstable bits"
    Next I
    GC_Sample = Expected & "," & CStr(N) & "," & Replace(Format$(Elapsed / Frequency * 1000000#, "0.000000"), ",", ".")
End Function
