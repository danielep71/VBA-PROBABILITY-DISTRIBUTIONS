"""
Timing instrumentation of the VBA regression harness (#298).

tests/M_STATS_PROBDIST_TEST.bas records how long each test procedure takes:
the suite drivers, and the study driver of #300, call TimingStart once and TimingLap after every test
procedure, and one marked block holds the three procedures that keep the
record. The CI bridge (ci/Run-ExcelVbaTests.ps1) reads it after the run and
writes it to the Excel log, which the certification record binds by SHA-256.

The instrumentation is generated, never edited by hand: apply() inserts it
deterministically and strip() removes it exactly. Studies that compare the
test module with a frozen version (public_integration.restore and the tests
built on it) call strip() first, so timing never reads as a change of a
frozen study. test_harness_timing.py proves apply(strip(m)) == m for the
committed module.

A procedure in NESTED is a test procedure that only chains others (#13): each
call inside it gets a lap too, and no TimingStart, because a lap measures from
the previous mark. The procedure's own lap in its suite driver then holds only
what it runs after its last chained call.

Both functions work on LF text.
"""
import re

BEGIN = "' BEGIN harness timing (#298)"
END = "' END harness timing (#298)"
DRIVERS = ("RunCoreSuite", "RunNormalFamilySuite", "RunTFamilySuite",
           "RunContinuousSuite", "RunDiscreteSuite", "RunStudyBatteries",
           "RunSlowProductionBatteries")
NESTED = ("Test_CN_GammaSurvival",)
ANCHOR = "Private Sub BeginRun("
CALL = re.compile(r"^    (Test_\w+)$")
INSTRUMENT = re.compile(r'^    Timing(?:Start|Lap) "\w+"\n', re.M)

BLOCK = '''' BEGIN harness timing (#298)
Private Function HarnessTiming( _
    ByVal Action As String, _
    Optional ByVal Label As String = "") _
    As String
'
'==============================================================================
' HarnessTiming
'------------------------------------------------------------------------------
' PURPOSE
'   Records the elapsed time of each test procedure the suite drivers call, so
'   the CI log shows where the suite spends its time (#298).
'
' BEHAVIOR
'   "start" marks a suite, "lap" records the seconds since the previous mark,
'   "read" returns the record, "reset" clears it. The state is Static, so the
'   module needs no module-level declaration. Timer restarts at midnight; a
'   negative interval is corrected by one day.
'
' GENERATED
'   By benchmark/_harness_timing.py, with the TimingStart and TimingLap calls
'   of the suite drivers. Do not edit by hand.
'==============================================================================
'
    Static LapStart         As Double          'Timer at the previous mark
    Static Record           As String          'One line per mark
    Dim Clock               As Double          'Timer now
    Dim Elapsed             As Double          'Seconds since the previous mark

    Clock = Timer
    Select Case Action
        Case "start"
            Record = Record & "suite " & Label & vbLf
            LapStart = Clock
        Case "lap"
            Elapsed = Clock - LapStart
            If Elapsed < 0# Then Elapsed = Elapsed + 86400#
            Record = Record & Label & " " & Replace(Format$(Elapsed, "0.000"), ",", ".") & vbLf
            LapStart = Clock
        Case "read"
            HarnessTiming = Record
        Case "reset"
            Record = vbNullString
            LapStart = Clock
    End Select
End Function


Private Sub TimingStart(ByVal Label As String)
'
'==============================================================================
' TimingStart
'------------------------------------------------------------------------------
' PURPOSE
'   Marks the start of a suite driver in the timing record (#298).
'==============================================================================
'
    HarnessTiming "start", Label
End Sub


Private Sub TimingLap(ByVal Label As String)
'
'==============================================================================
' TimingLap
'------------------------------------------------------------------------------
' PURPOSE
'   Records the seconds the named test procedure took (#298).
'==============================================================================
'
    HarnessTiming "lap", Label
End Sub
' END harness timing (#298)


'''


class TimingError(ValueError):
    pass


def instrumented(text):
    return BEGIN in text or INSTRUMENT.search(text) is not None


def strip(text):
    """The module without timing: every TimingStart/TimingLap line and the block."""
    if BEGIN in text:
        if text.count(BLOCK) != 1:
            raise TimingError("the timing block differs from the generator's")
        text = text.replace(BLOCK, "", 1)
    return INSTRUMENT.sub("", text)


def _driver_span(text, name):
    start = text.find(f"Private Sub {name}()\n")
    if start < 0:
        raise TimingError(f"suite driver {name} not found")
    end = text.index("\nEnd Sub\n", start)
    return start, end


def apply(text):
    """Instrument an uninstrumented module; the exact inverse of strip()."""
    if instrumented(text):
        raise TimingError("the module is already instrumented")
    for name in DRIVERS:
        start, end = _driver_span(text, name)
        lines = text[start:end].split("\n")
        out, started = [], False
        for line in lines:
            match = CALL.match(line)
            if match and not started:
                out.append(f'    TimingStart "{name}"')
                started = True
            out.append(line)
            if match:
                out.append(f'    TimingLap "{match.group(1)}"')
        if not started:
            raise TimingError(f"suite driver {name} calls no test procedure")
        text = text[:start] + "\n".join(out) + text[end:]
    for name in NESTED:
        start, end = _driver_span(text, name)
        out, chained = [], 0
        for line in text[start:end].split("\n"):
            out.append(line)
            match = CALL.match(line)
            if match:
                out.append(f'    TimingLap "{match.group(1)}"')
                chained += 1
        if not chained:
            raise TimingError(f"{name} chains no test procedure")
        text = text[:start] + "\n".join(out) + text[end:]
    if text.count(ANCHOR) != 1:
        raise TimingError("BeginRun anchor not found exactly once")
    return text.replace(ANCHOR, BLOCK + ANCHOR, 1)
