"""
The study batteries of the VBA harness, separated from the regression (#300).

Six batteries in tests/M_STATS_PROBDIST_TEST.bas compare private copies of a
study's arithmetic, held in the test module, with that study's frozen Python
model bit for bit. They do not exercise production code, and each study's
certified Excel runs are archived with it. They took most of the suite's ten
minutes (#298), so they leave the regression that every pull request runs:

  - Test_CN_GammaSurvival no longer calls them;
  - Test_STATS_PROBDIST_RunStudies runs them, through the timed driver
    RunStudyBatteries. CI runs it on dispatched and export waves, against
    .github/excel-study-policy.json.

The separation is generated, never edited by hand: apply() moves the six calls
deterministically and strip() restores the original layout exactly. Studies
that compare the test module with frozen versions call strip() first. Apply it
before _harness_timing.apply(), and strip it after _harness_timing.strip().

Both functions work on LF text.
"""

BEGIN = "' BEGIN study batteries (#300)"
END = "' END study batteries (#300)"
BATTERIES = ("Test_PrecisionV2_Fidelity", "Test_ScaledProduct_Fidelity",
             "Test_ScaledDD_Fidelity", "Test_ScaledGamma_Fidelity",
             "Test_ExtendedGamma_Fidelity", "Test_WideGamma_Fidelity")
# Assertions each battery adds, from its generator's fidelity record
# (benchmark/gamma_cf_study/coverage/*_fidelity.json and install_v2_fidelity.py).
ASSERTIONS = {"Test_PrecisionV2_Fidelity": 500, "Test_ScaledProduct_Fidelity": 4657,
              "Test_ScaledDD_Fidelity": 611, "Test_ScaledGamma_Fidelity": 475,
              "Test_ExtendedGamma_Fidelity": 1772, "Test_WideGamma_Fidelity": 886}
STUDY_ASSERTIONS = sum(ASSERTIONS.values())
HOST = "Private Sub Test_CN_GammaSurvival()\n"
CALLS = "".join(f"    {name}\n" for name in BATTERIES)
FOLLOWER = "    Test_PublicGamma_Integration\n"
ANCHOR = "Private Sub BeginRun("

BLOCK = '''' BEGIN study batteries (#300)
Public Sub Test_STATS_PROBDIST_RunStudies()
'
'==============================================================================
' Test_STATS_PROBDIST_RunStudies
'------------------------------------------------------------------------------
' PURPOSE
'   Runs the study batteries: the bit-for-bit fidelity checks of the private
'   arithmetic ports of the #26 studies, kept apart from the regression (#300).
'
' BEHAVIOR
'   - Resets the shared counters.
'   - Runs RunStudyBatteries.
'   - Prints one consolidated PASS/FAIL summary.
'
' GENERATED
'   By benchmark/_harness_studies.py. Do not edit by hand.
'==============================================================================
'
    BeginRun "STUDY BATTERIES"
    RunStudyBatteries
    EndRun
End Sub


Private Sub RunStudyBatteries()
'
'==============================================================================
' RunStudyBatteries
'------------------------------------------------------------------------------
' PURPOSE
'   Calls the six study batteries that compare private ports of the #26 study
'   arithmetic with their frozen Python models (#300).
'==============================================================================
'
    Debug.Print "== STUDY BATTERIES: private arithmetic ports of the #26 studies"
''' + CALLS + '''End Sub
' END study batteries (#300)


'''


class StudiesError(ValueError):
    pass


def separated(text):
    return BEGIN in text


def _host_span(text):
    start = text.find(HOST)
    if start < 0:
        raise StudiesError("Test_CN_GammaSurvival not found")
    return start, text.index("\nEnd Sub\n", start)


def apply(text):
    """Move the six battery calls out of Test_CN_GammaSurvival."""
    if separated(text):
        raise StudiesError("the study batteries are already separated")
    start, end = _host_span(text)
    body = text[start:end + 1]
    if body.count(CALLS + FOLLOWER) != 1:
        raise StudiesError("the six battery calls are not contiguous before the public integration")
    text = text[:start] + body.replace(CALLS, "", 1) + text[end + 1:]
    if text.count(ANCHOR) != 1:
        raise StudiesError("BeginRun anchor not found exactly once")
    return text.replace(ANCHOR, BLOCK + ANCHOR, 1)


def strip(text):
    """The original layout: the block removed and the calls back in place."""
    if not separated(text):
        return text
    if text.count(BLOCK) != 1:
        raise StudiesError("the study block differs from the generator's")
    text = text.replace(BLOCK, "", 1)
    start, end = _host_span(text)
    body = text[start:end + 1]
    if body.count(FOLLOWER) != 1 or CALLS in body:
        raise StudiesError("Test_CN_GammaSurvival does not have the separated layout")
    return text[:start] + body.replace(FOLLOWER, CALLS + FOLLOWER, 1) + text[end + 1:]


STUDY_POLICY = ".github/excel-study-policy.json"


def live_total(root):
    """Regression plus study assertions, as the two committed policies declare.
    The study policy must declare exactly the batteries and counts above."""
    import json
    from pathlib import Path
    root = Path(root)
    regression = json.loads((root / ".github/excel-evidence-policy.json").read_text())
    study = json.loads((root / STUDY_POLICY).read_text())
    if study.get("batteries") != dict(ASSERTIONS) or study.get("expected_assertions") != STUDY_ASSERTIONS:
        raise StudiesError(f"{STUDY_POLICY} does not declare the separated batteries")
    return regression["expected_assertions"] + study["expected_assertions"]
