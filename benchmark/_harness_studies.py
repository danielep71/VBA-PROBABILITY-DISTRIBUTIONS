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

Two more batteries leave the pull-request regression, in a class of their own
(SLOW): Test_GammaCF_CapAndXSeam and Test_PublicGamma_Integration. Unlike the
six, they DO exercise production code: the iteration cap and the monotonic
seams of the incomplete-gamma kernel, and the bit-level regression of the
public Gamma and Chi-square surfaces. Together they took 349 of the
regression's 383 seconds (#13 timing). RunSlowProductionBatteries runs them,
with the studies: on every dispatched run and on every export wave, where a
failure fails the job before any grid is exported. No VBA change reaches a
release branch without them, because a VBA change needs an export wave; a
pull-request run alone does not execute them.

The separation is generated, never edited by hand: apply() moves the calls
deterministically and strip() restores the original layout exactly. strip()
also accepts the layout of #300, which moved the six batteries only: the
pinned sources of the closed studies still have it. Apply it before
_harness_timing.apply(), and strip it after _harness_timing.strip().

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
# The slow production batteries and their assertions. The cap-and-seam battery
# is 64 flat assertions; the public integration is 496 cases of 8, 7 invalid
# inputs of 2, 24 inverses of 2 and 2 reach cases of 4 (public_integration.py).
SLOW = {"Test_GammaCF_CapAndXSeam": 64, "Test_PublicGamma_Integration": 4038}
SLOW_ASSERTIONS = sum(SLOW.values())
HOST = "Private Sub Test_CN_GammaSurvival()\n"
CALLS = "".join(f"    {name}\n" for name in BATTERIES)
FOLLOWER = "    Test_PublicGamma_Integration\n"
PRECEDER = "    Test_GammaCF_PrecisionCounterexample\n"
SEAM = "    Test_GammaCF_CapAndXSeam\n"
ANCHOR = "Private Sub BeginRun("

LEGACY_BLOCK = '''' BEGIN study batteries (#300)
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


BLOCK = '''' BEGIN study batteries (#300)
Public Sub Test_STATS_PROBDIST_RunStudies()
'
'==============================================================================
' Test_STATS_PROBDIST_RunStudies
'------------------------------------------------------------------------------
' PURPOSE
'   Runs the batteries kept apart from the pull-request regression: the
'   bit-for-bit fidelity checks of the private arithmetic ports of the #26
'   studies (#300), and the slow production batteries of the Gamma kernel
'   (#13).
'
' BEHAVIOR
'   - Resets the shared counters.
'   - Runs RunStudyBatteries, then RunSlowProductionBatteries.
'   - Prints one consolidated PASS/FAIL summary.
'
' GENERATED
'   By benchmark/_harness_studies.py. Do not edit by hand.
'==============================================================================
'
    BeginRun "STUDY BATTERIES"
    RunStudyBatteries
    RunSlowProductionBatteries
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


Private Sub RunSlowProductionBatteries()
'
'==============================================================================
' RunSlowProductionBatteries
'------------------------------------------------------------------------------
' PURPOSE
'   Calls the slow production batteries: the iteration cap and seams of the
'   incomplete-gamma kernel, and the bit-level regression of the public Gamma
'   and Chi-square surfaces. They test production code, and they run on
'   dispatched and export waves, not on a pull-request run (#13).
'==============================================================================
'
    Debug.Print "== SLOW PRODUCTION BATTERIES: Gamma cap, seams and public integration"
''' + SEAM + FOLLOWER + '''End Sub
' END study batteries (#300)


'''


class StudiesError(ValueError):
    pass


def separated(text):
    return BEGIN in text


def legacy(text):
    """The layout of #300: the six batteries moved, the slow battery in place."""
    return LEGACY_BLOCK in text


def _host_span(text):
    start = text.find(HOST)
    if start < 0:
        raise StudiesError("Test_CN_GammaSurvival not found")
    return start, text.index("\nEnd Sub\n", start)


def apply(text):
    """Move the six study calls and the slow batteries out of Test_CN_GammaSurvival."""
    if separated(text):
        raise StudiesError("the study batteries are already separated")
    start, end = _host_span(text)
    body = text[start:end + 1]
    if body.count(SEAM + PRECEDER + CALLS + FOLLOWER) != 1:
        raise StudiesError("the layout is not cap-and-seam, precision counterexample, the six "
                           "battery calls, public integration")
    text = text[:start] + body.replace(SEAM + PRECEDER + CALLS + FOLLOWER, PRECEDER, 1) + text[end + 1:]
    if text.count(ANCHOR) != 1:
        raise StudiesError("BeginRun anchor not found exactly once")
    return text.replace(ANCHOR, BLOCK + ANCHOR, 1)


def strip(text):
    """The original layout: the block removed and the calls back in place."""
    if not separated(text):
        return text
    if legacy(text):
        if text.count(LEGACY_BLOCK) != 1:
            raise StudiesError("the study block differs from the generator's")
        text = text.replace(LEGACY_BLOCK, "", 1)
        start, end = _host_span(text)
        body = text[start:end + 1]
        if body.count(FOLLOWER) != 1 or CALLS in body:
            raise StudiesError("Test_CN_GammaSurvival does not have the separated layout")
        return text[:start] + body.replace(FOLLOWER, CALLS + FOLLOWER, 1) + text[end + 1:]
    if text.count(BLOCK) != 1:
        raise StudiesError("the study block differs from the generator's")
    text = text.replace(BLOCK, "", 1)
    start, end = _host_span(text)
    body = text[start:end + 1]
    if FOLLOWER in body or CALLS in body or SEAM in body or body.count(PRECEDER) != 1:
        raise StudiesError("Test_CN_GammaSurvival does not have the separated layout")
    return (text[:start] + body.replace(PRECEDER, SEAM + PRECEDER + CALLS + FOLLOWER, 1)
            + text[end + 1:])


STUDY_POLICY = ".github/excel-study-policy.json"


def live_total(root):
    """Regression plus separated assertions, as the two committed policies declare.
    The study policy must declare exactly the batteries and counts above: the
    six studies and the slow production battery, or, at a commit before the
    slow battery moved (#300 layout), the six studies alone."""
    import json
    from pathlib import Path
    root = Path(root)
    regression = json.loads((root / ".github/excel-evidence-policy.json").read_text())
    study = json.loads((root / STUDY_POLICY).read_text())
    slow = study.get("production_batteries")
    expected = STUDY_ASSERTIONS + (SLOW_ASSERTIONS if slow is not None else 0)
    if (study.get("batteries") != dict(ASSERTIONS) or slow not in (None, dict(SLOW))
            or study.get("expected_assertions") != expected):
        raise StudiesError(f"{STUDY_POLICY} does not declare the separated batteries")
    return regression["expected_assertions"] + study["expected_assertions"]
