"""
Proofs for the separated study batteries of #300 (_harness_studies.py).

  1. The committed test module is the generator's output: the six batteries
     are called only from RunStudyBatteries, Test_CN_GammaSurvival no longer
     calls them, and strip() restores the original contiguous layout.
  2. The six batteries call no production procedure beyond PROB_IsFinite and
     the binary64 series and continued fraction: they test private ports.
  3. The two policies split the 14,888 assertions of run 42766552574 into
     5,987 regression and 8,901 study assertions, as the batteries' own
     fidelity records state.
  4. The CI bridge runs the batteries only with -RunStudies, after a green
     regression, against the study policy, and fails the job on a mismatch;
     the workflow requests it on dispatch and on every export.
  5. A hand-edited block and a second separation are refused.

Run: python3 test_harness_studies.py   (exit 0 = pass)
"""
import json
import os
import re
import sys

HERE = os.path.dirname(os.path.abspath(__file__))
ROOT = os.path.dirname(HERE)
sys.path.insert(0, HERE)
import _harness_studies as S  # noqa: E402
import _harness_timing as T  # noqa: E402

COVERAGE = os.path.join(HERE, "gamma_cf_study", "coverage")
ALLOWED = {"PROB_IsFinite", "PROB_TryGammaSeriesP", "PROB_TryGammaContinuedFractionQ"}


def read(rel):
    with open(os.path.join(ROOT, rel), "rb") as f:
        return f.read().decode("utf-8").replace("\r\n", "\n")


def main():
    fails = []

    def check(condition, message):
        if not condition:
            fails.append(message)

    module = read("tests/M_STATS_PROBDIST_TEST.bas")
    untimed = T.strip(module)
    original = S.strip(untimed)

    # 1. Generator output and layout.
    check(S.apply(original) == untimed, "the module is not the generator's output")
    procs = {m.group(2): m.group(0) for m in re.finditer(
        r"^(?:Private|Public) (Sub|Function) (\w+)\b.*?^End \1\n", module, re.M | re.S)}
    host = procs["Test_CN_GammaSurvival"]
    driver = procs["RunStudyBatteries"]
    for name in S.BATTERIES:
        check(f"    {name}\n" not in host, f"Test_CN_GammaSurvival still calls {name}")
        check(f"    {name}\n" in driver, f"RunStudyBatteries does not call {name}")
        callers = re.findall(rf"^    {name}$", module, re.M)
        check(len(callers) == 1, f"{name} is called {len(callers)} times, not once")
    check(S.CALLS + S.FOLLOWER in original, "strip() does not restore the contiguous calls")

    # 2. The batteries test private ports, not production.
    src = "".join(read(f"src/{n}") for n in sorted(os.listdir(os.path.join(ROOT, "src")))
                  if n.endswith(".bas"))
    production = set(re.findall(r"^Public (?:Function|Sub) (\w+)", src, re.M))

    def closure(name, seen):
        if name in seen or name not in procs:
            return seen
        seen.add(name)
        for token in set(re.findall(r"\b([A-Za-z_]\w*)\b", procs[name])):
            closure(token, seen)
        return seen

    for name in S.BATTERIES:
        used = set()
        for proc in closure(name, set()):
            used |= set(re.findall(r"\b(PROB_\w+|K_STATS_\w+)\b", procs[proc])) & production
        check(used <= ALLOWED, f"{name} calls production code: {sorted(used - ALLOWED)}")

    # 3. The policies and the batteries' records agree.
    regression = json.loads(read(".github/excel-evidence-policy.json"))
    study = json.loads(read(S.STUDY_POLICY))
    check(regression["expected_assertions"] == 5987, "the regression policy is not 5,987")
    check(S.live_total(ROOT) == 14888, "regression plus studies is not run 42766552574's 14,888")
    check(study["entry_point"] == "Test_STATS_PROBDIST_RunStudies_CI", "wrong study entry point")
    records = {"Test_ScaledProduct_Fidelity": "scaled_product_fidelity.json",
               "Test_ScaledDD_Fidelity": "scaled_dd_fidelity.json",
               "Test_ScaledGamma_Fidelity": "scaled_gamma_fidelity.json",
               "Test_ExtendedGamma_Fidelity": "extended_gamma_fidelity.json",
               "Test_WideGamma_Fidelity": "wide_gamma_fidelity.json"}
    for name, path in records.items():
        with open(os.path.join(COVERAGE, path)) as f:
            added = json.load(f)["additional_assertions"]
        check(S.ASSERTIONS[name] == added, f"{name}: {S.ASSERTIONS[name]} != {path} {added}")
    with open(os.path.join(COVERAGE, "install_v2_fidelity.py")) as f:
        installer = f.read()
    check('"expected_assertions": 1947' in installer and '"expected_assertions": 2447' in installer
          and S.ASSERTIONS["Test_PrecisionV2_Fidelity"] == 2447 - 1947,
          "Precision v2's 500 assertions are not the installer's 1,947 -> 2,447")

    # 4. The bridge and the workflow.
    bridge = read("ci/Run-ExcelVbaTests.ps1")
    check("[switch]$RunStudies" in bridge and "if ($RunStudies) {" in bridge,
          "the runner has no -RunStudies switch")
    check(bridge.index('$stages.regression.status = "PASS"') < bridge.index("if ($RunStudies) {"),
          "the studies do not run after a green regression")
    check("excel-study-policy.json" in bridge and "$exitCode = 1" in bridge.split("if ($RunStudies) {")[1],
          "a study failure does not fail the job against the study policy")
    run_studies = re.search(r"Public Function Test_STATS_PROBDIST_RunStudies_CI\(\).*?End Function",
                            bridge, re.S)
    check(run_studies is not None and "RunStudyBatteries" in run_studies.group(0),
          "the study entry point does not run RunStudyBatteries")
    workflow = read(".github/workflows/excel-vba-regression.yml")
    check("run_studies:" in workflow and "'-RunStudies'" in workflow
          and "github.event_name == 'workflow_dispatch'" in workflow
          and "inputs.export_main" in workflow.split("'-RunStudies'")[0].splitlines()[-1],
          "the workflow does not request the studies on dispatch and exports")

    # 5. Drift is refused.
    try:
        S.strip(untimed.replace("RunStudyBatteries\n    EndRun", "RunStudyBatteries\n\n    EndRun", 1))
        check(False, "a hand-edited study block was accepted")
    except S.StudiesError:
        pass
    try:
        S.apply(untimed)
        check(False, "a second separation was accepted")
    except S.StudiesError:
        pass

    # 6. The retained log of run 42775390763 is the certified one and shows
    #    both parts passing with the declared counts.
    import hashlib
    log = os.path.join(HERE, "harness_timing", "run_42775390763_log.txt")
    with open(log, "rb") as f:
        raw = f.read()
    crlf = raw.replace(b"\r\n", b"\n").replace(b"\n", b"\r\n")
    check(hashlib.sha256(crlf).hexdigest()
          == "16ce9124b779fa507f54ac5823489196357e319974980c6e4bffab290b08b3a7",
          "the retained log is not the one run 42775390763's record binds")
    text = raw.decode("utf-8-sig")
    check("VBA result: TOTAL=5987;PASS=5987;FAIL=0" in text
          and "STUDIES: PASS (8901/8901 study assertions)" in text,
          "the retained log does not show both parts passing")

    if fails:
        print("FAIL: #300 study batteries")
        for f in fails:
            print("  - " + f)
        return 1
    print(f"PASS: #300 study batteries: {len(S.BATTERIES)} private-port batteries "
          f"({S.STUDY_ASSERTIONS} assertions) run apart from the 5,987-assertion regression")
    return 0


if __name__ == "__main__":
    sys.exit(main())
