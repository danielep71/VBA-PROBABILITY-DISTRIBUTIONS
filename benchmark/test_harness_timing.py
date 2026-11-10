"""
Proofs for the harness timing of #298 (_harness_timing.py).

  1. The committed test module is exactly apply(strip(module)): five
     TimingStart lines, one TimingLap after every test procedure the five
     suite drivers call, and the generated block before BeginRun.
  2. strip() removes everything apply() adds, and nothing else.
  3. A hand-edited timing block, a driver call without its lap, and a
     second instrumentation are refused or detected.
  4. The CI bridge resets the record before the run, reads it after, and
     times the Gamma kernel only after a complete green regression.

Run: python3 test_harness_timing.py   (exit 0 = pass)
"""
import os
import re
import sys

HERE = os.path.dirname(os.path.abspath(__file__))
ROOT = os.path.dirname(HERE)
sys.path.insert(0, HERE)
import _harness_timing as H  # noqa: E402


def read(rel):
    with open(os.path.join(ROOT, rel), "rb") as f:
        return f.read().decode("utf-8").replace("\r\n", "\n")


def main():
    fails = []

    def check(condition, message):
        if not condition:
            fails.append(message)

    module = read("tests/M_STATS_PROBDIST_TEST.bas")
    bare = H.strip(module)

    # 1. The committed module is the generator's output.
    check(H.apply(bare) == module, "the committed module is not apply(strip(module))")
    check(module.count('    TimingStart "') == 5, "expected one TimingStart per suite driver")
    calls = laps = 0
    for name in H.DRIVERS:
        start, end = H._driver_span(module, name)
        body = module[start:end].split("\n")
        for i, line in enumerate(body):
            m = H.CALL.match(line)
            if m:
                calls += 1
                ok = i + 1 < len(body) and body[i + 1] == f'    TimingLap "{m.group(1)}"'
                laps += ok
                check(ok, f"{name}: {m.group(1)} has no lap")
    check(calls == laps and calls > 0, f"{laps} laps for {calls} calls")
    check(module.count(H.BLOCK) == 1 and module.index(H.BLOCK) < module.index(H.ANCHOR),
          "the timing block is not once, before BeginRun")

    # 2. strip() removes exactly the instrumentation.
    check(not H.instrumented(bare), "strip() left instrumentation behind")
    added = len(module.splitlines()) - len(bare.splitlines())
    check(added == 5 + calls + len(H.BLOCK.splitlines()),
          f"strip() removed {added} lines, not the instrumentation alone")

    # 3. Drift is refused or detected.
    try:
        H.strip(module.replace("Static Record           As String", "Static Record As String", 1))
        check(False, "a hand-edited timing block was accepted")
    except H.TimingError:
        pass
    try:
        H.apply(module)
        check(False, "a second instrumentation was accepted")
    except H.TimingError:
        pass
    dropped = module.replace('    TimingLap "Test_Core_Log1p"\n', "", 1)
    check(H.apply(H.strip(dropped)) != dropped, "a driver call without its lap is not detected")

    # 4. The CI bridge uses the record and times the kernel after a green run.
    bridge = read("ci/Run-ExcelVbaTests.ps1")
    run_all = re.search(r"Public Function Test_STATS_PROBDIST_RunAll_CI\(\).*?End Function", bridge, re.S)
    check(run_all is not None and run_all.group(0).index('HarnessTiming "reset"')
          < run_all.group(0).index("BeginRun"), "the bridge does not reset the timing before the run")
    check('HarnessTiming("read")' in bridge and 'Write-CiLog "TIMING $line"' in bridge,
          "the bridge does not log the timing record")
    passed = bridge.index('$stages.regression.status = "PASS"')
    check(passed < bridge.index("Test_STATS_PROBDIST_GammaCost_CI\"") and
          'Write-CiLog "COST $line"' in bridge,
          "the cost benchmark does not run after a green regression")

    if fails:
        print("FAIL: #298 harness timing")
        for f in fails:
            print("  - " + f)
        return 1
    print(f"PASS: #298 harness timing: {calls} test procedures timed in 5 drivers; "
          "the module is the generator's output and the CI bridge logs it")
    return 0


if __name__ == "__main__":
    sys.exit(main())
