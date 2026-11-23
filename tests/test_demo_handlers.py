"""The demo workbook's VBA leaves Excel's state as it found it, also on error (#482).

Excel's events are application-wide: a routine that turns them off and fails
before turning them back on leaves every open workbook without events. These
static checks read every routine of src/workbook/*.bas.
"""
import re
import unittest
from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]
WORKBOOK = ROOT / "src" / "workbook"
ROUTINE = re.compile(r"^(?:Public|Private) (?:Sub|Function) (\w+)\(.*?^End (?:Sub|Function)$", re.S | re.M)


def code_lines(body: str) -> list[str]:
    """The lines of a routine without its comments."""
    return [line for line in body.splitlines() if line.strip() and not line.strip().startswith("'")]


def first(lines: list[str], pattern: str) -> int | None:
    for index, line in enumerate(lines):
        if re.search(pattern, line):
            return index
    return None


def routines():
    for path in sorted(WORKBOOK.glob("*.bas")):
        text = path.read_text(encoding="latin-1").replace("\r\n", "\n")
        for match in ROUTINE.finditer(text):
            yield path.name, match.group(1), code_lines(match.group(0))


class HandlerStateTests(unittest.TestCase):
    def test_the_events_state_is_saved_before_the_error_handler_is_set(self):
        checked = 0
        for module, name, lines in routines():
            restore = first(lines, r"Application\.EnableEvents = SavedEvents")
            handler = first(lines, r"On Error GoTo \w+")
            if restore is None or handler is None:
                continue
            save = first(lines, r"SavedEvents = Application\.EnableEvents")
            with self.subTest(routine=f"{module}:{name}"):
                self.assertIsNotNone(save)
                self.assertLess(save, handler, "the handler would restore a state that was never saved")
            checked += 1
        self.assertGreaterEqual(checked, 20)

    def test_no_routine_forces_events_on(self):
        for module, name, lines in routines():
            with self.subTest(routine=f"{module}:{name}"):
                self.assertIsNone(first(lines, r"Application\.EnableEvents = True"))

    def test_a_routine_that_unprotects_a_sheet_protects_it_again_on_error(self):
        checked = 0
        for module, name, lines in routines():
            unprotect = first(lines, r"\.Unprotect\b")
            if unprotect is None or name.startswith("NavigationBar") or module == "M_STATS_PROBDIST_DEMO.bas":
                continue
            with self.subTest(routine=f"{module}:{name}"):
                handler = first(lines, r"On Error GoTo (\w+)")
                self.assertIsNotNone(handler, "no error handler around the unprotected sheet")
                label = re.search(r"On Error GoTo (\w+)", lines[handler]).group(1)
                start = first(lines, rf"^{label}:")
                self.assertIsNotNone(start)
                self.assertIsNotNone(first(lines[start:], r"\.Protect\b"), "the error path leaves the sheet unprotected")
                # Nothing that can fail runs between the unprotect and the handler.
                self.assertLessEqual(handler, unprotect + 1)
            checked += 1
        self.assertGreaterEqual(checked, 13)


if __name__ == "__main__":
    unittest.main()
