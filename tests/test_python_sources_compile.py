"""Every tracked Python file compiles without a warning.

Python reports an invalid escape sequence such as "\\[" in an ordinary string
as a SyntaxWarning since 3.12 and will make it an error. Compiling with
warnings as errors finds it today, on every supported version.

Run: python3 -m unittest tests.test_python_sources_compile
"""
import subprocess
import unittest
import warnings
from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]


def problems(source, name):
    with warnings.catch_warnings():
        warnings.simplefilter("error")
        try:
            compile(source, name, "exec")
        except (SyntaxError, SyntaxWarning, DeprecationWarning) as error:
            return f"{name}: {type(error).__name__}: {error}"
    return None


class PythonSourcesCompile(unittest.TestCase):
    def test_every_tracked_file_compiles_without_a_warning(self):
        listed = subprocess.run(["git", "-C", str(ROOT), "ls-files", "-z", "*.py"],
                                check=True, capture_output=True).stdout
        paths = [p for p in listed.decode("utf-8").split("\0") if p]
        self.assertGreater(len(paths), 100)
        found = []
        for rel in paths:
            source = (ROOT / rel).read_text(encoding="utf-8")
            found.append(problems(source, rel))
        self.assertEqual([f for f in found if f], [])

    def test_an_invalid_escape_is_reported(self):
        self.assertIn("invalid escape sequence", problems('"""a \\[ b"""\n', "fixture.py"))
        self.assertIsNone(problems('r"""a \\[ b"""\n', "fixture.py"))


if __name__ == "__main__":
    unittest.main()
