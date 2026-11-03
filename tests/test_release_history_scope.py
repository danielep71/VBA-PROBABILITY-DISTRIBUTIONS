"""The release history policy does not apply to this repository.

`tools/check_release_semantics.py` inspects the commit history since the
previous release tag only when `.github/repository-profile.json` declares the
template repository itself (`"mode": "template"`). This repository's profile
records only its name, so the history policy, which reports any multi-parent
commit without a reviewed exception, never runs here. That matters because the
`Protect main` ruleset allows merge commits only: every pull request reaches
`main` as a merge commit.

These tests pin that scope. They use the checker's own history fixtures, so
they also show that the same history would be rejected if the template policy
were enabled by mistake.
"""

from __future__ import annotations

import importlib.util
import json
import shutil
import sys
import tempfile
import unittest
from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(ROOT / "tools"))  # the checker imports _gatelib from tools/
SPEC = importlib.util.spec_from_file_location(
    "check_release_semantics", ROOT / "tools" / "check_release_semantics.py")
assert SPEC is not None and SPEC.loader is not None
semantics = importlib.util.module_from_spec(SPEC)
sys.modules[SPEC.name] = semantics
SPEC.loader.exec_module(semantics)

PROFILE = ROOT / semantics.CONFIG_PATH


def codes(report: dict[str, object]) -> set[str]:
    findings = report["findings"]
    assert isinstance(findings, list)
    return {str(finding["code"]) for finding in findings}


@unittest.skipIf(shutil.which("git") is None, "git is not installed")
class ReleaseHistoryScopeTest(unittest.TestCase):
    def setUp(self) -> None:
        directory = tempfile.TemporaryDirectory(prefix="release-history-scope-")
        self.addCleanup(directory.cleanup)
        self.root = Path(directory.name)

    def merge_history(self) -> None:
        """A release tag, then an unreviewed pull request merge commit after it."""
        report = semantics._history_merge_case(self.root, approved=False)
        self.assertIn("unapproved-merge-commit", codes(report))

    def use_repository_profile(self) -> dict[str, object]:
        """Commit this repository's profile, VERSION and CHANGELOG on top of the merge."""
        for name in (semantics.CONFIG_PATH, "VERSION", "CHANGELOG.md"):
            (self.root / name).write_bytes((ROOT / name).read_bytes())
        semantics._fixture_git(self.root, "add", "--all")
        semantics._fixture_git(self.root, "commit", "-m", "Use this repository's profile")
        config = json.loads(PROFILE.read_text(encoding="utf-8"))
        self.assertIsInstance(config, dict)
        return config

    def test_repository_profile_does_not_enable_the_template_history_policy(self) -> None:
        config = json.loads(PROFILE.read_text(encoding="utf-8"))
        self.assertNotEqual(config.get("mode"), "template",
                            "the template history policy would report every merge commit")
        self.assertFalse((ROOT / semantics.HISTORY_POLICY_PATH).exists())

    def test_pull_request_merge_commit_passes_with_this_repository_profile(self) -> None:
        self.merge_history()
        config = self.use_repository_profile()
        history = semantics._release_history_report(self.root, config)
        self.assertFalse(history["applicable"])
        self.assertEqual(history["findings"], [])
        report = semantics.run_check(self.root)
        self.assertEqual(report["status"], "pass", report["findings"])

    def test_template_mode_still_rejects_the_same_merge(self) -> None:
        # The checker itself is unchanged: with the template profile, the same
        # unreviewed merge commit after a release tag is a finding.
        self.merge_history()
        config = json.loads((self.root / semantics.CONFIG_PATH).read_text(encoding="utf-8"))
        self.assertEqual(config.get("mode"), "template")
        history = semantics._release_history_report(self.root, config)
        self.assertTrue(history["applicable"])
        self.assertIn("unapproved-merge-commit", codes(history))


if __name__ == "__main__":
    unittest.main()
