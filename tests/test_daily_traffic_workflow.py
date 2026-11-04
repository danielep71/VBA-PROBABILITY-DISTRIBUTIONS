"""Credential wiring of the daily traffic export (#142).

The traffic API needs ``Administration: read``. The workflow mints a
short-lived GitHub App installation token for it, scoped to this repository
and to that one permission, and gives it only to the fetch step. These checks
read the workflow as text, so they need no YAML parser, and each rule has a
negative fixture.
"""

from __future__ import annotations

import re
import unittest
from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]
WORKFLOW = ROOT / ".github" / "workflows" / "daily-traffic.yml"

APP_TOKEN_USE = re.compile(
    r"^\s+uses: actions/create-github-app-token@(?P<ref>\S+)(?: # (?P<tag>\S+))?$",
    re.MULTILINE,
)
TOKEN_OUTPUT = "${{ steps.traffic-token.outputs.token }}"


def _step(text: str, name: str) -> str:
    """Return the text of the step called ``name``, up to the next step or job."""
    match = re.search(
        rf"^(?P<indent>\s+)- name: {re.escape(name)}\n(?P<body>.*?)(?=^(?P=indent)- name: |^  \S|\Z)",
        text,
        re.MULTILINE | re.DOTALL,
    )
    return match.group("body") if match else ""


def problems(text: str) -> list[str]:
    """List every way ``text`` breaks the traffic credential rules."""
    found: list[str] = []
    if "secrets.TRAFFIC_TOKEN" in text:
        found.append("reads the retired TRAFFIC_TOKEN personal access token")

    uses = list(APP_TOKEN_USE.finditer(text))
    if len(uses) != 1:
        found.append(f"expected one create-github-app-token step, found {len(uses)}")
    elif not re.fullmatch(r"[0-9a-f]{40}", uses[0].group("ref")):
        found.append("create-github-app-token is not pinned to a full commit SHA")
    elif not uses[0].group("tag"):
        found.append("create-github-app-token pin has no version comment")

    mint = _step(text, "Mint traffic API token")
    if not mint:
        found.append("no 'Mint traffic API token' step")
    else:
        if "id: traffic-token" not in mint:
            found.append("the mint step has no id traffic-token")
        if "secrets.TRAFFIC_APP_PRIVATE_KEY" not in mint:
            found.append("the mint step does not read TRAFFIC_APP_PRIVATE_KEY")
        if "TRAFFIC_APP_CLIENT_ID" not in mint:
            found.append("the mint step does not read TRAFFIC_APP_CLIENT_ID")
        if "repositories: ${{ github.event.repository.name }}" not in mint:
            found.append("the token is not limited to this repository")
        grants = re.findall(r"^\s+(permission-[a-z-]+): (\S+)$", mint, re.MULTILINE)
        if grants != [("permission-administration", "read")]:
            found.append(
                f"the token must grant administration: read only, got {grants}"
            )

    fetch = _step(text, "Fetch GitHub traffic API data")
    if TOKEN_OUTPUT not in fetch:
        found.append("the fetch step does not use the minted token")
    if text.count(TOKEN_OUTPUT) != 1:
        found.append("the minted token must be used exactly once, by the fetch step")

    if text.count("environment: analytics") != 1:
        found.append(
            "the analytics environment must be declared by the traffic job only"
        )
    alert = text[text.index("\n  alert:") :] if "\n  alert:" in text else ""
    if not alert:
        found.append("no alert job")
    elif (
        re.search(r"^\s+environment:", alert, re.MULTILINE)
        or "traffic-token" in alert
        or "TRAFFIC_APP" in alert
    ):
        found.append("the alert job can reach the analytics credential")
    return found


class DailyTrafficCredentialTests(unittest.TestCase):
    def setUp(self) -> None:
        self.text = WORKFLOW.read_text(encoding="utf-8")

    def assertBroken(self, text: str, fragment: str) -> None:
        found = problems(text)
        self.assertTrue(
            any(fragment in item for item in found),
            f"expected a problem containing {fragment!r}, got {found}",
        )

    def test_the_workflow_follows_every_rule(self) -> None:
        self.assertEqual(problems(self.text), [])

    def test_the_retired_token_is_rejected(self) -> None:
        self.assertBroken(
            self.text.replace(TOKEN_OUTPUT, "${{ secrets.TRAFFIC_TOKEN }}", 1),
            "TRAFFIC_TOKEN",
        )

    def test_a_tag_pin_is_rejected(self) -> None:
        mutated = re.sub(
            r"create-github-app-token@[0-9a-f]{40}",
            "create-github-app-token@v3",
            self.text,
        )
        self.assertBroken(mutated, "full commit SHA")

    def test_an_extra_permission_is_rejected(self) -> None:
        mutated = self.text.replace(
            "          permission-administration: read\n",
            "          permission-administration: read\n          permission-contents: write\n",
        )
        self.assertBroken(mutated, "administration: read only")

    def test_a_write_permission_is_rejected(self) -> None:
        mutated = self.text.replace(
            "permission-administration: read", "permission-administration: write"
        )
        self.assertBroken(mutated, "administration: read only")

    def test_an_unscoped_token_is_rejected(self) -> None:
        mutated = self.text.replace(
            "          repositories: ${{ github.event.repository.name }}\n", ""
        )
        self.assertBroken(mutated, "this repository")

    def test_the_token_must_stay_in_the_fetch_step(self) -> None:
        mutated = self.text.replace(
            "\n  alert:", f"\n  alert:\n    # leaked: {TOKEN_OUTPUT}", 1
        )
        self.assertBroken(mutated, "exactly once")

    def test_the_alert_job_must_not_declare_the_environment(self) -> None:
        mutated = self.text.replace(
            "\n  alert:", "\n  alert:\n    environment: analytics", 1
        )
        self.assertBroken(mutated, "analytics environment")


if __name__ == "__main__":
    unittest.main()
