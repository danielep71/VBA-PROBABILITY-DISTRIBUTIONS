"""Fixtures for tools/traffic_alert.py, including the issue it opens passing Issue triage."""

from __future__ import annotations

import csv
import json
import sys
import tempfile
import unittest
from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(ROOT / "tools"))

import check_issues  # noqa: E402
import traffic_alert  # noqa: E402

TRAFFIC_FIELDS = ["snapshot_date", "snapshot_datetime_utc", "views_14d", "unique_views_14d",
                  "clones_14d", "unique_clones_14d", "stars", "forks", "watchers", "open_issues"]
REFERRER_FIELDS = ["snapshot_date", "rank", "referrer", "count", "uniques"]
PATH_FIELDS = ["snapshot_date", "rank", "path", "title", "count", "uniques"]
REPOSITORY = "danielep71/VBA-PROBABILITY-DISTRIBUTIONS"
DUE = "2026-12-18"


def snapshot(day: str, views: int = 57, clones: int = 434, stars: int = 2,
             forks: int = 0) -> dict[str, str]:
    return {"snapshot_date": day, "snapshot_datetime_utc": f"{day}T06:00:00+00:00",
            "views_14d": str(views), "unique_views_14d": "24", "clones_14d": str(clones),
            "unique_clones_14d": "130", "stars": str(stars), "forks": str(forks),
            "watchers": "1", "open_issues": "30"}


class TrafficAlertTest(unittest.TestCase):
    def setUp(self) -> None:
        directory = tempfile.TemporaryDirectory(prefix="traffic-alert-")
        self.addCleanup(directory.cleanup)
        self.data = Path(directory.name)

    def write(self, name: str, fields: list[str], rows: list[dict[str, str]]) -> None:
        with (self.data / name).open("w", newline="", encoding="utf-8") as handle:
            writer = csv.DictWriter(handle, fieldnames=fields)
            writer.writeheader()
            writer.writerows(rows)

    def history(self, previous: dict[str, str], current: dict[str, str],
                referrers: list[dict[str, str]] | None = None) -> None:
        self.write(traffic_alert.TRAFFIC, TRAFFIC_FIELDS, [previous, current])
        self.write(traffic_alert.REFERRERS, REFERRER_FIELDS, referrers or [])
        self.write(traffic_alert.PATHS, PATH_FIELDS, [
            {"snapshot_date": current["snapshot_date"], "rank": "1",
             "path": "/danielep71/VBA-PROBABILITY-DISTRIBUTIONS", "title": "Overview",
             "count": "8", "uniques": "6"}])

    def test_the_2026_11_03_clones_spike_raises_no_alert(self) -> None:
        # The snapshot behind #126: clones 434 -> 1169 from CI checkouts, nothing else.
        self.history(snapshot("2026-11-02", views=57, clones=434),
                     snapshot("2026-11-03", views=66, clones=1169))
        self.assertIsNone(traffic_alert.alert(self.data, DUE))

    def test_views_spike_raises_an_alert_that_still_shows_clones(self) -> None:
        self.history(snapshot("2026-11-02", views=57), snapshot("2026-11-03", views=200, clones=900))
        issue = traffic_alert.alert(self.data, DUE)
        assert issue is not None
        self.assertEqual(issue["title"], "Traffic alert: 2026-11-03")
        self.assertIn("views spike", issue["body"])
        self.assertIn("| Clones | 900 | 434 |", issue["body"])
        self.assertIn("- Target end date: 2026-11-10 (Europe/Rome)", issue["body"])

    def test_views_growth_below_the_threshold_raises_nothing(self) -> None:
        # Doubled, but by less than 100 views.
        self.history(snapshot("2026-11-02", views=40), snapshot("2026-11-03", views=90))
        self.assertIsNone(traffic_alert.alert(self.data, DUE))

    def test_stars_and_forks(self) -> None:
        self.history(snapshot("2026-11-02"), snapshot("2026-11-03", stars=3, forks=1))
        issue = traffic_alert.alert(self.data, DUE)
        assert issue is not None
        self.assertIn("stars increased, forks increased", issue["body"])

    def test_new_referrer_needs_five_views_and_no_earlier_sighting(self) -> None:
        row = {"rank": "1", "count": "6", "uniques": "3"}
        seen = [{**row, "snapshot_date": "2026-11-02", "referrer": "github.com"},
                {**row, "snapshot_date": "2026-11-03", "referrer": "github.com"},
                {**row, "snapshot_date": "2026-11-03", "referrer": "few.example", "count": "4"}]
        self.history(snapshot("2026-11-02"), snapshot("2026-11-03"), seen)
        self.assertIsNone(traffic_alert.alert(self.data, DUE))
        self.history(snapshot("2026-11-02"), snapshot("2026-11-03"),
                     [*seen, {**row, "snapshot_date": "2026-11-03", "referrer": "blog.example"}])
        issue = traffic_alert.alert(self.data, DUE)
        assert issue is not None
        self.assertIn("new referrer", issue["body"])
        self.assertIn("- blog.example: 6 views, 3 unique", issue["body"])

    def test_target_date_stays_before_the_milestone_due_date(self) -> None:
        self.history(snapshot("2026-12-14"), snapshot("2026-12-15", stars=3))
        issue = traffic_alert.alert(self.data, DUE)
        assert issue is not None
        self.assertIn("- Target end date: 2026-12-17 (Europe/Rome)", issue["body"])
        with self.assertRaises(traffic_alert.TrafficAlertError):
            traffic_alert.alert(self.data, "2026-12-15")

    def test_fewer_than_two_snapshots_raise_nothing(self) -> None:
        self.write(traffic_alert.TRAFFIC, TRAFFIC_FIELDS, [snapshot("2026-11-03", stars=9)])
        self.assertIsNone(traffic_alert.alert(self.data, DUE))

    def test_command_line_writes_json_only_when_an_alert_is_due(self) -> None:
        output = self.data / "alert.json"
        self.history(snapshot("2026-11-02"), snapshot("2026-11-03", clones=5000))
        self.assertEqual(traffic_alert.main(["--data", str(self.data), "--milestone-due", DUE,
                                             "--output", str(output)]), 0)
        self.assertFalse(output.exists())
        self.history(snapshot("2026-11-02"), snapshot("2026-11-03", forks=1))
        self.assertEqual(traffic_alert.main(["--data", str(self.data), "--milestone-due", DUE,
                                             "--output", str(output)]), 0)
        self.assertEqual(json.loads(output.read_text(encoding="utf-8"))["title"],
                         "Traffic alert: 2026-11-03")
        self.assertEqual(traffic_alert.main(["--data", str(self.data), "--milestone-due",
                                             "2026-11-01", "--output", str(output)]), 1)

    def test_alert_issue_passes_issue_triage(self) -> None:
        # The workflow adds the form's assignee, P3 + repository and the milestone;
        # with them, the generated body must pass tools/check_issues.py unchanged.
        self.history(snapshot("2026-11-02"), snapshot("2026-11-03", stars=3))
        issue = traffic_alert.alert(self.data, DUE)
        assert issue is not None
        assignee = check_issues.configured_assignee(ROOT)
        item = {"number": 999, "state": "open", "body": issue["body"],
                "labels": [{"name": "P3"}, {"name": "repository"}],
                "assignees": [{"login": assignee}],
                "milestone": {"title": "v1.0.0", "due_on": f"{DUE}T00:00:00Z", "state": "open"}}
        findings = check_issues.audit([check_issues.issue_from_api(item)], {999: "open"},
                                      REPOSITORY, check_issues.catalogue_labels(ROOT),
                                      assignee=assignee)
        self.assertEqual(findings, [])


if __name__ == "__main__":
    unittest.main()
