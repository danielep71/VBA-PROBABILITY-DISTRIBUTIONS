#!/usr/bin/env python3
"""Decide whether the daily traffic snapshot needs an alert issue, and write it.

The *Daily traffic export* workflow keeps a longer traffic history on the
`traffic-history` branch (data/traffic.csv, data/referrers.csv and
data/popular_paths.csv). Its alert job runs this script on that data, without
the analytics environment. The script compares the latest snapshot with the
one before. An alert is due on any of these:

- a views spike: the 14-day views at least double and grow by 100 or more;
- more stars or more forks;
- a referrer never seen before that brings at least 5 views.

Clones never raise an alert. Every workflow checkout counts as a clone, so a
busy day of pull requests would otherwise raise one on the repository's own
CI. Clones are still shown in the alert.

The issue it writes passes Issue triage (tools/check_issues.py): the workflow
adds the assignee, the P3 and repository labels and the milestone, and the
body carries a target end date before that milestone's due date and a
dependency section.

Usage:
    python3 tools/traffic_alert.py --data history/data --milestone-due 2026-12-18 \\
        --output alert.json

When an alert is due, the output file holds {"title": ..., "body": ...}.
Otherwise no file is written. The exit code is 0 in both cases.
"""

from __future__ import annotations

import argparse
import csv
import json
import sys
from collections.abc import Mapping, Sequence
from datetime import date, timedelta
from pathlib import Path

TRAFFIC = "traffic.csv"
REFERRERS = "referrers.csv"
PATHS = "popular_paths.csv"

# A 14-day total that at least doubles and grows by this much is a spike.
SPIKE_GROWTH = {"views_14d": 100}
NEW_REFERRER_VIEWS = 5
TOP_ROWS = 10
# The alert's target end date: this many days after the snapshot, and always
# before the milestone's due date, as Issue triage requires.
ALERT_DAYS = 7

Row = Mapping[str, str]


class TrafficAlertError(ValueError):
    """The data or the milestone cannot produce a valid alert."""


def read_rows(path: Path) -> list[dict[str, str]]:
    if not path.is_file():
        return []
    with path.open(newline="", encoding="utf-8") as handle:
        return [dict(row) for row in csv.DictReader(handle)]


def count(row: Row, key: str) -> int:
    value = (row.get(key) or "").strip()
    try:
        return int(value)
    except ValueError as error:
        raise TrafficAlertError(f"{key} is not a whole number: {value!r}") from error


def top(rows: Sequence[Row], limit: int = TOP_ROWS) -> list[Row]:
    return sorted(rows, key=lambda row: count(row, "count"), reverse=True)[:limit]


def listed(rows: Sequence[Row], label: str, empty: str) -> list[str]:
    return [f"- {row[label]}: {row['count']} views, {row['uniques']} unique"
            for row in rows] or [f"- {empty}"]


def reasons_for(previous: Row, current: Row, new_referrers: Sequence[Row]) -> list[str]:
    reasons = []
    for key, growth in SPIKE_GROWTH.items():
        before, now = count(previous, key), count(current, key)
        if now >= max(before * 2, before + growth):
            reasons.append(f"{key.removesuffix('_14d').capitalize()} spike")
    for key in ("stars", "forks"):
        if count(current, key) > count(previous, key):
            reasons.append(f"{key.capitalize()} increased")
    if new_referrers:
        reasons.append("New referrer")
    return reasons


def alert(data: Path, milestone_due: str) -> dict[str, str] | None:
    """The alert issue for the latest snapshot, or None when nothing stands out."""
    snapshots = sorted((row for row in read_rows(data / TRAFFIC) if row.get("snapshot_date")),
                       key=lambda row: row["snapshot_date"])
    if len(snapshots) < 2:
        return None
    previous, current = snapshots[-2], snapshots[-1]
    day = current["snapshot_date"]

    referrers = read_rows(data / REFERRERS)
    known = {row["referrer"] for row in referrers if row["snapshot_date"] < day}
    today = [row for row in referrers if row["snapshot_date"] == day]
    new = [row for row in today if row["referrer"] and row["referrer"] not in known
           and count(row, "count") >= NEW_REFERRER_VIEWS]
    reasons = reasons_for(previous, current, new)
    if not reasons:
        return None

    try:
        snapshot_day, due = date.fromisoformat(day), date.fromisoformat(milestone_due)
    except ValueError as error:
        raise TrafficAlertError(f"not an ISO date: {error}") from error
    if due <= snapshot_day:
        raise TrafficAlertError(f"the milestone is due {milestone_due}, not after the snapshot "
                                f"date {day}; open a milestone with a later due date")
    target = min(snapshot_day + timedelta(days=ALERT_DAYS), due - timedelta(days=1)).isoformat()

    def compared(key: str, name: str) -> str:
        return f"| {name} | {current.get(key) or 'not recorded'} | {previous.get(key) or 'not recorded'} |"

    paths = [row for row in read_rows(data / PATHS) if row["snapshot_date"] == day]
    body = [
        "## Objective", "",
        f"Find out what drove the traffic recorded on {day}: " + ", ".join(reasons).lower() + ".", "",
        "## Context", "",
        "Opened by the *Daily traffic export* workflow. Clones are shown but never raise an "
        "alert: every workflow checkout counts as a clone.", "",
        f"| 14-day snapshot | {day} | {previous['snapshot_date']} |", "| --- | --- | --- |",
        compared("views_14d", "Views"), compared("unique_views_14d", "Unique visitors"),
        compared("clones_14d", "Clones"), compared("unique_clones_14d", "Unique cloners"),
        compared("stars", "Stars"), compared("forks", "Forks"), "",
        f"**New referrers** (at least {NEW_REFERRER_VIEWS} views):", "",
        *listed(top(new), "referrer", "None."), "",
        "**Top referrers:**", "", *listed(top(today), "referrer", "None reported."), "",
        "**Popular paths:**", "", *listed(top(paths), "path", "None reported."), "",
        "## Acceptance criteria", "",
        "- [ ] The sources are checked in **Insights → Traffic**: which referrer (a blog, forum, "
        "newsletter or another project) and which page drew the visits.",
        "- [ ] What was found, and any follow-up issue, is recorded in a Resolution.", "",
        "## Dependencies", "", "None.", "",
        "## Planning", "",
        f"- Planned start date: {day} (UTC collection date)",
        f"- Target end date: {target} (Europe/Rome)",
        f"- Milestone due date: {milestone_due}",
    ]
    return {"title": f"Traffic alert: {day}", "body": "\n".join(body) + "\n"}


def main(argv: Sequence[str] | None = None) -> int:
    parser = argparse.ArgumentParser(description=__doc__.splitlines()[0])
    parser.add_argument("--data", type=Path, required=True,
                        help="the data/ directory of the traffic-history branch")
    parser.add_argument("--milestone-due", required=True,
                        help="due date (YYYY-MM-DD) of the milestone the alert is filed under")
    parser.add_argument("--output", type=Path, required=True,
                        help="where to write the alert as JSON, if one is due")
    args = parser.parse_args(argv)
    try:
        issue = alert(args.data, args.milestone_due)
    except (TrafficAlertError, KeyError) as error:
        print(f"::error::traffic alert: {error}", file=sys.stderr)
        return 1
    if issue is None:
        print("No traffic alert is due.")
        return 0
    args.output.write_text(json.dumps(issue, indent=2) + "\n", encoding="utf-8")
    print(f"Traffic alert due: {issue['title']}")
    return 0


if __name__ == "__main__":
    sys.exit(main())
