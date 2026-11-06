"""Tests for the repository settings tool (tools/setup_repository.py).

Every test runs offline against GitHub REST responses recorded from this
repository on 2026-11-02 (tests/fixtures/repository_settings/), or against
copies of them changed to show one departure from the baseline.
"""

from __future__ import annotations

import argparse
import copy
import io
import json
import re
import sys
import tempfile
import unittest
from contextlib import redirect_stdout
from datetime import datetime, timezone
from pathlib import Path
from typing import Any
from unittest.mock import patch

ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(ROOT / "tools"))

import setup_repository as setup  # noqa: E402

RECORDED = ROOT / "tests" / "fixtures" / "repository_settings" / "recorded-2026-11-02.json"
REPO = "danielep71/VBA-PROBABILITY-DISTRIBUTIONS"
API = f"repos/{REPO}"
NOW = datetime(2026, 11, 2, 14, 0, 0, tzinfo=timezone.utc)
INTEGRATION = 15368


class RecordedAPI:
    """Serves recorded responses and applies writes to its own copy of them."""

    def __init__(self) -> None:
        recorded = json.loads(RECORDED.read_text(encoding="utf-8"))["responses"]
        self.responses: dict[str, dict[str, Any]] = copy.deepcopy(recorded)
        self.writes: list[tuple[str, str, Any]] = []
        self.ignore_writes = False
        self.check_conclusion = "success"

    # Shortcuts to the recorded state.
    @property
    def repo(self) -> dict[str, Any]:
        body: dict[str, Any] = self.responses[API]["body"]
        return body

    def ruleset(self, name: str) -> dict[str, Any]:
        rid = next(r["id"] for r in self.responses[API + "/rulesets"]["body"] if r["name"] == name)
        body: dict[str, Any] = self.responses[f"{API}/rulesets/{rid}"]["body"]
        return body

    def rule(self, name: str, rule_type: str) -> dict[str, Any]:
        return next(r for r in self.ruleset(name)["rules"] if r["type"] == rule_type)

    def remove_ruleset(self, name: str) -> None:
        listed = self.responses[API + "/rulesets"]["body"]
        rid = next(r["id"] for r in listed if r["name"] == name)
        self.responses[API + "/rulesets"]["body"] = [r for r in listed if r["id"] != rid]
        del self.responses[f"{API}/rulesets/{rid}"]

    def set_status(self, path: str, status: int, body: Any = None) -> None:
        self.responses[path] = {"status": status, "body": body}

    def __call__(self, method: str, endpoint: str, payload: Any = None) -> Any:
        path = endpoint.split("?", 1)[0]
        if method == "GET":
            return self.get(path, endpoint)
        self.writes.append((method, endpoint, copy.deepcopy(payload)))
        if not self.ignore_writes:
            self.write(method, path, copy.deepcopy(payload))
        return {}

    def get(self, path: str, endpoint: str) -> Any:
        if re.search(r"[?&]page=([2-9]|\d\d)", endpoint):
            return {"check_runs": []} if path.endswith("/check-runs") else []
        if path.endswith("/check-runs"):
            names = ("Repository integrity (static, no Excel)", "Strict accuracy gate (pure Python, no Excel)")
            return {"check_runs": [{"name": n, "app": {"id": INTEGRATION}, "status": "completed",
                                    "conclusion": self.check_conclusion} for n in names]}
        if path == f"{API}/branches/main":
            return {"commit": {"sha": "0" * 40}}
        response = self.responses.get(path)
        if response is None:
            raise setup.ApiError("GET", endpoint, "404")
        if response["status"] >= 300:
            raise setup.ApiError("GET", endpoint, str(response["status"]))
        return copy.deepcopy(response["body"])

    def write(self, method: str, path: str, payload: Any) -> None:
        analysis = self.repo.setdefault("security_and_analysis", {})
        if path == API + "/rulesets" and method == "POST":
            rid = 90000000 + len(self.writes)
            self.responses[API + "/rulesets"]["body"].append(
                {"id": rid, "name": payload["name"], "source_type": "Repository"})
            self.responses[f"{API}/rulesets/{rid}"] = {"status": 200, "body": dict(payload, id=rid)}
        elif "/rulesets/" in path and method == "PUT":
            rid = int(path.rsplit("/", 1)[1])
            self.responses[path] = {"status": 200, "body": dict(payload, id=rid)}
        elif path == API and method == "PATCH":
            analysis.update(payload.pop("security_and_analysis", {}))
            self.repo.update(payload)
        elif path.endswith("/vulnerability-alerts"):
            self.set_status(path, 204 if method == "PUT" else 404)
        elif path.endswith("/automated-security-fixes"):
            analysis["dependabot_security_updates"] = {"status": "enabled" if method == "PUT" else "disabled"}
        elif path.endswith("/private-vulnerability-reporting"):
            self.set_status(path, 200, {"enabled": method == "PUT"})
        else:
            raise AssertionError(f"unexpected write {method} {path}")


class SetupRepositoryTests(unittest.TestCase):
    def setUp(self) -> None:
        self.api = RecordedAPI()
        self.args = argparse.Namespace(repo=None, baseline=setup.BASELINE, preview=False, apply=False,
                                       verify=False, update_existing=False, evidence=None)

    def run_tool(self, **flags: Any) -> tuple[int, str]:
        for key, value in flags.items():
            setattr(self.args, key, value)
        output = io.StringIO()
        with redirect_stdout(output):
            code = setup.run(self.args, self.api, NOW)
        return code, output.getvalue()

    def verify(self) -> tuple[int, str]:
        return self.run_tool(verify=True)

    # The committed baseline and the recorded state.

    def test_recorded_state_verifies_without_writes(self) -> None:
        code, output = self.verify()
        self.assertEqual(code, 0, output)
        self.assertIn("VERIFIED", output)
        self.assertEqual(self.api.writes, [])

    def test_preview_of_recorded_state_plans_nothing(self) -> None:
        code, output = self.run_tool()
        self.assertEqual(code, 0)
        self.assertNotIn("PLAN", output)
        self.assertEqual(self.api.writes, [])

    def test_baseline_names_the_profile_repository(self) -> None:
        profile = json.loads((ROOT / ".github" / "repository-profile.json").read_text(encoding="utf-8"))
        self.assertEqual(setup.load_baseline(setup.BASELINE)["repository"], profile["repository"])

    def test_required_checks_are_workflow_job_names(self) -> None:
        jobs: set[str] = set()
        for workflow in (ROOT / ".github" / "workflows").glob("*.yml"):
            jobs |= set(re.findall(r"^ {4}name: *(.+?) *$", workflow.read_text(encoding="utf-8"), re.M))
        baseline = setup.load_baseline(setup.BASELINE)
        contexts = {r["name"]: {c["context"] for rule in r["rules"] if rule["type"] == "required_status_checks"
                                for c in rule["parameters"]["required_status_checks"]}
                    for r in baseline["rulesets"]}
        self.assertEqual(contexts["Protect main"], {"Repository integrity (static, no Excel)",
                                                    "Strict accuracy gate (pure Python, no Excel)"})
        for name, required in contexts.items():
            self.assertLessEqual(required, jobs, name)

    def test_main_allows_merge_commits_only(self) -> None:
        main = next(r for r in setup.load_baseline(setup.BASELINE)["rulesets"] if r["name"] == "Protect main")
        pull_request = next(rule for rule in main["rules"] if rule["type"] == "pull_request")
        self.assertEqual(pull_request["parameters"]["allowed_merge_methods"], ["merge"])

    # Drift is reported, never written, outside --apply.

    def test_ruleset_drift_names_the_changed_parameter(self) -> None:
        self.api.rule("Protect main", "pull_request")["parameters"]["allowed_merge_methods"] = ["merge", "squash"]
        code, output = self.verify()
        self.assertEqual(code, 1)
        self.assertIn('DIFF ruleset "Protect main" (22186928) rules.pull_request.allowed_merge_methods: '
                      'observed ["merge", "squash"], baseline ["merge"]', output)
        self.assertEqual(self.api.writes, [])

    def test_removed_required_check_is_drift(self) -> None:
        checks = self.api.rule("Protect main", "required_status_checks")["parameters"]["required_status_checks"]
        del checks[1]
        code, output = self.verify()
        self.assertEqual(code, 1)
        self.assertIn("required_status_checks.required_status_checks", output)

    def test_lax_up_to_date_policy_is_drift(self) -> None:
        params = self.api.rule("Protect main", "required_status_checks")["parameters"]
        params["strict_required_status_checks_policy"] = False
        self.assertEqual(self.verify()[0], 1)

    def test_tag_ruleset_without_update_rule_is_drift(self) -> None:
        tags = self.api.ruleset("Protect release tags")
        tags["rules"] = [r for r in tags["rules"] if r["type"] != "update"]
        code, output = self.verify()
        self.assertEqual(code, 1)
        self.assertIn("rules.update: observed \"<absent>\"", output)

    def test_bypass_actor_and_disabled_enforcement_are_drift(self) -> None:
        for change in ({"bypass_actors": [{"actor_id": 5, "actor_type": "RepositoryRole",
                                            "bypass_mode": "always"}]},
                       {"enforcement": "evaluate"}):
            with self.subTest(change=change):
                self.api = RecordedAPI()
                self.api.ruleset("Protect main").update(change)
                self.assertEqual(self.verify()[0], 1)

    def test_missing_ruleset_is_drift_and_preview_plans_its_creation(self) -> None:
        self.api.remove_ruleset("Protect traffic history")
        self.assertEqual(self.verify()[0], 1)
        code, output = self.run_tool(verify=False)
        self.assertEqual(code, 0)
        self.assertIn(f"PLAN POST {API}/rulesets", output)
        self.assertIn("PREVIEW ONLY", output)
        self.assertEqual(self.api.writes, [])

    def test_merge_method_and_feature_drift(self) -> None:
        for key in ("allow_squash_merge", "allow_auto_merge", "has_wiki", "delete_branch_on_merge"):
            with self.subTest(setting=key):
                self.api = RecordedAPI()
                self.api.repo[key] = not self.api.repo[key]
                code, output = self.verify()
                self.assertEqual(code, 1)
                self.assertIn(f"DIFF setting {key}", output)

    def test_security_feature_drift(self) -> None:
        for key in ("secret_scanning", "secret_scanning_push_protection", "dependabot_security_updates"):
            with self.subTest(feature=key):
                self.api = RecordedAPI()
                self.api.repo["security_and_analysis"][key] = {"status": "disabled"}
                self.assertEqual(self.verify()[0], 1)
        self.api = RecordedAPI()
        self.api.set_status(API + "/private-vulnerability-reporting", 200, {"enabled": False})
        code, output = self.verify()
        self.assertEqual(code, 1)
        self.assertIn("DIFF security private_vulnerability_reporting", output)

    def test_visibility_drift_is_reported_and_never_written(self) -> None:
        self.api.repo["visibility"] = "private"
        code, output = self.verify()
        self.assertEqual(code, 1)
        self.assertIn("DIFF visibility: observed private, baseline public", output)
        with self.assertRaisesRegex(setup.SetupError, "never changes visibility"):
            self.run_tool(verify=False, apply=True)
        self.assertEqual(self.api.writes, [])

    def test_unmanaged_ruleset_is_noted_not_drift(self) -> None:
        self.api.responses[API + "/rulesets"]["body"].append(
            {"id": 7, "name": "Extra", "source_type": "Repository"})
        self.api.set_status(f"{API}/rulesets/7", 200, {"id": 7, "name": "Extra"})
        code, output = self.verify()
        self.assertEqual(code, 0)
        self.assertIn('unmanaged repository ruleset "Extra"', output)

    def test_duplicate_ruleset_names_are_refused(self) -> None:
        self.api.responses[API + "/rulesets"]["body"].append(
            {"id": 7, "name": "Protect main", "source_type": "Repository"})
        with self.assertRaisesRegex(setup.SetupError, "Ambiguous"):
            self.verify()

    # Dependabot alerts: 204 on, 404 off, otherwise inferred or unreadable.

    def test_vulnerability_alerts_states(self) -> None:
        state = setup.security_states(self.api, API, self.api.repo)["vulnerability_alerts"]
        self.assertEqual(state["observed"], True)
        self.assertIn("inferred", state["source"])
        self.api.set_status(API + "/vulnerability-alerts", 204)
        self.assertEqual(self.verify()[0], 0)
        self.api.set_status(API + "/vulnerability-alerts", 404)
        code, output = self.verify()
        self.assertEqual(code, 1)
        self.assertIn("DIFF security vulnerability_alerts: observed false", output)

    def test_unreadable_alerts_without_security_updates_are_unverified(self) -> None:
        del self.api.repo["security_and_analysis"]["dependabot_security_updates"]
        code, output = self.verify()
        self.assertEqual(code, 1)
        self.assertIn("vulnerability_alerts: observed null", output)
        self.assertIn("automated_security_fixes: observed null", output)

    # Paused Dependabot security updates (#206).

    def pause(self, enabled: bool = True, paused: bool = True) -> None:
        self.api.set_status(API + "/automated-security-fixes", 200, {"enabled": enabled, "paused": paused})

    def test_paused_security_updates_are_drift(self) -> None:
        self.pause()
        with tempfile.TemporaryDirectory() as directory:
            target = Path(directory) / "evidence.json"
            code, output = self.run_tool(verify=True, evidence=target)
            record = json.loads(target.read_text(encoding="utf-8"))
        self.assertEqual(code, 1)
        self.assertIn('DIFF security automated_security_fixes: observed "on (paused)", baseline true', output)
        self.assertIn("answering or closing the open Dependabot pull requests resumes them", output)
        row = next(r for r in record["security"] if r["feature"] == "automated_security_fixes")
        self.assertEqual((row["observed"], row["paused"], row["matches"]), ("on (paused)", True, False))

    def test_enabled_security_updates_that_are_not_paused_match(self) -> None:
        self.pause(paused=False)
        self.assertEqual(self.verify()[0], 0)

    def test_apply_cannot_resume_paused_security_updates(self) -> None:
        self.pause()
        with self.assertRaisesRegex(setup.SetupError, "Read-back"):
            self.run_tool(apply=True)
        self.assertIn(("PUT", API + "/automated-security-fixes", None), self.api.writes)

    def test_unreadable_endpoint_leaves_the_paused_state_unknown(self) -> None:
        # The recording: the endpoint answered 403, security_and_analysis says enabled.
        state = setup.security_states(self.api, API, self.api.repo)["automated_security_fixes"]
        self.assertEqual(state["observed"], True)
        self.assertIsNone(state["paused"])
        self.assertIn("paused state unknown", state["source"])
        self.assertIn("HTTP 403", state["source"])

    def test_turning_paused_security_updates_off_deletes_them(self) -> None:
        self.pause()
        baseline = setup.load_baseline(setup.BASELINE)
        baseline["security"]["automated_security_fixes"] = False
        state = setup.observe(self.api, REPO)
        steps = setup.plan(baseline, state, setup.assess(baseline, state), REPO)
        self.assertIn(("DELETE", API + "/automated-security-fixes", None), steps)

    # GitHub Actions settings: read and reported, never written (#208).

    def actions(self, **changes: Any) -> None:
        values = {"actions/permissions": {"enabled": True, "allowed_actions": "all",
                                          "sha_pinning_required": True},
                  "actions/permissions/workflow": {"default_workflow_permissions": "read",
                                                   "can_approve_pull_request_reviews": False},
                  "actions/permissions/fork-pr-contributor-approval": {
                      "approval_policy": "all_external_contributors"},
                  "actions/permissions/artifact-and-log-retention": {"days": 90,
                                                                     "maximum_allowed_days": 400}}
        for endpoint, fields in changes.items():
            values[endpoint].update(fields)
        for endpoint, body in values.items():
            self.api.set_status(f"{API}/{endpoint}", 200, body)

    def test_actions_settings_that_match_verify(self) -> None:
        self.actions()
        with tempfile.TemporaryDirectory() as directory:
            target = Path(directory) / "evidence.json"
            code, output = self.run_tool(verify=True, evidence=target)
            record = json.loads(target.read_text(encoding="utf-8"))
        self.assertEqual(code, 0, output)
        rows = {row["setting"]: row for row in record["actions"]}
        self.assertEqual(rows["artifact_and_log_retention_days"]["observed"], 90)
        self.assertEqual(rows["fork_pr_approval_policy"]["observed"], "all_external_contributors")
        self.assertTrue(all(row["matches"] is True for row in rows.values()))

    def test_actions_drift_is_reported_and_never_written(self) -> None:
        self.actions(**{"actions/permissions/artifact-and-log-retention": {"days": 30},
                        "actions/permissions/workflow": {"default_workflow_permissions": "write"}})
        code, output = self.verify()
        self.assertEqual(code, 1)
        self.assertIn("DIFF actions artifact_and_log_retention_days: observed 30, baseline 90", output)
        self.assertIn('DIFF actions default_workflow_permissions: observed "write", baseline "read"', output)
        with self.assertRaisesRegex(setup.SetupError, "Read-back"):
            self.run_tool(verify=False, apply=True)
        self.assertEqual([w for w in self.api.writes if "/actions/" in w[1]], [])

    def test_unreadable_actions_settings_are_reported_not_fatal(self) -> None:
        # The recording holds no Actions response: every endpoint is unreadable.
        with tempfile.TemporaryDirectory() as directory:
            target = Path(directory) / "evidence.json"
            code, output = self.run_tool(verify=True, evidence=target)
            record = json.loads(target.read_text(encoding="utf-8"))
        self.assertEqual(code, 0, output)
        self.assertIn("NOTE: actions sha_pinning_required not verified: GET actions/permissions "
                      "unreadable (HTTP 404)", output)
        self.assertTrue(record["actions"])
        self.assertTrue(all(row["matches"] is None for row in record["actions"]))

    # Apply writes only with the flag, after its prerequisites, and reads back.

    def drift_everywhere(self) -> None:
        self.api.remove_ruleset("Protect traffic history")
        self.api.rule("Protect main", "pull_request")["parameters"]["allowed_merge_methods"] = ["squash"]
        self.api.repo["has_projects"] = True
        self.api.repo["security_and_analysis"]["secret_scanning_push_protection"] = {"status": "disabled"}
        self.api.set_status(API + "/vulnerability-alerts", 404)
        self.api.repo["security_and_analysis"]["dependabot_security_updates"] = {"status": "disabled"}

    def test_apply_repairs_drift_then_second_run_is_a_noop(self) -> None:
        self.drift_everywhere()
        code, _ = self.run_tool(apply=True, update_existing=True)
        self.assertEqual(code, 0)
        self.assertEqual([(m, e.removeprefix(API)) for m, e, _ in self.api.writes], [
            # Rulesets in baseline order, alerts before security updates, settings last.
            ("PUT", "/rulesets/22186928"), ("POST", "/rulesets"),
            ("PUT", "/vulnerability-alerts"), ("PUT", "/automated-security-fixes"),
            ("PATCH", ""), ("PATCH", "")])
        self.assertEqual(self.api.writes[4][2], {"security_and_analysis": {
            "secret_scanning_push_protection": {"status": "enabled"}}})
        self.assertEqual(self.api.writes[5][2], {"has_projects": False})
        self.api.writes.clear()
        self.assertEqual(self.run_tool()[0], 0)
        self.assertEqual(self.api.writes, [])

    def test_replacing_a_ruleset_needs_update_existing(self) -> None:
        self.api.rule("Protect main", "pull_request")["parameters"]["allowed_merge_methods"] = ["squash"]
        with self.assertRaisesRegex(setup.SetupError, "--update-existing"):
            self.run_tool(apply=True)
        self.assertEqual(self.api.writes, [])

    def test_apply_refuses_without_admin_or_with_red_checks(self) -> None:
        self.api.remove_ruleset("Protect traffic history")
        self.api.check_conclusion = "failure"
        with self.assertRaisesRegex(setup.SetupError, "has not succeeded"):
            self.run_tool(apply=True)
        self.api.repo["permissions"]["admin"] = False
        with self.assertRaisesRegex(setup.SetupError, "administrative"):
            self.run_tool(apply=True)
        self.assertEqual(self.api.writes, [])

    def test_apply_without_drift_writes_nothing(self) -> None:
        self.assertEqual(self.run_tool(apply=True)[0], 0)
        self.assertEqual(self.api.writes, [])

    def test_read_back_detects_ignored_writes(self) -> None:
        self.api.repo["has_projects"] = True
        self.api.ignore_writes = True
        with self.assertRaisesRegex(setup.SetupError, "Read-back"):
            self.run_tool(apply=True)

    # Evidence.

    def test_evidence_records_verification_and_drift(self) -> None:
        with tempfile.TemporaryDirectory() as directory:
            target = Path(directory) / "evidence" / "settings-{date}.json"
            self.assertEqual(self.run_tool(verify=True, evidence=target)[0], 0)
            written = Path(directory) / "evidence" / "settings-2026-11-02.json"
            record = json.loads(written.read_text(encoding="utf-8"))
            self.assertEqual(record["result"], "VERIFIED")
            self.assertEqual(record["checked_at"], "2026-11-02T14:00:00Z")
            self.assertEqual(record["mode"], "verify")
            self.assertEqual(record["baseline"]["path"], "docs/repo_governance/repository-settings.json")
            self.assertEqual(len(record["baseline"]["sha256"]), 64)
            self.assertEqual({r["status"] for r in record["rulesets"]}, {"matches"})
            self.assertEqual(record["planned_changes"], [])

            self.api.repo["allow_squash_merge"] = False
            self.assertEqual(self.run_tool(evidence=target)[0], 1)
            record = json.loads(written.read_text(encoding="utf-8"))
            self.assertEqual(record["result"], "DRIFT")
            row = next(r for r in record["settings"] if r["setting"] == "allow_squash_merge")
            self.assertEqual((row["observed"], row["baseline"], row["matches"]), (False, True, False))
            self.assertEqual(record["planned_changes"], [{"method": "PATCH", "endpoint": API}])
        self.assertEqual(self.api.writes, [])

    # Baseline validation.

    def broken(self, change: Any) -> None:
        baseline = json.loads(setup.BASELINE.read_text(encoding="utf-8"))
        change(baseline)
        with patch.object(Path, "read_text", return_value=json.dumps(baseline)):
            setup.load_baseline(Path("baseline.json"))

    def test_invalid_baselines_are_rejected(self) -> None:
        def ruleset(b: Any, name: str) -> Any:
            return next(r for r in b["rulesets"] if r["name"] == name)

        cases = {
            "Push protection requires": lambda b: b["security"].update(secret_scanning=False),
            "security updates require": lambda b: b["security"].update(vulnerability_alerts=False),
            "squash merges": lambda b: b["settings"].update(allow_squash_merge=False),
            "bypass actors": lambda b: ruleset(b, "Protect main").update(
                bypass_actors=[{"actor_id": 5, "actor_type": "RepositoryRole", "bypass_mode": "always"}]),
            "unique name": lambda b: b["rulesets"].append(copy.deepcopy(b["rulesets"][0])),
            "repository settings": lambda b: b["settings"].update(has_pages=False),
            "schema": lambda b: b.update(schema_version=99),
            "active": lambda b: ruleset(b, "Protect main").update(enforcement="evaluate"),
            "repeats a rule": lambda b: ruleset(b, "Protect main")["rules"].append({"type": "deletion"}),
            "integration ID": lambda b: ruleset(b, "Protect main")["rules"][3]["parameters"][
                "required_status_checks"].append({"context": "x"}),
        }
        for message, change in cases.items():
            with self.subTest(message=message):
                with self.assertRaisesRegex(setup.SetupError, message):
                    self.broken(change)

    # Public-only security features on a private baseline (#207).

    def test_private_baseline_with_public_only_features_is_rejected(self) -> None:
        def private(b: Any) -> None:
            b["visibility"] = "private"
        with self.assertRaises(setup.SetupError) as raised:
            self.broken(private)
        for feature in ("private_vulnerability_reporting", "secret_scanning",
                        "secret_scanning_push_protection"):
            self.assertIn(feature, str(raised.exception))

        def private_without_them(b: Any) -> None:
            b["visibility"] = "private"
            b["security"].update(private_vulnerability_reporting=False, secret_scanning=False,
                                 secret_scanning_push_protection=False)
        self.broken(private_without_them)

    def test_rule_order_is_not_drift_but_an_extra_rule_is(self) -> None:
        rule = copy.deepcopy(setup.load_baseline(setup.BASELINE)["rulesets"][0])
        reordered = copy.deepcopy(rule)
        reordered["rules"].reverse()
        self.assertEqual(setup.rule_spec(rule), setup.rule_spec(reordered))
        reordered["rules"].append({"type": "creation"})
        self.assertNotEqual(setup.rule_spec(rule), setup.rule_spec(reordered))

    def test_pagination_reads_beyond_the_first_page(self) -> None:
        calls: list[str] = []

        def call(method: str, endpoint: str) -> Any:
            calls.append(endpoint)
            return list(range(100)) if endpoint.endswith("&page=1") else [100]

        self.assertEqual(len(setup.pages(call, f"{API}/rulesets?includes_parents=true")), 101)
        self.assertEqual(len(calls), 2)

    # The gh call itself.

    def test_cli_uses_explicit_method_host_and_stdin(self) -> None:
        result = argparse.Namespace(returncode=0, stdout="{}", stderr="")
        with patch("setup_repository.subprocess.run", return_value=result) as run:
            setup.api("PATCH", "repos/a/b", {"has_projects": False})
        command = run.call_args.args[0]
        self.assertEqual(command[:7], ["gh", "api", "--hostname", "github.com", "--method", "PATCH", "-H"])
        self.assertEqual(command[-2:], ["--input", "-"])
        self.assertEqual(run.call_args.kwargs["input"], '{"has_projects": false}')
        self.assertEqual(run.call_args.kwargs["encoding"], "utf-8")
        self.assertNotIn("GH_DEBUG", run.call_args.kwargs["env"])

    def test_api_error_keeps_the_status_but_not_server_detail(self) -> None:
        result = argparse.Namespace(returncode=1, stdout="", stderr="secret token HTTP 403")
        with patch("setup_repository.subprocess.run", return_value=result):
            with self.assertRaises(setup.ApiError) as caught:
                setup.api("GET", "repos/a/b")
        self.assertEqual(caught.exception.status, "403")
        self.assertNotIn("secret token", str(caught.exception))


if __name__ == "__main__":
    unittest.main()
