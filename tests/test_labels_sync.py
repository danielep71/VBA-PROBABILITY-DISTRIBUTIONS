"""End-to-end tests of label pruning in .github/scripts/labels-sync.mjs (#201).

Each test runs the real script in reconcile mode against a local stand-in for
the GitHub REST API (GITHUB_API_URL), so it exercises the path the workflow
runs on the default branch: list the live labels, check what pruning would
delete, apply, and verify.
"""

from __future__ import annotations

import json
import re
import shutil
import subprocess
import threading
import unittest
from http.server import BaseHTTPRequestHandler, ThreadingHTTPServer
from pathlib import Path
from typing import Any
from urllib.parse import parse_qs, unquote, urlsplit

ROOT = Path(__file__).resolve().parents[1]
SCRIPT = ROOT / ".github/scripts/labels-sync.mjs"
WORKFLOW = ROOT / ".github/workflows/labels-sync.yml"
REPOSITORY = "owner/repo"


class FakeGitHub:
    """Labels and issues of one repository, and every request received."""

    def __init__(self, labels: list[dict[str, str]], issues: list[dict[str, Any]]) -> None:
        self.labels = [dict(label) for label in labels]
        self.issues = issues
        self.requests: list[tuple[str, str]] = []
        fake = self

        class Handler(BaseHTTPRequestHandler):
            def log_message(self, *args: object) -> None:
                pass

            def reply(self, status: int, body: object = None) -> None:
                data = b"" if body is None else json.dumps(body).encode()
                self.send_response(status)
                self.send_header("Content-Type", "application/json")
                self.send_header("Content-Length", str(len(data)))
                self.end_headers()
                self.wfile.write(data)

            def handle_any(self) -> None:
                length = int(self.headers.get("Content-Length") or 0)
                body = json.loads(self.rfile.read(length) or b"null")
                fake.requests.append((self.command, self.path))
                status, payload = fake.respond(self.command, self.path, body)
                self.reply(status, payload)

            do_GET = do_POST = do_PATCH = do_DELETE = handle_any

        self.server = ThreadingHTTPServer(("127.0.0.1", 0), Handler)
        self.thread = threading.Thread(target=self.server.serve_forever, daemon=True)

    @property
    def url(self) -> str:
        return f"http://127.0.0.1:{self.server.server_address[1]}"

    def __enter__(self) -> FakeGitHub:
        self.thread.start()
        return self

    def __exit__(self, *args: object) -> None:
        self.server.shutdown()
        self.server.server_close()

    def respond(self, method: str, path: str, body: Any) -> tuple[int, object]:
        parts = urlsplit(path)
        query = parse_qs(parts.query)
        prefix = f"/repos/{REPOSITORY}"
        route = parts.path[len(prefix):] if parts.path.startswith(prefix) else parts.path
        if method == "GET" and route == "/labels":
            page, size = int(query.get("page", ["1"])[0]), int(query.get("per_page", ["30"])[0])
            return 200, self.labels[(page - 1) * size:page * size]
        if method == "GET" and route == "/issues":
            items = self.issues
            if "labels" in query:
                wanted = {name.lower() for name in query["labels"][0].split(",")}
                items = [item for item in items
                         if wanted <= {label["name"].lower() for label in item["labels"]}]
            page, size = int(query.get("page", ["1"])[0]), int(query.get("per_page", ["30"])[0])
            return 200, items[(page - 1) * size:page * size]
        if method == "DELETE" and route.startswith("/labels/"):
            name = unquote(route[len("/labels/"):])
            self.labels = [label for label in self.labels if label["name"] != name]
            return 204, None
        if method == "POST" and route == "/labels":
            self.labels.append({key: body[key] for key in ("name", "color", "description")})
            return 201, body
        if method == "PATCH" and route.startswith("/labels/"):
            name = unquote(route[len("/labels/"):])
            for label in self.labels:
                if label["name"] == name:
                    label.update({"name": body.get("new_name", name), "color": body["color"],
                                  "description": body["description"]})
            return 200, body
        return 404, {"message": f"unexpected {method} {path}"}

    def writes(self) -> list[tuple[str, str]]:
        return [request for request in self.requests if request[0] != "GET"]

    def issue_queries(self) -> list[str]:
        return [path for method, path in self.requests
                if method == "GET" and urlsplit(path).path.endswith("/issues")]


def catalogue() -> list[dict[str, str]]:
    manifest = json.loads((ROOT / ".github/labels.json").read_text(encoding="utf-8"))
    return [{key: label[key] for key in ("name", "color", "description")}
            for label in manifest["labels"]]


def extra(name: str) -> dict[str, str]:
    return {"name": name, "color": "ABCDEF", "description": "outside the catalogue"}


def issue(number: int, *names: str, pull_request: bool = False, state: str = "closed") -> dict:
    item: dict[str, Any] = {"number": number, "state": state,
                            "labels": [{"name": name} for name in names]}
    if pull_request:
        item["pull_request"] = {"url": f"https://example.invalid/pulls/{number}"}
    return item


def reconcile(fake: FakeGitHub) -> subprocess.CompletedProcess[str]:
    node = shutil.which("node")
    assert node, "Node.js is required: the label synchronization script runs on it"
    env = {"PATH": str(Path(node).parent), "GITHUB_API_URL": fake.url,
           "GITHUB_TOKEN": "test-token", "GITHUB_REPOSITORY": REPOSITORY}
    return subprocess.run([node, str(SCRIPT), "--mode", "reconcile", "--repository", REPOSITORY],
                          cwd=ROOT, env=env, capture_output=True, text=True, timeout=60,
                          check=False)


class PruneGuardTests(unittest.TestCase):
    def test_label_on_a_closed_issue_is_not_deleted(self) -> None:
        with FakeGitHub(catalogue() + [extra("legacy"), extra("stale")],
                        [issue(7, "legacy")]) as fake:
            result = reconcile(fake)
        self.assertEqual(result.returncode, 1, result.stdout + result.stderr)
        self.assertEqual(fake.writes(), [], "a change was applied before the refusal")
        self.assertIn('"legacy"', result.stderr)
        self.assertNotIn('"stale"', result.stderr)
        self.assertIn("No change was applied", result.stderr)

    def test_label_only_on_a_pull_request_is_not_deleted(self) -> None:
        with FakeGitHub(catalogue() + [extra("from-pr")],
                        [issue(8, "from-pr", pull_request=True, state="open")]) as fake:
            result = reconcile(fake)
        self.assertEqual(result.returncode, 1, result.stdout + result.stderr)
        self.assertEqual(fake.writes(), [])
        self.assertIn('"from-pr"', result.stderr)

    def test_label_with_a_comma_is_checked_on_every_page(self) -> None:
        # 100 unrelated items fill the first page; the label is on page 2,
        # with different letter case.
        issues = [issue(n, "bug") for n in range(1, 101)] + [issue(101, "Area,Core")]
        with FakeGitHub(catalogue() + [extra("area,core")], issues) as fake:
            result = reconcile(fake)
        self.assertEqual(result.returncode, 1, result.stdout + result.stderr)
        self.assertEqual(fake.writes(), [])
        self.assertIn('"area,core"', result.stderr)
        queries = fake.issue_queries()
        self.assertTrue(queries and all("labels=" not in query for query in queries), queries)
        self.assertEqual(len(queries), 2, queries)

    def test_unused_label_is_checked_then_deleted(self) -> None:
        with FakeGitHub(catalogue() + [extra("stale")], [issue(9, "bug")]) as fake:
            result = reconcile(fake)
        self.assertEqual(result.returncode, 0, result.stdout + result.stderr)
        self.assertEqual(fake.writes(), [("DELETE", "/repos/owner/repo/labels/stale")])
        self.assertEqual(fake.issue_queries(),
                         ["/repos/owner/repo/issues?state=all&per_page=1&labels=stale"])
        self.assertLess(fake.requests.index(("GET", fake.issue_queries()[0])),
                        fake.requests.index(("DELETE", "/repos/owner/repo/labels/stale")))


class WorkflowPermissionTests(unittest.TestCase):
    def test_reconcile_job_permissions_are_exact(self) -> None:
        text = WORKFLOW.read_text(encoding="utf-8")
        job = re.search(r"^  reconcile:\n(.*?)(?=^  \S|\Z)", text, re.M | re.S)
        assert job, "labels-sync.yml has no reconcile job"
        block = re.search(r"^    permissions:\n((?:      .*\n|\s*#.*\n)+)", job.group(1), re.M)
        assert block, "the reconcile job declares no permissions"
        granted = dict(re.findall(r"^      ([a-z-]+):\s*([a-z]+)", block.group(1), re.M))
        self.assertEqual(granted, {"contents": "read", "issues": "write", "pull-requests": "read"})


if __name__ == "__main__":
    unittest.main()
