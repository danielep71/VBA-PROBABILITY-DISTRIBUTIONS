#!/usr/bin/env node
// Read-only detection of drift between .github/labels.json and the live
// repository labels.
//
// labels-sync.yml reconciles the live labels on trusted events and needs
// issues: write. This script never writes: it reports what that sync would
// change, so a label edited by hand in the GitHub UI is noticed without
// granting write permission to a scheduled job. The comparison mirrors the
// sync script exactly - names matched case-insensitively, then exact name,
// case-insensitive color and description (missing = "") compared, and extra
// live labels reported only when the manifest sets "prune": true. Change one
// and the other together.
//
// The idea follows the portfolio template's labels-drift workflow; the
// template's script is not reused because it reads a different labels.json
// schema (core plus profile/domain overlays).
//
//   node .github/scripts/labels-drift.mjs --self-test   offline fixtures
//   node .github/scripts/labels-drift.mjs --check       live; needs GITHUB_TOKEN
//                                                        and GITHUB_REPOSITORY
import { appendFile, readFile } from "node:fs/promises";
import { pathToFileURL } from "node:url";

const MANIFEST = ".github/labels.json";
const COLOR = /^[0-9a-fA-F]{6}$/;

export function validateManifest(manifest) {
  if (!manifest || typeof manifest !== "object") throw new Error("manifest must be an object");
  if (manifest.schema_version !== 1) throw new Error("schema_version must be 1");
  if (typeof manifest.prune !== "boolean") throw new Error("prune must be a boolean");
  if (!Array.isArray(manifest.labels) || manifest.labels.length === 0) {
    throw new Error("labels must be a non-empty array");
  }
  const seen = new Set();
  for (const label of manifest.labels) {
    if (typeof label.name !== "string" || label.name.trim() === "") throw new Error("label without a name");
    if (typeof label.color !== "string" || !COLOR.test(label.color)) {
      throw new Error(`label ${JSON.stringify(label.name)} has an invalid color`);
    }
    const key = label.name.toLowerCase();
    if (seen.has(key)) throw new Error(`duplicate label name ${JSON.stringify(label.name)}`);
    seen.add(key);
  }
  return manifest;
}

export function compareLabels(manifest, live) {
  validateManifest(manifest);
  const liveByKey = new Map(live.map((label) => [label.name.toLowerCase(), label]));
  const desiredKeys = new Set(manifest.labels.map((label) => label.name.toLowerCase()));
  const missing = [];
  const changed = [];
  for (const wanted of manifest.labels) {
    const existing = liveByKey.get(wanted.name.toLowerCase());
    if (!existing) {
      missing.push(wanted.name);
      continue;
    }
    const fields = [];
    if (existing.name !== wanted.name) fields.push("name");
    if (existing.color.toLowerCase() !== wanted.color.toLowerCase()) fields.push("color");
    if ((existing.description || "") !== (wanted.description || "")) fields.push("description");
    if (fields.length > 0) changed.push({ name: wanted.name, fields });
  }
  const extra = manifest.prune
    ? live.filter((label) => !desiredKeys.has(label.name.toLowerCase())).map((label) => label.name)
    : [];
  return { missing, changed, extra, drift: missing.length + changed.length + extra.length > 0 };
}

function report(result, total) {
  const list = (items) => (items.length > 0 ? items.join(", ") : "—");
  return [
    "## Label drift",
    "",
    `- **Status:** ${result.drift ? "DRIFT" : "in sync"}`,
    `- **Manifest labels:** ${total}`,
    `- **Missing on the repository:** ${list(result.missing)}`,
    `- **Different on the repository:** ${list(result.changed.map((c) => `${c.name} (${c.fields.join("/")})`))}`,
    `- **Extra on the repository (pruned by sync):** ${list(result.extra)}`,
    "",
    result.drift ? "Run **Label Taxonomy Sync** on main, or update `.github/labels.json`." : "",
  ].join("\n");
}

export async function listLiveLabels(repository, token, fetchImpl = fetch) {
  const labels = [];
  for (let page = 1; page <= 50; page += 1) {
    const response = await fetchImpl(
      `https://api.github.com/repos/${repository}/labels?per_page=100&page=${page}`,
      {
        headers: {
          Accept: "application/vnd.github+json",
          Authorization: `Bearer ${token}`,
          "User-Agent": "labels-drift",
          "X-GitHub-Api-Version": "2022-11-28",
        },
      },
    );
    if (!response.ok) throw new Error(`GitHub API returned ${response.status} listing labels`);
    const batch = await response.json();
    labels.push(...batch);
    if (batch.length < 100) return labels;
  }
  throw new Error("more than 5,000 labels; refusing to continue");
}

async function selfTest() {
  const manifest = {
    schema_version: 1,
    prune: true,
    labels: [
      { name: "bug", color: "d73a4a", description: "Something is wrong" },
      { name: "P1", color: "B60205", description: "" },
    ],
  };
  const live = () => [
    { name: "bug", color: "D73A4A", description: "Something is wrong" },
    { name: "P1", color: "b60205", description: null },
  ];
  const failures = [];
  const expect = (label, result, missing, changed, extra) => {
    const got = JSON.stringify([result.missing, result.changed.map((c) => `${c.name}:${c.fields}`), result.extra]);
    const want = JSON.stringify([missing, changed, extra]);
    if (got !== want) failures.push(`${label}: expected ${want}, got ${got}`);
  };
  const throws = (label, fn) => {
    try {
      fn();
      failures.push(`${label}: expected a rejection`);
    } catch {
      // expected
    }
  };

  expect("in sync; color case and null description are not drift", compareLabels(manifest, live()), [], [], []);
  expect("missing label", compareLabels(manifest, live().slice(0, 1)), ["P1"], [], []);
  expect("changed color", compareLabels(manifest, [live()[0], { ...live()[1], color: "000000" }]), [], ["P1:color"], []);
  expect("changed description", compareLabels(manifest, [{ ...live()[0], description: "x" }, live()[1]]), [], ["bug:description"], []);
  expect("name case differs", compareLabels(manifest, [live()[0], { ...live()[1], name: "p1" }]), [], ["P1:name"], []);
  expect("extra label is drift when pruning", compareLabels(manifest, [...live(), { name: "wontfix", color: "ffffff" }]), [], [], ["wontfix"]);
  expect("extra label is allowed without pruning", compareLabels({ ...manifest, prune: false }, [...live(), { name: "wontfix", color: "ffffff" }]), [], [], []);
  throws("duplicate names differing only by case", () => validateManifest({ ...manifest, labels: [...manifest.labels, { name: "BUG", color: "000000" }] }));
  throws("invalid color", () => validateManifest({ ...manifest, labels: [{ name: "x", color: "#fff" }] }));
  throws("prune must be boolean", () => validateManifest({ ...manifest, prune: "yes" }));

  // Pagination: 150 labels arrive as a full page of 100 and a final page of 50.
  const pages = [];
  const fakeFetch = async (url) => {
    pages.push(Number(new URL(url).searchParams.get("page")));
    const page = pages[pages.length - 1];
    const size = page === 1 ? 100 : page === 2 ? 50 : 0;
    return { ok: true, json: async () => Array.from({ length: size }, (_, i) => ({ name: `l${page}-${i}`, color: "000000" })) };
  };
  const all = await listLiveLabels("owner/repo", "token", fakeFetch);
  if (all.length !== 150 || pages.join() !== "1,2") failures.push(`pagination: got ${all.length} labels from pages ${pages}`);
  try {
    await listLiveLabels("owner/repo", "token", async () => ({ ok: false, status: 403 }));
    failures.push("API error: expected a rejection");
  } catch {
    // expected
  }
  return failures;
}

async function main(argv) {
  const manifest = validateManifest(JSON.parse(await readFile(MANIFEST, "utf8")));
  if (argv.includes("--self-test")) {
    const failures = await selfTest();
    if (failures.length > 0) {
      console.error(`SELF-TEST FAIL:\n  - ${failures.join("\n  - ")}`);
      return 1;
    }
    console.log(`SELF-TEST PASS: in-sync, missing, color, description, name-case, prune/no-prune extras and manifest validation, pagination and API errors; ${MANIFEST} (${manifest.labels.length} labels) is valid.`);
    return 0;
  }
  if (argv.includes("--check")) {
    const { GITHUB_REPOSITORY: repository, GITHUB_TOKEN: token } = process.env;
    if (!repository || !token) throw new Error("--check needs GITHUB_REPOSITORY and GITHUB_TOKEN");
    const result = compareLabels(manifest, await listLiveLabels(repository, token));
    const summary = report(result, manifest.labels.length);
    console.log(summary);
    if (process.env.GITHUB_STEP_SUMMARY) await appendFile(process.env.GITHUB_STEP_SUMMARY, `${summary}\n`, "utf8");
    return result.drift ? 1 : 0;
  }
  console.error("usage: labels-drift.mjs --self-test | --check");
  return 2;
}

// Run only when executed directly; importing the module (for example from a
// test harness, where process.argv[1] may be absent) must not start main().
if (process.argv[1] && import.meta.url === pathToFileURL(process.argv[1]).href) {
  main(process.argv.slice(2)).then(
    (code) => process.exit(code),
    (error) => {
      console.error(`ERROR: ${error.message}`);
      process.exit(1);
    },
  );
}
