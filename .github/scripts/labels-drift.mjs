#!/usr/bin/env node
// Read-only detection of drift between .github/labels.json and the live
// repository labels.
//
// labels-sync.yml reconciles the live labels on trusted events and needs
// issues: write. This script never writes: it reports what that sync would
// change, so a label edited by hand in the GitHub UI is noticed without
// granting write permission to a scheduled job. The comparison mirrors the
// sync script exactly - a pending rename (old live name, new manifest name)
// matched first, then names matched case-insensitively, then exact name,
// case-insensitive color and description (missing = "") compared, and extra
// live labels reported only when the manifest sets "prune": true. Change one
// and the other together.
//
// The manifest follows the portfolio template's label format rules: uppercase
// six-digit colors, names of at most 50 and single-line descriptions of 1-100
// characters, unique case-insensitively, sorted case-insensitively by name. The
// template's script itself is not reused because it reads a different
// labels.json schema (core plus profile/domain overlays).
//
// "renames" maps an old label name to its manifest name. Sync renames the
// live label in place, keeping it on every issue and pull request, instead of
// pruning the old label and creating a new, unused one.
//
//   node .github/scripts/labels-drift.mjs --self-test   offline fixtures
//   node .github/scripts/labels-drift.mjs --check       live; needs GITHUB_TOKEN
//                                                        and GITHUB_REPOSITORY
import { appendFile, readFile } from "node:fs/promises";
import { pathToFileURL } from "node:url";

const MANIFEST = ".github/labels.json";
const COLOR = /^[0-9A-F]{6}$/;
// The template's compareNames: code-unit order of the lower-cased names, with
// the exact names as tie-break. Locale-independent on purpose.
function byName(left, right) {
  const leftKey = left.toLowerCase();
  const rightKey = right.toLowerCase();
  if (leftKey < rightKey) return -1;
  if (leftKey > rightKey) return 1;
  return left < right ? -1 : left > right ? 1 : 0;
}

export function validateManifest(manifest) {
  if (!manifest || typeof manifest !== "object") throw new Error("manifest must be an object");
  if (manifest.schema_version !== 1) throw new Error("schema_version must be 1");
  if (typeof manifest.prune !== "boolean") throw new Error("prune must be a boolean");
  if (!Array.isArray(manifest.labels) || manifest.labels.length === 0) {
    throw new Error("labels must be a non-empty array");
  }
  const seen = new Set();
  let previous = null;
  for (const label of manifest.labels) {
    if (typeof label.name !== "string" || label.name.trim() === "" || label.name !== label.name.trim()) {
      throw new Error("label without a name, or with surrounding whitespace");
    }
    const name = JSON.stringify(label.name);
    if (label.name.length > 50 || /\r|\n/.test(label.name)) throw new Error(`label ${name} is not a single line of at most 50 characters`);
    if (typeof label.color !== "string" || !COLOR.test(label.color)) {
      throw new Error(`label ${name} color must be six uppercase hexadecimal characters without #`);
    }
    if (typeof label.description !== "string" || label.description.length === 0
        || label.description.length > 100 || /\r|\n/.test(label.description)) {
      throw new Error(`label ${name} description must be a single line of 1-100 characters`);
    }
    const key = label.name.toLowerCase();
    if (seen.has(key)) throw new Error(`duplicate label name ${name}`);
    seen.add(key);
    if (previous !== null && byName(previous, label.name) >= 0) {
      throw new Error(`labels must be sorted case-insensitively by name: ${JSON.stringify(previous)} precedes ${name}`);
    }
    previous = label.name;
  }
  const renames = manifest.renames ?? {};
  if (typeof renames !== "object" || Array.isArray(renames)) throw new Error("renames must be an object");
  for (const [from, to] of Object.entries(renames)) {
    if (typeof to !== "string" || !seen.has(to.toLowerCase()) || manifest.labels.every((label) => label.name !== to)) {
      throw new Error(`rename target ${JSON.stringify(to)} must be a manifest label name`);
    }
    if (seen.has(from.toLowerCase())) throw new Error(`rename source ${JSON.stringify(from)} must not be a manifest label`);
  }
  return manifest;
}

// Live labels keyed by the lower-case manifest name they correspond to. A live
// label under a rename source maps to its target, unless the target also
// exists live: then both stay as they are, and the source is reported extra
// (sync refuses that state rather than pruning a label that is still in use).
export function liveByManifestKey(manifest, live) {
  const byKey = new Map(live.map((label) => [label.name.toLowerCase(), label]));
  for (const [from, to] of Object.entries(manifest.renames ?? {})) {
    const source = byKey.get(from.toLowerCase());
    if (source && !byKey.has(to.toLowerCase())) {
      byKey.delete(from.toLowerCase());
      byKey.set(to.toLowerCase(), source);
    }
  }
  return byKey;
}

export function compareLabels(manifest, live) {
  validateManifest(manifest);
  const liveByKey = liveByManifestKey(manifest, live);
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
    ? [...liveByKey].filter(([key]) => !desiredKeys.has(key)).map(([, label]) => label.name)
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
      { name: "bug", color: "D73A4A", description: "Something is wrong" },
      { name: "P1", color: "B60205", description: "Critical" },
    ],
  };
  const live = () => [
    { name: "bug", color: "d73a4a", description: "Something is wrong" },
    { name: "P1", color: "b60205", description: "Critical" },
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

  expect("in sync; live color case is not drift", compareLabels(manifest, live()), [], [], []);
  expect("null live description is drift against a non-empty one", compareLabels(manifest, [live()[0], { ...live()[1], description: null }]), [], ["P1:description"], []);
  expect("missing label", compareLabels(manifest, live().slice(0, 1)), ["P1"], [], []);
  expect("changed color", compareLabels(manifest, [live()[0], { ...live()[1], color: "000000" }]), [], ["P1:color"], []);
  expect("changed description", compareLabels(manifest, [{ ...live()[0], description: "x" }, live()[1]]), [], ["bug:description"], []);
  expect("name case differs", compareLabels(manifest, [live()[0], { ...live()[1], name: "p1" }]), [], ["P1:name"], []);
  expect("extra label is drift when pruning", compareLabels(manifest, [...live(), { name: "wontfix", color: "ffffff" }]), [], [], ["wontfix"]);
  expect("extra label is allowed without pruning", compareLabels({ ...manifest, prune: false }, [...live(), { name: "wontfix", color: "ffffff" }]), [], [], []);
  const renamed = { ...manifest, renames: { testing: "P1" } };
  expect("pending rename is a name change, not missing plus extra", compareLabels(renamed, [live()[0], { ...live()[1], name: "testing" }]), [], ["P1:name"], []);
  expect("applied rename is in sync", compareLabels(renamed, live()), [], [], []);
  expect("rename source beside its live target is extra", compareLabels(renamed, [...live(), { name: "Testing", color: "006B75" }]), [], [], ["Testing"]);
  throws("rename target outside the manifest", () => validateManifest({ ...manifest, renames: { testing: "tests" } }));
  throws("rename source still in the manifest", () => validateManifest({ ...manifest, renames: { BUG: "P1" } }));
  throws("renames must be an object", () => validateManifest({ ...manifest, renames: ["testing"] }));
  throws("lowercase color", () => validateManifest({ ...manifest, labels: [{ name: "bug", color: "d73a4a", description: "x" }] }));
  throws("empty description", () => validateManifest({ ...manifest, labels: [{ name: "bug", color: "D73A4A", description: "" }] }));
  throws("description over 100 characters", () => validateManifest({ ...manifest, labels: [{ name: "bug", color: "D73A4A", description: "x".repeat(101) }] }));
  throws("name over 50 characters", () => validateManifest({ ...manifest, labels: [{ name: "x".repeat(51), color: "D73A4A", description: "x" }] }));
  throws("unsorted labels", () => validateManifest({ ...manifest, labels: [...manifest.labels].reverse() }));
  throws("duplicate names differing only by case", () => validateManifest({ ...manifest, labels: [{ name: "BUG", color: "000000", description: "x" }, ...manifest.labels] }));
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
    console.log(`SELF-TEST PASS: in-sync, missing, color, description, name-case, prune/no-prune extras, renames and manifest format validation, pagination and API errors; ${MANIFEST} (${manifest.labels.length} labels) is valid.`);
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
