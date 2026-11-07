# 📦 Distribution review

A release publishes more than the files in the bundle. Before each public
release, record the review below in an issue and close it. Do not publish the
release while that issue is open. [RELEASING.md](../../RELEASING.md) makes the
closed review a release step.

## 🚪 What counts as distribution

This repository is public, so every push to GitHub already publishes its
**whole history**: every branch, tag and reachable commit, and every issue,
pull request, comment, workflow run log and artifact. A file deleted in a later
commit is still published with the commit that added it. The review therefore
covers what is already public as well as what the release adds.

Run the review before any of these:

- publishing a GitHub Release or attaching a file to one;
- sharing a clone, an archive, an exported module or a workbook outside GitHub;
- changing the repository's visibility, which the settings baseline records.

## ✅ The review

Open an issue for the review, triaged like any other, with the `release` and
`security` labels. Copy this checklist into it, and record for each item what
you checked and what you found:

```markdown
## Distribution review

- What is distributed, to whom, and from which commit (full SHA):

- [ ] **License.** `LICENSE` (MPL-2.0) is still the intended license for the
      original content.
- [ ] **Provenance.** Every file is original, or its source, license and
      attribution are recorded (`tools/README.md` Provenance,
      `benchmark/PROVENANCE.md` and the study READMEs for reference data).
- [ ] **Third-party claims.** No employer, client or vendor has a claim on the
      content or on the methods it implements.
- [ ] **History scan.** The pinned secret scan below found no leak in any
      reachable commit, and the manual searches found no sensitive material.
      Scanner version, commit count, pattern categories and redacted findings:
- [ ] **Binary files.** Each binary file ever committed was inspected,
      including every workbook as described below.
- [ ] **Issues and pull requests.** Their text, comments, edit history and
      attachments contain no confidential material.
- [ ] **Actions.** No workflow run log or artifact holds sensitive output.
- [ ] **Repository settings.** `python3 tools/setup_repository.py --verify
      --evidence <path>` reports VERIFIED. Date and result of the record:
```

Close the issue only when every item is checked. If the review finds
something, handle it as *Handling a finding* describes and review again.

## 🔍 History scan

Scan every commit reachable from any branch or tag, not only the current tree,
from a fresh, complete clone. A shallow clone hides history: check that
`git rev-parse --is-shallow-repository` prints `false`.

```sh
git clone https://github.com/danielep71/VBA-PROBABILITY-DISTRIBUTIONS.git review
cd review
git fetch origin '+refs/heads/*:refs/remotes/origin/*' '+refs/tags/*:refs/tags/*'
git rev-parse --is-shallow-repository
git rev-list --all --count
```

### Pinned secret scan

Use [Gitleaks](https://github.com/gitleaks/gitleaks) **8.30.1**. Download the
release archive for your platform and check its SHA-256 against the
release's `gitleaks_8.30.1_checksums.txt`; for Linux x86-64:

```sh
curl -fsSLO https://github.com/gitleaks/gitleaks/releases/download/v8.30.1/gitleaks_8.30.1_linux_x64.tar.gz
echo '551f6fc83ea457d62a0d98237cbad105af8d557003051f41f3e7ca7b3f2470eb  gitleaks_8.30.1_linux_x64.tar.gz' \
  | sha256sum --check --strict
tar -xzf gitleaks_8.30.1_linux_x64.tar.gz gitleaks
```

Then scan every commit of every ref, with the default rules and findings
redacted, from the clone's root:

```sh
./gitleaks git --log-opts="--all" --redact --no-banner \
  --report-format json --report-path ../gitleaks-report.json .
```

Exit status `0` with `no leaks found` passes. Exit status `1` means findings:
read the report, which stays outside the repository, and handle each one. The
scanner reads the patches of non-merge commits; it does not look inside
binary files, so the inspection below still applies.

### Manual searches

The scanner looks for credentials only. Search the text of every commit for
the other categories too:

```sh
# Every reachable blob, with the path it was committed under.
git rev-list --objects --all
# Every binary file ever committed, under each path it had: inspect each one.
git log --all --numstat --no-renames --format= | awk -F'\t' '$1 == "-" {print $3}' | sort -u
# Search the text of every commit for a pattern.
git rev-list --all | xargs git grep -n -I -i -E 'PATTERN'
```

Search for at least:

- credentials: tokens, keys, passwords, connection strings and private
  certificates;
- people: names, email addresses and user names other than your own;
- organizations: employer, client, counterparty and vendor names, internal
  host names, paths and system identifiers;
- real data: portfolio, trade, account or position identifiers, real
  amounts, and any data that is not synthetic or published reference data.

The review issue is public, so it must not become another copy of what the
scan looks for. Record in it the scanner version, the scanned commit count,
the categories of patterns, such as "client names" or "internal host names",
and each finding redacted to its commit, path and category. Keep the literal
patterns and any sensitive evidence outside the repository, in a private note;
report a suspected vulnerability through the private channel in
[SECURITY.md](../../SECURITY.md). The `traffic-history` branch holds aggregate
counts only; it is scanned with the rest.

## 📗 Workbooks

A workbook holds content that its exported source and a text diff do not
show. Open each workbook ever committed (the binary-file list above names
them) in Excel and run **File → Info → Check for Issues → Inspect Document**,
or unzip the `.xlsm` and read its parts. Look for:

| Content | Where it is stored |
| --- | --- |
| Author, last editor, company and manager | `docProps/core.xml`, `docProps/app.xml` |
| Hidden and very hidden sheets | `state` attributes in `xl/workbook.xml` |
| Defined names, including hidden ones | `definedNames` in `xl/workbook.xml` |
| Links to other workbooks and data connections | `xl/externalLinks/`, `xl/connections.xml` |
| Comments and threaded comments, with their authors | `xl/comments*.xml`, `xl/threadedComments/`, `xl/persons/` |
| Cached source data of pivot tables | `xl/pivotCache/` |
| Values left in cells, rows or columns outside the visible range | the sheet parts under `xl/worksheets/` |

The VBA project inside a workbook can differ from the exported modules in
`src/` and `tests/`. Export the modules from the workbook itself and read
them too.

## 🧹 Handling a finding

- **A credential:** revoke or rotate it first. Removing it from the history
  does not undo its exposure.
- **A file or text in the history:** stop the release and decide, in a
  private note, whether to rewrite history. Rewriting needs `git filter-repo`
  on every affected branch and tag, a force-push that the protection rulesets
  forbid until they are relaxed, and new SHAs that invalidate the exact-SHA
  evidence bound to the old commits. If you do it, restore the rulesets and
  run `python3 tools/setup_repository.py --verify` afterwards. Every existing
  clone keeps the old history. On GitHub, pull request references and cached
  views can keep it too; GitHub's guide to
  [removing sensitive data](https://docs.github.com/en/authentication/keeping-your-account-and-data-secure/removing-sensitive-data-from-a-repository)
  explains what needs GitHub Support.
- **Issue or pull request text:** edit it, then delete the earlier revisions
  from the comment's edit history, which readers can open. Removing the link
  to an attachment does not delete the file; ask GitHub Support to remove it.
- **A workflow run log or artifact:** delete the run.

Then run the review again from the start.
