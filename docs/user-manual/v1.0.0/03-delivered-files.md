# 3. Delivered files

## 3.1 The release

| File | Purpose |
|---|---|
| `VBA-Probability-Distributions-X.Y.Z.zip` | The bundle |
| `SHA256SUMS` | Checksum of the bundle |
| `VBA-Probability-Distributions-X.Y.Z.intoto.jsonl` | Build-provenance attestation |

`X.Y.Z` is the version number. The release also carries the source archive that GitHub creates from the tag.

## 3.2 Contents of the bundle

| Content | Description |
|---|---|
| Six production modules | The library (section 3.3) |
| `M_STATS_PROBDIST_TEST.bas` | Regression harness, optional |
| `LICENSE` | Mozilla Public License 2.0 |
| `README.md` | Overview and evidence summary |
| `INSTALLATION.md` | Installation and upgrade guide |
| `CHANGELOG.md` | Version history |
| `SECURITY.md` | Vulnerability reporting |
| `VERSION` | Version identity |
| `MANIFEST.sha256` | Checksum of every member |

The `.bas` files use CRLF line endings, ready for import into the VBA editor.

## 3.3 The six production modules

| Order | Module | Content |
|---:|---|---|
| 1 | `M_STATS_PROBDIST_CORE` | Constants, checks, diagnostics |
| 2 | `M_STATS_PROBDIST_SPECIALFUNCS` | Shared mathematical kernels |
| 3 | `M_STATS_PROBDIST_NORMALFAMILY` | Normal, Lognormal |
| 4 | `M_STATS_PROBDIST_TFAMILY` | Student t, Chi-square, F |
| 5 | `M_STATS_PROBDIST_CONTINUOUS` | Gamma, Beta, Exponential, Weibull, Uniform |
| 6 | `M_STATS_PROBDIST_DISCRETE` | The six discrete distributions |

The first two modules are internal and contain no worksheet functions. The other four contain the `K_STATS_` functions. All six are required.

The regression harness is not needed to use the library. Import it only to check an installation (section 4.5).

## 3.4 Verifying the bundle

Verify the bundle before you import anything from it.

```text
sha256sum --check SHA256SUMS
gh attestation verify VBA-Probability-Distributions-X.Y.Z.zip --repo danielep71/VBA-PROBABILITY-DISTRIBUTIONS
```

After extracting, run this inside the bundle folder to verify every module and document:

```text
sha256sum --check MANIFEST.sha256
```

On Windows, `Get-FileHash -Algorithm SHA256` gives the same digests.

## 3.5 Other sources

A source archive or a Git clone of the repository also contains the six modules, in the `src` folder. Take all six from the same version. Do not combine files from different tags, commits, release assets or copied workbooks.

## 3.6 Reference material in the repository

| Document | Content |
|---|---|
| [`docs/PUBLIC_API.md`](../../PUBLIC_API.md) | Public API reference |
| [`docs/PUBLIC_API.txt`](../../PUBLIC_API.txt) | Exact declarations |
| [`accuracy_guarantees.csv`](../../../benchmark/accuracy_guarantees.csv) | Guarantee per contract |
| [`numerical_limitations.csv`](../../../benchmark/numerical_limitations.csv) | Limitation register |
| [`accuracy_summary.md`](../../../benchmark/accuracy_summary.md) | Current verdicts |
