# Public VBA API compatibility contract

`docs/PUBLIC_API.txt` is the machine-readable compatibility baseline for the
worksheet-facing VBA product surface of VBA-PROBABILITY-DISTRIBUTIONS.

## Scope

The product API is the set of `Public K_STATS_*` declarations in the six
production modules under `src/`:

- `M_STATS_PROBDIST_CORE.bas`
- `M_STATS_PROBDIST_SPECIALFUNCS.bas`
- `M_STATS_PROBDIST_NORMALFAMILY.bas`
- `M_STATS_PROBDIST_TFAMILY.bas`
- `M_STATS_PROBDIST_CONTINUOUS.bas`
- `M_STATS_PROBDIST_DISCRETE.bas`

The v1.0.0 baseline contains **112 declarations**.

The manifest deliberately does not equate VBA `Public` scope with product API
scope. `PROB_*` routines may be `Public` inside `Option Private Module` modules
so production modules can share numerical kernels while those helpers remain
project-internal rather than worksheet-facing. Benchmark and test exporters are
also outside the product API.

## What is frozen

Each manifest row records:

1. production module path;
2. declaration kind;
3. `K_STATS_*` name; and
4. normalized VBA declaration.

The normalized declaration preserves compatibility-significant details,
including parameter order, parameter names and types, `ByVal`/`ByRef`,
`Optional`, default values, and return type. VBA physical line wrapping,
ordinary declaration whitespace, comments, and function bodies are not part of
the compatibility key.

Moving a function between production modules is also treated as drift. This is
intentional: module ownership is part of the maintained source contract and can
matter to downstream import/maintenance workflows even when Excel worksheet
syntax is unchanged.

## Gate

Run from the repository root:

```bash
python benchmark/check_public_api.py
```

The check regenerates the canonical surface from source and compares it with
`docs/PUBLIC_API.txt`. It fails on an added, removed, renamed, moved, or
signature-changed `K_STATS_*` declaration.

Negative controls are in:

```bash
python benchmark/test_public_api.py
```

Those fixtures prove that name, parameter type/order, `ByVal`/`ByRef`,
`Optional`/default, return type, additive API, and module movement are detected,
while function-body edits, harmless formatting, `PROB_*` project helpers, and
benchmark exporters do not create product-API drift.

Both commands are included in `benchmark/test_evidence_tools.py`, so the hosted
portable evidence gate exercises the checker and its controls together.

## Intentional API changes

Do not edit the manifest merely to make a failing check green. For an
intentional API change:

1. document the user-facing compatibility decision and owning issue;
2. update the production declaration and its regression/documentation coverage;
3. review the semantic-version impact;
4. regenerate the baseline with:

   ```bash
   python benchmark/check_public_api.py --write
   ```

5. review the `docs/PUBLIC_API.txt` diff declaration by declaration; and
6. commit the source, tests, documentation, changelog, and reviewed baseline in
   the same change set.

Once a stable release exists, removing or incompatibly changing a supported
signature is a breaking change. Adding a new supported worksheet function is an
API addition and must also be explicit rather than slipping through an internal
numerical change.

## Non-goals

This gate does not establish numerical correctness, supported-domain accuracy,
or Excel-runtime behavior. Those remain owned by the regression harness,
accuracy contracts, source/grid provenance, and independent holdout evidence.
The API gate protects interface compatibility only.
