"""
Fail-closed disposition checker for the main accuracy grid (issue #22).

Every `evidence_set = main grid` row must be claimed by an active or
characterization contract, or by a final explicit row exemption.  Strict mode
fails on any missing disposition.  During the v1.0.0 transition only, transition
mode permits the unresolved subset of one frozen binary64 fingerprint and
reports it as KNOWN COVERAGE DEBT; it rejects new, changed, or removed debt rows.

Run:
    python check_grid_coverage.py --mode strict
    python check_grid_coverage.py --mode transition --check-summary
"""
import argparse
import csv
import hashlib
import json
import math
import os
import struct
from collections import Counter

HERE = os.path.dirname(os.path.abspath(__file__))
IDENTITY_SCHEMA = "function+arg1..arg4(binary64)+regime+evidence_set/v1"
FINGERPRINT_SCHEMA = "coverage-debt/v1"
EXEMPTION_SCHEMA = "accuracy-row-exemptions/v1"
CONTRACT_STATUSES = {"active", "characterization_only"}
EXEMPTION_STATUS = "active"
ARG_FIELDS = ("arg1", "arg2", "arg3", "arg4")


def normalized_hash(path):
    """SHA-256 over LF-normalized bytes, matching the provenance policy."""
    data = open(path, "rb").read().replace(b"\r\n", b"\n").replace(b"\r", b"\n")
    return "sha256:" + hashlib.sha256(data).hexdigest()


def _bits(text):
    """Canonical binary64 argument identity; blank arguments remain blank."""
    value = (text or "").strip()
    if not value:
        return ""
    number = float(value)
    if not math.isfinite(number):
        raise ValueError(f"non-finite argument {value!r}")
    return struct.pack(">d", number).hex()


def row_identity(row):
    """Return the canonical, JSON-serializable identity of one grid row."""
    function = (row.get("function") or "").strip()
    regime = (row.get("regime") or "all").strip()
    evidence_set = (row.get("evidence_set") or "").strip()
    if not function or not regime:
        raise ValueError("row has a blank function or regime")
    identity = {"function": function}
    for field in ARG_FIELDS:
        identity[field + "_bits"] = _bits(row.get(field, ""))
    identity["regime"] = regime
    identity["evidence_set"] = evidence_set
    return identity


def identity_key(identity):
    return tuple(identity[k] for k in (
        "function", "arg1_bits", "arg2_bits", "arg3_bits", "arg4_bits",
        "regime", "evidence_set"))


def _identity_from_json(value):
    required = ("function", "arg1_bits", "arg2_bits", "arg3_bits",
                "arg4_bits", "regime", "evidence_set")
    if not isinstance(value, dict) or any(k not in value for k in required):
        raise ValueError("identity is missing canonical fields")
    for field in ("arg1_bits", "arg2_bits", "arg3_bits", "arg4_bits"):
        bits = value[field]
        if bits and (len(bits) != 16 or any(c not in "0123456789abcdef" for c in bits)):
            raise ValueError(f"invalid binary64 field {field}")
    return {k: str(value[k]) for k in required}


def _load_json(path, description):
    try:
        with open(path, encoding="utf-8") as f:
            return json.load(f), []
    except FileNotFoundError:
        return None, [f"missing {description}: {os.path.basename(path)}"]
    except (OSError, ValueError) as exc:
        return None, [f"malformed {description}: {type(exc).__name__}: {exc}"]


def _load_exemptions(path):
    data, errors = _load_json(path, "row-exemption registry")
    if errors:
        return [], errors
    if not isinstance(data, dict) or data.get("schema_version") != EXEMPTION_SCHEMA:
        return [], [f"row-exemption schema must be {EXEMPTION_SCHEMA!r}"]
    values = data.get("exemptions")
    if not isinstance(values, list):
        return [], ["row-exemption registry must contain an exemptions list"]
    return values, []


def evaluate_coverage(rows, contracts, exemptions):
    """Classify main-grid rows and validate every explicit exemption."""
    errors = []
    contract_keys = {
        ((c.get("function") or "").strip(), (c.get("regime") or "all").strip())
        for c in contracts if (c.get("status") or "").strip() in CONTRACT_STATUSES
    }

    main = []
    for index, row in enumerate(rows, start=2):
        if (row.get("evidence_set") or "").strip() != "main grid":
            continue
        try:
            ident = row_identity(row)
        except (ValueError, OverflowError) as exc:
            errors.append(f"grid row {index}: {exc}")
            continue
        main.append((row, ident, identity_key(ident)))

    unclaimed = [(row, ident, key) for row, ident, key in main
                 if ((row.get("function") or "").strip(),
                     (row.get("regime") or "all").strip()) not in contract_keys]
    unclaimed_keys = Counter(key for _, _, key in unclaimed)
    duplicate_missing = [key for key, count in unclaimed_keys.items() if count > 1]
    if duplicate_missing:
        errors.append(f"{len(duplicate_missing)} duplicate unclaimed canonical identity(ies)")

    exemption_keys = set()
    exemption_ids = set()
    for pos, exemption in enumerate(exemptions, start=1):
        if not isinstance(exemption, dict):
            errors.append(f"exemption {pos}: entry is not an object")
            continue
        exemption_id = str(exemption.get("exemption_id") or "").strip()
        if not exemption_id:
            errors.append(f"exemption {pos}: exemption_id is required")
        elif exemption_id in exemption_ids:
            errors.append(f"duplicate exemption_id {exemption_id!r}")
        exemption_ids.add(exemption_id)
        if exemption.get("status") != EXEMPTION_STATUS:
            errors.append(f"exemption {exemption_id or pos}: unknown status "
                          f"{exemption.get('status')!r}")
        for field in ("reason", "owner", "review_policy"):
            if not str(exemption.get(field) or "").strip():
                errors.append(f"exemption {exemption_id or pos}: {field} is required")
        try:
            key = identity_key(_identity_from_json(exemption.get("identity")))
        except (ValueError, TypeError) as exc:
            errors.append(f"exemption {exemption_id or pos}: {exc}")
            continue
        if key in exemption_keys:
            errors.append(f"duplicate/conflicting exemption identity in {exemption_id or pos}")
        exemption_keys.add(key)
        if key not in unclaimed_keys:
            errors.append(f"stale exemption {exemption_id or pos}: matches no unclaimed main-grid row")

    claimed = [(row, ident, key) for row, ident, key in main
               if ((row.get("function") or "").strip(),
                   (row.get("regime") or "all").strip()) in contract_keys]
    exempt = [(row, ident, key) for row, ident, key in unclaimed if key in exemption_keys]
    missing = [(row, ident, key) for row, ident, key in unclaimed if key not in exemption_keys]
    return {
        "total_rows": len(rows),
        "main_rows": len(main),
        "claimed_rows": len(claimed),
        "exempt_rows": len(exempt),
        "missing_rows": len(missing),
        "main_identities": Counter(key for _, _, key in main),
        "missing_identities": Counter(key for _, _, key in missing),
        "missing_identity_objects": [ident for _, ident, _ in missing],
        "missing_groups": Counter((row["function"], row.get("regime", "all") or "all")
                                  for row, _, _ in missing),
        "errors": errors,
    }


def build_fingerprint(root, grid_path, result, source_commit, owner, issue,
                      milestone_expiry):
    """Build the one-time reviewed v1.0.0 transition fingerprint."""
    src = {}
    src_dir = os.path.join(root, "src")
    for name in sorted(os.listdir(src_dir)):
        if name.lower().endswith(".bas"):
            rel = "src/" + name
            src[rel] = normalized_hash(os.path.join(src_dir, name))
    rows = sorted(result["missing_identity_objects"], key=identity_key)
    return {
        "schema_version": FINGERPRINT_SCHEMA,
        "identity_schema": IDENTITY_SCHEMA,
        "owner": owner,
        "issue": int(issue),
        "milestone_expiry": milestone_expiry,
        "source_identity": {
            "commit_sha": source_commit,
            "production_modules": src,
            "grid_sha256": normalized_hash(grid_path),
        },
        "initial_audit_baseline_count": len(rows),
        "rows": rows,
    }


def validate_transition(fingerprint, result):
    """Permit only resolved or still-present members of the frozen debt set."""
    errors = list(result["errors"])
    if not isinstance(fingerprint, dict):
        return errors + ["coverage-debt fingerprint is not an object"]
    for field, expected in (("schema_version", FINGERPRINT_SCHEMA),
                            ("identity_schema", IDENTITY_SCHEMA),
                            ("issue", 22),
                            ("milestone_expiry", "v1.0.0")):
        if fingerprint.get(field) != expected:
            errors.append(f"fingerprint {field} must be {expected!r}")
    if not str(fingerprint.get("owner") or "").strip():
        errors.append("fingerprint owner is required")
    source = fingerprint.get("source_identity")
    if not isinstance(source, dict):
        errors.append("fingerprint source_identity is required")
    else:
        sha = str(source.get("commit_sha") or "")
        if len(sha) != 40 or any(c not in "0123456789abcdef" for c in sha):
            errors.append("fingerprint source commit must be a full lowercase SHA")
        if not isinstance(source.get("production_modules"), dict) or not source["production_modules"]:
            errors.append("fingerprint production-module identity is required")
        if not str(source.get("grid_sha256") or "").startswith("sha256:"):
            errors.append("fingerprint grid identity is required")

    rows = fingerprint.get("rows")
    if not isinstance(rows, list):
        return errors + ["fingerprint rows must be a list"]
    baseline = Counter()
    for pos, value in enumerate(rows, start=1):
        try:
            baseline[identity_key(_identity_from_json(value))] += 1
        except (ValueError, TypeError) as exc:
            errors.append(f"fingerprint row {pos}: {exc}")
    if fingerprint.get("initial_audit_baseline_count") != sum(baseline.values()):
        errors.append("fingerprint initial count does not match its canonical rows")
    if any(count > 1 for count in baseline.values()):
        errors.append("fingerprint contains duplicate canonical identities")

    unauthorized = result["missing_identities"] - baseline
    if unauthorized:
        errors.append(f"{sum(unauthorized.values())} new or changed unclaimed row(s) "
                      "are outside the frozen coverage-debt fingerprint")
    removed = baseline - result["main_identities"]
    if removed:
        errors.append(f"{sum(removed.values())} fingerprinted debt row(s) were removed "
                      "instead of receiving a disposition")
    return errors


def validate_strict(result):
    errors = list(result["errors"])
    if result["missing_rows"]:
        errors.append(f"{result['missing_rows']} main-grid row(s) have no contract or exemption")
    return errors


def summary_markdown(result, mode, errors):
    lines = [
        "# Main-grid disposition summary",
        "",
        "Generated by `check_grid_coverage.py` from canonical binary64 row identities.",
        "",
        f"- mode: `{mode}`",
        f"- total grid rows: {result['total_rows']:,}",
        f"- main-grid rows: {result['main_rows']:,}",
        f"- claimed main-grid rows: {result['claimed_rows']:,}",
        f"- explicitly exempt main-grid rows: {result['exempt_rows']:,}",
        f"- missing main-grid dispositions: {result['missing_rows']:,}",
        "",
    ]
    if mode == "transition":
        lines += [f"> **KNOWN COVERAGE DEBT — {result['missing_rows']:,} unresolved row(s).** "
                  "This is a temporary anti-regression fence, not a contract, exemption, "
                  "or completeness PASS.", ""]
    elif not result["missing_rows"] and not errors:
        lines += ["> **STRICT COVERAGE COMPLETE.** Every main-grid row has a final disposition.", ""]
    if result["missing_groups"]:
        lines += ["| Missing function | Regime | Rows |", "|---|---|---:|"]
        for (function, regime), count in sorted(result["missing_groups"].items()):
            lines.append(f"| {function} | {regime} | {count} |")
        lines.append("")
    if errors:
        lines += ["## Blocking diagnostics", ""] + [f"- {error}" for error in errors] + [""]
    return "\n".join(lines)


def evaluate_paths(grid_path, contracts_path, exemptions_path):
    with open(grid_path, newline="") as f:
        rows = list(csv.DictReader(f))
    with open(contracts_path, newline="") as f:
        contracts = list(csv.DictReader(f))
    exemptions, exemption_errors = _load_exemptions(exemptions_path)
    result = evaluate_coverage(rows, contracts, exemptions)
    result["errors"] = exemption_errors + result["errors"]
    return result


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--mode", choices=("auto", "strict", "transition"), default="auto")
    ap.add_argument("--grid", default=os.path.join(HERE, "probability_accuracy_grid.csv"))
    ap.add_argument("--contracts", default=os.path.join(HERE, "accuracy_contracts.csv"))
    ap.add_argument("--exemptions", default=os.path.join(HERE, "accuracy_row_exemptions.json"))
    ap.add_argument("--fingerprint", default=os.path.join(HERE, "coverage_debt_v1_0_0.json"))
    ap.add_argument("--out", default=os.path.join(HERE, "coverage_summary.md"))
    ap.add_argument("--check-summary", action="store_true")
    ap.add_argument("--write-baseline", action="store_true")
    ap.add_argument("--source-commit")
    args = ap.parse_args()

    result = evaluate_paths(args.grid, args.contracts, args.exemptions)
    if args.write_baseline:
        if result["errors"]:
            raise SystemExit("refusing to write a fingerprint from invalid coverage data")
        if not args.source_commit:
            raise SystemExit("--source-commit is required with --write-baseline")
        root = os.path.dirname(HERE)
        fingerprint = build_fingerprint(root, args.grid, result, args.source_commit,
                                        "danielep71", 22, "v1.0.0")
        with open(args.fingerprint, "w", encoding="utf-8", newline="\n") as f:
            json.dump(fingerprint, f, indent=2, ensure_ascii=False)
            f.write("\n")
        print(f"wrote frozen coverage-debt fingerprint: {args.fingerprint} "
              f"({result['missing_rows']} rows)")

    mode = args.mode
    if mode == "auto":
        mode = "transition" if os.path.exists(args.fingerprint) else "strict"

    if mode == "transition":
        fingerprint, load_errors = _load_json(args.fingerprint, "coverage-debt fingerprint")
        errors = result["errors"] + load_errors
        if not load_errors:
            errors = validate_transition(fingerprint, result)
    else:
        errors = validate_strict(result)

    rendered = summary_markdown(result, mode, errors)
    if args.check_summary:
        try:
            committed = open(args.out, encoding="utf-8").read()
        except OSError as exc:
            errors.append(f"cannot read committed coverage summary: {exc}")
        else:
            if committed != rendered:
                errors.append("coverage_summary.md is stale; regenerate it")
    else:
        with open(args.out, "w", encoding="utf-8", newline="\n") as f:
            f.write(rendered)

    print(f"coverage: total={result['total_rows']} main={result['main_rows']} "
          f"claimed={result['claimed_rows']} exempt={result['exempt_rows']} "
          f"missing={result['missing_rows']} mode={mode}")
    if mode == "transition":
        print(f"KNOWN COVERAGE DEBT: {result['missing_rows']} unresolved row(s)")
    for error in errors:
        print("  - " + error)
    if errors:
        raise SystemExit(1)
    if mode == "transition":
        print("transition guard satisfied; completeness remains unresolved")
    else:
        print("strict coverage complete")


if __name__ == "__main__":
    main()
