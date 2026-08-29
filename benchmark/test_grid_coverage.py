"""Fixtures for the strict and transitional main-grid coverage checker."""
import copy

import check_grid_coverage as C

fails = []


def check(condition, message):
    if not condition:
        fails.append(message)


def row(function="F_InverseCumulative", regime="envelope_domain",
        arg1="0.9", arg2="1000000", evidence_set="main grid"):
    return {
        "function": function, "arg1": arg1, "arg2": arg2,
        "arg3": "3", "arg4": "", "regime": regime,
        "evidence_set": evidence_set,
    }


def contract(function="F_InverseCumulative", regime="envelope_domain"):
    return {"contract_id": function + "." + regime,
            "function": function, "regime": regime, "status": "active"}


def fingerprint_for(result):
    return {
        "schema_version": C.FINGERPRINT_SCHEMA,
        "identity_schema": C.IDENTITY_SCHEMA,
        "owner": "danielep71",
        "issue": 22,
        "milestone_expiry": "v1.0.0",
        "source_identity": {
            "commit_sha": "b" * 40,
            "production_modules": {"src/a.bas": "sha256:" + "a" * 64},
            "grid_sha256": "sha256:" + "b" * 64,
        },
        "initial_audit_baseline_count": result["missing_rows"],
        "rows": copy.deepcopy(result["missing_identity_objects"]),
    }


# 1. clean strict grid: every main row is claimed; non-main study rows are ignored
clean = C.evaluate_coverage([row(), row(evidence_set="study")], [contract()], [])
check(C.validate_strict(clean) == [], "clean strict grid passes")
check(clean["main_rows"] == 1 and clean["claimed_rows"] == 1,
      "non-main study row excluded")

# 2. lost contract is a strict missing disposition
lost = C.evaluate_coverage([row()], [], [])
check(any("no contract or exemption" in x for x in C.validate_strict(lost)),
      "lost contract fails strict mode")

# 3. a complete named explicit exemption is a final disposition
ident = C.row_identity(row())
exemption = {
    "exemption_id": "EX-1", "identity": ident, "reason": "structural invariant",
    "owner": "danielep71", "review_policy": "review at every release", "status": "active",
}
exempt = C.evaluate_coverage([row()], [], [exemption])
check(C.validate_strict(exempt) == [] and exempt["exempt_rows"] == 1,
      "explicit exemption passes strict mode")

# 4. stale exemptions and unknown statuses fail
stale = C.evaluate_coverage([row()], [contract()], [exemption])
check(any("stale exemption" in x for x in C.validate_strict(stale)),
      "stale exemption fails")
bad_status = copy.deepcopy(exemption); bad_status["status"] = "temporary"
bad = C.evaluate_coverage([row()], [], [bad_status])
check(any("unknown status" in x for x in C.validate_strict(bad)),
      "unknown exemption status fails")

# 5. exact frozen transition baseline is accepted only as visible known debt
baseline = C.evaluate_coverage([row()], [], [])
frozen = fingerprint_for(baseline)
check(C.validate_transition(frozen, baseline) == [],
      "exact transition baseline is fenced")

# 6. a new row outside the fingerprint fails
new_row = row(arg1="0.8")
grown = C.evaluate_coverage([row(), new_row], [], [])
check(any("new or changed" in x for x in C.validate_transition(frozen, grown)),
      "new unclaimed row fails transition")

# 7. changed arguments fail as a new identity and a removed baseline identity
changed_arg = C.evaluate_coverage([row(arg2="1000001")], [], [])
arg_errors = C.validate_transition(frozen, changed_arg)
check(any("new or changed" in x for x in arg_errors), "changed arguments fail transition")

# 8. changed regime fails as a new identity
changed_regime = C.evaluate_coverage([row(regime="all")], [], [])
check(any("new or changed" in x for x in C.validate_transition(frozen, changed_regime)),
      "changed regime fails transition")

# 9. a resolved row is allowed only when the same identity remains and gains a contract
resolved = C.evaluate_coverage([row()], [contract()], [])
check(C.validate_transition(frozen, resolved) == [],
      "resolved fingerprint row is allowed")

# 10. deleting a debt row is not accepted as resolution
deleted = C.evaluate_coverage([], [], [])
check(any("were removed" in x for x in C.validate_transition(frozen, deleted)),
      "deleted debt row fails transition")

# 11. corrupted fingerprint metadata and row identity fail closed
corrupt = copy.deepcopy(frozen); corrupt["initial_audit_baseline_count"] = 999
check(any("initial count" in x for x in C.validate_transition(corrupt, baseline)),
      "corrupted fingerprint count fails")
corrupt_row = copy.deepcopy(frozen); corrupt_row["rows"][0]["arg1_bits"] = "not-bits"
check(any("invalid binary64" in x for x in C.validate_transition(corrupt_row, baseline)),
      "corrupted fingerprint identity fails")

if fails:
    print("FAIL: main-grid coverage fixtures")
    for failure in fails:
        print("  - " + failure)
    raise SystemExit(1)
print("PASS: main-grid coverage checker (strict, exemption, baseline, new/change, "
      "resolved/deleted, corrupted fingerprint, non-main row)")
