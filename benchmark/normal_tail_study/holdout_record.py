"""
One-shot holdout exposure record of the #72 normal-tail study (#80).

PREREGISTRATION.md §5 evaluates the holdout H once, after the validation set V
has passed and has been written. results.json carries that exposure
explicitly, in its `holdout_exposure` block, and validate_kernel.py enforces
it through this module:

  status               meaning
  NOT_EVALUATED        H has not been evaluated; acceptance.A6 is absent.
  EVALUATION_STARTED   a first evaluation began and did not record A6. H
                       counts as exposed and is never evaluated again.
  EVALUATED            H was evaluated once; acceptance.A6 holds the result
                       and a6_sha256 its canonical SHA-256.

The record also binds the identity it was made under: the SHA-256 of the
coefficient file (the candidate), the byte length and SHA-256 of the
preregistration (the design; later amendments may only append to it), and the
holdout definition with the SHA-256 of its 15,000 seeded points (the input).
Generating those points is not an evaluation: no kernel and no oracle is
computed on them.

Every mode reads and checks the existing record before it computes or writes
anything. A missing, malformed or inconsistent record is refused, never
reset. A write goes through an exclusively created temporary file and
os.replace, so a crash leaves either the old or the new file; a temporary file
left behind by a killed run blocks every mode until someone inspects it.

  init              new study only: results.json must not exist. Writes V with
                    status NOT_EVALUATED.
  verify            read-only replay: recomputes everything except H and
                    compares it with results.json. Never evaluates H and never
                    writes; the recorded A6 is reported as retained, not as a
                    new measurement.
  regenerate        recomputes everything except H and rewrites results.json,
                    carrying holdout_exposure and acceptance.A6 over unchanged
                    and appending the reason to `regenerations`.
  first_evaluation  refused unless the status is NOT_EVALUATED. Writes V with
                    EVALUATION_STARTED, evaluates H, then writes EVALUATED.

Standard library only, so the fixtures in test_holdout_record.py run on
synthetic studies without mpmath.
"""
from __future__ import annotations

import hashlib
import json
import os
import random
import struct
import time
from dataclasses import dataclass
from typing import Any, Callable

NOT_EVALUATED = "NOT_EVALUATED"
STARTED = "EVALUATION_STARTED"
EVALUATED = "EVALUATED"
STATES = {NOT_EVALUATED: 0, STARTED: 1, EVALUATED: 1}   # status -> exposures

PASS, FAIL, PENDING = "PASS", "FAIL", "PENDING_HOLDOUT"

OK, FAILED, USAGE, REFUSED = 0, 1, 2, 3

# Key order of results.json, so a regeneration rewrites the file in place.
ORDER = ("study", "preregistration", "coefficients_sha256", "status", "problems", "generated_utc",
         "versions", "bound", "oracle", "amendments", "acceptance", "bands", "seams", "predictions",
         "holdout_559_inspected", "runtime_s", "holdout_exposure", "regenerations")
# Fields a replay does not reproduce: when and where it ran, and the regeneration log.
VOLATILE = ("generated_utc", "versions", "runtime_s", "regenerations")
EXPOSURE_ORDER = ("status", "exposures", "started_utc", "evaluated_utc", "candidate", "design", "holdout",
                  "a6_sha256", "recorded_by", "note", "code_committed_with_record")
HOLDOUT_PROBLEM = "A6 "                 # every holdout problem starts with this


class StateError(Exception):
    """The existing record does not allow the requested mode."""


@dataclass
class Study:
    output: str
    preregistration: str
    coefficients: str
    holdout_definition: list[dict[str, Any]]
    compute: Callable[[], dict[str, Any]]
    evaluate_holdout: Callable[[list[float]], dict[str, Any]]
    log: Callable[[str], None] = print


# ---------------------------------------------------------------------------
# Identity
# ---------------------------------------------------------------------------

def canonical(obj: Any) -> bytes:
    return json.dumps(obj, sort_keys=True, separators=(",", ":")).encode("utf-8")


def digest(obj: Any) -> str:
    return hashlib.sha256(canonical(obj)).hexdigest()


def normalise(obj: Any) -> Any:
    """The value as results.json stores it (tuples become lists)."""
    return json.loads(json.dumps(obj))


def file_sha256(path: str) -> str:
    with open(path, "rb") as f:
        return hashlib.sha256(f.read()).hexdigest()


def holdout_points(definition: list[dict[str, Any]]) -> list[float]:
    zs: list[float] = []
    for part in definition:
        rnd = random.Random(part["seed"])
        zs += [rnd.uniform(part["low"], part["high"]) for _ in range(part["count"])]
    return zs


def points_sha256(zs: list[float]) -> str:
    return hashlib.sha256(b"".join(struct.pack(">d", z) for z in zs)).hexdigest()


def identity(study: Study) -> dict[str, Any]:
    with open(study.preregistration, "rb") as f:
        design = f.read()
    zs = holdout_points(study.holdout_definition)
    return {
        "candidate": {"coefficients_sha256": file_sha256(study.coefficients)},
        "design": {"preregistration": os.path.basename(study.preregistration), "bytes": len(design),
                   "sha256": hashlib.sha256(design).hexdigest()},
        "holdout": {"definition": normalise(study.holdout_definition), "points": len(zs),
                    "points_sha256": points_sha256(zs)},
    }


def check_identity(exposure: dict[str, Any], study: Study) -> list[str]:
    """Differences between the recorded identity and the files on disk."""
    now = identity(study)
    out = []
    recorded = exposure["candidate"].get("coefficients_sha256")
    if recorded != now["candidate"]["coefficients_sha256"]:
        out.append(f"candidate changed: the record binds coefficients {recorded}, "
                   f"{os.path.basename(study.coefficients)} is {now['candidate']['coefficients_sha256']}. "
                   "The recorded exposure belongs to the recorded candidate")
    design = exposure["design"]
    with open(study.preregistration, "rb") as f:
        text = f.read()
    size = design.get("bytes")
    if (not isinstance(size, int) or len(text) < size
            or hashlib.sha256(text[:size]).hexdigest() != design.get("sha256")):
        out.append(f"{os.path.basename(study.preregistration)} no longer begins with the {size} bytes "
                   f"recorded (SHA-256 {design.get('sha256')}): the frozen text was changed, not amended "
                   "by appending")
    if exposure["holdout"] != now["holdout"]:
        out.append(f"holdout definition changed: recorded {exposure['holdout']}, now {now['holdout']}")
    return out


# ---------------------------------------------------------------------------
# Record consistency
# ---------------------------------------------------------------------------

def holdout_problems(exposure: dict[str, Any], a6: dict[str, Any] | None) -> list[str]:
    if exposure["status"] == STARTED:
        return [f"A6 not recorded: the holdout evaluation started at {exposure.get('started_utc')} has no "
                "recorded result; the holdout counts as exposed"]
    if exposure["status"] == EVALUATED and a6 is not None and a6.get("pass") is not True:
        return ["A6 failed"]
    return []


def derive_status(problems: list[str], exposure_status: str) -> str:
    if problems or exposure_status == STARTED:
        return FAIL
    return PASS if exposure_status == EVALUATED else PENDING


def check_record(result: Any) -> list[str]:
    """Every way the record is malformed or contradicts itself; [] if none."""
    if not isinstance(result, dict):
        return ["the top level is not a JSON object"]
    out = []
    for key, kind in (("status", str), ("problems", list), ("acceptance", dict),
                      ("coefficients_sha256", str), ("holdout_exposure", dict)):
        if not isinstance(result.get(key), kind):
            out.append(f"'{key}' is missing or is not a {kind.__name__}")
    if out:
        if "A6" in (result.get("acceptance") or {}) and not isinstance(result.get("holdout_exposure"), dict):
            out.append("acceptance.A6 is present without a holdout_exposure record")
        return out
    ex, acc = result["holdout_exposure"], result["acceptance"]
    status = ex.get("status")
    if status not in STATES:
        return [f"holdout_exposure.status {status!r} is not one of {', '.join(STATES)}"]
    for key in ("candidate", "design", "holdout"):
        if not isinstance(ex.get(key), dict):
            out.append(f"holdout_exposure.{key} is missing or is not an object")
    if out:
        return out
    if ex["candidate"].get("coefficients_sha256") != result["coefficients_sha256"]:
        out.append("holdout_exposure.candidate does not match coefficients_sha256")
    if type(ex.get("exposures")) is not int or ex["exposures"] != STATES[status]:
        out.append(f"holdout_exposure.exposures is {ex.get('exposures')!r}; status {status} requires "
                   f"{STATES[status]}")
    a6 = acc.get("A6")
    if status == NOT_EVALUATED:
        if "A6" in acc:
            out.append("acceptance.A6 is present but holdout_exposure.status is NOT_EVALUATED")
        for key in ("started_utc", "evaluated_utc", "a6_sha256"):
            if key in ex:
                out.append(f"holdout_exposure.{key} is present but the status is NOT_EVALUATED")
    else:
        if not isinstance(ex.get("started_utc"), str):
            out.append(f"holdout_exposure.started_utc is missing for status {status}")
    if status == STARTED:
        if "A6" in acc:
            out.append("acceptance.A6 is present but holdout_exposure.status is EVALUATION_STARTED")
        if "a6_sha256" in ex:
            out.append("holdout_exposure.a6_sha256 is present but the status is EVALUATION_STARTED")
    if status == EVALUATED:
        if not isinstance(a6, dict):
            out.append("holdout_exposure.status is EVALUATED but acceptance.A6 is missing")
            return out
        if a6.get("evaluated_once") is not True:
            out.append("acceptance.A6.evaluated_once is not true")
        if not isinstance(a6.get("pass"), bool):
            out.append("acceptance.A6.pass is not a Boolean")
        if ex.get("a6_sha256") != digest(a6):
            out.append(f"acceptance.A6 does not match holdout_exposure.a6_sha256 ({ex.get('a6_sha256')}; "
                       f"the block hashes to {digest(a6)})")
    problems = result["problems"]
    base = [p for p in problems if not str(p).startswith(HOLDOUT_PROBLEM)]
    if problems != base + holdout_problems(ex, a6 if isinstance(a6, dict) else None):
        out.append(f"problems {problems} do not match the exposure record")
    elif result["status"] != derive_status(problems, status):
        out.append(f"status {result['status']} does not follow from the problems and exposure "
                   f"(expected {derive_status(problems, status)})")
    for entry in result.get("regenerations", []):
        if not isinstance(entry, dict) or entry.get("holdout_evaluated") is not False:
            out.append(f"regeneration entry {entry!r} does not record holdout_evaluated: false")
    return out


def assemble(body: dict[str, Any], exposure: dict[str, Any], a6: dict[str, Any] | None = None,
             regenerations: list[dict[str, Any]] | None = None) -> dict[str, Any]:
    """A results.json record from the non-holdout body and the exposure."""
    merged = dict(body)
    problems = [p for p in body["problems"]] + holdout_problems(exposure, a6)
    merged["problems"] = problems
    merged["status"] = derive_status(problems, exposure["status"])
    acceptance = dict(body["acceptance"])
    if a6 is not None:
        acceptance["A6"] = a6
    merged["acceptance"] = acceptance
    merged["holdout_exposure"] = ordered(exposure, EXPOSURE_ORDER)
    if regenerations:
        merged["regenerations"] = regenerations
    return ordered(merged, ORDER)


def ordered(record: dict[str, Any], order: tuple[str, ...]) -> dict[str, Any]:
    out = {k: record[k] for k in order if k in record}
    out.update({k: v for k, v in record.items() if k not in out})
    return out


def differences(expected: Any, recorded: Any, path: str = "") -> list[str]:
    if isinstance(expected, dict) and isinstance(recorded, dict):
        out = []
        for key in list(expected) + [k for k in recorded if k not in expected]:
            sub = f"{path}.{key}" if path else str(key)
            if key not in recorded:
                out.append(f"{sub}: recomputed but not recorded")
            elif key not in expected:
                out.append(f"{sub}: recorded but not recomputed")
            else:
                out += differences(expected[key], recorded[key], sub)
        return out
    if isinstance(expected, list) and isinstance(recorded, list) and len(expected) == len(recorded):
        out = []
        for i, (a, b) in enumerate(zip(expected, recorded)):
            out += differences(a, b, f"{path}[{i}]")
        return out
    if expected == recorded and type(expected) is type(recorded):
        return []
    return [f"{path}: recorded {str(recorded)[:80]}, recomputed {str(expected)[:80]}"]


# ---------------------------------------------------------------------------
# Reading and writing
# ---------------------------------------------------------------------------

def temporary_path(output: str) -> str:
    return os.path.join(os.path.dirname(output), "." + os.path.basename(output) + ".tmp")


def read_state(study: Study) -> tuple[dict[str, Any] | None, str | None]:
    """The checked record and the SHA-256 of its bytes, or (None, None) if absent."""
    tmp = temporary_path(study.output)
    if os.path.lexists(tmp):
        raise StateError(f"{tmp} exists: an earlier write was interrupted. {study.output} was not replaced "
                         "by it. Inspect both, keep the one that holds the exposure record (git has the "
                         "committed one), delete the temporary file and run again")
    if not os.path.exists(study.output):
        return None, None
    with open(study.output, "rb") as f:
        raw = f.read()
    try:
        result = json.loads(raw.decode("utf-8"))
    except (UnicodeDecodeError, ValueError) as e:
        raise StateError(f"{study.output} is not valid JSON ({e}); it may be a partial write. Restore it "
                         "from git; it is never reset") from e
    problems = check_record(result)
    if problems:
        raise StateError(f"{study.output} holds a missing, malformed or inconsistent holdout exposure "
                         "record; it is never reset:\n  - " + "\n  - ".join(problems))
    return result, hashlib.sha256(raw).hexdigest()


def require_identity(state: dict[str, Any], study: Study) -> None:
    problems = check_identity(state["holdout_exposure"], study)
    if problems:
        raise StateError("the record was made under a different identity:\n  - " + "\n  - ".join(problems))


def write_atomic(study: Study, result: dict[str, Any], expected: str | None) -> str:
    """Replace the output, provided it still has the bytes this run read."""
    current = file_sha256(study.output) if os.path.exists(study.output) else None
    if current != expected:
        raise StateError(f"{study.output} changed while this run was computing (another run?); "
                         "it was not overwritten")
    data = (json.dumps(result, indent=1) + "\n").encode("utf-8")
    tmp = temporary_path(study.output)
    try:
        fd = os.open(tmp, os.O_WRONLY | os.O_CREAT | os.O_EXCL, 0o644)
    except FileExistsError as e:
        raise StateError(f"{tmp} appeared while this run was computing (another run?); "
                         f"{study.output} was not overwritten") from e
    try:
        with os.fdopen(fd, "wb") as f:
            f.write(data)
            f.flush()
            os.fsync(f.fileno())
        os.replace(tmp, study.output)
    except BaseException:
        if os.path.lexists(tmp):
            os.unlink(tmp)
        raise
    return hashlib.sha256(data).hexdigest()


def utc_now() -> str:
    return time.strftime("%Y-%m-%dT%H:%M:%SZ", time.gmtime())


def compute(study: Study) -> tuple[dict[str, Any], float]:
    started = time.time()
    body = normalise(study.compute())
    if any(str(p).startswith(HOLDOUT_PROBLEM) for p in body["problems"]) or "A6" in body["acceptance"]:
        raise StateError("the non-holdout computation returned a holdout result; nothing was written")
    return body, started


def finish(result: dict[str, Any], started: float) -> dict[str, Any]:
    result["generated_utc"] = utc_now()
    result["runtime_s"] = round(time.time() - started, 1)
    return ordered(result, ORDER)


def describe(exposure: dict[str, Any]) -> str:
    if exposure["status"] == NOT_EVALUATED:
        return "holdout not yet evaluated"
    when = exposure.get("started_utc")
    if exposure["status"] == STARTED:
        return f"holdout evaluation started at {when} and did not finish; the holdout counts as exposed"
    return f"holdout evaluated once, in the run started at {when}"


# ---------------------------------------------------------------------------
# Modes
# ---------------------------------------------------------------------------

def init(study: Study, reason: str | None = None) -> int:
    state, _ = read_state(study)
    if state is not None:
        raise StateError(f"{study.output} already exists; --init is for a new study only. "
                         "Use --verify or --regenerate")
    body, started = compute(study)
    exposure = {"status": NOT_EVALUATED, "exposures": 0, **identity(study)}
    result = finish(assemble(body, exposure), started)
    write_atomic(study, result, None)
    study.log(f"INITIALISED: {result['status']}; {describe(exposure)}; {len(result['problems'])} problem(s)")
    for p in result["problems"]:
        study.log(f"  - {p}")
    return OK if result["status"] == PENDING else FAILED


def verify(study: Study, reason: str | None = None) -> int:
    state, _ = read_state(study)
    if state is None:
        raise StateError(f"{study.output} does not exist. --verify never creates a record; restore it from git")
    require_identity(state, study)
    exposure = state["holdout_exposure"]
    body, _ = compute(study)
    expected = normalise(assemble(body, exposure, state["acceptance"].get("A6")))
    diffs = differences({k: v for k, v in expected.items() if k not in VOLATILE},
                        {k: v for k, v in state.items() if k not in VOLATILE})
    study.log("REPLAY (verification only): the holdout was not evaluated by this run, and nothing was written. "
              f"Record: {describe(exposure)}.")
    if "A6" in state["acceptance"]:
        study.log(f"  acceptance.A6 is the retained record of that evaluation (SHA-256 {exposure['a6_sha256']}), "
                  "not a new measurement.")
    if body.get("versions") != state.get("versions"):
        study.log(f"  note: recorded with {state.get('versions')}, replayed with {body.get('versions')}")
    if diffs:
        study.log(f"NOT REPRODUCED: {len(diffs)} recorded value(s) differ from the replay:")
        for d in diffs[:40]:
            study.log(f"  - {d}")
        return FAILED
    if exposure["status"] == STARTED:
        study.log("INCOMPLETE: the replay matches, but the exposure record has no A6 result")
        return FAILED
    study.log(f"VERIFIED: every recorded non-holdout value reproduces; recorded status {state['status']} "
              "retained, not re-established by a new holdout evaluation")
    return OK


def regenerate(study: Study, reason: str | None = None) -> int:
    if not reason or not reason.strip():
        raise StateError("--regenerate needs --reason TEXT, recorded in results.json")
    state, raw = read_state(study)
    if state is None:
        raise StateError(f"{study.output} does not exist. --regenerate needs the existing exposure record; "
                         "restore it from git (or use --init for a new study)")
    require_identity(state, study)
    exposure, a6 = state["holdout_exposure"], state["acceptance"].get("A6")
    body, started = compute(study)
    regenerations = list(state.get("regenerations", [])) + [
        {"utc": utc_now(), "reason": reason.strip(), "holdout_evaluated": False}]
    result = finish(assemble(body, exposure, a6, regenerations), started)
    if result["holdout_exposure"] != exposure or result["acceptance"].get("A6") != a6:
        raise StateError("internal error: the exposure record would change; nothing was written")
    write_atomic(study, result, raw)
    study.log(f"REGENERATED: {result['status']}. Everything except the holdout was recomputed and rewritten; "
              f"the holdout was not evaluated. holdout_exposure and acceptance.A6 were carried over unchanged "
              f"({describe(exposure)}).")
    for p in result["problems"]:
        study.log(f"  - {p}")
    return OK if result["status"] in (PASS, PENDING) else FAILED


def first_evaluation(study: Study, reason: str | None = None) -> int:
    state, raw = read_state(study)
    if state is None:
        raise StateError(f"{study.output} does not exist. A first evaluation needs the NOT_EVALUATED record "
                         "that --init writes; a missing record is never taken to mean the holdout is untouched")
    exposure = state["holdout_exposure"]
    if exposure["status"] != NOT_EVALUATED:
        raise StateError(f"second first-evaluation attempt refused: {describe(exposure)} "
                         f"(holdout_exposure.status {exposure['status']}, exposures {exposure['exposures']}). "
                         "PREREGISTRATION.md §5 allows one evaluation. Use --verify to replay the non-holdout "
                         "checks, or --regenerate to rewrite them with the recorded exposure kept")
    require_identity(state, study)
    ident = identity(study)
    body, started = compute(study)
    if body["problems"]:
        result = finish(assemble(body, {"status": NOT_EVALUATED, "exposures": 0, **ident}), started)
        write_atomic(study, result, raw)
        study.log("FAIL: the holdout was not evaluated, because a condition before it failed (§6):")
        for p in result["problems"]:
            study.log(f"  - {p}")
        return FAILED
    exposure = {"status": STARTED, "exposures": 1, "started_utc": utc_now(), **ident,
                "recorded_by": "validate_kernel.py --first-evaluation"}
    raw = write_atomic(study, finish(assemble(body, exposure), started), raw)
    a6 = normalise(study.evaluate_holdout(holdout_points(study.holdout_definition)))
    exposure = {**exposure, "status": EVALUATED, "evaluated_utc": utc_now(), "a6_sha256": digest(a6)}
    result = finish(assemble(body, exposure, a6), started)
    write_atomic(study, result, raw)
    study.log(f"{result['status']}: the holdout was evaluated once; the exposure is recorded as EVALUATED")
    for p in result["problems"]:
        study.log(f"  - {p}")
    return OK if result["status"] == PASS else FAILED


MODES: dict[str, Callable[[Study, str | None], int]] = {
    "init": init, "verify": verify, "regenerate": regenerate, "first_evaluation": first_evaluation}


def run(study: Study, mode: str, reason: str | None = None) -> int:
    def fingerprint() -> str | None:
        return file_sha256(study.output) if os.path.isfile(study.output) else None

    before = fingerprint()
    try:
        return MODES[mode](study, reason)
    except StateError as e:
        study.log(f"REFUSED ({mode.replace('_', '-')}): {e}")
        after = fingerprint()
        if after == before:
            study.log(f"  {study.output} is unchanged (SHA-256 {after}).")
        else:
            study.log(f"  {study.output} changed during this run: SHA-256 {before} -> {after}.")
        return REFUSED
