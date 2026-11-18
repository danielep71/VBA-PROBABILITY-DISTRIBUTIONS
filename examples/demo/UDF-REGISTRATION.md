# Function Wizard registration

`M_STATS_PROBDIST_REGISTER.bas` adds descriptions and argument help for the
complete worksheet surface in `docs/PUBLIC_API.txt`, under the Excel category
**Probability Distributions**. It adapts the lifecycle of
[K-PRICING's registration module](https://github.com/danielep71/K-PRICING/blob/main/src/modules/KPR_REGISTER_PUBLIC_UDFS.bas).

The Explorer's `Workbook_Open` calls `PROB_Register_PublicUDFs` before its
existing initialization. Macros must be enabled. For an existing workbook,
import this module and add that call to the existing open event, retaining its
other statements. No new references or Trust Center changes are required.

The module's `LoadManifest` is the single source of Function Wizard help text.
Names and argument order are checked against `docs/PUBLIC_API.txt` by
`tests/test_demo_registration.py`; descriptions are bounded to Excel's limits.
Help distinguishes rates from scales, log-space from arithmetic parameters,
upper from lower tails, and discrete failure counts. Optional `Status` is a
VBA diagnostic output and should be omitted in worksheet formulas.

## Entry points

- `PROB_Register_PublicUDFs`: repeatable registration, returning accepted calls.
- `PROB_Register_ClearPublicUDFs`: blank help and return entries to User Defined.
  Excel cannot unregister functions while their VBA project remains loaded.
- `PROB_Register_LastCount` and `PROB_Register_LastReport`: last accepted count
  and per-function failures. Infrastructure failure does not stop calculation.
- `PROB_Register_Manifest*`: inspect names, argument counts and argument help.

`Option Private Module` keeps these infrastructure entry points out of the
worksheet function list. Call them from the Immediate window or with the
workbook-qualified `Application.Run`. Registration does not select, activate,
recalculate or change events/alerts/calculation mode. It restores the incoming
Saved flag and attempts every record independently. A bare-name retry is only
allowed while ThisWorkbook is active; otherwise a failed qualified call is
reported instead of risking a different workbook's same-named function.

## Validation scope

`ci/Test-UdfRegistration.ps1` checks automatic startup, complete metadata,
repeat register/clear, Saved flags, inactive-workbook isolation and recovery,
and a numerical smoke test. MacroOptions has no metadata read-back API:
accepted calls are host execution evidence, not a visual inspection of every
Function Wizard entry. The Explorer workflow also runs its existing workbook,
calculation, Ribbon and save/reopen checks.

This implements registration for the demo and provides an importable module.
It does not complete issue #329's separate XLAM packaging/release scope.

## Development evidence — 2026-11-17

Source `5ab14b156d48c7805fb458de654ad3e9ea9f25ae` passed
[Distribution Explorer run 44179591754](https://github.com/danielep71/VBA-PROBABILITY-DISTRIBUTIONS/actions/runs/44179591754)
on Excel 16.0 build 20710 x64. All API entries were accepted at startup,
repeated registration/clear passed, and registration while another workbook
was active accepted every entry without changing that workbook or application
state. The full Explorer workflow also passed. Its tested workbook digest was
`826dd0106e6f6ea43cd0cf44f9491b2ba4472d4d4c1427973ea8feee0238c93e`.
A later edit corrects only a copied return-count comment, with no executable
change. The module's API coverage test also passes after that edit.

The user-selected workbook was updated separately, preserving every package
part except `xl/vbaProject.bin`, all existing module source except one inserted
startup call, and its existing Explorer initialization. Its source digest is
`e9008d8fe4b8390659ed6f8655c2c63688774a2dabb0b4322537b3433c5041bb`;
the delivered digest is
`b225bfcf5b2b33f5a54094b6ac8f37093626426a868945dca9c181362c4a06c9`.
That patch has ZIP, extracted-source and unchanged-part checks. The native
Excel evidence above applies to the repository-built workbook and registration
module; the separately patched user workbook was not reopened on that host.
