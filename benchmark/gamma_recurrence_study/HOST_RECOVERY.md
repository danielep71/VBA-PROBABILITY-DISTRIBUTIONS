# Host failure and recovery — 2026-11-11

Run 43030484382 at 2c82328753b413f9da8c5a6a6a281914051c464c stopped
at sample 193/216: Excel COM returned RPC failure 0x800706BE, followed by
cleanup failure 0x800706BA. The report preserves both errors. The artifact ZIP
SHA-256 was verified before extraction; metadata and all 193 partial samples
are retained in results/run_43030484382_incomplete.json.

All completed samples match baseline bits, but the incomplete run is FAIL and
is not used for candidate selection or a speedup claim. Host timing is visibly
variable, including controls whose executed production path was unchanged.
The exact reason Excel stopped responding is not established.

Before repeating the complete experiment, bound each Excel instance to one
round (72 source reloads rather than up to 216 in one instance), close it and
start a fresh instance for the next round. This is a robustness mitigation,
not a proven diagnosis of the COM failure. All points, mode order, warmups,
call counts, numerical transformations, rejection criteria and timing target
remain as preregistered. Re-run all three rounds; do not splice partial samples.

Use a dedicated measurement branch so updating the review PR does not launch
another benchmark behind the ordinary Excel regression. No production code or
accuracy manifest changes. A successful run remains preliminary performance
evidence and needs confirmation plus full numerical validation before promotion.
