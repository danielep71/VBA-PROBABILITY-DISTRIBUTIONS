# #307 Gamma routing study

#307: since #242, every public Gamma P and Q call runs the double-double
kernel, at 5 ms to 170 ms per call. The binary64 routines it replaced take 2
to 39 microseconds. This study looks for a routing rule: binary64 wherever it
meets the contracts, double-double only where it does not.

Step 1 is a map of the binary64 route in Excel. It is a fitting set; the rule
derived from it will be preregistered and validated on untouched points.

| File | Role |
| --- | --- |
| [`MAP_DESIGN.md`](MAP_DESIGN.md) | The frozen map design: shapes, arguments, references and scoring |
| `design.py` | The map's inputs |
| `generate_binary64_map.py` | Builds `binary64_map_grid.csv` with high-precision references; `--check` proves the committed grid is the design |
| `M_STATS_PROBDIST_GBR.bas` | `Export_GammaBinary64Map`: calls the binary64 series and continued fraction at every point. The runner exports it as the study grid `gbr_binary64_map` |
| `analyze_binary64_map.py` | Scores the pre-#242 route against the governing contract at every point |
| `test_binary64_map.py` | Proves the grid is the design and the analyzer forms the pre-#242 complement; it runs in the evidence-tool suite |
| [`MAP_RESULT.md`](MAP_RESULT.md), `map_results.json` | The map's result (run 42783017314) |
| `analyze_exact_deviance.py` | Diagnosis: the binary64 route with an exact deviance |
