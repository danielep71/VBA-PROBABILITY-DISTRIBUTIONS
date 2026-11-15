# User manual, v1.0.0 — source

The manual is written in Markdown, one chapter per file, and built into a PDF.
This directory belongs to v1.0.0: after that release it is not edited, and a
later revision of the manual goes in a directory of its own version.

| File | Content |
| --- | --- |
| `00-front-matter.md` to `13-quick-reference.md` | The chapters, in order |
| `manual.css` | Page size, typography, running footer and page numbers |
| `build_manual.py` | The build and its checks |
| `checks/example_values.py` | Recomputes every numeric result the manual shows |

## Build

```text
python3 build_manual.py --check
```

It writes `build/manual.html`,
`build/VBA-Probability-Distributions-User-Manual-v1.0.0.pdf` and
`build/build_record.json`, which holds the PDF's SHA-256 and the versions of
the two tools. `build/` is not committed.

Tools:

- pandoc 3.1 or later, for Markdown to HTML with MathML and the contents;
- a Chromium or Chrome binary, version 131 or later, for HTML to PDF. The
  script looks for `chromium`, `chromium-browser`, `google-chrome` and
  `chrome`; set `CHROME` to a path to choose one.

The PDF is not reproducible byte for byte: the browser writes the creation
time into it. The checksum to publish is the one of the file attached to the
release.

## Checks

`build_manual.py --check` fails when a `K_STATS_` name in a chapter is not in
`docs/PUBLIC_API.txt`, or when a call written with its arguments has a number
of arguments its declaration does not accept.
`tests/test_user_manual.py` runs the same check with the unit tests.

`python3 checks/example_values.py` recomputes the expected results of chapters
4, 7 and 8 with mpmath at 30 digits, independently of the library. It needs
mpmath and exits 0 when it has printed every value.

These checks do not run Excel. The functions the examples call are covered by
the Excel regression and the accuracy evidence of the release; the manual's
own formulas were not typed into a workbook by a test.

## Before the release

- Rebuild on the release commit and read every page: clipping, broken
  formulas, a table split badly.
- Compare chapter 10 with `benchmark/accuracy_guarantees.csv` and
  `benchmark/numerical_limitations.csv` as they are on that commit.
- Attach the PDF to the release and add its SHA-256 to the checksum inventory.
