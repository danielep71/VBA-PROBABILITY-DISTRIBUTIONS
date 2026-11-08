#!/usr/bin/env python3
"""Render a decision guide from editable text; reconcile the implemented inventory."""
import argparse
from html import escape
from pathlib import Path

from generate_distribution_map import CONTINUOUS, DISCRETE, validate_api

ROOT = Path(__file__).resolve().parents[2]
SVG = ROOT / 'assets/choosing-distribution.svg'
PNG = ROOT / 'assets/choosing-distribution.png'
# Question, candidate surfaces, support, assumption lines. These are alternatives,
# not an automatic fitting algorithm or a numerical certification statement.
DISCRETE_ROUTES = [
    ('Successes in a fixed number of trials?', ('Binomial',), 'k = 0, …, n',
     ('Independent Bernoulli trials; fixed n and p.', 'Changing p or dependence needs another model.')),
    ('Sampling a finite population without replacement?', ('Hypergeometric',), 'max(0, n − N + K) ≤ k ≤ min(n, K)',
     ('Uniform random sample: size n from N, K successes.', 'The population composition is fixed.')),
    ('Events over a fixed exposure?', ('Poisson',), 'k = 0, 1, 2, …',
     ('Independent increments and a constant event rate.', 'Check exposure, clustering and mean versus variance.')),
    ('Failures before success number r?', ('Geometric', 'Negative Binomial'), 'k = 0, 1, 2, …',
     ('Independent Bernoulli trials with constant p.', 'r = 1: Geometric; r ≥ 1: Negative Binomial.')),
    ('Equally likely consecutive integers?', ('Discrete Uniform',), 'k = a, a + 1, …, b',
     ('Every integer in the finite range has equal mass.', 'Bounds must describe the actual sample space.')),
]
CONTINUOUS_ROUTES = [
    ('Real-valued and approximately symmetric?', ('Normal', 'Student t', 'Standard Normal'), 'x ∈ ℝ',
     ('Normal-shaped or heavier, t-shaped tails?', 'Standard Normal is the fixed N(0,1) surface.')),
    ('Positive quantity with a Normal log?', ('Lognormal',), 'x > 0',
     ('Check log-scale fit and multiplicative variation.', 'Zeros and negative values do not fit this model.')),
    ('Nonnegative duration or lifetime?', ('Exponential', 'Gamma', 'Weibull'), 'x ≥ 0',
     ('Constant hazard: Exponential; flexible shape: Gamma.', 'Monotone hazard: Weibull; account for censoring.')),
    ('Continuous proportion or bounded measurement?', ('Beta', 'Uniform'), '0 ≤ x ≤ 1  /  a ≤ x ≤ b',
     ('Beta: flexible shape on [0,1]; rescale if needed.', 'Uniform: constant density, not just bounded data.')),
    ('A statistic with a known sampling construction?', ('Chi-square', 'F'), 'x ≥ 0',
     ('Chi-square: sum of independent squared N(0,1)s.', 'F: ratio of independent chi-squares / their df.')),
]


def build_svg():
    validate_api()
    used = {n for _, names, _, _ in DISCRETE_ROUTES + CONTINUOUS_ROUTES for n in names}
    if used != {r[0] for r in CONTINUOUS + DISCRETE}:
        raise ValueError('Decision guide and Distribution Map inventories differ')
    parts = ['<svg xmlns="http://www.w3.org/2000/svg" width="1400" height="1660" viewBox="0 0 1400 1660" role="img" aria-labelledby="title desc">',
             '<title id="title">Choosing a distribution</title>',
             '<desc id="desc">Start with the variable and its support. Follow discrete or continuous questions to candidate families, then check assumptions and fit. All seventeen implemented pre-release API surfaces are included. This is not automatic model selection.</desc>',
             '<rect width="1400" height="1660" fill="white"/>',
             '<g font-family="DejaVu Sans, Arial, sans-serif">']

    def text(x, y, value, size=20, color='#2B3A4E', bold=False):
        parts.append(f'<text x="{x}" y="{y}" font-size="{size}" fill="{color}" font-weight="{"bold" if bold else "normal"}">{escape(value)}</text>')

    def box(x, y, w, h, color):
        parts.append(f'<rect x="{x}" y="{y}" width="{w}" height="{h}" rx="9" fill="{color}"/>')

    text(40, 42, 'VBA PROBABILITY DISTRIBUTIONS', 18, '#1F5AA6', True)
    text(40, 96, 'Choosing a distribution', 44, '#0B2A5B', True)
    text(40, 136, 'Start with the variable. Check the mechanism. Compare candidate models.', 25)
    box(40, 164, 1320, 102, '#EAF1FA')
    text(60, 198, '1  DEFINE WHAT X MEASURES', 25, '#0B2A5B', True)
    text(60, 230, 'Counts or measurements? What values are possible? Fixed trials, exposure or stopping rule?', 22)
    text(60, 255, 'These routes can overlap. A matching support narrows the options; it does not establish a good fit.', 19)
    for x, title, routes, color in ((40, '2A  DISCRETE · counts / integer outcomes', DISCRETE_ROUTES, '#096D72'),
                                    (714, '2B  CONTINUOUS · measurements / statistics', CONTINUOUS_ROUTES, '#1F5AA6')):
        box(x, 290, 646, 52, color)
        text(x + 16, 324, title, 23, 'white', True)
        for i, (question, names, support, assumptions) in enumerate(routes):
            y = 356 + i * 162
            box(x, y, 646, 150, '#F0F5FB')
            text(x + 16, y + 29, question, 21, '#0B2A5B', True)
            text(x + 16, y + 60, '→ ' + ' / '.join(names), 22, color, True)
            text(x + 16, y + 86, support, 19)
            for j, line in enumerate(assumptions):
                text(x + 16, y + 112 + 24 * j, line, 18)
    box(40, 1188, 1320, 165, '#0B2A5B')
    text(60, 1227, '3  CHECK THE CANDIDATES BEFORE USING THEM', 27, 'white', True)
    for y, line in zip((1262, 1295, 1328), (
        'Check support, zeros, rounding, truncation, censoring, dependence and stability over time.',
        'Estimate parameters; examine probability plots, tail fit and out-of-sample performance.',
        'Reject inadequate models. Similar means or a good histogram alone do not validate the choice.')):
        text(60, y, line, 22, 'white')
    box(40, 1376, 1320, 112, '#EDF7F6')
    text(60, 1411, '4  CHECK THE IMPLEMENTATION AND ITS LIMITS', 25, '#096D72', True)
    text(60, 1444, 'Confirm parameter conventions, supported operations, accuracy contracts and known limitations.', 22)
    text(60, 1473, 'Geometric / Negative Binomial count failures, not trials. Beta has no point mass at 0 or 1.', 21)
    text(40, 1527, 'CURRENT API / PRE-RELEASE · 17 surfaces, reconciled with the Distribution Map.', 22, '#0B2A5B', True)
    text(40, 1562, 'Planned noncentral distributions and dependence extensions are outside this guide.', 22)
    text(40, 1596, 'Examples: defects in n items → Binomial; waiting with constant hazard → Exponential (if assumptions hold).', 20)
    text(40, 1630, 'Companion: docs/visuals/choosing-distribution.md · No automatic selection or release certification.', 19)
    return '\n'.join(parts + ['</g>', '</svg>', ''])


def main():
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument('--check', action='store_true')
    parser.add_argument('--png', action='store_true')
    args = parser.parse_args()
    svg = build_svg()
    if args.check:
        if SVG.read_text() != svg:
            raise SystemExit('Decision guide SVG is stale')
    else:
        SVG.write_text(svg, encoding='utf-8')
    if args.png:
        import cairosvg
        data = cairosvg.svg2png(bytestring=svg.encode(), output_width=2800, output_height=3320)
        if args.check:
            if PNG.read_bytes() != data:
                raise SystemExit('Decision guide PNG is stale')
        else:
            PNG.write_bytes(data)
    print('PASS: decision guide currency and all 17 API surfaces')


if __name__ == '__main__':
    main()
