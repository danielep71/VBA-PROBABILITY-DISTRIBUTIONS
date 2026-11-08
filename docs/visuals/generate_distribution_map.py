#!/usr/bin/env python3
"""Generate the editable distribution overview; optionally render its PNG."""
import argparse
from html import escape
from pathlib import Path
import re

ROOT = Path(__file__).resolve().parents[2]
SVG = ROOT / 'assets/distribution-map.svg'
PNG = ROOT / 'assets/distribution-map.png'
# Name, API stem, mathematical support (nondegenerate parameters), example.
CONTINUOUS = [
    ('Standard Normal', 'NormalStandard', 'x ∈ ℝ  ·  mean 0, variance 1', 'Standardized scores'),
    ('Normal', 'Normal', 'x ∈ ℝ', 'Symmetric measurement variation'),
    ('Lognormal', 'Lognormal', 'x > 0', 'Positive multiplicative quantities'),
    ('Student t', 'StudentT', 'x ∈ ℝ', 'Studentized statistics; heavy tails'),
    ('Chi-square', 'ChiSquare', 'x ≥ 0', 'Sums of squared standard normals'),
    ('F', 'F', 'x ≥ 0', 'Ratios of independent scaled chi-squares'),
    ('Gamma', 'Gamma', 'x ≥ 0  ·  shape and scale', 'Positive waiting times'),
    ('Beta', 'Beta', '0 ≤ x ≤ 1', 'Bounded proportions'),
    ('Exponential', 'Exponential', 'x ≥ 0  ·  rate λ', 'Memoryless waiting times'),
    ('Weibull', 'Weibull', 'x ≥ 0  ·  shape and scale', 'Lifetimes with varying hazard'),
    ('Uniform', 'Uniform', 'a ≤ x ≤ b', 'Equal density on a finite interval'),
]
DISCRETE = [
    ('Binomial', 'Binomial', 'k = 0, …, n', 'Successes in n independent trials'),
    ('Poisson', 'Poisson', 'k = 0, 1, 2, …', 'Event counts over a fixed exposure'),
    ('Geometric', 'Geometric', 'k = 0, 1, 2, …', 'Failures before the first success'),
    ('Negative Binomial', 'NegativeBinomial', 'k = 0, 1, 2, …', 'Failures before the r-th success'),
    ('Hypergeometric', 'Hypergeometric', 'max(0, n − N + K) ≤ k ≤ min(n, K)', 'Successes when sampling without replacement'),
    ('Discrete Uniform', 'DiscreteUniform', 'k = a, a + 1, …, b', 'Equal mass on consecutive integers'),
]
NAVY, BLUE, INK, TEAL = '#0B2A5B', '#1F5AA6', '#2B3A4E', '#096D72'


def validate_api():
    api = (ROOT / 'docs/PUBLIC_API.txt').read_text()
    names = set(re.findall(r'\t(K_STATS_\w+)\t', api))
    stems = {name.split('_')[2] for name in names}
    expected = {row[1] for row in CONTINUOUS + DISCRETE}
    if stems != expected:
        raise ValueError(f'API families changed: {stems ^ expected}')
    for rows, measure in ((CONTINUOUS, 'Density'), (DISCRETE, 'PMF')):
        for _, stem, _, _ in rows:
            for operation in (measure, 'Cumulative', 'Survival', 'InverseCumulative'):
                if f'K_STATS_{stem}_{operation}' not in names:
                    raise ValueError(f'Missing operation: {stem}_{operation}')
    inverse_survival = {n.split('_')[2] for n in names if n.endswith('_InverseSurvival')}
    if inverse_survival != {'NormalStandard', 'Normal', 'Lognormal'}:
        raise ValueError('Inverse-survival capability changed')


def build_svg():
    parts = ['<svg xmlns="http://www.w3.org/2000/svg" width="1200" height="1530" viewBox="0 0 1200 1530" role="img" aria-labelledby="title desc">',
             '<title id="title">Distribution map — VBA Probability Distributions</title>',
             '<desc id="desc">Seventeen current API surfaces: eleven continuous and six discrete. Each row gives mathematical support and an illustrative use case. Implemented pre-release capabilities are not release certification; numerical domains are narrower than mathematical support.</desc>',
             '<rect width="1200" height="1530" fill="white"/>',
             '<g font-family="DejaVu Sans, Arial, sans-serif">']

    def text(x, y, label, size=20, color=INK, bold=False):
        parts.append(f'<text x="{x}" y="{y}" font-size="{size}" fill="{color}" font-weight="{"bold" if bold else "normal"}">{escape(label)}</text>')

    def rect(x, y, w, h, color, radius=8):
        parts.append(f'<rect x="{x}" y="{y}" width="{w}" height="{h}" rx="{radius}" fill="{color}"/>')

    text(36, 43, 'VBA PROBABILITY DISTRIBUTIONS', 17, BLUE, True)
    text(36, 94, 'A map of the distributions', 40, NAVY, True)
    text(36, 132, '17 API surfaces · mathematical support · illustrative applications', 23)
    rect(36, 156, 1128, 40, '#EAF1FA')
    text(52, 183, 'IMPLEMENTED / PRE-RELEASE    •    Standard Normal is a dedicated API surface.', 18, NAVY)

    for x, label, rows, color in ((36, 'CONTINUOUS  /  11 surfaces', CONTINUOUS, BLUE),
                                   (612, 'DISCRETE  /  6 surfaces', DISCRETE, TEAL)):
        rect(x, 220, 552, 51, color)
        text(x + 18, 254, label, 23, 'white', True)
        for j, (name, _, support, use) in enumerate(rows):
            y = 285 + j * 88
            rect(x, y, 552, 80, '#F0F5FB' if j % 2 == 0 else '#F8FAFD')
            text(x + 16, y + 25, name, 23, NAVY, True)
            text(x + 16, y + 49, support, 19)
            text(x + 16, y + 70, use, 17, '#465970')

    rect(612, 840, 552, 176, '#EAF1FA')
    text(630, 871, 'COMMON OPERATIONS', 22, NAVY, True)
    for y, line in zip((905, 935, 965, 995), (
        'Continuous: density · CDF · survival · inverse CDF',
        'Discrete: PMF · CDF · survival · inverse CDF',
        'CDF = P(X ≤ x)     Survival = P(X > x)',
        'Inverse survival: all three Normal-family surfaces')):
        text(630, y, line, 17)
    rect(612, 1034, 552, 209, '#EDF7F6')
    text(630, 1066, 'READ THE SUPPORT CORRECTLY', 22, TEAL, True)
    for y, line in zip((1100, 1128, 1156, 1184, 1212), (
        'Discrete k is integer; n is the trial or sample size.',
        'Hypergeometric: N population size, K successes.',
        'Uniform: a and b are the lower and upper bounds.',
        'Zero may be a boundary where a density is singular.',
        'Degenerate parameter cases can collapse the support.')):
        text(630, y, line, 17)

    rect(36, 1270, 1128, 114, NAVY)
    text(56, 1304, 'CHOOSE A MODEL, THEN CHECK ITS ASSUMPTIONS', 24, 'white', True)
    text(56, 1336, 'Use cases illustrate possible applications. They do not establish fit, independence or calibration.', 20, 'white')
    text(56, 1365, 'Mathematical support is not a promise of numerical accuracy across the whole domain.', 20, 'white')
    text(36, 1421, 'Before use: check parameters, validated domains, error behavior and known limitations.', 22, NAVY, True)
    text(36, 1456, 'Current API only. Planned noncentral distributions and dependence extensions are not shown.', 20)
    text(36, 1495, 'Companion: docs/visuals/distribution-map.md  •  No release certification is implied.', 17, '#465970')
    return '\n'.join(parts + ['</g>', '</svg>', ''])


def main():
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument('--check', action='store_true', help='Check SVG currency and API capabilities')
    parser.add_argument('--png', action='store_true', help='Render matching PNG (CairoSVG 2.8.2)')
    args = parser.parse_args()
    validate_api()
    svg = build_svg()
    if args.check:
        if SVG.read_text() != svg:
            raise SystemExit('Distribution map SVG is stale')
    else:
        SVG.write_text(svg, encoding='utf-8')
    if args.png:
        import cairosvg
        rendered = cairosvg.svg2png(bytestring=svg.encode(), output_width=2400, output_height=3060)
        if args.check:
            if PNG.read_bytes() != rendered:
                raise SystemExit('Distribution map PNG differs from the current SVG rendering')
        else:
            PNG.write_bytes(rendered)
    print('PASS: 17 API surfaces and core operations reconciled; map is current')


if __name__ == '__main__':
    main()
