#!/usr/bin/env python3
"""Render a calculated Normal PDF/CDF/survival/quantile guide (#245)."""
import argparse
from html import escape
import math
from pathlib import Path
import re
from statistics import NormalDist

ROOT = Path(__file__).resolve().parents[2]
SVG = ROOT / 'assets/probability-functions.svg'
PNG = ROOT / 'assets/probability-functions.png'
NAVY, BLUE, TEAL, INK = '#0B2A5B', '#1F5AA6', '#096D72', '#2B3A4E'
THRESHOLD = 1.0
NORMAL = NormalDist()
OPERATIONS = ('Density', 'Cumulative', 'Survival', 'InverseCumulative')


def density(x):
    return math.exp(-x*x/2) / math.sqrt(2*math.pi)


def cdf(x):
    return math.erfc(-x/math.sqrt(2))/2


def survival(x):
    return math.erfc(x/math.sqrt(2))/2


def samples():
    return [-4 + i/100 for i in range(801)]


def validate_api():
    names = set(re.findall(r'\t(K_STATS_\w+)\t',
                          (ROOT / 'docs/PUBLIC_API.txt').read_text(encoding='utf-8')))
    for op in OPERATIONS:
        if 'K_STATS_NormalStandard_' + op not in names:
            raise ValueError(f'Missing NormalStandard operation: {op}')


def verify_reference():
    """Independent high-precision integration/inversion, not a VBA test."""
    import mpmath as mp
    with mp.workdps(80):
        phi = lambda x: mp.exp(-x*x/2)/mp.sqrt(2*mp.pi)
        p = mp.quad(phi, [-mp.inf, 0, THRESHOLD])
        q = mp.quad(phi, [THRESHOLD, mp.inf])
        if abs(mp.mpf(cdf(THRESHOLD))-p) > mp.mpf('2e-16'):
            raise ValueError('CDF annotation failed quadrature reference')
        if abs(mp.mpf(survival(THRESHOLD))-q) > mp.mpf('2e-16'):
            raise ValueError('Survival annotation failed quadrature reference')
        # mpmath receives the exact plotted binary64 arguments.
        for x in samples():
            xx = mp.mpf(x)
            expected = (phi(xx), mp.erfc(-xx/mp.sqrt(2))/2, mp.erfc(xx/mp.sqrt(2))/2)
            actual = (density(x), cdf(x), survival(x))
            if any(abs(mp.mpf(a)-b) > mp.mpf('3e-16') for a, b in zip(actual, expected)):
                raise ValueError(f'Curve reference mismatch at {x}')
        for x in [-3+i/100 for i in range(601)] + [THRESHOLD]:
            probability = cdf(x)
            ref = mp.sqrt(2)*mp.erfinv(2*mp.mpf(probability)-1)
            if abs(mp.mpf(NORMAL.inv_cdf(probability))-ref) > mp.mpf('3e-15'):
                raise ValueError(f'Quantile reference mismatch at {probability}')
    print('PASS: 801 PDF/CDF/SF samples, 602 quantiles and area annotations; mpmath 80 digits')


def build_svg():
    p, q = cdf(THRESHOLD), survival(THRESHOLD)
    out = ['<svg xmlns="http://www.w3.org/2000/svg" width="1200" height="1240" viewBox="0 0 1200 1240" role="img" aria-labelledby="title desc">',
           '<title id="title">Density, cumulative, survival and quantile — one Normal example</title>',
           '<desc id="desc">For Z standard Normal at z=1: density 0.241971, cumulative area 0.841345, survival area 0.158655, and inverse cumulative at the full precision cumulative probability returns approximately 1. Density is height, probability is area. Survival means P(Z greater than z).</desc>',
           '<rect width="1200" height="1240" fill="white"/>',
           '<g font-family="DejaVu Sans, Arial, sans-serif">']

    def text(x, y, s, size=18, color=INK, bold=False, anchor='start'):
        out.append(f'<text x="{x}" y="{y}" font-size="{size}" fill="{color}" font-weight="{"bold" if bold else "normal"}" text-anchor="{anchor}">{escape(s)}</text>')

    def rect(x, y, w, h, fill):
        out.append(f'<rect x="{x}" y="{y}" width="{w}" height="{h}" rx="9" fill="{fill}"/>')

    def line(x1, y1, x2, y2, color=INK, dash=False):
        out.append(f'<path d="M{x1:.3f},{y1:.3f} L{x2:.3f},{y2:.3f}" fill="none" stroke="{color}" stroke-width="1.5"'+(' stroke-dasharray="5 4"' if dash else '')+'/>')

    def path(points, color, fill='none', close=False):
        d = 'M' + ' L'.join(f'{x:.3f},{y:.3f}' for x, y in points) + (' Z' if close else '')
        out.append(f'<path d="{d}" fill="{fill}" stroke="{color}" stroke-width="2.5"/>')

    text(36, 40, 'VBA PROBABILITY DISTRIBUTIONS', 17, BLUE, True)
    text(36, 91, 'One distribution. Four ways to read it.', 38, NAVY, True)
    text(36, 131, 'Z ~ Normal(0, 1)    •    threshold z = 1    •    p = F(1), q = S(1)', 23)
    rect(36, 154, 1128, 44, '#EAF1FA')
    text(54, 183, 'Same model and threshold throughout. Probabilities shown to 6 decimals; calculations use full precision.', 18, NAVY)
    for index, (title, op) in enumerate(zip(('1  Density · height is not probability',
                                           '2  Cumulative · area to the left',
                                           '3  Survival · area to the right',
                                           '4  Quantile · probability to threshold'), OPERATIONS)):
        left, top = 36+(index % 2)*576, 220+(index//2)*410
        rect(left, top, 552, 390, '#F0F5FB')
        text(left+18, top+32, title, 22, NAVY, True)
        x0, y0, w, h = left+72, top+100, 444, 194
        inverse = index == 3
        xmin, xmax, ymin, ymax = (0, 1, -3, 3) if inverse else (-4, 4, 0, .45 if index == 0 else 1)
        def point(x, y):
            return x0+(x-xmin)/(xmax-xmin)*w, y0+h-(y-ymin)/(ymax-ymin)*h
        def graph_line(a, b, color=INK, dash=False):
            line(*point(*a), *point(*b), color, dash)
        xticks = (0, .5, 1) if inverse else (-4, -2, 0, 1, 2, 4)
        yticks = (-3, 0, 3) if inverse else (0, .2, .4) if index == 0 else (0, .5, 1)
        for y in yticks:
            graph_line((xmin,y), (xmax,y), '#D5DFED')
            text(x0-12, point(xmin,y)[1]+5, f'{y:g}', 15, anchor='end')
        for x in xticks:
            px, py = point(x,ymin)
            line(px,py,px,py+5)
            text(px,py+23,f'{x:g}',15,anchor='middle')
        graph_line((xmin,ymin),(xmax,ymin))
        graph_line((xmin,ymin),(xmin,ymax))
        text(x0, y0-17, 'quantile z' if inverse else ('density f(z)' if index == 0 else 'probability'), 17)
        text(x0+w, y0+h+43, 'probability p' if inverse else 'threshold z', 16, anchor='end')
        if index == 0:
            xs = samples()
            for subset, fill in (([x for x in xs if x <= THRESHOLD], '#B8CEE9'),
                                 ([x for x in xs if x >= THRESHOLD], '#B5DCD6')):
                area = [(subset[0],0)]+[(x,density(x)) for x in subset]+[(subset[-1],0)]
                path([point(*v) for v in area], 'none', fill, True)
            path([point(x,density(x)) for x in xs],BLUE)
            graph_line((1,0),(1,density(1)),NAVY,True)
            text(left+18,top+65,f'f(1) = {density(1):.6f}   •   P(Z = 1) = 0',19)
            text(left+18,top+347,f'Areas: left p = {p:.6f}   •   right q = {q:.6f}',17)
        elif inverse:
            xs = [-3+i/100 for i in range(601)]
            path([point(cdf(x),NORMAL.inv_cdf(cdf(x))) for x in xs], BLUE)
            graph_line((p,-3),(p,1),TEAL,True)
            graph_line((0,1),(p,1),TEAL,True)
            px, py = point(p,1)
            out.append(f'<circle cx="{px:.3f}" cy="{py:.3f}" r="5" fill="{TEAL}"/>')
            text(left+18,top+65,f'F⁻¹(p) ≈ 1   when p = F(1) ≈ {p:.6f}',18)
            text(px+8,py-12,'z = 1',17,TEAL,True)
        else:
            fn, color, level = (cdf,BLUE,p) if index == 1 else (survival,TEAL,q)
            path([point(x,fn(x)) for x in samples()],color)
            graph_line((1,0),(1,level),color,True)
            graph_line((-4,level),(1,level),color,True)
            px, py = point(1,level)
            out.append(f'<circle cx="{px:.3f}" cy="{py:.3f}" r="5" fill="{color}"/>')
            text(left+18,top+65,('F(1) = P(Z ≤ 1) = ' if index == 1 else 'S(1) = P(Z > 1) = ')+f'{level:.6f}',19,color)
        text(left+18,top+372,'K_STATS_NormalStandard_'+op,16,NAVY)
    rect(36, 1044, 1128, 82, NAVY)
    text(54,1076,'DENSITY ≠ PROBABILITY',22,'white',True)
    text(54,1105,'For continuous Z: P(Z = 1) = 0; probability is an area. For discrete X: PMF(k) = P(X = k).',20,'white')
    text(36,1160,'Tail convention: S(x) = P(X > x) = 1 − F(x). Use direct Survival for small upper tails.',20,NAVY)
    text(36,1191,'Plots show finite windows; area labels include the full infinite tails. Quantile endpoints are unbounded.',17)
    text(36,1221,'Illustration of implemented APIs, not Excel certification. Availability and valid domains vary by operation.',17)
    return '\n'.join(out+['</g>','</svg>',''])


def main():
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument('--check', action='store_true')
    parser.add_argument('--png', action='store_true')
    parser.add_argument('--verify-reference', action='store_true', help='Verify using mpmath 1.3.0 at 80 digits')
    args = parser.parse_args()
    validate_api()
    if args.verify_reference:
        verify_reference()
    svg = build_svg()
    if args.check:
        if SVG.read_text(encoding='utf-8') != svg:
            raise SystemExit('Probability-functions SVG is stale')
    else:
        SVG.write_text(svg,encoding='utf-8')
    if args.png:
        import cairosvg
        rendered = cairosvg.svg2png(bytestring=svg.encode(), output_width=2400, output_height=2480)
        if args.check:
            if PNG.read_bytes() != rendered:
                raise SystemExit('Probability-functions PNG is stale')
        else:
            PNG.write_bytes(rendered)
    print('PASS: NormalStandard API mapping and probability-functions visual')


if __name__ == '__main__':
    main()
