# Gamma continued-fraction caller coverage

Source inventory at B6 integration 5618526. Source search covers every executable
call to `PROB_TryGammaContinuedFractionQ`, `PROB_TryGammaRegularizedP`,
`PROB_TryGammaRegularizedQ` and `PROB_TryGammaInvP` in `src/*.bas`.

| Public surface | Route to upper CF | Permanent probe coverage |
| --- | --- | --- |
| Gamma Cumulative | RegularizedP; 1-Q above X=A+1 | 89 pairs, scales .5/1/2, P+Q, adjacent-X seam |
| Gamma Survival | RegularizedQ | Same pairs/scales, shape seam, underflow |
| Gamma InverseCumulative | GammaInvP; P or Q objective according to probability | A=.01/.5/1/2; p=.1/.5/.9 |
| ChiSquare Cumulative | RegularizedP at A=df/2, X=input/2 | Equivalent calls for all 89 pairs |
| ChiSquare Survival | RegularizedQ at A=df/2, X=input/2 | Equivalent calls, underflow |
| ChiSquare InverseCumulative | GammaInvP plus guarded rescale | Same twelve inverse pairs |
| Poisson Cumulative | PoissonCDF -> RegularizedQ(K+1, mean) | K=0/1/10; mean below/at/above K+2 |
| Poisson Survival | PoissonSF -> RegularizedP(K+1, mean) | Same nine points |
| Poisson InverseCumulative | Integer quantile search -> PoissonCDF/SF | mean=.5/2/12; p=.1/.5/.9; both defining inequalities |

The incomplete-gamma inverse itself evaluates RegularizedP for the lower
probability objective and RegularizedQ for the upper objective. Thus forward
and inverse callers both require testing; direct distribution density routines
do not call this continued fraction. Student-t, F and beta inverses use the
incomplete-beta family, not this Gamma CF.

Poisson reaches the changed A<=1 region only at K=0 (A=1); K>0 checks retain
coverage of the unchanged A>1 route. Chi-square's public df cap prevents a
natural 1E14-shape cap probe, so the cap fixture exercises the direct kernels
and the Gamma public CDF/SF. It does not claim runtime cap coverage for every
public distribution wrapper.

## Coverage limits

The inventory covers the callers, not every possible argument combination.
The 25 residual accuracy failures in RESULT.md stay blocking. The 64 additional
cap/recovery and adjacent-X assertions at 15bb7ce have a separate Excel result;
they are not included in the 1881-assertion paired result. The second paired run
now confirms all 12 cap/recovery assertions pass and 15 X-seam assertions fail,
including two new B6 regressions at A=.5, X=1.5; see CAP_SEAM_RESULT.md. No main-grid promotion
or fresh export is implied by this caller map.
