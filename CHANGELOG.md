# EPPIB model: v3.4 to v3.5 (theory-based forecasts)

`FINM3008_EPPIB_Model_v3.5.xlsx` is built from v3.4 by `tools/make_v35.sh`. The `Forecast` tab now
builds every asset-class forecast from a method in the literature, using sourced September 2026
inputs. Every forecast-dependent Solver set has then been re-optimised.

## Method by asset class (`Forecast` sections 5 to 7)

| Asset class | Method | Key sources |
|---|---|---|
| Australian, world and EM equities | Grinold-Kroner (dividend yield - net dilution + real earnings growth + inflation + repricing), averaged with a risk-premium build-up (expected cash + equity premium) | Grinold & Kroner (2002); CFA Institute, *Capital Market Expectations*; Dimson, Marsh & Staunton (2026); Bernstein & Arnott (2003); Campbell & Shiller (1998) |
| Foreign assets in AUD | Relative purchasing power parity (AUD drifts with the inflation gap) | Rogoff (1996) |
| Listed property | Grinold-Kroner on REIT yields; build-up using a beta estimated from the data | Nareit (2026) |
| Direct property | Cap-rate model: net rental yield + rent growth - change in yield | CFA Institute; Fox & Tulip (2014, RBA); Cotality (2026) |
| Cash | Expectations hypothesis (5-year yield less term premium), averaged with today's rate | Hambur & Finlay (2018, RBA) |
| Australian fixed income | Starting yield to maturity less credit losses | Leibowitz; Lozada (2015); Vanguard VAF fact sheet |
| World fixed income (hedged) | Covered interest parity: foreign yields + AUD/foreign cash-rate gap - basis | CFA Institute; central bank and market yields, Sep-2026 |
| Commodities | Collateral + long-run excess return + PPP drift | Erb & Harvey (2006); Levine et al. (2018) |
| Hedge funds | Cash + equity beta x premium, zero net alpha | Fung & Hsieh (2004); Dichev & Yu (2011) |
| Private equity | Cash + levered beta x premium + net premium | Harris, Jenkinson & Kaplan (2014); Korteweg (2019) |

* All views are compound returns. They are converted to arithmetic (+ σ²/2) before the
  Black-Litterman-style blend with the equilibrium prior, because the optimiser needs arithmetic
  means. Confidence still falls as the views disagree.
* Target tests now use compound returns: brief objective (a), "target reachable", the constraint-cost
  table, the return-advantage checks, the forecast stress test and the franking view.
* Section 6c estimates betas from the data. Private equity and hedge fund estimates are distorted by
  AUD-denominated indices (hedge fund R² is 0.00), so literature values are used and the data
  estimates are shown alongside.
* New sections: 8 (90% ranges), 9 (comparison with v3.4, 20-year history and published assumptions)
  and 10 (rationale text for the report). Section 4, the v3.4 method, is kept for comparison.

## Net forecasts (compound, p.a.)

| Asset class | v3.5 | v3.4 (old method) |
|---|---|---|
| Australian equities | 6.60% | 8.27%* |
| World equities | 6.92% | 8.22%* |
| Emerging markets | 6.52% | 7.88%* |
| Listed property | 6.51% | 7.32%* |
| Commodities | 3.97% | 5.03%* |
| Direct property | 3.97% | 5.44%* |
| Hedge funds | 3.47% | 3.32%* |
| Private equity | 4.34% | 5.87%* |
| Australian fixed income | 4.93% | 4.79%* |
| World fixed income (hedged) | 5.25% | 4.77%* |
| Cash | 4.77% | 4.60%* |

\* v3.4 mixed compound and arithmetic inputs, so its figures are not strictly comparable.

## Re-optimisation

The Python replicas in `tools/optim.py` reproduce the v3.4 Solver weights to within 0.01% for
every set. They were then re-run on the new forecasts: Robust, Sample and Resampled MVO, minimum
variance, the three sensitivity sets, the five constraint-cost sets and both 10-point frontiers.
Unchanged because they do not use the forecasts: Naive MVO (historical means), risk parity and the
2006-16 out-of-sample fits.

**New recommendation:** Australian equities 26.8%, world equities 27.6%, emerging markets 5.5%,
listed property 8.9%, Australian fixed income 5.0% (floor), world fixed income 15.2%, cash 2.0%.
Stage 2 weights are unchanged.

* Compound expected return 6.64% against 6.61% for the current fund; volatility 7.90% against
  8.38%; parametric TE 1.10%; bootstrap TE 1.39%. All brief constraints are met, and (a) is
  maximum feasible.
* Main change: about 8 points move from Australian into hedged world bonds. With the AUD cash rate
  above foreign rates, hedging adds carry.
* REVIEW items to disclose: interim TE during the transition; the robust vs resampled gap (5.9%,
  Australian fixed income at its floor); the return advantage after transition costs (-0.002%);
  the return advantage after franking (-0.11%). Equity including listed property now passes
  (68.8%).

## Inputs to verify before submission

`Forecast` section 5 flags these as "verify":
* the long-run forward P/E averages for ASX 200 (14.8x), MSCI World (15.5x) and EM (12.0x)
* the EM dividend yield (2.0%)
* the 2026 fact-sheet figures for VAF and VIF

Check them against J.P. Morgan's *Guide to the Markets* and the MSCI and Vanguard fact sheets, then
re-run `tools/make_v35.sh`, or Solver.

---

# EPPIB model: v3.3 to v3.4

`FINM3008_EPPIB_Model_v3.4.xlsx` is built from `source/FINM3008_EPPIB_Model_v3.3.xlsx` by
`tools/make.sh`. Every addition uses the existing boxed layout and palette. Nothing existing was
moved: new sections sit below or beside current content, so chart positions and cross-sheet links
are unchanged. The 19 original charts are carried over byte for byte. The optimisation, risk,
bootstrap, stress and Solver outputs are unchanged (every one of their cells was compared with v3.3).

## Corrections

| Where | Problem | Fix |
|---|---|---|
| `Quarterly returns` Z:AD (feeds `Track record`, `Validation` back-test, Figures 3, 4, 10) | Historical portfolio returns were gross of fees, although the data notice says fees must be deducted | Each portfolio's own fee is deducted quarterly |
| `Risk` P59:P62, `Summary` K41 | `SUMPRODUCT(MAX(ABS(...)))` returned `#VALUE!` in Excel | Explicit `MAX(ABS(..),..)` over the 11 weights |
| `Validation` E61 | Same formula pattern silently showed 2.46%; the true largest gap between current and recommended is 13.0% (Australian equities) | As above |
| `Oil` S column | FRED 3-month rate is 0 for Nov-1969 (missing value), creating a fake -5.6pp / +5.65pp move | Zero values interpolated from neighbouring months; the 60-month rate beta moves from 2.63 to 2.61 |
| `Summary` H7 | Objective (b) was labelled "risk no higher than current fund" | Relabelled "Minimise risk: volatility vs current fund" |

## New: `Performance` tab (criterion 1)

Monthly rebalancing, as the brief specifies, net of fees, Mar-2006 to Mar-2026.

1. Settings
2. Performance summary for 5 and 20 years: return (net and gross), volatility, Sharpe ratio, tracking
   error, information ratio, beta, maximum drawdown, worst 12 months, up and down capture, hit rate, target
3. Non-overlapping 5-year blocks plus rolling 60-month statistics ("is performance improving?")
4. Attribution of active return by asset class (allocation effect plus fee effect); it reconciles
   exactly to the arithmetic active return (checked in `Checks`)
5. Method notes

## Oil tab (criterion 3)

* 7: post-1986 subsample (after WTI deregulation) and AU 5-year yield regressions at 1, 12 and 60 months
* 8: oil betas for all 11 asset classes at 1, 12 and 60 months
* 9: US$150 portfolio impact by horizon, regression approach against the episode approach
* 10: 2008 oil spike as a third episode. The switch defaults to 0, so the central case is unchanged
  and 2008 serves as the severe case.
* 11: mitigation options compared on return, volatility, TE, oil loss and brief limits
* 12: other mitigation tools (qualitative)
* Panel extension (from column AG): 5-year yield and asset-class log changes

## Implementation, tax and integrity

* `Inputs` 4 and 5: transaction costs, transition ramp, franking parameters (blue = judgement inputs)
* `Portfolio` 5: transition costs; 6: transition plan by phase with constraint tests; 7: franking
  credit sensitivity; 8: equity band if listed property counts as equity
* `Validation` 7 (the empty section 7 slot): every pasted Solver weight set re-tested against current
  inputs (sum, no shorts, cash floor, equity band, TE limit)
* `Checks`: 14 new tests. A new REVIEW status marks items to disclose in the report; the master flag
  still counts FAIL only.
* `Summary` 9 to 11; `Report figures` 14 (rolling 5-year active return) and 15 (attribution)

## Findings worth using in the report

* Net of fees with monthly rebalancing, the fund returned 6.92% a year over Mar-21 to Mar-26, against
  7.35% for the benchmark (-0.43%). Over 20 years the figures are 6.49% and 6.52% (-0.03%).
* Active return by 5-year block: +0.12%, +0.06%, +0.14%, then -0.43%. The fund has not improved; the
  latest block is the weakest.
* Over 20 years: beta 1.20, down capture 1.25, maximum drawdown -28.9% against -22.8%. The extra
  equity risk was not paid for.
* Largest 5-year drags: no commodities (-0.29%), world equity underweight (-0.24%), cash overweight
  (-0.19%), no private equity (-0.16%). Offsets: lower fees (+0.23%), world fixed income underweight
  (+0.26%), Australian equity overweight (+0.18%).
* At US$150 oil, the regression approach gives about +1.0% over 12 months but -5.0% over 60 months;
  the episode central case is -7.8%; the 2008 severe case is -13.1%. Oil raises the 5-year yield by
  about 0.63pp, which hurts bonds.
* Every oil hedge tested breaches the 1.1% calibrated TE limit. Commodities +2% funded from
  Australian equities cuts the oil loss by 0.75pp, at a TE of 1.27% and a 0.06% return cost.
* Transition costs are about A$4.1m (0.14% of the fund). Amortised over 5 years, they leave only
  +0.01% a year of the expected-return advantage.
* Items flagged REVIEW (disclose these):
  * Franking: allowing for franking credits turns the advantage negative (-0.10% a year).
  * Equity definition: counting listed property as equity gives 69.5% against a 69% ceiling.
  * Transition: interim TE is 1.45% in phase 1 and 1.22% at the end of year 1.
  * Methods: the robust and resampled weights differ by more than the 3% stability limit.

## Not changed (needs Solver or your judgement)

* Forecast blend weights (for example the 20-year history weight on bonds and cash). Changing them
  means re-running Solver for every pasted weight set; `Validation` 7 will show REVIEW until you do.
* Currency hedging is discussed qualitatively (`Oil` 12) but not modelled.

## Rebuild

```
tools/make.sh source/FINM3008_EPPIB_Model_v3.3.xlsx FINM3008_EPPIB_Model_v3.4.xlsx
```

Recalculation uses LibreOffice. It reproduces 39,541 of the 39,542 numeric cells that Excel cached
in v3.3; the exception is the `Validation` E61 bug fixed above. Excel recalculates everything on
open.
