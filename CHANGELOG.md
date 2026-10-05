# EPPIB model: v3.7 (fixes from the v3.6 review)

`FINM3008_EPPIB_Model_v3.7.xlsx` is rebuilt end to end from v3.3 by `tools/make_v37.sh`.

## 1. Oil mitigation built into the recommendation (essential)
* New optimiser constraint: the US$150 central-case loss must be no worse than the current fund's
  (`Inputs` B50 = `Oil` F43). It applies to every forecast-dependent optimised set: robust, sample,
  resampled (each draw), minimum variance, the sensitivity sets, the constraint-cost cases, both
  frontiers and the scope portfolios. It does not apply to risk parity, naive MVO or the out-of-sample
  fits, which do not use the forecasts.
* Shown in `Portfolio` section 3 (row 39), `Summary` compliance row (f), `Checks` row 50 and a new
  oil column in the `Validation` integrity table.
* Cost: under 0.1bp of compound return. Commodities rise from 2.4% to 2.7%, funded mainly from
  emerging markets.
* A stricter floor at the benchmark's loss is priced in the constraint-cost table ("Oil floor at
  benchmark loss"). It brings the oil loss to -7.4% but costs about 4bp, moves equity down to 58.6%
  and adds 1.4% direct property. It also pushes the robust vs resampled gap to 3.4 points, above the
  3-point stability rule, so it is shown as an alternative and is not the base case.
* The Stage 2 overweight cap (`Inputs` B47, 2%) now caps the alternatives' upper bounds in the
  optimiser (`Portfolio` G10:G13). It does not bind in the base case.

## 2. TE setting relabelled
`Inputs` C45: "Judgement: buffer below the 1.5% brief limit". The value stays at 1.10%; bootstrap TE
is 1.17%.

## 3. Scorecard return scores follow the forecasts
`Portfolio` B23:B26: the net arithmetic forecast scaled from 1 (lowest of the 11 classes) to 5
(highest). Commodities 4, direct property 2, hedge funds 1, private equity 2. The averages (4.2, 2.2,
2.0, 2.0) now agree with the optimiser, which holds commodities only. The fixed comparison case
(commodities 3%, direct property 5%, private equity 1%) is renamed "Alternatives at fixed weights"
because the scorecard no longer implies those weights.

## 4. Disclosures (report, not model)
Two REVIEW items remain, and both are disclosures: franking (-4bp advantage after franking) and
equity including listed property (70.4%).

## 5. Text trimmed
`Oil` section 12 (other mitigation tools) moved to `docs/oil_mitigation_notes.md`. Notes in
Performance, Portfolio and Validation are shortened.

## Result
| | Current fund | Benchmark | Recommended |
|---|---|---|---|
| Expected return (compound) | 6.61% | 6.38% | 6.70% |
| Volatility | 8.38% | 7.01% | 7.84% |
| Tracking error (parametric / bootstrap) | 2.16% / 1.56% | 0 | 1.10% / 1.17% |
| 5-year loss probability (bootstrap) | 7.7% | 6.2% | 7.2% |
| GFC / US$150 oil (central) | -24.2% / -8.3% | -19.3% / -7.4% | -23.5% / -8.3% |

Weights: Australian equities 27.5%, world equities 27.9%, emerging markets 6.5%, listed property
8.5%, commodities 2.7%, Australian fixed income 12.0%, world fixed income 12.9%, cash 2.0%.
Checks: ALL CHECKS PASS, 0 error cells, 21 charts. Robust vs resampled gap 2.8 points. All four
transition phases MET.

# EPPIB model: v3.6 (decisions from the v3.5 review)

`FINM3008_EPPIB_Model_v3.6.xlsx` is rebuilt end to end from v3.3 by `tools/make_v36.sh`.

## Decisions
1. **Alternatives are optimised, not fixed.** The optimiser holds no direct property, hedge funds or
   private equity, and 2.4% commodities. The Stage 2 scorecard stays visible as a cross-check
   (`Portfolio` section 2). The scorecard-fixed case is now a comparison portfolio in section 9 and
   the constraint-cost table.
2. **Australian fixed income floor of 12%** (`Inputs` E13). A 10% floor left the robust vs
   resampled gap at 4.1 points; at 12% it is 2.7 points, inside the 3-point rule. The return cost is
   under 1bp. Risk parity and Naive MVO are unchanged (they do not use the forecasts).
3. **Trimmed.**
   * Oil asset-class panel cut to the 12-month horizon; section 2 is untouched. Oil formulas fell
     from 44,700 to 22,300, and the workbook total from 73,900 to 51,600.
   * Forecast rationale text removed and saved to `docs/forecast_rationale_notes.md`, to rewrite in
     your own words.

## Other fixes from the review
* The transition plan and the oil mitigation options are now tested against the 1.5% brief TE limit.
  All four phases and all options pass.
* "Verify" source labels are relabelled as judgement inputs. They still need checking (see below).
* New `Forecast` section 9: a sensitivity table for the P/E reversion share. Its 50% row reproduces
  section 1.
* Arithmetic and compound returns are labelled in Methods, Portfolio, Risk, Validation and Report
  figures.

## Result
| | Current fund | Recommended |
|---|---|---|
| Expected return (compound) | 6.61% | 6.70% |
| Volatility | 8.38% | 7.84% |
| Tracking error (parametric / bootstrap) | 2.16% / 1.56% | 1.10% / 1.18% |
| 5-year loss probability (bootstrap) | 7.7% | 7.3% |
| Fee | 0.11% | 0.13% |
| GFC / US$150 oil (central) | -24.2% / -8.3% | -23.5% / -8.4% |

Weights: Australian equities 27.4%, world equities 27.9%, emerging markets 6.7%, listed property
8.7%, commodities 2.4%, Australian fixed income 12.0%, world fixed income 12.9%, cash 2.0%.

* Turnover is 22%. Transition cost is A$1.6m, and the advantage after costs is +0.07% a year.
* REVIEW flags left to disclose:
  * franking: the advantage is -0.04% after franking credits
  * equity including listed property: 70.7% against a 69% ceiling under the wider definition
* Oil: the recommended portfolio loses slightly more than the current fund (-8.4% vs -8.3%). A 2%
  commodity tilt from world equities cuts the loss to -7.5% for 1bp of return (`Oil` section 11).

## Still to do before submission
* Replace the long-run P/E sources:
  * ASX 200 (14.8x) currently cites an IG article. Use J.P. Morgan *Guide to the Markets Australia*.
  * MSCI World (15.5x) and EM (12.0x) are judgement.
* Check the EM dividend yield (2.0%) against the MSCI factsheet.
* Re-run Solver in Excel for the main portfolios to confirm the pasted weights.
* Open the file in Excel at least once.

---

# EPPIB model: v3.4 to v3.5 (theory-based forecasts)

`FINM3008_EPPIB_Model_v3.5.xlsx` is built from v3.4 by `tools/make_v35.sh`. The `Forecast` tab now
builds every asset-class forecast from a method in the literature, using sourced September 2026
inputs. The earlier building-block section, its parameter box and the weighted-excess column have
been removed, so the tab presents a single forecasting framework with no reference to earlier versions. Every forecast-dependent Solver set has then been re-optimised.

## Method by asset class (`Forecast` sections 4 to 6)

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
* Section 5c estimates betas from the data. Private equity and hedge fund estimates are distorted by
  AUD-denominated indices (hedge fund R² is 0.00), so literature values are used and the data
  estimates are shown alongside.
* Further sections: 7 (90% ranges), 8 (comparison with the 20-year history and published
  assumptions) and 9 (rationale text for the report). Section 3 (implied returns) is kept as a
  cross-check and uses the section 4 inflation and growth assumptions.

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

## Asset-class scope (`Portfolio` section 9, `Summary` box 12)

Five portfolios, each optimised under the same forecasts, risk model and brief constraints and scored
on the same measures, including the bootstrap (new columns AC:AF and L:S in `Bootstrap`). The
recommendation is unchanged.

| | Recommended | Alternatives optimised | No alternatives | Alternatives at benchmark | Core five only |
|---|---|---|---|---|---|
| Expected return (compound) | 6.64% | 6.70% | 6.66% | 6.57% | 6.50% |
| Volatility | 7.90% | 7.80% | 7.66% | 7.98% | 7.35% |
| Tracking error, bootstrap | 1.39% | 1.14% | 1.13% | 1.56% | 1.14% |
| Fee | 0.22% | 0.14% | 0.13% | 0.34% | 0.12% |
| US$150 oil (central) | -7.7% | -8.3% | -9.1% | -8.2% | -8.4% |
| Transition cost | A$4.4m | A$2.0m | A$1.8m | A$5.2m | A$1.7m |
| Brief constraints | Met | Met | Met | Not met (TE) | Met |

* When the alternatives are optimised rather than fixed, the optimiser holds no illiquids and 2.4%
  commodities.
* Holding alternatives at benchmark weights breaches the bootstrap TE limit and raises fees.
* The core five cannot reach the 1.1% calibrated TE (its minimum is 1.38%), so it is optimised at
  the brief's 1.5% limit.

## Charts (report-ready theme)

All 21 charts are regenerated from one specification (`tools/charts.py`):
* **Consistent colours:** Recommended `#2B5C9E`, current fund `#D9641E`, benchmark `#18998A`, return
  target `#D99E00` (dashed), other series `#6B5CA5` and `#5C9BD6`, reference lines grey (dashed).
  The palette passes the dataviz validator: lightness band, chroma floor, colour-blind separation
  and normal-vision floor.
* **Legends:** below the plot area. They previously overlaid the data on every chart.
* **Labels:** shown only on the series the chart is about, usually the recommendation. On the
  frontier chart only the current fund, benchmark and recommendation are labelled, and the other
  methods are told apart by marker shape.
* **Fills:** explicit fills on every bar. The stress-test bars previously showed white in Excel.
* **Sizing and fonts:** Arial throughout, light gridlines, category labels kept at the bottom or
  left even when values are negative. Charts are 6.3 in wide to fit an A4 page with 1-inch margins.
* **Report figures:** the charts carry no in-chart title. Use the "Figure N." row above each one as
  the caption in the report.
* **Data fixes:** the histogram is labelled by bin lower edge. Figure 7 now plots compound returns
  against the (compound) target. The frontier axes say "arithmetic". The growth chart axis reads
  "net of fees".

## Inputs to verify before submission

`Forecast` section 4 flags these as "verify":
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
