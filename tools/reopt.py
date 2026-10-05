"""Re-run every forecast-dependent Solver problem and paste the weights into the workbook.

Usage: python reopt.py <built.xlsx (formulas)> <calculated.xlsx (values)> <out.xlsx>

The Python problems reproduce the pasted v3.4 Solver weights to within 0.01% (see tools/optim.py),
so they are used here to refresh the pasted values after the forecast change. Sets that do not
depend on the forecasts (Naive MVO uses history, risk parity, the 2006-16 out-of-sample fits) are
left unchanged.
"""
import sys

import numpy as np
import openpyxl

from optim import frontier, load, solve

ROWS = range(6, 17)


def shrink(S, delta):
    sd = np.sqrt(np.diag(S))
    C = S / np.outer(sd, sd)
    n = len(sd)
    rbar = (C.sum() - n) / (n * (n - 1))
    F = rbar * np.outer(sd, sd)
    np.fill_diagonal(F, np.diag(S))
    return (1 - delta) * S + delta * F


def resample(d, n_sets=150, T=80, seed=20261016, delta=0.35, fix=True):
    """Michaud resampling: simulate T quarters from the forecast distribution, re-estimate means and
    covariance (with the same constant-correlation shrinkage), re-optimise, average the weights."""
    rng = np.random.default_rng(seed)
    mu_q, S_q = d["mu"] / 4, d["S_shr"] / 4
    W = []
    for _ in range(n_sets):
        R = rng.multivariate_normal(mu_q, S_q, size=T)
        mu_hat = R.mean(0) * 4
        S_hat = shrink(np.cov(R, rowvar=False) * 4, delta)
        w = solve(d, mu=mu_hat, S=S_hat, fix_stage2=fix)
        if w is not None:
            W.append(w)
    W = np.array(W)
    return W.mean(0), W.std(0, ddof=1), np.percentile(W, 10, axis=0), np.percentile(W, 90, axis=0), len(W)


SCORECARD = [0.03, 0.05, 0.0, 0.01]     # Stage 2 scorecard weights (comparison case)


def main(built, calc, out, mode="v35"):
    d, _ = load(calc)
    v37 = mode == "v37"
    v36 = mode in ("v36", "v37")
    FIX = not v36                       # v3.6: alternatives optimised everywhere
    wb = openpyxl.load_workbook(built)
    M, V, P = wb["Methods"], wb["Validation"], wb["Portfolio"]

    def put(ws, col, r0, w):
        """Paste at 4 dp (as Solver output was pasted) with the rounding residual on the largest
        weight so every set sums to exactly 100%."""
        w = np.round(np.asarray(w, float), 4)
        w[np.argmax(w)] += round(1 - w.sum(), 4)
        for i, x in enumerate(w):
            ws[f"{col}{r0 + i}"] = round(float(x), 4)

    res = {}
    res["robust"] = solve(d, fix_stage2=FIX)
    res["sample"] = solve(d, S=d["S_smp"], fix_stage2=FIX)
    res["minvar"] = solve(d, objective="minvar", fix_stage2=FIX)
    mean, sd, p10, p90, n = resample(d, fix=FIX)
    res["resampled"] = mean
    put(M, "D", 6, res["robust"])
    put(M, "E", 6, res["sample"])
    put(M, "F", 6, mean / mean.sum())
    put(M, "G", 6, res["minvar"])
    for i in range(11):
        if FIX and i in (4, 5, 6, 7):        # Stage 2 fixed: no spread
            continue
        M[f"F{41 + i}"] = round(float(sd[i]), 4)
        M[f"G{41 + i}"] = round(float(p10[i]), 4)
        M[f"H{41 + i}"] = round(float(p90[i]), 4)
    M["A39"] = f"4. Resampling spread ({n} resampled input sets)"

    d_sc = dict(d)
    d_sc["stage2"] = np.array(SCORECARD)
    put(V, "C", 30, solve(d, S=d["S_smp"], fix_stage2=FIX))           # unshrunk covariance
    put(V, "D", 30, solve(d, te=0.0075, fix_stage2=FIX))              # tight TE
    put(V, "E", 30, solve(d, te=0.015, fix_stage2=FIX))               # loose TE

    put(P, "C", 43, solve(d, te=0.015, fix_stage2=FIX))
    put(P, "D", 43, solve(d, te=False, fix_stage2=FIX))
    put(P, "E", 43, solve(d, eqmax=0.72, fix_stage2=FIX))
    put(P, "F", 43, solve(d_sc) if v36 else solve(d, fix_stage2=False))
    put(P, "G", 43, solve(d, te=False, fix_stage2=False, longonly_only=True))
    if v37:                                                            # cost of the oil floor
        put(P, "H", 43, solve(d, fix_stage2=False, oil=False))
        put(P, "I", 43, solve(d, fix_stage2=False, oil=float(d["bench"] @ d["oil"])))

    for r0, te in ((113, False), (131, d["te"])):
        pts = frontier(d, te=te, fix=FIX)
        for k, w in enumerate(pts):
            put(V, "BCDEFGHIJK"[k], r0, w)

    # asset-class scope portfolios (Portfolio section 9; weights rows 127-137, columns C:F)
    def restricted(excl):
        d2 = dict(d)
        ub, lb = d["ub"].copy(), d["lb"].copy()
        for i in excl:
            ub[i] = lb[i] = 0
        d2["ub"], d2["lb"] = ub, lb
        return d2
    d_bench_alts = dict(d)
    d_bench_alts["stage2"] = d["bench"][[4, 5, 6, 7]]
    scope = {"C": lambda **k: solve(d_sc, **k) if v36 else solve(d, fix_stage2=False, **k),
             "D": lambda **k: solve(restricted([4, 5, 6, 7]), fix_stage2=False, **k),
             "E": lambda **k: solve(d_bench_alts, **k),
             "F": lambda **k: solve(restricted([2, 3, 4, 5, 6, 7]), fix_stage2=False, te=0.015, **k)}
    for j, (col, f) in enumerate(scope.items()):
        w = f()
        if w is None and v37:          # oil floor not attainable with this asset set
            w = f(oil=False)
            P[f"B{143 + j}"] = P[f"B{143 + j}"].value + "; oil floor not attainable, solved without it"
            print("scope", col, "solved without the oil floor")
        put(P, col, 127, w)

    wb.save(out)
    w = res["robust"]
    print("robust weights:", np.round(w, 4), "| max gap to resampled %.4f" % np.max(np.abs(w - mean / mean.sum())))
    v = np.sqrt(w @ d["S_shr"] @ w)
    print("expected return (arith) %.4f comp %.5f vol %.4f TE %.4f" % (
        d["mu"] @ w, d["mu"] @ w - v * v / 2, v, np.sqrt((w - d["bench"]) @ d["S_shr"] @ (w - d["bench"]))))
    if "oil" in d:
        print("oil central %.4f floor %.4f" % (d["oil"] @ w, d["oilmin"]))


if __name__ == "__main__":
    main(*sys.argv[1:5])
