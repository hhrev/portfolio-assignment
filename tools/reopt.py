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


def resample(d, n_sets=150, T=80, seed=20261016, delta=0.35):
    """Michaud resampling: simulate T quarters from the forecast distribution, re-estimate means and
    covariance (with the same constant-correlation shrinkage), re-optimise, average the weights."""
    rng = np.random.default_rng(seed)
    mu_q, S_q = d["mu"] / 4, d["S_shr"] / 4
    W = []
    for _ in range(n_sets):
        R = rng.multivariate_normal(mu_q, S_q, size=T)
        mu_hat = R.mean(0) * 4
        S_hat = shrink(np.cov(R, rowvar=False) * 4, delta)
        w = solve(d, mu=mu_hat, S=S_hat)
        if w is not None:
            W.append(w)
    W = np.array(W)
    return W.mean(0), W.std(0, ddof=1), np.percentile(W, 10, axis=0), np.percentile(W, 90, axis=0), len(W)


def main(built, calc, out):
    d, _ = load(calc)
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
    res["robust"] = solve(d)
    res["sample"] = solve(d, S=d["S_smp"])
    res["minvar"] = solve(d, objective="minvar")
    mean, sd, p10, p90, n = resample(d)
    res["resampled"] = mean
    put(M, "D", 6, res["robust"])
    put(M, "E", 6, res["sample"])
    put(M, "F", 6, mean / mean.sum())
    put(M, "G", 6, res["minvar"])
    for i in range(11):
        if i in (4, 5, 6, 7):        # Stage 2 fixed: no spread
            continue
        M[f"F{41 + i}"] = round(float(sd[i]), 4)
        M[f"G{41 + i}"] = round(float(p10[i]), 4)
        M[f"H{41 + i}"] = round(float(p90[i]), 4)
    M["A39"] = f"4. Resampling spread ({n} resampled input sets)"

    put(V, "C", 30, solve(d, S=d["S_smp"]))           # unshrunk covariance
    put(V, "D", 30, solve(d, te=0.0075))              # tight TE
    put(V, "E", 30, solve(d, te=0.015))               # loose TE

    put(P, "C", 43, solve(d, te=0.015))
    put(P, "D", 43, solve(d, te=False))
    put(P, "E", 43, solve(d, eqmax=0.72))
    put(P, "F", 43, solve(d, fix_stage2=False))
    put(P, "G", 43, solve(d, te=False, fix_stage2=False, longonly_only=True))

    for r0, te in ((113, False), (131, d["te"])):
        pts = frontier(d, te=te)
        for k, w in enumerate(pts):
            put(V, "BCDEFGHIJK"[k], r0, w)

    wb.save(out)
    w = res["robust"]
    print("robust weights:", np.round(w, 4))
    print("expected return (arith) %.4f  TE %.4f" % (d["mu"] @ w,
                                                    np.sqrt((w - d["bench"]) @ d["S_shr"] @ (w - d["bench"]))))


if __name__ == "__main__":
    main(*sys.argv[1:4])
