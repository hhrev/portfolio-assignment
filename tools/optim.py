"""Python replica of the workbook's Solver problems (used to re-optimise after forecast changes).

Robust MVO (the recommendation): maximise net expected return subject to
  parametric TE vs benchmark <= TE limit (shrunk covariance), weights sum to 1,
  asset bounds (Inputs E:F), Stage 2 weights fixed, equity within band, illiquids <= max, cash >= min.
From v3.7 the Stage 2 upper bounds are capped at benchmark + the overweight cap (Inputs B47) and the
US$150 oil central-case loss must be no worse than the floor in Inputs B50 (Oil!F28:F38 responses).
"""
import warnings

import numpy as np
import openpyxl
from scipy.optimize import minimize

warnings.filterwarnings("ignore")
STAGE2 = [4, 5, 6, 7]   # Commodities, Direct Property, Hedge Funds, Private Equity


def load(path):
    wb = openpyxl.load_workbook(path, data_only=True)
    inp, risk, port = wb["Inputs"], wb["Risk"], wb["Portfolio"]
    g = lambda ws, rng: np.array([[c.value for c in row] for row in ws[rng]], dtype=float)
    d = {
        "mu": g(risk, "B54:L54")[0],
        "S_shr": g(risk, "B41:L51"),
        "S_smp": g(risk, "B5:L15"),
        "bench": g(inp, "C5:C15")[:, 0],
        "lb": g(inp, "E5:E15")[:, 0], "ub": g(inp, "F5:F15")[:, 0],
        "eq": g(inp, "G5:G15")[:, 0], "ill": g(inp, "H5:H15")[:, 0],
        "stage2": g(port, "H23:H26")[:, 0],
        "te": inp["B45"].value, "eqmin": inp["B31"].value, "eqmax": inp["B32"].value,
        "illmax": inp["B33"].value, "cashmin": inp["B34"].value,
    }
    if isinstance(inp["B50"].value, (int, float)) and "floor" in str(inp["A50"].value).lower():
        d["oil"] = g(wb["Oil"], "F28:F38")[:, 0]
        d["oilmin"] = inp["B50"].value
        cap = inp["B47"].value
        for i in STAGE2:
            d["ub"][i] = min(d["ub"][i], d["bench"][i] + cap)
    return d, wb


def solve(d, mu=None, S=None, te=None, eqmax=None, fix_stage2=True, longonly_only=False, x0=None,
          objective="return", target=None, oil=None):
    """oil: None = the floor in d (if any), False = no oil floor, a number = that floor.
    Returns None if no start converges (e.g. the problem is infeasible)."""
    mu = d["mu"] if mu is None else mu
    S = d["S_shr"] if S is None else S
    te = d["te"] if te is None else (None if te is False else te)
    eqmax = d["eqmax"] if eqmax is None else eqmax
    b = d["bench"]
    n = len(mu)
    cons = [{"type": "eq", "fun": lambda w: w.sum() - 1}]
    if not longonly_only:
        cons += [{"type": "ineq", "fun": lambda w: d["eq"] @ w - d["eqmin"]},
                 {"type": "ineq", "fun": lambda w: eqmax - d["eq"] @ w},
                 {"type": "ineq", "fun": lambda w: d["illmax"] - d["ill"] @ w}]
    cons += [{"type": "ineq", "fun": lambda w: w[10] - d["cashmin"]}]
    if te is not None:
        cons.append({"type": "ineq", "fun": lambda w: te ** 2 - (w - b) @ S @ (w - b)})
    floor = d.get("oilmin") if oil is None else (None if oil is False else oil)
    if floor is not None and not longonly_only:
        cons.append({"type": "ineq", "fun": lambda w: d["oil"] @ w - floor})
    if target is not None:
        cons.append({"type": "ineq", "fun": lambda w: mu @ w - target})
    if longonly_only:
        bounds = [(0, 1)] * n
    else:
        bounds = list(zip(d["lb"], d["ub"]))
    if fix_stage2:
        for k, i in enumerate(STAGE2):
            bounds[i] = (d["stage2"][k], d["stage2"][k])
    if objective == "return":
        f = lambda w: -(mu @ w)
        jac = lambda w: -mu
    elif objective == "te":   # minimum tracking error
        f = lambda w: (w - b) @ S @ (w - b)
        jac = lambda w: 2 * S @ (w - b)
    else:   # minimum variance
        f = lambda w: w @ S @ w
        jac = lambda w: 2 * S @ w
    best = None
    starts = [x0] if x0 is not None else []
    starts += [b.copy(), np.clip(b, [x[0] for x in bounds], [x[1] for x in bounds])]
    for s in starts:
        s = np.clip(s, [x[0] for x in bounds], [x[1] for x in bounds])
        r = minimize(f, s, jac=jac, bounds=bounds, constraints=cons, method="SLSQP",
                     options={"maxiter": 2000, "ftol": 1e-12})
        if r.success and (best is None or r.fun < best.fun):
            best = r
    return None if best is None else best.x


def frontier(d, S=None, te=None, n=10, fix=True):
    S = d["S_shr"] if S is None else S
    w_lo = solve(d, S=S, te=te, objective="minvar", fix_stage2=fix)
    w_hi = solve(d, S=S, te=te, fix_stage2=fix)
    r_lo, r_hi = d["mu"] @ w_lo, d["mu"] @ w_hi
    pts = []
    for k in range(n):
        t = r_lo + (r_hi - r_lo) * k / (n - 1)
        if k == 0:
            pts.append(w_lo)
        elif k == n - 1:
            pts.append(w_hi)
        else:
            pts.append(solve(d, S=S, te=te, objective="minvar", target=t - 1e-9, fix_stage2=fix))
    return pts
