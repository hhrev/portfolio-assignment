"""v3.6 changes applied on top of the v3.5 build (before re-optimisation).

* Alternatives are optimised (0 to upper bound) rather than fixed; the Stage 2 scorecard is kept
  as a cross-check on the optimised weights.
* Australian fixed income floor raised to 12% (robust vs resampled stability test, Methods 3).
* Transition plan tested against the 1.5% brief TE limit.
* Forecast: rationale text removed (it belongs in the report), source labels tidied, and a
  sensitivity table added for the P/E reversion assumption.
* Arithmetic and compound returns labelled wherever they appear.
"""
import sys
from copy import copy

import openpyxl
from openpyxl.styles import Alignment, Border, Font, PatternFill

from style import BROWN, box

ASSETS = ["Australian Equities", "World Equities", "Emerging Markets", "Listed Property",
          "Commodities", "Direct Property", "Hedge Funds", "Private Equity",
          "Australian Fixed Income", "World Fixed Income", "Cash"]
PCT2 = "0.00%"
AFI_FLOOR = 0.12


def find_row(ws, prefix, col=1, start=1, end=400):
    for r in range(start, end):
        v = ws.cell(r, col).value
        if isinstance(v, str) and v.startswith(prefix):
            return r
    raise KeyError(prefix)


def blank(c):
    c.value = None
    c.border = Border()
    c.fill = PatternFill(fill_type=None)
    c.font = Font(name="Arial", size=10)
    c.number_format = "General"
    c.alignment = Alignment()


def inputs(wb):
    ws = wb["Inputs"]
    ws["E13"] = AFI_FLOOR
    ws["A18"] = ("Australian fixed income floor 12%: home-currency defensive anchor; keeps the bond split "
                 "stable against resampling (Methods section 3).")
    ws["A18"].font = Font(name="Arial", size=9, italic=True, color="FF595959")


def portfolio(wb):
    ws = wb["Portfolio"]
    for k in range(11):
        ws[f"E{6 + k}"] = f"=Methods!I{6 + k}"
        if k in (4, 5, 6, 7):
            ws[f"B{6 + k}"] = "2: optimised"
    ws["A20"] = "2. Stage 2 scorecard (1 to 5): cross-check on the optimised alternatives weights"
    ws["H22"] = "Optimised weight"
    for j, r in enumerate(range(23, 27)):
        ws[f"H{r}"] = f"=E{10 + j}"
    ws["F42"] = "Alternatives fixed by scorecard"
    ws["A32"] = "Expected return, arithmetic (maximise)"
    ws["A55"] = "Expected return (arithmetic)"
    ws["K5"] = "Net forecast (arithmetic)"
    # transition plan: test against the 1.5% brief limit
    r = find_row(ws, "Brief constraints (TE 1.1%")
    ws[f"A{r}"] = "Brief constraints (TE 1.5% brief limit, equity band, illiquid, cash)"
    for c in "CDEF":
        v = ws[f"{c}{r}"].value
        ws[f"{c}{r}"] = v.replace("Inputs!$B$45+0.00001", "Inputs!$B$27+0.00001")


def methods(wb):
    ws = wb["Methods"]
    ws["A3"] = ("1. Method weights (alternatives optimised; Australian fixed income floor 12%; risk parity "
                "holds the scorecard alternatives)")
    ws["B21"] = "Expected return (arithmetic)"
    # alternatives rows were links to the scorecard; they are now optimiser output (re-pasted by
    # reopt.py), with risk parity holding the scorecard weights
    for j, v in enumerate([0.03, 0.05, 0.0, 0.01]):
        for c in "DEFGH":
            ws[f"{c}{10 + j}"] = v
    for i in range(4, 8):          # spread rows for the alternatives now come from the resampling
        r = 41 + i
        ws[f"B{r}"] = f"=D{6 + i}"
        ws[f"C{r}"] = f"=F{6 + i}"
        ws[f"D{r}"] = f"=B{r}-C{r}"
        ws[f"E{r}"] = f"=ABS(D{r})"
        ws[f"I{r}"] = f'=IF(E{r}<=F{r},"Yes","No")'
        for c in "BCDEFGHI":
            ws[f"{c}{r}"]._style = copy(ws[f"{c}41"]._style)
    wb["Risk"]["N55"] = "Expected return (arithmetic)"
    wb["Validation"]["B55"] = "Expected return (arithmetic)"
    wb["Report figures"]["O281"] = "Expected return (arithmetic)"


def forecast(wb):
    f = wb["Forecast"]
    # source labels
    relabel = {
        "Australian equities: long-run average forward P/E": "Judgement: long-run average forward P/E (IG, ASX 200 outlook 2026: 14.8x)",
        "World equities: long-run average forward P/E": "Judgement: long-run average forward P/E for MSCI World",
        "Emerging markets: dividend yield": "iShares MSCI EM ETF trailing 12-month yield, May-2026",
        "Emerging markets: long-run average forward P/E": "Judgement: long-run average forward P/E for MSCI EM",
    }
    top4 = find_row(f, "4. Market and economic inputs")
    rows = {}
    for r in range(top4 + 2, top4 + 60):
        lab = f.cell(r, 1).value
        if lab in relabel:
            f.cell(r, 3).value = relabel[lab]
        if isinstance(lab, str):
            rows[lab] = r
    # remove the rationale section (the reasoning belongs in the report)
    t9 = find_row(f, "9. Forecast rationale")
    text = []
    for r in range(t9 + 2, t9 + 13):
        text.append((f.cell(r, 1).value, f.cell(r, 2).value, f.cell(r, 7).value))
    for rng in list(f.merged_cells.ranges):
        if rng.min_row >= t9:
            f.unmerge_cells(str(rng))
    for r in range(t9, t9 + 14):
        for c in range(1, 10):
            blank(f.cell(r, c))
        f.row_dimensions[r].height = None

    # ---- 9. Sensitivity to the P/E reversion assumption ----
    t5b = find_row(f, "5b.")
    t6 = find_row(f, "6. Blend")
    eq = {"Australian Equities": (t5b + 2, 6, "ae"), "World Equities": (t5b + 3, 7, "we"),
          "Emerging Markets": (t5b + 4, 8, "em")}
    K = lambda lab: f"$B${rows[lab]}"
    names = {"ae": "Australian equities", "we": "World equities", "em": "Emerging markets"}
    shares = [0, 0.25, 0.5, 0.75, 1.0]
    tbl = []
    top = t9
    for i, x in enumerate(shares):
        r = top + 2 + i
        tbl.append([x, None, None, None, None])
    last = box(f, top, 1, "9. Sensitivity: share of the P/E gap closed over 5 years (the most influential "
               "forecast judgement)", ["Share closed", "Australian equities net (compound)",
                                       "World equities net (compound)", "Emerging markets net (compound)",
                                       "Recommended portfolio (compound, weights held)"],
               tbl, BROWN, fmts=["0%", PCT2, PCT2, PCT2, PCT2], header_height=40, inputs={(None, 0)})

    def post(asset, x_cell):
        """Gross arithmetic posterior for one equity class at a given share closed."""
        g, s1, p = eq[asset]
        lr = rows[f"{names[p]}: long-run average forward P/E"]
        cur = rows[f"{names[p]}: forward P/E"]
        gk = (f"({K(names[p] + ': ' + ('forward dividend yield' if p != 'em' else 'dividend yield'))}"
              f"-{K(names[p] + ': net dilution')}+{K(names[p] + ': real earnings growth')}"
              f"+{K('Australian inflation, 5-year average')}+($B${lr}/$B${cur})^({x_cell}/5)-1)")
        vb = f"$I${g}"
        comp = f"(({gk}+{vb})/2)"
        arith = f"({comp}+$H${s1}^2/2)"
        floor = f"$H${t6 + 2 + ASSETS.index(asset)}"
        disp = f"MAX(ABS({gk}-{vb})/SQRT(2),{floor})"
        conf = f"(Inputs!$B$42-(Inputs!$B$42-Inputs!$B$43)*MIN(1,{disp}/Inputs!$B$44))"
        return f"({conf}*{arith}+(1-{conf})*$C${s1})"
    for i in range(len(shares)):
        r = top + 2 + i
        x = f"$A{r}"
        ae, we, em = post("Australian Equities", x), post("World Equities", x), post("Emerging Markets", x)
        ae_g = f"IF(Inputs!$B$46=1,({ae}+{we})/2,{ae})"
        we_g = f"IF(Inputs!$B$46=1,({ae}+{we})/2,{we})"
        f[f"B{r}"] = f"={ae_g}-$F$6-$H$6^2/2"
        f[f"C{r}"] = f"={we_g}-$F$7-$H$7^2/2"
        f[f"D{r}"] = f"={em}-$F$8-$H$8^2/2"
        f[f"E{r}"] = (f"=Summary!$D$24+Portfolio!$E$6*(B{r}+$H$6^2/2-$G$6)+Portfolio!$E$7*(C{r}+$H$7^2/2-$G$7)"
                      f"+Portfolio!$E$8*(D{r}+$H$8^2/2-$G$8)")
    f[f"A{last + 1}"] = ("The 50% row reproduces section 1. Moving from 50% to 0% (no reversion) or 100% (full "
                         "reversion) shows how much of the equity forecast rests on this judgement.")
    f[f"A{last + 1}"].font = Font(name="Arial", size=9, italic=True, color="FF595959")
    return text


def main(src, dst, notes):
    wb = openpyxl.load_workbook(src)
    inputs(wb)
    portfolio(wb)
    methods(wb)
    text = forecast(wb)
    wb.calculation.fullCalcOnLoad = True
    wb.save(dst)
    with open(notes, "w") as fh:
        fh.write("# Forecast rationale notes (removed from the model; rewrite in your own words for the report)\n\n")
        for a, m, w in text:
            if a:
                fh.write(f"## {a}\n\n- Method and drivers: {m}\n- Why it differs from history: {w}\n\n")


if __name__ == "__main__":
    main(*sys.argv[1:4])
