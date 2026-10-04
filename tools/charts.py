"""Report-ready chart theme: every chart in the workbook is regenerated from one specification.

Theme (validated with the dataviz palette checker: lightness band, chroma floor, colour-blind
separation and normal-vision floor all pass on a white surface):
  Recommended  #2B5C9E   Current fund  #D9641E   Benchmark  #18998A
  Target       #D99E00 (dashed)   Other series  #6B5CA5, #5C9BD6   Reference lines  #8C8C8C
Rules: legend below the plot (never overlaying data); labels only on the series the chart is about;
light gridlines; Arial; charts sized for an A4 page with 1-inch margins (6.3 in wide).
"""
from xml.sax.saxutils import escape

NAVY, ORANGE, TEAL, GOLD, VIOLET, LBLUE = "2B5C9E", "D9641E", "18998A", "D99E00", "6B5CA5", "5C9BD6"
REF_GREY, TEXT, TEXT2, GRID, AXIS = "8C8C8C", "262626", "595959", "E3E3E3", "BFBFBF"
NS = ('xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart" '
      'xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" '
      'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"')
RF = "'Report figures'!"


def _txpr(sz=900, color=TEXT2, bold=False, rot=None):
    r = f' rot="{rot}" vert="horz"' if rot is not None else ""
    b = ' b="1"' if bold else ' b="0"'
    return (f'<c:txPr><a:bodyPr{r}/><a:lstStyle/><a:p><a:pPr><a:defRPr sz="{sz}"{b}>'
            f'<a:solidFill><a:srgbClr val="{color}"/></a:solidFill><a:latin typeface="Arial"/>'
            f'<a:cs typeface="Arial"/></a:defRPr></a:pPr><a:endParaRPr lang="en-AU"/></a:p></c:txPr>')


def _rich(text, sz=900, color=TEXT2, bold=False, rot=None):
    r = f' rot="{rot}" vert="horz"' if rot is not None else ""
    b = ' b="1"' if bold else ' b="0"'
    return (f'<c:tx><c:rich><a:bodyPr{r}/><a:lstStyle/><a:p><a:pPr><a:defRPr sz="{sz}"{b}>'
            f'<a:solidFill><a:srgbClr val="{color}"/></a:solidFill><a:latin typeface="Arial"/></a:defRPr></a:pPr>'
            f'<a:r><a:rPr lang="en-AU" sz="{sz}"{b}><a:solidFill><a:srgbClr val="{color}"/></a:solidFill>'
            f'<a:latin typeface="Arial"/></a:rPr><a:t>{escape(text)}</a:t></a:r></a:p></c:rich></c:tx>')


def _title(text, sz=900, rot=None):
    return f'<c:title>{_rich(text, sz=sz, rot=rot)}<c:overlay val="0"/></c:title>'


def _line_sp(color, w=22225, dash="solid"):
    return (f'<c:spPr><a:ln w="{w}" cap="rnd"><a:solidFill><a:srgbClr val="{color}"/></a:solidFill>'
            f'<a:prstDash val="{dash}"/><a:round/></a:ln></c:spPr>')


def _fill_sp(color):
    return (f'<c:spPr><a:solidFill><a:srgbClr val="{color}"/></a:solidFill>'
            f'<a:ln w="9525"><a:solidFill><a:srgbClr val="FFFFFF"/></a:solidFill></a:ln></c:spPr>')


def _tx(name):
    if name.startswith("="):
        return f'<c:tx><c:strRef><c:f>{escape(name[1:])}</c:f></c:strRef></c:tx>'
    return f'<c:tx><c:v>{escape(name)}</c:v></c:tx>'


def _dlbls(fmt="0.0%", pos="outEnd", ser_name=False, sz=800):
    return (f'<c:dLbls><c:numFmt formatCode="{escape(fmt)}" sourceLinked="0"/>'
            f'<c:spPr><a:noFill/><a:ln><a:noFill/></a:ln></c:spPr>{_txpr(sz, TEXT)}'
            f'<c:dLblPos val="{pos}"/><c:showLegendKey val="0"/><c:showVal val="{0 if ser_name else 1}"/>'
            f'<c:showCatName val="0"/><c:showSerName val="{1 if ser_name else 0}"/><c:showPercent val="0"/>'
            f'<c:showBubbleSize val="0"/></c:dLbls>')


def _no_dlbls():
    return ('<c:dLbls><c:showLegendKey val="0"/><c:showVal val="0"/><c:showCatName val="0"/>'
            '<c:showSerName val="0"/><c:showPercent val="0"/><c:showBubbleSize val="0"/></c:dLbls>')


def _ref(kind, f):
    tag = "numRef" if kind == "num" else "strRef"
    return f"<c:{tag}><c:f>{escape(f)}</c:f></c:{tag}>"


def _cat(f, kind="str"):
    return f"<c:cat>{_ref(kind, f)}</c:cat>"


def _val(f):
    return f"<c:val>{_ref('num', f)}</c:val>"


# ---------------------------------------------------------------------------------------------
def _axes_cat_val(cat_id, val_id, horizontal=False, val_fmt="0%", val_title=None, cat_title=None,
                  date=False, date_fmt="yyyy", label_low=True, val_min=None, val_max=None, major=None,
                  cat_fmt="General"):
    cat_pos, val_pos = ("l", "b") if horizontal else ("b", "l")
    orient = "maxMin" if horizontal else "minMax"
    tag = "dateAx" if date else "catAx"
    ct = _title(cat_title, rot=-5400000 if horizontal else None) if cat_title else ""
    vt = _title(val_title, rot=None if horizontal else -5400000) if val_title else ""
    cat = (f'<c:{tag}><c:axId val="{cat_id}"/><c:scaling><c:orientation val="{orient}"/></c:scaling>'
           f'<c:delete val="0"/><c:axPos val="{cat_pos}"/>{ct}'
           f'<c:numFmt formatCode="{escape(date_fmt if date else cat_fmt)}" sourceLinked="0"/>'
           f'<c:majorTickMark val="none"/><c:minorTickMark val="none"/>'
           f'<c:tickLblPos val="{"low" if label_low else "nextTo"}"/>'
           f'<c:spPr><a:ln w="9525"><a:solidFill><a:srgbClr val="{AXIS}"/></a:solidFill></a:ln></c:spPr>'
           f'{_txpr(900)}<c:crossAx val="{val_id}"/><c:crosses val="autoZero"/><c:auto val="1"/>')
    if date:
        cat += ('<c:lblOffset val="100"/><c:baseTimeUnit val="months"/><c:majorUnit val="2"/>'
                '<c:majorTimeUnit val="years"/>')
    else:
        cat += '<c:lblAlgn val="ctr"/><c:lblOffset val="100"/><c:noMultiLvlLbl val="0"/>'
    cat += f'</c:{tag}>'
    sc = '<c:orientation val="minMax"/>'
    if val_max is not None:
        sc += f'<c:max val="{val_max}"/>'
    if val_min is not None:
        sc += f'<c:min val="{val_min}"/>'
    val = (f'<c:valAx><c:axId val="{val_id}"/><c:scaling>{sc}</c:scaling><c:delete val="0"/>'
           f'<c:axPos val="{val_pos}"/><c:majorGridlines><c:spPr><a:ln w="6350"><a:solidFill>'
           f'<a:srgbClr val="{GRID}"/></a:solidFill></a:ln></c:spPr></c:majorGridlines>{vt}'
           f'<c:numFmt formatCode="{escape(val_fmt)}" sourceLinked="0"/><c:majorTickMark val="none"/>'
           f'<c:minorTickMark val="none"/><c:tickLblPos val="nextTo"/><c:spPr><a:ln><a:noFill/></a:ln></c:spPr>'
           f'{_txpr(900)}<c:crossAx val="{cat_id}"/>'
           f'<c:crosses val="{"max" if horizontal else "autoZero"}"/><c:crossBetween val="between"/>'
           + (f'<c:majorUnit val="{major}"/>' if major else "") + '</c:valAx>')
    return cat + val


def _wrap(plot, title=None, legend=True):
    t = _title(title, sz=1100) if title else ""
    atd = '<c:autoTitleDeleted val="0"/>' if title else '<c:autoTitleDeleted val="1"/>'
    leg = (f'<c:legend><c:legendPos val="b"/><c:overlay val="0"/><c:spPr><a:noFill/><a:ln><a:noFill/></a:ln>'
           f'</c:spPr>{_txpr(900, TEXT)}</c:legend>') if legend else ""
    return (f'<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<c:chartSpace {NS}>'
            f'<c:roundedCorners val="0"/><c:chart>{t}{atd}<c:plotArea><c:layout/>{plot}'
            f'<c:spPr><a:noFill/><a:ln><a:noFill/></a:ln></c:spPr></c:plotArea>{leg}'
            f'<c:plotVisOnly val="1"/><c:dispBlanksAs val="gap"/></c:chart>'
            f'<c:spPr><a:solidFill><a:srgbClr val="FFFFFF"/></a:solidFill><a:ln><a:noFill/></a:ln></c:spPr>'
            f'{_txpr(900)}</c:chartSpace>')


def bar(series, cats, horizontal=False, val_fmt="0%", lbl_fmt="0.0%", val_title=None, cat_title=None,
        lines=(), gap=60, overlap=-8, title=None, legend=True, val_min=None, val_max=None, major=None,
        cat_kind="str", cat_fmt="General"):
    """series: list of (name, values_ref, colour, label?)  lines: list of (name, ref, colour, dash, markers)"""
    s = ""
    for i, (name, ref, col, lbl) in enumerate(series):
        s += (f'<c:ser><c:idx val="{i}"/><c:order val="{i}"/>{_tx(name)}{_fill_sp(col)}'
              f'<c:invertIfNegative val="0"/>{_dlbls(lbl_fmt) if lbl else ""}{_cat(cats, cat_kind)}{_val(ref)}</c:ser>')
    plot = (f'<c:barChart><c:barDir val="{"bar" if horizontal else "col"}"/><c:grouping val="clustered"/>'
            f'<c:varyColors val="0"/>{s}<c:gapWidth val="{gap}"/><c:overlap val="{overlap}"/>'
            f'<c:axId val="1001"/><c:axId val="1002"/></c:barChart>')
    if lines:
        ls = ""
        for j, (name, ref, col, dash, mk) in enumerate(lines):
            i = len(series) + j
            marker = (f'<c:marker><c:symbol val="dash"/><c:size val="20"/><c:spPr><a:solidFill><a:srgbClr val="{col}"/>'
                      f'</a:solidFill><a:ln><a:solidFill><a:srgbClr val="{col}"/></a:solidFill></a:ln></c:spPr></c:marker>'
                      if mk else '<c:marker><c:symbol val="none"/></c:marker>')
            sp = '<c:spPr><a:ln><a:noFill/></a:ln></c:spPr>' if mk else _line_sp(col, 19050, dash)
            ls += (f'<c:ser><c:idx val="{i}"/><c:order val="{i}"/>{_tx(name)}{sp}{marker}'
                   f'{_cat(cats)}{_val(ref)}<c:smooth val="0"/></c:ser>')
        plot += (f'<c:lineChart><c:grouping val="standard"/><c:varyColors val="0"/>{ls}'
                 f'<c:marker val="1"/><c:axId val="1001"/><c:axId val="1002"/></c:lineChart>')
    plot += _axes_cat_val(1001, 1002, horizontal, val_fmt, val_title, cat_title, val_min=val_min,
                          val_max=val_max, major=major, cat_fmt=cat_fmt)
    return _wrap(plot, title, legend)


def line(series, cats, val_fmt="0%", val_title=None, date=True, title=None, date_fmt="yyyy"):
    """series: list of (name, ref, colour, dash, width)"""
    s = ""
    for i, (name, ref, col, dash, w) in enumerate(series):
        s += (f'<c:ser><c:idx val="{i}"/><c:order val="{i}"/>{_tx(name)}{_line_sp(col, w, dash)}'
              f'<c:marker><c:symbol val="none"/></c:marker>{_cat(cats, "num" if date else "str")}{_val(ref)}'
              f'<c:smooth val="0"/></c:ser>')
    plot = (f'<c:lineChart><c:grouping val="standard"/><c:varyColors val="0"/>{s}<c:marker val="1"/>'
            f'<c:axId val="1001"/><c:axId val="1002"/></c:lineChart>')
    plot += _axes_cat_val(1001, 1002, False, val_fmt, val_title, None, date=date, date_fmt=date_fmt)
    return _wrap(plot, title)


def scatter(lines_, points, x_title, y_title, title=None):
    """lines_: (name, xref, yref, colour, dash, width); points: (name, xref, yref, colour, symbol, size, label_pos)"""
    s, i = "", 0
    for name, xr, yr, col, dash, w in lines_:
        s += (f'<c:ser><c:idx val="{i}"/><c:order val="{i}"/>{_tx(name)}{_line_sp(col, w, dash)}'
              f'<c:marker><c:symbol val="none"/></c:marker><c:xVal>{_ref("num", xr)}</c:xVal>'
              f'<c:yVal>{_ref("num", yr)}</c:yVal><c:smooth val="0"/></c:ser>')
        i += 1
    for name, xr, yr, col, sym, size, pos in points:
        mk = (f'<c:marker><c:symbol val="{sym}"/><c:size val="{size}"/><c:spPr><a:solidFill><a:srgbClr val="{col}"/>'
              f'</a:solidFill><a:ln w="12700"><a:solidFill><a:srgbClr val="FFFFFF"/></a:solidFill></a:ln></c:spPr></c:marker>')
        lbl = _dlbls(pos=pos, ser_name=True, sz=800) if pos else _no_dlbls()
        s += (f'<c:ser><c:idx val="{i}"/><c:order val="{i}"/>{_tx(name)}<c:spPr><a:ln w="19050"><a:noFill/></a:ln></c:spPr>'
              f'{mk}{lbl}<c:xVal>{_ref("num", xr)}</c:xVal><c:yVal>{_ref("num", yr)}</c:yVal><c:smooth val="0"/></c:ser>')
        i += 1
    plot = (f'<c:scatterChart><c:scatterStyle val="lineMarker"/><c:varyColors val="0"/>{s}'
            f'<c:axId val="1001"/><c:axId val="1002"/></c:scatterChart>')
    ax = lambda aid, cross, pos, t, rot, grid: (
        f'<c:valAx><c:axId val="{aid}"/><c:scaling><c:orientation val="minMax"/></c:scaling><c:delete val="0"/>'
        f'<c:axPos val="{pos}"/>' + (f'<c:majorGridlines><c:spPr><a:ln w="6350"><a:solidFill><a:srgbClr val="{GRID}"/>'
                                     f'</a:solidFill></a:ln></c:spPr></c:majorGridlines>' if grid else "")
        + f'{_title(t, rot=rot)}<c:numFmt formatCode="0.0%" sourceLinked="0"/><c:majorTickMark val="none"/>'
        f'<c:minorTickMark val="none"/><c:tickLblPos val="low"/><c:spPr><a:ln w="9525"><a:solidFill>'
        f'<a:srgbClr val="{AXIS}"/></a:solidFill></a:ln></c:spPr>{_txpr(900)}<c:crossAx val="{cross}"/>'
        f'<c:crosses val="autoZero"/><c:crossBetween val="midCat"/></c:valAx>')
    plot += ax(1001, 1002, "b", x_title, None, False) + ax(1002, 1001, "l", y_title, -5400000, True)
    return _wrap(plot, title)


# ---------------------------------------------------------------------------------------------
def specs():
    """chart file -> (xml, (width_in, height_in))"""
    W, H, HT = 6.3, 3.6, 4.4
    S = {}
    port3 = lambda sheet, c0, c1, c2, cats, hdr, lbl_rec=True: [
        (f"={sheet}{c0}{hdr}", f"{sheet}${c0}${cats[0]}:${c0}${cats[1]}", ORANGE, False),
        (f"={sheet}{c1}{hdr}", f"{sheet}${c1}${cats[0]}:${c1}${cats[1]}", TEAL, False),
        (f"={sheet}{c2}{hdr}", f"{sheet}${c2}${cats[0]}:${c2}${cats[1]}", NAVY, lbl_rec)]
    # asset allocation (Summary and Figure 2): horizontal bars, label the recommendation
    S["chart1.xml"] = (bar(port3("Summary!", "B", "C", "D", (6, 16), 5), "Summary!$A$6:$A$16",
                           horizontal=True, val_fmt="0%", title="Asset allocation: current, benchmark and recommended",
                           gap=40), (W, 3.6))
    S["chart4.xml"] = (bar(port3(RF, "O", "P", "Q", (48, 58), 47), f"{RF}$N$48:$N$58", horizontal=True,
                           val_fmt="0%", gap=40), (W, HT))

    def frontier_points(extra):
        pts = [("=" + RF + "$N$33", f"{RF}$O$33", f"{RF}$P$33", ORANGE, "circle", 9, "r"),
               ("=" + RF + "$N$34", f"{RF}$O$34", f"{RF}$P$34", TEAL, "circle", 9, "b"),
               ("=" + RF + "$N$35", f"{RF}$O$35", f"{RF}$P$35", NAVY, "diamond", 11, "t")]
        if extra:
            pts += [("=" + RF + "$N$36", f"{RF}$O$36", f"{RF}$P$36", VIOLET, "triangle", 7, None),
                    ("=" + RF + "$N$37", f"{RF}$O$37", f"{RF}$P$37", VIOLET, "square", 7, None),
                    ("=" + RF + "$N$38", f"{RF}$O$38", f"{RF}$P$38", LBLUE, "triangle", 7, None),
                    ("=" + RF + "$N$39", f"{RF}$O$39", f"{RF}$P$39", LBLUE, "square", 7, None)]
        return pts
    flines = [("Frontier without the tracking-error limit", f"{RF}$O$22:$O$31", f"{RF}$P$22:$P$31", REF_GREY, "dash", 15875),
              ("Efficient frontier, all mandate limits", f"{RF}$Q$22:$Q$31", f"{RF}$R$22:$R$31", NAVY, "solid", 25400)]
    S["chart2.xml"] = (scatter(flines, frontier_points(False), "Volatility (p.a.)", "Expected return (arithmetic, net, p.a.)",
                               title="Risk and return: recommended portfolio on the frontier"), (W, 3.6))
    S["chart3.xml"] = (scatter(flines, frontier_points(True), "Volatility (p.a.)", "Expected return (arithmetic, net, p.a.)"),
                       (W, 4.0))
    # Figure 3: growth of A$1
    S["chart5.xml"] = (line([(f"={RF}$W$73", f"{RF}$W$74:$W$154", ORANGE, "solid", 22225),
                             (f"={RF}$X$73", f"{RF}$X$74:$X$154", TEAL, "solid", 22225),
                             (f"={RF}$Y$73", f"{RF}$Y$74:$Y$154", NAVY, "solid", 25400),
                             (f"={RF}$Z$73", f"{RF}$Z$74:$Z$154", GOLD, "dash", 19050)],
                            f"{RF}$V$74:$V$154", val_fmt="0.0", val_title="Value of A$1 (net of fees)"), (W, H))
    # Figure 4: annualised returns by horizon, target as markers
    S["chart6.xml"] = (bar(port3(RF, "O", "P", "Q", (100, 103), 99), f"{RF}$N$100:$N$103",
                           lines=[(f"={RF}$R$99", f"{RF}$R$100:$R$103", GOLD, "solid", True)],
                           val_title="Annualised return (net of fees)"), (W, H))
    # Figure 5: stress tests
    S["chart7.xml"] = (bar(port3(RF, "O", "P", "Q", (126, 129), 125), f"{RF}$N$126:$N$129",
                           val_title="Portfolio return in scenario"), (W, H))
    # Figure 6: bootstrap distribution
    S["chart8.xml"] = (bar([(f"={RF}$O$151", f"{RF}$O$152:$O$166", ORANGE, False),
                            (f"={RF}$P$151", f"{RF}$P$152:$P$166", NAVY, False)], f"{RF}$Q$152:$Q$166",
                           val_fmt="0", val_title="Number of paths (of 1,000)",
                           cat_title="Cumulative 5-year return (lower edge of each 10-point bin)",
                           gap=30, overlap=-5, cat_kind="num", cat_fmt="0%"), (W, H))
    # Figure 7: return target out of reach
    S["chart9.xml"] = (bar([(f"={RF}$O$177", f"{RF}$O$178:$O$183", NAVY, True)], f"{RF}$N$178:$N$183",
                           lines=[(f"={RF}$P$177", f"{RF}$P$178:$P$183", GOLD, "dash", False)],
                           val_fmt="0%", lbl_fmt="0.00%", val_title="Expected return (compound, net, p.a.)", val_min=0), (W, H))
    # Figure 8: oil estimates
    S["chart10.xml"] = (bar(port3(RF, "O", "P", "Q", (204, 207), 203), f"{RF}$N$204:$N$207",
                            val_title="Portfolio return at US$150 oil"), (W, H))
    # Figure 9 (and Oil tab): oil betas by horizon
    beta = bar([(f"={RF}$O$229", f"{RF}$O$230:$O$232", LBLUE, True),
                (f"={RF}$P$229", f"{RF}$P$230:$P$232", NAVY, True)], f"{RF}$N$230:$N$232",
               val_fmt="0.00", lbl_fmt="0.00", val_title="Beta to log change in oil price")
    S["chart11.xml"] = (beta, (W, H))
    S["chart18.xml"] = (bar([(f"={RF}$O$229", f"{RF}$O$230:$O$232", LBLUE, True),
                             (f"={RF}$P$229", f"{RF}$P$230:$P$232", NAVY, True)], f"{RF}$N$230:$N$232",
                            val_fmt="0.00", lbl_fmt="0.00", val_title="Beta to log change in oil price",
                            title="Oil beta of equity returns by horizon"), (W, H))
    # Figure 10 (and Track record): rolling 5-year return vs target
    roll = [(f"={RF}$AC$255", f"{RF}$AC$256:$AC$316", ORANGE, "solid", 22225),
            (f"={RF}$AD$255", f"{RF}$AD$256:$AD$316", NAVY, "solid", 25400),
            (f"={RF}$AE$255", f"{RF}$AE$256:$AE$316", GOLD, "dash", 19050)]
    S["chart12.xml"] = (line(roll, f"{RF}$AB$256:$AB$316", val_title="Annualised 5-year return (net)"), (W, H))
    S["chart19.xml"] = (line(roll, f"{RF}$AB$256:$AB$316", val_title="Annualised 5-year return (net)",
                             title="Rolling 5-year return against the target"), (W, H))
    # Figure 11 (and Methods): construction methods scorecard
    meth = [(f"={RF}$O$281", f"{RF}$O$282:$O$288", NAVY, True),
            (f"={RF}$P$281", f"{RF}$P$282:$P$288", LBLUE, False),
            (f"={RF}$Q$281", f"{RF}$Q$282:$Q$288", VIOLET, False)]
    S["chart13.xml"] = (bar(meth, f"{RF}$N$282:$N$288", gap=50), (W, H))
    S["chart16.xml"] = (bar(meth, f"{RF}$N$282:$N$288", gap=50, title="Portfolio construction methods"), (W, H))
    # Figure 12: forecasts vs history
    S["chart14.xml"] = (bar([(f"={RF}$O$307", f"{RF}$O$308:$O$318", NAVY, True),
                             (f"={RF}$P$307", f"{RF}$P$308:$P$318", LBLUE, False)], f"{RF}$N$308:$N$318",
                            horizontal=True, val_title="Return (p.a.)", gap=40), (W, HT))
    # Figure 13: forecast stress test
    S["chart15.xml"] = (bar([(f"={RF}$O$333", f"{RF}$O$334:$O$344", NAVY, False),
                             (f"={RF}$P$333", f"{RF}$P$334:$P$344", LBLUE, False)], f"{RF}$N$334:$N$344",
                            horizontal=True, val_fmt="0.00%", val_title="Change in advantage over the current fund",
                            gap=40), (W, HT))
    # Figure 14: rolling active return
    S["chart20.xml"] = (line([(f"={RF}$AC$359", f"{RF}$AC$360:$AC$540", NAVY, "solid", 25400),
                              (f"={RF}$AD$359", f"{RF}$AD$360:$AD$540", REF_GREY, "dash", 12700)],
                             f"{RF}$AB$360:$AB$540", val_fmt="0.0%", val_title="Active return, annualised"), (W, H))
    # Figure 15: attribution
    S["chart21.xml"] = (bar([(f"={RF}$O$385", f"{RF}$O$386:$O$397", NAVY, False),
                             (f"={RF}$P$385", f"{RF}$P$386:$P$397", LBLUE, False)], f"{RF}$N$386:$N$397",
                            horizontal=True, val_fmt="0.0%", val_title="Contribution to active return (p.a.)",
                            gap=40), (W, HT))
    # Methods tab: weights by method
    mw = [(f"=Methods!${c}$5", f"Methods!${c}$6:${c}$16", col, False)
          for c, col in zip("DEFGH", [NAVY, TEAL, VIOLET, LBLUE, GOLD])]
    S["chart17.xml"] = (bar(mw, "Methods!$A$6:$A$16", horizontal=True, gap=40, title="Weights by method"),
                        (W, 5.0))
    return S


def apply(parts):
    """Replace every chart part and resize its drawing anchor."""
    import re
    S = specs()
    for name, (xml, _) in S.items():
        parts[f"xl/charts/{name}"] = xml.encode("utf-8")
    # anchors: find each drawing's rels to map rId -> chart file
    for dname in [n for n in parts if re.match(r"xl/drawings/drawing\d+\.xml$", n)]:
        rel = parts[dname.replace("drawings/", "drawings/_rels/") + ".rels"].decode()
        rid2chart = dict(re.findall(r'Id="(rId\d+)"[^>]*Target="../charts/([^"]+)"', rel))
        rid2chart.update({b: a for a, b in re.findall(r'Target="../charts/([^"]+)"[^>]*Id="(rId\d+)"', rel)})
        x = parts[dname].decode()

        def fix(m):
            a = m.group(0)
            rid = re.search(r'r:id="(rId\d+)"', a).group(1)
            ch = rid2chart.get(rid)
            if ch in S:
                w, h = S[ch][1]
                a = re.sub(r'<xdr:ext cx="\d+" cy="\d+"/>', f'<xdr:ext cx="{int(w * 914400)}" cy="{int(h * 914400)}"/>', a, 1)
            return a
        x = re.sub(r"<xdr:oneCellAnchor>.*?</xdr:oneCellAnchor>", fix, x, flags=re.S)
        parts[dname] = x.encode("utf-8")
