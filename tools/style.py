"""Style helpers that reproduce the boxed section layout used throughout the EPPIB model.

Every section in the workbook follows the same pattern:
  * title row: dark fill, white bold Arial 10, medium dark border on the outside
  * header row: light fill of the same hue, dark bold text, wrapped, thin dark bottom rule
  * body: thin light-grey (D0D0D0) inner grid, medium dark border around the outside
"""
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side

# (dark, light) pairs already used in the workbook
NAVY = ("FF203864", "FFD9E1F2")
BLUE = ("FF1F4E79", "FFDDEBF7")
GOLD = ("FF806000", "FFFFF2CC")
PURPLE = ("FF5B2C83", "FFE4DFEC")
GREEN = ("FF375623", "FFE2EFDA")
GREY = ("FF595959", "FFEDEDED")
OLIVE = ("FF4F6228", "FFEBF1DE")
TEAL = ("FF1B6B73", "FFDAEEF3")
RED = ("FFA50021", "FFFBE2D5")
ORANGE = ("FFC55A11", "FFFCE4D6")
BROWN = ("FF843C0C", "FFF4E1D2")

GRID = "FFD0D0D0"
INPUT_BLUE = "FF0000FF"
FONT = "Arial"


def _side(style, color):
    return Side(style=style, color=color)


def font(color=None, bold=False, size=10, italic=False):
    return Font(name=FONT, size=size, bold=bold, italic=italic, color=color)


def box(ws, r0, c0, title, headers, rows, palette, fmts=None, header_height=None,
        inputs=None, bold_rows=(), bold_first_col=False, notes_col=None):
    """Write a boxed section.

    rows: list of lists (values or formulas). None leaves a cell blank but styled.
    fmts: list of number formats per column (applied to body cells), or a dict
          {(row_idx, col_idx): fmt} overrides via a callable.
    inputs: set of (row_idx, col_idx) body cells to colour blue (hard-coded inputs);
            use (row_idx, None) for a whole row or (None, col_idx) for a column.
    Returns the last row written.
    """
    dark, light = palette
    ncols = max(len(headers) if headers else 0, max((len(r) for r in rows), default=0), 1)
    inputs = inputs or set()
    thin = _side("thin", GRID)
    med = _side("medium", dark)
    last_body = r0 + (1 if headers else 0) + len(rows)

    def border(row, ci, kind):
        left = med if ci == 0 else thin
        right = med if ci == ncols - 1 else thin
        top = med if kind == "title" else thin
        if kind == "title":
            bottom = thin
        elif kind == "header":
            bottom = _side("thin", dark)
        else:
            bottom = med if row == last_body else thin
        return Border(left=left, right=right, top=top, bottom=bottom)

    # title
    for ci in range(ncols):
        c = ws.cell(r0, c0 + ci)
        c.value = title if ci == 0 else None
        c.fill = PatternFill("solid", fgColor=dark)
        c.font = font("FFFFFFFF", bold=True)
        c.border = border(r0, ci, "title")
    r = r0 + 1
    if headers:
        for ci in range(ncols):
            c = ws.cell(r, c0 + ci)
            c.value = headers[ci] if ci < len(headers) else None
            c.fill = PatternFill("solid", fgColor=light)
            c.font = font(dark, bold=True)
            c.alignment = Alignment(horizontal="left" if ci == 0 else "center",
                                    vertical="center", wrap_text=True)
            c.border = border(r, ci, "header")
        if header_height:
            ws.row_dimensions[r].height = header_height
        r += 1
    for ri, vals in enumerate(rows):
        for ci in range(ncols):
            v = vals[ci] if ci < len(vals) else None
            c = ws.cell(r, c0 + ci)
            c.value = v
            is_input = (ri, ci) in inputs or (ri, None) in inputs or (None, ci) in inputs
            bold = ri in bold_rows or (bold_first_col and ci == 0)
            size = 9 if notes_col is not None and ci == notes_col else 10
            c.font = font(INPUT_BLUE if is_input else None, bold=bold, size=size)
            if fmts:
                f = fmts(ri, ci) if callable(fmts) else (fmts[ci] if ci < len(fmts) else None)
                if f:
                    c.number_format = f
            if ci > 0 and isinstance(v, str) and not v.startswith("=") and notes_col != ci:
                c.alignment = Alignment(horizontal="center")
            c.border = border(r, ci, "body")
        r += 1
    return r - 1


def sheet_title(ws, text):
    c = ws["A1"]
    c.value = text
    c.font = font("FF1F3864", bold=True, size=15)
    ws.row_dimensions[1].height = 19


def status_cf(ws, rng):
    """Same status colouring as the Summary and Checks tabs, plus REVIEW (amber)."""
    from openpyxl.formatting.rule import CellIsRule
    green = PatternFill("solid", fgColor="FFC6EFCE")
    red = PatternFill("solid", fgColor="FFFFC7CE")
    amber = PatternFill("solid", fgColor="FFFFEB9C")
    for row in ws[rng] if ":" in rng else [[ws[rng]]]:
        for c in row:
            c.alignment = Alignment(horizontal="center")
    for word, fill in (("MET", green), ("PASS", green), ("INFO", green), ("NOT MET", red),
                       ("FAIL", red), ("MAX FEASIBLE", amber), ("REVIEW", amber)):
        ws.conditional_formatting.add(rng, CellIsRule(operator="equal", formula=[f'"{word}"'],
                                                      fill=fill))
