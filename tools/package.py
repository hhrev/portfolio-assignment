"""Post-process the openpyxl output.

1. Restore the original Excel chart and drawing parts byte-for-byte (openpyxl drops ~30% of the
   chart formatting on a round trip).
2. Add Figures 14 and 15 as clones of existing charts so they inherit the house chart style.
3. Write calculated values (from a LibreOffice full recalculation) into every formula cell so the
   file opens with numbers in any viewer; Excel still recalculates on load.
4. Rebuild every chart's data caches from those values.

Usage: python package.py built.xlsx calculated.xlsx original.xlsx figs.json out.xlsx
"""
import datetime as dt
import json
import re
import sys
import warnings
import zipfile

import openpyxl
from lxml import etree

warnings.filterwarnings("ignore")
NS_MAIN = "http://schemas.openxmlformats.org/spreadsheetml/2006/main"
NS_R = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
NS_PR = "http://schemas.openxmlformats.org/package/2006/relationships"
NS_C = "http://schemas.openxmlformats.org/drawingml/2006/chart"
CT_DRAWING = "application/vnd.openxmlformats-officedocument.drawing+xml"
CT_CHART = "application/vnd.openxmlformats-officedocument.drawingml.chart+xml"


def read_zip(path):
    with zipfile.ZipFile(path) as z:
        return {n: z.read(n) for n in z.namelist()}


def sheet_files(parts):
    wbx = etree.fromstring(parts["xl/workbook.xml"])
    rels = etree.fromstring(parts["xl/_rels/workbook.xml.rels"])
    tgt = {r.get("Id"): r.get("Target") for r in rels}
    out = {}
    for s in wbx.iter(f"{{{NS_MAIN}}}sheet"):
        t = tgt[s.get(f"{{{NS_R}}}id")].lstrip("/")
        out[s.get("name")] = t if t.startswith("xl/") else "xl/" + t
    return out


def rels_name(part):
    d, f = part.rsplit("/", 1)
    return f"{d}/_rels/{f}.rels"


def excel_serial(v):
    if isinstance(v, dt.datetime):
        delta = v - dt.datetime(1899, 12, 30)
        return delta.days + delta.seconds / 86400
    if isinstance(v, dt.date):
        return (v - dt.date(1899, 12, 30)).days
    return v


def fmt_num(v):
    if isinstance(v, bool):
        return "1" if v else "0"
    if isinstance(v, int):
        return str(v)
    return repr(float(v))


# ---------------------------------------------------------------------------------------------
def restore_charts(parts, orig, figs):
    built_sheets = sheet_files(parts)
    orig_sheets = sheet_files(orig)
    for n in [n for n in parts if n.startswith(("xl/drawings/", "xl/charts/"))]:
        del parts[n]
    sheet_drawing = {}
    for name, sf in orig_sheets.items():
        rn = rels_name(sf)
        if rn in orig:
            for r in etree.fromstring(orig[rn]):
                if r.get("Type").endswith("/drawing"):
                    sheet_drawing[name] = "xl/drawings/" + r.get("Target").rsplit("/", 1)[1]
    for n, b in orig.items():
        if n.startswith(("xl/drawings/", "xl/charts/")):
            parts[n] = b
    for name, dpart in sheet_drawing.items():
        rn = rels_name(built_sheets[name])
        rel = etree.fromstring(parts[rn])
        for r in rel:
            if r.get("Type").endswith("/drawing"):
                r.set("Target", "/" + dpart)
        parts[rn] = etree.tostring(rel, xml_declaration=True, encoding="UTF-8", standalone=True)

    # ---- new charts: Figure 14 (clone of Figure 10 line chart), Figure 15 (clone of Figure 13 bars)
    fig_drawing = sheet_drawing["Report figures"]
    if figs is None:
        _content_types(parts)
        return
    new = [("chart20.xml", "chart12.xml", figs["fig14"], 358,
            "Rolling 5-year active return, net of fees (current less benchmark)",
            {"Annualised 5-year return": "Active return, annualised"}),
           ("chart21.xml", "chart15.xml", figs["fig15"], 384,
            "Attribution of active return, net of fees",
            {})]
    dxml = parts[fig_drawing].decode("utf-8")
    drels = etree.fromstring(parts[rels_name(fig_drawing)])
    anchors = re.findall(r"<xdr:oneCellAnchor>.*?</xdr:oneCellAnchor>", dxml, re.S)
    template = anchors[-1]
    ids = [int(x) for x in re.findall(r'<xdr:cNvPr id="(\d+)"', dxml)]
    next_id = max(ids) + 1
    rid_n = max(int(r.get("Id")[3:]) for r in drels) + 1
    for fname, src, spec, title_row, title, renames in new:
        x = parts[f"xl/charts/{src}"].decode("utf-8")
        q = "'Report figures'!"
        sers = re.findall(r"<c:ser>.*?</c:ser>", x, re.S)
        if src == "chart12.xml":
            x = x.replace(sers[1], "")            # keep series 1 (data) and series 3 (reference line)
            sers = [sers[0], sers[2]]
        for k, s in enumerate(sers[:2]):
            s2 = s
            fs = re.findall(r"<c:f>([^<]*)</c:f>", s)
            s2 = s2.replace(f"<c:f>{fs[0]}</c:f>", f"<c:f>{q}{spec['s' + str(k + 1)]}</c:f>", 1)
            s2 = s2.replace(f"<c:f>{fs[1]}</c:f>", f"<c:f>{q}{spec['cat']}</c:f>", 1)
            s2 = s2.replace(f"<c:f>{fs[2]}</c:f>", f"<c:f>{q}{spec['v' + str(k + 1)]}</c:f>", 1)
            s2 = re.sub(r'<c:idx val="\d+"/><c:order val="\d+"/>', f'<c:idx val="{k}"/><c:order val="{k}"/>', s2, 1)
            x = x.replace(s, s2)
        old_title = re.findall(r"<a:t>([^<]*)</a:t>", re.search(r"<c:title>.*?</c:title>", x, re.S).group(0))
        x = x.replace(f"<a:t>{old_title[0]}</a:t>", f"<a:t>{title}</a:t>", 1)
        for a, b in renames.items():
            x = x.replace(f"<a:t>{a}</a:t>", f"<a:t>{b}</a:t>")
        parts[f"xl/charts/{fname}"] = x.encode("utf-8")
        rid = f"rId{rid_n}"
        rid_n += 1
        a = re.sub(r"<xdr:row>\d+</xdr:row>", f"<xdr:row>{title_row + 1}</xdr:row>", template, 1)
        a = re.sub(r'<xdr:cNvPr id="\d+" name="[^"]*"', f'<xdr:cNvPr id="{next_id}" name="Chart {next_id}"', a, 1)
        a = re.sub(r'(a16:creationId[^>]*id="\{)[0-9A-F]{8}', lambda m: m.group(1) + f"{next_id:08X}", a, 1)
        a = re.sub(r'r:id="rId\d+"', f'r:id="{rid}"', a, 1)
        next_id += 1
        dxml = dxml.replace("</xdr:wsDr>", a + "</xdr:wsDr>")
        el = etree.SubElement(drels, f"{{{NS_PR}}}Relationship")
        el.set("Id", rid)
        el.set("Type", "http://schemas.openxmlformats.org/officeDocument/2006/relationships/chart")
        el.set("Target", f"../charts/{fname}")
    parts[fig_drawing] = dxml.encode("utf-8")
    parts[rels_name(fig_drawing)] = etree.tostring(drels, xml_declaration=True, encoding="UTF-8", standalone=True)

    _content_types(parts)


def _content_types(parts):
    ct = etree.fromstring(parts["[Content_Types].xml"])
    for o in list(ct):
        pn = o.get("PartName") or ""
        if pn.startswith(("/xl/drawings/", "/xl/charts/")):
            ct.remove(o)
    ns = ct.nsmap[None]
    for n in sorted(parts):
        if re.match(r"xl/drawings/drawing\d+\.xml$", n) or re.match(r"xl/charts/chart\d+\.xml$", n):
            o = etree.SubElement(ct, f"{{{ns}}}Override")
            o.set("PartName", "/" + n)
            o.set("ContentType", CT_DRAWING if "drawings" in n else CT_CHART)
    parts["[Content_Types].xml"] = etree.tostring(ct, xml_declaration=True, encoding="UTF-8", standalone=True)


# ---------------------------------------------------------------------------------------------
def inject_values(parts, calc):
    files = sheet_files(parts)
    n_set = 0
    for name, sf in files.items():
        ws = calc[name]
        root = etree.fromstring(parts[sf])
        for c in root.iter(f"{{{NS_MAIN}}}c"):
            f = c.find(f"{{{NS_MAIN}}}f")
            if f is None:
                continue
            v = c.find(f"{{{NS_MAIN}}}v")
            if v is None:
                v = etree.SubElement(c, f"{{{NS_MAIN}}}v")
            val = ws[c.get("r")].value
            val = excel_serial(val)
            if val is None:
                c.remove(v)
                c.attrib.pop("t", None)
                continue
            if isinstance(val, bool):
                c.set("t", "b")
                v.text = fmt_num(val)
            elif isinstance(val, (int, float)):
                c.attrib.pop("t", None)
                v.text = fmt_num(val)
            elif isinstance(val, str) and re.match(r"^#(N/A|VALUE!|REF!|DIV/0!|NUM!|NAME\?|NULL!)$", val):
                c.set("t", "e")
                v.text = val
            else:
                c.set("t", "str")
                v.text = str(val)
            n_set += 1
        parts[sf] = etree.tostring(root, xml_declaration=True, encoding="UTF-8", standalone=True)
    return n_set


def cell_values(calc, ref):
    sheet, rng = ref.rsplit("!", 1)
    sheet = sheet.strip("'").replace("''", "'")
    ws = calc[sheet]
    rng = rng.replace("$", "")
    if ":" in rng:
        cells = ws[rng]
        flat = [c for row in cells for c in row]
    else:
        flat = [ws[rng]]
    return [excel_serial(c.value) for c in flat]


def refresh_caches(parts, calc):
    for n in [n for n in parts if re.match(r"xl/charts/chart\d+\.xml$", n)]:
        root = etree.fromstring(parts[n])
        for kind in ("numRef", "strRef"):
            for ref in root.iter(f"{{{NS_C}}}{kind}"):
                f = ref.find(f"{{{NS_C}}}f").text
                vals = cell_values(calc, f)
                cache_tag = "numCache" if kind == "numRef" else "strCache"
                old = ref.find(f"{{{NS_C}}}{cache_tag}")
                fmt = None
                if old is not None:
                    fc = old.find(f"{{{NS_C}}}formatCode")
                    fmt = fc.text if fc is not None else None
                    ref.remove(old)
                cache = etree.SubElement(ref, f"{{{NS_C}}}{cache_tag}")
                if kind == "numRef":
                    fc = etree.SubElement(cache, f"{{{NS_C}}}formatCode")
                    fc.text = fmt or "General"
                etree.SubElement(cache, f"{{{NS_C}}}ptCount").set("val", str(len(vals)))
                for i, v in enumerate(vals):
                    if v is None or (kind == "numRef" and not isinstance(v, (int, float))):
                        continue
                    pt = etree.SubElement(cache, f"{{{NS_C}}}pt")
                    pt.set("idx", str(i))
                    etree.SubElement(pt, f"{{{NS_C}}}v").text = fmt_num(v) if kind == "numRef" else str(v)
        parts[n] = etree.tostring(root, xml_declaration=True, encoding="UTF-8", standalone=True)


def main(built, calc_path, orig_path, figs_path, out):
    parts = read_zip(built)
    orig = read_zip(orig_path)
    figs = None if figs_path == "-" else json.load(open(figs_path))
    calc = openpyxl.load_workbook(calc_path, data_only=True)
    restore_charts(parts, orig, figs)
    n = inject_values(parts, calc)
    refresh_caches(parts, calc)
    # keep the original document properties (author, application)
    if "docProps/core.xml" in orig:
        parts["docProps/core.xml"] = orig["docProps/core.xml"]
    with zipfile.ZipFile(out, "w", zipfile.ZIP_DEFLATED) as z:
        order = ["[Content_Types].xml"] + [p for p in parts if p != "[Content_Types].xml"]
        for p in order:
            z.writestr(p, parts[p])
    print("values written:", n)


if __name__ == "__main__":
    main(*sys.argv[1:6])
