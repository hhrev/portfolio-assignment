import sys, warnings, openpyxl
warnings.filterwarnings("ignore")
wb = openpyxl.load_workbook(sys.argv[1], data_only=True)
bad = []
for ws in wb.worksheets:
    for row in ws.iter_rows():
        for c in row:
            if isinstance(c.value, str) and c.value.startswith("#") and c.value[1:4].isupper():
                bad.append((ws.title, c.coordinate, c.value))
print("error cells:", len(bad)); print(bad[:40])
