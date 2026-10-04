#!/bin/sh
# Full pipeline: v3.3 -> v3.4 (build, recalculate, restore charts, write values)
# Usage: tools/make.sh source/FINM3008_EPPIB_Model_v3.3.xlsx FINM3008_EPPIB_Model_v3.4.xlsx
# Needs: python3 with openpyxl and lxml; LibreOffice Calc (soffice) for recalculation.
set -e
SRC=$(realpath "$1"); OUT=$(realpath -m "$2"); export T=${T:-/tmp/w}
mkdir -p "$T"
cd "$(dirname "$0")"
python3 build_v34.py "$SRC" "$T/built.xlsx"
./recalc.sh "$T/built.xlsx" >/dev/null
cp "$T/calc/built.xlsx" "$T/calc_built.xlsx"
python3 package.py "$T/built.xlsx" "$T/calc_built.xlsx" "$SRC" "$T/built.xlsx.figs.json" "$OUT"
python3 errscan.py "$OUT"
