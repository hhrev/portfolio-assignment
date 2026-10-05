#!/bin/sh
# Full pipeline v3.3 -> v3.7: base build, theory forecasts, v3.7 decisions, re-optimisation, charts.
# Usage: tools/make_v37.sh source/FINM3008_EPPIB_Model_v3.3.xlsx FINM3008_EPPIB_Model_v3.7.xlsx
set -e
SRC=$(realpath "$1"); OUT=$(realpath -m "$2"); export T=${T:-/tmp/w7}
mkdir -p "$T"
cd "$(dirname "$0")"
T="$T" ./make.sh "$SRC" "$T/v34.xlsx" >/dev/null
python3 build_v35.py "$T/v34.xlsx" "$T/b1.xlsx" 2>/dev/null
python3 build_v36.py "$T/b1.xlsx" "$T/b1a.xlsx" ../docs/forecast_rationale_notes.md
python3 build_v37.py "$T/b1a.xlsx" "$T/b1b.xlsx" ../docs/oil_mitigation_notes.md
./recalc.sh "$T/b1b.xlsx" >/dev/null && cp "$T/calc/b1b.xlsx" "$T/c1.xlsx"
python3 reopt.py "$T/b1b.xlsx" "$T/c1.xlsx" "$T/b2.xlsx" v37
./recalc.sh "$T/b2.xlsx" >/dev/null && cp "$T/calc/b2.xlsx" "$T/c2.xlsx"
python3 package.py "$T/b2.xlsx" "$T/c2.xlsx" "$T/v34.xlsx" - "$OUT"
python3 errscan.py "$OUT"
