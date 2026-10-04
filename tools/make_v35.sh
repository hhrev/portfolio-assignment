#!/bin/sh
# v3.4 -> v3.5: theory-based forecasts, then re-optimise every forecast-dependent Solver set.
# Usage: tools/make_v35.sh FINM3008_EPPIB_Model_v3.4.xlsx FINM3008_EPPIB_Model_v3.5.xlsx
set -e
SRC=$(realpath "$1"); OUT=$(realpath -m "$2"); export T=${T:-/tmp/w5}
mkdir -p "$T"
cd "$(dirname "$0")"
python3 build_v35.py "$SRC" "$T/b1.xlsx"
./recalc.sh "$T/b1.xlsx" >/dev/null && cp "$T/calc/b1.xlsx" "$T/c1.xlsx"
python3 reopt.py "$T/b1.xlsx" "$T/c1.xlsx" "$T/b2.xlsx"
./recalc.sh "$T/b2.xlsx" >/dev/null && cp "$T/calc/b2.xlsx" "$T/c2.xlsx"
python3 package.py "$T/b2.xlsx" "$T/c2.xlsx" "$SRC" - "$OUT"
python3 errscan.py "$OUT"
