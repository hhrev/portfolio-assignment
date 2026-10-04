#!/bin/sh
# Recalculate an xlsx with LibreOffice (forced full recalculation on load) into $T/calc/
T=${T:-/tmp/w}
export HOME=$T/lohome
CFG=$HOME/.config/libreoffice/4/user
mkdir -p "$CFG"
cat > "$CFG/registrymodifications.xcu" <<'XCU'
<?xml version="1.0" encoding="UTF-8"?>
<oor:items xmlns:oor="http://openoffice.org/2001/registry" xmlns:xs="http://www.w3.org/2001/XMLSchema" xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance">
<item oor:path="/org.openoffice.Office.Calc/Formula/Load"><prop oor:name="OOXMLRecalcMode" oor:op="fuse"><value>0</value></prop></item>
</oor:items>
XCU
rm -rf "$T/calc" && timeout 900 soffice --headless --norestore --convert-to xlsx --outdir "$T/calc" "$1" >/dev/null 2>&1
ls "$T/calc"
