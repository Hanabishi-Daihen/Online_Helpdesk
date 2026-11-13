#!/usr/bin/env bash
# Run every test: server (.gs in a Node sandbox) then browser (index.html with a fake server, in Chromium).
cd "$(dirname "$0")" || exit 1
[ -x /opt/pw-browsers/chromium ] && export CHROMIUM_PATH=/opt/pw-browsers/chromium
fail=0
echo "== syntax"
cat ../*.gs > out_all.js 2>/dev/null; mkdir -p out; mv out_all.js out/all.js; node --check out/all.js && echo "  .gs OK" || fail=1
echo "== server"
for f in server/*.test.js; do r=$(node "$f" 2>&1 | tail -1); echo "  $(basename "$f"): $r"; echo "$r" | grep -q " 0 failed" || fail=1; done
echo "== browser"
python3 browser/build.py order.html order.html && python3 browser/build.py doc.html doc.html || fail=1
for f in browser/*.js; do if node "$f" > "out/$(basename "$f" .js).log" 2>&1; then echo "  $(basename "$f"): OK"; else echo "  $(basename "$f"): FAIL (see tests/out/$(basename "$f" .js).log)"; fail=1; fi; done
[ $fail = 0 ] && echo "ALL PASSED" || echo "SOME FAILED"
exit $fail
