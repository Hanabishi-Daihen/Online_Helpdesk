// Syntax check: the app module, the sign-in script inside index.html, and all .gs files together
const fs = require('fs'), path = require('path'), { execFileSync } = require('child_process');
const ROOT = path.resolve(__dirname, '../..'), OUT = path.join(__dirname, '../out');
fs.mkdirSync(OUT, { recursive: true });
const s = fs.readFileSync(path.join(ROOT, 'index.html'), 'utf8');
const mod = [...s.matchAll(/<script type="module"[^>]*>([\s\S]*?)<\/script>/g)].map(x => x[1]).sort((a, b) => b.length - a.length)[0];
fs.writeFileSync(path.join(OUT, 'm.mjs'), mod);
const i = s.indexOf('<!-- zq-auth:start'), j = s.indexOf('<!-- zq-auth:end');
if (i >= 0) { const b = s.slice(i, j); fs.writeFileSync(path.join(OUT, 'auth.js'), b.slice(b.indexOf('<script>') + 8, b.lastIndexOf('</script>'))); }
fs.writeFileSync(path.join(OUT, 'all.js'), fs.readdirSync(ROOT).filter(f => f.endsWith('.gs')).map(f => fs.readFileSync(path.join(ROOT, f), 'utf8')).join('\n'));
let bad = 0;
for (const f of ['m.mjs', 'auth.js', 'all.js']) {
  const p = path.join(OUT, f); if (!fs.existsSync(p)) continue;
  try { execFileSync(process.execPath, ['--check', p], { stdio: 'pipe' }); console.log('  ' + f + ' OK'); }
  catch (e) { bad++; console.log('  ' + f + ' FAIL\n' + String(e.stderr || e).split('\n').slice(0, 6).join('\n')); }
}
process.exit(bad ? 1 : 0);
