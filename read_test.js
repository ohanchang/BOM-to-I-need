const XLSX = require('../BOM Fix/static/js/xlsx.full.min.js');
const fs = require('fs');
const files = ['good.xls', 'good 2.xls', 'good 3.xls'];
files.forEach(f => {
  console.log(`=== FILE: ${f} ===`);
  const buf = fs.readFileSync('E:/Antigravity/BOM to I need/' + f);
  const wb = XLSX.read(buf, { type: 'buffer' });
  const ws = wb.Sheets[wb.SheetNames[0]];
  const raw = XLSX.utils.sheet_to_json(ws, {header:1, defval:''});
  for (let r = 0; r < Math.min(13, raw.length); r++) {
    console.log(`Row ${r}:`, JSON.stringify(raw[r]));
  }
});
