const XLSX = require('../BOM Fix/static/js/xlsx.full.min.js');
const fs = require('fs');
const path = require('path');

const STANDARD_HEADER_ROW = 11;
const STANDARD_COL_COUNT = 13;
const STANDARD_HEADERS = [
  "No", "階層", "料號", "規格項目-名稱說明", "UOM",
  "使用數量", "插件位置", "ERP Status", "LT",
  "是否為安規料", "單價", "ASL狀態", "最新報價日(日/月/年)"
];

function processBOMWorkbook(wb, filename, defaultRate) {
  const sheetName = wb.SheetNames[0];
  const ws = wb.Sheets[sheetName];
  
  const hasHighLevelSheet = wb.SheetNames.includes('大分類統計');
  const hasCategorySheet = wb.SheetNames.includes('物料拆解分析');
  const raw = XLSX.utils.sheet_to_json(ws, { header: 1, defval: '' });
  
  let headerRow = -1;
  for (let r = 0; r < Math.min(25, raw.length); r++) {
    const row = raw[r];
    if (row && String(row[0]).trim() === 'No' && String(row[1]).trim() === '階層' && String(row[2]).trim() === '料號') {
      headerRow = r;
      break;
    }
  }

  const isL12Product = (headerRow !== -1 && raw[headerRow] && String(raw[headerRow][11]).trim() === '乘積');
  
  if (hasHighLevelSheet || hasCategorySheet || isL12Product) {
    return { filename, status: 'processed', fixedWb: wb, issues: ['此檔案已是整理過後的 BOM 規格'] };
  }

  if (headerRow === -1) {
    return { filename, status: 'error', fixedWb: null, issues: ['找不到標準標題列 (No/階層/料號)'] };
  }

  const issues = [];
  const fixLog = [];
  let needsFix = false;

  if (headerRow !== STANDARD_HEADER_ROW) {
    issues.push(`標題列位置偏移：Row ${headerRow} → 應為 Row ${STANDARD_HEADER_ROW}`);
    needsFix = true;
  }

  const maxCols = raw.reduce((m, r) => Math.max(m, r.length), 0);
  if (maxCols > STANDARD_COL_COUNT) {
    issues.push(`欄位數異常：${maxCols} 欄 → 標準為 ${STANDARD_COL_COUNT} 欄`);
    needsFix = true;
  }

  const dataStart = headerRow + 1;
  let shiftedCount = 0;
  const shiftedRows = [];

  for (let r = dataStart; r < raw.length; r++) {
    const row = raw[r];
    if (!row || row.length === 0) continue;
    const noVal = String(row[0]).trim();
    if (!noVal) continue;

    const col9 = String(row[9] ?? '').trim();
    const col10 = String(row[10] ?? '').trim();
    const col11 = String(row[11] ?? '').trim();

    if (col9 === '' && col10 === '' && col11 !== '' && row.length > 13) {
      shiftedCount++;
      shiftedRows.push(r);
    }
  }

  if (shiftedCount > 0) {
    issues.push(`發現 ${shiftedCount} 筆子料行欄位偏移 (已自動修正搬移)`);
    needsFix = true;
  }

  const fixed = raw.map(r => [...r]);

  for (const r of shiftedRows) {
    const row = fixed[r];
    row[9] = row[11];
    row[10] = row[12];
    row[11] = row[13] ?? '';
    row[12] = row[14] ?? '';
  }

  for (let r = 0; r < fixed.length; r++) {
    fixed[r] = fixed[r].slice(0, STANDARD_COL_COUNT);
  }

  const findMetaRow = (rawRows, hRow, key) => {
    const normKey = key.replace(/[:：\s]/g, '').toLowerCase();
    for (let r = 0; r < hRow; r++) {
      const row = rawRows[r];
      if (!row || row.length === 0) continue;
      const firstCell = String(row[0] ?? '').replace(/[:：\s]/g, '').toLowerCase();
      if (firstCell === normKey || firstCell.startsWith(normKey)) {
        return row;
      }
    }
    return null;
  };

  const dataRows = fixed.slice(dataStart);
  const rebuilt = [];

  for (let i = 0; i < 3; i++) {
    rebuilt.push(new Array(STANDARD_COL_COUNT).fill(''));
  }

  const META_KEYS = [
    '成品/半成品料號：',
    'BOM 版本：',
    'Change NO：',
    'BOM Cost：',
    '構成總筆數：',
    'UOM：',
    '品名規格：'
  ];

  for (const key of META_KEYS) {
    const foundRow = findMetaRow(fixed, headerRow, key);
    if (foundRow) {
      const sliced = foundRow.slice(0, STANDARD_COL_COUNT);
      while (sliced.length < STANDARD_COL_COUNT) sliced.push('');
      if (sliced[1] !== '' && sliced[3] === '') {
        sliced[3] = sliced[1];
        sliced[1] = '';
      }
      rebuilt.push(sliced);
    } else {
      const defaultRow = new Array(STANDARD_COL_COUNT).fill('');
      defaultRow[0] = key;
      rebuilt.push(defaultRow);
    }
  }

  rebuilt.push(new Array(STANDARD_COL_COUNT).fill(''));
  rebuilt.push([...STANDARD_HEADERS]);

  for (const dr of dataRows) {
    const hasAnyData = dr.some(c => String(c ?? '').trim() !== '');
    if (hasAnyData) {
      const sliced = dr.slice(0, STANDARD_COL_COUNT);
      while (sliced.length < STANDARD_COL_COUNT) sliced.push('');
      rebuilt.push(sliced);
    }
  }

  let rate = defaultRate;
  if (wb.SheetNames.includes('Ctrl Q')) {
    const ctrlQSheet = wb.Sheets['Ctrl Q'];
    const ctrlQRaw = XLSX.utils.sheet_to_json(ctrlQSheet, { header: 1, defval: '' });
    if (ctrlQRaw && ctrlQRaw[5] && ctrlQRaw[5][0] !== '') {
      const parsedRate = parseFloat(ctrlQRaw[5][0]);
      if (!isNaN(parsedRate) && parsedRate > 0) {
        rate = parsedRate;
      }
    }
  }

  const finalBOMHeaders = [
    "No", "階層", "料號", "規格項目-名稱說明", "UOM",
    "使用數量", "插件位置", "ERP Status", "LT",
    "是否為安規料", "單價", "乘積", "ASL狀態", "最新報價日(日/月/年)"
  ];

  const processedBOMRows = [];
  for (let r = 0; r < 11; r++) {
    const row = [...rebuilt[r]];
    row.splice(11, 0, '');
    processedBOMRows.push(row);
  }

  processedBOMRows[1][4] = '匯率';
  processedBOMRows[2][4] = rate;
  processedBOMRows[3][4] = { f: 'L10', v: 0.0 };
  processedBOMRows[1][6] = new Date().toLocaleDateString('zh-TW');
  processedBOMRows[2][6] = '美金';
  processedBOMRows[3][6] = { f: 'E4/E3', v: 0.0 };
  processedBOMRows[8][11] = 'TOTAL:';
  processedBOMRows[9][11] = { f: 'SUBTOTAL(9,L13:L1500)', v: 0.0 };
  processedBOMRows.push(finalBOMHeaders);

  const lastRowIndex = rebuilt.length;
  const formulaLastRow = Math.max(1500, lastRowIndex);

  const highLevelSum = {
    '1xxx 主動': 0.0, '2xxx 被動': 0.0, '3xxx 機電與保護': 0.0, '4xxx 連接器線材': 0.0,
    '5xxx 機構五金': 0.0, '6xxx 輔料與包材': 0.0, '7xxx 磁芯': 0.0, '8xxx 磁性元件': 0.0
  };
  const categoryBreakdown = {
    'DIODE': 0.0, 'IC': 0.0, 'MOS/ TR.': 0.0, 'CAP.': 0.0, 'DIP Res./Cap.': 0.0,
    'SMD Res./Cap.': 0.0, 'Magnetic': 0.0, 'PCB': 0.0, '機構件/ Other': 0.0,
    'FAN': 0.0, 'FAN GUARD': 0.0, 'CASE': 0.0, 'O/P WIRE': 0.0
  };

  let totalBOMCost = 0.0;

  for (let r = 12; r < rebuilt.length; r++) {
    const rawRow = rebuilt[r];
    const level = String(rawRow[1]).trim();
    const partNo = String(rawRow[2]).trim().toUpperCase();
    const desc = String(rawRow[3]).trim().toUpperCase();
    const qty = parseFloat(rawRow[5]) || 0;
    const price = parseFloat(rawRow[10]) || 0;
    const product = qty * price;

    const row = [...rawRow];
    const excelRow = r + 1;
    row.splice(11, 0, { f: `IFERROR(F${excelRow}*K${excelRow},0)`, v: product });

    if (level !== '') {
      totalBOMCost += product;
      const firstChar = partNo.charAt(0);
      const hlKey = firstChar === '1' ? '1xxx 主動' :
                    firstChar === '2' ? '2xxx 被動' :
                    firstChar === '3' ? '3xxx 機電與保護' :
                    firstChar === '4' ? '4xxx 連接器線材' :
                    firstChar === '5' ? '5xxx 機構五金' :
                    firstChar === '6' ? '6xxx 輔料與包材' :
                    firstChar === '7' ? '7xxx 磁芯' :
                    firstChar === '8' ? '8xxx 磁性元件' : null;
      if (hlKey) highLevelSum[hlKey] += product;

      let cat = '機構件/ Other';
      if (partNo.startsWith('1D')) {
        cat = 'DIODE';
      } else if (partNo.startsWith('1I') || partNo.startsWith('1S') || partNo.startsWith('1OT')) {
        cat = 'IC';
      } else if (partNo.startsWith('1T') || partNo.startsWith('1M')) {
        cat = 'MOS/ TR.';
      } else if ((partNo.startsWith('2EC') || partNo.startsWith('2EG') || partNo.startsWith('2EL')) && !desc.includes('SMD')) {
        cat = 'CAP.';
      } else if (['2CX', '2CY', '2EV', '2FN', '2FS', '2OP', '2VR', '3RD', '3UC'].some(p => partNo.startsWith(p)) && !desc.includes('SMD')) {
        cat = 'DIP Res./Cap.';
      } else if (partNo.startsWith('2') && desc.includes('SMD')) {
        cat = 'SMD Res./Cap.';
      } else if ((partNo.startsWith('7') || partNo.startsWith('8')) && !partNo.startsWith('8B')) {
        cat = 'Magnetic';
      } else if (partNo.startsWith('3B') || partNo.startsWith('8B')) {
        cat = 'PCB';
      } else if (desc.includes('FAN') && !desc.includes('FAN GUARD') && !partNo.startsWith('1')) {
        cat = 'FAN';
      } else if (desc.includes('FAN GUARD')) {
        cat = 'FAN GUARD';
      } else if (desc.includes('CASE') || desc.includes('CHASSIS')) {
        cat = 'CASE';
      } else if ((partNo.startsWith('4W') || desc.includes('WIRE')) && !desc.includes('GUARD')) {
        cat = 'O/P WIRE';
      }
      categoryBreakdown[cat] += product;
    }

    processedBOMRows.push(row);
  }

  processedBOMRows[3][4].v = totalBOMCost;
  processedBOMRows[3][6].v = totalBOMCost / rate;
  processedBOMRows[9][11].v = totalBOMCost;

  const sheet1 = XLSX.utils.aoa_to_sheet(processedBOMRows);
  sheet1['E4'].t = 'n';
  sheet1['G4'].t = 'n';
  sheet1['L10'].t = 'n';
  for (let r = 12; r < rebuilt.length; r++) {
    const key = 'L' + (r + 1);
    if (sheet1[key]) sheet1[key].t = 'n';
  }
  sheet1['!autofilter'] = { ref: `A12:N${rebuilt.length}` };
  sheet1['!cols'] = [
    { wch: 5 }, { wch: 5 }, { wch: 16 }, { wch: 78 }, { wch: 6 },
    { wch: 10 }, { wch: 15 }, { wch: 12 }, { wch: 8 }, { wch: 14 },
    { wch: 12 }, { wch: 12 }, { wch: 12 }, { wch: 18 }
  ];

  // Sheet 2: 大分類統計
  const sheet2Rows = [['大分類', '金額 (NTD)', '百分比']];
  const hlKeys = Object.keys(highLevelSum);
  hlKeys.forEach((key, idx) => {
    const excelRow = idx + 2;
    const digit = key.charAt(0);
    sheet2Rows.push([
      key,
      { f: `SUMIFS(BOM!$L$13:$L$${formulaLastRow},BOM!$C$13:$C$${formulaLastRow},"${digit}*",BOM!$B$13:$B$${formulaLastRow},"<>\")`, v: highLevelSum[key] },
      { f: `IFERROR(B${excelRow}/$B$10,0)`, v: totalBOMCost > 0 ? highLevelSum[key] / totalBOMCost : 0.0 }
    ]);
  });
  sheet2Rows.push([
    'Total',
    { f: 'SUM(B2:B9)', v: totalBOMCost },
    { f: 'SUM(C2:C9)', v: totalBOMCost > 0 ? 1.0 : 0.0 }
  ]);
  const sheet2 = XLSX.utils.aoa_to_sheet(sheet2Rows);
  for (let idx = 2; idx <= 10; idx++) {
    sheet2['B' + idx].t = 'n';
    sheet2['C' + idx].t = 'n';
    sheet2['C' + idx].z = '0.00%';
  }
  sheet2['!cols'] = [{ wch: 22 }, { wch: 18 }, { wch: 12 }];

  // Sheet 3: 物料拆解分析
  const sheet3Rows = [['項目', '金額 (NTD)', '百分比']];
  const breakdownKeys = [
    'DIODE', 'IC', 'MOS/ TR.', 'CAP.', 'DIP Res./Cap.',
    'SMD Res./Cap.', 'Magnetic', 'PCB', '機構件/ Other',
    'FAN', 'FAN GUARD', 'CASE', 'O/P WIRE'
  ];
  breakdownKeys.forEach((key, idx) => {
    const excelRow = idx + 2;
    const percentageFormula = `IFERROR(B${excelRow}/$B$16,0)`;
    let formula = '';
    const rangeC = `BOM!$C$13:$C$${formulaLastRow}`;
    const rangeB = `BOM!$B$13:$B$${formulaLastRow}`;
    const rangeD = `BOM!$D$13:$D$${formulaLastRow}`;
    const rangeL = `BOM!$L$13:$L$${formulaLastRow}`;

    switch (key) {
      case 'DIODE': formula = `SUMPRODUCT((LEFT(${rangeC},2)="1D")*(TRIM(${rangeB})<>"")*(${rangeL}))`; break;
      case 'IC': formula = `SUMPRODUCT(((LEFT(${rangeC},2)="1I")+(LEFT(${rangeC},2)="1S")+(LEFT(${rangeC},3)="1OT")>0)*(TRIM(${rangeB})<>"")*(${rangeL}))`; break;
      case 'MOS/ TR.': formula = `SUMPRODUCT(((LEFT(${rangeC},2)="1T")+(LEFT(${rangeC},2)="1M")>0)*(TRIM(${rangeB})<>"")*(${rangeL}))`; break;
      case 'CAP.': formula = `SUMPRODUCT((((LEFT(${rangeC},3)="2EC")+(LEFT(${rangeC},3)="2EG")+(LEFT(${rangeC},3)="2EL"))>0)*ISERR(SEARCH("SMD",${rangeD}))*(TRIM(${rangeB})<>"")*(${rangeL}))`; break;
      case 'DIP Res./Cap.': formula = `SUMPRODUCT((((LEFT(${rangeC},3)="2CX")+(LEFT(${rangeC},3)="2CY")+(LEFT(${rangeC},3)="2EV")+(LEFT(${rangeC},3)="2FN")+(LEFT(${rangeC},3)="2FS")+(LEFT(${rangeC},3)="2OP")+(LEFT(${rangeC},3)="2VR")+(LEFT(${rangeC},3)="3RD")+(LEFT(${rangeC},3)="3UC"))>0)*ISERR(SEARCH("SMD",${rangeD}))*(TRIM(${rangeB})<>"")*(${rangeL}))`; break;
      case 'SMD Res./Cap.': formula = `SUMPRODUCT((LEFT(${rangeC},1)="2")*ISNUMBER(SEARCH("SMD",${rangeD}))*(TRIM(${rangeB})<>"")*(${rangeL}))`; break;
      case 'Magnetic': formula = `SUMPRODUCT(((LEFT(${rangeC},1)="7")+(LEFT(${rangeC},1)="8")>0)*(LEFT(${rangeC},2)<>"8B")*(TRIM(${rangeB})<>"")*(${rangeL}))`; break;
      case 'PCB': formula = `SUMPRODUCT(((LEFT(${rangeC},2)="3B")+(LEFT(${rangeC},2)="8B")>0)*(TRIM(${rangeB})<>"")*(${rangeL}))`; break;
      case 'FAN': formula = `SUMPRODUCT(ISNUMBER(SEARCH("FAN",${rangeD}))*ISERR(SEARCH("FAN GUARD",${rangeD}))*(LEFT(${rangeC},1)<>"1")*(TRIM(${rangeB})<>"")*(${rangeL}))`; break;
      case 'FAN GUARD': formula = `SUMPRODUCT(ISNUMBER(SEARCH("FAN GUARD",${rangeD}))*(TRIM(${rangeB})<>"")*(${rangeL}))`; break;
      case 'CASE': formula = `SUMPRODUCT(((ISNUMBER(SEARCH("CASE",${rangeD})))+(ISNUMBER(SEARCH("CHASSIS",${rangeD})))>0)*(TRIM(${rangeB})<>"")*(${rangeL}))`; break;
      case 'O/P WIRE': formula = `SUMPRODUCT(((LEFT(${rangeC},2)="4W")+ISNUMBER(SEARCH("WIRE",${rangeD}))>0)*ISERR(SEARCH("GUARD",${rangeD}))*(TRIM(${rangeB})<>"")*(${rangeL}))`; break;
      case '機構件/ Other': formula = `SUMPRODUCT((TRIM(${rangeB})<>"")*(${rangeL}))-SUM(B2:B9)-SUM(B11:B14)`; break;
    }
    sheet3Rows.push([
      key,
      { f: formula, v: categoryBreakdown[key] },
      { f: percentageFormula, v: totalBOMCost > 0 ? categoryBreakdown[key] / totalBOMCost : 0.0 }
    ]);
  });
  sheet3Rows.push([
    'Total',
    { f: 'SUM(B2:B14)', v: totalBOMCost },
    { f: 'SUM(C2:C14)', v: totalBOMCost > 0 ? 1.0 : 0.0 }
  ]);
  const sheet3 = XLSX.utils.aoa_to_sheet(sheet3Rows);
  for (let idx = 2; idx <= 15; idx++) {
    sheet3['B' + idx].t = 'n';
    sheet3['C' + idx].t = 'n';
    sheet3['C' + idx].z = '0.00%';
  }
  sheet3['!cols'] = [{ wch: 22 }, { wch: 18 }, { wch: 12 }];

  const newWb = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(newWb, sheet1, 'BOM');
  XLSX.utils.book_append_sheet(newWb, sheet2, '大分類統計');
  XLSX.utils.book_append_sheet(newWb, sheet3, '物料拆解分析');

  return { filename, status: needsFix ? 'fixed' : 'raw_normal', fixedWb: newWb, totalBOMCost, rate };
}

// Main Test Loop
const files = ['good.xls', 'good 2.xls', 'good 3.xls'];
const outputDir = 'E:/Antigravity/BOM to I need/FIXED_BOM_TEST';
if (!fs.existsSync(outputDir)) {
  fs.mkdirSync(outputDir, { recursive: true });
}

console.log('=== STARTING VERIFICATION ===');
let successCount = 0;

files.forEach(f => {
  console.log(`Processing: ${f}...`);
  const buf = fs.readFileSync('E:/Antigravity/BOM to I need/' + f);
  const wb = XLSX.read(buf, { type: 'buffer' });
  
  const res = processBOMWorkbook(wb, f, 30.0);
  console.log(`- Status: ${res.status}`);
  console.log(`- Total Cost: ${res.totalBOMCost.toFixed(4)} NTD`);
  console.log(`- Exchange Rate: ${res.rate}`);

  if (res.status !== 'error') {
    const outPath = path.join(outputDir, f.replace(/\.xls$/, '_fixed.xlsx'));
    const outBuf = XLSX.write(res.fixedWb, { bookType: 'xlsx', type: 'buffer' });
    fs.writeFileSync(outPath, outBuf);
    console.log(`- Saved: ${outPath}`);
    
    const checkBuf = fs.readFileSync(outPath);
    const checkWb = XLSX.read(checkBuf, { type: 'buffer' });
    console.log(`- Worksheets check:`, checkWb.SheetNames.join(', '));
    if (checkWb.SheetNames.includes('BOM') && checkWb.SheetNames.includes('大分類統計') && checkWb.SheetNames.includes('物料拆解分析')) {
      console.log(`- [PASS] worksheets successfully validated!`);
      successCount++;
    } else {
      console.error(`- [FAIL] missing sheets in exported file!`);
    }
  } else {
    console.error(`- [FAIL] error in processing file!`);
  }
  console.log('--------------------------------');
});

console.log(`=== VERIFICATION COMPLETED: ${successCount}/${files.length} PASSED ===`);
if (successCount === files.length) {
  process.exit(0);
} else {
  process.exit(1);
}
