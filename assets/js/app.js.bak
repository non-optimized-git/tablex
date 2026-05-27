// ==================== STATE ====================
let headers = [];
let rows = [];
let selectedQuestions = [];
let groupingColIdx = null;
let groups = [];
let config = { groups: [], ordered: {}, mean_cols: {}, multi_idx: [] };

// ==================== CONSTANTS ====================
const MODULE_KEYWORDS = {
  '用户画像': ['性别', '年龄', '职业', '学历', '收入', '城市'],
  '购车意向': ['购车', '买车', '换车', '预算', '品牌'],
  '智能驾驶': ['智驾', '自动驾驶', '辅助驾驶', '泊车', '车道'],
  '家庭用车': ['家庭', '孩子', '家人', '空间'],
  '产品体验': ['满意度', '体验', '质量', '服务', '售后']
};

function detectModule(header) {
  if (!header) return '其他';
  for (const [mod, keywords] of Object.entries(MODULE_KEYWORDS)) {
    for (const kw of keywords) {
      if (header.includes(kw)) return mod;
    }
  }
  return '其他';
}

function getShortTitle(header) {
  if (!header) return '';
  // Strip HTML tags and color codes like ff2f00">text
  const cleaned = String(header).replace(/<[^>]*>/g, '').replace(/^[a-fA-F0-9]{6}">/, '').trim();
  const parts = cleaned.split(/[\s_#]+/);
  for (let i = parts.length - 1; i >= 0; i--) {
    if (parts[i].length >= 4) return parts[i];
  }
  return cleaned.slice(0, 20);
}

function escapeHtml(str) {
  if (str === null || str === undefined) return '';
  return String(str).replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;').replace(/"/g, '&quot;');
}

// ==================== EXCEL PARSER ====================
function parseExcel(file) {
  return new Promise((resolve, reject) => {
    const reader = new FileReader();
    reader.onload = (e) => {
      try {
        const data = new Uint8Array(e.target.result);
        const workbook = XLSX.read(data, { type: 'array', cellDates: true });
        const sheet = workbook.Sheets[workbook.SheetNames[0]];
        const range = XLSX.utils.decode_range(sheet['!ref']);

        const headers = [];
        for (let c = range.s.c; c <= range.e.c; c++) {
          const addr = XLSX.utils.encode_cell({ r: 0, c });
          headers.push(sheet[addr]?.v ?? null);
        }

        const rows = [];
        for (let r = 1; r <= range.e.r; r++) {
          const row = [];
          let hasValue = false;
          for (let c = range.s.c; c <= range.e.c; c++) {
            const addr = XLSX.utils.encode_cell({ r, c });
            const val = sheet[addr]?.v ?? null;
            row.push(val);
            if (val !== null && val !== '') hasValue = true;
          }
          if (hasValue) rows.push(row);
        }
        resolve({ headers, rows });
      } catch (err) {
        reject(err);
      }
    };
    reader.onerror = () => reject(new Error('读取文件失败'));
    reader.readAsArrayBuffer(file);
  });
}

// ==================== ANALYZER ====================
function safe(v) {
  if (v === null || v === undefined) return null;
  const s = String(v).trim();
  return s || null;
}

function splitOpts(val) {
  if (!val) return [];
  return String(val).split(',').map(o => o.trim()).filter(o => o);
}

function groupRowsCustom(allRows, customGroups) {
  const groups = {};
  for (const groupDef of customGroups) {
    const groupName = groupDef.name;
    const conditions = groupDef.conditions;
    groups[groupName] = [];
    for (const row of allRows) {
      let match = true;
      for (const [colIdx, expectedVal] of conditions) {
        if (safe(row[colIdx]) !== expectedVal) {
          match = false;
          break;
        }
      }
      if (match) groups[groupName].push(row);
    }
  }
  return groups;
}

function countOptions(allRows, colIdx, isMulti) {
  const counts = {};
  for (const row of allRows) {
    const val = safe(row[colIdx]);
    if (val === null) continue;
    if (isMulti) {
      for (const opt of splitOpts(val)) {
        counts[opt] = (counts[opt] || 0) + 1;
      }
    } else {
      counts[val] = (counts[val] || 0) + 1;
    }
  }
  return counts;
}

function countUsers(allRows, colIdx) {
  return allRows.filter(row => safe(row[colIdx]) !== null).length;
}

function isMultiChoice(allRows, colIdx) {
  let total = 0, hasComma = 0;
  for (const row of allRows) {
    const val = row[colIdx];
    if (val === null || val === undefined) continue;
    total++;
    if (String(val).includes(',')) hasComma++;
  }
  return total > 0 && hasComma > total * 0.2;
}

function extractMidValue(opt) {
  if (!opt) return null;
  const s = String(opt);
  let m = s.match(/^(\d+)岁以下/);
  if (m) return parseInt(m[1]);
  m = s.match(/^(\d+)岁以上/);
  if (m) return parseInt(m[1]);
  m = s.match(/^(\d+)-(\d+)/);
  if (m) return (parseInt(m[1]) + parseInt(m[2])) / 2;
  return null;
}

function calcNumericMean(allRows, colIdx) {
  let total = 0, sum = 0;
  for (const row of allRows) {
    const val = row[colIdx];
    if (val === null || val === undefined) continue;
    const strVal = String(val).trim();
    if (!strVal) continue;
    const mid = extractMidValue(strVal);
    if (mid !== null) {
      total++;
      sum += mid;
    }
  }
  return total > 0 ? sum / total : null;
}

function r3(x) {
  if (x === null || x === undefined) return null;
  return Math.abs(x) < 0.0005 ? 0 : Math.round(x * 1000) / 1000;
}

// ==================== EXCEL WRITER ====================
// ==================== EXCEL STYLE HELPERS (ExcelJS) ====================
function makeExcelJSBorder() {
  return {
    top: { style: 'thin' }, bottom: { style: 'thin' },
    left: { style: 'thin' }, right: { style: 'thin' }
  };
}

function applyTitleStyle(cell) {
  cell.font = { name: '黑体', size: 12 };
}

function applySubtitleStyle(cell) {
  cell.font = { name: '黑体', size: 9 };
}

function applyHeaderStyle(cell) {
  cell.font = { name: '黑体', size: 9, bold: true };
  cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFD9D9D9' } };
  cell.alignment = { horizontal: 'center' };
  cell.border = makeExcelJSBorder();
}

function applyDataStyle(cell) {
  cell.font = { name: '黑体', size: 9 };
  cell.border = makeExcelJSBorder();
  cell.alignment = { horizontal: 'center' };
}

function applyPctStyle(cell) {
  cell.font = { name: '黑体', size: 9 };
  cell.border = makeExcelJSBorder();
  cell.alignment = { horizontal: 'center' };
  cell.numFmt = '0.0%';
}

function applyRedStyle(cell) {
  cell.font = { name: '黑体', size: 9 };
  cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFFFCCCC' } };
  cell.border = makeExcelJSBorder();
  cell.alignment = { horizontal: 'center' };
  cell.numFmt = '0.0%';
}

function applyGreenStyle(cell) {
  cell.font = { name: '黑体', size: 9 };
  cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFCCFFCC' } };
  cell.border = makeExcelJSBorder();
  cell.alignment = { horizontal: 'center' };
  cell.numFmt = '0.0%';
}

// ==================== EXCEL GENERATION (ExcelJS) ====================

function colLetter(colIdx) {
  // colIdx 0-based: 0=A, 1=B, ..., 25=Z, 26=AA, etc.
  let result = '';
  let num = colIdx;
  while (num >= 0) {
    result = String.fromCharCode(65 + (num % 26)) + result;
    num = Math.floor(num / 26) - 1;
    if (num < 0) break;
  }
  return result;
}

async function generateExcelWorkbook() {
  const questions = selectedQuestions;
  if (!questions.length || !headers.length || !rows.length) return null;

  const orderedGroups = groups.map(g => g.name);
  const numGroups = orderedGroups.length;
  const hasGroups = numGroups > 0;

  const wb = new ExcelJS.Workbook();

  // Iterate in original order, create new sheet when module changes
  let currentModule = null;
  let ws = null;

  for (const colIdx of questions) {
    const header = headers[colIdx] || `列${colIdx}`;
    const module = detectModule(header);

    // Start new sheet when module changes
    if (module !== currentModule) {
      if (ws) wb.addWorksheet(currentModule.slice(0, 31));
      currentModule = module;
      ws = wb.addWorksheet(module.slice(0, 31));

      // Set column widths for 3-table layout
      ws.getColumn(1).width = 18;
      ws.getColumn(2).width = 10;
      ws.getColumn(3).width = 10;
      ws.getColumn(4).width = 10;
      ws.getColumn(5).width = 2;
      ws.getColumn(6).width = 18;
      ws.getColumn(7).width = 10;
      ws.getColumn(8).width = 10;
      ws.getColumn(9).width = 10;
      ws.getColumn(10).width = 2;
      ws.getColumn(11).width = 18;
      ws.getColumn(12).width = 10;
      ws.getColumn(13).width = 10;
      ws.getColumn(14).width = 10;
    }

    const title = getShortTitle(header);
    const isMulti = isMultiChoice(rows, colIdx);
    let rowNum = ws.rowCount + 1;

    // Row 1: Title
    ws.getCell(`A${rowNum}`).value = title;
    applyTitleStyle(ws.getCell(`A${rowNum}`));
    rowNum++;

    // Row 2: Subtitle
    ws.getCell(`A${rowNum}`).value = `题目: ${header} (${isMulti ? '多选' : '单选'})`;
    applySubtitleStyle(ws.getCell(`A${rowNum}`));
    rowNum++;

    if (!hasGroups) {
      // No groups: single simple table
      const counts = countOptions(rows, colIdx, isMulti);
      const totalUsers = countUsers(rows, colIdx);
      const allOpts = Object.keys(counts).sort((a, b) => counts[b] - counts[a]);

      // Header row
      ws.getCell(`A${rowNum}`).value = '选项';
      applyHeaderStyle(ws.getCell(`A${rowNum}`));
      ws.getCell(`B${rowNum}`).value = 'Total';
      applyHeaderStyle(ws.getCell(`B${rowNum}`));
      ws.getCell(`C${rowNum}`).value = '占比';
      applyHeaderStyle(ws.getCell(`C${rowNum}`));
      rowNum++;

      // Base row
      ws.getCell(`A${rowNum}`).value = 'base';
      applyDataStyle(ws.getCell(`A${rowNum}`));
      ws.getCell(`B${rowNum}`).value = totalUsers;
      applyDataStyle(ws.getCell(`B${rowNum}`));
      ws.getCell(`C${rowNum}`).value = '100%';
      applyDataStyle(ws.getCell(`C${rowNum}`));
      rowNum++;

      for (const opt of allOpts) {
        const val = counts[opt] || 0;
        const pct = totalUsers > 0 ? val / totalUsers : 0;
        ws.getCell(`A${rowNum}`).value = opt;
        applyDataStyle(ws.getCell(`A${rowNum}`));
        ws.getCell(`B${rowNum}`).value = val;
        applyDataStyle(ws.getCell(`B${rowNum}`));
        ws.getCell(`C${rowNum}`).value = pct;
        applyPctStyle(ws.getCell(`C${rowNum}`));
        rowNum++;
      }

      const meanVal = calcNumericMean(rows, colIdx);
      if (meanVal !== null) {
        ws.getCell(`A${rowNum}`).value = '平均值';
        applyDataStyle(ws.getCell(`A${rowNum}`));
        ws.getCell(`B${rowNum}`).value = meanVal;
        applyDataStyle(ws.getCell(`B${rowNum}`));
        ws.getCell(`B${rowNum}`).numFmt = '0.0';
        ws.getCell(`C${rowNum}`).value = '';
        applyDataStyle(ws.getCell(`C${rowNum}`));
        rowNum++;
      }

    } else {
      // With groups: 3 tables side by side
      const grpData = groupRowsCustom(rows, groups);
      const groupCounts = {};
      const groupUsers = {};
      for (const g of orderedGroups) {
        groupCounts[g] = countOptions(grpData[g], colIdx, isMulti);
        groupUsers[g] = countUsers(grpData[g], colIdx);
      }
      const totalCounts = countOptions(rows, colIdx, isMulti);
      const totalAllUsers = countUsers(rows, colIdx);

      const allOptsSet = new Set();
      for (const counts of Object.values(groupCounts)) {
        for (const opt of Object.keys(counts)) allOptsSet.add(opt);
      }
      for (const opt of Object.keys(totalCounts)) allOptsSet.add(opt);

      const allOptsCount = {};
      for (const counts of Object.values(groupCounts)) {
        for (const [opt, count] of Object.entries(counts)) {
          allOptsCount[opt] = (allOptsCount[opt] || 0) + count;
        }
      }
      const allOpts = [...allOptsSet].sort((a, b) => (allOptsCount[b] || 0) - (allOptsCount[a] || 0));

      // Table column positions (0-based): t1=A-D(0-3), t2=F-I(5-8), t3=K-N(10-13)
      const t1 = { opt: 0, tot: 1, g1: 2, g2: 3 };
      const t2 = { opt: 5, tot: 6, g1: 7, g2: 8 };
      const t3 = { opt: 10, tot: 11, g1: 12, g2: 13 };

      // Row 3: Headers
      const headerRow = rowNum;
      for (const t of [t1, t2, t3]) {
        ws.getCell(`${colLetter(t.opt)}${headerRow}`).value = '选项';
        applyHeaderStyle(ws.getCell(`${colLetter(t.opt)}${headerRow}`));
        ws.getCell(`${colLetter(t.tot)}${headerRow}`).value = 'Total';
        applyHeaderStyle(ws.getCell(`${colLetter(t.tot)}${headerRow}`));
        ws.getCell(`${colLetter(t.g1)}${headerRow}`).value = orderedGroups[0];
        applyHeaderStyle(ws.getCell(`${colLetter(t.g1)}${headerRow}`));
        ws.getCell(`${colLetter(t.g2)}${headerRow}`).value = orderedGroups[1];
        applyHeaderStyle(ws.getCell(`${colLetter(t.g2)}${headerRow}`));
      }
      rowNum++;

      // Row 4: Base row
      for (const t of [t1, t2, t3]) {
        ws.getCell(`${colLetter(t.opt)}${rowNum}`).value = 'base';
        applyDataStyle(ws.getCell(`${colLetter(t.opt)}${rowNum}`));
        ws.getCell(`${colLetter(t.tot)}${rowNum}`).value = totalAllUsers;
        applyDataStyle(ws.getCell(`${colLetter(t.tot)}${rowNum}`));
        ws.getCell(`${colLetter(t.g1)}${rowNum}`).value = groupUsers[orderedGroups[0]] || 0;
        applyDataStyle(ws.getCell(`${colLetter(t.g1)}${rowNum}`));
        ws.getCell(`${colLetter(t.g2)}${rowNum}`).value = groupUsers[orderedGroups[1]] || 0;
        applyDataStyle(ws.getCell(`${colLetter(t.g2)}${rowNum}`));
      }
      rowNum++;

      for (const opt of allOpts) {
        const pcts = {};
        for (const g of orderedGroups) {
          const val = groupCounts[g][opt] || 0;
          const total = groupUsers[g] || 0;
          pcts[g] = total > 0 ? val / total : 0;
        }
        const totalVal = totalCounts[opt] || 0;
        const totalPct = totalAllUsers > 0 ? totalVal / totalAllUsers : 0;

        const allPcts = [totalPct, pcts[orderedGroups[0]], pcts[orderedGroups[1]]];
        const maxPct = Math.max(...allPcts);
        const nonZeroPcts = allPcts.filter(p => p > 0);
        const minPct = nonZeroPcts.length > 0 ? Math.min(...nonZeroPcts) : null;
        const hasMultipleDistinctValues = maxPct > 0 && minPct !== null && maxPct !== minPct;

        // Table 1: counts
        ws.getCell(`${colLetter(t1.opt)}${rowNum}`).value = opt;
        applyDataStyle(ws.getCell(`${colLetter(t1.opt)}${rowNum}`));
        ws.getCell(`${colLetter(t1.tot)}${rowNum}`).value = totalVal;
        applyDataStyle(ws.getCell(`${colLetter(t1.tot)}${rowNum}`));
        ws.getCell(`${colLetter(t1.g1)}${rowNum}`).value = groupCounts[orderedGroups[0]][opt] || 0;
        applyDataStyle(ws.getCell(`${colLetter(t1.g1)}${rowNum}`));
        ws.getCell(`${colLetter(t1.g2)}${rowNum}`).value = groupCounts[orderedGroups[1]][opt] || 0;
        applyDataStyle(ws.getCell(`${colLetter(t1.g2)}${rowNum}`));

        // Table 2: percentages (no fill)
        ws.getCell(`${colLetter(t2.opt)}${rowNum}`).value = opt;
        applyDataStyle(ws.getCell(`${colLetter(t2.opt)}${rowNum}`));
        ws.getCell(`${colLetter(t2.tot)}${rowNum}`).value = totalPct;
        applyPctStyle(ws.getCell(`${colLetter(t2.tot)}${rowNum}`));
        ws.getCell(`${colLetter(t2.g1)}${rowNum}`).value = pcts[orderedGroups[0]];
        applyPctStyle(ws.getCell(`${colLetter(t2.g1)}${rowNum}`));
        ws.getCell(`${colLetter(t2.g2)}${rowNum}`).value = pcts[orderedGroups[1]];
        applyPctStyle(ws.getCell(`${colLetter(t2.g2)}${rowNum}`));

        // Table 3: percentages with red (max) / green (min)
        ws.getCell(`${colLetter(t3.opt)}${rowNum}`).value = opt;
        applyDataStyle(ws.getCell(`${colLetter(t3.opt)}${rowNum}`));

        // Total column
        if (totalPct === maxPct && maxPct > 0 && hasMultipleDistinctValues) {
          ws.getCell(`${colLetter(t3.tot)}${rowNum}`).value = totalPct;
          applyRedStyle(ws.getCell(`${colLetter(t3.tot)}${rowNum}`));
        } else if (totalPct === minPct && minPct !== null && hasMultipleDistinctValues) {
          ws.getCell(`${colLetter(t3.tot)}${rowNum}`).value = totalPct;
          applyGreenStyle(ws.getCell(`${colLetter(t3.tot)}${rowNum}`));
        } else {
          ws.getCell(`${colLetter(t3.tot)}${rowNum}`).value = totalPct;
          applyPctStyle(ws.getCell(`${colLetter(t3.tot)}${rowNum}`));
        }

        // Group 1 column
        const g1Pct = pcts[orderedGroups[0]];
        if (g1Pct === maxPct && maxPct > 0 && hasMultipleDistinctValues) {
          ws.getCell(`${colLetter(t3.g1)}${rowNum}`).value = g1Pct;
          applyRedStyle(ws.getCell(`${colLetter(t3.g1)}${rowNum}`));
        } else if (g1Pct === minPct && minPct !== null && hasMultipleDistinctValues) {
          ws.getCell(`${colLetter(t3.g1)}${rowNum}`).value = g1Pct;
          applyGreenStyle(ws.getCell(`${colLetter(t3.g1)}${rowNum}`));
        } else {
          ws.getCell(`${colLetter(t3.g1)}${rowNum}`).value = g1Pct;
          applyPctStyle(ws.getCell(`${colLetter(t3.g1)}${rowNum}`));
        }

        // Group 2 column
        const g2Pct = pcts[orderedGroups[1]];
        if (g2Pct === maxPct && maxPct > 0 && hasMultipleDistinctValues) {
          ws.getCell(`${colLetter(t3.g2)}${rowNum}`).value = g2Pct;
          applyRedStyle(ws.getCell(`${colLetter(t3.g2)}${rowNum}`));
        } else if (g2Pct === minPct && minPct !== null && hasMultipleDistinctValues) {
          ws.getCell(`${colLetter(t3.g2)}${rowNum}`).value = g2Pct;
          applyGreenStyle(ws.getCell(`${colLetter(t3.g2)}${rowNum}`));
        } else {
          ws.getCell(`${colLetter(t3.g2)}${rowNum}`).value = g2Pct;
          applyPctStyle(ws.getCell(`${colLetter(t3.g2)}${rowNum}`));
        }

        rowNum++;
      }

      // Mean row
      const means = {};
      for (const g of orderedGroups) {
        means[g] = calcNumericMean(grpData[g], colIdx);
      }
      const totalMean = calcNumericMean(rows, colIdx);

      if (totalMean !== null || Object.values(means).some(m => m !== null)) {
        const allMeans = Object.values(means).filter(m => m !== null);
        if (totalMean !== null) allMeans.push(totalMean);
        const maxM = allMeans.length > 0 ? Math.max(...allMeans) : null;
        const minM = allMeans.length > 0 ? Math.min(...allMeans.filter(m => m > 0)) : null;
        const hasMultipleDistinctMeans = maxM !== null && minM !== null && maxM !== minM;

        // Table 1 mean
        ws.getCell(`${colLetter(t1.opt)}${rowNum}`).value = '平均值';
        applyDataStyle(ws.getCell(`${colLetter(t1.opt)}${rowNum}`));
        ws.getCell(`${colLetter(t1.tot)}${rowNum}`).value = totalMean !== null ? totalMean : 0;
        applyDataStyle(ws.getCell(`${colLetter(t1.tot)}${rowNum}`));
        ws.getCell(`${colLetter(t1.tot)}${rowNum}`).numFmt = '0.0';
        ws.getCell(`${colLetter(t1.g1)}${rowNum}`).value = means[orderedGroups[0]] !== null ? means[orderedGroups[0]] : 0;
        applyDataStyle(ws.getCell(`${colLetter(t1.g1)}${rowNum}`));
        ws.getCell(`${colLetter(t1.g1)}${rowNum}`).numFmt = '0.0';
        ws.getCell(`${colLetter(t1.g2)}${rowNum}`).value = means[orderedGroups[1]] !== null ? means[orderedGroups[1]] : 0;
        applyDataStyle(ws.getCell(`${colLetter(t1.g2)}${rowNum}`));
        ws.getCell(`${colLetter(t1.g2)}${rowNum}`).numFmt = '0.0';

        // Table 2 mean
        ws.getCell(`${colLetter(t2.opt)}${rowNum}`).value = '平均值';
        applyDataStyle(ws.getCell(`${colLetter(t2.opt)}${rowNum}`));
        ws.getCell(`${colLetter(t2.tot)}${rowNum}`).value = totalMean !== null ? totalMean : 0;
        applyDataStyle(ws.getCell(`${colLetter(t2.tot)}${rowNum}`));
        ws.getCell(`${colLetter(t2.tot)}${rowNum}`).numFmt = '0.0';
        ws.getCell(`${colLetter(t2.g1)}${rowNum}`).value = means[orderedGroups[0]] !== null ? means[orderedGroups[0]] : 0;
        applyDataStyle(ws.getCell(`${colLetter(t2.g1)}${rowNum}`));
        ws.getCell(`${colLetter(t2.g1)}${rowNum}`).numFmt = '0.0';
        ws.getCell(`${colLetter(t2.g2)}${rowNum}`).value = means[orderedGroups[1]] !== null ? means[orderedGroups[1]] : 0;
        applyDataStyle(ws.getCell(`${colLetter(t2.g2)}${rowNum}`));
        ws.getCell(`${colLetter(t2.g2)}${rowNum}`).numFmt = '0.0';

        // Table 3 mean with red/green
        ws.getCell(`${colLetter(t3.opt)}${rowNum}`).value = '平均值';
        applyDataStyle(ws.getCell(`${colLetter(t3.opt)}${rowNum}`));

        const m0 = totalMean;
        if (m0 !== null && m0 === maxM && hasMultipleDistinctMeans) {
          ws.getCell(`${colLetter(t3.tot)}${rowNum}`).value = m0;
          applyRedStyle(ws.getCell(`${colLetter(t3.tot)}${rowNum}`));
        } else if (m0 !== null && m0 === minM && hasMultipleDistinctMeans) {
          ws.getCell(`${colLetter(t3.tot)}${rowNum}`).value = m0;
          applyGreenStyle(ws.getCell(`${colLetter(t3.tot)}${rowNum}`));
        } else {
          ws.getCell(`${colLetter(t3.tot)}${rowNum}`).value = m0 !== null ? m0 : 0;
          applyDataStyle(ws.getCell(`${colLetter(t3.tot)}${rowNum}`));
        }
        ws.getCell(`${colLetter(t3.tot)}${rowNum}`).numFmt = '0.0';

        const m1 = means[orderedGroups[0]];
        if (m1 !== null && m1 === maxM && hasMultipleDistinctMeans) {
          ws.getCell(`${colLetter(t3.g1)}${rowNum}`).value = m1;
          applyRedStyle(ws.getCell(`${colLetter(t3.g1)}${rowNum}`));
        } else if (m1 !== null && m1 === minM && hasMultipleDistinctMeans) {
          ws.getCell(`${colLetter(t3.g1)}${rowNum}`).value = m1;
          applyGreenStyle(ws.getCell(`${colLetter(t3.g1)}${rowNum}`));
        } else {
          ws.getCell(`${colLetter(t3.g1)}${rowNum}`).value = m1 !== null ? m1 : 0;
          applyDataStyle(ws.getCell(`${colLetter(t3.g1)}${rowNum}`));
        }
        ws.getCell(`${colLetter(t3.g1)}${rowNum}`).numFmt = '0.0';

        const m2 = means[orderedGroups[1]];
        if (m2 !== null && m2 === maxM && hasMultipleDistinctMeans) {
          ws.getCell(`${colLetter(t3.g2)}${rowNum}`).value = m2;
          applyRedStyle(ws.getCell(`${colLetter(t3.g2)}${rowNum}`));
        } else if (m2 !== null && m2 === minM && hasMultipleDistinctMeans) {
          ws.getCell(`${colLetter(t3.g2)}${rowNum}`).value = m2;
          applyGreenStyle(ws.getCell(`${colLetter(t3.g2)}${rowNum}`));
        } else {
          ws.getCell(`${colLetter(t3.g2)}${rowNum}`).value = m2 !== null ? m2 : 0;
          applyDataStyle(ws.getCell(`${colLetter(t3.g2)}${rowNum}`));
        }
        ws.getCell(`${colLetter(t3.g2)}${rowNum}`).numFmt = '0.0';

        rowNum++;
      }
    }
  }

  return wb;
}
        const groupCounts = {};
        const groupUsers = {};
        for (const g of orderedGroups) {
          groupCounts[g] = countOptions(grpData[g], colIdx, isMulti);
          groupUsers[g] = countUsers(grpData[g], colIdx);
        }
        const totalCounts = countOptions(rows, colIdx, isMulti);
        const totalAllUsers = countUsers(rows, colIdx);

        const allOptsSet = new Set();
        for (const counts of Object.values(groupCounts)) {
          for (const opt of Object.keys(counts)) allOptsSet.add(opt);
        }
        for (const opt of Object.keys(totalCounts)) allOptsSet.add(opt);

        const allOptsCount = {};
        for (const counts of Object.values(groupCounts)) {
          for (const [opt, count] of Object.entries(counts)) {
            allOptsCount[opt] = (allOptsCount[opt] || 0) + count;
          }
        }
        const allOpts = [...allOptsSet].sort((a, b) => (allOptsCount[b] || 0) - (allOptsCount[a] || 0));

        // Table column positions (0-based): t1=A-D(0-3), t2=F-I(5-8), t3=K-N(10-13)
        const t1 = { opt: 0, tot: 1, g1: 2, g2: 3 };
        const t2 = { opt: 5, tot: 6, g1: 7, g2: 8 };
        const t3 = { opt: 10, tot: 11, g1: 12, g2: 13 };

        // Row 3: Headers
        const headerRow = rowNum;
        for (const t of [t1, t2, t3]) {
          ws.getCell(`${colLetter(t.opt)}${headerRow}`).value = '选项';
          applyHeaderStyle(ws.getCell(`${colLetter(t.opt)}${headerRow}`));
          ws.getCell(`${colLetter(t.tot)}${headerRow}`).value = 'Total';
          applyHeaderStyle(ws.getCell(`${colLetter(t.tot)}${headerRow}`));
          ws.getCell(`${colLetter(t.g1)}${headerRow}`).value = orderedGroups[0];
          applyHeaderStyle(ws.getCell(`${colLetter(t.g1)}${headerRow}`));
          ws.getCell(`${colLetter(t.g2)}${headerRow}`).value = orderedGroups[1];
          applyHeaderStyle(ws.getCell(`${colLetter(t.g2)}${headerRow}`));
        }
        rowNum++;

        // Row 4: Base row
        for (const t of [t1, t2, t3]) {
          ws.getCell(`${colLetter(t.opt)}${rowNum}`).value = 'base';
          applyDataStyle(ws.getCell(`${colLetter(t.opt)}${rowNum}`));
          ws.getCell(`${colLetter(t.tot)}${rowNum}`).value = totalAllUsers;
          applyDataStyle(ws.getCell(`${colLetter(t.tot)}${rowNum}`));
          ws.getCell(`${colLetter(t.g1)}${rowNum}`).value = groupUsers[orderedGroups[0]] || 0;
          applyDataStyle(ws.getCell(`${colLetter(t.g1)}${rowNum}`));
          ws.getCell(`${colLetter(t.g2)}${rowNum}`).value = groupUsers[orderedGroups[1]] || 0;
          applyDataStyle(ws.getCell(`${colLetter(t.g2)}${rowNum}`));
        }
        rowNum++;

        for (const opt of allOpts) {
          const pcts = {};
          for (const g of orderedGroups) {
            const val = groupCounts[g][opt] || 0;
            const total = groupUsers[g] || 0;
            pcts[g] = total > 0 ? val / total : 0;
          }
          const totalVal = totalCounts[opt] || 0;
          const totalPct = totalAllUsers > 0 ? totalVal / totalAllUsers : 0;

          const allPcts = [totalPct, pcts[orderedGroups[0]], pcts[orderedGroups[1]]];
          const maxPct = Math.max(...allPcts);
          const nonZeroPcts = allPcts.filter(p => p > 0);
          const minPct = nonZeroPcts.length > 0 ? Math.min(...nonZeroPcts) : null;
          const hasMultipleDistinctValues = maxPct > 0 && minPct !== null && maxPct !== minPct;

          // Table 1: counts
          ws.getCell(`${colLetter(t1.opt)}${rowNum}`).value = opt;
          applyDataStyle(ws.getCell(`${colLetter(t1.opt)}${rowNum}`));
          ws.getCell(`${colLetter(t1.tot)}${rowNum}`).value = totalVal;
          applyDataStyle(ws.getCell(`${colLetter(t1.tot)}${rowNum}`));
          ws.getCell(`${colLetter(t1.g1)}${rowNum}`).value = groupCounts[orderedGroups[0]][opt] || 0;
          applyDataStyle(ws.getCell(`${colLetter(t1.g1)}${rowNum}`));
          ws.getCell(`${colLetter(t1.g2)}${rowNum}`).value = groupCounts[orderedGroups[1]][opt] || 0;
          applyDataStyle(ws.getCell(`${colLetter(t1.g2)}${rowNum}`));

          // Table 2: percentages (no fill)
          ws.getCell(`${colLetter(t2.opt)}${rowNum}`).value = opt;
          applyDataStyle(ws.getCell(`${colLetter(t2.opt)}${rowNum}`));
          ws.getCell(`${colLetter(t2.tot)}${rowNum}`).value = totalPct;
          applyPctStyle(ws.getCell(`${colLetter(t2.tot)}${rowNum}`));
          ws.getCell(`${colLetter(t2.g1)}${rowNum}`).value = pcts[orderedGroups[0]];
          applyPctStyle(ws.getCell(`${colLetter(t2.g1)}${rowNum}`));
          ws.getCell(`${colLetter(t2.g2)}${rowNum}`).value = pcts[orderedGroups[1]];
          applyPctStyle(ws.getCell(`${colLetter(t2.g2)}${rowNum}`));

          // Table 3: percentages with red (max) / green (min)
          ws.getCell(`${colLetter(t3.opt)}${rowNum}`).value = opt;
          applyDataStyle(ws.getCell(`${colLetter(t3.opt)}${rowNum}`));

          // Total column
          if (totalPct === maxPct && maxPct > 0 && hasMultipleDistinctValues) {
            ws.getCell(`${colLetter(t3.tot)}${rowNum}`).value = totalPct;
            applyRedStyle(ws.getCell(`${colLetter(t3.tot)}${rowNum}`));
          } else if (totalPct === minPct && minPct !== null && hasMultipleDistinctValues) {
            ws.getCell(`${colLetter(t3.tot)}${rowNum}`).value = totalPct;
            applyGreenStyle(ws.getCell(`${colLetter(t3.tot)}${rowNum}`));
          } else {
            ws.getCell(`${colLetter(t3.tot)}${rowNum}`).value = totalPct;
            applyPctStyle(ws.getCell(`${colLetter(t3.tot)}${rowNum}`));
          }

          // Group 1 column
          const g1Pct = pcts[orderedGroups[0]];
          if (g1Pct === maxPct && maxPct > 0 && hasMultipleDistinctValues) {
            ws.getCell(`${colLetter(t3.g1)}${rowNum}`).value = g1Pct;
            applyRedStyle(ws.getCell(`${colLetter(t3.g1)}${rowNum}`));
          } else if (g1Pct === minPct && minPct !== null && hasMultipleDistinctValues) {
            ws.getCell(`${colLetter(t3.g1)}${rowNum}`).value = g1Pct;
            applyGreenStyle(ws.getCell(`${colLetter(t3.g1)}${rowNum}`));
          } else {
            ws.getCell(`${colLetter(t3.g1)}${rowNum}`).value = g1Pct;
            applyPctStyle(ws.getCell(`${colLetter(t3.g1)}${rowNum}`));
          }

          // Group 2 column
          const g2Pct = pcts[orderedGroups[1]];
          if (g2Pct === maxPct && maxPct > 0 && hasMultipleDistinctValues) {
            ws.getCell(`${colLetter(t3.g2)}${rowNum}`).value = g2Pct;
            applyRedStyle(ws.getCell(`${colLetter(t3.g2)}${rowNum}`));
          } else if (g2Pct === minPct && minPct !== null && hasMultipleDistinctValues) {
            ws.getCell(`${colLetter(t3.g2)}${rowNum}`).value = g2Pct;
            applyGreenStyle(ws.getCell(`${colLetter(t3.g2)}${rowNum}`));
          } else {
            ws.getCell(`${colLetter(t3.g2)}${rowNum}`).value = g2Pct;
            applyPctStyle(ws.getCell(`${colLetter(t3.g2)}${rowNum}`));
          }

          rowNum++;
        }

        // Mean row
        const means = {};
        for (const g of orderedGroups) {
          means[g] = calcNumericMean(grpData[g], colIdx);
        }
        const totalMean = calcNumericMean(rows, colIdx);

        if (totalMean !== null || Object.values(means).some(m => m !== null)) {
          const allMeans = Object.values(means).filter(m => m !== null);
          if (totalMean !== null) allMeans.push(totalMean);
          const maxM = allMeans.length > 0 ? Math.max(...allMeans) : null;
          const minM = allMeans.length > 0 ? Math.min(...allMeans.filter(m => m > 0)) : null;
          const hasMultipleDistinctMeans = maxM !== null && minM !== null && maxM !== minM;

          // Table 1 mean
          ws.getCell(`${colLetter(t1.opt)}${rowNum}`).value = '平均值';
          applyDataStyle(ws.getCell(`${colLetter(t1.opt)}${rowNum}`));
          ws.getCell(`${colLetter(t1.tot)}${rowNum}`).value = totalMean !== null ? totalMean : 0;
          applyDataStyle(ws.getCell(`${colLetter(t1.tot)}${rowNum}`));
          ws.getCell(`${colLetter(t1.tot)}${rowNum}`).numFmt = '0.0';
          ws.getCell(`${colLetter(t1.g1)}${rowNum}`).value = means[orderedGroups[0]] !== null ? means[orderedGroups[0]] : 0;
          applyDataStyle(ws.getCell(`${colLetter(t1.g1)}${rowNum}`));
          ws.getCell(`${colLetter(t1.g1)}${rowNum}`).numFmt = '0.0';
          ws.getCell(`${colLetter(t1.g2)}${rowNum}`).value = means[orderedGroups[1]] !== null ? means[orderedGroups[1]] : 0;
          applyDataStyle(ws.getCell(`${colLetter(t1.g2)}${rowNum}`));
          ws.getCell(`${colLetter(t1.g2)}${rowNum}`).numFmt = '0.0';

          // Table 2 mean
          ws.getCell(`${colLetter(t2.opt)}${rowNum}`).value = '平均值';
          applyDataStyle(ws.getCell(`${colLetter(t2.opt)}${rowNum}`));
          ws.getCell(`${colLetter(t2.tot)}${rowNum}`).value = totalMean !== null ? totalMean : 0;
          applyDataStyle(ws.getCell(`${colLetter(t2.tot)}${rowNum}`));
          ws.getCell(`${colLetter(t2.tot)}${rowNum}`).numFmt = '0.0';
          ws.getCell(`${colLetter(t2.g1)}${rowNum}`).value = means[orderedGroups[0]] !== null ? means[orderedGroups[0]] : 0;
          applyDataStyle(ws.getCell(`${colLetter(t2.g1)}${rowNum}`));
          ws.getCell(`${colLetter(t2.g1)}${rowNum}`).numFmt = '0.0';
          ws.getCell(`${colLetter(t2.g2)}${rowNum}`).value = means[orderedGroups[1]] !== null ? means[orderedGroups[1]] : 0;
          applyDataStyle(ws.getCell(`${colLetter(t2.g2)}${rowNum}`));
          ws.getCell(`${colLetter(t2.g2)}${rowNum}`).numFmt = '0.0';

          // Table 3 mean with red/green
          ws.getCell(`${colLetter(t3.opt)}${rowNum}`).value = '平均值';
          applyDataStyle(ws.getCell(`${colLetter(t3.opt)}${rowNum}`));

          const m0 = totalMean;
          if (m0 !== null && m0 === maxM && hasMultipleDistinctMeans) {
            ws.getCell(`${colLetter(t3.tot)}${rowNum}`).value = m0;
            applyRedStyle(ws.getCell(`${colLetter(t3.tot)}${rowNum}`));
          } else if (m0 !== null && m0 === minM && hasMultipleDistinctMeans) {
            ws.getCell(`${colLetter(t3.tot)}${rowNum}`).value = m0;
            applyGreenStyle(ws.getCell(`${colLetter(t3.tot)}${rowNum}`));
          } else {
            ws.getCell(`${colLetter(t3.tot)}${rowNum}`).value = m0 !== null ? m0 : 0;
            applyDataStyle(ws.getCell(`${colLetter(t3.tot)}${rowNum}`));
          }
          ws.getCell(`${colLetter(t3.tot)}${rowNum}`).numFmt = '0.0';

          const m1 = means[orderedGroups[0]];
          if (m1 !== null && m1 === maxM && hasMultipleDistinctMeans) {
            ws.getCell(`${colLetter(t3.g1)}${rowNum}`).value = m1;
            applyRedStyle(ws.getCell(`${colLetter(t3.g1)}${rowNum}`));
          } else if (m1 !== null && m1 === minM && hasMultipleDistinctMeans) {
            ws.getCell(`${colLetter(t3.g1)}${rowNum}`).value = m1;
            applyGreenStyle(ws.getCell(`${colLetter(t3.g1)}${rowNum}`));
          } else {
            ws.getCell(`${colLetter(t3.g1)}${rowNum}`).value = m1 !== null ? m1 : 0;
            applyDataStyle(ws.getCell(`${colLetter(t3.g1)}${rowNum}`));
          }
          ws.getCell(`${colLetter(t3.g1)}${rowNum}`).numFmt = '0.0';

          const m2 = means[orderedGroups[1]];
          if (m2 !== null && m2 === maxM && hasMultipleDistinctMeans) {
            ws.getCell(`${colLetter(t3.g2)}${rowNum}`).value = m2;
            applyRedStyle(ws.getCell(`${colLetter(t3.g2)}${rowNum}`));
          } else if (m2 !== null && m2 === minM && hasMultipleDistinctMeans) {
            ws.getCell(`${colLetter(t3.g2)}${rowNum}`).value = m2;
            applyGreenStyle(ws.getCell(`${colLetter(t3.g2)}${rowNum}`));
          } else {
            ws.getCell(`${colLetter(t3.g2)}${rowNum}`).value = m2 !== null ? m2 : 0;
            applyDataStyle(ws.getCell(`${colLetter(t3.g2)}${rowNum}`));
          }
          ws.getCell(`${colLetter(t3.g2)}${rowNum}`).numFmt = '0.0';

          rowNum++;
        }
      }
    }
  }

  return wb;
}

async function downloadExcel() {
  try {
    console.log('selectedQuestions:', selectedQuestions);
    console.log('groups:', groups);
    const wb = await generateExcelWorkbook();
    if (!wb) {
      alert('请先上传 Excel 文件并选择题目');
      return;
    }
    const buf = await wb.xlsx.writeBuffer();
    const blob = new Blob([buf], { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' });
    const url = URL.createObjectURL(blob);
    const a = document.createElement('a');
    a.href = url;
    a.download = '分析结果.xlsx';
    document.body.appendChild(a);
    a.click();
    document.body.removeChild(a);
    URL.revokeObjectURL(url);
  } catch (err) {
    console.error('Excel export error:', err);
    alert('导出失败: ' + err.message);
  }
}

// Old XLSX-based functions removed - using ExcelJS now

// ==================== UI HANDLERS ====================
async function handleFile(file) {
  if (!file) return;
  try {
    const result = await parseExcel(file);
    headers = result.headers;
    rows = result.rows;
    selectedQuestions = headers.map((_, i) => i);
    groupingColIdx = null;
    groups = [];
    const hint = document.getElementById('uploadHint');
    if (hint) hint.textContent = `${file.name} 已上传，共 ${rows.length} 行数据`;
    render();
    loadSavedConfig();
  } catch (err) {
    alert('读取文件失败: ' + err.message);
  }
}

function handleGroupColChange(e) {
  groupingColIdx = e.target.value === '' ? null : parseInt(e.target.value);
  groups = [];
  render();
  saveConfig();
}

function toggleGroupValue(val) {
  const exists = groups.find(g => g.name === val && g.conditions[0][0] === groupingColIdx);
  if (exists) {
    groups = groups.filter(g => g !== exists);
  } else {
    groups.push({ name: val, conditions: [[groupingColIdx, val]] });
  }
  render();
  saveConfig();
}

function selectAllQuestions() {
  selectedQuestions = headers.map((_, i) => i);
  render();
  saveConfig();
}

function selectNoneQuestions() {
  selectedQuestions = [];
  render();
  saveConfig();
}

function toggleQuestion(idx) {
  if (selectedQuestions.includes(idx)) {
    selectedQuestions = selectedQuestions.filter(i => i !== idx);
  } else {
    selectedQuestions.push(idx);
  }
  render();
  saveConfig();
}

// Drag and drop for groups
let dragVal = null;
let dropTargetVal = null;

function handleGroupDragStart(e, val) {
  dragVal = val;
  e.dataTransfer.effectAllowed = 'move';
  e.currentTarget.classList.add('dragging-item');
}

function handleGroupDragEnter(e, val) {
  e.preventDefault();
  if (dragVal !== null && dragVal !== val) {
    dropTargetVal = val;
    updateDropIndicator(val);
  }
}

function handleGroupDragLeave(e) {
  // Only clear if we're leaving to outside the drop zone
  if (!e.currentTarget.contains(e.relatedTarget)) {
    dropTargetVal = null;
    clearDropIndicator();
  }
}

function handleGroupDragOver(e, val) {
  e.preventDefault();
  e.dataTransfer.dropEffect = 'move';
}

function handleGroupDrop(e, val) {
  e.preventDefault();
  if (dragVal === null || dragVal === val) {
    dragVal = null;
    dropTargetVal = null;
    clearDropIndicator();
    return;
  }
  const dragIdx = groups.findIndex(g => g.name === dragVal && g.conditions[0][0] === groupingColIdx);
  const targetIdx = groups.findIndex(g => g.name === val && g.conditions[0][0] === groupingColIdx);
  if (dragIdx === -1 || targetIdx === -1) {
    dragVal = null;
    dropTargetVal = null;
    clearDropIndicator();
    return;
  }
  const [removed] = groups.splice(dragIdx, 1);
  groups.splice(targetIdx, 0, removed);
  dragVal = null;
  dropTargetVal = null;
  clearDropIndicator();
  render();
  saveConfig();
}

function handleGroupDragEnd(e) {
  e.currentTarget.classList.remove('dragging-item');
  dragVal = null;
  dropTargetVal = null;
  clearDropIndicator();
  saveConfig();
}

function updateDropIndicator(targetVal) {
  const items = document.querySelectorAll('.group-tag');
  items.forEach(item => {
    const itemVal = item.getAttribute('data-val');
    if (itemVal === targetVal) {
      item.classList.add('drop-target');
    } else {
      item.classList.remove('drop-target');
    }
  });
}

function clearDropIndicator() {
  document.querySelectorAll('.group-tag').forEach(item => {
    item.classList.remove('drop-target');
  });
}

// ==================== RENDER ====================
function render() {
  const uploadZone = document.getElementById('uploadZone');
  uploadZone.classList.remove('dragging');

  const groupCard = document.getElementById('groupSelectorCard');
  const groupColSelect = document.getElementById('groupColSelect');
  const groupValuesContainer = document.getElementById('groupValuesContainer');

  if (headers.length === 0) {
    groupCard.style.display = 'none';
    document.getElementById('questionSelectorCard').style.display = 'none';
    document.getElementById('resultPreviewCard').style.display = 'none';
    return;
  }

  groupCard.style.display = 'block';

  const colOptions = headers.map((h, i) => ({ idx: i, name: getShortTitle(h) || `列${i + 1}` })).filter(h => h.name);
  groupColSelect.innerHTML = '<option value="">-- 选择分组列 --</option>' +
    colOptions.map(opt => `<option value="${opt.idx}" ${groupingColIdx === opt.idx ? 'selected' : ''}>${opt.name}</option>`).join('');

  if (groupingColIdx === null) {
    groupValuesContainer.style.display = 'none';
  } else {
    groupValuesContainer.style.display = 'block';

    const values = new Set();
    for (const row of rows) {
      const v = safe(row[groupingColIdx]);
      if (v !== null) values.add(v);
    }
    const allValues = [...values].sort();
    // Show selected values first (in groups order), then unselected (alphabetically)
    const selectedVals = groups.filter(g => g.conditions[0][0] === groupingColIdx).map(g => g.name);
    const unselectedVals = allValues.filter(v => !selectedVals.includes(v));
    const sortedValues = [...selectedVals, ...unselectedVals];

    document.getElementById('groupValues').innerHTML = sortedValues.map((val) => `
      <div class="group-tag ${groups.some(g => g.name === val && g.conditions[0][0] === groupingColIdx) ? 'selected' : 'unselected'}"
           draggable="true"
           data-val="${escapeHtml(val)}"
           ondragstart="handleGroupDragStart(event, '${val.replace(/'/g, "\\'")}')"
           ondragenter="handleGroupDragEnter(event, '${val.replace(/'/g, "\\'")}')"
           ondragleave="handleGroupDragLeave(event)"
           ondragover="handleGroupDragOver(event, '${val.replace(/'/g, "\\'")}')"
           ondrop="handleGroupDrop(event, '${val.replace(/'/g, "\\'")}')"
           ondragend="handleGroupDragEnd(event)"
           onclick="toggleGroupValue('${val.replace(/'/g, "\\'")}')">
        <span class="drag-handle">⋮⋮</span>
        <span>${escapeHtml(val)}</span>
      </div>
    `).join('');
  }

  const qCard = document.getElementById('questionSelectorCard');
  if (headers.length === 0) {
    qCard.style.display = 'none';
    return;
  }
  qCard.style.display = 'block';

  const modules = {};
  for (let i = 0; i < headers.length; i++) {
    const mod = detectModule(headers[i]);
    const title = getShortTitle(headers[i]) || `列${i + 1}`;
    const isMulti = rows.length > 0 && isMultiChoice(rows, i);
    if (!modules[mod]) modules[mod] = [];
    modules[mod].push({ idx: i, title, isMulti });
  }

  document.getElementById('questionCount').textContent = `已选择 ${selectedQuestions.length} / ${headers.length} 题`;

  document.getElementById('questionList').innerHTML = Object.entries(modules).map(([modName, questions]) => `
    <div class="question-module">
      <h4>${escapeHtml(modName)}</h4>
      ${questions.map(q => `
        <label class="question-item">
          <input type="checkbox" ${selectedQuestions.includes(q.idx) ? 'checked' : ''} onchange="toggleQuestion(${q.idx})">
          <span>${escapeHtml(q.title)}</span>
          <span class="type-tag">${q.isMulti ? '多选' : '单选'}</span>
        </label>
      `).join('')}
    </div>
  `).join('');

  const rCard = document.getElementById('resultPreviewCard');
  if (headers.length === 0 || selectedQuestions.length === 0) {
    rCard.style.display = 'none';
    return;
  }
  rCard.style.display = 'block';
  document.getElementById('downloadBtn').disabled = selectedQuestions.length === 0;

  const grpData = groupRowsCustom(rows, groups);
  const orderedGroups = groups.map(g => g.name);
  const hasGroups = orderedGroups.length > 0;

  const previewHTML = [];

  for (const colIdx of selectedQuestions) {
    const header = headers[colIdx] || `列${colIdx}`;
    const title = getShortTitle(header);
    const isMulti = isMultiChoice(rows, colIdx);

    if (!hasGroups) {
      const counts = countOptions(rows, colIdx, isMulti);
      const totalUsers = countUsers(rows, colIdx);
      const allOpts = Object.keys(counts).sort((a, b) => counts[b] - counts[a]);

      let html = `<div class="result-item">`;
      html += `<h4>${escapeHtml(title)} <span class="type-tag">${isMulti ? '多选' : '单选'}</span></h4>`;
      html += `<table><thead><tr><th>选项</th><th class="center">次数</th><th class="center">占比</th></tr></thead><tbody>`;
      html += `<tr class="base-row"><td>base</td><td class="center">${totalUsers}</td><td class="center">100%</td></tr>`;
      for (const opt of allOpts) {
        const val = counts[opt] || 0;
        const pct = totalUsers > 0 ? r3(val / totalUsers * 100) : 0;
        html += `<tr><td>${escapeHtml(opt)}</td><td class="center">${val}</td><td class="center">${pct}%</td></tr>`;
      }
      const meanVal = calcNumericMean(rows, colIdx);
      if (meanVal !== null) {
        html += `<tr class="mean-row"><td>平均值</td><td class="center" colspan="2">${r3(meanVal)}</td></tr>`;
      }
      html += `</tbody></table></div>`;
      previewHTML.push(html);
    } else {
      const groupCounts = {};
      const groupUsers = {};
      for (const g of orderedGroups) {
        groupCounts[g] = countOptions(grpData[g], colIdx, isMulti);
        groupUsers[g] = countUsers(grpData[g], colIdx);
      }
      const totalCounts = countOptions(rows, colIdx, isMulti);
      const totalAllUsers = countUsers(rows, colIdx);

      const allOptsSet = new Set();
      for (const counts of Object.values(groupCounts)) {
        for (const opt of Object.keys(counts)) allOptsSet.add(opt);
      }
      const allOptsCount = {};
      for (const counts of Object.values(groupCounts)) {
        for (const [opt, count] of Object.entries(counts)) {
          allOptsCount[opt] = (allOptsCount[opt] || 0) + count;
        }
      }
      const allOpts = [...allOptsSet].sort((a, b) => (allOptsCount[b] || 0) - (allOptsCount[a] || 0));

      let html = `<div class="result-item">`;
      html += `<h4>${escapeHtml(title)} <span class="type-tag">${isMulti ? '多选' : '单选'}</span></h4>`;
      html += `<div style="overflow-x:auto"><table><thead><tr><th>选项</th><th class="center">Total</th>`;
      for (const g of orderedGroups) html += `<th class="center">${escapeHtml(g)}</th>`;
      html += `</tr></thead><tbody>`;
      html += `<tr class="base-row"><td>base</td><td class="center">${totalAllUsers}</td>`;
      for (const g of orderedGroups) html += `<td class="center">${groupUsers[g] || 0}</td>`;
      html += `</tr>`;
      for (const opt of allOpts) {
        const totalVal = totalCounts[opt] || 0;
        const totalPct = totalAllUsers > 0 ? r3(totalVal / totalAllUsers * 100) : 0;
        html += `<tr><td>${escapeHtml(opt)}</td><td class="center">${totalPct}%</td>`;
        for (const g of orderedGroups) {
          const val = groupCounts[g][opt] || 0;
          const total = groupUsers[g] || 0;
          const pct = total > 0 ? r3(val / total * 100) : 0;
          html += `<td class="center">${pct}%</td>`;
        }
        html += `</tr>`;
      }
      html += `</tbody></table></div></div>`;
      previewHTML.push(html);
    }
  }

  document.getElementById('previewList').innerHTML = previewHTML.join('') || '<div class="empty-state">请先上传 Excel 文件并选择题目</div>';
}

// ==================== CONFIG ====================
function saveConfig() {
  const data = { groups, selectedQuestions, groupingColIdx, config };
  localStorage.setItem('survey_analysis_config', JSON.stringify(data));
}

function loadSavedConfig() {
  try {
    const saved = localStorage.getItem('survey_analysis_config');
    if (saved) {
      const data = JSON.parse(saved);
      if (data.groups) groups = data.groups;
      if (data.selectedQuestions) selectedQuestions = data.selectedQuestions;
      if (data.groupingColIdx !== undefined) groupingColIdx = data.groupingColIdx;
      if (data.config) config = data.config;
      render();
    }
  } catch (e) {
    console.error('Failed to load config:', e);
  }
}

function resetConfig() {
  groups = [];
  selectedQuestions = headers.map((_, i) => i);
  groupingColIdx = null;
  config = { groups: [], ordered: {}, mean_cols: {}, multi_idx: [] };
  localStorage.removeItem('survey_analysis_config');
  render();
}

function exportConfig() {
  const data = { groups, selectedQuestions, groupingColIdx, config };
  const blob = new Blob([JSON.stringify(data, null, 2)], { type: 'application/json' });
  const url = URL.createObjectURL(blob);
  const a = document.createElement('a');
  a.href = url;
  a.download = 'config.json';
  document.body.appendChild(a);
  a.click();
  document.body.removeChild(a);
  URL.revokeObjectURL(url);
}

function importConfig(e) {
  const file = e.target.files?.[0];
  if (!file) return;
  const reader = new FileReader();
  reader.onload = (evt) => {
    try {
      const data = JSON.parse(evt.target.result);
      if (data.groups) groups = data.groups;
      if (data.selectedQuestions) selectedQuestions = data.selectedQuestions;
      if (data.groupingColIdx !== undefined) groupingColIdx = data.groupingColIdx;
      if (data.config) config = data.config;
      saveConfig();
      render();
      alert('配置已加载');
    } catch (err) {
      alert('加载失败: ' + err.message);
    }
  };
  reader.readAsText(file);
  e.target.value = '';
}

// ==================== INIT ====================
document.addEventListener('DOMContentLoaded', function() {
  const uploadZone = document.getElementById('uploadZone');
  const fileInput = document.getElementById('fileInput');
  const configInput = document.getElementById('configInput');

  uploadZone.addEventListener('click', () => fileInput.click());

  uploadZone.addEventListener('dragover', (e) => {
    e.preventDefault();
    uploadZone.classList.add('dragging');
  });

  uploadZone.addEventListener('dragleave', () => {
    uploadZone.classList.remove('dragging');
  });

  uploadZone.addEventListener('drop', (e) => {
    e.preventDefault();
    uploadZone.classList.remove('dragging');
    const file = e.dataTransfer.files?.[0];
    if (file) handleFile(file);
  });

  fileInput.addEventListener('change', () => {
    const file = fileInput.files?.[0];
    if (file) handleFile(file);
  });

  configInput.addEventListener('change', importConfig);

  document.getElementById('exportConfigBtn').addEventListener('click', exportConfig);
  document.getElementById('resetConfigBtn').addEventListener('click', resetConfig);
  document.getElementById('selectAllBtn').addEventListener('click', selectAllQuestions);
  document.getElementById('selectNoneBtn').addEventListener('click', selectNoneQuestions);
  document.getElementById('downloadBtn').addEventListener('click', downloadExcel);
  document.getElementById('groupColSelect').addEventListener('change', handleGroupColChange);

  render();
  loadSavedConfig();
});