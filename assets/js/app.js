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

function getPrefix(header) {
  if (!header) return '';
  const cleaned = String(header).replace(/<[^>]*>/g, '').replace(/^[a-fA-F0-9]{6}">/, '').trim();
  const match = cleaned.match(/^([A-Za-z]?\d+[.\-]?\d*|[A-Za-z]+\d+|[一-龥]+)/);
  return match ? match[0] : cleaned.slice(0, 4);
}

function prefixSimilarity(a, b) {
  let i = 0;
  while (i < a.length && i < b.length && a[i] === b[i]) i++;
  return i / Math.max(a.length, b.length);
}

function buildSheets(questions) {
  // Always produce 1 sheet per question — sheet grouping is handled by the caller
  // by passing individual questions one at a time.
  if (!questions.length) return [];
  return questions.map(q => [q]);
}

function estimateRows(colIdx) {
  const isMulti = isMultiChoice(rows, colIdx);
  const opts = Object.keys(countOptions(rows, colIdx, isMulti));
  return 3 + opts.length + (calcNumericMean(rows, colIdx) !== null ? 1 : 0);
}

function partitionSheets(questions, maxSheets) {
  // Evenly distribute questions across maxSheets while preserving order.
  if (questions.length <= maxSheets) {
    return questions.map(q => [q]);
  }
  const perSheet = Math.ceil(questions.length / maxSheets);
  const sheets = [];
  for (let i = 0; i < maxSheets; i++) {
    const start = i * perSheet;
    if (start >= questions.length) break;
    sheets.push(questions.slice(start, start + perSheet));
  }
  return sheets;
}

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
  const cleaned = String(header).replace(/<[^>]*>/g, '').replace(/^[a-fA-F0-9]{6}">/, '').trim();
  const parts = cleaned.split(/[\s_#]+/);
  for (let i = parts.length - 1; i >= 0; i--) {
    if (parts[i].length >= 4) return parts[i];
  }
  return cleaned.slice(0, 20);
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
    // Always split by English comma (regardless of isMulti) if comma present
    if (String(val).includes(',')) {
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
  let total = 0, hasEnglishComma = 0;
  for (const row of allRows) {
    const val = row[colIdx];
    if (val === null || val === undefined) continue;
    total++;
    if (String(val).includes(',')) hasEnglishComma++;
  }
  return total > 0 && hasEnglishComma > total * 0.2;
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

// ==================== EXCEL WRITER (ExcelJS) ====================
const RED_FILL = 'FFFFCCCC';
const GREEN_FILL = 'FFCCFFCC';
const GRAY_FILL = 'FFD9D9D9';

function makeBorder() {
  return { top: { style: 'thin' }, bottom: { style: 'thin' }, left: { style: 'thin' }, right: { style: 'thin' } };
}

function makeHeaderStyle(align) {
  return { font: { name: '黑体', size: 10 }, fill: { type: 'pattern', pattern: 'solid', fgColor: { argb: GRAY_FILL } }, alignment: { horizontal: align }, border: makeBorder() };
}

function makeDataStyle(align, numFmt) {
  const s = { font: { name: '黑体', size: 11 }, alignment: { horizontal: align }, border: makeBorder() };
  if (numFmt) s.numFmt = numFmt;
  return s;
}

function makeFillStyle(color, align, numFmt) {
  const s = { font: { name: '黑体', size: 11 }, fill: { type: 'pattern', pattern: 'solid', fgColor: { argb: color } }, alignment: { horizontal: align }, border: makeBorder() };
  if (numFmt) s.numFmt = numFmt;
  return s;
}

function colNumToLetter(num) {
  let result = '';
  while (num > 0) {
    const rem = (num - 1) % 26;
    result = String.fromCharCode(65 + rem) + result;
    num = Math.floor((num - 1) / 26);
  }
  return result;
}

function generateExcelWorkbook() {
  const questions = selectedQuestions;
  if (!questions.length || !headers.length || !rows.length) return null;

  console.log('>>> generateExcelWorkbook START. groups length:', groups.length);
  const orderedGroups = groups.map(g => g.name);
  console.log('>>> orderedGroups:', orderedGroups);
  const numGroups = orderedGroups.length;
  const hasGroups = numGroups > 0;

  const table2Start = 2 + numGroups + 1;
  const table3Start = table2Start + 2 + numGroups + 1;

  const wb = new ExcelJS.Workbook();
  wb.creator = 'TableX';
  wb.created = new Date();

  // Build sheets preserving order, max 10 sheets
  const sheetGroups = partitionSheets(questions, 10);
  console.log('sheetGroups count:', sheetGroups.length, 'from', questions.length, 'questions');

  for (const groupIndices of sheetGroups) {
    const firstHeader = headers[groupIndices[0]] || `列${groupIndices[0]}`;
    const rawName = getPrefix(firstHeader).slice(0, 31) || '题目';
    let safeName = rawName.replace(/[\[\]:\\/?*]/g, '_').replace(/\|/g, '_');
    // Ensure unique sheet name
    let idx = 1;
    const usedSheets = wb.worksheets.map(ws => ws.name);
    while (usedSheets.includes(safeName)) {
      safeName = rawName.replace(/[\[\]:\\/?*]/g, '_').replace(/\|/g, '_').slice(0, 28) + '_' + idx;
      idx++;
    }
    console.log('sheet:', safeName, 'questions:', groupIndices.length);
    const ws = wb.addWorksheet(safeName);

    // Set column widths
    const colWidths = [];
    if (!hasGroups) {
      colWidths.push({ width: 1 }, { width: 18 }, { width: 10 }, { width: 1 }, { width: 18 }, { width: 10 }, { width: 1 }, { width: 18 }, { width: 10 });
    } else {
      // 3 tables each with (option + Total + numGroups groups), separated by spacers
      // Total cols = 3*(numGroups+2) + 2 spacers = 3*numGroups + 8
      const lastCol = 3 * numGroups + 8;
      colWidths.length = 0;
      for (let c = 0; c < lastCol; c++) {
        // Option col of each table: start of table1, start of table2, start of table3
        if (c === 0 || c === numGroups + 3 || c === 2 * (numGroups + 3)) colWidths.push({ width: 18 });
        // Spacer cols: after table1 and after table2
        else if (c === numGroups + 2 || c === 2 * (numGroups + 2) + 1) colWidths.push({ width: 1 });
        else colWidths.push({ width: 10 });
      }
    }
    console.log('generateExcelWorkbook called. hasGroups:', hasGroups, 'numGroups:', numGroups, 'groups:', groups.map(g => g.name));
    ws.columns = colWidths;

    let firstQuestion = true;
    for (const colIdx of groupIndices) {
      // Add spacer row between questions (but not before the first)
      if (!firstQuestion) {
        ws.addRow();
      }
      firstQuestion = false;
      const header = headers[colIdx] || `列${colIdx}`;
      const title = getShortTitle(header);
      const isMulti = isMultiChoice(rows, colIdx);

      // Title row: 题干 in row 1, 题目 in row 2 — each table gets its own cell
      const titleRow = ws.addRow();
      const subRow = ws.addRow();
      if (!hasGroups) {
        const t1 = { opt: 2, tot: 3 };
        const t2 = { opt: 6, tot: 7 };
        const t3 = { opt: 10, tot: 11 };
        const tables = [t1, t2, t3];
        // Row 1: 题干
        for (const t of tables) {
          const cell = titleRow.getCell(t.opt);
          cell.value = title;
          cell.font = { name: '黑体', size: 14 };
          cell.alignment = { horizontal: 'left' };
        }
        // Row 2: 题目 info
        for (const t of tables) {
          const cell = subRow.getCell(t.opt);
          cell.value = `题目: ${title} (${isMulti ? '多选' : '单选'})`;
          cell.font = { name: '黑体', size: 9 };
          cell.alignment = { horizontal: 'left' };
        }
      } else {
        for (let table = 0; table < 3; table++) {
          const tableStart = 1 + table * (numGroups + 3);
          const titleCell = titleRow.getCell(tableStart);
          titleCell.value = title;
          titleCell.font = { name: '黑体', size: 14 };
          titleCell.alignment = { horizontal: 'left' };

          const subCell = subRow.getCell(tableStart);
          subCell.value = `题目: ${title} (${isMulti ? '多选' : '单选'})`;
          subCell.font = { name: '黑体', size: 9 };
          subCell.alignment = { horizontal: 'left' };
        }
      }

      if (!hasGroups) {
        // No groups: 3 side-by-side tables with spacer cols at col 1, 5, 9 (0-indexed)
        // col layout: [spacer] | B(opt) C(Total) | [spacer] | F(opt) G(Total) | [spacer] | J(opt) K(Total)
        const t1 = { opt: 2, tot: 3 };
        const t2 = { opt: 6, tot: 7 };
        const t3 = { opt: 10, tot: 11 };

        const counts = countOptions(rows, colIdx, isMulti);
        const totalUsers = countUsers(rows, colIdx);
        const allOpts = Object.keys(counts).sort((a, b) => counts[b] - counts[a]);

        // Header row
        const hRow = ws.addRow();
        for (const t of [t1, t2, t3]) {
          hRow.getCell(t.opt).value = '选项'; hRow.getCell(t.opt).style = makeHeaderStyle('left');
          hRow.getCell(t.tot).value = 'Total'; hRow.getCell(t.tot).style = makeHeaderStyle('center');
        }

        // Base row (italic)
        const bRow = ws.addRow();
        for (const t of [t1, t2, t3]) {
          const c1 = bRow.getCell(t.opt); c1.value = 'base'; c1.style = makeDataStyle('left'); c1.font = { name: '黑体', size: 11, italic: true };
          const c2 = bRow.getCell(t.tot); c2.value = totalUsers; c2.style = makeDataStyle('center'); c2.font = { name: '黑体', size: 11, italic: true };
        }

        // Data rows
        const pctByOpt = allOpts.map(opt => {
          const val = counts[opt] || 0;
          return totalUsers > 0 ? r3(val / totalUsers) : 0;
        });
        const maxPct = Math.max(...pctByOpt);
        const minPct = Math.min(...pctByOpt);

        for (let oi = 0; oi < allOpts.length; oi++) {
          const opt = allOpts[oi];
          const val = counts[opt] || 0;
          const pct = pctByOpt[oi];

          const dRow = ws.addRow();
          // Table 1: counts
          dRow.getCell(t1.opt).value = opt; dRow.getCell(t1.opt).style = makeDataStyle('left');
          dRow.getCell(t1.tot).value = val; dRow.getCell(t1.tot).style = makeDataStyle('center');

          // Table 2: pct (no highlight)
          dRow.getCell(t2.opt).value = opt; dRow.getCell(t2.opt).style = makeDataStyle('left');
          dRow.getCell(t2.tot).value = pct; dRow.getCell(t2.tot).style = makeDataStyle('center', '0.0%');

          // Table 3: pct with highlight (red=highest, green=lowest, only if distinct)
          dRow.getCell(t3.opt).value = opt; dRow.getCell(t3.opt).style = makeDataStyle('left');
          if (maxPct !== minPct) {
            if (pct === maxPct) dRow.getCell(t3.tot).style = makeFillStyle(GREEN_FILL, 'center', '0.0%');
            else if (pct === minPct) dRow.getCell(t3.tot).style = makeFillStyle(RED_FILL, 'center', '0.0%');
            else dRow.getCell(t3.tot).style = makeDataStyle('center', '0.0%');
          } else {
            dRow.getCell(t3.tot).style = makeDataStyle('center', '0.0%');
          }
          dRow.getCell(t3.tot).value = pct;
        }

        // No mean row

      } else {
        // With groups: 3 tables of (选项, Total, groups...)
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

        // Header row
        const hRow = ws.addRow();
        let cellIdx = 1;
        // Table 1: opt, Total, groups
        hRow.getCell(cellIdx++).value = '选项'; hRow.getCell(cellIdx - 1).style = makeHeaderStyle('left');
        hRow.getCell(cellIdx++).value = 'Total'; hRow.getCell(cellIdx - 1).style = makeHeaderStyle('center');
        for (let i = 0; i < numGroups; i++) { hRow.getCell(cellIdx++).value = orderedGroups[i]; hRow.getCell(cellIdx - 1).style = makeHeaderStyle('center'); }
        // Spacer after table 1
        cellIdx++;
        // Table 2: opt, Total, groups
        hRow.getCell(cellIdx++).value = '选项'; hRow.getCell(cellIdx - 1).style = makeHeaderStyle('left');
        hRow.getCell(cellIdx++).value = 'Total'; hRow.getCell(cellIdx - 1).style = makeHeaderStyle('center');
        for (let i = 0; i < numGroups; i++) { hRow.getCell(cellIdx++).value = orderedGroups[i]; hRow.getCell(cellIdx - 1).style = makeHeaderStyle('center'); }
        // Spacer after table 2
        cellIdx++;
        // Table 3: opt, Total, groups
        hRow.getCell(cellIdx++).value = '选项'; hRow.getCell(cellIdx - 1).style = makeHeaderStyle('left');
        hRow.getCell(cellIdx++).value = 'Total'; hRow.getCell(cellIdx - 1).style = makeHeaderStyle('center');
        for (let i = 0; i < numGroups; i++) { hRow.getCell(cellIdx++).value = orderedGroups[i]; hRow.getCell(cellIdx - 1).style = makeHeaderStyle('center'); }

        // Base row
        const bRow = ws.addRow();
        cellIdx = 1;
        const applyBaseStyle = (cell, value) => { cell.value = value; cell.style = makeDataStyle('center'); cell.font = { name: '黑体', size: 11, italic: true }; };
        const applyBaseStyleLeft = (cell, value) => { cell.value = value; cell.style = makeDataStyle('left'); cell.font = { name: '黑体', size: 11, italic: true }; };
        applyBaseStyleLeft(bRow.getCell(cellIdx++), 'base');
        applyBaseStyle(bRow.getCell(cellIdx++), totalAllUsers);
        for (let i = 0; i < numGroups; i++) { applyBaseStyle(bRow.getCell(cellIdx++), groupUsers[orderedGroups[i]] || 0); }
        cellIdx++; // spacer
        for (let t = 0; t < 2; t++) {
          applyBaseStyleLeft(bRow.getCell(cellIdx++), 'base');
          applyBaseStyle(bRow.getCell(cellIdx++), totalAllUsers);
          for (let i = 0; i < numGroups; i++) { applyBaseStyle(bRow.getCell(cellIdx++), groupUsers[orderedGroups[i]] || 0); }
          if (t < 1) cellIdx++; // spacer between tables
        }

        // Data rows
        for (const opt of allOpts) {
          const pcts = {};
          for (const g of orderedGroups) {
            const val = groupCounts[g][opt] || 0;
            const total = groupUsers[g] || 0;
            pcts[g] = total > 0 ? r3(val / total) : 0;
          }
          const totalVal = totalCounts[opt] || 0;
          const totalPct = totalAllUsers > 0 ? r3(totalVal / totalAllUsers) : 0;

          const maxPct = Math.max(...Object.values(pcts), totalPct);
          const minPct = Math.min(...Object.values(pcts), totalPct);

          const dRow = ws.addRow();
          cellIdx = 1;
          // Table 1: opt, totalVal, groups (counts)
          dRow.getCell(cellIdx++).value = opt; dRow.getCell(cellIdx - 1).style = makeDataStyle('left');
          dRow.getCell(cellIdx++).value = totalVal; dRow.getCell(cellIdx - 1).style = makeDataStyle('center');
          for (let i = 0; i < numGroups; i++) { dRow.getCell(cellIdx++).value = groupCounts[orderedGroups[i]][opt] || 0; dRow.getCell(cellIdx - 1).style = makeDataStyle('center'); }
          // Spacer
          cellIdx++;
          // Table 2: opt, totalPct, groups (pct, no highlight)
          dRow.getCell(cellIdx++).value = opt; dRow.getCell(cellIdx - 1).style = makeDataStyle('left');
          dRow.getCell(cellIdx++).value = totalPct; dRow.getCell(cellIdx - 1).style = makeDataStyle('center', '0.0%');
          for (let i = 0; i < numGroups; i++) { dRow.getCell(cellIdx++).value = pcts[orderedGroups[i]]; dRow.getCell(cellIdx - 1).style = makeDataStyle('center', '0.0%'); }
          // Spacer
          cellIdx++;
          // Table 3: opt, totalPct, groups (pct with highlight)
          dRow.getCell(cellIdx++).value = opt; dRow.getCell(cellIdx - 1).style = makeDataStyle('left');
          dRow.getCell(cellIdx++).value = totalPct; dRow.getCell(cellIdx - 1).style = makeDataStyle('center', '0.0%');
          for (let i = 0; i < numGroups; i++) {
            const g = orderedGroups[i];
            const cellPct = pcts[g];
            if (maxPct !== minPct) {
              if (cellPct === maxPct) dRow.getCell(cellIdx++).style = makeFillStyle(GREEN_FILL, 'center', '0.0%');
              else if (cellPct === minPct) dRow.getCell(cellIdx++).style = makeFillStyle(RED_FILL, 'center', '0.0%');
              else dRow.getCell(cellIdx++).style = makeDataStyle('center', '0.0%');
            } else {
              dRow.getCell(cellIdx++).style = makeDataStyle('center', '0.0%');
            }
            dRow.getCell(cellIdx - 1).value = cellPct;
          }
        }

        // No mean row
      }
    }
  }

  return wb;
}

async function downloadExcel() {
  try {
    console.log('questions selected:', selectedQuestions.length, 'headers:', headers.length);
    const wb = generateExcelWorkbook();
    if (!wb) {
      alert('请先上传 Excel 文件并选择题目');
      return;
    }
    const buffer = await wb.xlsx.writeBuffer();
    const blob = new Blob([buffer], { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' });
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

// ==================== UI HANDLERS ====================
async function handleFile(file) {
  if (!file) return;
  try {
    // Reset all state before loading new file
    headers = [];
    rows = [];
    selectedQuestions = [];
    groupingColIdx = null;
    groups = [];
    config = { groups: [], ordered: {}, mean_cols: {}, multi_idx: [] };
    localStorage.removeItem('survey_analysis_config');

    const result = await parseExcel(file);
    headers = result.headers;
    rows = result.rows;
    selectedQuestions = headers.map((_, i) => i);

    // Update upload zone with file info
    const uploadHint = document.getElementById('uploadHint');
    if (uploadHint) {
      uploadHint.textContent = `${file.name} (${rows.length} 行 × ${headers.length} 列)`;
    }

    render();
    // Only restore saved config if the saved grouping col still exists in the new file
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
let dragIdx = null;

function handleGroupDragStart(e, idx) {
  dragIdx = idx;
  e.dataTransfer.effectAllowed = 'move';
  e.currentTarget.classList.add('dragging-item');
}

function handleGroupDragOver(e, idx) {
  e.preventDefault();
  e.dataTransfer.dropEffect = 'move';
  // Remove drag-over class from all items
  document.querySelectorAll('.drag-over-target').forEach(el => el.classList.remove('drag-over-target'));
  // Add drag-over class to hovered item
  e.currentTarget.classList.add('drag-over-target');
  if (dragIdx === null || dragIdx === idx) return;
  const [removed] = groups.splice(dragIdx, 1);
  groups.splice(idx, 0, removed);
  dragIdx = idx;
  renderSelectedGroups();
}

function handleGroupDragEnd(e) {
  e.currentTarget.classList.remove('dragging-item');
  document.querySelectorAll('.drag-over-target').forEach(el => el.classList.remove('drag-over-target'));
  dragIdx = null;
  saveConfig();
}

// ==================== RENDER ====================
function hide(el) { if (el) el.style.display = 'none'; }
function show(el) { if (el) el.style.display = 'block'; }

function render() {
  const uploadZone = document.getElementById('uploadZone');
  uploadZone.classList.remove('dragging');

  const groupCard = document.getElementById('groupSelectorCard');
  const groupColSelect = document.getElementById('groupColSelect');
  const groupValuesContainer = document.getElementById('groupValuesContainer');
  const selectedGroupsContainer = document.getElementById('selectedGroupsContainer');
  const qCard = document.getElementById('questionSelectorCard');
  if (headers.length === 0) {
    hide(groupCard);
    hide(document.getElementById('questionSelectorCard'));
    return;
  }

  show(groupCard);

  const colOptions = headers.map((h, i) => ({ idx: i, name: getShortTitle(h) || `列${i + 1}` })).filter(h => h.name);
  groupColSelect.innerHTML = '<option value="">-- 选择分组列 --</option>' +
    colOptions.map(opt => `<option value="${opt.idx}" ${groupingColIdx === opt.idx ? 'selected' : ''}>${opt.name}</option>`).join('');

  if (groupingColIdx === null) {
    hide(groupValuesContainer);
    hide(selectedGroupsContainer);
  } else {
    show(groupValuesContainer);

    const values = new Set();
    const valueCounts = {};
    for (const row of rows) {
      const v = safe(row[groupingColIdx]);
      if (v !== null) {
        values.add(v);
        valueCounts[v] = (valueCounts[v] || 0) + 1;
      }
    }
    const sortedValues = [...values].sort();
    const selectedValues = groups.filter(g => g.conditions[0][0] === groupingColIdx).map(g => g.name);

    document.getElementById('groupValues').innerHTML = sortedValues.map(val => `
      <button class="group-tag ${selectedValues.includes(val) ? 'selected' : 'unselected'}"
              onclick="toggleGroupValue('${val.replace(/'/g, "\\'")}')">${val} <span style="opacity:0.6;font-size:11px">(${valueCounts[val]})</span></button>
    `).join('');

    if (groups.length > 0) {
      show(selectedGroupsContainer);
      renderSelectedGroups();
    } else {
      hide(selectedGroupsContainer);
    }
  }

  if (headers.length === 0) {
    hide(qCard);
    return;
  }
  show(qCard);

  // Build flat list preserving original order, inserting module labels when topic changes
  const questionItems = [];
  let lastMod = null;
  for (let i = 0; i < headers.length; i++) {
    const mod = detectModule(headers[i]);
    const title = getShortTitle(headers[i]) || `列${i + 1}`;
    const isMulti = rows.length > 0 && isMultiChoice(rows, i);
    if (mod !== lastMod) {
      questionItems.push({ type: 'module', name: mod });
      lastMod = mod;
    }
    questionItems.push({ type: 'question', idx: i, title, isMulti });
  }

  document.getElementById('questionCount').textContent = `已选择 ${selectedQuestions.length} / ${headers.length} 题`;

  document.getElementById('questionList').innerHTML = questionItems.map((item, i) => {
    if (item.type === 'module') {
      const prev = questionItems[i - 1];
      const close = prev && prev.type === 'question' ? '</div>' : '';
      return `${close}<div class="question-module"><h4>${item.name}</h4>`;
    }
    return `
        <label class="question-item">
          <input type="checkbox" ${selectedQuestions.includes(item.idx) ? 'checked' : ''} onchange="toggleQuestion(${item.idx})">
          <span>${item.title}</span>
          <span class="type-tag">${item.isMulti ? '多选' : '单选'}</span>
        </label>`;
  }).join('') + '</div>';

  document.getElementById('downloadBtn').disabled = selectedQuestions.length === 0;
}

function renderSelectedGroups() {
  const container = document.getElementById('selectedGroups');
  if (!container) return;
  container.innerHTML = groups.map((g, i) => `
    <div class="group-tag selected" draggable="true"
         ondragstart="handleGroupDragStart(event, ${i})"
         ondragover="handleGroupDragOver(event, ${i})"
         ondragend="handleGroupDragEnd(event)">
      <span class="drag-handle">⋮⋮</span>
      <span>${g.name}</span>
    </div>
  `).join('');
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
      // Only restore grouping if the column index is still valid in the new file
      if (data.groupingColIdx !== null && data.groupingColIdx !== undefined && data.groupingColIdx < headers.length) {
        groupingColIdx = data.groupingColIdx;
      } else {
        groupingColIdx = null;
      }
      // Only restore groups if groupingColIdx is valid
      if (groupingColIdx !== null && data.groups && Array.isArray(data.groups)) {
        groups = data.groups;
      } else {
        groups = [];
      }
      if (data.selectedQuestions) selectedQuestions = data.selectedQuestions;
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