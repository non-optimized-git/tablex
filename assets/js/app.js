// ==================== STATE ====================
let headers = [];
let rows = [];
let selectedQuestions = [];
let groupingColIdx = null;
let groups = [];
let config = { groups: [], ordered: {}, mean_cols: {}, multi_idx: [] };
// modules: [{id: number, name: string, questionIdxs: number[]}]
let modules = [];
let nextModuleId = 1;
let uploadedFileName = ''; // 上传的文件名（不含扩展名）
// 表头编辑器状态
let editingGroupIndex = -1; // -1 表示新建，>=0 表示编辑
let editingRelation = 'and'; // 条件组之间的关系：'and' | 'or'
let editingClauses = []; // [{relation: 'and'|'or', conditions: [{colIdx, values}]}]，空数组=编辑器关闭

// 拖拽选择状态
let isDraggingSelect = false;
let dragSelectValue = false; // true=选中, false=取消

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

function cleanHeader(header) {
  if (!header) return '';
  return String(header).replace(/<[^>]*>/g, '').replace(/^[a-fA-F0-9]{2,6}[\s_#:]+/, '').trim();
}

function getShortTitle(header) {
  if (!header) return '';
  let s = String(header).replace(/<[^>]*>/g, '').replace(/^[a-fA-F0-9]{6}">/, '').trim();
  // Remove type tag
  s = s.replace(/[（(]\s*(单选|多选|量表|矩阵|排序|填空|开放)\s*[）)]/g, '');
  // Remove leading question phrases
  s = s.replace(/^(请问\s*你(目前|目前对|是否|进店时|进店后|在进店看过)?\s*)/, '');
  s = s.replace(/^(请问\s*新乐道L60的?\s*)/, '');
  s = s.replace(/^(你的?|您)\s*/, '');
  s = s.replace(/^(请问\s*)/, '');
  // Remove brand name
  s = s.replace(/新乐道L60/gi, '');
  // Remove trailing question phrases
  s = s.replace(/\s*哪些\s*[?？].*$/, ''); // 哪些...？ → 哪些
  s = s.replace(/\s*什么\s*[?？].*$/, ''); // 什么...？ → 什么
  s = s.replace(/\s*了吗\s*[?？]?.*$/, ''); // ...了吗？ → (empty)
  s = s.replace(/[?？].*$/, ''); // remaining question mark + after
  s = s.replace(/\s*让你比较犹豫\s*$/, '');
  s = s.replace(/\s*让你比较不满意\s*$/, '');
  s = s.replace(/\s*有没有\s*[?？]?.*$/, ''); // 有没有...？
  s = s.replace(/了\s*$/, '');
  s = s.replace(/[是的了]\s*$/, '');
  s = s.replace(/^[a-fA-F0-9]{2,6}[\s_#:]+/, '');
  s = s.replace(/\s+/g, ' ').trim();
  if (!s) s = String(header).slice(0, 20);
  return s.length > 16 ? s.slice(0, 16) : s;
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

// ==================== 表头条件 ====================
// 表头定义：{ name, relation: 'and'|'or', clauses: [{ relation, conditions: [{colIdx, values}] }] }
// 语义：条件值内 OR（命中任一值即可），条件间按组内 relation，条件组间按顶层 relation，
// 可表达 (A & B) | (C & D) 等分组括号组合。migrateGroup 兼容历史格式
// conditions: [[colIdx, val]] / [{colIdx, values}]。
function migrateGroup(g) {
  if (Array.isArray(g.clauses)) {
    return {
      name: g.name,
      relation: g.relation === 'or' ? 'or' : 'and',
      clauses: g.clauses.map(cl => ({
        relation: cl.relation === 'or' ? 'or' : 'and',
        conditions: (Array.isArray(cl.conditions) ? cl.conditions : []).map(c => ({
          colIdx: c.colIdx,
          values: Array.isArray(c.values) ? c.values : [],
          valuesRelation: c.valuesRelation === 'and' ? 'and' : 'or'
        }))
      }))
    };
  }
  const conds = Array.isArray(g.conditions) ? g.conditions.map(c =>
    Array.isArray(c)
      ? { colIdx: c[0], values: [c[1]], valuesRelation: 'or' }
      : { colIdx: c.colIdx, values: Array.isArray(c.values) ? c.values : [], valuesRelation: c.valuesRelation === 'and' ? 'and' : 'or' }
  ) : [];
  return { name: g.name, relation: 'and', clauses: [{ relation: 'and', conditions: conds }] };
}

// 基础表头的单一条件（单条件组、单条件），否则返回 null
function getBasicCondition(g) {
  const m = migrateGroup(g);
  if (m.clauses.length !== 1 || m.clauses[0].conditions.length !== 1) return null;
  return m.clauses[0].conditions[0];
}

function condMatches(row, cond) {
  // 选项之间：任一（或，默认）/ 全部（且，多选题）；多选单元格按英文逗号分割后逐项匹配
  // （与选项数据同口径），整值命中保留以兼容历史配置
  const v = safe(row[cond.colIdx]);
  if (v === null) return false;
  const parts = splitOpts(v);
  const hit = val => v === val || parts.includes(val);
  return cond.valuesRelation === 'and' ? cond.values.every(hit) : cond.values.some(hit);
}

function clauseMatches(row, cl) {
  return cl.relation === 'or'
    ? cl.conditions.some(c => condMatches(row, c))
    : cl.conditions.every(c => condMatches(row, c));
}

function groupRowsCustom(allRows, customGroups) {
  const result = {};
  for (const groupDef of customGroups) {
    const g = migrateGroup(groupDef);
    result[g.name] = allRows.filter(row =>
      g.relation === 'or'
        ? g.clauses.some(cl => clauseMatches(row, cl))
        : g.clauses.every(cl => clauseMatches(row, cl))
    );
  }
  return result;
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

function r3(x) {
  if (x === null || x === undefined) return null;
  return x < 0.0005 ? 0 : Math.round(x * 1000) / 1000;
}

// 从选项值中提取数值分数，如 "5.0-非常满意" → 5.0，"4.5" → 4.5
function extractScoreValue(val) {
  if (val === null || val === undefined) return null;
  const m = String(val).match(/^([\d.]+)/);
  if (!m) return null;
  const n = parseFloat(m[1]);
  return isNaN(n) ? null : n;
}

// 判断是否为量表题：所有选项都能提取出数值分数
function isScoreQuestion(allRows, colIdx) {
  const counts = countOptions(allRows, colIdx, false);
  const opts = Object.keys(counts);
  if (opts.length === 0) return false;
  return opts.every(opt => extractScoreValue(opt) !== null);
}

// 计算加权平均分：Σ(分数 × 人数) / 总人数
function calcWeightedMean(allRows, colIdx) {
  let totalWeight = 0;
  let totalCount = 0;
  for (const row of allRows) {
    const val = safe(row[colIdx]);
    if (val === null) continue;
    const parts = String(val).includes(',') ? splitOpts(val) : [val];
    for (const p of parts) {
      const score = extractScoreValue(p);
      if (score !== null) {
        totalWeight += score;
        totalCount++;
      }
    }
  }
  return totalCount > 0 ? totalWeight / totalCount : null;
}

// ==================== EXCEL WRITER (ExcelJS) ====================
const RED_FILL = 'FFD96459';
const GREEN_FILL = 'FF6AA84F';
const GRAY_FILL = 'FFDEE0E3';
const BASE_FILL = 'FFA89888';
const AVG_FILL = 'FFFFF3E0'; // 暖琥珀色，用于平均分行

const BORDER = { top: { style: 'thin' }, bottom: { style: 'thin' }, left: { style: 'thin' }, right: { style: 'thin' } };
const FONT_DATA = { name: '微软雅黑', size: 10 };
const FONT_HEADER = { name: '微软雅黑', size: 10 };

function makeBorder() { return BORDER; }

function makeHeaderStyle(align) {
  return { font: FONT_HEADER, fill: { type: 'pattern', pattern: 'solid', fgColor: { argb: GRAY_FILL } }, alignment: { horizontal: align }, border: BORDER };
}

function makeDataStyle(align, numFmt) {
  const s = { font: FONT_DATA, alignment: { horizontal: align }, border: BORDER };
  if (numFmt) s.numFmt = numFmt;
  return s;
}

function makeFillStyle(color, align, numFmt) {
  const s = { font: FONT_DATA, fill: { type: 'pattern', pattern: 'solid', fgColor: { argb: color } }, alignment: { horizontal: align }, border: BORDER };
  if (numFmt) s.numFmt = numFmt;
  return s;
}

// Table 3 红→绿色阶：t=0 → RED_FILL（最低），t=1 → GREEN_FILL（最高），
// 中间按数据高低在两端点色之间线性插值。端点色为全局常量，全表统一。
function mixFill(t) {
  const parse = argb => [parseInt(argb.slice(2, 4), 16), parseInt(argb.slice(4, 6), 16), parseInt(argb.slice(6, 8), 16)];
  const c1 = parse(RED_FILL);
  const c2 = parse(GREEN_FILL);
  const hex = n => Math.max(0, Math.min(255, Math.round(n))).toString(16).padStart(2, '0').toUpperCase();
  return 'FF' + [0, 1, 2].map(i => hex(c1[i] + (c2[i] - c1[i]) * t)).join('');
}

function generateExcelWorkbook() {
  const questions = selectedQuestions;
  if (!questions.length || !headers.length || !rows.length) return null;

  const orderedGroups = groups.map(g => g.name);
  const numGroups = orderedGroups.length;
  const hasGroups = numGroups > 0;

  const table2Start = 2 + numGroups + 1;
  const table3Start = table2Start + 2 + numGroups + 1;

  const wb = new ExcelJS.Workbook();
  wb.creator = 'TableX';
  wb.created = new Date();

  // Build sheets per module
  for (const mod of modules) {
    const modQuestions = mod.questionIdxs.filter(idx => selectedQuestions.includes(idx));
    if (modQuestions.length === 0) continue;

    const rawName = mod.name || '题目';
    let safeName = rawName.replace(/[\[\]:\\/?*]/g, '_').replace(/\|/g, '_').slice(0, 31) || '题目';
    // Ensure unique sheet name
    let idx = 1;
    const usedSheets = wb.worksheets.map(ws => ws.name);
    while (usedSheets.includes(safeName)) {
      safeName = rawName.replace(/[\[\]:\\/?*]/g, '_').replace(/\|/g, '_').slice(0, 28) + '_' + idx;
      idx++;
    }
    const ws = wb.addWorksheet(safeName);

    // Set column widths
    const colWidths = [];
    if (!hasGroups) {
      // col layout: [spacer] | opt | Total | [spacer] | opt | Total | [spacer] | opt | Total
      // col 1: spacer, col 2: opt, col 3: Total, col 4: spacer, col 5: spacer, col 6: opt, col 7: Total, col 8: spacer, col 9: spacer, col 10: opt, col 11: Total
      colWidths.push(
        { width: 1 }, { width: 18 }, { width: 10 },  // Table 1
        { width: 1 }, { width: 1 },                    // Spacer
        { width: 18 }, { width: 10 },                  // Table 2
        { width: 1 }, { width: 1 },                    // Spacer
        { width: 18 }, { width: 10 }                   // Table 3
      );
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
    ws.columns = colWidths;

    let firstQuestion = true;
    for (const colIdx of modQuestions) {
      // Add spacer row between questions (but not before the first)
      if (!firstQuestion) {
        ws.addRow();
      }
      firstQuestion = false;
      const header = headers[colIdx] || `列${colIdx}`;
      const title = getShortTitle(header);
      const fullTitle = cleanHeader(header);
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
          cell.font = { name: '微软雅黑', size: 14, bold: true };
          cell.alignment = { horizontal: 'left' };
        }
        // Row 2: 题目 info
        for (const t of tables) {
          const cell = subRow.getCell(t.opt);
          cell.value = `题目: ${fullTitle} (${isMulti ? '多选' : '单选'})`;
          cell.font = { name: '微软雅黑', size: 10 };
          cell.alignment = { horizontal: 'left' };
        }
      } else {
        for (let table = 0; table < 3; table++) {
          const tableStart = 1 + table * (numGroups + 3);
          const titleCell = titleRow.getCell(tableStart);
          titleCell.value = title;
          titleCell.font = { name: '微软雅黑', size: 14, bold: true };
          titleCell.alignment = { horizontal: 'left' };

          const subCell = subRow.getCell(tableStart);
          subCell.value = `题目: ${fullTitle} (${isMulti ? '多选' : '单选'})`;
          subCell.font = { name: '微软雅黑', size: 10 };
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
          const c1 = bRow.getCell(t.opt); c1.value = 'base'; c1.style = makeDataStyle('left'); c1.font = { name: '微软雅黑', size: 10, italic: true, color: { argb: BASE_FILL } };
          const c2 = bRow.getCell(t.tot); c2.value = totalUsers; c2.style = makeDataStyle('center'); c2.font = { name: '微软雅黑', size: 10, italic: true, color: { argb: BASE_FILL } };
        }

        // 平均分行（仅量表题）
        if (isScoreQuestion(rows, colIdx)) {
          const avgMean = calcWeightedMean(rows, colIdx);
          const avgRow = ws.addRow();
          for (const t of [t1, t2, t3]) {
            const c1 = avgRow.getCell(t.opt); c1.value = '平均分'; c1.style = makeFillStyle(AVG_FILL, 'left'); c1.font = { name: '微软雅黑', size: 10, bold: true };
            const c2 = avgRow.getCell(t.tot); c2.value = avgMean !== null ? Math.round(avgMean * 100) / 100 : ''; c2.style = makeFillStyle(AVG_FILL, 'center'); c2.font = { name: '微软雅黑', size: 10, bold: true }; c2.numFmt = '0.00';
          }
        }

        // Data rows
        const pctByOpt = allOpts.map(opt => {
          const val = counts[opt] || 0;
          return totalUsers > 0 ? r3(val / totalUsers) : 0;
        });

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

          // Table 3: pct（Total 列不上色）
          dRow.getCell(t3.opt).value = opt; dRow.getCell(t3.opt).style = makeDataStyle('left');
          dRow.getCell(t3.tot).value = pct; dRow.getCell(t3.tot).style = makeDataStyle('center', '0.0%');
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

        // 收集所有选项（从 total 中获取，确保不遗漏）
        const allOpts = Object.keys(totalCounts).sort((a, b) => totalCounts[b] - totalCounts[a]);

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
        const applyBaseStyle = (cell, value) => { cell.value = value; cell.style = makeDataStyle('center'); cell.font = { name: '微软雅黑', size: 10, italic: true, color: { argb: BASE_FILL } }; };
        const applyBaseStyleLeft = (cell, value) => { cell.value = value; cell.style = makeDataStyle('left'); cell.font = { name: '微软雅黑', size: 10, italic: true, color: { argb: BASE_FILL } }; };
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

        // 平均分行（仅量表题）
        if (isScoreQuestion(rows, colIdx)) {
          const avgRow = ws.addRow();
          cellIdx = 1;
          const avgStyle = (cell, value) => { cell.value = value; cell.style = makeFillStyle(AVG_FILL, 'center'); cell.font = { name: '微软雅黑', size: 10, bold: true }; cell.numFmt = '0.00'; };
          const avgStyleLeft = (cell, value) => { cell.value = value; cell.style = makeFillStyle(AVG_FILL, 'left'); cell.font = { name: '微软雅黑', size: 10, bold: true }; };
          const totalAvg = calcWeightedMean(rows, colIdx);
          const totalAvgR = totalAvg !== null ? Math.round(totalAvg * 100) / 100 : '';

          // Table 1
          avgStyleLeft(avgRow.getCell(cellIdx++), '平均分');
          avgStyle(avgRow.getCell(cellIdx++), totalAvgR);
          for (let i = 0; i < numGroups; i++) {
            const gAvg = calcWeightedMean(grpData[orderedGroups[i]], colIdx);
            avgStyle(avgRow.getCell(cellIdx++), gAvg !== null ? Math.round(gAvg * 100) / 100 : '');
          }
          cellIdx++; // spacer
          // Table 2 & 3
          for (let t = 0; t < 2; t++) {
            avgStyleLeft(avgRow.getCell(cellIdx++), '平均分');
            avgStyle(avgRow.getCell(cellIdx++), totalAvgR);
            for (let i = 0; i < numGroups; i++) {
              const gAvg = calcWeightedMean(grpData[orderedGroups[i]], colIdx);
              avgStyle(avgRow.getCell(cellIdx++), gAvg !== null ? Math.round(gAvg * 100) / 100 : '');
            }
            if (t < 1) cellIdx++; // spacer between tables
          }
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

          // 色阶只在分组列之间归一化（Total 列不上色，也不参与极值）
          const maxPct = Math.max(...Object.values(pcts));
          const minPct = Math.min(...Object.values(pcts));

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
          // Table 3: opt, totalPct, groups (pct with red→green gradient per row; Total 列不上色)
          const gradStyle = pctVal => maxPct !== minPct
            ? makeFillStyle(mixFill((pctVal - minPct) / (maxPct - minPct)), 'center', '0.0%')
            : makeDataStyle('center', '0.0%');
          dRow.getCell(cellIdx++).value = opt; dRow.getCell(cellIdx - 1).style = makeDataStyle('left');
          const t3TotCell = dRow.getCell(cellIdx++); t3TotCell.value = totalPct; t3TotCell.style = makeDataStyle('center', '0.0%');
          for (let i = 0; i < numGroups; i++) {
            const cellPct = pcts[orderedGroups[i]];
            const gCell = dRow.getCell(cellIdx++); gCell.value = cellPct; gCell.style = gradStyle(cellPct);
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
    // 文件名：上传文件名_MMddHHmm.xlsx
    const now = new Date();
    const timeStr = String(now.getMonth() + 1).padStart(2, '0') +
                    String(now.getDate()).padStart(2, '0') +
                    String(now.getHours()).padStart(2, '0') +
                    String(now.getMinutes()).padStart(2, '0');
    a.download = uploadedFileName ? `${uploadedFileName}_${timeStr}.xlsx` : `${timeStr}.xlsx`;
    document.body.appendChild(a);
    a.click();
    document.body.removeChild(a);
    URL.revokeObjectURL(url);
  } catch (err) {
    console.error('Excel export error:', err);
    alert('导出失败: ' + err.message);
  }
}

// ==================== MODULE MANAGEMENT ====================
function initModules() {
  // 所有题目放在一个默认模块中
  modules = [];
  if (selectedQuestions.length > 0) {
    modules.push({ id: nextModuleId++, name: '题目', questionIdxs: [...selectedQuestions] });
  }
}

function addModuleAfterQuestion(beforeIdx) {
  // Find which module contains this question
  const currentModule = modules.find(m => m.questionIdxs.includes(beforeIdx));
  if (!currentModule) return;

  // Find position of this question in the module
  const pos = currentModule.questionIdxs.indexOf(beforeIdx);

  // Create new module with questions after this one
  const newModule = {
    id: nextModuleId++,
    name: `模块 ${modules.length + 1}`,
    questionIdxs: currentModule.questionIdxs.slice(pos + 1)
  };

  // Keep only questions up to and including this one in current module
  currentModule.questionIdxs = currentModule.questionIdxs.slice(0, pos + 1);

  // Insert new module after current module
  const moduleIdx = modules.indexOf(currentModule);
  modules.splice(moduleIdx + 1, 0, newModule);

  saveConfig();
  render();
}

function moveQuestionToModule(questionIdx, moduleId) {
  // Remove from all modules
  for (const m of modules) {
    m.questionIdxs = m.questionIdxs.filter(i => i !== questionIdx);
  }
  // Add to target module
  const target = modules.find(m => m.id === moduleId);
  if (target) {
    target.questionIdxs.push(questionIdx);
  }
  saveConfig();
}

function renameModule(moduleId, newName) {
  const m = modules.find(m => m.id === moduleId);
  if (m) {
    m.name = newName;
    saveConfig();
  }
}

function deleteModule(moduleId) {
  if (modules.length <= 1) return; // Keep at least one module
  const m = modules.find(m => m.id === moduleId);
  if (!m) return;
  // Move all questions to the previous module (or next if first)
  const idx = modules.indexOf(m);
  const targetModule = modules[idx - 1] || modules[idx + 1];
  if (targetModule) {
    // Insert questions into target module, maintaining original order
    for (const qIdx of m.questionIdxs) {
      // Find the correct insertion point to maintain original order
      const originalPos = selectedQuestions.indexOf(qIdx);
      let insertIdx = 0;
      for (let i = 0; i < targetModule.questionIdxs.length; i++) {
        if (selectedQuestions.indexOf(targetModule.questionIdxs[i]) < originalPos) {
          insertIdx = i + 1;
        }
      }
      targetModule.questionIdxs.splice(insertIdx, 0, qIdx);
    }
  }
  modules = modules.filter(mod => mod.id !== moduleId);
  saveConfig();
  render();
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
    editingGroupIndex = -1;
    editingRelation = 'and';
    editingClauses = [];
    localStorage.removeItem('survey_analysis_config');

    const result = await parseExcel(file);
    headers = result.headers;
    rows = result.rows;
    selectedQuestions = headers.map((_, i) => i);

    // 保存上传的文件名（不含扩展名）
    uploadedFileName = file.name.replace(/\.[^/.]+$/, '');

    // Initialize modules after selectedQuestions is populated
    initModules();

    // Update upload zone with file info
    const uploadHint = document.getElementById('uploadHint');
    if (uploadHint) {
      uploadHint.innerHTML = `<strong>${file.name}</strong> (${rows.length} 行 × ${headers.length} 列)`;
    }

    render();
    // Only restore saved config if the saved grouping col still exists in the new file
    loadSavedConfig();
  } catch (err) {
    alert('读取文件失败: ' + err.message);
  }
}

// 搜索选择框相关函数
function updateGroupColDropdown() {
  const dropdown = document.getElementById('groupColDropdown');
  const input = document.getElementById('groupColInput');
  if (!dropdown || !input) return;

  const searchText = input.value.toLowerCase();
  const colOptions = headers.map((h, i) => ({ idx: i, name: cleanHeader(h) || `列${i + 1}` })).filter(h => h.name);
  const filtered = colOptions.filter(opt => opt.name.toLowerCase().includes(searchText));

  dropdown.innerHTML = filtered.map(opt => `
    <div class="search-select-option ${groupingColIdx === opt.idx ? 'selected' : ''}"
         data-idx="${opt.idx}" data-name="${opt.name}">
      ${opt.name}
    </div>
  `).join('');

  dropdown.classList.add('show');
}

function selectGroupCol(idx, name) {
  groupingColIdx = idx;
  document.getElementById('groupColInput').value = name;
  document.getElementById('groupColDropdown').classList.remove('show');
  render();
  saveConfig();
}

function initGroupColSearch() {
  const container = document.getElementById('groupColSearchSelect');
  if (!container) return;

  // 点击选项
  container.addEventListener('click', (e) => {
    const option = e.target.closest('.search-select-option');
    if (option) {
      selectGroupCol(parseInt(option.dataset.idx), option.dataset.name);
      return;
    }
  });

  // 输入框获得焦点时显示下拉
  container.addEventListener('focusin', (e) => {
    if (e.target.id === 'groupColInput') {
      updateGroupColDropdown();
    }
  });

  // 输入时搜索
  container.addEventListener('input', (e) => {
    if (e.target.id === 'groupColInput') {
      groupingColIdx = null;
      updateGroupColDropdown();
      hide(document.getElementById('groupValuesContainer'));
    }
  });

  // 点击外部关闭下拉
  document.addEventListener('click', (e) => {
    if (!e.target.closest('#groupColSearchSelect')) {
      const dropdown = document.getElementById('groupColDropdown');
      if (dropdown) dropdown.classList.remove('show');
    }
  });
}

function toggleGroupValue(val) {
  // 检查是否已存在基础表头（单条件组、单条件、单值）
  const exists = groups.find(g => {
    const c = getBasicCondition(g);
    return c && c.colIdx === groupingColIdx && c.values.length === 1 && c.values[0] === val;
  });
  if (exists) {
    groups = groups.filter(g => g !== exists);
  } else {
    groups.push({ name: val, relation: 'and', clauses: [{ relation: 'and', conditions: [{ colIdx: groupingColIdx, values: [val] }] }] });
  }
  render();
  saveConfig();
}

function openGroupEditor(groupIndex) {
  editingGroupIndex = groupIndex;
  if (groupIndex >= 0) {
    const g = migrateGroup(groups[groupIndex]);
    editingRelation = g.relation;
    editingClauses = JSON.parse(JSON.stringify(g.clauses));
    if (editingClauses.length === 0) {
      editingClauses = [{ relation: 'and', conditions: [{ colIdx: null, values: [] }] }];
    }
  } else {
    editingRelation = 'and';
    editingClauses = [{ relation: 'and', conditions: [{ colIdx: null, values: [] }] }];
  }
  render();
}

function cancelGroupEditor() {
  editingGroupIndex = -1;
  editingRelation = 'and';
  editingClauses = [];
  render();
}

function addClause() {
  editingClauses.push({ relation: 'and', conditions: [{ colIdx: null, values: [] }] });
  render();
}

function removeClause(clauseIdx) {
  editingClauses.splice(clauseIdx, 1);
  if (editingClauses.length === 0) {
    editingClauses.push({ relation: 'and', conditions: [{ colIdx: null, values: [] }] });
  }
  render();
}

function setTopRelation(rel) {
  editingRelation = rel === 'or' ? 'or' : 'and';
  render();
}

function setClauseRelation(clauseIdx, rel) {
  editingClauses[clauseIdx].relation = rel === 'or' ? 'or' : 'and';
  render();
}

function addCondition(clauseIdx) {
  editingClauses[clauseIdx].conditions.push({ colIdx: null, values: [] });
  render();
}

function removeCondition(clauseIdx, condIdx) {
  const conds = editingClauses[clauseIdx].conditions;
  conds.splice(condIdx, 1);
  if (conds.length === 0) {
    conds.push({ colIdx: null, values: [] });
  }
  render();
}

function handleConditionColChange(clauseIdx, condIdx, colIdx) {
  const cond = editingClauses[clauseIdx].conditions[condIdx];
  cond.colIdx = colIdx === '' ? null : parseInt(colIdx);
  cond.values = [];
  render();
}

function toggleConditionValue(clauseIdx, condIdx, val) {
  const cond = editingClauses[clauseIdx].conditions[condIdx];
  const idx = cond.values.indexOf(val);
  if (idx >= 0) {
    cond.values.splice(idx, 1);
  } else {
    cond.values.push(val);
  }
  render();
}

// 选项之间关系：多选题可要求同时含某几个选项（and）
function setConditionValuesRelation(clauseIdx, condIdx, rel) {
  const cond = editingClauses[clauseIdx].conditions[condIdx];
  cond.valuesRelation = rel === 'and' ? 'and' : 'or';
  render();
}

// 仅保留已完整的条件（已选列且已选值），丢弃空条件组 —— 实时预览与保存共用
function effectiveClauses() {
  return editingClauses
    .map(cl => ({
      relation: cl.relation,
      conditions: cl.conditions.filter(c => c.colIdx !== null && c.values.length > 0)
        .map(c => ({ colIdx: c.colIdx, values: c.values, valuesRelation: c.valuesRelation === 'and' ? 'and' : 'or' }))
    }))
    .filter(cl => cl.conditions.length > 0);
}

// 实时命中人数：与已选表头样本数同口径
function countEditingMatches() {
  const clauses = effectiveClauses();
  if (clauses.length === 0) return 0;
  return groupRowsCustom(rows, [{ name: '__preview__', relation: editingRelation, clauses }])['__preview__'].length;
}

function saveGroup() {
  const clauses = effectiveClauses();
  if (clauses.length === 0) {
    alert('请至少添加一个有效条件');
    return;
  }

  // 生成表头名称：条件内各值 '+'，条件间/组间按关系 ' & ' 或 ' | '，多条件组加括号
  const op = rel => (rel === 'or' ? ' | ' : ' & ');
  const clauseStrs = clauses.map(cl => {
    const s = cl.conditions.map(c => c.values.join(c.valuesRelation === 'and' ? '&' : '+')).join(op(cl.relation));
    return clauses.length > 1 && cl.conditions.length > 1 ? `(${s})` : s;
  });
  let name = clauseStrs.join(op(editingRelation));

  // 检查重名
  const existingNames = groups.map(g => g.name);
  if (existingNames.includes(name)) {
    let i = 2;
    while (existingNames.includes(`${name} (${i})`)) i++;
    name = `${name} (${i})`;
  }

  const newGroup = { name, relation: editingRelation, clauses };
  if (editingGroupIndex >= 0) {
    groups[editingGroupIndex] = newGroup;
  } else {
    groups.push(newGroup);
  }

  editingGroupIndex = -1;
  editingRelation = 'and';
  editingClauses = [];
  render();
  saveConfig();
}

function deleteGroup(groupIndex) {
  groups.splice(groupIndex, 1);
  render();
  saveConfig();
}

function startDragSelect(event, idx) {
  // 只响应左键
  if (event.button !== 0) return;
  isDraggingSelect = true;
  // 记录初始状态：如果当前未选中，则拖拽操作为选中；反之为取消
  const isSelected = selectedQuestions.includes(idx);
  dragSelectValue = !isSelected;
}

function handleDragSelect(idx) {
  if (!isDraggingSelect) return;
  setQuestionSelected(idx, dragSelectValue);
}

function setQuestionSelected(idx, selected) {
  const wasSelected = selectedQuestions.includes(idx);
  if (selected && !wasSelected) {
    selectedQuestions.push(idx);
    updateCheckbox(idx, true);
  } else if (!selected && wasSelected) {
    selectedQuestions = selectedQuestions.filter(i => i !== idx);
    updateCheckbox(idx, false);
  }
}

function updateCheckbox(idx, checked) {
  const checkbox = document.querySelector(`.question-item[data-question-idx="${idx}"] input[type="checkbox"]`);
  if (checkbox) checkbox.checked = checked;
  // 更新下载按钮状态
  document.getElementById('downloadBtn').disabled = selectedQuestions.length === 0;
}

function endDragSelect() {
  if (isDraggingSelect) {
    isDraggingSelect = false;
    saveConfig();
  }
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
  const complexGroupSection = document.getElementById('complexGroupSection');
  const selectedGroupsContainer = document.getElementById('selectedGroupsContainer');
  const qCard = document.getElementById('questionSelectorCard');
  if (headers.length === 0) {
    hide(groupCard);
    hide(qCard);
    return;
  }

  show(groupCard);
  complexGroupSection.style.display = 'flex';

  // 基础表头选择 - 搜索选择框
  const groupColInput = document.getElementById('groupColInput');
  const colOptions = headers.map((h, i) => ({ idx: i, name: cleanHeader(h) || `列${i + 1}` })).filter(h => h.name);

  // 更新输入框显示
  if (groupingColIdx !== null) {
    const selectedCol = colOptions.find(opt => opt.idx === groupingColIdx);
    groupColInput.value = selectedCol ? selectedCol.name : '';
  } else {
    groupColInput.value = '';
  }

  if (groupingColIdx === null) {
    hide(groupValuesContainer);
  } else {
    show(groupValuesContainer);

    // 多选单元格按英文逗号分割成独立选项，计数与选项数据同口径
    const valueCounts = countOptions(rows, groupingColIdx, false);
    const sortedValues = Object.keys(valueCounts).sort();
    const selectedValues = groups.map(getBasicCondition).filter(c =>
      c && c.colIdx === groupingColIdx && c.values.length === 1
    ).map(c => c.values[0]);

    document.getElementById('groupValues').innerHTML = sortedValues.map(val => `
      <button class="group-tag ${selectedValues.includes(val) ? 'selected' : 'unselected'}"
              onclick="toggleGroupValue('${val.replace(/'/g, "\\'")}')">${val} <span style="opacity:0.6;font-size:11px">(${valueCounts[val]})</span></button>
    `).join('');
  }

  // 复杂表头编辑器
  const groupEditorForm = document.getElementById('groupEditorForm');
  const addGroupBtn = document.getElementById('addGroupBtn');
  const groupConditions = document.getElementById('groupConditions');

  if (editingGroupIndex >= 0 || editingClauses.length > 0) {
    show(groupEditorForm);
    hide(addGroupBtn);

    const relSelect = (handler, current, andLabel, orLabel) => `
      <select onchange="${handler}" style="width:auto;">
        <option value="and" ${current === 'and' ? 'selected' : ''}>${andLabel}</option>
        <option value="or" ${current === 'or' ? 'selected' : ''}>${orLabel}</option>
      </select>`;

    // 条件组配色（黑白灰阶循环），多条件组时以色带区分各组
    const CLAUSE_COLORS = ['#2B2B2B', '#555555', '#808080', '#3D3D3D', '#6B6B6B', '#999999'];

    groupConditions.innerHTML = `
      ${editingClauses.length > 1 ? `
        <div style="display:flex; align-items:center; gap:12px; margin-bottom:12px;">
          <span style="font-size:12px; color:var(--muted);">条件组之间</span>
          ${relSelect('setTopRelation(this.value)', editingRelation, '且（同时满足所有组）', '或（满足任一组）')}
        </div>` : ''}
      ${editingClauses.map((cl, clauseIdx) => {
        const condRows = cl.conditions.map((cond, condIdx) => {
          const selectedCol = cond.colIdx !== null ? cond.colIdx : '';
          let valueButtons = '';

          if (selectedCol !== '') {
            // 多选单元格按英文逗号分割成独立选项，计数与选项数据同口径
            const valueCounts = countOptions(rows, parseInt(selectedCol), false);
            const sortedValues = Object.keys(valueCounts).sort();
            valueButtons = sortedValues.map(val => `
              <button class="group-tag ${cond.values.includes(val) ? 'selected' : 'unselected'}"
                      onclick="toggleConditionValue(${clauseIdx}, ${condIdx}, '${val.replace(/'/g, "\\'")}')">${val} <span style="opacity:0.6;font-size:11px">(${valueCounts[val]})</span></button>
            `).join('');
          }

          return `
            <div class="condition-item">
              <div style="display:flex; align-items:center; gap:12px; margin-bottom:12px;">
                <span style="font-size:12px; color:var(--muted);">条件 ${condIdx + 1}</span>
                <select onchange="handleConditionColChange(${clauseIdx}, ${condIdx}, this.value)" style="flex:1;">
                  <option value="">-- 选择列 --</option>
                  ${colOptions.map(opt => `<option value="${opt.idx}" ${parseInt(selectedCol) === opt.idx ? 'selected' : ''}>${opt.name}</option>`).join('')}
                </select>
                ${cl.conditions.length > 1 ? `<button onclick="removeCondition(${clauseIdx}, ${condIdx})" style="background:none; border:none; color:var(--muted); cursor:pointer;">×</button>` : ''}
              </div>
              ${selectedCol !== '' ? `
                <div style="display:flex; align-items:center; gap:8px; margin-bottom:8px;">
                  <span style="font-size:12px; color:var(--muted);">选项之间</span>
                  ${relSelect(`setConditionValuesRelation(${clauseIdx}, ${condIdx}, this.value)`, cond.valuesRelation === 'and' ? 'and' : 'or', '全部（且，多选题）', '任一（或）')}
                </div>
                <div class="group-values">${valueButtons}</div>` : ''}
            </div>
          `;
        }).join('');

        const clColor = CLAUSE_COLORS[clauseIdx % CLAUSE_COLORS.length];
        return `
          <div ${editingClauses.length > 1 ? `style="border-left:3px solid ${clColor}; background:${clColor}14; padding:10px 12px; margin-bottom:12px; border-radius:4px;"` : ''}>
            <div style="display:flex; align-items:center; gap:12px; margin-bottom:12px;">
              <span style="font-size:12px; ${editingClauses.length > 1 ? `color:${clColor}; font-weight:600;` : 'color:var(--muted);'}">${editingClauses.length > 1 ? `条件组 ${clauseIdx + 1}` : '条件关系'}</span>
              ${relSelect(`setClauseRelation(${clauseIdx}, this.value)`, cl.relation,
                editingClauses.length > 1 ? '组内：且' : '且（同时满足）',
                editingClauses.length > 1 ? '组内：或' : '或（满足任一）')}
              ${editingClauses.length > 1 ? `<button onclick="removeClause(${clauseIdx})" style="margin-left:auto; background:none; border:none; color:var(--muted); cursor:pointer;" title="删除条件组">×</button>` : ''}
            </div>
            ${condRows}
            <button class="btn btn-secondary" style="margin-top:12px; font-size:12px; padding:5px 12px;" onclick="addCondition(${clauseIdx})">+ 添加条件</button>
          </div>
        `;
      }).join('')}
    `;

    // 保存表头按钮右侧实时显示当前条件命中人数（口径与已选表头样本数一致）
    const matchCountEl = document.getElementById('groupMatchCount');
    if (matchCountEl) {
      const matchCount = countEditingMatches();
      matchCountEl.innerHTML = `命中 <strong style="color:var(--accent)">${matchCount}</strong> 人`;
    }
  } else {
    hide(groupEditorForm);
    show(addGroupBtn);
  }

  // 已选表头列表
  if (groups.length > 0) {
    show(selectedGroupsContainer);
    renderSelectedGroups();
  } else {
    hide(selectedGroupsContainer);
  }

  if (headers.length === 0) {
    hide(qCard);
    return;
  }
  show(qCard);

  // Build question list using modules array
  const questionItems = [];
  for (const mod of modules) {
    questionItems.push({ type: 'module-start', moduleId: mod.id, moduleName: mod.name });
    for (const qIdx of mod.questionIdxs) {
      const title = cleanHeader(headers[qIdx]) || `列${qIdx + 1}`;
      const isMulti = rows.length > 0 && isMultiChoice(rows, qIdx);
      questionItems.push({ type: 'question', idx: qIdx, title, isMulti, moduleId: mod.id });
    }
    questionItems.push({ type: 'module-end' });
  }

  let html = '';
  for (let i = 0; i < questionItems.length; i++) {
    const item = questionItems[i];
    if (item.type === 'module-start') {
      html += `<div class="question-module" data-module-id="${item.moduleId}">
        <div class="module-header">
          <input type="text" class="module-name-input" value="${item.moduleName}" onchange="renameModule(${item.moduleId}, this.value)">
          <button class="module-delete-btn" onclick="deleteModule(${item.moduleId})" title="Delete module">×</button>
        </div>
        <div class="module-questions">`;
    } else if (item.type === 'module-end') {
      html += `</div></div>`;
    } else if (item.type === 'question') {
      html += `
        <label class="question-item" data-module-id="${item.moduleId}" data-question-idx="${item.idx}"
               onmousedown="startDragSelect(event, ${item.idx})" onmouseenter="handleDragSelect(${item.idx})">
          <input type="checkbox" ${selectedQuestions.includes(item.idx) ? 'checked' : ''} onchange="toggleQuestion(${item.idx})">
          <span>${item.title}</span>
          <button class="add-module-btn" onclick="addModuleAfterQuestion(${item.idx})" title="Add module after this">+</button>
        </label>`;
    }
  }

  document.getElementById('questionList').innerHTML = '<div class="question-list">' + html + '</div>';

  document.getElementById('downloadBtn').disabled = selectedQuestions.length === 0;
}

function renderSelectedGroups() {
  const container = document.getElementById('selectedGroups');
  if (!container) return;

  // 计算每个表头的样本数量
  const grpData = groupRowsCustom(rows, groups);

  container.innerHTML = groups.map((g, i) => {
    const count = grpData[g.name] ? grpData[g.name].length : 0;
    return `
      <div class="group-tag selected" draggable="true"
           ondragstart="handleGroupDragStart(event, ${i})"
           ondragover="handleGroupDragOver(event, ${i})"
           ondragend="handleGroupDragEnd(event)">
        <span class="drag-handle">☰</span>
        <span>${g.name}</span>
        <span style="opacity:0.6; font-size:11px; margin-left:4px;">(${count})</span>
        <span class="group-action-btn" onclick="openGroupEditor(${i})" title="编辑">✎</span>
        <span class="group-action-btn" onclick="deleteGroup(${i})" title="删除">×</span>
      </div>
    `;
  }).join('');
}

// ==================== CONFIG ====================
function saveConfig() {
  const data = { groups, selectedQuestions, groupingColIdx, config, modules };
  localStorage.setItem('survey_analysis_config', JSON.stringify(data));
}

function loadSavedConfig() {
  try {
    const saved = localStorage.getItem('survey_analysis_config');
    if (saved) {
      const data = JSON.parse(saved);

      // 兼容旧格式：conditions: [[colIdx, val]] / [{colIdx, values}] → clauses 结构（见 migrateGroup）
      if (data.groups && Array.isArray(data.groups)) {
        groups = data.groups.map(migrateGroup);
      } else {
        groups = [];
      }

      if (data.selectedQuestions) selectedQuestions = data.selectedQuestions;
      if (data.config) config = data.config;
      // Restore modules or initialize
      if (data.modules && Array.isArray(data.modules) && data.modules.length > 0) {
        modules = data.modules;
        nextModuleId = Math.max(...modules.map(m => m.id)) + 1;
      } else {
        initModules();
      }
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
  editingGroupIndex = -1;
  editingRelation = 'and';
  editingClauses = [];
  initModules();
  localStorage.removeItem('survey_analysis_config');
  render();
}

function exportConfig() {
  const data = { groups, selectedQuestions, groupingColIdx, config, modules };
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
      if (data.groups) groups = data.groups.map(migrateGroup);
      if (data.selectedQuestions) selectedQuestions = data.selectedQuestions;
      if (data.groupingColIdx !== undefined) groupingColIdx = data.groupingColIdx;
      if (data.config) config = data.config;
      if (data.modules && Array.isArray(data.modules) && data.modules.length > 0) {
        modules = data.modules;
        nextModuleId = Math.max(...modules.map(m => m.id)) + 1;
      }
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
  document.getElementById('addGroupBtn').addEventListener('click', () => openGroupEditor(-1));
  document.getElementById('addClauseBtn').addEventListener('click', addClause);
  document.getElementById('saveGroupBtn').addEventListener('click', saveGroup);
  document.getElementById('cancelGroupBtn').addEventListener('click', cancelGroupEditor);

  // 全局 mouseup 事件结束拖拽选择
  document.addEventListener('mouseup', endDragSelect);

  render();
  initGroupColSearch();
  loadSavedConfig();
});