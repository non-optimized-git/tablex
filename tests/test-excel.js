// TableX Excel 导出色阶测试 —— 沙箱运行真实 app.js + stub ExcelJS
// 规则：红→黄→绿三色渐变（同截图/Excel 三色色阶）；数据区整体归一化（中点=中位数）；
// Total 列不上色、不参与极值；max=min 不上色；无表头不着色
const fs = require('fs');
const assert = require('assert');
const src = fs.readFileSync(require('path').join(__dirname, '..', 'assets', 'js', 'app.js'), 'utf8');

function fakeEl() {
  return {
    addEventListener() {}, value: '', innerHTML: '', style: {}, textContent: '',
    classList: { add() {}, remove() {}, toggle() {} },
    click() {}, appendChild() {}, removeChild() {}, focus() {},
    files: null, dataset: {},
  };
}
const fakeDoc = {
  addEventListener() {}, getElementById() { return fakeEl(); },
  createElement() { return fakeEl(); }, querySelectorAll() { return []; },
  body: { appendChild() {}, removeChild() {} },
};
const fakeStorage = { getItem: () => null, setItem() {}, removeItem() {} };

class FakeWorkbook {
  constructor() { this.worksheets = []; }
  addWorksheet(name) {
    const ws = {
      name, columns: null, _rows: [],
      addRow() {
        const row = { _cells: {}, getCell(i) { if (!this._cells[i]) this._cells[i] = {}; return this._cells[i]; } };
        this._rows.push(row);
        return row;
      },
    };
    this.worksheets.push(ws);
    return ws;
  }
}
const fakeExcelJS = { Workbook: FakeWorkbook };

const wrapped = src + `
;function __setState(h, r, g, m, sq) { headers = h; rows = r; groups = g; modules = m; selectedQuestions = sq; }
return { generateExcelWorkbook, mixFill, r3, __setState };`;
const api = new Function('document', 'localStorage', 'ExcelJS', 'alert', 'window', 'navigator', wrapped)(
  fakeDoc, fakeStorage, fakeExcelJS, () => {}, {}, {}
);

let n = 0;
const eq = (actual, expected, msg) => { assert.strictEqual(actual, expected, msg); n++; };

// --- mixFill 三色渐变锚点 ---
const RED = 'FFF8696B', YELLOW = 'FFFFEB84', GREEN = 'FF63BE7B';
eq(api.mixFill(0), RED, 't=0 → 红端点精确一致');
eq(api.mixFill(0.5), YELLOW, 't=0.5 → 黄中点精确一致');
eq(api.mixFill(1), GREEN, 't=1 → 绿端点精确一致');

const fillOf = c => c.style && c.style.fill ? c.style.fill.fgColor.argb : null;
const rowHasAnyFill = row => Object.values(row._cells).some(c => c.style && c.style.fill);

// ============ 场景1：无表头（仅 Total 列 → 不上色） ============
api.__setState(['Q1'], [...Array(5).fill(['A']), ...Array(2).fill(['B']), ['C']], [], [{ id: 1, name: 'M', questionIdxs: [0] }], [0]);
let wb = api.generateExcelWorkbook();
let ws = wb.worksheets[0];
for (const r of [4, 5, 6]) {
  eq(rowHasAnyFill(ws._rows[r]), false, '无表头数据行不上色（Total 列不着色）');
}

// ============ 场景2：3 个表头（数据区整体归一化，Total 不上色不参与极值） ============
// 男(4行) a,a,a,b | 女(3行) b,c,c | 其他(2行) a,c
const grp = (name, col, val) => ({ name, relation: 'and', clauses: [{ relation: 'and', conditions: [{ colIdx: col, values: [val] }] }] });
api.__setState(['性别', 'Q2'],
  [['男', 'a'], ['男', 'a'], ['男', 'a'], ['男', 'b'], ['女', 'b'], ['女', 'c'], ['女', 'c'], ['其他', 'a'], ['其他', 'c']],
  [grp('男', 0, '男'), grp('女', 0, '女'), grp('其他', 0, '其他')],
  [{ id: 1, name: 'M', questionIdxs: [1] }], [1]);
wb = api.generateExcelWorkbook();
ws = wb.worksheets[0];
const r3 = api.r3;
// 数据区 9 个值（分组列 × 行）：排序后 0,0,0,0.25,0.333,0.5,0.5,0.667,0.75
// min=0, mid=0.333(中位数), max=0.75 —— 整体一个色阶，不再逐行
const areaVals = [0.75, 0, 0.5, 0, r3(2 / 3), 0.5, 0.25, r3(1 / 3), 0].sort((a, b) => a - b);
const minV = areaVals[0], maxV = areaVals[areaVals.length - 1];
const midV = areaVals.length % 2 ? areaVals[(areaVals.length - 1) / 2] : (areaVals[areaVals.length / 2 - 1] + areaVals[areaVals.length / 2]) / 2;
const gradT = v => v <= minV ? 0 : v >= maxV ? 1
  : v <= midV ? 0.5 * (v - minV) / (midV - minV) : 0.5 + 0.5 * (v - midV) / (maxV - midV);
// 行序: title, sub, header, base, a/c/b（Total=4/3/2）。Table 3: opt=13, Total=14, 男=15, 女=16, 其他=17
const rowA = ws._rows[4], rowC = ws._rows[5], rowB = ws._rows[6];
// 行 a: 男=0.75(全局最高→绿), 女=0(全局最低→红), 其他=0.5(整体插值)
eq(fillOf(rowA.getCell(15)), GREEN, '行a 男(全局最高) → 绿端点');
eq(fillOf(rowA.getCell(16)), RED, '行a 女(全局最低) → 红端点');
eq(fillOf(rowA.getCell(17)), api.mixFill(gradT(0.5)), '行a 其他 → 整体插值');
eq(fillOf(rowA.getCell(14)), null, '行a Total → 不上色');
// 行 c: 男=0(→红), 女≈0.667(整体插值), 其他=0.5(与行a 其他同值同色 —— 整体标记验证)
eq(fillOf(rowC.getCell(15)), RED, '行c 男(全局最低) → 红端点');
eq(fillOf(rowC.getCell(16)), api.mixFill(gradT(r3(2 / 3))), '行c 女 → 整体插值');
eq(fillOf(rowC.getCell(17)), api.mixFill(gradT(0.5)), '行c 其他 → 整体插值');
eq(fillOf(rowC.getCell(17)), fillOf(rowA.getCell(17)), '同值跨行同色（整体归一化）');
eq(fillOf(rowC.getCell(14)), null, '行c Total → 不上色');
// 行 b: 男=0.25(整体插值), 女=0.333(中位数→黄), 其他=0(→红)
eq(fillOf(rowB.getCell(15)), api.mixFill(gradT(0.25)), '行b 男 → 整体插值');
eq(fillOf(rowB.getCell(16)), YELLOW, '行b 女(中位数) → 黄中点');
eq(fillOf(rowB.getCell(17)), RED, '行b 其他(全局最低) → 红端点');
eq(fillOf(rowB.getCell(14)), null, '行b Total → 不上色');
// Table 2 不上色
for (const row of [rowA, rowC, rowB]) {
  for (const col of [9, 10, 11]) eq(fillOf(row.getCell(col)), null, 'Table 2 不上色');
}
// 表头加粗（同截图）；选项列不加粗
eq(ws._rows[2].getCell(13).style.font.bold, true, '表头加粗');
eq(rowA.getCell(13).style.font.bold, undefined, '选项列不加粗');

// ============ 场景3：分组间全等（max=min）→ 不上色 ============
api.__setState(['性别', 'Q2'],
  [['男', 'x'], ['女', 'x']],
  [grp('男', 0, '男'), grp('女', 0, '女')],
  [{ id: 1, name: 'M', questionIdxs: [1] }], [1]);
wb = api.generateExcelWorkbook();
ws = wb.worksheets[0];
// Table 3: opt=11, Total=12, 男=13, 女=14
eq(rowHasAnyFill(ws._rows[4]), false, 'max=min → 分组列不上色');

// ============ 场景4：Total 不参与极值（分组全 0，Total 100% → 仍不上色） ============
// 男(2行, 空) | 其他(2行, a)；表头=男+女 → 分组 pcts 全 0，Total=100%
api.__setState(['性别', 'Q2'],
  [['男', ''], ['男', ''], ['其他', 'a'], ['其他', 'a']],
  [grp('男', 0, '男'), grp('女', 0, '女')],
  [{ id: 1, name: 'M', questionIdxs: [1] }], [1]);
wb = api.generateExcelWorkbook();
ws = wb.worksheets[0];
const rowQ4 = ws._rows[4];
eq(rowQ4.getCell(12).value, 1, '场景4 Total 百分比 = 100%');
eq(fillOf(rowQ4.getCell(12)), null, 'Total 不上色');
eq(rowHasAnyFill(rowQ4), false, 'Total 不参与极值（分组全 0 → 无着色）');

console.log('EXCEL GRADIENT TESTS: ' + n + ' assertions PASSED');
