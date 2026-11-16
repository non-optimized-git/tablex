// TableX 表头多选分割测试 —— 选项列表与匹配均按英文逗号分割（与选项数据同口径）
const fs = require('fs');
const assert = require('assert');
const src = fs.readFileSync(require('path').join(__dirname, '..', 'assets', 'js', 'app.js'), 'utf8');

function extractFn(name) {
  const start = src.indexOf('function ' + name + '(');
  if (start < 0) throw new Error('not found: ' + name);
  let depth = 0;
  for (let j = src.indexOf('{', start); j < src.length; j++) {
    if (src[j] === '{') depth++;
    else if (src[j] === '}') { depth--; if (depth === 0) return src.slice(start, j + 1); }
  }
}
for (const fn of ['safe', 'splitOpts', 'countOptions', 'condMatches', 'clauseMatches', 'groupRowsCustom', 'migrateGroup', 'getBasicCondition']) {
  eval(extractFn(fn));
}

const count = (rows, g) => groupRowsCustom(rows, [g])[g.name].length;
const basic = (name, colIdx, values, valuesRelation) => ({
  name, relation: 'and',
  clauses: [{ relation: 'and', conditions: [valuesRelation ? { colIdx, values, valuesRelation } : { colIdx, values }] }]
});

// --- 选项列表口径：countOptions 按逗号分割，与题目选项数据一致 ---
const counts = countOptions([['A,B'], ['A,C'], ['A'], ['A,B']], 0, false);
assert.deepStrictEqual(counts, { A: 4, B: 2, C: 1 }, '多选拆分计数（与选项数据同口径）');
assert.deepStrictEqual(Object.keys(counts).sort(), ['A', 'B', 'C'], '选项列表为拆分后的独立选项');

// --- 匹配：拆分选项命中多选单元格 ---
assert.strictEqual(count([['A,B'], ['B,C'], ['C']], basic('g', 0, ['A'])), 1, '拆分选项 A 命中 "A,B" 行');
assert.strictEqual(count([['A,B'], ['A,C'], ['A,D']], basic('g', 0, ['B', 'C'])), 2, '拆分选项多值 OR');
assert.strictEqual(count([['A, X'], ['A']], basic('g', 0, ['X'])), 1, '拆分后 trim（" X"→"X"）');

// --- 兼容：历史配置存的整值仍可命中 ---
assert.strictEqual(count([['A,B'], ['A,B,C'], ['A,B']], basic('g', 0, ['A,B'])), 2, '整值命中兼容历史配置');

// --- 回归：单选列不受影响 ---
const rows1 = [['男'], ['女'], ['男']];
assert.strictEqual(count(rows1, basic('男', 0, ['男'])), 2, '单选列回归');

// --- 跨列 AND，两边都是多选 ---
const rows2 = [['A,B', 'X, Y'], ['A', 'Y,Z'], ['A,B', 'X']];
assert.strictEqual(count(rows2, {
  name: 'g', relation: 'and',
  clauses: [{ relation: 'and', conditions: [{ colIdx: 0, values: ['A'] }, { colIdx: 1, values: ['Y'] }] }]
}), 2, '跨列 AND 两侧均拆分匹配');

// --- 选项之间 且/或（多选题） ---
assert.strictEqual(count([['A,B'], ['A,B,C'], ['A'], ['B']], basic('g', 0, ['A', 'B'], 'and')), 2, '全部（且）：须同时含 A 和 B');
assert.strictEqual(count([['B,A']], basic('g', 0, ['A', 'B'], 'and')), 1, '且 与选项顺序无关');
assert.strictEqual(count([['A'], ['B'], ['C']], basic('g', 0, ['A', 'B'], 'or')), 2, '任一（或）显式指定');
assert.strictEqual(count([['A'], ['B'], ['C']], basic('g', 0, ['A', 'B'])), 2, '缺省 valuesRelation = 或');
assert.strictEqual(count([['男'], ['女']], basic('g', 0, ['男', '女'], 'and')), 0, '单选列 且 不可能同时命中');
const mg1 = migrateGroup({ name: 'x', relation: 'and', clauses: [{ relation: 'and', conditions: [{ colIdx: 0, values: ['A', 'B'], valuesRelation: 'and' }] }] });
assert.strictEqual(mg1.clauses[0].conditions[0].valuesRelation, 'and', 'migrateGroup 保留 and');
const mg2 = migrateGroup({ name: 'x', conditions: [{ colIdx: 0, values: ['A'] }] });
assert.strictEqual(mg2.clauses[0].conditions[0].valuesRelation, 'or', 'migrateGroup 缺省 or');

console.log('MULTI-CHOICE HEADER TESTS: 15/15 PASSED');
