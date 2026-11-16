// TableX 复杂表头（分组括号表达式）逻辑测试 —— 直接抽取 app.js 真实函数
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
for (const fn of ['safe', 'splitOpts', 'migrateGroup', 'getBasicCondition', 'condMatches', 'clauseMatches', 'groupRowsCustom', 'effectiveClauses', 'countEditingMatches']) {
  eval(extractFn(fn));
}

// 6 行数据：0=性别, 1=年龄, 2=城市
const rows = [
  ['男', '18-25岁', '北京'], // 0
  ['男', '26-35岁', '北京'], // 1
  ['女', '18-25岁', '上海'], // 2
  ['女', '36-45岁', '上海'], // 3
  ['男', '18-25岁', '上海'], // 4
  ['女', '26-35岁', '北京'], // 5
];
globalThis.rows = rows;

const count = g => groupRowsCustom(rows, [g])[g.name].length;

// --- migrateGroup 三种历史格式 ---
const v1 = migrateGroup({ name: '男', conditions: [[0, '男']] });
assert.strictEqual(v1.clauses.length, 1);
assert.deepStrictEqual(v1.clauses[0].conditions, [{ colIdx: 0, values: ['男'], valuesRelation: 'or' }]);
assert.strictEqual(count(v1), 3, '旧格式1（pair数组）');

const v2 = migrateGroup({ name: '男', conditions: [{ colIdx: 0, values: ['男'] }] });
assert.strictEqual(count(v2), 3, '旧格式2（conditions）');

const v3 = migrateGroup({ name: 'x', relation: 'or', clauses: [{ relation: 'or', conditions: [{ colIdx: 0, values: ['男'] }] }] });
assert.strictEqual(v3.relation, 'or');
assert.strictEqual(v3.clauses[0].relation, 'or', '新格式透传');

// --- 分组括号语义 ---
// (性别=男 且 年龄=18-25岁) 或 (性别=女 且 年龄=26-35岁) → 行 0,4,5 = 3
assert.strictEqual(count({
  name: 'p1', relation: 'or',
  clauses: [
    { relation: 'and', conditions: [{ colIdx: 0, values: ['男'] }, { colIdx: 1, values: ['18-25岁'] }] },
    { relation: 'and', conditions: [{ colIdx: 0, values: ['女'] }, { colIdx: 1, values: ['26-35岁'] }] },
  ]
}), 3, '(A且B)或(C且D)');

// 性别=男 或 城市=北京（单组内 OR，跨列）→ 行 0,1,4,5 = 4
assert.strictEqual(count({
  name: 'p2', relation: 'and',
  clauses: [{ relation: 'or', conditions: [{ colIdx: 0, values: ['男'] }, { colIdx: 2, values: ['北京'] }] }]
}), 4, '跨列 OR');

// 性别=女 且 (年龄=18-25岁 或 年龄=36-45岁) → 行 2,3 = 2
assert.strictEqual(count({
  name: 'p3', relation: 'and',
  clauses: [
    { relation: 'and', conditions: [{ colIdx: 0, values: ['女'] }] },
    { relation: 'or', conditions: [{ colIdx: 1, values: ['18-25岁'] }, { colIdx: 1, values: ['36-45岁'] }] },
  ]
}), 2, 'A且(B或C)');

// 旧语义回归：单组 AND 跨列 → 性别=男 且 城市=北京 → 行 0,1 = 2
assert.strictEqual(count({
  name: 'p4', relation: 'and',
  clauses: [{ relation: 'and', conditions: [{ colIdx: 0, values: ['男'] }, { colIdx: 2, values: ['北京'] }] }]
}), 2, '旧语义 AND 回归');

// --- getBasicCondition ---
assert.deepStrictEqual(getBasicCondition({ name: '男', clauses: [{ relation: 'and', conditions: [{ colIdx: 0, values: ['男'] }] }] }), { colIdx: 0, values: ['男'], valuesRelation: 'or' });
assert.strictEqual(getBasicCondition({ name: 'x', clauses: [{ relation: 'and', conditions: [{ colIdx: 0, values: ['男'] }, { colIdx: 1, values: ['a'] }] }] }), null);

// --- effectiveClauses / countEditingMatches 预览口径 ---
globalThis.editingRelation = 'and';
globalThis.editingClauses = [{ relation: 'and', conditions: [{ colIdx: null, values: [] }] }];
assert.strictEqual(countEditingMatches(), 0, '空条件 → 0');

globalThis.editingClauses = [
  { relation: 'and', conditions: [{ colIdx: 0, values: ['男'] }, { colIdx: null, values: [] }] },
  { relation: 'and', conditions: [{ colIdx: null, values: [] }] },
];
assert.strictEqual(countEditingMatches(), 3, '不完整条件/空组被忽略 → 男=3');

globalThis.editingRelation = 'or';
globalThis.editingClauses = [
  { relation: 'and', conditions: [{ colIdx: 0, values: ['男'] }, { colIdx: 1, values: ['18-25岁'] }] },
  { relation: 'and', conditions: [{ colIdx: 0, values: ['女'] }, { colIdx: 1, values: ['26-35岁'] }] },
];
assert.strictEqual(countEditingMatches(), 3, '预览与保存同口径 (A且B)或(C且D) → 3');

console.log('CLAUSE TESTS: 13/13 PASSED');
