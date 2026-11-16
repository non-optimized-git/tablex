// TableX 独立运行版（tablex-standalone.html）同步检查
// 1) 内联 <script> 语法可解析；2) 关键导出常量与 app.js 一致（两份文件手动同步的最小回归）
const fs = require('fs');
const assert = require('assert');
const path = require('path');
const root = path.join(__dirname, '..');

const appSrc = fs.readFileSync(path.join(root, 'assets', 'js', 'app.js'), 'utf8');
const saSrc = fs.readFileSync(path.join(root, 'tablex-standalone.html'), 'utf8');

// 提取无 src 的内联 <script> 块
const scripts = [...saSrc.matchAll(/<script(?![^>]*\bsrc=)[^>]*>([\s\S]*?)<\/script>/g)].map(m => m[1]);
assert.ok(scripts.length > 0, 'standalone 应含内联 script');
for (const [i, code] of scripts.entries()) {
  // eslint-disable-next-line no-new-func
  new Function(code); // 仅语法检查，不执行
  console.log('inline script #' + (i + 1) + ' syntax OK (' + code.length + ' chars)');
}

// 关键常量两份文件必须一致
const consts = ['RED_FILL', 'YELLOW_FILL', 'GREEN_FILL', 'GRAY_FILL', 'BASE_FILL', 'AVG_FILL'];
const grab = (src, name) => {
  const m = src.match(new RegExp('const ' + name + " = '([^']+)'"));
  assert.ok(m, name + ' 未找到');
  return m[1];
};
for (const c of consts) {
  assert.strictEqual(grab(saSrc, c), grab(appSrc, c), c + ' 不一致：standalone 与 app.js 需手动同步');
}

console.log('STANDALONE CHECK: ' + scripts.length + ' script(s), ' + consts.length + ' constants OK');
