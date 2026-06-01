const ExcelJS = require('exceljs');

async function main() {
  const wb = new ExcelJS.Workbook();
  await wb.xlsx.readFile('/Users/yuanyi.li/Downloads/分析结果 (6).xlsx');
  const ws = wb.getWorksheet(1);

  // === 核心：精确分析前3行的每列值 ===
  console.log('=== 前3行完整数据 (所有列) ===');
  for (let r = 1; r <= 3; r++) {
    const row = ws.getRow(r);
    console.log(`\nR${r}:`);
    for (let c = 1; c <= ws.columnCount; c++) {
      const cell = row.getCell(c);
      if (cell.value !== null && cell.value !== undefined) {
        const w = ws.getColumn(c).width;
        console.log(`  col${c} (w=${w}): "${cell.value}"`);
      }
    }
  }

  // === 验证 app.js 代码的逻辑问题 ===
  console.log('\n=== 代码逻辑分析 ===');
  // hasGroups=true 时，列宽公式: 3*numGroups+8 = 3*5+8 = 23 (匹配实际23列)
  // spacer 位置 (0-indexed): c === numGroups+2 = 7, 或 c === 2*(numGroups+2)+1 = 15
  // 即 1-indexed: 8, 16 —— 匹配实际 spacer 列 8, 16
  // 但 R3 中 col8 = "选项" 而不是 spacer！这说明 cellIdx 偏移了
  //
  // 原因：在 hasGroups=true 路径中，header row 从 cellIdx=1 开始顺序写入
  // cellIdx 递增序列: 1(选项) 2(Total) 3(男) 4(女) 5(5个) 6(7个) 7(6个) 8(选项)
  // → col8="选项" 但 col8 应该是 spacer (width=1)!
  //
  // 根本原因：代码写入 header 时没有考虑 spacer 列的位置，
  // spacer 列应该被 cellIdx 跳过但实际没有被跳过！

  const numGroups = 5; // 从 R3 推出: 7列/表 - 2 = 5组
  console.log('推算 numGroups =', numGroups);
  console.log('hasGroups=true 时的 spacer 列(1-indexed):', (numGroups+2+1), '和', (2*(numGroups+2)+1+1));
  console.log('实际 spacer 列 (width=1): 8, 16 ✓');

  console.log('\n=== 问题根因 ===');
  console.log('hasGroups=true 时，header row 的 cellIdx 从 1 开始顺序累加');
  console.log('它没有跳过 spacer 位置的列');
  console.log('导致 col8（应该是 spacer）被写入了"选项"');

  console.log('\n=== 修复方向 ===');
  console.log('hasGroups=true 时，header 写入必须考虑 spacer 列位置');
  console.log('每个表格写入前，需要先判断 spacer 列的位置并跳过');
  console.log('或者：始终按 9列无分组结构写入，spacer 用 width=1 列占据');
  console.log('');
  console.log('正确做法：每个表格写入 opt+Total+numGroups 列后，额外跳过 1 个 spacer 位');
}

main().catch(console.error);