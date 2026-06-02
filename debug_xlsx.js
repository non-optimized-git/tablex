const ExcelJS = require('exceljs');

async function main() {
  const wb = new ExcelJS.Workbook();
  await wb.xlsx.readFile('/Users/yuanyi.li/Downloads/分析结果 (10).xlsx');

  let totalFills = 0;
  wb.worksheets.forEach((ws, si) => {
    const spacers = [];
    for (let c = 1; c <= ws.columnCount; c++) {
      if (ws.getColumn(c).width === 1) spacers.push(c);
    }
    if (spacers.length < 2) { console.log('Sheet ' + (si+1) + ' 无spacer列'); return; }

    const numGroups = spacers[0] - 3; // col before spacer - opt - Total = group cols
    const table3Start = spacers[1] + 1; // table3 first col
    const table3GroupCols = [];
    for (let i = 0; i < numGroups; i++) {
      table3GroupCols.push(table3Start + 2 + i); // opt=table3Start, Total=table3Start+1, then groups
    }

    // Find all rows with table3 data (skip title rows)
    // Strategy: find rows where table3 opt col has a text value
    const optCol = table3Start;
    const headerRow = ws.getRow(7); // R7 is header
    const baseRow = ws.getRow(8);

    let greens = 0, reds = 0, others = 0;
    let rowsChecked = 0;

    // Scan ALL rows in the sheet
    for (let r = 1; r <= ws.rowCount; r++) {
      const row = ws.getRow(r);
      const optVal = row.getCell(optCol).value;
      if (!optVal || typeof optVal !== 'string' || /^(base|选项|题目|平均值)/.test(optVal)) continue;

      rowsChecked++;
      const groupPcts = [];
      const groupFills = [];

      for (let i = 0; i < numGroups; i++) {
        const c = table3GroupCols[i];
        const cell = row.getCell(c);
        const fill = cell.fill;
        const argb = fill && fill.fgColor && fill.fgColor.argb ? fill.fgColor.argb : 'none';
        groupPcts.push(cell.value);
        groupFills.push(argb);
        if (argb !== 'none' && argb !== '00000000') totalFills++;
      }

      const pcts = groupPcts.filter(v => v !== null && v !== undefined && v !== '').map(Number);
      if (pcts.length < 2) continue;

      const max = Math.max(...pcts);
      const min = Math.min(...pcts);

      // Count colored cells
      const redCount = groupFills.filter(f => f === 'FFFFCCCC').length;
      const greenCount = groupFills.filter(f => f === 'FFCCFFCC').length;

      if (redCount > 1 || greenCount > 1) {
        console.log(`\nSheet${si+1} R${r} opt="${optVal.slice(0,10)}" pcts=[${pcts.map(p=>p.toFixed(4)).join(',')}] max=${max.toFixed(4)} min=${min.toFixed(4)}`);
        console.log('  fills:', groupFills.join(', '), '-> REDS=' + redCount + ' GREENS=' + greenCount);
      }

      greens += greenCount;
      reds += redCount;
    }

    console.log(`\nSheet ${si+1} "${ws.name}": ${rowsChecked} rows checked, reds=${reds} greens=${greens}`);
  });

  console.log('\n总 fill 数:', totalFills);
}

main().catch(console.error);