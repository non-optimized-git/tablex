import ExcelJS from 'exceljs';
import { detectModule, getShortTitle } from './constants.js';
import { isMultiChoice, countOptions, countUsers } from './analyzer.js';

function r3(x) {
  if (x === null || x === undefined) return null;
  return Math.abs(x) < 0.0005 ? 0 : Math.round(x * 1000) / 1000;
}

const RED_FILL = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFFFCCCC' } };
const GREEN_FILL = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFCCFFCC' } };
const GRAY_FILL = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFD9D9D9' } };

function border(cell) {
  cell.border = { top: { style: 'thin' }, bottom: { style: 'thin' }, left: { style: 'thin' }, right: { style: 'thin' } };
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

function calcNumericMean(rows, colIdx) {
  let total = 0, sum = 0;
  for (const row of rows) {
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

export async function generateExcel(questions, headers, rows, groups, orderedGroups, config) {
  const workbook = new ExcelJS.Workbook();

  const modules = {};
  for (const colIdx of questions) {
    const module = detectModule(headers[colIdx] || `列${colIdx}`);
    if (!modules[module]) modules[module] = [];
    modules[module].push(colIdx);
  }

  const numGroups = orderedGroups.length;
  const hasGroups = numGroups > 0;
  // 三列布局: 频数表(选项,Total,组...), 占比表(选项,Total,组...), 镜像表(选项,Total,组...)
  // A=选项, B=Total, C..(numGroups) = 组, 空1列, F=选项, G=Total, H.. = 组, 空1列, K=选项, L=Total, M.. = 组
  const table2Start = 2 + 1 + numGroups + 1; // B + Total + numGroups + gap
  const table3Start = table2Start + 1 + 1 + numGroups + 1; // F + Total + numGroups + gap

  for (const [moduleName, colList] of Object.entries(modules)) {
    const worksheet = workbook.addWorksheet(moduleName.slice(0, 31));

    for (const colIdx of colList) {
      const header = headers[colIdx] || `列${colIdx}`;
      const title = getShortTitle(header);
      const isMulti = isMultiChoice(rows, colIdx);

      const titleRow = worksheet.addRow();
      titleRow.getCell(1).value = title;
      titleRow.getCell(1).font = { name: '黑体', size: 14 };

      const qRow = worksheet.addRow();
      qRow.getCell(1).value = `题目: ${title} (${isMulti ? '多选' : '单选'})`;
      qRow.getCell(1).font = { name: '黑体', size: 9 };

      if (!hasGroups) {
        // 无分组模式：选项 | 次数 | 占比
        const counts = countOptions(rows, colIdx, isMulti);
        const totalUsers = countUsers(rows, colIdx);
        const allOpts = Object.keys(counts).sort((a, b) => counts[b] - counts[a]);

        const hRow = worksheet.addRow(['选项', '次数', '占比']);
        hRow.getCell(1).font = { name: '黑体', size: 10 };
        hRow.getCell(1).fill = GRAY_FILL;
        hRow.getCell(1).alignment = { horizontal: 'left' };
        hRow.getCell(2).font = { name: '黑体', size: 10 };
        hRow.getCell(2).fill = GRAY_FILL;
        hRow.getCell(2).alignment = { horizontal: 'center' };
        hRow.getCell(3).font = { name: '黑体', size: 10 };
        hRow.getCell(3).fill = GRAY_FILL;
        hRow.getCell(3).alignment = { horizontal: 'center' };
        border(hRow.getCell(1));
        border(hRow.getCell(2));
        border(hRow.getCell(3));

        const baseRow = worksheet.addRow(['base', totalUsers, '100%']);
        baseRow.getCell(1).font = { name: '黑体', size: 10 };
        baseRow.getCell(1).alignment = { horizontal: 'left' };
        baseRow.getCell(2).font = { name: '黑体', size: 10 };
        baseRow.getCell(2).alignment = { horizontal: 'center' };
        baseRow.getCell(3).font = { name: '黑体', size: 10 };
        baseRow.getCell(3).alignment = { horizontal: 'center' };
        border(baseRow.getCell(1));
        border(baseRow.getCell(2));
        border(baseRow.getCell(3));

        for (const opt of allOpts) {
          const val = counts[opt] || 0;
          const pct = totalUsers > 0 ? r3(val / totalUsers) : 0;
          const dataRow = worksheet.addRow([opt, val, pct]);
          dataRow.getCell(1).font = { name: '黑体', size: 11 };
          dataRow.getCell(1).alignment = { horizontal: 'left' };
          dataRow.getCell(2).font = { name: '黑体', size: 11 };
          dataRow.getCell(2).alignment = { horizontal: 'center' };
          dataRow.getCell(3).font = { name: '黑体', size: 11 };
          dataRow.getCell(3).alignment = { horizontal: 'center' };
          dataRow.getCell(3).numFmt = '0.0%';
          border(dataRow.getCell(1));
          border(dataRow.getCell(2));
          border(dataRow.getCell(3));
        }

        const meanVal = calcNumericMean(rows, colIdx);
        if (meanVal !== null) {
          const meanRow = worksheet.addRow(['平均值', meanVal, '']);
          meanRow.getCell(1).font = { name: '黑体', size: 10 };
          meanRow.getCell(1).alignment = { horizontal: 'left' };
          meanRow.getCell(2).font = { name: '黑体', size: 10 };
          meanRow.getCell(2).alignment = { horizontal: 'center' };
          meanRow.getCell(2).numFmt = '0.0';
          meanRow.getCell(3).font = { name: '黑体', size: 10 };
          border(meanRow.getCell(1));
          border(meanRow.getCell(2));
          border(meanRow.getCell(3));
        }

        worksheet.addRow();
        worksheet.getColumn(1).width = 25;
        worksheet.getColumn(2).width = 10;
        worksheet.getColumn(3).width = 10;

      } else {
        // 分组模式：三列表格并排
        const groupCounts = {};
        const groupUsers = {};
        for (const g of orderedGroups) {
          groupCounts[g] = countOptions(groups[g], colIdx, isMulti);
          groupUsers[g] = countUsers(groups[g], colIdx);
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

        // 三表Header
        const hRow = worksheet.addRow();
        // 表1: 频数
        hRow.getCell(1).value = '选项';
        hRow.getCell(1).font = { name: '黑体', size: 10 };
        hRow.getCell(1).fill = GRAY_FILL;
        hRow.getCell(1).alignment = { horizontal: 'left' };
        border(hRow.getCell(1));
        hRow.getCell(2).value = 'Total';
        hRow.getCell(2).font = { name: '黑体', size: 10 };
        hRow.getCell(2).fill = GRAY_FILL;
        hRow.getCell(2).alignment = { horizontal: 'center' };
        border(hRow.getCell(2));
        for (let i = 0; i < numGroups; i++) {
          const cell = hRow.getCell(3 + i);
          cell.value = orderedGroups[i];
          cell.font = { name: '黑体', size: 10 };
          cell.fill = GRAY_FILL;
          cell.alignment = { horizontal: 'center' };
          border(cell);
        }
        // 表2: 占比（无颜色）
        hRow.getCell(table2Start).value = '选项';
        hRow.getCell(table2Start).font = { name: '黑体', size: 10 };
        hRow.getCell(table2Start).fill = GRAY_FILL;
        hRow.getCell(table2Start).alignment = { horizontal: 'left' };
        border(hRow.getCell(table2Start));
        hRow.getCell(table2Start + 1).value = 'Total';
        hRow.getCell(table2Start + 1).font = { name: '黑体', size: 10 };
        hRow.getCell(table2Start + 1).fill = GRAY_FILL;
        hRow.getCell(table2Start + 1).alignment = { horizontal: 'center' };
        border(hRow.getCell(table2Start + 1));
        for (let i = 0; i < numGroups; i++) {
          const cell = hRow.getCell(table2Start + 2 + i);
          cell.value = orderedGroups[i];
          cell.font = { name: '黑体', size: 10 };
          cell.fill = GRAY_FILL;
          cell.alignment = { horizontal: 'center' };
          border(cell);
        }
        // 表3: 镜像+红绿
        hRow.getCell(table3Start).value = '选项';
        hRow.getCell(table3Start).font = { name: '黑体', size: 10 };
        hRow.getCell(table3Start).fill = GRAY_FILL;
        hRow.getCell(table3Start).alignment = { horizontal: 'left' };
        border(hRow.getCell(table3Start));
        hRow.getCell(table3Start + 1).value = 'Total';
        hRow.getCell(table3Start + 1).font = { name: '黑体', size: 10 };
        hRow.getCell(table3Start + 1).fill = GRAY_FILL;
        hRow.getCell(table3Start + 1).alignment = { horizontal: 'center' };
        border(hRow.getCell(table3Start + 1));
        for (let i = 0; i < numGroups; i++) {
          const cell = hRow.getCell(table3Start + 2 + i);
          cell.value = orderedGroups[i];
          cell.font = { name: '黑体', size: 10 };
          cell.fill = GRAY_FILL;
          cell.alignment = { horizontal: 'center' };
          border(cell);
        }

        // Base行
        const baseRow = worksheet.addRow();
        // 表1 Base
        baseRow.getCell(1).value = 'base';
        baseRow.getCell(1).font = { name: '黑体', size: 10 };
        baseRow.getCell(1).alignment = { horizontal: 'left' };
        border(baseRow.getCell(1));
        baseRow.getCell(2).value = totalAllUsers;
        baseRow.getCell(2).font = { name: '黑体', size: 10 };
        baseRow.getCell(2).alignment = { horizontal: 'center' };
        border(baseRow.getCell(2));
        for (let i = 0; i < numGroups; i++) {
          const cell = baseRow.getCell(3 + i);
          cell.value = groupUsers[orderedGroups[i]] || 0;
          cell.font = { name: '黑体', size: 10 };
          cell.alignment = { horizontal: 'center' };
          border(cell);
        }
        // 表2 Base
        baseRow.getCell(table2Start).value = 'base';
        baseRow.getCell(table2Start).font = { name: '黑体', size: 10 };
        baseRow.getCell(table2Start).alignment = { horizontal: 'left' };
        border(baseRow.getCell(table2Start));
        baseRow.getCell(table2Start + 1).value = totalAllUsers;
        baseRow.getCell(table2Start + 1).font = { name: '黑体', size: 10 };
        baseRow.getCell(table2Start + 1).alignment = { horizontal: 'center' };
        border(baseRow.getCell(table2Start + 1));
        for (let i = 0; i < numGroups; i++) {
          const cell = baseRow.getCell(table2Start + 2 + i);
          cell.value = groupUsers[orderedGroups[i]] || 0;
          cell.font = { name: '黑体', size: 10 };
          cell.alignment = { horizontal: 'center' };
          border(cell);
        }
        // 表3 Base
        baseRow.getCell(table3Start).value = 'base';
        baseRow.getCell(table3Start).font = { name: '黑体', size: 10 };
        baseRow.getCell(table3Start).alignment = { horizontal: 'left' };
        border(baseRow.getCell(table3Start));
        baseRow.getCell(table3Start + 1).value = totalAllUsers;
        baseRow.getCell(table3Start + 1).font = { name: '黑体', size: 10 };
        baseRow.getCell(table3Start + 1).alignment = { horizontal: 'center' };
        border(baseRow.getCell(table3Start + 1));
        for (let i = 0; i < numGroups; i++) {
          const cell = baseRow.getCell(table3Start + 2 + i);
          cell.value = groupUsers[orderedGroups[i]] || 0;
          cell.font = { name: '黑体', size: 10 };
          cell.alignment = { horizontal: 'center' };
          border(cell);
        }

        // Data行
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

          const dataRow = worksheet.addRow();

          // 表1: 频数
          dataRow.getCell(1).value = opt;
          dataRow.getCell(1).font = { name: '黑体', size: 11 };
          dataRow.getCell(1).alignment = { horizontal: 'left' };
          border(dataRow.getCell(1));
          dataRow.getCell(2).value = totalVal;
          dataRow.getCell(2).font = { name: '黑体', size: 11 };
          dataRow.getCell(2).alignment = { horizontal: 'center' };
          border(dataRow.getCell(2));
          for (let i = 0; i < numGroups; i++) {
            const cell = dataRow.getCell(3 + i);
            cell.value = groupCounts[orderedGroups[i]][opt] || 0;
            cell.font = { name: '黑体', size: 11 };
            cell.alignment = { horizontal: 'center' };
            border(cell);
          }

          // 表2: 占比（无颜色）
          dataRow.getCell(table2Start).value = opt;
          dataRow.getCell(table2Start).font = { name: '黑体', size: 11 };
          dataRow.getCell(table2Start).alignment = { horizontal: 'left' };
          border(dataRow.getCell(table2Start));
          dataRow.getCell(table2Start + 1).value = totalPct;
          dataRow.getCell(table2Start + 1).font = { name: '黑体', size: 11 };
          dataRow.getCell(table2Start + 1).alignment = { horizontal: 'center' };
          dataRow.getCell(table2Start + 1).numFmt = '0.0%';
          border(dataRow.getCell(table2Start + 1));
          for (let i = 0; i < numGroups; i++) {
            const cell = dataRow.getCell(table2Start + 2 + i);
            cell.value = pcts[orderedGroups[i]];
            cell.font = { name: '黑体', size: 11 };
            cell.alignment = { horizontal: 'center' };
            cell.numFmt = '0.0%';
            border(cell);
          }

          // 表3: 镜像+红绿
          dataRow.getCell(table3Start).value = opt;
          dataRow.getCell(table3Start).font = { name: '黑体', size: 11 };
          dataRow.getCell(table3Start).alignment = { horizontal: 'left' };
          border(dataRow.getCell(table3Start));
          dataRow.getCell(table3Start + 1).value = totalPct;
          dataRow.getCell(table3Start + 1).font = { name: '黑体', size: 11 };
          dataRow.getCell(table3Start + 1).alignment = { horizontal: 'center' };
          dataRow.getCell(table3Start + 1).numFmt = '0.0%';
          border(dataRow.getCell(table3Start + 1));
          if (totalPct === maxPct && maxPct > 0) {
            dataRow.getCell(table3Start + 1).fill = RED_FILL;
          } else if (totalPct === minPct && minPct > 0 && maxPct !== minPct) {
            dataRow.getCell(table3Start + 1).fill = GREEN_FILL;
          }
          for (let i = 0; i < numGroups; i++) {
            const g = orderedGroups[i];
            const cellPct = pcts[g];
            const cell = dataRow.getCell(table3Start + 2 + i);
            cell.value = cellPct;
            cell.numFmt = '0.0%';
            cell.font = { name: '黑体', size: 11 };
            cell.alignment = { horizontal: 'center' };
            border(cell);
            if (cellPct === maxPct && maxPct > 0) {
              cell.fill = RED_FILL;
            } else if (cellPct === minPct && minPct > 0 && maxPct !== minPct) {
              cell.fill = GREEN_FILL;
            }
          }
        }

        // 均值行
        const means = {};
        for (const g of orderedGroups) {
          means[g] = calcNumericMean(groups[g], colIdx);
        }
        const totalMean = calcNumericMean(rows, colIdx);
        const validMeans = Object.fromEntries(Object.entries(means).filter(([_, v]) => v !== null));
        if (totalMean !== null) validMeans['Total'] = totalMean;
        const maxM = Object.values(validMeans).length > 0 ? Math.max(...Object.values(validMeans)) : null;
        const minM = Object.values(validMeans).length > 0 ? Math.min(...Object.values(validMeans)) : null;

        const meanRow = worksheet.addRow();
        // 表1 均值
        meanRow.getCell(1).value = '平均值';
        meanRow.getCell(1).font = { name: '黑体', size: 10 };
        meanRow.getCell(1).alignment = { horizontal: 'left' };
        border(meanRow.getCell(1));
        meanRow.getCell(2).value = totalMean;
        meanRow.getCell(2).font = { name: '黑体', size: 10 };
        meanRow.getCell(2).alignment = { horizontal: 'center' };
        if (totalMean !== null) meanRow.getCell(2).numFmt = '0.0';
        border(meanRow.getCell(2));
        for (let i = 0; i < numGroups; i++) {
          const m = means[orderedGroups[i]];
          const cell = meanRow.getCell(3 + i);
          cell.value = m;
          cell.font = { name: '黑体', size: 10 };
          cell.alignment = { horizontal: 'center' };
          if (m !== null) cell.numFmt = '0.0';
          border(cell);
        }
        // 表2 均值
        meanRow.getCell(table2Start).value = '平均值';
        meanRow.getCell(table2Start).font = { name: '黑体', size: 10 };
        meanRow.getCell(table2Start).alignment = { horizontal: 'left' };
        border(meanRow.getCell(table2Start));
        meanRow.getCell(table2Start + 1).value = totalMean;
        meanRow.getCell(table2Start + 1).font = { name: '黑体', size: 10 };
        meanRow.getCell(table2Start + 1).alignment = { horizontal: 'center' };
        if (totalMean !== null) meanRow.getCell(table2Start + 1).numFmt = '0.0';
        border(meanRow.getCell(table2Start + 1));
        for (let i = 0; i < numGroups; i++) {
          const m = means[orderedGroups[i]];
          const cell = meanRow.getCell(table2Start + 2 + i);
          cell.value = m;
          cell.font = { name: '黑体', size: 10 };
          cell.alignment = { horizontal: 'center' };
          if (m !== null) cell.numFmt = '0.0';
          border(cell);
        }
        // 表3 均值+红绿
        meanRow.getCell(table3Start).value = '平均值';
        meanRow.getCell(table3Start).font = { name: '黑体', size: 10 };
        meanRow.getCell(table3Start).alignment = { horizontal: 'left' };
        border(meanRow.getCell(table3Start));
        meanRow.getCell(table3Start + 1).value = totalMean;
        meanRow.getCell(table3Start + 1).font = { name: '黑体', size: 10 };
        meanRow.getCell(table3Start + 1).alignment = { horizontal: 'center' };
        if (totalMean !== null) meanRow.getCell(table3Start + 1).numFmt = '0.0';
        border(meanRow.getCell(table3Start + 1));
        if (totalMean !== null && Object.keys(validMeans).length >= 2) {
          if (totalMean === maxM) meanRow.getCell(table3Start + 1).fill = RED_FILL;
          else if (totalMean === minM && maxM !== minM) meanRow.getCell(table3Start + 1).fill = GREEN_FILL;
        }
        for (let i = 0; i < numGroups; i++) {
          const g = orderedGroups[i];
          const m = means[g];
          const cell = meanRow.getCell(table3Start + 2 + i);
          cell.value = m;
          cell.font = { name: '黑体', size: 10 };
          cell.alignment = { horizontal: 'center' };
          if (m !== null) cell.numFmt = '0.0';
          border(cell);
          if (m !== null && Object.keys(validMeans).length >= 2) {
            if (m === maxM) cell.fill = RED_FILL;
            else if (m === minM && maxM !== minM) cell.fill = GREEN_FILL;
          }
        }

        worksheet.addRow();

        // 列宽
        worksheet.getColumn(1).width = 18;
        worksheet.getColumn(2).width = 10;
        for (let i = 0; i < numGroups; i++) {
          worksheet.getColumn(3 + i).width = 10;
        }
        worksheet.getColumn(table2Start).width = 18;
        worksheet.getColumn(table2Start + 1).width = 10;
        for (let i = 0; i < numGroups; i++) {
          worksheet.getColumn(table2Start + 2 + i).width = 10;
        }
        worksheet.getColumn(table3Start).width = 18;
        worksheet.getColumn(table3Start + 1).width = 10;
        for (let i = 0; i < numGroups; i++) {
          worksheet.getColumn(table3Start + 2 + i).width = 10;
        }
      }
    }
  }

  return workbook;
}

export async function downloadExcel(workbook, filename) {
  const buffer = await workbook.xlsx.writeBuffer();
  const blob = new Blob([buffer], { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' });
  const url = URL.createObjectURL(blob);
  const a = document.createElement('a');
  a.href = url;
  a.download = filename;
  document.body.appendChild(a);
  a.click();
  document.body.removeChild(a);
  URL.revokeObjectURL(url);
}