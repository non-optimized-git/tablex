// 严格对应 Python analyze_survey.py 的统计逻辑

function safe(v) {
  if (v === null || v === undefined) return null;
  const s = String(v).trim();
  return s || null;
}

function splitOpts(val) {
  if (!val) return [];
  return String(val).split(',').map(o => o.trim()).filter(o => o);
}

// group_rows_custom - AND 关系分组
export function groupRowsCustom(rows, customGroups) {
  const groups = {};
  for (const groupDef of customGroups) {
    const groupName = groupDef.name;
    const conditions = groupDef.conditions;
    groups[groupName] = [];
    for (const row of rows) {
      let match = true;
      for (const [colIdx, expectedVal] of conditions) {
        if (safe(row[colIdx]) !== expectedVal) {
          match = false;
          break;
        }
      }
      if (match) {
        groups[groupName].push(row);
      }
    }
  }
  return groups;
}

// count_options - 多选题按逗号分割
export function countOptions(rows, colIdx, isMulti) {
  const counts = {};
  for (const row of rows) {
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

// count_users - 非空作答人数
export function countUsers(rows, colIdx) {
  return rows.filter(row => safe(row[colIdx]) !== null).length;
}

// calc_mean - 加权均值（精确按 value_map）
export function calcMean(rows, colIdx, valueMap) {
  let totalWeight = 0;
  let totalValue = 0;
  for (const row of rows) {
    let val = safe(row[colIdx]);
    if (val === null) continue;
    // 多选题只取第一个选项
    if (String(val).includes(',')) {
      const opts = splitOpts(val);
      val = opts[0] || null;
      if (val === null) continue;
    }
    for (const [key, mval] of valueMap) {
      if (val.includes(key)) {
        totalWeight++;
        totalValue += mval;
        break;
      }
    }
  }
  return totalWeight > 0 ? totalValue / totalWeight : null;
}

// is_multi_choice - 20% 阈值
export function isMultiChoice(rows, colIdx) {
  let total = 0, hasComma = 0;
  for (const row of rows) {
    const val = row[colIdx];
    if (val === null || val === undefined) continue;
    total++;
    if (String(val).includes(',')) hasComma++;
  }
  return total > 0 && hasComma > total * 0.2;
}

// 从 ordered 自动生成中间值（用于 calc_mean）
function buildValueMapFromOrdered(optOrder) {
  const valueMap = [];
  for (const opt of optOrder) {
    let m = opt.match(/^(\d+)岁以下/);
    if (m) { valueMap.push([opt, parseInt(m[1])]); continue; }
    m = opt.match(/^(\d+)岁以上/);
    if (m) { valueMap.push([opt, parseInt(m[1])]); continue; }
    m = opt.match(/^(\d+)-(\d+)/);
    if (m) { valueMap.push([opt, (parseInt(m[1]) + parseInt(m[2])) / 2]); continue; }
    m = opt.match(/^(\d+)$/);
    if (m) { valueMap.push([opt, parseFloat(m[1])]); continue; }
  }
  return valueMap;
}

// auto_calc_mean - 从 ordered 配置自动计算均值
export function autoCalcMean(rows, colIdx, orderedConfig) {
  if (!orderedConfig || !orderedConfig.length) return null;
  const valueMap = buildValueMapFromOrdered(orderedConfig);
  if (!valueMap.length) return null;
  return calcMean(rows, colIdx, valueMap);
}

// 按频率排序（非数字选项）
function sortByFrequency(allOpts, groupCounts) {
  const sorted = [];
  const counter = {};
  for (const counts of groupCounts) {
    for (const [opt, count] of Object.entries(counts)) {
      counter[opt] = (counter[opt] || 0) + count;
    }
  }
  return allOpts.sort((a, b) => (counter[b] || 0) - (counter[a] || 0));
}

// 提取数字选项的起始数字（用于排序）
function extractNumKey(s) {
  const m = String(s).match(/^(\d+)/);
  return m ? parseInt(m[1]) : 0;
}

// sort_options - 数字选项按自然顺序，非数字选项按频率
export function sortOptions(allOpts, colIdx, groupCounts, orderedConfig) {
  const numericOpts = allOpts.filter(o => /\d/.test(String(o)));
  const nonNumericOpts = allOpts.filter(o => !/\d/.test(String(o)));

  const sorted = [];
  // 数字选项：按 ordered 顺序或起始数字排序
  if (orderedConfig && orderedConfig.length) {
    sorted.push(...numericOpts.sort((a, b) => {
      const ai = orderedConfig.indexOf(a);
      const bi = orderedConfig.indexOf(b);
      if (ai !== -1 && bi !== -1) return ai - bi;
      if (ai !== -1) return -1;
      if (bi !== -1) return 1;
      return extractNumKey(a) - extractNumKey(b);
    }));
  } else {
    sorted.push(...numericOpts.sort((a, b) => extractNumKey(a) - extractNumKey(b)));
  }
  // 非数字选项：按频率降序
  if (nonNumericOpts.length) {
    sorted.push(...sortByFrequency(nonNumericOpts, groupCounts));
  }
  return sorted;
}

// 统计并返回完整结果（用于预览）
export function analyzeQuestions(questions, headers, rows, groups, orderedGroups, config) {
  const results = [];
  for (const colIdx of questions) {
    const header = headers[colIdx] || `列${colIdx}`;
    const isMulti = isMultiChoice(rows, colIdx);

    if (groups && Object.keys(groups).length > 0) {
      const groupCounts = {};
      const groupUsers = {};
      for (const g of orderedGroups) {
        groupCounts[g] = countOptions(groups[g], colIdx, isMulti);
        groupUsers[g] = countUsers(groups[g], colIdx);
      }
      const allOptsSet = new Set();
      for (const counts of Object.values(groupCounts)) {
        for (const opt of Object.keys(counts)) allOptsSet.add(opt);
      }
      const allOpts = sortOptions([...allOptsSet], colIdx, Object.values(groupCounts), config.ordered?.[colIdx]);
      const pcts = {};
      for (const g of orderedGroups) {
        pcts[g] = {};
        for (const opt of allOpts) {
          const val = groupCounts[g][opt] || 0;
          const total = groupUsers[g] || 0;
          pcts[g][opt] = total > 0 ? val / total : 0;
        }
      }
      // 计算均值
      const means = {};
      for (const g of orderedGroups) {
        if (config.mean_cols?.[colIdx]) {
          means[g] = calcMean(groups[g], colIdx, config.mean_cols[colIdx].values);
        } else if (config.ordered?.[colIdx]) {
          means[g] = autoCalcMean(groups[g], colIdx, config.ordered[colIdx]);
        } else {
          means[g] = null;
        }
      }
      results.push({ colIdx, header, isMulti, allOpts, groupCounts, groupUsers, pcts, means });
    } else {
      const counts = countOptions(rows, colIdx, isMulti);
      const totalUsers = countUsers(rows, colIdx);
      const allOpts = sortOptions(Object.keys(counts), colIdx, [counts], null);
      let meanVal = null;
      if (config.mean_cols?.[colIdx]) {
        meanVal = calcMean(rows, colIdx, config.mean_cols[colIdx].values);
      }
      results.push({ colIdx, header, isMulti, counts, totalUsers, allOpts, meanVal });
    }
  }
  return results;
}