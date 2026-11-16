# TableX 表析

问卷数据分析工具。解析 Excel 导出文件，按模块分组生成三个并排表格（计数、百分比、百分比+高亮），导出为 Excel。

## 快速启动

```bash
cd ~/Documents/GitHub/tablex
python3 -m http.server 9000
# 访问 http://localhost:9000/index.html
```

## 技术栈

- 纯 HTML/CSS/JS，无构建步骤，无 package.json
- 解析：SheetJS (`assets/lib/xlsx.full.min.js`)
- 导出：ExcelJS (`assets/lib/exceljs.min.js`)
- 部署：Vercel（`vercel.json` outputDirectory = `.`）

## 项目结构

```
tablex/
├── index.html              # 唯一入口页面
├── assets/
│   ├── js/
│   │   ├── app.js          # 全部核心逻辑
│   │   ├── app.js.bak      # 备份（旧版，可清理）
│   │   └── app.js.fix      # 备份（旧版，可清理）
│   ├── css/styles.css      # 样式（黑白灰极简）
│   └── lib/                # 第三方库（xlsx.full.min.js, exceljs.min.js）
├── vercel.json
├── README.md
└── CLAUDE.md               # 本文件
```

## 数据流

```
上传 Excel → parseExcel() → headers[], rows[]
                                    ↓
                        handleFile() 重置状态 + 保存文件名
                                    ↓
                        initModules() 创建单个默认模块
                                    ↓
                    render() 渲染表头设置 + 题目选择
                                    ↓
              downloadExcel() → generateExcelWorkbook() → 导出 .xlsx
```

## 状态变量

| 变量 | 类型 | 说明 |
|------|------|------|
| `headers[]` | `string[]` | 表头（列名） |
| `rows[]` | `any[][]` | 数据行 |
| `selectedQuestions[]` | `number[]` | 选中的题目列索引 |
| `groupingColIdx` | `number\|null` | 基础表头的列索引 |
| `groups[]` | `object[]` | 已选表头定义 `{name, relation, clauses}`（见「表头系统」） |
| `modules[]` | `object[]` | 模块定义 `{id, name, questionIdxs[]}` |
| `nextModuleId` | `number` | 模块 ID 自增计数器 |
| `uploadedFileName` | `string` | 上传的文件名（不含扩展名） |
| `config` | `object` | 扩展配置 |
| `editingGroupIndex` | `number` | 表头编辑器状态：-1=新建，>=0=编辑 |
| `editingRelation` | `string` | 编辑器中条件组之间的关系（'and'/'or'）|
| `editingClauses` | `array` | 编辑器中条件组列表（空数组 = 编辑器关闭）|
| `isDraggingSelect` | `boolean` | 拖拽选择状态 |
| `dragSelectValue` | `boolean` | 拖拽选择操作类型（选中/取消） |

## 表头系统

### 基础表头
- 选择一个列 → 显示该列所有选项（多选列按英文逗号分割为独立选项，括号内计数与选项数据同口径）
- 点击选项切换选中状态
- 每个选项独立成为一个表头

### 复杂表头
- 支持跨列选择与分组括号组合，可表达 `(A 且 B) 或 (C 且 D)` 等任意组合
- **条件值内 且/或**：同一条件勾选多个值可选 任一（或，默认）/ 全部（且，多选题须同时含所选选项）
- **条件间关系**：同一条件组内各条件可选 且/或（默认 且）
- **条件组之间关系**：可选 且/或（默认 且）
- **允许重叠**：同一行数据可出现在多个表头中（重复计算）
- **条件组配色**：多条件组时每组按黑白灰阶循环着色（左边框 + 浅底色 + 标签同色，6 级灰循环），单条件组保持素色
- **多选匹配**：多选单元格按英文逗号分割后逐项命中条件值（与选项数据同口径），整值命中保留以兼容历史配置

数据结构：

```js
groups[] = [{
  name: string,
  relation: 'and' | 'or',              // 条件组之间
  clauses: [{
    relation: 'and' | 'or',            // 组内条件之间
    conditions: [{ colIdx: number, values: string[], valuesRelation: 'and' | 'or' }]  // 选项之间（默认 or）
  }]
}]
```

`migrateGroup()` 兼容历史格式 `conditions: [[colIdx,val]]` / `[{colIdx,values}]`。
表头自动命名：条件内各值 `+`（或）或 `&`（且）连接，条件间/组间按关系 ` & ` 或 ` | ` 连接，多条件组加括号。

## 核心函数

### 解析 & 分析
| 函数 | 作用 |
|------|------|
| `parseExcel(file)` | 解析 Excel，返回 `{headers, rows}` |
| `countOptions(allRows, colIdx, isMulti)` | 统计选项计数，多选逗号分割 |
| `countUsers(allRows, colIdx)` | 统计有效用户数（非空值） |
| `isMultiChoice(allRows, colIdx)` | 判断多选题（20%+ 含逗号） |
| `groupRowsCustom(allRows, customGroups)` | 按表头分组括号条件过滤行（值内OR，条件/组间按 relation） |
| `migrateGroup(g)` | 历史格式表头定义 → clauses 结构 |
| `safe(v)` | 空值归一化（null/空字符串 → null） |
| `extractScoreValue(val)` | 从选项提取数值分数（如 "5.0-非常满意" → 5.0） |
| `isScoreQuestion(allRows, colIdx)` | 判断是否为量表题（所有选项均为数值） |
| `calcWeightedMean(allRows, colIdx)` | 计算加权平均分 Σ(分数×人数)/总人数 |

### 表头编辑器
| 函数 | 作用 |
|------|------|
| `handleGroupColChange(e)` | 基础表头列选择 |
| `toggleGroupValue(val)` | 基础表头选项切换 |
| `openGroupEditor(groupIndex)` | 打开复杂表头编辑器 |
| `cancelGroupEditor()` | 取消编辑 |
| `addClause()` / `removeClause(clauseIdx)` | 添加/删除条件组 |
| `setTopRelation(rel)` / `setClauseRelation(clauseIdx, rel)` | 组间 / 组内 且或关系 |
| `addCondition(clauseIdx)` / `removeCondition(clauseIdx, condIdx)` | 添加/删除条件行 |
| `handleConditionColChange(clauseIdx, condIdx, colIdx)` | 条件列选择变更 |
| `toggleConditionValue(clauseIdx, condIdx, val)` | 切换条件选项选中状态 |
| `setConditionValuesRelation(clauseIdx, condIdx, rel)` | 选项之间 且/或（多选题可选全部命中）|
| `effectiveClauses()` / `countEditingMatches()` | 取有效条件 / 实时命中人数（保存按钮旁显示） |
| `saveGroup()` | 保存复杂表头 |
| `deleteGroup(groupIndex)` | 删除表头 |

### 拖拽选择
| 函数 | 作用 |
|------|------|
| `startDragSelect(event, idx)` | 开始拖拽选择 |
| `handleDragSelect(idx)` | 拖拽过程中应用选择 |
| `setQuestionSelected(idx, selected)` | 设置题目选中状态 |
| `endDragSelect()` | 结束拖拽选择 |

### 模块管理
| 函数 | 作用 |
|------|------|
| `initModules()` | 创建单个默认模块（不自动分类） |
| `addModuleAfterQuestion(idx)` | 在指定题目后插入新模块 |
| `moveQuestionToModule(questionIdx, moduleId)` | 移动题目到指定模块 |
| `renameModule(moduleId, newName)` | 修改模块名称 |
| `deleteModule(moduleId)` | 删除模块 |

### Excel 导出
| 函数 | 作用 |
|------|------|
| `generateExcelWorkbook()` | 生成 ExcelJS Workbook |
| `mixFill(t)` | 红→绿线性插值填充色（t=0 红端点，t=1 绿端点）|
| `downloadExcel()` | 触发下载（文件名：`上传文件名_MMddHHmm.xlsx`） |

### 配置持久化
| 函数 | 作用 |
|------|------|
| `saveConfig()` | 保存到 localStorage |
| `loadSavedConfig()` | 从 localStorage 恢复（兼容新旧格式） |
| `resetConfig()` | 重置为默认 |
| `exportConfig()` / `importConfig(e)` | 导出/导入 JSON 配置 |

## Excel 输出格式

### Sheet 结构
每个模块 → 一个 Sheet（目前只有一个"题目"模块）

### 每个题目的布局
3 个并排表格，中间 spacer 列分隔：

| 表格 | 内容 | 高亮 |
|------|------|------|
| Table 1 | 计数（次数） | 无 |
| Table 2 | 百分比 | 无 |
| Table 3 | 百分比 | 红→绿色阶：最低=红端点 `#D96459`，最高=绿端点 `#6AA84F`，中间按数值线性插值（max=min 时不上色）|

Table 3 色阶只作用于分组列：**Total 列不上色、不参与极值**（无表头时无着色单元格），有表头时按行（选项）在分组列之间归一化。两端点色为全局常量（`mixFill(0)`/`mixFill(1)` 精确等于端点色），全表统一。注意：界面为黑白灰极简，但 Excel 导出的色阶保持红→绿梯度。

布局：`选项 | Total | 分组1 | 分组2 | ...` × 3

### 行结构
1. **标题行**：题干 + 题目信息（单选/多选）
2. **表头行**：选项、Total、各分组名（底色 `#DEE0E3` 灰色）
3. **base 行**：有效用户数（斜体，棕色文字）
4. **平均分行**（仅量表题）：加权平均分（暖琥珀底色 `#FFF3E0`，加粗，两位小数）
5. **数据行**：各选项的计数或百分比

### 排序规则
- **题目顺序**：按 Excel 原始列顺序
- **选项顺序**：按 Total 的选择人数倒序排列
- **表头顺序**：按用户添加/拖拽排序的顺序
- **选项完整性**：即使某个表头没有选择某选项，只要 Total 有，也会显示

### 文件命名
`上传文件名_MMddHHmm.xlsx`

示例：`问卷数据_07241420.xlsx`（7月24日14点20分）

## UI 设计

### 色彩方案（黑白灰极简）
- 背景：`#F7F7F7`（浅灰）
- 卡片：`#FFFFFF`（纯白）
- 边框：`#E2E2E2`（浅灰）
- 文字：`#1A1A1A`（近黑）
- 强调色：`#222222`（黑）

### 交互细节
- 模块名称可直接点击编辑
- 每个题目末尾有 `+` 按钮添加新模块
- 表头支持拖拽排序
- 题目支持滑动批量选择（按住鼠标滑过）
- 已选表头显示样本数量
- 编辑/删除按钮为白色样式

## localStorage

- key: `survey_analysis_config`
- 存储: `{groups, selectedQuestions, groupingColIdx, config, modules}`
- 兼容新旧格式的 groups 结构

## 注意事项

- 无测试框架，无测试文件
- `assets/js/` 下有 `app.js.bak` 和 `app.js.fix` 两个备份文件，可安全删除
- 第三方库直接 vendored，不通过 npm 管理
