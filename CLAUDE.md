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
│   ├── css/styles.css      # 样式（暖色调）
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
| `groups[]` | `object[]` | 已选表头定义 `{name, conditions: [{colIdx, values}]}` |
| `modules[]` | `object[]` | 模块定义 `{id, name, questionIdxs[]}` |
| `nextModuleId` | `number` | 模块 ID 自增计数器 |
| `uploadedFileName` | `string` | 上传的文件名（不含扩展名） |
| `config` | `object` | 扩展配置 |
| `editingGroupIndex` | `number` | 表头编辑器状态：-1=新建，>=0=编辑 |
| `editingConditions` | `array` | 表头编辑器当前条件列表 |
| `isDraggingSelect` | `boolean` | 拖拽选择状态 |
| `dragSelectValue` | `boolean` | 拖拽选择操作类型（选中/取消） |

## 表头系统

### 基础表头
- 选择一个列 → 显示该列所有唯一值
- 点击选项切换选中状态
- 每个选项独立成为一个表头

### 复杂表头
- 支持跨列选择
- **列内 OR**：同一列选择多个选项，满足任一即可
- **列间 AND**：不同列的条件必须同时满足
- **允许重叠**：同一行数据可出现在多个表头中（重复计算）

数据结构：`groups[].conditions = [{colIdx: number, values: string[]}, ...]`

## 核心函数

### 解析 & 分析
| 函数 | 作用 |
|------|------|
| `parseExcel(file)` | 解析 Excel，返回 `{headers, rows}` |
| `countOptions(allRows, colIdx, isMulti)` | 统计选项计数，多选逗号分割 |
| `countUsers(allRows, colIdx)` | 统计有效用户数（非空值） |
| `isMultiChoice(allRows, colIdx)` | 判断多选题（20%+ 含逗号） |
| `groupRowsCustom(allRows, customGroups)` | 按表头条件过滤行（列内OR，列间AND） |
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
| `addCondition()` / `removeCondition()` | 添加/删除条件行 |
| `handleConditionColChange()` | 条件列选择变更 |
| `toggleConditionValue()` | 切换条件选项选中状态 |
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
| Table 3 | 百分比 | 绿=最高，红=最低（仅当 maxPct ≠ minPct）|

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

### 色彩方案（暖色调）
- 背景：`#FAF7F2`（米色）
- 卡片：`#FFFDF9`（微暖白）
- 边框：`#E8E0D4`（暖灰）
- 文字：`#3E2723`（深棕）
- 强调色：`#5C4033`（深棕）

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
