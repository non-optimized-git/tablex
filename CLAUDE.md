# TableX 表析

问卷数据分析工具。解析 Excel 导出文件，按模块分组生成三个并排表格（计数、百分比、百分比+高亮），导出为 Excel。

## 快速启动

```bash
cd /Users/yuanyi.li/Documents/GitHub/tablex
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
├── index.html              # 唯一入口页面（76 行）
├── assets/
│   ├── js/
│   │   ├── app.js          # 全部核心逻辑（949 行）
│   │   ├── app.js.bak      # 备份（旧版，可清理）
│   │   └── app.js.fix      # 备份（旧版，可清理）
│   ├── css/styles.css      # 样式（398 行，MUJI 风格）
│   └── lib/                # 第三方库（xlsx.full.min.js, exceljs.min.js）
├── vercel.json
├── README.md
└── CLAUDE.md               # 本文件
```

## 数据流

```
上传 Excel → parseExcel() → headers[], rows[]
                                    ↓
                            handleFile() 重置状态
                                    ↓
                        initModules() 自动按关键词分模块
                                    ↓
                    render() 渲染分组设置 + 题目选择（模块化展示）
                                    ↓
              downloadExcel() → generateExcelWorkbook() → 导出 .xlsx
```

## 状态变量

| 变量 | 类型 | 说明 |
|------|------|------|
| `headers[]` | `string[]` | 表头（列名） |
| `rows[]` | `any[][]` | 数据行 |
| `selectedQuestions[]` | `number[]` | 选中的题目列索引 |
| `groupingColIdx` | `number\|null` | 分组列索引 |
| `groups[]` | `object[]` | 已选分组定义 `{name, conditions}` |
| `modules[]` | `object[]` | 模块定义 `{id, name, questionIdxs[]}` |
| `nextModuleId` | `number` | 模块 ID 自增计数器 |
| `config` | `object` | 扩展配置 `{groups, ordered, mean_cols, multi_idx}` |

## 核心函数

### 解析 & 分析
| 函数 | 作用 |
|------|------|
| `parseExcel(file)` | 解析 Excel，返回 `{headers, rows}` |
| `countOptions(allRows, colIdx, isMulti)` | 统计选项计数，多选逗号分割 |
| `countUsers(allRows, colIdx)` | 统计有效用户数（非空值） |
| `isMultiChoice(allRows, colIdx)` | 判断多选题（20%+ 含逗号） |
| `groupRowsCustom(allRows, customGroups)` | 按分组条件过滤行 |
| `safe(v)` | 空值归一化（null/空字符串 → null） |

### 模块管理
| 函数 | 作用 |
|------|------|
| `initModules()` | 自动按关键词分类题目到模块（最多 10 个） |
| `detectModule(header)` | 根据关键词匹配模块名 |
| `addModuleAfterQuestion(idx)` | 在指定题目后插入新模块 |
| `moveQuestionToModule(questionIdx, moduleId)` | 移动题目到指定模块 |
| `renameModule(moduleId, newName)` | 修改模块名称 |
| `deleteModule(moduleId)` | 删除模块，题目合并到上一个模块 |

### Excel 导出
| 函数 | 作用 |
|------|------|
| `generateExcelWorkbook()` | 生成 ExcelJS Workbook（主逻辑） |
| `downloadExcel()` | 触发下载 |
| `makeBorder()` / `makeHeaderStyle()` / `makeDataStyle()` / `makeFillStyle()` | 样式工厂 |
| `r3(x)` | 四舍五入保留 3 位小数 |

### UI 渲染
| 函数 | 作用 |
|------|------|
| `render()` | 主渲染函数（分组设置 + 题目列表） |
| `renderSelectedGroups()` | 渲染已选分组（支持拖拽排序） |
| `handleFile(file)` | 上传后重置状态并解析 |
| `handleGroupColChange(e)` | 分组列切换 |
| `toggleGroupValue(val)` | 切换分组选中状态 |
| `getShortTitle(header)` | 提取题目短标题 |

### 配置持久化
| 函数 | 作用 |
|------|------|
| `saveConfig()` | 保存到 localStorage（key: `survey_analysis_config`） |
| `loadSavedConfig()` | 从 localStorage 恢复 |
| `resetConfig()` | 重置为默认 |
| `exportConfig()` | 导出配置为 JSON 文件 |
| `importConfig(e)` | 从 JSON 文件导入配置 |

## Excel 输出格式

每个模块 → 一个 Sheet，每个题目输出 3 个并排表格：

| 表格 | 内容 | 高亮 |
|------|------|------|
| Table 1 | 计数（次数） | 无 |
| Table 2 | 百分比 | 无 |
| Table 3 | 百分比 | 绿=最高，红=最低（仅当 maxPct ≠ minPct）|

布局：`选项 | Total | 分组1 | 分组2 | ...` × 3 个 table，中间 spacer 列分隔。

## 关键 UI 细节

- 模块背景：浅米色 `#E8E4DC`，文字 `#3D3D3D`（MUJI 风格）
- 模块名称可直接点击编辑（`<input>` + `onchange`）
- 每个题目末尾有 `+` 按钮添加新模块，模块头有 `×` 删除
- 单选/多选标签显示在题目末尾
- 题目顺序始终按原始列顺序，不受模块操作影响
- 分组支持拖拽排序
- 配置自动保存到 localStorage，上传新文件时清除旧配置

## localStorage

- key: `survey_analysis_config`
- 存储: `{groups, selectedQuestions, groupingColIdx, config, modules}`
- 重新加载时验证列索引有效性

## 注意事项

- 无测试框架，无测试文件
- `assets/js/` 下有 `app.js.bak` 和 `app.js.fix` 两个备份文件，是重构遗留，可安全删除
- 第三方库直接 vendored，不通过 npm 管理
