# TableX 表析

问卷数据分析工具，解析 Excel 导出文件，生成三个并排表格（计数、百分比无高亮、百分比有高亮）。

## 快速启动

```bash
cd /Users/yuanyi.li/Documents/GitHub/tablex
python3 -m http.server 9000
# 访问 http://localhost:9000/index.html
```

## 技术栈

- 纯 HTML/CSS/JS，无构建步骤
- 解析：SheetJS (xlsx.full.min.js) — 本地 `assets/lib/`
- 导出：ExcelJS — 本地 `assets/lib/`
- 部署：Vercel（vercel.json 配置 outputDirectory = .）

## 项目结构

```
tablex/
├── index.html          # 主入口页面
├── tablex.html         # 单文件独立版（所有 CSS/JS 内联）
├── assets/
│   ├── js/app.js       # 核心逻辑
│   ├── css/styles.css  # 样式
│   └── lib/            # 第三方库（xlsx.full.min.js, exceljs.min.js）
├── debug_xlsx.js       # xlsx 调试脚本（node debug_xlsx.js 运行）
└── CLAUDE.md           # 本文件
```

## 主要功能

### 上传与解析
- 拖拽或点击上传 Excel 文件（.xlsx, .xls）
- 上传后显示文件名、行数、列数
- 自动识别单选/多选题（逗号分割判断多选）

### 分组设置
- 选择分组列后，列出所有分组值及样本数量
- 点击选择分组，支持拖拽排序已选分组
- 分组配置自动保存到 localStorage

### 题目选择与模块分组
- 按原始顺序展示，自动按关键词分类到模块（用户画像、购车意向、智能驾驶、家庭用车、产品体验、其他）
- 模块数量上限 10 个，超出自动合并
- 每个题目下方有 `+` 按钮，可在该题后添加新模块
- 模块头部有 `×` 按钮可删除模块（题目合并到相邻模块）
- 模块名称可直接点击修改
- MUJI 风格模块背景（浅米色 #E8E4DC + 深色文字 #3D3D3D）
- 单选/多选标签统一显示在题目末尾
- 题目顺序始终保持不变，不受模块操作影响
- 全选/全不选
- 下载 Excel 按钮在题目选择区

### Excel 输出格式

- 每个模块生成一个 Sheet，Sheet 名称与模块名称一致
- 每个题目输出 3 个并排表格（含 spacer 空列）：

| 表格 | 内容 | 高亮 |
|------|------|------|
| Table 1 | 计数（次数） | 无 |
| Table 2 | 百分比 | 无 |
| Table 3 | 百分比 | 绿=最高，红=最低（仅当 maxPct !== minPct）|

- 无平均值行
- base 行斜体
- 题干和题目分两行，3 个 table 列对齐

### 分组输出格式
- 3 个表格横向排列：Table 1（计数）| spacer | Table 2（百分比）| spacer | Table 3（百分比+高亮）
- 每个表格内：选项 | Total | 分组1 | 分组2 | ...

## 关键函数

| 函数 | 作用 |
|------|------|
| `parseExcel(file)` | 解析 Excel 文件，返回 headers 和 rows |
| `countOptions(allRows, colIdx, isMulti)` | 统计选项计数，逗号分割处理 |
| `isMultiChoice(allRows, colIdx)` | 判断是否多选题（20%以上含逗号）|
| `countUsers(allRows, colIdx)` | 统计有效用户数 |
| `generateExcelWorkbook()` | 生成 ExcelJS Workbook |
| `handleFile(file)` | 上传文件后重置状态并解析 |
| `loadSavedConfig()` | 从 localStorage 恢复分组配置 |
| `initModules()` | 自动分类题目到模块（最多10个）|
| `addModuleAfterQuestion(idx)` | 在指定题目后添加新模块 |
| `deleteModule(moduleId)` | 删除模块，题目合并到相邻模块 |
| `renameModule(moduleId, newName)` | 修改模块名称 |

## 状态管理

- `headers[]` — 表头
- `rows[]` — 数据行
- `selectedQuestions[]` — 选中的题目索引
- `groupingColIdx` — 分组列索引
- `groups[]` — 已选分组定义
- `config` — 扩展配置（保留字段）
- `modules[]` — 模块定义，每个模块包含 `{id, name, questionIdxs[]}`

## localStorage

- `survey_analysis_config` — 保存配置，key 为 `groups`, `selectedQuestions`, `groupingColIdx`, `config`, `modules`
- 上传新文件时清除旧配置，重新加载时验证列索引有效性
