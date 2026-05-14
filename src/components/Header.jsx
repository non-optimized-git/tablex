import { useStore } from '../store/useStore.js';
import { exportConfig, importConfig } from '../lib/configManager.js';
import { groupRowsCustom } from '../lib/analyzer.js';
import { generateExcel, downloadExcel } from '../lib/excelWriter.js';

export default function Header() {
  const { headers, rows, config, selectedQuestions, setConfig, resetConfig } = useStore();

  const canDownload = headers.length > 0 && rows.length > 0 && selectedQuestions.length > 0;

  const handleDownload = async () => {
    if (!canDownload) return;
    const groups = config.groups?.length > 0 ? groupRowsCustom(rows, config.groups) : {};
    const orderedGroups = config.groups?.map(g => g.name) || [];
    const workbook = await generateExcel(selectedQuestions, headers, rows, groups, orderedGroups, config);
    await downloadExcel(workbook, '分析结果.xlsx');
  };

  const handleExport = () => exportConfig(config);

  const handleImport = async (e) => {
    const file = e.target.files?.[0];
    if (!file) return;
    try {
      const imported = await importConfig(file);
      setConfig(imported);
      alert('配置已加载');
    } catch (err) {
      alert('加载失败: ' + err.message);
    }
    e.target.value = '';
  };

  return (
    <header className="bg-white border-b border-gray-200 px-6 py-5">
      <div className="flex items-center justify-between">
        <div className="flex items-center gap-4">
          <div className="w-12 h-12 bg-primary rounded-xl flex items-center justify-center shadow-soft">
            <svg className="w-6 h-6 text-white" fill="none" viewBox="0 0 24 24" strokeWidth={2} stroke="currentColor">
              <path strokeLinecap="round" strokeLinejoin="round" d="M9 19v-6a2 2 0 00-2-2H5a2 2 0 00-2 2v6a2 2 0 002 2h2a2 2 0 002-2zm0 0V9a2 2 0 012-2h2a2 2 0 012 2v10m-6 0a2 2 0 002 2h2a2 2 0 002-2m0 0V5a2 2 0 012-2h2a2 2 0 012 2v14a2 2 0 01-2 2h-2a2 2 0 01-2-2z" />
            </svg>
          </div>
          <div>
            <h1 className="text-xl font-semibold text-gray-900">问卷分析工具</h1>
            <p className="text-sm text-gray-500 mt-0.5">Survey Analysis Tool</p>
          </div>
        </div>

        <div className="flex items-center gap-3">
          <label className="px-4 py-2 text-sm bg-white border border-muji-lightGray text-muji-charcoal rounded cursor-pointer hover:border-muji-beige transition-colors">
            导入配置
            <input type="file" accept=".json" onChange={handleImport} className="hidden" />
          </label>
          <button
            onClick={handleExport}
            className="px-4 py-2 text-sm bg-muji-beige text-muji-charcoal rounded hover:bg-muji-lightGray transition-colors"
          >
            导出配置
          </button>
          <button
            onClick={resetConfig}
            className="px-4 py-2 text-sm bg-white border border-muji-lightGray text-muji-gray rounded hover:border-muji-beige transition-colors"
          >
            重置
          </button>
        </div>
      </div>
    </header>
  );
}