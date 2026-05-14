import { useMemo } from 'react';
import { useStore } from '../store/useStore.js';
import { analyzeQuestions, groupRowsCustom } from '../lib/analyzer.js';
import { getShortTitle } from '../lib/constants.js';
import { generateExcel, downloadExcel } from '../lib/excelWriter.js';

function r3(x) {
  if (x === null || x === undefined) return null;
  return Math.abs(x) < 0.0005 ? 0 : Math.round(x * 1000) / 1000;
}

export default function ResultPreview() {
  const { headers, rows, config, selectedQuestions } = useStore();

  const canDownload = headers.length > 0 && rows.length > 0 && selectedQuestions.length > 0;

  const handleDownload = async () => {
    if (!canDownload) return;
    const groups = config.groups?.length > 0 ? groupRowsCustom(rows, config.groups) : {};
    const orderedGroups = config.groups?.map(g => g.name) || [];
    const workbook = await generateExcel(selectedQuestions, headers, rows, groups, orderedGroups, config);
    await downloadExcel(workbook, '分析结果.xlsx');
  };

  const results = useMemo(() => {
    if (!headers.length || !rows.length || !selectedQuestions.length) return null;

    const groups = config.groups?.length > 0
      ? groupRowsCustom(rows, config.groups)
      : {};
    const orderedGroups = config.groups?.map(g => g.name) || [];

    return analyzeQuestions(selectedQuestions, headers, rows, groups, orderedGroups, config);
  }, [headers, rows, config, selectedQuestions]);

  if (!results) {
    return (
      <div className="bg-muji-warmWhite border border-muji-lightGray rounded-lg p-6 text-center text-muji-gray">
        请先上传 Excel 文件并选择题目
      </div>
    );
  }

  return (
    <div className="bg-muji-warmWhite border border-muji-lightGray rounded-lg p-4">
      <div className="flex justify-between items-center mb-3">
        <h3 className="text-sm font-medium text-muji-charcoal">结果预览</h3>
        <button
          onClick={handleDownload}
          disabled={!canDownload}
          className={`px-4 py-2 text-sm font-medium rounded transition-colors ${
            canDownload
              ? 'bg-primary text-white hover:bg-primary-hover'
              : 'bg-muji-lightGray text-muji-gray cursor-not-allowed'
          }`}
        >
          下载 Excel
        </button>
      </div>
      <div className="space-y-4 max-h-[500px] overflow-y-auto">
        {results.map((res) => {
          const title = getShortTitle(res.header);
          if (!res.allOpts?.length) return null;

          return (
            <div key={res.colIdx} className="border border-muji-lightGray rounded p-3">
              <h4 className="text-sm font-medium text-muji-charcoal mb-2">
                {title} ({res.isMulti ? '多选' : '单选'})
              </h4>

              {res.groupCounts ? (
                // 分组模式
                <div className="overflow-x-auto">
                  <table className="w-full text-xs">
                    <thead>
                      <tr>
                        <th className="text-left py-1 text-muji-gray">选项</th>
                        {res.groupUsers && Object.keys(res.groupCounts).map(g => (
                          <th key={g} className="py-1 px-2 text-muji-gray text-center">{g}</th>
                        ))}
                      </tr>
                    </thead>
                    <tbody>
                      <tr className="bg-muji-cream">
                        <td className="py-1 text-muji-gray">base</td>
                        {res.groupUsers && Object.entries(res.groupUsers).map(([g, n]) => (
                          <td key={g} className="py-1 px-2 text-center text-muji-charcoal">{n}</td>
                        ))}
                      </tr>
                      {res.allOpts.map(opt => (
                        <tr key={opt}>
                          <td className="py-1 text-muji-charcoal">{opt}</td>
                          {res.pcts && Object.entries(res.pcts).map(([g, optPcts]) => (
                            <td key={g} className="py-1 px-2 text-center text-muji-charcoal">
                              {r3(optPcts[opt] * 100)}%
                            </td>
                          ))}
                        </tr>
                      ))}
                      {res.means && Object.values(res.means).some(m => m !== null) && (
                        <tr className="border-t border-muji-lightGray">
                          <td className="py-1 text-muji-gray font-medium">平均值</td>
                          {Object.entries(res.means).map(([g, m]) => (
                            <td key={g} className="py-1 px-2 text-center text-muji-charcoal">
                              {m !== null ? r3(m) : '-'}
                            </td>
                          ))}
                        </tr>
                      )}
                    </tbody>
                  </table>
                </div>
              ) : (
                // 无分组模式
                <table className="w-full text-xs">
                  <thead>
                    <tr>
                      <th className="text-left py-1 text-muji-gray">选项</th>
                      <th className="py-1 px-2 text-muji-gray text-center">次数</th>
                      <th className="py-1 px-2 text-muji-gray text-center">占比</th>
                    </tr>
                  </thead>
                  <tbody>
                    <tr className="bg-muji-cream">
                      <td className="py-1 text-muji-gray">base</td>
                      <td className="py-1 px-2 text-center text-muji-charcoal">{res.totalUsers}</td>
                      <td className="py-1 px-2 text-center text-muji-charcoal">100%</td>
                    </tr>
                    {res.allOpts.map(opt => (
                      <tr key={opt}>
                        <td className="py-1 text-muji-charcoal">{opt}</td>
                        <td className="py-1 px-2 text-center text-muji-charcoal">{res.counts[opt]}</td>
                        <td className="py-1 px-2 text-center text-muji-charcoal">
                          {r3((res.counts[opt] / res.totalUsers) * 100)}%
                        </td>
                      </tr>
                    ))}
                    {res.meanVal !== null && (
                      <tr className="border-t border-muji-lightGray">
                        <td className="py-1 text-muji-gray font-medium">平均值</td>
                        <td colSpan={2} className="py-1 px-2 text-center text-muji-charcoal">{r3(res.meanVal)}</td>
                      </tr>
                    )}
                  </tbody>
                </table>
              )}
            </div>
          );
        })}
      </div>
    </div>
  );
}