import { useCallback } from 'react';
import { useStore } from '../store/useStore.js';
import { parseExcel } from '../lib/excelParser.js';

export default function FileUploader() {
  const { setHeaders, setRows, setSelectedQuestions } = useStore();

  const handleFile = useCallback(async (e) => {
    const file = e.target.files?.[0];
    if (!file) return;

    try {
      const { headers, rows } = await parseExcel(file);
      setHeaders(headers);
      setRows(rows);
      setSelectedQuestions(headers.map((_, i) => i));
    } catch (err) {
      alert('读取文件失败: ' + err.message);
    }
  }, [setHeaders, setRows, setSelectedQuestions]);

  return (
    <div className="bg-muji-warmWhite border border-muji-lightGray rounded-lg p-6">
      <label className="block">
        <span className="text-sm text-muji-gray mb-2 block">上传 Excel 文件</span>
        <input
          type="file"
          accept=".xlsx,.xls"
          onChange={handleFile}
          className="w-full text-sm text-muji-charcoal file:mr-4 file:py-2 file:px-4 file:rounded file:border-0 file:bg-muji-beige file:text-muji-charcoal file:cursor-pointer hover:file:bg-muji-lightGray file:transition-colors file:font-medium"
        />
      </label>
    </div>
  );
}