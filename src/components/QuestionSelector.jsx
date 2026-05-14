import { useMemo } from 'react';
import { useStore } from '../store/useStore.js';
import { detectModule, getShortTitle } from '../lib/constants.js';
import { isMultiChoice } from '../lib/analyzer.js';

export default function QuestionSelector() {
  const { headers, rows, selectedQuestions, setSelectedQuestions } = useStore();

  // 按模块分组
  const modules = useMemo(() => {
    if (!headers.length) return {};
    const mods = {};
    headers.forEach((h, i) => {
      const mod = detectModule(h);
      const title = getShortTitle(h) || `列${i + 1}`;
      const isMulti = rows.length > 0 ? isMultiChoice(rows, i) : false;
      if (!mods[mod]) mods[mod] = [];
      mods[mod].push({ idx: i, title, isMulti });
    });
    return mods;
  }, [headers, rows]);

  const toggleQuestion = (idx) => {
    if (selectedQuestions.includes(idx)) {
      setSelectedQuestions(selectedQuestions.filter(i => i !== idx));
    } else {
      setSelectedQuestions([...selectedQuestions, idx]);
    }
  };

  const selectAll = () => setSelectedQuestions(headers.map((_, i) => i));
  const selectNone = () => setSelectedQuestions([]);

  if (!headers.length) return null;

  return (
    <div className="bg-muji-warmWhite border border-muji-lightGray rounded-lg p-4">
      <div className="flex justify-between items-center mb-3">
        <h3 className="text-sm font-medium text-muji-charcoal">题目选择</h3>
        <div className="flex gap-2">
          <button
            onClick={selectAll}
            className="text-xs text-muji-gray hover:text-muji-charcoal"
          >
            全选
          </button>
          <span className="text-muji-lightGray">|</span>
          <button
            onClick={selectNone}
            className="text-xs text-muji-gray hover:text-muji-charcoal"
          >
            全不选
          </button>
        </div>
      </div>

      <p className="text-xs text-muji-gray mb-3">
        已选择 {selectedQuestions.length} / {headers.length} 题
      </p>

      <div className="space-y-3 max-h-96 overflow-y-auto">
        {Object.entries(modules).map(([modName, questions]) => (
          <div key={modName}>
            <h4 className="text-xs font-medium text-muji-gray mb-2">{modName}</h4>
            <div className="space-y-1 pl-2">
              {questions.map(q => {
                const isSelected = selectedQuestions.includes(q.idx);
                return (
                  <label
                    key={q.idx}
                    className={`flex items-center gap-2 p-2 rounded cursor-pointer transition-colors ${
                      isSelected ? 'bg-muji-cream' : 'hover:bg-muji-cream'
                    }`}
                  >
                    <input
                      type="checkbox"
                      checked={isSelected}
                      onChange={() => toggleQuestion(q.idx)}
                      className="w-4 h-4 accent-muji-beige"
                    />
                    <span className="text-sm text-muji-charcoal">{q.title}</span>
                    <span className="text-xs text-muji-gray">({q.isMulti ? '多选' : '单选'})</span>
                  </label>
                );
              })}
            </div>
          </div>
        ))}
      </div>
    </div>
  );
}