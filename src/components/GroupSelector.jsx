import { useMemo, useState } from 'react';
import { useStore } from '../store/useStore.js';
import { getShortTitle } from '../lib/constants.js';

function safe(v) {
  if (v === null || undefined) return null;
  const s = String(v).trim();
  return s || null;
}

export default function GroupSelector() {
  const { headers, rows, config, groupingColIdx, setGroupingColIdx, updateGroups } = useStore();
  const [dragIdx, setDragIdx] = useState(null);

  const colOptions = useMemo(() => {
    if (!headers.length) return [];
    return headers.map((h, i) => ({ idx: i, name: getShortTitle(h) || `列${i + 1}` }))
      .filter(h => h.name);
  }, [headers]);

  const uniqueValues = useMemo(() => {
    if (groupingColIdx === null || !rows.length) return [];
    const values = new Set();
    for (const row of rows) {
      const v = safe(row[groupingColIdx]);
      if (v !== null) values.add(v);
    }
    return [...values].sort();
  }, [groupingColIdx, rows]);

  const selectedValues = useMemo(() => {
    if (!config.groups?.length) return [];
    const vals = [];
    for (const g of config.groups) {
      for (const [colIdx, val] of g.conditions) {
        if (colIdx === groupingColIdx && !vals.includes(val)) vals.push(val);
      }
    }
    return vals;
  }, [config.groups, groupingColIdx]);

  const handleColChange = (e) => {
    const colIdx = parseInt(e.target.value);
    if (isNaN(colIdx)) return;
    setGroupingColIdx(colIdx);
    updateGroups([]);
  };

  const handleValueToggle = (val) => {
    const currentGroups = config.groups ? [...config.groups] : [];
    if (selectedValues.includes(val)) {
      const newGroups = currentGroups
        .map(g => ({
          ...g,
          conditions: g.conditions.filter(([c]) => c !== groupingColIdx)
        }))
        .filter(g => g.conditions.length > 0);
      updateGroups(newGroups);
    } else {
      currentGroups.push({
        name: val,
        conditions: [[groupingColIdx, val]]
      });
      updateGroups(currentGroups);
    }
  };

  const handleDragStart = (e, idx) => {
    setDragIdx(idx);
    e.dataTransfer.effectAllowed = 'move';
  };

  const handleDragOver = (e, idx) => {
    e.preventDefault();
    e.dataTransfer.dropEffect = 'move';
    if (dragIdx === null || dragIdx === idx) return;

    const newGroups = [...(config.groups || [])];
    const [removed] = newGroups.splice(dragIdx, 1);
    newGroups.splice(idx, 0, removed);
    updateGroups(newGroups);
    setDragIdx(idx);
  };

  const handleDragEnd = () => {
    setDragIdx(null);
  };

  if (!headers.length) return null;

  return (
    <div className="bg-muji-warmWhite border border-muji-lightGray rounded-lg p-4">
      <h3 className="text-sm font-medium text-muji-charcoal mb-3">分组设置</h3>

      <div className="mb-4">
        <label className="text-xs text-muji-gray block mb-1">选择分组列</label>
        <select
          value={groupingColIdx ?? ''}
          onChange={handleColChange}
          className="w-full bg-white border border-muji-lightGray rounded px-3 py-2 text-sm text-muji-charcoal focus:outline-none focus:border-muji-beige"
        >
          <option value="">-- 选择列 --</option>
          {colOptions.map(opt => (
            <option key={opt.idx} value={opt.idx}>{opt.name}</option>
          ))}
        </select>
      </div>

      {groupingColIdx !== null && uniqueValues.length > 0 && (
        <div>
          <label className="text-xs text-muji-gray block mb-2">选择分组选项</label>
          <div className="flex flex-wrap gap-2">
            {uniqueValues.map(val => (
              <button
                key={val}
                onClick={() => handleValueToggle(val)}
                className={`px-3 py-1 text-sm rounded border transition-colors ${
                  selectedValues.includes(val)
                    ? 'bg-muji-beige border-muji-beige text-muji-charcoal'
                    : 'bg-white border-muji-lightGray text-muji-gray hover:border-muji-beige'
                }`}
              >
                {val}
              </button>
            ))}
          </div>
          <p className="text-xs text-muji-gray mt-2">
            已选择 {selectedValues.length} 个分组
          </p>
        </div>
      )}

      {config.groups?.length > 0 && (
        <div className="mt-3 pt-3 border-t border-muji-lightGray">
          <p className="text-xs text-muji-gray mb-1">当前分组（拖动调整顺序）：</p>
          <div className="flex flex-wrap gap-2">
            {config.groups.map((g, i) => (
              <div
                key={i}
                draggable
                onDragStart={(e) => handleDragStart(e, i)}
                onDragOver={(e) => handleDragOver(e, i)}
                onDragEnd={handleDragEnd}
                className={`flex items-center gap-1 text-xs bg-muji-beige px-2 py-1 rounded text-muji-charcoal cursor-move ${
                  dragIdx === i ? 'opacity-50' : ''
                }`}
              >
                <span>⋮⋮</span>
                <span>{g.name}</span>
              </div>
            ))}
          </div>
        </div>
      )}
    </div>
  );
}