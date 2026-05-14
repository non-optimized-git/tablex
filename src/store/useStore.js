import { create } from 'zustand';
import { loadConfig, saveConfig } from '../lib/configManager.js';

const initialConfig = {
  groups: [],
  ordered: {},
  mean_cols: {},
  multi_idx: []
};

export const useStore = create((set, get) => ({
  headers: [],
  rows: [],
  config: { ...initialConfig },
  selectedQuestions: [],
  groupingColIdx: null, // 单独跟踪分组列

  setHeaders: (headers) => set({ headers }),
  setRows: (rows) => set({ rows }),

  setConfig: (config) => {
    saveConfig(config);
    set({ config });
  },

  updateGroups: (groups) => {
    const config = { ...get().config, groups };
    saveConfig(config);
    set({ config });
  },

  setGroupingColIdx: (idx) => set({ groupingColIdx: idx }),

  setSelectedQuestions: (questions) => set({ selectedQuestions: questions }),

  loadSavedConfig: () => {
    const saved = loadConfig();
    if (saved) {
      set({ config: saved });
      // 从保存的配置中恢复 groupingColIdx
      if (saved.groups?.length > 0) {
        const colIdx = saved.groups[0].conditions?.[0]?.[0];
        if (colIdx !== undefined) set({ groupingColIdx: colIdx });
      }
      return saved;
    }
    return null;
  },

  resetConfig: () => {
    const config = { ...initialConfig };
    saveConfig(config);
    set({ config, groupingColIdx: null });
  },

  getOrderedForCol: (colIdx) => {
    const { config } = get();
    if (config.ordered?.[colIdx]) return config.ordered[colIdx];
    return null;
  },

  getMeanColsForCol: (colIdx) => {
    const { config } = get();
    if (config.mean_cols?.[colIdx]) return config.mean_cols[colIdx];
    return null;
  }
}));