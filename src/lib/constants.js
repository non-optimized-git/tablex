// 模块检测关键字（对应 Python 的 detect_module）
export const MODULES = [
  { name: '用户画像', keywords: ['年龄', '性别', '职业', '学历', '收入'] },
  { name: '购车意向', keywords: ['车', '换车', '购车', '预算'] },
  { name: '智能驾驶', keywords: ['智能', '驾驶', '智驾', '安全'] },
  { name: '家庭用车', keywords: ['家庭', '孩子', '人口', '需求'] },
  { name: '产品体验', keywords: ['空间', '座椅', '内饰', '外观', '续航'] },
];

export function detectModule(header) {
  if (!header) return '其他';
  const s = String(header);
  for (const mod of MODULES) {
    if (mod.keywords.some(kw => s.includes(kw))) {
      return mod.name;
    }
  }
  return '其他';
}

export function getShortTitle(header) {
  if (!header) return '';
  let s = String(header);
  s = s.replace(/<[^>]+>/g, '');
  s = s.replace(/（[^）]+）/g, '');
  s = s.replace(/\([^)]+\)/g, '');
  s = s.replace(/[-－—_/\\|]/g, ' ');
  s = s.replace(/\s+/g, ' ').trim();
  return s;
}

export const STORAGE_KEY = 'survey-analysis-config';