// 文件名分类 + 大小格式化(UI 与 Worker 共用同一组关键词,改关键词只改这里)
import type { FileRole } from '../types/pipeline';

/** 字符串关键词 → 子串匹配的正则(关键词里没有正则元字符,直接用) */
const kw = (words: string[]): RegExp[] => words.map((w) => new RegExp(w));

/**
 * 「拿货历史」类文件名关键词(2026-11-16 用户指定,两种分析通用)。
 * 前三个都包含「拿货历史」,本来被最后一个覆盖;全部列出来是为了跟需求一一对应,
 * 以后有人删掉短词时长词也不会跟着失效。
 */
const SALES_HISTORY_KEYWORDS = kw(['各商品客户拿货历史', '商品客户拿货历史', '客户拿货历史', '拿货历史']);

/**
 * 「商品客户拿货·含散客」:中间的点在不同系统/输入法里可能是 · ・ • ‧ ∙ . 中的任一个,
 * 也可能没有点,这里统一放宽,避免因为点的写法不同识别失败。
 */
const HAN_SAN_KE = /商品客户拿货[·・•‧∙.]?含散客/;

const FILENAME_PATTERNS: { role: FileRole; tests: RegExp[] }[] = [
  // 先判滞销:文件名同时含「滞销」和「拿货历史」时算滞销商品
  { role: '滞销表', tests: [/滞销/] },
  {
    role: '销售明细',
    tests: [
      ...SALES_HISTORY_KEYWORDS,
      HAN_SAN_KE,
      // 旧关键词保留兼容
      /销售单明细/, /销售明细/, /销售流水/,
    ],
  },
];

export function classifyFile(name: string): FileRole | null {
  for (const p of FILENAME_PATTERNS) {
    if (p.tests.some((rx) => rx.test(name))) return p.role;
  }
  return null;
}

// 偏好分析只需 1 个拿货历史文件
const PREFERENCE_PATTERNS: RegExp[] = [
  ...SALES_HISTORY_KEYWORDS,
  HAN_SAN_KE,
  // 旧关键词保留兼容(「拿货」本身已覆盖上面所有拿货类写法)
  /客户偏好/, /拿货/,
];

export function classifyPreferenceFile(name: string): FileRole | null {
  return PREFERENCE_PATTERNS.some((rx) => rx.test(name)) ? '拿货历史' : null;
}

export function formatBytes(n: number): string {
  if (n < 1024) return `${n} B`;
  if (n < 1024 * 1024) return `${(n / 1024).toFixed(1)} KB`;
  return `${(n / (1024 * 1024)).toFixed(1)} MB`;
}
