// 文件名分类 + 大小格式化,规则照搬旧版 src/mockApi.jsx(UI 与后端共用同一组关键词)
import type { FileRole } from '../types/pipeline';

const FILENAME_PATTERNS: { role: FileRole; tests: RegExp[] }[] = [
  { role: '滞销表', tests: [/滞销/] },
  // 销售明细可以是「各商品客户拿货历史」(含「拿货历史」关键字)
  { role: '销售明细', tests: [/销售单明细/, /销售明细/, /销售流水/, /拿货历史/] },
];

export function classifyFile(name: string): FileRole | null {
  for (const p of FILENAME_PATTERNS) {
    if (p.tests.some((rx) => rx.test(name))) return p.role;
  }
  return null;
}

export function formatBytes(n: number): string {
  if (n < 1024) return `${n} B`;
  if (n < 1024 * 1024) return `${(n / 1024).toFixed(1)} KB`;
  return `${(n / (1024 * 1024)).toFixed(1)} MB`;
}
