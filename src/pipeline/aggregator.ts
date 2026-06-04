// aggregator.ts — 聚合阶段,算法逐行照搬 pipeline.py:174-282。
// 产出 base(滞销表全量货号 + 销售汇总 + 透传字段)和 货号×日期 透视(画图/Sheet3 用)。
//
// 字段透传是本工具最重要的可扩展性设计:滞销表里所有不在 KNOWN_ZHIXIAO_FIELDS
// 的字段自动进 extraFields,客户加新字段不需要改代码。跟拿货历史聚合字段
// (销售量/总销售金额)重名的自动加 "_滞销表" 后缀。
import type { Cell, SalesTable, ZhixiaoTable } from '../types/excel';

export interface BaseItem {
  货号: string;
  /** 契约 10 个元数据字段,滞销表缺哪个字段对应值为 null(输出留空) */
  品类: Cell;
  品牌: Cell;
  季节: Cell;       // 来自滞销表「年份」
  设计师: Cell;     // 来自滞销表「设计师品牌」
  上市天数: Cell;
  未成交天数: Cell; // 来自「未成交天数:」(带冒号)或「未成交天数」
  /** 0~1 浮点(原始 '12.5%' → 0.125);整列缺失为 null */
  销进率: number | null;
  库存价值: Cell;
  可售库存: Cell;
  /** 拿货历史聚合:净销售量合计(无记录 → 0) */
  销售量: number;
  /** 拿货历史聚合:净销售金额合计 round(2) */
  总销售金额: number;
  /** 滞销表盈利金额(coerce 失败 → 0,round(2));整列缺失为 null */
  盈利金额: number | null;
  /** 透传字段值,key 为输出列名(可能带 _滞销表 后缀) */
  extra: Record<string, Cell>;
}

export interface AggregatedBase {
  items: BaseItem[];          // 顺序 = 滞销表原始行序(去重后)
  extraFields: string[];      // 透传字段输出列名,按滞销表原始列序
  startDate: string;          // "YYYY-MM-DD"
  endDate: string;
  windowDays: number;
  dateRange: string[];        // 逐日,长度 = windowDays
  /** 货号 → 每日净销售量(与 dateRange 对齐);只含拿货历史里出现过的货号 */
  dailyByItem: Map<string, Float64Array>;
  itemsTotal: number;
  itemsWithSales: number;
}

const KNOWN_ZHIXIAO_FIELDS = new Set([
  '货号', '货号_k',
  '品类', '品牌', '设计师品牌', '年份',
  '上市天数', '未成交天数:', '未成交天数',
  '销进率', '库存价值', '可售库存', '盈利金额',
]);
const SALES_AGG_NAMES = new Set(['销售量', '总销售金额']);

/** 已知字段 src → 输出名 dst 的映射表(顺序同 pipeline.py:211-223) */
const KNOWN_FIELD_MAP: [src: string, dst: string][] = [
  ['品类', '品类'],
  ['品牌', '品牌'],
  ['设计师品牌', '设计师'],
  ['年份', '季节'],
  ['上市天数', '上市天数'],
  ['未成交天数:', '未成交天数'],
  ['未成交天数', '未成交天数'],
  ['销进率', '销进率'],
  ['库存价值', '库存价值'],
  ['可售库存', '可售库存'],
  ['盈利金额', '盈利金额'],
];

/** np.round 的 banker's rounding(round-half-even),保 2 位小数 */
export function npRound2(x: number): number {
  const y = x * 100;
  const f = Math.floor(y);
  const diff = y - f;
  let r: number;
  if (diff > 0.5) r = f + 1;
  else if (diff < 0.5) r = f;
  else r = f % 2 === 0 ? f : f + 1;
  return r / 100;
}

/** pd.to_numeric(errors='coerce') 等价(同 reader 内部版,这里独立一份) */
function toNumeric(v: Cell): number | null {
  if (v == null) return null;
  if (typeof v === 'number') return Number.isFinite(v) ? v : null;
  if (typeof v === 'boolean') return v ? 1 : 0;
  if (v instanceof Date) return null;
  const s = String(v).trim();
  if (!s) return null;
  const n = Number(s);
  return Number.isFinite(n) ? n : null;
}

/**
 * 销进率:'0%' / '12.5%' → 0.0 / 0.125。
 * 照搬 pandas:astype(str).rstrip('%').strip() → to_numeric(coerce) → fillna(0) → /100。
 * 单元格缺失(NaN→"nan")也走 coerce 失败 → 0。
 */
function parseXjl(v: Cell): number {
  const s = String(v ?? 'nan').replace(/%+$/, '').trim();
  const n = Number(s);
  return (Number.isFinite(n) ? n : 0) / 100;
}

function msToDateStr(ms: number): string {
  const d = new Date(ms);
  const mm = String(d.getUTCMonth() + 1).padStart(2, '0');
  const dd = String(d.getUTCDate()).padStart(2, '0');
  return `${d.getUTCFullYear()}-${mm}-${dd}`;
}

export function aggregate(zhixiao: ZhixiaoTable, sales: SalesTable): AggregatedBase {
  // ── 时间窗:拿货历史最早 ~ 最晚下单时间(不做智能裁断,用户拒绝过)──
  let minDate = sales.records[0].date;
  let maxDate = sales.records[0].date;
  for (const r of sales.records) {
    if (r.date < minDate) minDate = r.date;
    if (r.date > maxDate) maxDate = r.date;
  }
  const startMs = Date.parse(minDate);
  const windowDays = Math.round((Date.parse(maxDate) - startMs) / 86400000) + 1;
  const dateRange: string[] = [];
  for (let i = 0; i < windowDays; i++) dateRange.push(msToDateStr(startMs + i * 86400000));

  // ── 日维度透视(货号 × 日 × 净销售量)+ 销售汇总 ──────────────────────
  const dailyByItem = new Map<string, Float64Array>();
  const salesQty = new Map<string, number>();
  const salesAmt = new Map<string, number>();
  for (const r of sales.records) {
    let row = dailyByItem.get(r.key);
    if (!row) {
      row = new Float64Array(windowDays);
      dailyByItem.set(r.key, row);
    }
    const dayIdx = Math.round((Date.parse(r.date) - startMs) / 86400000);
    row[dayIdx] += r.qty;
    salesQty.set(r.key, (salesQty.get(r.key) ?? 0) + r.qty);
    salesAmt.set(r.key, (salesAmt.get(r.key) ?? 0) + r.amt);
  }

  // ── 滞销表字段:已知字段映射 + 其余全部透传 ───────────────────────────
  const colSet = new Set(zhixiao.columns);
  const ziKeep = new Set<string>(); // 已认领的源列
  /** dst → src(已知字段;同 dst 多 src 时取先到的,对齐 pipeline.py 循环顺序) */
  const knownSrc = new Map<string, string>();
  for (const [src, dst] of KNOWN_FIELD_MAP) {
    if (colSet.has(src) && !ziKeep.has(src)) {
      ziKeep.add(src);
      if (!knownSrc.has(dst)) knownSrc.set(dst, src);
    }
  }

  const extraFields: string[] = [];
  const extraFieldSources = new Map<string, string>(); // 输出列名 → 滞销表原列名
  for (const c of zhixiao.columns) {
    if (KNOWN_ZHIXIAO_FIELDS.has(c)) continue;
    if (ziKeep.has(c)) continue;
    const outName = SALES_AGG_NAMES.has(c) ? `${c}_滞销表` : c;
    extraFields.push(outName);
    extraFieldSources.set(outName, c);
    ziKeep.add(c);
  }

  const hasXjl = colSet.has('销进率');
  const hasProfit = colSet.has('盈利金额');

  // ── 拼装 base:滞销表货号为基础,左连接销售汇总 ────────────────────────
  const items: BaseItem[] = [];
  let itemsWithSales = 0;
  for (const row of zhixiao.rows) {
    const cells = row.cells;
    const get = (dst: string): Cell => {
      const src = knownSrc.get(dst);
      return src != null ? cells[src] ?? null : null;
    };
    const qty = Math.trunc(salesQty.get(row.key) ?? 0); // fillna(0).astype(int)
    if (qty > 0) itemsWithSales++;
    const extra: Record<string, Cell> = {};
    for (const f of extraFields) extra[f] = cells[extraFieldSources.get(f)!] ?? null;
    items.push({
      货号: row.key,
      品类: get('品类'),
      品牌: get('品牌'),
      季节: get('季节'),
      设计师: get('设计师'),
      上市天数: get('上市天数'),
      未成交天数: get('未成交天数'),
      销进率: hasXjl ? parseXjl(get('销进率')) : null,
      库存价值: get('库存价值'),
      可售库存: get('可售库存'),
      销售量: qty,
      总销售金额: npRound2(salesAmt.get(row.key) ?? 0),
      盈利金额: hasProfit ? npRound2(toNumeric(get('盈利金额')) ?? 0) : null,
      extra,
    });
  }

  return {
    items,
    extraFields,
    startDate: minDate,
    endDate: maxDate,
    windowDays,
    dateRange,
    dailyByItem,
    itemsTotal: items.length,
    itemsWithSales,
  };
}
