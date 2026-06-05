// preferenceReader.ts — 偏好分析的读取 + 列规范化,逐行照搬
// preference_pipeline.py 的 _load(:69-112) 与 _normalize_columns(:115-168)。
//
// 与趋势分析 reader 的三处不同(都是 Python 版自己的差异,如实复刻):
//   1. 嗅探关键词是 [客户名称, 销售, 下单时间, 货号](trend 是 货号/净销售/...)
//   2. 列对齐:== 改名、> 截断;< 保留原名由 concat 按列名取并集(trend 只处理 >=)
//   3. 读完做模糊列名规范化(「颜色:」→「颜色」等),毛值列(销售量/销售金额)被净值顶掉后丢弃
import { PipelineError } from './errors';
import {
  pandasColumns,
  parseDateValue,
  readWorkbook,
  sheetToRows,
  sheetWidth,
  toNumeric,
} from './reader';
import type { Cell, SalesSheetMeta } from '../types/excel';

/** _analyze/_build_excel 会用到的规范化列(销退单ID 之类的无关列只留名字不留数据) */
const USED_COLUMNS = new Set([
  '店铺', '货号', '分类', '年份', '设计师品牌', '品牌', '颜色', '尺码',
  '下单时间', '客户名称', '销售', '销售量', '销售金额',
]);

export interface PreferenceData {
  sheetNames: string[];
  sheetMeta: SalesSheetMeta[];
  rawRowCount: number;
  /** 规范化后的列名(顺序保留,含未映射的透传名) */
  columns: string[];
  /** 被丢弃的毛值冲突列(原始列名) */
  droppedConflicts: string[];
  /** 规范化列名 → 原始值数组(长度 rawRowCount;只存 USED_COLUMNS) */
  cols: Map<string, (Cell | null)[]>;
  /** 销售量 / 销售金额:to_numeric(coerce) 后(NaN 表缺失) */
  qty: Float64Array;
  amt: Float64Array;
  /** 下单时间:to_datetime(coerce) 后的 UTC ms(NaN 表 NaT);整列缺失为 null */
  orderMs: Float64Array | null;
}

/** _normalize_columns 的列名映射(if/elif 顺序敏感,照搬) */
function remapName(c: string): string | null {
  const s = c.trim().replace(/[:：]+$/, '');
  if (s.includes('店铺')) return '店铺';
  if (s.includes('货号')) return '货号';
  if (s === '分类' || s === '品类' || s === '类别') return '分类';
  if (s.includes('年份')) return '年份';
  if (s.includes('设计师')) return '设计师品牌';
  if (s === '品牌') return '品牌';
  if (s.includes('颜色')) return '颜色';
  if (s.includes('尺码') || s.includes('码数')) return '尺码';
  if (s.includes('下单时间') || s.includes('日期')) return '下单时间';
  if (s === '客户名称' || s === '客户') return '客户名称';
  if (s === '销售' || s === '销售员' || s === '业务员') return '销售';
  if (s.includes('净销售金额')) return '销售金额';
  if (s.includes('净销售量')) return '销售量';
  return null;
}

export function loadPreference(data: ArrayBuffer | Uint8Array): PreferenceData {
  let wb;
  try {
    wb = readWorkbook(data);
  } catch (e) {
    throw new PipelineError(
      `读不开这个 Excel：${e instanceof Error ? e.message : e}`,
      '确认文件能在 Excel/Numbers 里正常打开,没有损坏。',
    );
  }
  const date1904 = !!wb.Workbook?.WBProps?.date1904;

  // ── _load:逐 sheet 嗅探表头,<10 列丢弃 ────────────────────────────────
  const SNIFF = ['客户名称', '销售', '下单时间', '货号'];
  interface Frame { columns: string[]; rows: Cell[][]; start: number }
  const frames: Frame[] = [];
  const sheetMeta: SalesSheetMeta[] = [];
  for (const name of wb.SheetNames) {
    const ws = wb.Sheets[name];
    const rows = sheetToRows(ws);
    const width = sheetWidth(ws);
    let hrow: number | null = null;
    for (let i = 0; i < Math.min(3, rows.length); i++) {
      const joined = rows[i].map((v) => (v == null ? 'nan' : String(v))).join(' ');
      if (SNIFF.some((k) => joined.includes(k))) {
        hrow = i;
        break;
      }
    }
    const headerIdx = hrow ?? 0;
    const columns = pandasColumns(rows[headerIdx] ?? [], width);
    const kept = columns.length >= 10;
    sheetMeta.push({
      name, headerRow: hrow, cols: columns.length,
      rows: Math.max(rows.length - headerIdx - 1, 0), kept,
    });
    if (kept) frames.push({ columns, rows, start: headerIdx + 1 });
  }
  if (frames.length === 0) {
    throw new PipelineError(
      'Excel 里所有 sheet 列数都太少（<10 列），无法识别为有效销售明细',
      '请检查上传的是不是「各商品客户拿货历史」之类的销售明细文件。',
    );
  }

  // 列对齐(preference_pipeline.py:103-109):== 改名、> 截断改名、< 保留原名
  const baseCols = frames[0].columns;
  for (let i = 1; i < frames.length; i++) {
    if (frames[i].columns.length >= baseCols.length) {
      frames[i] = { ...frames[i], columns: baseCols };
    }
  }
  // concat 取列名并集(出现顺序)
  const rawColumns: string[] = [...baseCols];
  for (let i = 1; i < frames.length; i++) {
    for (const c of frames[i].columns) if (!rawColumns.includes(c)) rawColumns.push(c);
  }
  let rawRowCount = 0;
  for (const f of frames) rawRowCount += f.rows.length - f.start;

  // ── _normalize_columns:remap → 冲突丢列 → rename ───────────────────────
  const remap = new Map<string, string>();
  for (const c of rawColumns) {
    const t = remapName(c);
    if (t != null) remap.set(c, t);
  }
  const targetNames = new Set(remap.values());
  const droppedConflicts = rawColumns.filter(
    (c) => !remap.has(c) && targetNames.has(c.trim().replace(/[:：]+$/, '')),
  );
  const droppedSet = new Set(droppedConflicts);
  const columns = rawColumns
    .filter((c) => !droppedSet.has(c))
    .map((c) => remap.get(c) ?? c);

  for (const need of ['客户名称', '销售量', '销售金额']) {
    if (!columns.includes(need)) {
      throw new PipelineError(
        `上传的文件缺少必需字段「${need}」`,
        '销售明细必须包含 客户名称 / 净销售量 / 净销售金额 三列（取净值，已扣退货），请检查文件表头是不是含「净」前缀。',
      );
    }
  }

  // ── 物化列式数据(只存用得上的列)────────────────────────────────────────
  const keptRaw = rawColumns.filter((c) => !droppedSet.has(c));
  const cols = new Map<string, (Cell | null)[]>();
  for (let i = 0; i < keptRaw.length; i++) {
    const norm = columns[i];
    if (USED_COLUMNS.has(norm)) cols.set(norm, new Array(rawRowCount).fill(null));
  }
  let at = 0;
  for (const f of frames) {
    // 该 frame 的列名 → 在 keptRaw/columns 里的位置(concat 按列名对齐)
    const idxOf = new Map<string, number>();
    f.columns.forEach((c, i) => {
      if (!droppedSet.has(c) && !idxOf.has(c)) idxOf.set(c, i);
    });
    for (let r = f.start; r < f.rows.length; r++, at++) {
      const row = f.rows[r];
      for (let k = 0; k < keptRaw.length; k++) {
        const arr = cols.get(columns[k]);
        if (!arr) continue;
        const srcIdx = idxOf.get(keptRaw[k]);
        if (srcIdx !== undefined) arr[at] = row[srcIdx] ?? null;
      }
    }
  }

  // ── 数值/日期 coerce(同 pandas,保留 NaN 不 fillna)─────────────────────
  const qty = new Float64Array(rawRowCount).fill(NaN);
  const amt = new Float64Array(rawRowCount).fill(NaN);
  const qtyRaw = cols.get('销售量')!;
  const amtRaw = cols.get('销售金额')!;
  for (let i = 0; i < rawRowCount; i++) {
    const q = toNumeric(qtyRaw[i]);
    if (q != null) qty[i] = q;
    const a = toNumeric(amtRaw[i]);
    if (a != null) amt[i] = a;
  }
  let orderMs: Float64Array | null = null;
  const timeRaw = cols.get('下单时间');
  if (timeRaw) {
    orderMs = new Float64Array(rawRowCount).fill(NaN);
    for (let i = 0; i < rawRowCount; i++) {
      const ms = parseDateValue(timeRaw[i], date1904);
      if (ms != null) orderMs[i] = ms;
    }
  }

  return {
    sheetNames: wb.SheetNames,
    sheetMeta,
    rawRowCount,
    columns,
    droppedConflicts,
    cols,
    qty,
    amt,
    orderMs,
  };
}
