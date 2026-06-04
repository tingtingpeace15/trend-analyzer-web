// reader.ts — SheetJS 读 Excel + 字段嗅探,算法逐行照搬 pipeline.py 步骤 2 的读取部分
// (pipeline.py:116-172)。pandas 的隐式行为(列名规则、header 消耗、NaN 处理)也一并复刻,
// 否则行数/列名跟 Python 版对不上,M7 的 diff 过不了。
import * as XLSX from 'xlsx';
import { PipelineError } from './errors';
import type {
  Cell,
  SalesRecord,
  SalesSheetMeta,
  SalesTable,
  ZhixiaoRow,
  ZhixiaoTable,
} from '../types/excel';

// ── 基础工具 ─────────────────────────────────────────────────────────────

function readWorkbook(data: ArrayBuffer | Uint8Array): XLSX.WorkBook {
  const u8 = data instanceof Uint8Array ? data : new Uint8Array(data);
  // raw 数值模式(日期保持 Excel serial,自己转,避免时区坑);关掉不需要的解析省内存
  return XLSX.read(u8, {
    type: 'array',
    cellDates: false,
    cellFormula: false,
    cellHTML: false,
    cellText: false,
    dense: true,
  });
}

/**
 * pandas read_excel 默认 na_values:这些字符串(以及空串)一律读成 NaN。
 * SheetJS 原样返回字符串,这里归一化成 null,否则空字符串单元格会跟 pandas 产生分歧。
 */
const PANDAS_NA_VALUES = new Set([
  '', '#N/A', '#N/A N/A', '#NA', '-1.#IND', '-1.#QNAN', '-NaN', '-nan',
  '1.#IND', '1.#QNAN', '<NA>', 'N/A', 'NA', 'NULL', 'NaN', 'None', 'n/a', 'nan', 'null',
]);

/** sheet → 行数组(含空行,trim 掉尾部全空行)。等价 pandas 不跳过中间空行的行为 */
function sheetToRows(ws: XLSX.WorkSheet): Cell[][] {
  const rows: Cell[][] = XLSX.utils.sheet_to_json(ws, {
    header: 1,
    defval: null,
    blankrows: true,
    raw: true,
  });
  for (const row of rows) {
    for (let c = 0; c < row.length; c++) {
      if (typeof row[c] === 'string' && PANDAS_NA_VALUES.has(row[c] as string)) row[c] = null;
    }
  }
  let end = rows.length;
  while (end > 0 && rows[end - 1].every((v) => v == null)) end--;
  return rows.slice(0, end);
}

function sheetWidth(ws: XLSX.WorkSheet): number {
  const ref = ws['!ref'];
  if (!ref) return 0;
  return XLSX.utils.decode_range(ref).e.c + 1;
}

/**
 * pandas 风格列名:空表头 → "Unnamed: {i}",重名 → "name.1" / "name.2"。
 * 宽度按整张 sheet 算(数据行比表头宽时 pandas 也会补 Unnamed 列)。
 */
function pandasColumns(headerCells: Cell[], width: number): string[] {
  const names: string[] = [];
  const counts = new Map<string, number>();
  for (let i = 0; i < width; i++) {
    const v = headerCells[i];
    let name = v == null || v === '' ? `Unnamed: ${i}` : String(v);
    const seen = counts.get(name) ?? 0;
    counts.set(name, seen + 1);
    if (seen > 0) name = `${name}.${seen}`;
    names.push(name);
  }
  return names;
}

/**
 * 货号 → key:对齐 pandas `astype(str).str.strip()`。
 * 注意 pandas 对 NaN 的 astype(str) 得到字符串 "nan"(空行不会被滞销表 dropna,
 * 而是变成 key="nan" 参与去重)——这里照搬,保持行数一致。
 */
export function keyOf(v: Cell): string {
  if (v == null) return 'nan';
  return String(v).trim();
}

/** pd.to_numeric(errors='coerce'):失败返回 null */
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
 * pd.to_datetime(errors='coerce') 的等价:支持 Excel serial(数字)、Date、
 * "2026-01-01 12:30:45" / "2026/1/5" 等字符串。返回 UTC ms,失败 null。
 */
function parseDateValue(v: Cell, date1904: boolean): number | null {
  if (v == null) return null;
  if (v instanceof Date) return v.getTime();
  if (typeof v === 'number') {
    if (!Number.isFinite(v)) return null;
    const epoch = date1904 ? Date.UTC(1904, 0, 1) : Date.UTC(1899, 11, 30);
    return epoch + Math.round(v * 86400000);
  }
  if (typeof v === 'string') {
    const s = v.trim();
    const m = s.match(
      /^(\d{4})[-/](\d{1,2})[-/](\d{1,2})(?:[ T](\d{1,2}):(\d{2})(?::(\d{2}))?)?$/,
    );
    if (!m) return null;
    return Date.UTC(+m[1], +m[2] - 1, +m[3], +(m[4] ?? 0), +(m[5] ?? 0), +(m[6] ?? 0));
  }
  return null;
}

function msToDateStr(ms: number): string {
  const d = new Date(ms);
  const mm = String(d.getUTCMonth() + 1).padStart(2, '0');
  const dd = String(d.getUTCDate()).padStart(2, '0');
  return `${d.getUTCFullYear()}-${mm}-${dd}`;
}

// ── 滞销商品表(pipeline.py:116-126)────────────────────────────────────

/**
 * 读滞销商品:表头在第 2 行(header=1,第 1 行是合并大标题),
 * 必需字段只有「货号」;按 货号_k 去重保留首行。所有列原样保留(透传用)。
 */
export function readZhixiao(data: ArrayBuffer | Uint8Array): ZhixiaoTable {
  let wb: XLSX.WorkBook;
  try {
    wb = readWorkbook(data);
  } catch (e) {
    throw new PipelineError(
      `无法读取滞销商品:${e instanceof Error ? e.message : e}`,
      '请确认文件未损坏,且第一行是合并标题、第二行是真表头。',
    );
  }
  const ws = wb.Sheets[wb.SheetNames[0]];
  const rows = sheetToRows(ws);
  const width = sheetWidth(ws);
  const HEADER_ROW = 1; // pandas header=1
  if (rows.length <= HEADER_ROW) {
    throw new PipelineError(
      '无法读取滞销商品:文件行数不足',
      '请确认文件未损坏,且第一行是合并标题、第二行是真表头。',
    );
  }
  const columns = pandasColumns(rows[HEADER_ROW], width);
  if (!columns.includes('货号')) {
    throw new PipelineError(
      '滞销商品缺少必需字段「货号」',
      '请检查文件第二行是否包含 货号 列。',
    );
  }
  const keyIdx = columns.indexOf('货号');

  const seen = new Set<string>();
  const out: ZhixiaoRow[] = [];
  for (let r = HEADER_ROW + 1; r < rows.length; r++) {
    const row = rows[r];
    const key = keyOf(row[keyIdx] ?? null);
    if (seen.has(key)) continue; // drop_duplicates(keep='first')
    seen.add(key);
    const cells: Record<string, Cell> = {};
    for (let c = 0; c < columns.length; c++) cells[columns[c]] = row[c] ?? null;
    out.push({ key, cells });
  }
  return { columns, rows: out };
}

// ── 拿货历史(pipeline.py:128-172)───────────────────────────────────────

const SNIFF_KEYWORDS = ['货号', '净销售', '下单时间', '客户名称'];
const REQUIRED_SALES_FIELDS = ['货号', '下单时间', '净销售量', '净销售金额'] as const;

/**
 * 读拿货历史:多 sheet。每个 sheet 在前 3 行嗅探表头(含任一关键词的行);
 * 嗅探失败按 header=0(首行当表头,即使是数据——Python 版就这样,照搬)。
 * 列数 <10 的 sheet 丢弃;多 sheet 按列名对齐合并;清洗后只保留聚合需要的 4 个字段。
 */
export function readSales(data: ArrayBuffer | Uint8Array, fileName = ''): SalesTable {
  let wb: XLSX.WorkBook;
  try {
    wb = readWorkbook(data);
  } catch (e) {
    throw new PipelineError(`无法读取拿货历史:${e instanceof Error ? e.message : e}`, '');
  }

  const date1904 = !!wb.Workbook?.WBProps?.date1904;
  const sheetMeta: SalesSheetMeta[] = [];
  interface Frame { columns: string[]; rows: Cell[][]; start: number }
  const frames: Frame[] = [];

  for (const name of wb.SheetNames) {
    const ws = wb.Sheets[name];
    const rows = sheetToRows(ws);
    const width = sheetWidth(ws);
    // 嗅探:前 3 行里找含关键词的行(pandas 把整行 str() 后 join 查子串)
    let hrow: number | null = null;
    for (let i = 0; i < Math.min(3, rows.length); i++) {
      const joined = rows[i].map((v) => (v == null ? 'nan' : String(v))).join(' ');
      if (SNIFF_KEYWORDS.some((k) => joined.includes(k))) {
        hrow = i;
        break;
      }
    }
    const headerIdx = hrow ?? 0;
    const columns = pandasColumns(rows[headerIdx] ?? [], width);
    const dataRows = rows.length - headerIdx - 1;
    const kept = columns.length >= 10;
    sheetMeta.push({ name, headerRow: hrow, cols: columns.length, rows: Math.max(dataRows, 0), kept });
    if (kept) frames.push({ columns, rows, start: headerIdx + 1 });
  }

  if (frames.length === 0) {
    throw new PipelineError(
      `拿货历史 ${fileName} 里所有 sheet 列数都太少(<10 列),无法识别为有效销售明细。`,
      '请检查上传的是不是「各商品客户拿货历史」之类的销售明细文件。',
    );
  }

  // 列对齐(pipeline.py:150-155):后续 sheet 列数 ≥ 首 sheet 时截断并改名为首 sheet 列;
  // 列数更少的保持原名,由 concat 按列名对齐(并集)
  const baseCols = frames[0].columns;
  for (let i = 1; i < frames.length; i++) {
    if (frames[i].columns.length >= baseCols.length) {
      frames[i] = { ...frames[i], columns: baseCols };
    }
  }
  const columns: string[] = [...baseCols];
  for (let i = 1; i < frames.length; i++) {
    for (const c of frames[i].columns) if (!columns.includes(c)) columns.push(c);
  }

  for (const need of REQUIRED_SALES_FIELDS) {
    if (!columns.includes(need)) {
      throw new PipelineError(
        `拿货历史缺少必需字段「${need}」。`,
        '请确认表头含 货号 / 下单时间 / 净销售量 / 净销售金额。',
      );
    }
  }

  // 合并 + 清洗(pipeline.py:167-172):
  // dropna(下单时间, 货号) → 货号_k trim → 净销售量/金额 coerce fillna(0) → 日期
  let rawRowCount = 0;
  const records: SalesRecord[] = [];
  for (const f of frames) {
    const idx = {
      key: f.columns.indexOf('货号'),
      time: f.columns.indexOf('下单时间'),
      qty: f.columns.indexOf('净销售量'),
      amt: f.columns.indexOf('净销售金额'),
    };
    for (let r = f.start; r < f.rows.length; r++) {
      rawRowCount++;
      const row = f.rows[r];
      const rawKey = idx.key >= 0 ? row[idx.key] ?? null : null;
      if (rawKey == null) continue; // dropna(货号)
      const ms = parseDateValue(idx.time >= 0 ? row[idx.time] ?? null : null, date1904);
      if (ms == null) continue; // to_datetime coerce 失败 → dropna(下单时间)
      records.push({
        key: keyOf(rawKey),
        date: msToDateStr(ms),
        qty: toNumeric(idx.qty >= 0 ? row[idx.qty] ?? null : null) ?? 0,
        amt: toNumeric(idx.amt >= 0 ? row[idx.amt] ?? null : null) ?? 0,
      });
    }
  }

  return { sheetNames: wb.SheetNames, sheetMeta, columns, rawRowCount, records };
}
