// excel.ts — 读取阶段的数据结构(对应 pipeline.py 步骤 2 的 df_zi / df_h)

/** 单元格原始值(SheetJS raw 模式) */
export type Cell = string | number | boolean | Date | null;

/** 滞销表一行:key = 货号_k(trim 后),cells 按原始列名存全部字段(透传用) */
export interface ZhixiaoRow {
  key: string;
  cells: Record<string, Cell>;
}

/** 滞销表(已按 货号_k 去重,保留首行) */
export interface ZhixiaoTable {
  /** 原始列名,pandas 风格(空表头 → "Unnamed: N",重名 → ".1" 后缀),顺序保留 */
  columns: string[];
  rows: ZhixiaoRow[];
}

/** 拿货历史清洗后的一条销售记录(下游聚合只需要这 4 个字段,其余列丢弃) */
export interface SalesRecord {
  key: string;   // 货号_k
  date: string;  // 日期 "YYYY-MM-DD"(下单时间的日期部分)
  qty: number;   // 净销售量(coerce 失败 → 0)
  amt: number;   // 净销售金额(coerce 失败 → 0)
}

/** 拿货历史单个 sheet 的读取元数据(测试对照 pandas 用) */
export interface SalesSheetMeta {
  name: string;
  /** 嗅探到的表头行号(0-based);嗅探失败为 null(此时按 header=0 读,跟 Python 版一致) */
  headerRow: number | null;
  cols: number;
  /** 表头之后的数据行数 */
  rows: number;
  /** 列数 ≥10 才保留 */
  kept: boolean;
}

/** 拿货历史(多 sheet 合并 + 清洗后) */
export interface SalesTable {
  sheetNames: string[];
  sheetMeta: SalesSheetMeta[];
  /** 合并后的列名(pandas concat 语义:按列名对齐,取并集) */
  columns: string[];
  /** concat 后、清洗(dropna)前的行数 */
  rawRowCount: number;
  records: SalesRecord[];
}
