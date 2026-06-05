// writer.ts — ExcelJS 组装输出 Excel,逐行照搬 pipeline.py:440-755(步骤 4)。
// 3 个 sheet:商品销售趋势(小图)/ 款趋势明细图(大图)/ 款日销量明细(日期矩阵)。
// 列宽/行高/颜色/格式/冻结/筛选/锚点 全部沿用 Python 版数字。
import ExcelJS from 'exceljs';
import type { AggregatedBase, BaseItem } from './aggregator';
import type { Cell } from '../types/excel';
import type { LogKind } from '../types/pipeline';

export interface ItemImages { sm: Uint8Array; dt: Uint8Array }
export type LogFn = (text: string, kind: LogKind, step: number) => void;

// ── 样式常量(pipeline.py:489-495)───────────────────────────────────────
const HEADER_FILL: ExcelJS.Fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: '002F5597' } };
const ALT_FILL: ExcelJS.Fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: '00F2F6FB' } };
const HEADER_FONT: Partial<ExcelJS.Font> = { color: { argb: '00FFFFFF' }, bold: true, size: 11 };
const CENTER: Partial<ExcelJS.Alignment> = { horizontal: 'center', vertical: 'middle' };
const LEFT: Partial<ExcelJS.Alignment> = { horizontal: 'left', vertical: 'middle', indent: 1 };
const THIN: Partial<ExcelJS.Border> = { style: 'thin', color: { argb: '00DDDDDD' } };
const BORDER: Partial<ExcelJS.Borders> = { left: THIN, right: THIN, top: THIN, bottom: THIN };

const RED = '00E74C3C';
const GREEN = '0027AE60';
const GRAY = '00999999';

const HEAD_KNOWN = ['货号', '品类', '品牌', '季节', '设计师',
  '上市天数', '未成交天数', '销进率', '库存价值', '可售库存'];
const HEAD_TAIL = ['销售量', '总销售金额', '盈利金额'];

/**
 * 固定排位的透传字段:滞销表里出现时不进默认透传区(可售库存之后),
 * 而是插在「销售量」之后。2026-06-05 用户要求(A价)。
 * 注意:这是对 Python 版列序的有意偏离;若 Python 版同步改,diff 才会一致。
 */
const PINNED_AFTER_QTY = ['A价'];
const HEAD_KNOWN_WIDTHS = [18, 14, 12, 10, 10, 10, 12, 10, 12, 10];
const DETAIL_KNOWN_WIDTHS = [16, 14, 12, 10, 10, 10, 12, 10, 12, 10];
const META3_KNOWN_WIDTHS = [14, 14, 12, 10, 10, 10, 12, 10, 12, 10];
const HEAD_TAIL_WIDTHS = [10, 14, 14];
const EXTRA_WIDTH = 12;

function colLetter(n: number): string {
  let s = '';
  while (n > 0) {
    const m = (n - 1) % 26;
    s = String.fromCharCode(65 + m) + s;
    n = Math.floor((n - 1) / 26);
  }
  return s;
}

// ── pandas str()/数值 语义 ────────────────────────────────────────────────

/**
 * 列是否为 pandas float64(全数值 + 存在空值或小数)。
 * float64 列里整数 str() 出来带 ".0"(如 "387.0"),要原样复刻。
 */
function columnIsFloat64(values: Cell[]): boolean {
  let hasNullOrFloat = false;
  for (const v of values) {
    if (v == null) {
      hasNullOrFloat = true;
      continue;
    }
    if (typeof v !== 'number') return false; // object dtype
    if (!Number.isInteger(v)) hasNullOrFloat = true;
  }
  return hasNullOrFloat;
}

/** as_str(pipeline.py:473-481)+ pandas dtype 的字符串化语义 */
function asStr(v: Cell, isFloat64: boolean): string {
  if (v == null) return '';
  if (typeof v === 'number') {
    if (isFloat64 && Number.isInteger(v)) return `${v}.0`;
    return String(v);
  }
  if (typeof v === 'boolean') return v ? 'True' : 'False';
  return String(v);
}

/** as_num(pipeline.py:465-471):NaN → null(留空),其余原样 */
function asNum(v: Cell): Cell {
  return v ?? null;
}

/** build_comment 里的 fmt(pipeline.py:440-442) */
function fmtC(v: number): string {
  return v === Math.trunc(v) ? String(Math.trunc(v)) : String(v);
}

/** 批注文本(pipeline.py:444-463) */
function buildComment(values: ArrayLike<number>, dateRange: string[]): string {
  const n = dateRange.length;
  let total = 0, active = 0, peakVal = -Infinity, peakI = 0;
  for (let i = 0; i < n; i++) {
    const v = values[i] ?? 0;
    total += v;
    if (v !== 0) active++;
    if (v > peakVal) {
      peakVal = v;
      peakI = i;
    }
  }
  if (active === 0) return `${n}日内无销售记录`;
  const lines = [`${n}日销售量: ${fmtC(total)}`];
  if (peakVal > 0) lines.push(`峰值: ${fmtC(peakVal)} (${dateRange[peakI].slice(5)})`);
  lines.push(`有销量天数: ${active}/${n}`);
  lines.push('-'.repeat(24));
  lines.push('日期      销售量');
  for (let i = 0; i < n; i++) {
    const v = values[i] ?? 0;
    if (v !== 0) lines.push(`${dateRange[i].slice(5)}     ${fmtC(v).padStart(6)}`);
  }
  return lines.join('\n');
}

// ── 排序(pipeline.py:483-486:多列降序,NaN 在后,稳定)──────────────────

interface SortedRow { item: BaseItem; origIdx: number }

function sortBase(agg: AggregatedBase): SortedRow[] {
  const hasKucun = agg.items.some((it) => it.可售库存 != null);
  const keyNum = (v: Cell): number | null =>
    typeof v === 'number' && Number.isFinite(v) ? v : null;
  const rows: SortedRow[] = agg.items.map((item, origIdx) => ({ item, origIdx }));
  // JS Array.sort 是稳定排序(同 np.lexsort)
  rows.sort((a, b) => {
    if (a.item.销售量 !== b.item.销售量) return b.item.销售量 - a.item.销售量;
    if (hasKucun) {
      const ka = keyNum(a.item.可售库存);
      const kb = keyNum(b.item.可售库存);
      if (ka == null && kb == null) return 0;
      if (ka == null) return 1;  // na_position='last'
      if (kb == null) return -1;
      if (ka !== kb) return kb - ka;
    }
    return 0;
  });
  return rows;
}

// ── 嵌图锚点(lib/anchor_util.py:TwoCellAnchor,2px 内边距)──────────────
//
// 注意:不能用 ExcelJS 的小数坐标(tl: {col: 13.03}),它内部按「列宽×10000 EMU」
// 换算偏移,而真实像素是 9525 EMU/px——差 ~3.5 倍,图会被压扁(实测翻车过)。
// 改为直接传原生 EMU 偏移,跟 openpyxl 的 AnchorMarker(col, colOff) 一一对应:
//   from = (col-1, +2px), to = (col, -2px),editAs='twoCell' 图随排序/筛选伸缩。

const EMU_PER_PX = 9525;

function addAnchoredImage(ws: ExcelJS.Worksheet, imageId: number, col1: number, row1: number) {
  const pad = 2 * EMU_PER_PX;
  ws.addImage(imageId, {
    tl: { nativeCol: col1 - 1, nativeColOff: pad, nativeRow: row1 - 1, nativeRowOff: pad },
    br: { nativeCol: col1, nativeColOff: -pad, nativeRow: row1, nativeRowOff: -pad },
    editAs: 'twoCell',
  } as unknown as Parameters<ExcelJS.Worksheet['addImage']>[1]);
}

// ── 主入口 ────────────────────────────────────────────────────────────────

export async function writeExcel(
  agg: AggregatedBase,
  images: Map<number, ItemImages>,
  log: LogFn,
): Promise<Uint8Array> {
  const { extraFields, dateRange } = agg;
  const sorted = sortBase(agg);
  const zeroValues = new Float64Array(dateRange.length);
  const valuesOf = (item: BaseItem): ArrayLike<number> =>
    agg.dailyByItem.get(item.货号) ?? zeroValues;

  // pandas 列 dtype(as_str 字段)
  const f64 = {
    品类: columnIsFloat64(agg.items.map((i) => i.品类)),
    品牌: columnIsFloat64(agg.items.map((i) => i.品牌)),
    季节: columnIsFloat64(agg.items.map((i) => i.季节)),
    设计师: columnIsFloat64(agg.items.map((i) => i.设计师)),
    上市天数: columnIsFloat64(agg.items.map((i) => i.上市天数)),
    未成交天数: columnIsFloat64(agg.items.map((i) => i.未成交天数)),
  };
  const extraF64 = new Map(
    extraFields.map((f) => [f, columnIsFloat64(agg.items.map((i) => i.extra[f]))]),
  );

  const wb = new ExcelJS.Workbook();
  // 同一张图(placeholder)只注册一次
  const imageIds = new Map<Uint8Array, number>();
  const imageId = (png: Uint8Array): number => {
    let id = imageIds.get(png);
    if (id === undefined) {
      id = wb.addImage({ buffer: png.buffer as never, extension: 'png' });
      imageIds.set(png, id);
    }
    return id;
  };

  // 透传字段分两组:固定排位组(销售量之后)+ 默认组(可售库存之后)
  const pinnedAfterQty = PINNED_AFTER_QTY.filter((f) => extraFields.includes(f));
  const extraRest = extraFields.filter((f) => !pinnedAfterQty.includes(f));
  const headers = [
    ...HEAD_KNOWN, ...extraRest,
    HEAD_TAIL[0], ...pinnedAfterQty, HEAD_TAIL[1], HEAD_TAIL[2],
    '商品销售量趋势图',
  ];
  const SALES_QTY_COL = 11 + extraRest.length;
  const SALES_AMT_COL = SALES_QTY_COL + 1 + pinnedAfterQty.length;
  const PROFIT_COL = SALES_AMT_COL + 1;
  /** 含固定排位透传列的尾段列宽:[销售量, ...pinned, 总销售金额, 盈利金额] */
  const tailWidths = [
    HEAD_TAIL_WIDTHS[0], ...Array(pinnedAfterQty.length).fill(EXTRA_WIDTH),
    HEAD_TAIL_WIDTHS[1], HEAD_TAIL_WIDTHS[2],
  ];

  /**
   * 元数据 13+extra 列的写入(三个 sheet 共用)。
   * 差异:货号对齐方式;Sheet3 零销量的销售量不染灰(pipeline.py:725 只有 if 没有 else)
   */
  function writeMetaCells(
    ws: ExcelJS.Worksheet, r: number, item: BaseItem,
    huohaoAlign: Partial<ExcelJS.Alignment>, qtyGrayWhenZero = true,
  ) {
    const set = (c: number, value: Cell, align = CENTER) => {
      const cell = ws.getCell(r, c);
      cell.value = value as ExcelJS.CellValue;
      cell.alignment = align;
      return cell;
    };
    set(1, item.货号, huohaoAlign);
    set(2, asStr(item.品类, f64.品类));
    set(3, asStr(item.品牌, f64.品牌));
    set(4, asStr(item.季节, f64.季节));
    set(5, asStr(item.设计师, f64.设计师));
    set(6, asStr(item.上市天数, f64.上市天数));
    set(7, asStr(item.未成交天数, f64.未成交天数));
    const rcell = set(8, asNum(item.销进率));
    rcell.numFmt = '0%';
    if (item.销进率 != null && item.销进率 === 0) rcell.font = { color: { argb: GRAY } };
    set(9, asNum(item.库存价值));
    set(10, asNum(item.可售库存));
    const setExtra = (c: number, fname: string) => {
      const raw = item.extra[fname];
      if (typeof raw === 'number') set(c, raw);
      else set(c, asStr(raw, extraF64.get(fname)!));
    };
    extraRest.forEach((fname, ei) => setExtra(11 + ei, fname));
    pinnedAfterQty.forEach((fname, pi) => setExtra(SALES_QTY_COL + 1 + pi, fname));
    const qcell = set(SALES_QTY_COL, item.销售量);
    if (item.销售量 > 0) qcell.font = { color: { argb: RED }, bold: true };
    else if (qtyGrayWhenZero) qcell.font = { color: { argb: GRAY } };
    const acell = set(SALES_AMT_COL, item.总销售金额);
    acell.numFmt = '#,##0.00';
    const pcell = set(PROFIT_COL, asNum(item.盈利金额));
    pcell.numFmt = '#,##0.00';
    if (item.盈利金额 != null) {
      if (item.盈利金额 > 0) pcell.font = { color: { argb: RED } };
      else if (item.盈利金额 < 0) pcell.font = { color: { argb: GREEN } };
    }
  }

  /** Sheet 1 / Sheet 2 共用的图表 sheet 写入 */
  function writeChartSheet(opts: {
    name: string; widths: number[]; imgKey: 'sm' | 'dt'; rowH: number;
  }) {
    const ws = wb.addWorksheet(opts.name);
    headers.forEach((h, j0) => {
      const c = ws.getCell(1, j0 + 1);
      c.value = h;
      c.fill = HEADER_FILL;
      c.font = HEADER_FONT;
      c.alignment = CENTER;
      c.border = BORDER;
      ws.getColumn(j0 + 1).width = opts.widths[j0];
    });
    ws.getRow(1).height = 28;
    const trendCol = headers.length;

    let r = 2;
    for (const { item, origIdx } of sorted) {
      writeMetaCells(ws, r, item, LEFT);
      const trendCell = ws.getCell(r, trendCol);
      trendCell.alignment = CENTER;
      const im = images.get(origIdx);
      if (im) addAnchoredImage(ws, imageId(im[opts.imgKey]), trendCol, r);
      const values = valuesOf(item);
      trendCell.note = {
        texts: [{ text: buildComment(values, dateRange) }],
      };
      ws.getRow(r).height = opts.rowH;
      if (r % 2 === 0) {
        for (let j = 1; j <= headers.length; j++) ws.getCell(r, j).fill = ALT_FILL;
      }
      for (let j = 1; j <= headers.length; j++) ws.getCell(r, j).border = BORDER;
      r += 1;
    }
    ws.views = [{ state: 'frozen', xSplit: 1, ySplit: 1 }]; // freeze B2
    ws.autoFilter = `A1:${colLetter(headers.length)}${r - 1}`;
    return r - 2;
  }

  // ---------- Sheet 1:商品销售趋势 ----------
  const widths1 = [...HEAD_KNOWN_WIDTHS, ...Array(extraRest.length).fill(EXTRA_WIDTH), ...tailWidths, 54];
  const rows1 = writeChartSheet({
    name: '商品销售趋势', widths: widths1, imgKey: 'sm', rowH: 62,
  });
  log(`  写入 sheet「商品销售趋势」 ${rows1} 行 × ${headers.length} 列`, 'normal', 4);

  // ---------- Sheet 2:款趋势明细图 ----------
  const widths2 = [...DETAIL_KNOWN_WIDTHS, ...Array(extraRest.length).fill(EXTRA_WIDTH), ...tailWidths, 92];
  writeChartSheet({
    name: '款趋势明细图', widths: widths2, imgKey: 'dt', rowH: 148,
  });
  log('  写入 sheet「款趋势明细图」', 'normal', 4);

  // ---------- Sheet 3:款日销量明细 ----------
  const ws3 = wb.addWorksheet('款日销量明细');
  const metaHeaders3 = [...HEAD_KNOWN, ...extraRest, HEAD_TAIL[0], ...pinnedAfterQty, HEAD_TAIL[1], HEAD_TAIL[2]];
  const widths3 = [...META3_KNOWN_WIDTHS, ...Array(extraRest.length).fill(EXTRA_WIDTH), ...tailWidths];
  metaHeaders3.forEach((h, j0) => {
    ws3.getCell(1, j0 + 1).value = h;
    ws3.getColumn(j0 + 1).width = widths3[j0];
  });
  const dayStartCol = metaHeaders3.length + 1;
  dateRange.forEach((d, k) => {
    ws3.getCell(1, dayStartCol + k).value = d.slice(5);
    ws3.getColumn(dayStartCol + k).width = 7;
  });
  const totalCol = dayStartCol + dateRange.length;
  ws3.getCell(1, totalCol).value = '合计';
  ws3.getColumn(totalCol).width = 10;
  for (let c = 1; c <= totalCol; c++) {
    const cell = ws3.getCell(1, c);
    cell.font = HEADER_FONT;
    cell.fill = HEADER_FILL;
    cell.alignment = CENTER;
    cell.border = BORDER;
  }
  ws3.getRow(1).height = 26;

  let r3 = 2;
  for (const { item } of sorted) {
    writeMetaCells(ws3, r3, item, CENTER, false); // Sheet3 货号居中,零销量不染灰
    const values = valuesOf(item);
    let totalV = 0;
    for (let k = 0; k < dateRange.length; k++) {
      const v = values[k] ?? 0;
      totalV += v;
      const cell = ws3.getCell(r3, dayStartCol + k);
      cell.value = v !== 0 ? v : null;
      cell.alignment = CENTER;
      if (v > 0) cell.font = { color: { argb: RED } };
    }
    const tcell = ws3.getCell(r3, totalCol);
    tcell.value = totalV;
    tcell.alignment = CENTER;
    tcell.font = totalV > 0 ? { bold: true, color: { argb: RED } } : { bold: true };
    r3 += 1;
  }
  ws3.views = [{ state: 'frozen', xSplit: 1, ySplit: 1 }];
  ws3.autoFilter = `A1:${colLetter(totalCol)}${r3 - 1}`;
  log(`  写入 sheet「款日销量明细」 ${r3 - 2} 行 × ${dateRange.length} 列`, 'normal', 4);

  const buf = await wb.xlsx.writeBuffer();
  return new Uint8Array(buf as ArrayBuffer);
}
