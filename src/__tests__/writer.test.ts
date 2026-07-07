// writer.test.ts — JS writer 输出 xlsx,回读后对照 pandas/openpyxl 基准
// (writer.golden.json:排序、Sheet1 单元格值含 pandas 字符串语义)+ 结构断言。
import { readFileSync } from 'node:fs';
import { resolve } from 'node:path';
import ExcelJS from 'exceljs';
import { beforeAll, describe, expect, it } from 'vitest';
import { readSales, readZhixiao } from '../pipeline/reader';
import { aggregate } from '../pipeline/aggregator';
import { writeExcel } from '../pipeline/writer';
import type { ItemImages } from '../pipeline/writer';
import readerGolden from './golden/reader.golden.json';
import golden from './golden/writer.golden.json';

const BASE_DIR = resolve(process.cwd(), '..', '趋势分析工具 -最终app版本');

// 1×1 透明 PNG(测试里图片内容无关紧要,结构/锚点才是断言点)
const TINY_PNG = new Uint8Array(Buffer.from(
  'iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJAAAADUlEQVR42mNk+M9QDwADhgGAWjR9awAAAABJRU5ErkJggg==',
  'base64',
));

let wb: ExcelJS.Workbook;
let agg: ReturnType<typeof aggregate>;

beforeAll(async () => {
  const zhixiao = readZhixiao(new Uint8Array(readFileSync(resolve(BASE_DIR, readerGolden.zhixiao.file))));
  const sales = readSales(new Uint8Array(readFileSync(resolve(BASE_DIR, readerGolden.sales.file))));
  agg = aggregate(zhixiao, sales);
  const images = new Map<number, ItemImages>(
    agg.items.map((_, i) => [i, { sm: TINY_PNG, dt: TINY_PNG }]),
  );
  const bytes = await writeExcel(agg, images, () => {});
  wb = new ExcelJS.Workbook();
  await wb.xlsx.load(bytes.buffer as ArrayBuffer);
}, 180_000);

describe('writeExcel(3 个 sheet)', () => {
  it('sheet 名与顺序', () => {
    expect(wb.worksheets.map((w) => w.name)).toEqual(['商品销售趋势', '款趋势明细图', '款日销量明细']);
  });

  it('Sheet1 表头与列宽', () => {
    const ws = wb.getWorksheet('商品销售趋势')!;
    const headers = golden.headers;
    headers.forEach((h, j0) => {
      expect(ws.getCell(1, j0 + 1).value, `header ${j0 + 1}`).toBe(h);
    });
    // 列宽抽查:货号 18 / 品类 14 / 趋势图 54;透传列 12
    expect(ws.getColumn(1).width).toBe(18);
    expect(ws.getColumn(2).width).toBe(14);
    expect(ws.getColumn(11).width).toBe(12); // 是否盈利(extra)
    expect(ws.getColumn(headers.length).width).toBe(54);
    expect(ws.getRow(1).height).toBe(28);
  });

  it('排序结果与 pandas 一致(首 10 + 末 5)', () => {
    const ws = wb.getWorksheet('商品销售趋势')!;
    const first10 = Array.from({ length: 10 }, (_, i) => ws.getCell(2 + i, 1).value);
    expect(first10).toEqual(golden.sortedKeysFirst10);
    const total = agg.itemsTotal;
    const last5 = Array.from({ length: 5 }, (_, i) => ws.getCell(2 + total - 5 + i, 1).value);
    expect(last5).toEqual(golden.sortedKeysLast5);
  });

  it('Sheet1 抽查行逐单元格一致(含 pandas 字符串/数值语义)', () => {
    const ws = wb.getWorksheet('商品销售趋势')!;
    for (const [idxStr, cells] of Object.entries(golden.sheet1Rows)) {
      const r = 2 + Number(idxStr);
      (cells as (string | number | null)[]).forEach((expected, j0) => {
        const actual = ws.getCell(r, j0 + 1).value;
        const label = `row[${idxStr}] col${j0 + 1}`;
        if (typeof expected === 'number' && !Number.isInteger(expected)) {
          expect(actual, label).toBeCloseTo(expected, 6);
        } else if (expected === null) {
          expect(actual, label).toBeNull();
        } else {
          expect(actual, label).toBe(expected);
        }
      });
    }
  });

  it('数字格式与染色(销进率 0% / 金额 #,##0.00 / 红绿灰)', () => {
    const ws = wb.getWorksheet('商品销售趋势')!;
    const extra = 1; // 是否盈利
    const r = 2; // 第一行:销售量 3186 > 0,盈利 43270.62 > 0
    expect(ws.getCell(r, 8).numFmt).toBe('0%');
    expect(ws.getCell(r, 12 + extra).numFmt).toBe('#,##0.00');
    expect(ws.getCell(r, 13 + extra).numFmt).toBe('#,##0.00');
    const qFont = ws.getCell(r, 11 + extra).font;
    expect(qFont?.color?.argb).toBe('00E74C3C');
    expect(qFont?.bold).toBe(true);
    expect(ws.getCell(r, 13 + extra).font?.color?.argb).toBe('00E74C3C');
    // 零销量行(最后一行):销售量灰字
    const rz = 1 + agg.itemsTotal;
    expect(ws.getCell(rz, 11 + extra).font?.color?.argb).toBe('00999999');
  });

  it('冻结窗格 B2 + 自动筛选(3 个 sheet)', () => {
    for (const ws of wb.worksheets) {
      const v = ws.views[0];
      expect(v?.state, ws.name).toBe('frozen');
      expect((v as ExcelJS.WorksheetViewFrozen).xSplit, ws.name).toBe(1);
      expect((v as ExcelJS.WorksheetViewFrozen).ySplit, ws.name).toBe(1);
      expect(ws.autoFilter, ws.name).toBeTruthy();
    }
  });

  it('嵌图:零销量行不嵌图,有销量行仍嵌图', () => {
    const ws = wb.getWorksheet('商品销售趋势')!;
    const imgs = ws.getImages();
    expect(imgs.length).toBe(agg.items.filter((item) => item.销售量 > 0).length);
    // 全部有销量行用 TINY_PNG → 媒体应只有 1 份(写入时按字节引用去重)
    expect(wb.model.media?.length ?? 0).toBe(1);
  });

  it('嵌图锚点:原生 EMU 偏移 ±2px,跨满整格(回归:ExcelJS 小数坐标会压扁图)', () => {
    const ws = wb.getWorksheet('商品销售趋势')!;
    const trendCol0 = golden.headers.length - 1; // 0-based 趋势图列
    const img = ws.getImages()[0];
    const { tl, br } = img.range as unknown as {
      tl: { nativeCol: number; nativeColOff: number; nativeRow: number; nativeRowOff: number };
      br: { nativeCol: number; nativeColOff: number; nativeRow: number; nativeRowOff: number };
    };
    const PAD = 2 * 9525; // 2px(同 openpyxl pixels_to_EMU(2))
    expect(tl.nativeCol).toBe(trendCol0);
    expect(tl.nativeColOff).toBe(PAD);
    expect(tl.nativeRow).toBe(1); // 第 2 行(0-based)
    expect(tl.nativeRowOff).toBe(PAD);
    expect(br.nativeCol).toBe(trendCol0 + 1);
    expect(br.nativeColOff).toBe(-PAD);
    expect(br.nativeRow).toBe(2);
    expect(br.nativeRowOff).toBe(-PAD);
  });

  it('趋势列批注(有销量含峰值,零销量显示无销售且不加批注)', () => {
    const ws = wb.getWorksheet('商品销售趋势')!;
    const trendCol = golden.headers.length;
    const noteTop = ws.getCell(2, trendCol).note;
    const textOf = (n: ExcelJS.Comment | string) =>
      typeof n === 'string' ? n : (n.texts ?? []).map((t) => t.text).join('');
    expect(textOf(noteTop)).toContain('日销售量: 3186');
    expect(textOf(noteTop)).toContain('峰值:');
    const zeroCell = ws.getCell(1 + agg.itemsTotal, trendCol);
    expect(zeroCell.value).toBe('无销售');
    expect(zeroCell.note).toBeUndefined();
    expect(zeroCell.font?.color?.argb).toBe('00999999');
  });

  it('A价 固定排位:出现时插在销售量之后(2026-06-05 需求)', async () => {
    // 在真实滞销表上注入合成 A价 列 → 透传应分组:是否盈利留默认区,A价 钉在销售量后
    const zhixiao = readZhixiao(new Uint8Array(readFileSync(resolve(BASE_DIR, readerGolden.zhixiao.file))));
    // 用隐形变体「À价」注入:验证宽容匹配 + 输出列头规范化为 A价
    zhixiao.columns.push('À价');
    zhixiao.rows.forEach((r, i) => { r.cells['À价'] = 100 + i; });
    const sales = readSales(new Uint8Array(readFileSync(resolve(BASE_DIR, readerGolden.sales.file))));
    const agg2 = aggregate(zhixiao, sales);
    expect(agg2.extraFields).toEqual(['是否盈利', 'À价']);
    // 只取前 20 款,加速写入
    agg2.items = agg2.items.slice(0, 20);
    const images = new Map<number, ItemImages>(
      agg2.items.map((_, i) => [i, { sm: TINY_PNG, dt: TINY_PNG }]),
    );
    const wb2 = new ExcelJS.Workbook();
    await wb2.xlsx.load((await writeExcel(agg2, images, () => {})).buffer as ArrayBuffer);

    const ws = wb2.getWorksheet('商品销售趋势')!;
    const hdr = Array.from({ length: 16 }, (_, j) => ws.getCell(1, j + 1).value);
    expect(hdr).toEqual([
      '货号', '品类', '品牌', '季节', '设计师', '上市天数', '未成交天数',
      '销进率', '库存价值', '可售库存', '是否盈利',
      '销售量', 'A价', '总销售金额', '盈利金额', '商品销售量趋势图',
    ]);
    expect(ws.getColumn(13).width).toBe(12); // A价 透传列宽
    // 数值原样透传(排序后第一行对应原始某行的 100+i)
    expect(typeof ws.getCell(2, 13).value).toBe('number');
    // Sheet3 同样排位
    const ws3 = wb2.getWorksheet('款日销量明细')!;
    expect(ws3.getCell(1, 12).value).toBe('销售量');
    expect(ws3.getCell(1, 13).value).toBe('A价');
    expect(ws3.getCell(1, 14).value).toBe('总销售金额');
  }, 120_000);

  it('Sheet3:日期列 + 合计列', () => {
    const ws3 = wb.getWorksheet('款日销量明细')!;
    const metaLen = golden.headers.length - 1; // 无趋势图列
    expect(ws3.getCell(1, metaLen + 1).value).toBe(agg.dateRange[0].slice(5));
    const totalCol = metaLen + agg.windowDays + 1;
    expect(ws3.getCell(1, totalCol).value).toBe('合计');
    expect(ws3.getColumn(totalCol).width).toBe(10);
    // 第一行(销量最高款)合计 = 3186
    expect(ws3.getCell(2, totalCol).value).toBe(3186);
    expect(ws3.getCell(2, totalCol).font?.bold).toBe(true);
    expect(ws3.getCell(2, totalCol).font?.color?.argb).toBe('00E74C3C');
  });
});
