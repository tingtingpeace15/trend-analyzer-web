// reader.test.ts — JS 版 reader 对照 pandas 版黄金基准(reader.golden.json,
// 由 scripts/gen_golden_reader.py 用真实客户数据生成)。
// 任何一项不等都意味着跟 Python 版输出会产生分歧,必须修 reader 而不是改基准。
import { readFileSync } from 'node:fs';
import { resolve } from 'node:path';
import { describe, expect, it } from 'vitest';
import { readSales, readZhixiao } from '../pipeline/reader';
import golden from './golden/reader.golden.json';

const BASE_DIR = resolve(process.cwd(), '..', '趋势分析工具 -最终app版本');

function load(name: string): Uint8Array {
  return new Uint8Array(readFileSync(resolve(BASE_DIR, name)));
}

describe('readZhixiao(滞销商品表)', () => {
  const table = readZhixiao(load(golden.zhixiao.file));

  it('行数(按货号去重后)与 pandas 一致', () => {
    expect(table.rows.length).toBe(golden.zhixiao.rows);
  });

  it('列名与列序与 pandas 完全一致', () => {
    expect(table.columns).toEqual(golden.zhixiao.columns);
  });

  it('前 5 个货号一致', () => {
    expect(table.rows.slice(0, 5).map((r) => r.key)).toEqual(golden.zhixiao.first5货号);
  });
});

describe('readSales(拿货历史)', () => {
  const table = readSales(load(golden.sales.file), golden.sales.file);

  it('sheet 名单与逐 sheet 元数据(表头行/列数/行数/是否保留)一致', () => {
    expect(table.sheetNames).toEqual(golden.sales.sheetNames);
    expect(table.sheetMeta).toEqual(golden.sales.sheetMeta);
  });

  it('合并后行列数一致', () => {
    expect(table.rawRowCount).toBe(golden.sales.rawRows);
    expect(table.columns.length).toBe(golden.sales.cols);
    expect(table.columns).toEqual(golden.sales.columns);
  });

  it('清洗后(dropna 下单时间/货号)记录数一致', () => {
    expect(table.records.length).toBe(golden.sales.cleanRows);
  });

  it('时间窗一致', () => {
    const dates = table.records.map((r) => r.date);
    expect(dates.reduce((a, b) => (a < b ? a : b))).toBe(golden.sales.minDate);
    expect(dates.reduce((a, b) => (a > b ? a : b))).toBe(golden.sales.maxDate);
  });

  it('净销售量/净销售金额总和与 pandas 一致', () => {
    let qty = 0;
    let amt = 0;
    for (const r of table.records) {
      qty += r.qty;
      amt += r.amt;
    }
    expect(qty).toBeCloseTo(golden.sales.totalQty, 6);
    expect(amt).toBeCloseTo(golden.sales.totalAmt, 2);
  });

  it('货号去重数一致', () => {
    expect(new Set(table.records.map((r) => r.key)).size).toBe(golden.sales.uniqueItems);
  });
});
