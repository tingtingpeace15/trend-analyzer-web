// readerAliases.test.ts — 拿货历史列名别名(不依赖基准目录,内存里造 xlsx)
import * as XLSX from 'xlsx';
import { describe, expect, it } from 'vitest';
import { msToDateStr, readSales } from '../pipeline/reader';
import { loadPreference } from '../pipeline/preferenceReader';

/** 造一个单 sheet 的拿货历史:表头 + 两行数据,补足到 ≥10 列 */
function makeSales(timeHeader: string, extraHeaders: string[] = [], extraCells: unknown[] = []): Uint8Array {
  const header = ['客户名称', '货号', timeHeader, '净销售量', '净销售金额', ...extraHeaders];
  while (header.length < 10) header.push(`备注${header.length}`);
  const row = (key: string, date: string, qty: number, amt: number) => {
    const r: unknown[] = ['张三', key, date, qty, amt, ...extraCells];
    while (r.length < header.length) r.push('');
    return r;
  };
  const ws = XLSX.utils.aoa_to_sheet([
    header,
    row('A001', '2026-10-01 10:00:00', 2, 200),
    row('A002', '2026/10/15', 1, 99),
  ]);
  const wb = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(wb, ws, 'Sheet1');
  return new Uint8Array(XLSX.write(wb, { type: 'array', bookType: 'xlsx' }));
}

describe('readSales 列名别名:创建时间 = 下单时间', () => {
  it('只有「创建时间」列时当作下单时间读取', () => {
    const t = readSales(makeSales('创建时间'), 'x.xlsx');
    expect(t.columns).toContain('下单时间');
    expect(t.records.map((r) => r.date)).toEqual(['2026-10-01', '2026-10-15']);
  });

  it('列名前后有空格也认', () => {
    const t = readSales(makeSales(' 创建时间 '), 'x.xlsx');
    expect(t.records).toHaveLength(2);
  });

  it('「下单时间」和「创建时间」同时存在时以下单时间为准', () => {
    const t = readSales(makeSales('下单时间', ['创建时间'], ['2020-01-01']), 'x.xlsx');
    expect(t.records.map((r) => r.date)).toEqual(['2026-10-01', '2026-10-15']);
  });

  it('两者都没有时报缺少下单时间', () => {
    expect(() => readSales(makeSales('开单时间'), 'x.xlsx')).toThrow('下单时间');
  });
});

describe('loadPreference 列名别名:创建时间 = 下单时间', () => {
  const dates = (bytes: Uint8Array) =>
    Array.from(loadPreference(bytes).orderMs ?? []).map((ms) => msToDateStr(ms));

  it('只有「创建时间」列时当作下单时间读取', () => {
    expect(dates(makeSales('创建时间'))).toEqual(['2026-10-01', '2026-10-15']);
  });

  it('「创建时间:」带冒号、前后空格也认', () => {
    expect(dates(makeSales(' 创建时间: '))).toEqual(['2026-10-01', '2026-10-15']);
  });

  it('「下单时间」和「创建时间」同时存在时以下单时间为准,创建时间原样保留', () => {
    const d = loadPreference(makeSales('下单时间', ['创建时间'], ['2020-01-01']));
    expect(Array.from(d.orderMs ?? []).map((ms) => msToDateStr(ms))).toEqual(['2026-10-01', '2026-10-15']);
    expect(d.columns.filter((c) => c === '下单时间')).toHaveLength(1);
    expect(d.columns).toContain('创建时间');
  });

  it('含「日期」的列也优先于创建时间(沿用原规则)', () => {
    expect(dates(makeSales('下单日期', ['创建时间'], ['2020-01-01']))).toEqual(['2026-10-01', '2026-10-15']);
  });

  it('都没有时 orderMs 为 null(时间类统计为空,不报错)', () => {
    expect(loadPreference(makeSales('开单时间')).orderMs).toBeNull();
  });
});
