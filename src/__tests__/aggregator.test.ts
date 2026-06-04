// aggregator.test.ts — JS 聚合 vs pandas 黄金基准(aggregator.golden.json)
import { readFileSync } from 'node:fs';
import { resolve } from 'node:path';
import { describe, expect, it } from 'vitest';
import { readSales, readZhixiao } from '../pipeline/reader';
import { aggregate } from '../pipeline/aggregator';
import type { BaseItem } from '../pipeline/aggregator';
import readerGolden from './golden/reader.golden.json';
import golden from './golden/aggregator.golden.json';

const BASE_DIR = resolve(process.cwd(), '..', '趋势分析工具 -最终app版本');

function load(name: string): Uint8Array {
  return new Uint8Array(readFileSync(resolve(BASE_DIR, name)));
}

const zhixiao = readZhixiao(load(readerGolden.zhixiao.file));
const sales = readSales(load(readerGolden.sales.file), readerGolden.sales.file);
const agg = aggregate(zhixiao, sales);

/** 数值字段近似比较(浮点),其余精确比较 */
function expectValue(actual: unknown, expected: unknown, label: string) {
  if (typeof expected === 'number' && typeof actual === 'number' && !Number.isInteger(expected)) {
    expect(actual, label).toBeCloseTo(expected, 6);
  } else {
    expect(actual, label).toEqual(expected);
  }
}

describe('aggregate(聚合 + 字段透传)', () => {
  it('时间窗与 pandas 一致', () => {
    expect(agg.startDate).toBe(golden.startDate);
    expect(agg.endDate).toBe(golden.endDate);
    expect(agg.windowDays).toBe(golden.windowDays);
    expect(agg.dateRange.length).toBe(golden.windowDays);
    expect(agg.dateRange[0]).toBe(golden.startDate);
    expect(agg.dateRange[agg.dateRange.length - 1]).toBe(golden.endDate);
  });

  it('货号数 / 有销量款数一致', () => {
    expect(agg.itemsTotal).toBe(golden.itemsTotal);
    expect(agg.itemsWithSales).toBe(golden.itemsWithSales);
  });

  it('透传字段列表一致(含顺序)', () => {
    expect(agg.extraFields).toEqual(golden.extraFields);
  });

  it('透视表货号数一致', () => {
    expect(agg.dailyByItem.size).toBe(golden.pivotItemCount);
  });

  it('销售量 / 总销售金额 / 盈利金额 / 销进率 全表合计一致', () => {
    let qty = 0, amt = 0, profit = 0, xjl = 0;
    for (const it of agg.items) {
      qty += it.销售量;
      amt += it.总销售金额;
      profit += it.盈利金额 ?? 0;
      xjl += it.销进率 ?? 0;
    }
    expect(qty).toBe(golden.sumQty);
    expect(amt).toBeCloseTo(golden.sumAmt, 2);
    if (golden.sum盈利金额 != null) expect(profit).toBeCloseTo(golden.sum盈利金额, 2);
    if (golden.sum销进率 != null) expect(xjl).toBeCloseTo(golden.sum销进率, 4);
  });

  it('抽查行逐字段一致(含透传 extra)', () => {
    golden.sampleIndexes.forEach((idx, i) => {
      const expected = golden.sampleItems[i];
      const actual = agg.items[idx];
      for (const [k, v] of Object.entries(expected)) {
        if (k === 'extra') {
          for (const [ek, ev] of Object.entries(v as Record<string, unknown>)) {
            expectValue(actual.extra[ek], ev, `items[${idx}].extra.${ek}`);
          }
        } else {
          expectValue(actual[k as keyof BaseItem], v, `items[${idx}].${k}`);
        }
      }
    });
  });

  it('透视行抽查(行合计 / 非零天数 / 首末非零点)一致', () => {
    for (const check of golden.pivotChecks) {
      const row = agg.dailyByItem.get(check.货号);
      expect(row, `pivot ${check.货号}`).toBeDefined();
      let sum = 0, nonZero = 0;
      let first: { date: string; value: number } | null = null;
      let last: { date: string; value: number } | null = null;
      row!.forEach((v, i) => {
        sum += v;
        if (v !== 0) {
          nonZero++;
          const point = { date: agg.dateRange[i], value: v };
          if (!first) first = point;
          last = point;
        }
      });
      expect(sum).toBeCloseTo(check.rowSum, 6);
      expect(nonZero).toBe(check.nonZeroDays);
      expect(first).toEqual(check.firstNonZero);
      expect(last).toEqual(check.lastNonZero);
    }
  });
});
