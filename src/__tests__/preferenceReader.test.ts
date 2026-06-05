// preferenceReader.test.ts — JS 偏好读取 vs pandas 黄金基准(preference_reader.golden.json)
import { readFileSync } from 'node:fs';
import { resolve } from 'node:path';
import { describe, expect, it } from 'vitest';
import { loadPreference } from '../pipeline/preferenceReader';
import { msToDateStr } from '../pipeline/reader';
import golden from './golden/preference_reader.golden.json';

const BASE_DIR = resolve(process.cwd(), '..', '趋势分析工具 -最终app版本');

const data = loadPreference(
  new Uint8Array(readFileSync(resolve(BASE_DIR, golden.file))),
);

describe('loadPreference(读取 + 列规范化)', () => {
  it('sheet 元数据与 pandas 一致(嗅探表头/列数/行数/保留)', () => {
    expect(data.sheetNames).toEqual(golden.sheetNames);
    expect(data.sheetMeta).toEqual(golden.sheetMeta);
  });

  it('合并行数与冲突丢列一致(毛值销售量列被净值顶掉)', () => {
    expect(data.rawRowCount).toBe(golden.rawRows);
    expect(data.droppedConflicts).toEqual(golden.droppedConflicts);
  });

  it('规范化后的列名与列序一致', () => {
    expect(data.columns).toEqual(golden.columns);
  });

  it('各列非空计数一致(验证模糊映射没接错线)', () => {
    for (const [col, expected] of Object.entries(golden.nonNullCounts)) {
      let actual: number;
      if (col === '销售量') actual = data.qty.filter((v) => !Number.isNaN(v)).length;
      else if (col === '销售金额') actual = data.amt.filter((v) => !Number.isNaN(v)).length;
      else if (col === '下单时间') actual = data.orderMs!.filter((v) => !Number.isNaN(v)).length;
      else if (data.cols.has(col)) actual = data.cols.get(col)!.filter((v) => v != null).length;
      else continue; // 无关列(销退单ID)只保留列名,不存数据
      expect(actual, col).toBe(expected);
    }
  });

  it('净销售量/净销售金额总和一致', () => {
    let q = 0, a = 0;
    for (const v of data.qty) if (!Number.isNaN(v)) q += v;
    for (const v of data.amt) if (!Number.isNaN(v)) a += v;
    expect(q).toBeCloseTo(golden.sumQty, 6);
    expect(a).toBeCloseTo(golden.sumAmt, 2);
  });

  it('客户去重数与时间范围一致', () => {
    const customers = new Set<string>();
    for (const v of data.cols.get('客户名称')!) if (v != null) customers.add(String(v));
    expect(customers.size).toBe(golden.uniqueCustomers);
    let min = Infinity, max = -Infinity;
    for (const v of data.orderMs!) {
      if (Number.isNaN(v)) continue;
      if (v < min) min = v;
      if (v > max) max = v;
    }
    expect(msToDateStr(min)).toBe(golden.dateMin);
    expect(msToDateStr(max)).toBe(golden.dateMax);
  });

  it('首行抽查(规范化后各字段值)', () => {
    for (const [col, expected] of Object.entries(golden.row0)) {
      if (col === '销退单ID') continue; // 无关列不存数据
      let actual: unknown;
      if (col === '销售量') actual = Number.isNaN(data.qty[0]) ? null : data.qty[0];
      else if (col === '销售金额') actual = Number.isNaN(data.amt[0]) ? null : data.amt[0];
      else if (col === '下单时间') {
        actual = Number.isNaN(data.orderMs![0]) ? null
          : new Date(data.orderMs![0]).toISOString().replace('T', ' ').replace(/\.\d+Z$/, '');
      } else actual = data.cols.get(col)![0];
      if (typeof expected === 'number' && !Number.isInteger(expected)) {
        expect(actual, col).toBeCloseTo(expected, 6);
      } else if (col === '下单时间') {
        // pandas str(Timestamp) "2026-01-02 00:00:00" vs 我们的 ISO 变体,归一化比较
        expect(String(actual), col).toBe(String(expected));
      } else if (typeof expected === 'number') {
        expect(Number(actual), col).toBe(expected);
      } else {
        expect(actual == null ? null : String(actual), col).toBe(expected);
      }
    }
  });
});
