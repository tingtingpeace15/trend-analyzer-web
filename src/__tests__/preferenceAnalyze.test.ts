// preferenceAnalyze.test.ts — JS _analyze 复刻 vs 真实 Python _analyze 的 R
// (preference_analyze.golden.json,由真实模块生成)。逐 key 深度对比。
import { readFileSync } from 'node:fs';
import { resolve } from 'node:path';
import { describe, expect, it } from 'vitest';
import { loadPreference } from '../pipeline/preferenceReader';
import { analyzePreference } from '../pipeline/preferenceAnalyze';
import { PyFloat } from '../pipeline/pyjson';
import readerGolden from './golden/preference_reader.golden.json';
import golden from './golden/preference_analyze.golden.json';

const BASE_DIR = resolve(process.cwd(), '..', '趋势分析工具 -最终app版本');

const data = loadPreference(
  new Uint8Array(readFileSync(resolve(BASE_DIR, readerGolden.file))),
);
const R = analyzePreference(data) as Record<string, unknown>;
const G = golden.R as Record<string, unknown>;

/** 深度对比:数字带浮点容差,其余精确(PyFloat 解包成 number) */
function deepCompare(rawActual: unknown, expected: unknown, path: string) {
  const actual = rawActual instanceof PyFloat ? rawActual.v : rawActual;
  if (typeof expected === 'number') {
    expect(typeof actual, path).toBe('number');
    if (Number.isInteger(expected)) expect(actual, path).toBe(expected);
    else expect(actual as number, path).toBeCloseTo(expected, 6);
    return;
  }
  if (Array.isArray(expected)) {
    expect(Array.isArray(actual), path).toBe(true);
    expect((actual as unknown[]).length, `${path}.length`).toBe(expected.length);
    expected.forEach((e, i) => deepCompare((actual as unknown[])[i], e, `${path}[${i}]`));
    return;
  }
  if (expected !== null && typeof expected === 'object') {
    for (const [k, v] of Object.entries(expected)) {
      deepCompare((actual as Record<string, unknown>)[k], v, `${path}.${k}`);
    }
    // key 集合也要一致(防 JS 多出/漏掉字段)
    expect(Object.keys(actual as object).sort(), `${path} keys`).toEqual(Object.keys(expected).sort());
    return;
  }
  expect(actual, path).toBe(expected);
}

describe('analyzePreference vs Python _analyze', () => {
  it('R 的 key 集合一致', () => {
    expect(Object.keys(R).sort()).toEqual(Object.keys(G).sort());
  });

  for (const key of Object.keys(golden.R)) {
    it(`R.${key} 一致`, () => {
      deepCompare(R[key], G[key], key);
    });
  }
});
