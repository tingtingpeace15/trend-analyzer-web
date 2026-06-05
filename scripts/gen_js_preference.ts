// 生成 JS 版偏好分析产物(html + xlsx),供 diff 对照 Python 基准。
// 用法: npx vite-node scripts/gen_js_preference.ts
// 输出: baseline/js-pref/客户偏好分析报告.html / 客户偏好分析数据.xlsx
import { mkdirSync, readFileSync, writeFileSync } from 'node:fs';
import { resolve } from 'node:path';
import { loadPreference } from '../src/pipeline/preferenceReader';
import { analyzePreference } from '../src/pipeline/preferenceAnalyze';
import { buildPreferenceHtml } from '../src/pipeline/preferenceHtml';
import { buildPreferenceExcel } from '../src/pipeline/preferenceExcel';

const ROOT = resolve(import.meta.dirname, '..');
const BASE_DIR = resolve(ROOT, '..', '趋势分析工具 -最终app版本');
const OUT_DIR = resolve(ROOT, 'baseline', 'js-pref');
mkdirSync(OUT_DIR, { recursive: true });

const t0 = Date.now();
const data = loadPreference(new Uint8Array(readFileSync(
  resolve(BASE_DIR, '各商品客户拿货历史_5.12.xlsx'))));
const R = analyzePreference(data);
writeFileSync(resolve(OUT_DIR, '客户偏好分析报告.html'), buildPreferenceHtml(R), 'utf-8');
const xlsx = await buildPreferenceExcel(data);
writeFileSync(resolve(OUT_DIR, '客户偏好分析数据.xlsx'), xlsx);
console.log(`done in ${((Date.now() - t0) / 1000).toFixed(1)}s, xlsx ${(xlsx.length / 1024).toFixed(0)} KB`);
