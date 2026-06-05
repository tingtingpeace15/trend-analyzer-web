// preferenceHtml.test.ts — JS 生成的 html 报告 vs 真实 Python 版产物,逐字节对比。
// 基准文件由 scripts/gen_baseline_preference.py 生成(不入库,本地跑)。
import { existsSync, readFileSync } from 'node:fs';
import { resolve } from 'node:path';
import { describe, expect, it } from 'vitest';
import { loadPreference } from '../pipeline/preferenceReader';
import { analyzePreference } from '../pipeline/preferenceAnalyze';
import { buildPreferenceHtml } from '../pipeline/preferenceHtml';
import { PREFERENCE_HTML_PATCHES } from '../pipeline/preferenceHtmlTemplate';
import readerGolden from './golden/preference_reader.golden.json';

const BASE_DIR = resolve(process.cwd(), '..', '趋势分析工具 -最终app版本');
const BASELINE = resolve(process.cwd(), 'baseline', 'python-pref', '客户偏好分析报告.html');

describe.skipIf(!existsSync(BASELINE))('buildPreferenceHtml vs Python 版(逐字节)', () => {
  it('html 与 Python 版完全一致(基准先打上有意修正补丁)', () => {
    const data = loadPreference(
      new Uint8Array(readFileSync(resolve(BASE_DIR, readerGolden.file))),
    );
    const html = buildPreferenceHtml(analyzePreference(data));
    let expected = readFileSync(BASELINE, 'utf-8');
    // 对 Python 基准应用同样的有意修正(extract_html_template.py 里记录的 PATCHES),
    // 保证除这些点外仍逐字节一致
    for (const [frm, to] of PREFERENCE_HTML_PATCHES) {
      expect(expected.includes(frm), `基准中应存在补丁源串: ${frm.slice(0, 50)}`).toBe(true);
      expected = expected.split(frm).join(to);
    }
    expect(html.length, 'length').toBe(expected.length);
    if (html !== expected) {
      // 找到第一个差异点,给出上下文方便定位
      let i = 0;
      while (html[i] === expected[i]) i++;
      throw new Error(
        `首个差异 @${i}:\n  js: …${html.slice(Math.max(0, i - 60), i + 60)}…\n  py: …${expected.slice(Math.max(0, i - 60), i + 60)}…`,
      );
    }
    expect(html).toBe(expected);
  });
});
