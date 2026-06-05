// preferenceHtml.ts — 生成客户偏好分析报告 html(对应 _build_html,
// preference_pipeline.py:435-569)。模板由脚本从 Python 源码机械提取,
// JSON 用 pyJsonDumps(Python json.dumps 字节级复刻)→ 输出与 Python 版逐字节一致。
import { PREFERENCE_HTML_TEMPLATE } from './preferenceHtmlTemplate';
import { pyJsonDumps } from './pyjson';
import { TOP_N } from './preferenceAnalyze';
import type { AnalyzeResult } from './preferenceAnalyze';

export function buildPreferenceHtml(R: AnalyzeResult): string {
  // 用 split/join 替换:JSON 里可能含 "$",String.replace 的替换模式会踩坑
  return PREFERENCE_HTML_TEMPLATE
    .split('__PREFERENCE_DATA__').join(pyJsonDumps(R))
    .split('__TOP_N__').join(String(TOP_N));
}
