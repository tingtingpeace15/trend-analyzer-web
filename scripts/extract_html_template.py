# -*- coding: utf-8 -*-
"""从 preference_pipeline.py 机械提取 HTML 模板,生成 TS 常量文件。

避免人工转写 130 行模板引入误差;Python 版模板改了就重跑本脚本。
输出: src/pipeline/preferenceHtmlTemplate.ts(自动生成,勿手改)
"""
import json
from pathlib import Path

HERE = Path(__file__).resolve().parent
SRC = HERE.parent.parent / "趋势分析工具 -最终app版本" / "backend" / "preference_pipeline.py"
OUT = HERE.parent / "src" / "pipeline" / "preferenceHtmlTemplate.ts"

text = SRC.read_text(encoding="utf-8")
start_marker = "return r'''"
start = text.index(start_marker) + len(start_marker)
end = text.index("'''", start)
template = text[start:end]

# ── 对原版模板的「有意修正」(每条都要附原因;html 对照测试会把同样的补丁
#    打到 Python 基准上再比,保证除这些点外仍逐字节一致)─────────────────────
PATCHES = [
    # 2026-06-05 用户报告:客户等级分布饼图扇区原版用「客户数」——四分位分层
    # 人数天然各占 1/4,图永远近似四等分,与图例里的金额占比(85.5%…)不符。
    # 改用「总金额」,扇区与占比一致;负净额档(纯退货)钳到 0,扇区消失但图例保留。
    (
        "datasets:[{data:D.tier_summary.map(r=>r['客户数']),backgroundColor:D.tier_summary.map(r=>cm[r['等级']]||'#999')}]",
        "datasets:[{data:D.tier_summary.map(r=>Math.max(r['总金额'],0)),backgroundColor:D.tier_summary.map(r=>cm[r['等级']]||'#999')}]",
    ),
]
for frm, to in PATCHES:
    assert template.count(frm) == 1, f"补丁源串不唯一/不存在: {frm[:60]}…"
    template = template.replace(frm, to)

ts = (
    "// 自动生成,勿手改 — 由 scripts/extract_html_template.py 从\n"
    "// preference_pipeline.py 的 _build_html 模板机械提取(含脚本里记录的有意修正)。\n"
    "// Python 版模板变更后重跑脚本。占位符:__PREFERENCE_DATA__(R 的 JSON)、__TOP_N__(30)。\n"
    f"export const PREFERENCE_HTML_TEMPLATE: string = {json.dumps(template, ensure_ascii=False)};\n\n"
    "/** 相对 Python 原版模板的有意修正(测试用它把基准补丁后再对比) */\n"
    f"export const PREFERENCE_HTML_PATCHES: [string, string][] = {json.dumps(PATCHES, ensure_ascii=False)};\n"
)
OUT.write_text(ts, encoding="utf-8")
print(f"written: {OUT} ({len(template)} chars, {len(PATCHES)} patch)")
