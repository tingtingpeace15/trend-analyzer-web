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

ts = (
    "// 自动生成,勿手改 — 由 scripts/extract_html_template.py 从\n"
    "// preference_pipeline.py 的 _build_html 模板机械提取。Python 版模板变更后重跑脚本。\n"
    "// 占位符:__PREFERENCE_DATA__(R 的 JSON)、__TOP_N__(30)。\n"
    f"export const PREFERENCE_HTML_TEMPLATE: string = {json.dumps(template, ensure_ascii=False)};\n"
)
OUT.write_text(ts, encoding="utf-8")
print(f"written: {OUT} ({len(template)} chars)")
