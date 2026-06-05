# -*- coding: utf-8 -*-
"""偏好分析聚合阶段黄金基准:直接 import 真实 preference_pipeline 模块,
调它的 _load + _analyze,把 R 整体 dump 成 JSON——零转写误差。

输出: src/__tests__/golden/preference_analyze.golden.json
用法: python3 scripts/gen_golden_preference_analyze.py
"""
import importlib.util
import json
import sys
from pathlib import Path

HERE = Path(__file__).resolve().parent
BASE = HERE.parent.parent / "趋势分析工具 -最终app版本"
OUT = HERE.parent / "src" / "__tests__" / "golden" / "preference_analyze.golden.json"
INPUT = BASE / "各商品客户拿货历史_5.12.xlsx"

spec = importlib.util.spec_from_file_location(
    "pref_pipeline", BASE / "backend" / "preference_pipeline.py")
mod = importlib.util.module_from_spec(spec)
sys.modules["pref_pipeline"] = mod
spec.loader.exec_module(mod)
mod._lazy_load_pd_np()


def log(text, kind="normal", step=1):
    pass


df = mod._load(INPUT, log)
R = mod._analyze(df, log)

golden = {"_comment": "由 scripts/gen_golden_preference_analyze.py 生成(直接调用真实 _analyze),勿手改。",
          "R": R}
OUT.write_text(json.dumps(golden, ensure_ascii=False, indent=1), encoding="utf-8")
print(f"written: {OUT}  ({OUT.stat().st_size/1024:.0f} KB)")
print("summary:", json.dumps(R["summary"], ensure_ascii=False))
print("keys:", list(R.keys()))
print("lens:", {k: len(v) for k, v in R.items() if isinstance(v, list)})
