# -*- coding: utf-8 -*-
"""用真实 preference_pipeline 生成偏好分析基准产物(html + xlsx)。

输出: baseline/python-pref/客户偏好分析报告.html / 客户偏好分析数据.xlsx
用法: python3 scripts/gen_baseline_preference.py
"""
import importlib.util
import sys
import time
from pathlib import Path

HERE = Path(__file__).resolve().parent
BASE = HERE.parent.parent / "趋势分析工具 -最终app版本"
OUT_DIR = HERE.parent / "baseline" / "python-pref"
OUT_DIR.mkdir(parents=True, exist_ok=True)

spec = importlib.util.spec_from_file_location("pref_pipeline", BASE / "backend" / "preference_pipeline.py")
mod = importlib.util.module_from_spec(spec)
sys.modules["pref_pipeline"] = mod
spec.loader.exec_module(mod)


def log(text, kind="normal", step=1):
    print(f"[{step}][{kind}] {text}")


t0 = time.time()
result = mod.run_pipeline(BASE / "各商品客户拿货历史_5.12.xlsx", OUT_DIR, log)
print(f"\nbaseline done in {time.time() - t0:.1f}s")
print({k: str(v) for k, v in result.items()})
