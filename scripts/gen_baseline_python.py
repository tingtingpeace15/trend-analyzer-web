# -*- coding: utf-8 -*-
"""用真实 Python 版 pipeline 生成基准输出 xlsx(M7 全字段 diff 用)。

只读基准项目的代码与测试数据;输出写到本项目 baseline/python/(已 gitignore)。

用法: python3 scripts/gen_baseline_python.py
输出: baseline/python/商品销售趋势.xlsx
"""
import importlib.util
import sys
import time
from pathlib import Path

HERE = Path(__file__).resolve().parent
BASE = HERE.parent.parent / "趋势分析工具 -最终app版本"
OUT_DIR = HERE.parent / "baseline" / "python"
OUT_DIR.mkdir(parents=True, exist_ok=True)

# 按文件路径加载 pipeline.py(它自己会把 ../lib 加进 sys.path)
spec = importlib.util.spec_from_file_location("baseline_pipeline", BASE / "backend" / "pipeline.py")
mod = importlib.util.module_from_spec(spec)
sys.modules["baseline_pipeline"] = mod
spec.loader.exec_module(mod)


def log(text, kind="normal", step=1):
    print(f"[{step}][{kind}] {text}")


t0 = time.time()
result = mod.run_pipeline(
    zhixiao_path=BASE / "滞销商品【按销售】_njfayxbugwDqARB.xlsx",
    sales_path=BASE / "各商品客户拿货历史_5.12.xlsx",
    job_dir=OUT_DIR,
    log=log,
)
print(f"\nbaseline done in {time.time() - t0:.1f}s")
print(result)
