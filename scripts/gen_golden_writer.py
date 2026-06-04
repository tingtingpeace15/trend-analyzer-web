# -*- coding: utf-8 -*-
"""生成 writer 阶段的黄金基准:排序结果 + Sheet1 首行单元格字符串。

排序逻辑照搬 pipeline.py:483-486(sort_values 多列降序,NaN 在后,稳定);
as_str 的 pandas float 格式("387.0")也在这里固化成基准。

输出: src/__tests__/golden/writer.golden.json
"""
import json
import math
from pathlib import Path

import numpy as np
import pandas as pd

HERE = Path(__file__).resolve().parent
BASE = HERE.parent.parent / "趋势分析工具 -最终app版本"
OUT = HERE.parent / "src" / "__tests__" / "golden" / "writer.golden.json"

# ── 重建 base(同 gen_golden_aggregator.py,代码照搬 pipeline.py)─────────
exec((HERE / "gen_golden_aggregator.py").read_text(encoding="utf-8")
     .split("# ── 导出基准")[0]
     .replace('OUT = Path(__file__).resolve().parent.parent / "src"', '_OUT_IGNORED = Path("/tmp")'))

# 排序(pipeline.py:483-486)
sort_cols = ["销售量_期间"]
if "可售库存" in base.columns:
    sort_cols.append("可售库存")
base_sorted = base.sort_values(sort_cols, ascending=False).reset_index()


def as_str(v):
    if v is None:
        return ""
    try:
        if pd.isna(v):
            return ""
    except Exception:
        pass
    return str(v)


def as_num(v):
    try:
        if pd.isna(v):
            return None
    except Exception:
        pass
    if isinstance(v, (np.integer,)):
        return int(v)
    if isinstance(v, (np.floating,)):
        return float(v)
    return v


def sheet1_row(row):
    """Sheet1 一行的写入值(列 1..13+extra,不含图列),顺序同 pipeline.py:524-558"""
    cells = [
        as_str(row["货号"]),
        as_str(row.get("品类")),
        as_str(row.get("品牌")),
        as_str(row.get("季节")),
        as_str(row.get("设计师")),
        as_str(row.get("上市天数")),
        as_str(row.get("未成交天数")),
        as_num(row.get("销进率")),
        as_num(row.get("库存价值")),
        as_num(row.get("可售库存")),
    ]
    for fname in extra_fields:
        raw_val = row.get(fname)
        if isinstance(raw_val, (int, float, np.integer, np.floating)) and not (
            isinstance(raw_val, (float, np.floating)) and pd.isna(raw_val)
        ):
            cells.append(as_num(raw_val))
        else:
            cells.append(as_str(raw_val))
    cells.append(int(row.get("销售量_期间", 0)))
    cells.append(float(row.get("总销售金额", 0.0) or 0.0))
    cells.append(as_num(row.get("盈利金额")))
    return cells


golden = {
    "_comment": "由 scripts/gen_golden_writer.py 生成,勿手改。",
    "sortedKeysFirst10": base_sorted["货号_k"].head(10).tolist(),
    "sortedKeysLast5": base_sorted["货号_k"].tail(5).tolist(),
    # 排序后前 3 行 + 第 100/500 行的 Sheet1 单元格值(字符串格式含 pandas float 行为)
    "sheet1Rows": {str(i): sheet1_row(base_sorted.iloc[i]) for i in [0, 1, 2, 99, 499, len(base_sorted) - 1]},
    "headers": ["货号", "品类", "品牌", "季节", "设计师", "上市天数", "未成交天数",
                "销进率", "库存价值", "可售库存"] + extra_fields
               + ["销售量", "总销售金额", "盈利金额", "商品销售量趋势图"],
}


def clean(o):
    if isinstance(o, float) and math.isnan(o):
        return None
    if isinstance(o, dict):
        return {k: clean(v) for k, v in o.items()}
    if isinstance(o, list):
        return [clean(v) for v in o]
    return o


OUT.write_text(json.dumps(clean(golden), ensure_ascii=False, indent=2), encoding="utf-8")
print(f"written: {OUT}")
print("first10:", golden["sortedKeysFirst10"])
print("row0:", golden["sheet1Rows"]["0"])
