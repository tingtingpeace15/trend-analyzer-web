# -*- coding: utf-8 -*-
"""生成聚合阶段的黄金基准数(代码逐行照搬 pipeline.py:113-282)。

输出: src/__tests__/golden/aggregator.golden.json
用法: python3 scripts/gen_golden_aggregator.py
"""
import json
import math
from pathlib import Path

import numpy as np
import pandas as pd

BASE = Path(__file__).resolve().parent.parent.parent / "趋势分析工具 -最终app版本"
OUT = Path(__file__).resolve().parent.parent / "src" / "__tests__" / "golden" / "aggregator.golden.json"

ZHIXIAO = BASE / "滞销商品【按销售】_njfayxbugwDqARB.xlsx"
SALES = BASE / "各商品客户拿货历史_5.12.xlsx"

# ── 读取(同 gen_golden_reader.py / pipeline.py:116-172)─────────────────
df_zi = pd.read_excel(ZHIXIAO, header=1)
df_zi["货号_k"] = df_zi["货号"].astype(str).str.strip()
df_zi = df_zi.drop_duplicates(subset=["货号_k"], keep="first")

xls = pd.ExcelFile(SALES)
frames_h = []
for sh in xls.sheet_names:
    raw = pd.read_excel(xls, sheet_name=sh, header=None, nrows=3)
    hrow = None
    for i in range(min(3, len(raw))):
        joined = " ".join(str(v) for v in raw.iloc[i])
        if any(k in joined for k in ["货号", "净销售", "下单时间", "客户名称"]):
            hrow = i
            break
    d = pd.read_excel(xls, sheet_name=sh, header=hrow if hrow is not None else 0)
    if len(d.columns) >= 10:
        frames_h.append(d)
base_cols = list(frames_h[0].columns)
for i in range(1, len(frames_h)):
    if len(frames_h[i].columns) >= len(base_cols):
        frames_h[i] = frames_h[i].iloc[:, : len(base_cols)]
        frames_h[i].columns = base_cols
df_h = pd.concat(frames_h, ignore_index=True)

df_h["下单时间"] = pd.to_datetime(df_h["下单时间"], errors="coerce")
df_h = df_h.dropna(subset=["下单时间", "货号"])
df_h["货号_k"] = df_h["货号"].astype(str).str.strip()
df_h["净销售量"] = pd.to_numeric(df_h["净销售量"], errors="coerce").fillna(0)
df_h["净销售金额"] = pd.to_numeric(df_h["净销售金额"], errors="coerce").fillna(0)
df_h["日期"] = df_h["下单时间"].dt.date

# ── 聚合(pipeline.py:174-282 原样)──────────────────────────────────────
end_date = df_h["日期"].max()
start_date = df_h["日期"].min()
n_days = (end_date - start_date).days + 1
date_range = pd.date_range(start_date, end_date, freq="D").date

kuan_daily = df_h.groupby(["货号_k", "日期"])["净销售量"].sum().reset_index()
kuan_pivot = kuan_daily.pivot_table(index="货号_k", columns="日期", values="净销售量", fill_value=0)
kuan_pivot = kuan_pivot.reindex(columns=date_range, fill_value=0)

sales_qty = df_h.groupby("货号_k")["净销售量"].sum().rename("销售量_期间")
sales_amt = df_h.groupby("货号_k")["净销售金额"].sum().rename("总销售金额")

KNOWN_ZHIXIAO_FIELDS = {
    "货号", "货号_k",
    "品类", "品牌", "设计师品牌", "年份",
    "上市天数", "未成交天数:", "未成交天数",
    "销进率", "库存价值", "可售库存", "盈利金额",
}
SALES_AGG_NAMES = {"销售量", "总销售金额"}

zi_keep = ["货号_k"]
zi_rename = {}
for src, dst in [
    ("品类", "品类"),
    ("品牌", "品牌"),
    ("设计师品牌", "设计师"),
    ("年份", "季节"),
    ("上市天数", "上市天数"),
    ("未成交天数:", "未成交天数"),
    ("未成交天数", "未成交天数"),
    ("销进率", "销进率"),
    ("库存价值", "库存价值"),
    ("可售库存", "可售库存"),
    ("盈利金额", "盈利金额"),
]:
    if src in df_zi.columns and src not in zi_keep:
        zi_keep.append(src)
        if src != dst:
            zi_rename[src] = dst

extra_fields = []
extra_field_sources = {}
for c in df_zi.columns:
    if c in KNOWN_ZHIXIAO_FIELDS:
        continue
    if c in zi_keep:
        continue
    out_name = f"{c}_滞销表" if c in SALES_AGG_NAMES else c
    extra_fields.append(out_name)
    extra_field_sources[out_name] = c
    zi_keep.append(c)
    if out_name != c:
        zi_rename[c] = out_name

zi_subset = df_zi[zi_keep].rename(columns=zi_rename)

base = pd.DataFrame({"货号_k": df_zi["货号_k"].tolist()})
base = base.merge(zi_subset, on="货号_k", how="left")
base = base.merge(sales_qty, on="货号_k", how="left")
base = base.merge(sales_amt, on="货号_k", how="left")
base["销售量_期间"] = base["销售量_期间"].fillna(0).astype(int)
base["总销售金额"] = base["总销售金额"].fillna(0).round(2)
if "盈利金额" in base.columns:
    base["盈利金额"] = pd.to_numeric(base["盈利金额"], errors="coerce").fillna(0).round(2)
else:
    base["盈利金额"] = np.nan
if "销进率" in base.columns:
    base["销进率"] = (
        pd.to_numeric(base["销进率"].astype(str).str.rstrip("%").str.strip(), errors="coerce")
        .fillna(0) / 100.0
    )
else:
    base["销进率"] = np.nan
base["货号"] = base["货号_k"]

items_total = len(base)
items_with_sales = int((base["销售量_期间"] > 0).sum())

# ── 导出基准 ─────────────────────────────────────────────────────────────


def jval(v):
    """pandas 值 → JSON(NaN→null,numpy 标量→python)"""
    if v is None:
        return None
    if isinstance(v, float) and math.isnan(v):
        return None
    if pd.isna(v):
        return None
    if isinstance(v, (np.integer,)):
        return int(v)
    if isinstance(v, (np.floating,)):
        return float(v)
    return v


CONTRACT_FIELDS = ["品类", "品牌", "季节", "设计师", "上市天数", "未成交天数",
                   "销进率", "库存价值", "可售库存", "盈利金额"]


def item_at(i):
    row = base.iloc[i]
    rec = {"货号": row["货号"]}
    for f in CONTRACT_FIELDS:
        rec[f] = jval(row[f]) if f in base.columns else None
    rec["销售量"] = int(row["销售量_期间"])
    rec["总销售金额"] = jval(row["总销售金额"])
    rec["extra"] = {f: jval(row[f]) for f in extra_fields}
    return rec


# 抽查行:前 5 + 中段 + 末尾,覆盖有销量/零销量
sample_idx = [0, 1, 2, 3, 4, 100, 400, items_total - 1]

# 透视抽查:取 3 个有销量货号(且在滞销表里),记录整行序列特征
keys_with_sales = base.loc[base["销售量_期间"] > 0, "货号_k"].tolist()
pivot_checks = []
for k in [keys_with_sales[0], keys_with_sales[len(keys_with_sales) // 2], keys_with_sales[-1]]:
    series = kuan_pivot.loc[k]
    nz = series[series != 0]
    pivot_checks.append({
        "货号": k,
        "rowSum": float(series.sum()),
        "nonZeroDays": int((series != 0).sum()),
        "firstNonZero": {"date": str(nz.index[0]), "value": float(nz.iloc[0])},
        "lastNonZero": {"date": str(nz.index[-1]), "value": float(nz.iloc[-1])},
    })

golden = {
    "_comment": "由 scripts/gen_golden_aggregator.py 生成,勿手改。对照 pipeline.py 聚合阶段。",
    "startDate": str(start_date),
    "endDate": str(end_date),
    "windowDays": int(n_days),
    "itemsTotal": items_total,
    "itemsWithSales": items_with_sales,
    "extraFields": extra_fields,
    "pivotItemCount": int(kuan_pivot.shape[0]),
    "sumQty": int(base["销售量_期间"].sum()),
    "sumAmt": round(float(base["总销售金额"].sum()), 2),
    "sum盈利金额": round(float(base["盈利金额"].sum()), 2) if base["盈利金额"].notna().any() else None,
    "sum销进率": round(float(base["销进率"].sum()), 6) if base["销进率"].notna().any() else None,
    "sampleIndexes": sample_idx,
    "sampleItems": [item_at(i) for i in sample_idx],
    "pivotChecks": pivot_checks,
}

OUT.write_text(json.dumps(golden, ensure_ascii=False, indent=2), encoding="utf-8")
print(f"written: {OUT}")
print(json.dumps({k: golden[k] for k in ["itemsTotal", "itemsWithSales", "extraFields",
                                          "pivotItemCount", "sumQty", "sumAmt",
                                          "sum盈利金额", "sum销进率"]}, ensure_ascii=False))
