# -*- coding: utf-8 -*-
"""偏好分析读取阶段黄金基准(代码照搬 preference_pipeline.py:69-168)。

输出: src/__tests__/golden/preference_reader.golden.json
用法: python3 scripts/gen_golden_preference_reader.py
"""
import json
from pathlib import Path

import pandas as pd

BASE = Path(__file__).resolve().parent.parent.parent / "趋势分析工具 -最终app版本"
OUT = Path(__file__).resolve().parent.parent / "src" / "__tests__" / "golden" / "preference_reader.golden.json"
INPUT = BASE / "各商品客户拿货历史_5.12.xlsx"

# ── _load(preference_pipeline.py:69-112)─────────────────────────────────
xls = pd.ExcelFile(INPUT)
frames = []
sheet_meta = []
for sh in xls.sheet_names:
    raw = pd.read_excel(xls, sheet_name=sh, header=None, nrows=3)
    hrow = None
    for i in range(min(3, len(raw))):
        joined = " ".join(str(v) for v in raw.iloc[i])
        if any(k in joined for k in ["客户名称", "销售", "下单时间", "货号"]):
            hrow = i
            break
    df = pd.read_excel(xls, sheet_name=sh, header=hrow if hrow is not None else 0)
    kept = len(df.columns) >= 10
    sheet_meta.append({"name": sh, "headerRow": hrow, "cols": int(df.shape[1]),
                       "rows": int(df.shape[0]), "kept": kept})
    if kept:
        frames.append(df)

base = list(frames[0].columns)
for i in range(1, len(frames)):
    if len(frames[i].columns) == len(base):
        frames[i].columns = base
    elif len(frames[i].columns) > len(base):
        frames[i] = frames[i].iloc[:, : len(base)]
        frames[i].columns = base
df = pd.concat(frames, ignore_index=True)
raw_rows = int(len(df))

# ── _normalize_columns(preference_pipeline.py:115-168)───────────────────
remap = {}
for c in df.columns:
    s = str(c).strip().rstrip(":：")
    if "店铺" in s:
        remap[c] = "店铺"
    elif "货号" in s:
        remap[c] = "货号"
    elif s in ("分类", "品类", "类别"):
        remap[c] = "分类"
    elif "年份" in s:
        remap[c] = "年份"
    elif "设计师" in s:
        remap[c] = "设计师品牌"
    elif s == "品牌":
        remap[c] = "品牌"
    elif "颜色" in s:
        remap[c] = "颜色"
    elif "尺码" in s or "码数" in s:
        remap[c] = "尺码"
    elif "下单时间" in s or "日期" in s:
        remap[c] = "下单时间"
    elif s in ("客户名称", "客户"):
        remap[c] = "客户名称"
    elif s in ("销售", "销售员", "业务员"):
        remap[c] = "销售"
    elif "净销售金额" in s:
        remap[c] = "销售金额"
    elif "净销售量" in s:
        remap[c] = "销售量"
target_names = set(remap.values())
conflicts = [c for c in df.columns
             if c not in remap and str(c).strip().rstrip(":：") in target_names]
if conflicts:
    df = df.drop(columns=conflicts)
df.rename(columns=remap, inplace=True)

for c in ["销售量", "销售金额"]:
    if c in df.columns:
        df[c] = pd.to_numeric(df[c], errors="coerce")
if "下单时间" in df.columns:
    df["下单时间"] = pd.to_datetime(df["下单时间"], errors="coerce")

# ── 导出 ─────────────────────────────────────────────────────────────────
golden = {
    "_comment": "由 scripts/gen_golden_preference_reader.py 生成,勿手改。",
    "file": INPUT.name,
    "sheetNames": xls.sheet_names,
    "sheetMeta": sheet_meta,
    "rawRows": raw_rows,
    "droppedConflicts": conflicts,
    "columns": list(df.columns),
    "nonNullCounts": {c: int(df[c].notna().sum()) for c in df.columns},
    "sumQty": float(df["销售量"].sum()),
    "sumAmt": round(float(df["销售金额"].sum()), 2),
    "uniqueCustomers": int(df["客户名称"].nunique()),
    "dateMin": str(df["下单时间"].min().date()),
    "dateMax": str(df["下单时间"].max().date()),
    # 抽查首行的关键字段(规范化后)
    "row0": {c: (None if pd.isna(df.iloc[0][c]) else
                 (str(df.iloc[0][c]) if c == "下单时间" else
                  (float(df.iloc[0][c]) if isinstance(df.iloc[0][c], float) else
                   (int(df.iloc[0][c]) if hasattr(df.iloc[0][c], "item") and isinstance(df.iloc[0][c].item(), int) else str(df.iloc[0][c])))))
             for c in df.columns},
}
OUT.write_text(json.dumps(golden, ensure_ascii=False, indent=2), encoding="utf-8")
print(f"written: {OUT}")
print(json.dumps({k: golden[k] for k in ["rawRows", "droppedConflicts", "columns",
                                          "sumQty", "sumAmt", "uniqueCustomers",
                                          "dateMin", "dateMax"]}, ensure_ascii=False))
