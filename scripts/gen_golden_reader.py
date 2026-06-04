# -*- coding: utf-8 -*-
"""生成 reader 阶段的黄金基准数(逻辑逐行照搬 pipeline.py 步骤 2 的读取部分)。

用真实测试数据跑 pandas 版读取,把行列数/列名/聚合校验值写成 JSON,
供 Vitest 对比 JS 版 reader 的输出。只读基准目录,绝不写入。

用法: python3 scripts/gen_golden_reader.py
输出: src/__tests__/golden/reader.golden.json
"""
import json
from pathlib import Path

import pandas as pd

BASE = Path(__file__).resolve().parent.parent.parent / "趋势分析工具 -最终app版本"
OUT = Path(__file__).resolve().parent.parent / "src" / "__tests__" / "golden" / "reader.golden.json"

ZHIXIAO = BASE / "滞销商品【按销售】_njfayxbugwDqARB.xlsx"
SALES = BASE / "各商品客户拿货历史_5.12.xlsx"


def read_zhixiao(path):
    df_zi = pd.read_excel(path, header=1)
    assert "货号" in df_zi.columns
    df_zi["货号_k"] = df_zi["货号"].astype(str).str.strip()
    df_zi = df_zi.drop_duplicates(subset=["货号_k"], keep="first")
    return df_zi


def read_sales(path):
    xls = pd.ExcelFile(path)
    frames_h = []
    sheet_meta = []
    for sh in xls.sheet_names:
        raw = pd.read_excel(xls, sheet_name=sh, header=None, nrows=3)
        hrow = None
        for i in range(min(3, len(raw))):
            joined = " ".join(str(v) for v in raw.iloc[i])
            if any(k in joined for k in ["货号", "净销售", "下单时间", "客户名称"]):
                hrow = i
                break
        d = pd.read_excel(xls, sheet_name=sh, header=hrow if hrow is not None else 0)
        sheet_meta.append({"name": sh, "headerRow": hrow, "cols": int(d.shape[1]),
                           "rows": int(d.shape[0]), "kept": len(d.columns) >= 10})
        if len(d.columns) >= 10:
            frames_h.append(d)
    base_cols = list(frames_h[0].columns)
    for i in range(1, len(frames_h)):
        if len(frames_h[i].columns) >= len(base_cols):
            frames_h[i] = frames_h[i].iloc[:, : len(base_cols)]
            frames_h[i].columns = base_cols
    df_h = pd.concat(frames_h, ignore_index=True)
    return df_h, xls.sheet_names, sheet_meta


df_zi = read_zhixiao(ZHIXIAO)
df_h, sheet_names, sheet_meta = read_sales(SALES)

# 清洗(照搬 pipeline.py 167-172 行)
df_h["下单时间"] = pd.to_datetime(df_h["下单时间"], errors="coerce")
df_h = df_h.dropna(subset=["下单时间", "货号"])
df_h["货号_k"] = df_h["货号"].astype(str).str.strip()
df_h["净销售量"] = pd.to_numeric(df_h["净销售量"], errors="coerce").fillna(0)
df_h["净销售金额"] = pd.to_numeric(df_h["净销售金额"], errors="coerce").fillna(0)
df_h["日期"] = df_h["下单时间"].dt.date

golden = {
    "_comment": "由 scripts/gen_golden_reader.py 生成,勿手改。对照 pipeline.py 读取阶段。",
    "zhixiao": {
        "file": ZHIXIAO.name,
        "rows": int(len(df_zi)),
        # 列数不含内部加的 货号_k
        "cols": int(df_zi.shape[1]) - 1,
        "columns": [c for c in df_zi.columns if c != "货号_k"],
        "first5货号": df_zi["货号_k"].head(5).tolist(),
    },
    "sales": {
        "file": SALES.name,
        "sheetNames": sheet_names,
        "sheetMeta": sheet_meta,
        "rawRows": int(sum(m["rows"] for m in sheet_meta if m["kept"])),
        "cols": int(df_h.shape[1]) - 2,  # 不含内部加的 货号_k / 日期
        "columns": [c for c in df_h.columns if c not in ("货号_k", "日期")],
        # 清洗后(dropna 下单时间/货号)
        "cleanRows": int(len(df_h)),
        "minDate": str(df_h["日期"].min()),
        "maxDate": str(df_h["日期"].max()),
        "totalQty": float(df_h["净销售量"].sum()),
        "totalAmt": round(float(df_h["净销售金额"].sum()), 2),
        "uniqueItems": int(df_h["货号_k"].nunique()),
    },
}

OUT.parent.mkdir(parents=True, exist_ok=True)
OUT.write_text(json.dumps(golden, ensure_ascii=False, indent=2), encoding="utf-8")
print(f"written: {OUT}")
print(json.dumps(golden["zhixiao"] | {"columns": "..."}, ensure_ascii=False))
print(json.dumps({k: v for k, v in golden["sales"].items() if k not in ("columns", "sheetMeta")}, ensure_ascii=False))
