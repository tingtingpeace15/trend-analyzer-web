# -*- coding: utf-8 -*-
"""生成画图阶段的黄金基准:参考 PNG + 几何参数 JSON。

draw()/draw_detail() 代码逐行照搬 pipeline.py:289-389。
对每个样本图额外导出几何参数(最终 PNG 像素尺寸、ylim、数据坐标→PNG 像素
的仿射、axes 像素框、刻度位置),JS 版 Canvas 按这些参数复现。

输出:
  src/__tests__/golden/charts/*.png        参考图(人工目检 + M7 对比用)
  src/__tests__/golden/charts.golden.json  几何参数
"""
import json
import struct
from pathlib import Path

import numpy as np
import pandas as pd
import matplotlib
matplotlib.use("Agg")
import matplotlib.pyplot as plt

BASE = Path(__file__).resolve().parent.parent.parent / "趋势分析工具 -最终app版本"
OUT_DIR = Path(__file__).resolve().parent.parent / "src" / "__tests__" / "golden" / "charts"
OUT_JSON = Path(__file__).resolve().parent.parent / "src" / "__tests__" / "golden" / "charts.golden.json"
OUT_DIR.mkdir(parents=True, exist_ok=True)

ZHIXIAO = BASE / "滞销商品【按销售】_njfayxbugwDqARB.xlsx"
SALES = BASE / "各商品客户拿货历史_5.12.xlsx"

# ── 重建 kuan_pivot(同 gen_golden_aggregator.py)────────────────────────
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
df_h["日期"] = df_h["下单时间"].dt.date
end_date = df_h["日期"].max()
start_date = df_h["日期"].min()
date_range = pd.date_range(start_date, end_date, freq="D").date
kuan_daily = df_h.groupby(["货号_k", "日期"])["净销售量"].sum().reset_index()
kuan_pivot = kuan_daily.pivot_table(index="货号_k", columns="日期", values="净销售量", fill_value=0)
kuan_pivot = kuan_pivot.reindex(columns=date_range, fill_value=0)

# ── 画图函数(pipeline.py:289-389 原样)+ 几何参数导出 ───────────────────


def fmt_val(v):
    iv = int(v) if v == int(v) else round(v, 1)
    return str(iv)


def png_size(path):
    with open(path, "rb") as f:
        head = f.read(24)
    w, h = struct.unpack(">II", head[16:24])
    return int(w), int(h)


def export_geometry(fig, ax, savefig_dpi, pad_inches):
    """savefig(bbox_inches='tight') 的最终图内几何:像素尺寸 + data→px 仿射 + axes 框"""
    fig.canvas.draw()
    renderer = fig.canvas.get_renderer()
    tb = fig.get_tightbbox(renderer).padded(pad_inches)  # inches
    scale = savefig_dpi / fig.dpi
    w_px = tb.width * savefig_dpi
    h_px = tb.height * savefig_dpi
    # data → display(fig.dpi 像素) → savefig 像素 → 减 tight 裁剪原点;y 翻转(PNG 从上往下)
    ox = tb.x0 * savefig_dpi
    oy = tb.y0 * savefig_dpi

    def to_px(xd, yd):
        X, Y = ax.transData.transform((xd, yd))
        return [X * scale - ox, h_px - (Y * scale - oy)]

    axb = ax.bbox  # display units @ fig.dpi
    axes_px = {
        "x0": axb.x0 * scale - ox,
        "y0": h_px - (axb.y1 * scale - oy),  # top
        "w": axb.width * scale,
        "h": axb.height * scale,
    }
    return tb, scale, w_px, h_px, to_px, axes_px


def draw_with_meta(values, filename):
    """draw() 原样 + 几何导出"""
    values = np.array(values, dtype=float)
    fig, ax = plt.subplots(figsize=(3.6, 0.95), dpi=120)
    if values.max() == 0 and values.min() == 0:
        ax.plot(range(len(values)), values, color="#dddddd", linewidth=1.0, linestyle="--")
        ax.set_ylim(-1, 1)
    else:
        color = "#e74c3c"
        x = np.arange(len(values))
        ax.plot(x, values, color=color, linewidth=1.1)
        ax.fill_between(x, values, 0, color=color, alpha=0.10)
        ax.axhline(0, color="#bbbbbb", linewidth=0.4)
        nonzero_idx = np.where(values != 0)[0]
        for i in nonzero_idx:
            v = values[i]
            ax.plot(i, v, "o", color=color, markersize=3.0,
                    markerfacecolor=color, markeredgecolor="white", markeredgewidth=0.5)
        for rank_pos, i in enumerate(nonzero_idx):
            v = values[i]
            close_neighbour = rank_pos > 0 and (i - nonzero_idx[rank_pos - 1]) <= 1
            y_off = 4
            if close_neighbour:
                y_off += 3
            ax.annotate(fmt_val(v), xy=(i, v), xytext=(0, y_off),
                        textcoords="offset points", ha="center", va="bottom",
                        fontsize=6.5, color=color, weight="bold", clip_on=False)
        ymin, ymax = values.min(), values.max()
        yrange = ymax - ymin
        pad = max(yrange * 0.35, 1.0)
        ax.set_ylim(ymin - pad * 0.3, ymax + pad * 1.2)
    ax.set_xlim(-0.5, len(values) - 0.5)
    ax.axis("off")
    fig.patch.set_alpha(0)
    plt.tight_layout(pad=0.2)

    _, scale, w_px, h_px, to_px, axes_px = export_geometry(fig, ax, 110, 0.08)
    n = len(values)
    meta = {
        "ylim": list(ax.get_ylim()),
        "xlim": list(ax.get_xlim()),
        "calcSize": [w_px, h_px],
        "axesPx": axes_px,
        # 仿射验证点:数据坐标 → 最终 PNG 像素
        "probe": {
            "p0": {"data": [0, 0], "px": to_px(0, 0)},
            "pEnd": {"data": [n - 1, float(values.max())], "px": to_px(n - 1, float(values.max()))},
        },
    }
    plt.savefig(filename, dpi=110, bbox_inches="tight", pad_inches=0.08, transparent=True)
    plt.close(fig)
    meta["pngSize"] = list(png_size(filename))
    return meta


def draw_detail_with_meta(values, dates, filename):
    """draw_detail() 原样 + 几何导出"""
    values = np.array(values, dtype=float)
    fig, ax = plt.subplots(figsize=(6.4, 1.8), dpi=90)
    n = len(values)
    x = np.arange(n)
    if values.max() == 0 and values.min() == 0:
        ax.plot(x, values, color="#cccccc", linewidth=1.2, linestyle="--")
        ax.text(n / 2, 0.4, "no sales", ha="center", va="center", color="#999999", fontsize=10)
        ax.set_ylim(-1, 1)
    else:
        color = "#e74c3c"
        ax.plot(x, values, color=color, linewidth=1.6, marker="o", markersize=4.5,
                markerfacecolor=color, markeredgecolor="white", markeredgewidth=0.8)
        ax.fill_between(x, values, 0, color=color, alpha=0.12)
        ax.axhline(0, color="#888888", linewidth=0.5)
        nonzero_idx = np.where(values != 0)[0]
        for rank_pos, i in enumerate(nonzero_idx):
            v = values[i]
            close_neighbour = rank_pos > 0 and (i - nonzero_idx[rank_pos - 1]) <= 1
            y_off = 8
            if close_neighbour:
                y_off += 4
            ax.annotate(fmt_val(v), xy=(i, v), xytext=(0, y_off),
                        textcoords="offset points", ha="center", va="bottom",
                        fontsize=8.5, color=color, weight="bold", clip_on=False)
        ymin, ymax = values.min(), values.max()
        yrange = ymax - ymin
        pad = max(yrange * 0.30, 1.0)
        ax.set_ylim(ymin - pad * 0.2, ymax + pad * 1.3)

    tick_pos = list(range(0, n, max(1, n // 10)))
    if (n - 1) not in tick_pos:
        tick_pos.append(n - 1)
    ax.set_xticks(tick_pos)
    ax.set_xticklabels([dates[i].strftime("%m-%d") for i in tick_pos], fontsize=8, color="#444444")
    ax.set_xlim(-0.5, n - 0.5)

    for sp in ["top", "right"]:
        ax.spines[sp].set_visible(False)
    for sp in ["left", "bottom"]:
        ax.spines[sp].set_color("#bbbbbb")
        ax.spines[sp].set_linewidth(0.6)
    ax.tick_params(left=False, bottom=False, colors="#666666", labelsize=7.5)
    ax.grid(axis="y", alpha=0.15, linewidth=0.4)
    fig.patch.set_alpha(0)
    plt.tight_layout(pad=0.4)

    _, scale, w_px, h_px, to_px, axes_px = export_geometry(fig, ax, 110, 0.1)
    ylim = ax.get_ylim()
    yticks_all = list(ax.get_yticks())
    yticks_visible = [t for t in yticks_all if ylim[0] <= t <= ylim[1]]
    ytick_labels = [lbl.get_text() for lbl, t in zip(ax.get_yticklabels(), yticks_all)
                    if ylim[0] <= t <= ylim[1]]
    meta = {
        "ylim": list(ylim),
        "xlim": list(ax.get_xlim()),
        "calcSize": [w_px, h_px],
        "axesPx": axes_px,
        "tickPos": tick_pos,
        "tickLabels": [dates[i].strftime("%m-%d") for i in tick_pos],
        "yticks": yticks_visible,
        "ytickLabels": ytick_labels,
        "probe": {
            "p0": {"data": [0, 0], "px": to_px(0, 0)},
            "pEnd": {"data": [n - 1, float(values.max())], "px": to_px(n - 1, float(values.max()))},
        },
    }
    plt.savefig(filename, dpi=110, bbox_inches="tight", pad_inches=0.1, transparent=True)
    plt.close(fig)
    meta["pngSize"] = list(png_size(filename))
    return meta


# ── 样本:零销量 + 3 个有销量货号(疏/中/密)──────────────────────────────
agg_golden = json.loads(
    (Path(__file__).resolve().parent.parent / "src" / "__tests__" / "golden" / "aggregator.golden.json")
    .read_text(encoding="utf-8"))
sample_keys = [c["货号"] for c in agg_golden["pivotChecks"]]

samples = [{"id": "zero", "货号": None, "values": [0.0] * len(date_range)}]
for k in sample_keys:
    samples.append({"id": f"item_{k}", "货号": k, "values": kuan_pivot.loc[k].values.tolist()})

dates = list(date_range)
out = {"_comment": "由 scripts/gen_golden_charts.py 生成,勿手改。", "nDays": len(dates), "samples": []}
for s in samples:
    sm = draw_with_meta(s["values"], str(OUT_DIR / f"sm_{s['id']}.png"))
    dt = draw_detail_with_meta(s["values"], dates, str(OUT_DIR / f"dt_{s['id']}.png"))
    out["samples"].append({"id": s["id"], "货号": s["货号"], "small": sm, "detail": dt})
    print(f"{s['id']}: small={sm['pngSize']} calc={[round(v,2) for v in sm['calcSize']]} "
          f"ylim={[round(v,3) for v in sm['ylim']]} | detail={dt['pngSize']} "
          f"calc={[round(v,2) for v in dt['calcSize']]} yticks={dt['yticks']}")

OUT_JSON.write_text(json.dumps(out, ensure_ascii=False, indent=2), encoding="utf-8")
print(f"written: {OUT_JSON}")
