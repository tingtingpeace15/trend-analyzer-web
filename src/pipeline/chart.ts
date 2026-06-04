// chart.ts — Canvas 复刻 matplotlib 趋势图(pipeline.py:289-389 的 draw / draw_detail)。
//
// 设计:几何(buildSmallScene / buildDetailScene,纯函数)与渲染(renderScene /
// sceneToPng,需要 OffscreenCanvas)分离——几何在 Node 里可以对照 matplotlib
// 导出的黄金参数测试(charts.golden.json),渲染只在浏览器 Worker 里跑。
//
// 几何常量来自对 matplotlib 实际输出的反推校准(scripts/gen_golden_charts.py
// 对真实数据导出的 axes 像素框/仿射,4 个样本拟合,误差 <0.5px):
//   - 小图(figsize 3.6×0.95 @ savefig dpi=110, tight pad 0.08in):
//       画布 407×115,axes 恒为 (8.8, 8.8, 389.89, 98.39),仅当边缘标注溢出时右侧收缩
//   - 详情图(figsize 6.4×1.8 @ 110, pad 0.1in):
//       画布 713×206,左边距 = 21.75 + y轴刻度标签最大宽度,右边距 24.2,底 30.92
// Excel 里的显示尺寸由 writer 写死(340×78 / 640×180),PNG 原始像素 ±1px 无碍。
/** 1 matplotlib point → 像素(savefig dpi=110,1pt = 1/72in) */
const PT = 110 / 72;

// ── 场景模型 ─────────────────────────────────────────────────────────────

export type SceneOp =
  | { kind: 'polyline'; pts: number[]; color: string; lw: number; dash?: number[]; alpha?: number }
  | { kind: 'polygon'; pts: number[]; color: string; alpha: number }
  | { kind: 'marker'; x: number; y: number; r: number; fill: string; edge: string; edgeLw: number }
  | {
      kind: 'text'; x: number; y: number; text: string; sizePx: number; color: string;
      bold?: boolean; align: 'center' | 'left' | 'right'; baseline: 'bottom' | 'middle' | 'top';
    };

export interface AxesRect { x0: number; y0: number; w: number; h: number }

export interface ChartScene {
  width: number;
  height: number;
  ops: SceneOp[];
  /** 测试对照用的几何元数据 */
  meta: {
    ax: AxesRect;
    ylim: [number, number];
    xlim: [number, number];
    yticks?: number[];
    ytickLabels?: string[];
    tickPos?: number[];
  };
}

// ── 文本宽度估算(DejaVu Sans 字宽表,em 单位)──────────────────────────
// matplotlib 用 DejaVu Sans;布局需要标签宽度,但 Worker 里没有该字体,
// 用字宽表估算(数字 0.6362em 等),已对照 matplotlib 输出校准。

const DEJAVU: Record<string, number> = { '.': 0.3179, '-': 0.414, '−': 0.838 };
const DEJAVU_DIGIT = 0.6362;
const DEJAVU_BOLD_DIGIT = 0.696;

function textWidthPx(s: string, sizePt: number, bold = false): number {
  let em = 0;
  for (const ch of s) {
    if (ch >= '0' && ch <= '9') em += bold ? DEJAVU_BOLD_DIGIT : DEJAVU_DIGIT;
    else em += DEJAVU[ch] ?? 0.6;
  }
  return em * sizePt * PT;
}

/** np.round 同款 banker's rounding,保 1 位小数 */
function npRound1(x: number): number {
  const y = x * 10;
  const f = Math.floor(y);
  const diff = y - f;
  let r: number;
  if (diff > 0.5) r = f + 1;
  else if (diff < 0.5) r = f;
  else r = f % 2 === 0 ? f : f + 1;
  return r / 10;
}

/** matplotlib 数值标注文案:整数原样,否则保留 1 位(round-half-even,同 Python round) */
export function fmtVal(v: number): string {
  if (v === Math.trunc(v)) return String(Math.trunc(v));
  return String(npRound1(v));
}

// ── ylim 公式(逐行对应 draw / draw_detail)──────────────────────────────

export function smallYlim(values: ArrayLike<number>): [number, number] {
  let min = Infinity, max = -Infinity;
  for (let i = 0; i < values.length; i++) {
    if (values[i] < min) min = values[i];
    if (values[i] > max) max = values[i];
  }
  if (min === 0 && max === 0) return [-1, 1];
  const pad = Math.max((max - min) * 0.35, 1.0);
  return [min - pad * 0.3, max + pad * 1.2];
}

export function detailYlim(values: ArrayLike<number>): [number, number] {
  let min = Infinity, max = -Infinity;
  for (let i = 0; i < values.length; i++) {
    if (values[i] < min) min = values[i];
    if (values[i] > max) max = values[i];
  }
  if (min === 0 && max === 0) return [-1, 1];
  const pad = Math.max((max - min) * 0.30, 1.0);
  return [min - pad * 0.2, max + pad * 1.3];
}

// ── MaxNLocator(matplotlib 默认 y 轴刻度,steps [1,2,2.5,5,10])─────────
// nbins=7:由详情图 axes 高度 ≈164px(=107pt)÷ (labelsize 7.5 × 2) 得出,
// 与 4 个黄金样本的实际刻度完全吻合。

export function maxNTicks(vmin: number, vmax: number, nbins = 7): { ticks: number[]; step: number } {
  const range = vmax - vmin;
  const rawStep = range / nbins;
  const mag = Math.pow(10, Math.floor(Math.log10(rawStep)));
  const steps = [1, 2, 2.5, 5, 10];
  let step = 10 * mag;
  for (const s of steps) {
    if (s * mag >= rawStep - 1e-12) {
      step = s * mag;
      break;
    }
  }
  const low = Math.floor(vmin / step + 1e-9) * step;
  const ticks: number[] = [];
  for (let t = low; t <= vmax + step * 1e-9; t += step) {
    // 消浮点误差(0.30000000000000004 → 0.3)
    ticks.push(Number((Math.round(t / step) * step).toPrecision(12)));
  }
  return { ticks: ticks.filter((t) => t >= vmin - range * 1e-9 && t <= vmax + range * 1e-9), step };
}

/** ScalarFormatter 的简化:按步长决定小数位(0.5→"0.5",20→"20") */
export function formatTick(t: number, step: number): string {
  const decimals = step >= 1 && Number.isInteger(step) ? 0 : Math.max(0, -Math.floor(Math.log10(step)));
  return t.toFixed(decimals);
}

// ── 标注布局(两图共用,y_off 规则同 pipeline.py:312-322 / 353-363)──────

interface Annotation { i: number; text: string; yOffPt: number }

function buildAnnotations(values: ArrayLike<number>, baseOffPt: number, bumpPt: number): Annotation[] {
  const nonzero: number[] = [];
  for (let i = 0; i < values.length; i++) if (values[i] !== 0) nonzero.push(i);
  return nonzero.map((i, rank) => {
    const close = rank > 0 && i - nonzero[rank - 1] <= 1;
    return { i, text: fmtVal(values[i]), yOffPt: baseOffPt + (close ? bumpPt : 0) };
  });
}

// ── 小缩略图(draw,Sheet 1 用)──────────────────────────────────────────

const SM = {
  PAD: 8.8,          // tight pad 0.08in × 110
  AX_W0: 389.89,     // 无标注溢出时的 axes 宽
  AX_H: 98.39,
  COLOR: '#e74c3c',
};

export function buildSmallScene(values: ArrayLike<number>): ChartScene {
  const n = values.length;
  const ylim = smallYlim(values);
  const isZero = ylim[0] === -1 && ylim[1] === 1 && (() => {
    for (let i = 0; i < n; i++) if (values[i] !== 0) return false;
    return true;
  })();

  // 右侧标注溢出 → axes 收缩(tight_layout 行为)
  let axW = SM.AX_W0;
  const annotations = isZero ? [] : buildAnnotations(values, 4, 3);
  if (!isZero) {
    let overhang = 0;
    for (const a of annotations) {
      const xPx = SM.PAD + ((a.i + 0.5) / n) * SM.AX_W0;
      const half = textWidthPx(a.text, 6.5, true) / 2;
      overhang = Math.max(overhang, xPx + half - (SM.PAD + SM.AX_W0));
    }
    axW = SM.AX_W0 - overhang;
  }

  const ax: AxesRect = { x0: SM.PAD, y0: SM.PAD, w: axW, h: SM.AX_H };
  const width = Math.floor(SM.PAD * 2 + SM.AX_W0);
  const height = Math.floor(SM.PAD * 2 + SM.AX_H);
  const xPx = (x: number) => ax.x0 + ((x + 0.5) / n) * ax.w;
  const yPx = (y: number) => ax.y0 + ((ylim[1] - y) / (ylim[1] - ylim[0])) * ax.h;

  const ops: SceneOp[] = [];
  if (isZero) {
    // 水平虚线(matplotlib '--' 默认 dash pattern [3.7, 1.6] × linewidth,单位 pt)
    const lw = 1.0 * PT;
    ops.push({
      kind: 'polyline', pts: [xPx(0), yPx(0), xPx(n - 1), yPx(0)],
      color: '#dddddd', lw, dash: [3.7 * lw, 1.6 * lw],
    });
  } else {
    // fill_between(zorder 低于线)
    const poly: number[] = [];
    for (let i = 0; i < n; i++) poly.push(xPx(i), yPx(values[i]));
    poly.push(xPx(n - 1), yPx(0), xPx(0), yPx(0));
    ops.push({ kind: 'polygon', pts: poly, color: SM.COLOR, alpha: 0.10 });
    // 主线
    const line: number[] = [];
    for (let i = 0; i < n; i++) line.push(xPx(i), yPx(values[i]));
    ops.push({ kind: 'polyline', pts: line, color: SM.COLOR, lw: 1.1 * PT });
    // axhline y=0(横贯整个 xlim)
    ops.push({
      kind: 'polyline', pts: [ax.x0, yPx(0), ax.x0 + ax.w, yPx(0)],
      color: '#bbbbbb', lw: 0.4 * PT,
    });
    // 非零点 marker(markersize=直径 pt)
    for (let i = 0; i < n; i++) {
      if (values[i] !== 0) {
        ops.push({
          kind: 'marker', x: xPx(i), y: yPx(values[i]), r: (3.0 / 2) * PT,
          fill: SM.COLOR, edge: 'white', edgeLw: 0.5 * PT,
        });
      }
    }
    // 数值标注
    for (const a of annotations) {
      ops.push({
        kind: 'text', x: xPx(a.i), y: yPx(values[a.i]) - a.yOffPt * PT,
        text: a.text, sizePx: 6.5 * PT, color: SM.COLOR, bold: true,
        align: 'center', baseline: 'bottom',
      });
    }
  }

  return { width, height, ops, meta: { ax, ylim, xlim: [-0.5, n - 0.5] } };
}

// ── 详情大图(draw_detail,Sheet 2 用)───────────────────────────────────

const DT = {
  W: 713,
  H: 206,
  PAD: 11,            // tight pad 0.1in × 110
  LEFT_BASE: 21.75,   // 左边距 = LEFT_BASE + y刻度标签最大宽度(校准值)
  RIGHT: 24.2,        // 右边距(末位日期标签溢出,实测恒定)
  BOTTOM: 30.92,      // 底边距(x 刻度标签区,实测恒定)
  YLABEL_HALF_H: 4.28, // y 标签半高(顶部刻度贴 ylim 上沿时把 axes 往下推)
  TICK_PAD: 5.35,     // 刻度标签到 axes 的距离(3.5pt)
  COLOR: '#e74c3c',
};

export function buildDetailScene(values: ArrayLike<number>, dates: string[]): ChartScene {
  const n = values.length;
  const ylim = detailYlim(values);
  let isZero = true;
  for (let i = 0; i < n; i++) if (values[i] !== 0) { isZero = false; break; }

  // y 刻度(零销量图 ylim 恒 (-1,1),也走同一个 locator)
  const { ticks: yticks, step } = maxNTicks(ylim[0], ylim[1]);
  const ytickLabels = yticks.map((t) => formatTick(t, step));
  const maxYLabelW = Math.max(...yticks.map((t, i) =>
    textWidthPx(t < 0 ? `−${ytickLabels[i].replace('-', '')}` : ytickLabels[i], 7.5)));

  // x 刻度位置(pipeline.py:369-371)
  const tickPos: number[] = [];
  const tickStep = Math.max(1, Math.floor(n / 10));
  for (let i = 0; i < n; i += tickStep) tickPos.push(i);
  if (!tickPos.includes(n - 1)) tickPos.push(n - 1);
  const tickLabels = tickPos.map((i) => dates[i].slice(5)); // "YYYY-MM-DD" → "MM-DD"

  // axes 框:顶部刻度贴 ylim 上沿时 y0 下移(标签半高溢出)
  const x0 = DT.LEFT_BASE + maxYLabelW;
  let y0 = DT.PAD;
  let h = DT.H - y0 - DT.BOTTOM;
  const topTick = yticks[yticks.length - 1];
  if (topTick !== undefined) {
    const dPx = ((ylim[1] - topTick) / (ylim[1] - ylim[0])) * h;
    y0 = DT.PAD + Math.max(0, DT.YLABEL_HALF_H - dPx);
    h = DT.H - y0 - DT.BOTTOM;
  }
  const ax: AxesRect = { x0, y0, w: DT.W - x0 - DT.RIGHT, h };
  const xPx = (x: number) => ax.x0 + ((x + 0.5) / n) * ax.w;
  const yPx = (y: number) => ax.y0 + ((ylim[1] - y) / (ylim[1] - ylim[0])) * ax.h;

  const ops: SceneOp[] = [];
  // 网格(axisbelow,最先画):y 刻度横线,#b0b0b0 alpha 0.15
  for (const t of yticks) {
    ops.push({
      kind: 'polyline', pts: [ax.x0, yPx(t), ax.x0 + ax.w, yPx(t)],
      color: '#b0b0b0', lw: 0.4 * PT, alpha: 0.15,
    });
  }
  // 左/下 spine
  ops.push({ kind: 'polyline', pts: [ax.x0, ax.y0, ax.x0, ax.y0 + ax.h], color: '#bbbbbb', lw: 0.6 * PT });
  ops.push({
    kind: 'polyline', pts: [ax.x0, ax.y0 + ax.h, ax.x0 + ax.w, ax.y0 + ax.h],
    color: '#bbbbbb', lw: 0.6 * PT,
  });

  if (isZero) {
    const lw = 1.2 * PT;
    ops.push({
      kind: 'polyline', pts: [xPx(0), yPx(0), xPx(n - 1), yPx(0)],
      color: '#cccccc', lw, dash: [3.7 * lw, 1.6 * lw],
    });
    ops.push({
      kind: 'text', x: xPx(n / 2), y: yPx(0.4), text: 'no sales',
      sizePx: 10 * PT, color: '#999999', align: 'center', baseline: 'middle',
    });
  } else {
    const poly: number[] = [];
    for (let i = 0; i < n; i++) poly.push(xPx(i), yPx(values[i]));
    poly.push(xPx(n - 1), yPx(0), xPx(0), yPx(0));
    ops.push({ kind: 'polygon', pts: poly, color: DT.COLOR, alpha: 0.12 });
    const line: number[] = [];
    for (let i = 0; i < n; i++) line.push(xPx(i), yPx(values[i]));
    ops.push({ kind: 'polyline', pts: line, color: DT.COLOR, lw: 1.6 * PT });
    ops.push({
      kind: 'polyline', pts: [ax.x0, yPx(0), ax.x0 + ax.w, yPx(0)],
      color: '#888888', lw: 0.5 * PT,
    });
    // marker='o' 在线上 → 每个点都有(含 0 值点)
    for (let i = 0; i < n; i++) {
      ops.push({
        kind: 'marker', x: xPx(i), y: yPx(values[i]), r: (4.5 / 2) * PT,
        fill: DT.COLOR, edge: 'white', edgeLw: 0.8 * PT,
      });
    }
    for (const a of buildAnnotations(values, 8, 4)) {
      ops.push({
        kind: 'text', x: xPx(a.i), y: yPx(values[a.i]) - a.yOffPt * PT,
        text: a.text, sizePx: 8.5 * PT, color: DT.COLOR, bold: true,
        align: 'center', baseline: 'bottom',
      });
    }
  }

  // y 刻度标签(右对齐,#666666 7.5pt)
  yticks.forEach((t, i) => {
    ops.push({
      kind: 'text', x: ax.x0 - DT.TICK_PAD, y: yPx(t), text: ytickLabels[i],
      sizePx: 7.5 * PT, color: '#666666', align: 'right', baseline: 'middle',
    });
  });
  // x 刻度标签(#444444 8pt)
  tickPos.forEach((p, i) => {
    ops.push({
      kind: 'text', x: xPx(p), y: ax.y0 + ax.h + DT.TICK_PAD, text: tickLabels[i],
      sizePx: 8 * PT, color: '#444444', align: 'center', baseline: 'top',
    });
  });

  return {
    width: DT.W, height: DT.H, ops,
    meta: { ax, ylim, xlim: [-0.5, n - 0.5], yticks, ytickLabels, tickPos },
  };
}

// ── 渲染(浏览器 / Worker;透明背景,同 matplotlib transparent=True)─────

export function renderScene(scene: ChartScene, ctx: OffscreenCanvasRenderingContext2D): void {
  ctx.clearRect(0, 0, scene.width, scene.height);
  ctx.lineJoin = 'round';
  ctx.lineCap = 'butt';
  for (const op of scene.ops) {
    ctx.globalAlpha = 1;
    switch (op.kind) {
      case 'polygon': {
        ctx.globalAlpha = op.alpha;
        ctx.fillStyle = op.color;
        ctx.beginPath();
        ctx.moveTo(op.pts[0], op.pts[1]);
        for (let i = 2; i < op.pts.length; i += 2) ctx.lineTo(op.pts[i], op.pts[i + 1]);
        ctx.closePath();
        ctx.fill();
        break;
      }
      case 'polyline': {
        if (op.alpha != null) ctx.globalAlpha = op.alpha;
        ctx.strokeStyle = op.color;
        ctx.lineWidth = op.lw;
        ctx.setLineDash(op.dash ?? []);
        ctx.beginPath();
        ctx.moveTo(op.pts[0], op.pts[1]);
        for (let i = 2; i < op.pts.length; i += 2) ctx.lineTo(op.pts[i], op.pts[i + 1]);
        ctx.stroke();
        ctx.setLineDash([]);
        break;
      }
      case 'marker': {
        ctx.beginPath();
        ctx.arc(op.x, op.y, op.r, 0, Math.PI * 2);
        ctx.fillStyle = op.fill;
        ctx.fill();
        ctx.lineWidth = op.edgeLw;
        ctx.strokeStyle = op.edge;
        ctx.stroke();
        break;
      }
      case 'text': {
        ctx.fillStyle = op.color;
        ctx.font = `${op.bold ? 'bold ' : ''}${op.sizePx}px "DejaVu Sans", "Helvetica Neue", Arial, sans-serif`;
        ctx.textAlign = op.align;
        ctx.textBaseline = op.baseline === 'bottom' ? 'alphabetic' : op.baseline;
        ctx.fillText(op.text, op.x, op.y);
        break;
      }
    }
  }
  ctx.globalAlpha = 1;
}

/** 场景 → PNG 字节(透明背景)。复用同一个 OffscreenCanvas 减少分配 */
export async function sceneToPng(scene: ChartScene, canvas?: OffscreenCanvas): Promise<Uint8Array> {
  const c = canvas && canvas.width === scene.width && canvas.height === scene.height
    ? canvas
    : new OffscreenCanvas(scene.width, scene.height);
  const ctx = c.getContext('2d');
  if (!ctx) throw new Error('OffscreenCanvas 2d context 不可用');
  renderScene(scene, ctx);
  const blob = await c.convertToBlob({ type: 'image/png' });
  return new Uint8Array(await blob.arrayBuffer());
}
