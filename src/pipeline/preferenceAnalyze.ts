// preferenceAnalyze.ts — 偏好分析聚合,逐行照搬 preference_pipeline.py 的
// _season_type(:172-180) 与 _analyze(:183-431)。
//
// pandas 语义雷区(全部如实复刻):
//   - groupby 默认丢 NaN 键、键排序(字符串按码点);sort_values 平局按键升序兜底
//   - round() 是 banker's rounding(四舍六入五留双),int() 是向零截断
//   - quantile 用线性插值;mean/median 跳过 NaN
//   - pd.cut 右闭区间 (0,50],(50,100]…;observed=True 按分类序输出
//   - to_period('M') 对 NaT 产生字符串 "NaT" 组
import { PipelineError } from './errors';
import { msToDateStr } from './reader';
import { pyFloat } from './pyjson';
import type { PyFloat } from './pyjson';
import type { PreferenceData } from './preferenceReader';
import type { Cell } from '../types/excel';

export const TOP_N = 30;

// ── Python 数值语义 ───────────────────────────────────────────────────────

/**
 * Python 内建 round(x, d):对 double 的精确十进制值做 half-even 舍入(BigInt 精确计算)。
 * 与 np.round / pandas.round(先乘 10^d,有浮点乘法误差)在边界值上不同:
 * round(88.35, 1) → Python 88.3,乘法路径 88.4。两种语义都要,按调用处选。
 */
export function pyRound(x: number, ndigits = 0): number {
  if (!Number.isFinite(x) || x === 0) return x;
  const neg = x < 0;
  const dv = new DataView(new ArrayBuffer(8));
  dv.setFloat64(0, Math.abs(x));
  const hi = dv.getUint32(0);
  const lo = dv.getUint32(4);
  const expBits = (hi >>> 20) & 0x7ff;
  let m = (BigInt(hi & 0xfffff) << 32n) | BigInt(lo);
  let e2: number;
  if (expBits === 0) {
    e2 = -1074;
  } else {
    m |= 1n << 52n;
    e2 = expBits - 1075;
  }
  // |x| = m × 2^e2;q ≈ |x|×10^d 的 half-even 整数
  const p10 = 10n ** BigInt(ndigits);
  let q: bigint;
  if (e2 >= 0) {
    q = (m * p10) << BigInt(e2);
  } else {
    const D = 1n << BigInt(-e2);
    const T = m * p10;
    q = T / D;
    const r = T % D;
    if (r !== 0n) {
      const twice = r * 2n;
      if (twice > D || (twice === D && (q & 1n) === 1n)) q += 1n;
    }
  }
  const res = Number(q) / Number(p10);
  return neg ? -res : res;
}

/** np.round / pandas .round() 语义:乘 10^d 后对乘积 half-even(保留乘法误差) */
export function npRound(x: number, decimals = 0): number {
  const scale = Math.pow(10, decimals);
  const y = x * scale;
  const f = Math.floor(y);
  const diff = y - f;
  let r: number;
  if (diff > 0.5) r = f + 1;
  else if (diff < 0.5) r = f;
  else r = f % 2 === 0 ? f : f + 1;
  return r / scale;
}

/** Python int(x):向零截断 */
export const pyInt = (x: number) => Math.trunc(x);

/** pandas quantile(线性插值),输入须升序 */
export function quantile(sortedAsc: number[], q: number): number {
  if (sortedAsc.length === 0) return NaN;
  const pos = (sortedAsc.length - 1) * q;
  const lo = Math.floor(pos);
  const hi = Math.ceil(pos);
  if (lo === hi) return sortedAsc[lo];
  return sortedAsc[lo] + (pos - lo) * (sortedAsc[hi] - sortedAsc[lo]);
}

export const mean = (xs: number[]) => (xs.length ? xs.reduce((a, b) => a + b, 0) / xs.length : NaN);

// ── groupby(键排序 + NaN 键丢弃)─────────────────────────────────────────

export type Key = string | number;

export function cmpKey(a: Key, b: Key): number {
  if (typeof a === 'number' && typeof b === 'number') return a - b;
  const sa = String(a), sb = String(b);
  return sa < sb ? -1 : sa > sb ? 1 : 0;
}

/** groupby col:返回按键升序的 [key, 行号数组];null 键丢弃(pandas dropna=True) */
export function groupBy(keys: (Cell | null)[], indices: number[]): [Key, number[]][] {
  const m = new Map<Key, number[]>();
  for (const i of indices) {
    const k = keys[i];
    if (k == null) continue;
    const kk = (typeof k === 'number' ? k : String(k)) as Key;
    let arr = m.get(kk);
    if (!arr) m.set(kk, (arr = []));
    arr.push(i);
  }
  return [...m.entries()].sort((a, b) => cmpKey(a[0], b[0]));
}

// ── numpy argsort('quicksort') 的精确移植 ─────────────────────────────────
// pandas sort_values 单列排序走 numpy 的间接 introsort(npysort/quicksort.c.src
// 的 aquicksort_*):>15 元素用 median-of-3 快排(平局会被分区打乱,不稳定但
// 确定),≤15 用插入排序(稳定),递归过深退化堆排。平局顺序由该算法逐步决定,
// 想跟 Python 输出逐位一致只能原样移植——黄金测试会在真实数据上逐位验证。

const SMALL_QUICKSORT = 15;

/** FLOAT_LT:NaN 视为最大(numpy npy_sort 的 LT 语义) */
const LT = (a: number, b: number) => a < b || (b !== b && a === a);

function aheapsort(v: number[], tosort: number[], off: number, n: number) {
  // numpy aheapsort:1-based 索引访问 tosort[off-1 + i]
  const a = (i: number) => tosort[off + i - 1];
  const setA = (i: number, x: number) => { tosort[off + i - 1] = x; };
  let i: number, j: number, l: number, tmp: number;
  for (l = n >> 1; l > 0; --l) {
    tmp = a(l);
    for (i = l, j = l << 1; j <= n;) {
      if (j < n && LT(v[a(j)], v[a(j + 1)])) j += 1;
      if (LT(v[tmp], v[a(j)])) { setA(i, a(j)); i = j; j += i; }
      else break;
    }
    setA(i, tmp);
  }
  for (; n > 1;) {
    tmp = a(n);
    setA(n, a(1));
    n -= 1;
    for (i = 1, j = 2; j <= n;) {
      if (j < n && LT(v[a(j)], v[a(j + 1)])) j++;
      if (LT(v[tmp], v[a(j)])) { setA(i, a(j)); i = j; j += i; }
      else break;
    }
    setA(i, tmp);
  }
}

/** numpy ndarray.argsort(kind='quicksort') 的间接 introsort,返回升序索引 */
export function npArgsort(v: number[]): number[] {
  const n = v.length;
  const t: number[] = Array.from({ length: n }, (_, i) => i); // tosort
  if (n < 2) return t;
  const swap = (i: number, j: number) => { const x = t[i]; t[i] = t[j]; t[j] = x; };

  let depthLimit = (31 - Math.clz32(n)) * 2; // npy_get_msb(num) * 2
  const stack: number[] = [];
  let pl = 0;
  let pr = n - 1;

  for (;;) {
    while (pr - pl > SMALL_QUICKSORT) {
      if (depthLimit-- < 0) {
        aheapsort(v, t, pl, pr - pl + 1);
        break;
      }
      // median-of-3 选轴
      const pm = pl + ((pr - pl) >> 1);
      if (LT(v[t[pm]], v[t[pl]])) swap(pm, pl);
      if (LT(v[t[pr]], v[t[pm]])) swap(pr, pm);
      if (LT(v[t[pm]], v[t[pl]])) swap(pm, pl);
      const vp = v[t[pm]];
      let pi = pl;
      let pj = pr - 1;
      swap(pm, pj);
      for (;;) {
        do pi++; while (LT(v[t[pi]], vp));
        do pj--; while (LT(vp, v[t[pj]]));
        if (pi >= pj) break;
        swap(pi, pj);
      }
      swap(pi, pr - 1);
      // 大分区压栈,小分区继续
      if (pi - pl < pr - pi) {
        stack.push(pi + 1, pr);
        pr = pi - 1;
      } else {
        stack.push(pl, pi - 1);
        pl = pi + 1;
      }
    }
    if (pr - pl <= SMALL_QUICKSORT) {
      // 插入排序
      for (let pi = pl + 1; pi <= pr; ++pi) {
        const vi = t[pi];
        const vp = v[vi];
        let pj = pi;
        let pk = pi - 1;
        while (pj > pl && LT(vp, v[t[pk]])) t[pj--] = t[pk--];
        t[pj] = vi;
      }
    }
    if (stack.length === 0) break;
    pr = stack.pop()!;
    pl = stack.pop()!;
  }
  return t;
}

/**
 * pandas 单列 sort_values(ascending=False) = nargsort(ascending=False):
 * **先反转输入 → 升序 argsort(introsort)→ 结果再反转**。
 * 双重反转使小数组(插入排序,稳定)的平局保持原序;大数组平局由 introsort
 * 在反转后输入上的分区决定。与 Python 输出逐位一致(黄金测试验证)。
 */
export function sortValuesDesc<T>(rows: T[], val: (r: T) => number): T[] {
  const n = rows.length;
  const revVals: number[] = new Array(n);
  for (let i = 0; i < n; i++) revVals[i] = val(rows[n - 1 - i]);
  const idx = npArgsort(revVals);
  const out: T[] = new Array(n);
  for (let i = 0; i < n; i++) out[i] = rows[n - 1 - idx[n - 1 - i]];
  return out;
}

/**
 * 多列 sort_values([A, 金额], [asc, desc]) 走稳定 lexsort:
 * 平局保持原顺序(键升序)。color_by_cat 用这个。
 */
export function sortByDescStable<T>(rows: T[], val: (r: T) => number): T[] {
  return [...rows].sort((a, b) => val(b) - val(a));
}

/** Series.sum 语义:朴素左到右累加(pandas+bottleneck 的 nansum) */
export const sum = (idx: number[], arr: Float64Array) => {
  let s = 0;
  for (const i of idx) s += arr[i];
  return s;
};

/** groupby().sum() 语义:Kahan 补偿求和(pandas Cython group_sum 逐操作复刻)。
 *  与朴素累加差 ~1e-12,但 round(1) 在 .x5 边界会翻面,必须分开。 */
export const gsum = (idx: number[], arr: Float64Array) => {
  let s = 0, comp = 0;
  for (const i of idx) {
    const y = arr[i] - comp;
    const t = s + y;
    comp = t - s - y;
    s = t;
  }
  return s;
};

/** numpy pairwise 求和(umath pairwise_sum 原样移植):8 累加器 + 128 分块递归。
 *  pandas 无 bottleneck 时 Series.sum()/mean() 走这条路(与朴素/Kahan 都不同)。 */
export function npPairwiseSum(a: ArrayLike<number>, off = 0, n = a.length - off): number {
  if (n < 8) {
    let res = 0;
    for (let i = 0; i < n; i++) res += a[off + i];
    return res;
  }
  if (n <= 128) {
    const r = [a[off], a[off + 1], a[off + 2], a[off + 3], a[off + 4], a[off + 5], a[off + 6], a[off + 7]];
    let i = 8;
    for (; i + 8 <= n; i += 8) {
      r[0] += a[off + i]; r[1] += a[off + i + 1]; r[2] += a[off + i + 2]; r[3] += a[off + i + 3];
      r[4] += a[off + i + 4]; r[5] += a[off + i + 5]; r[6] += a[off + i + 6]; r[7] += a[off + i + 7];
    }
    let res = ((r[0] + r[1]) + (r[2] + r[3])) + ((r[4] + r[5]) + (r[6] + r[7]));
    for (; i < n; i++) res += a[off + i];
    return res;
  }
  let n2 = Math.floor(n / 2);
  n2 -= n2 % 8;
  return npPairwiseSum(a, off, n2) + npPairwiseSum(a, off + n2, n - n2);
}

/** Series.sum 语义(无 bottleneck):对选定行物化后 pairwise */
export const psum = (idx: number[], arr: Float64Array) => npPairwiseSum(idx.map((i) => arr[i]));
/** Series.mean 语义:pairwise 和 / 个数 */
export const pmean = (xs: number[]) => (xs.length ? npPairwiseSum(xs) / xs.length : NaN);

/** groupby().mean() 语义:Kahan 和 / 个数(pandas group_mean) */
export const gmean = (xs: number[]) => {
  if (!xs.length) return NaN;
  let s = 0, comp = 0;
  for (const x of xs) {
    const y = x - comp;
    const t = s + y;
    comp = t - s - y;
    s = t;
  }
  return s / xs.length;
};
export const nunique = (idx: number[], col: (Cell | null)[]) => {
  const s = new Set<Key>();
  for (const i of idx) {
    const v = col[i];
    if (v != null) s.add(typeof v === 'number' ? v : String(v));
  }
  return s.size;
};

// ── _season_type(preference_pipeline.py:172-180)─────────────────────────

export function seasonType(y: Cell | null): string {
  if (y == null) return '未知';
  const s = String(y);
  if (s.includes('订货会')) return '订货会';
  if (s.includes('2026') || (s.includes('2025') && (s.includes('春') || s.includes('夏')))) {
    return '当季新品';
  }
  return '往季/折扣';
}

// ── 输出类型(对应 R 的 19 个 key,字段名与 Python 输出一致)──────────────

export interface AnalyzeResult {
  summary: Record<string, number | string | PyFloat>;
  [k: string]: unknown;
}

// ── 主函数(_analyze)─────────────────────────────────────────────────────

export function analyzePreference(data: PreferenceData): AnalyzeResult {
  const { qty, amt, orderMs, cols, rawRowCount } = data;
  const col = (name: string) => cols.get(name) ?? null;
  const has = (name: string) => cols.has(name);

  // d = df.dropna(subset=[销售量, 销售金额])
  const idx: number[] = [];
  for (let i = 0; i < rawRowCount; i++) {
    if (!Number.isNaN(qty[i]) && !Number.isNaN(amt[i])) idx.push(i);
  }
  if (idx.length === 0) {
    throw new PipelineError(
      '销售量/销售金额都缺失，无法分析',
      '请确认上传的是有效销售明细文件，且至少有 1 条记录的净销售量/净销售金额不为空。',
    );
  }

  // 单价 = 金额/数量(±inf → NaN)
  const price = new Float64Array(rawRowCount).fill(NaN);
  for (const i of idx) {
    const p = amt[i] / qty[i];
    if (Number.isFinite(p)) price[i] = p;
  }
  // 季节类型
  const yearCol = col('年份');
  const season: string[] = new Array(rawRowCount);
  for (const i of idx) season[i] = yearCol ? seasonType(yearCol[i]) : '未知';

  const custCol = col('客户名称')!;

  // crank:客户销售额降序;tops = 前 30
  const crank = sortValuesDesc(
    groupBy(custCol, idx).map(([k, ii]) => ({ k, v: gsum(ii, amt), ii })),
    (r) => r.v,
  );
  const tops = crank.slice(0, TOP_N).map((r) => r.k);
  const topsSet = new Set(tops);

  const R: AnalyzeResult = { summary: {} };

  // ── summary ──
  const pricesAll: number[] = [];
  for (const i of idx) if (!Number.isNaN(price[i])) pricesAll.push(price[i]);
  let dateFrom = '', dateTo = '';
  if (orderMs) {
    let min = Infinity, max = -Infinity;
    for (const i of idx) {
      const v = orderMs[i];
      if (Number.isNaN(v)) continue;
      if (v < min) min = v;
      if (v > max) max = v;
    }
    if (min !== Infinity) {
      dateFrom = msToDateStr(min);
      dateTo = msToDateStr(max);
    }
  }
  R.summary = {
    records: idx.length,
    amount: pyRound(psum(idx, amt)),
    qty: pyInt(psum(idx, qty)),
    customers: nunique(idx, custCol),
    avg_price: pyFloat(pricesAll.length ? pyRound(pmean(pricesAll), 1) : 0.0),
    categories: has('分类') ? nunique(idx, col('分类')!) : 0,
    brands: has('品牌') ? nunique(idx, col('品牌')!) : 0,
    products: has('货号') ? nunique(idx, col('货号')!) : 0,
    shops: has('店铺') ? nunique(idx, col('店铺')!) : 0,
    designers: has('设计师品牌') ? nunique(idx, col('设计师品牌')!) : 0,
    date_from: dateFrom,
    date_to: dateTo,
  };

  // ── agg_top(col):数量/金额/客户数,金额降序 head(n) ──
  const aggTop = (name: string, n: number) => {
    if (!has(name)) return [];
    const g = groupBy(col(name)!, idx).map(([k, ii]) => ({
      k, 数量: gsum(ii, qty), 金额: gsum(ii, amt), 客户数: nunique(ii, custCol),
    }));
    return sortValuesDesc(g, (r) => r.金额).slice(0, n).map((r) => ({
      [name]: String(r.k), 数量: pyInt(r.数量), 金额: pyRound(r.金额), 客户数: r.客户数,
    }));
  };
  R.cat_dist = aggTop('分类', 20);
  R.color_dist = aggTop('颜色', 30);
  R.brand_dist = aggTop('品牌', 15);
  R.designer_dist = aggTop('设计师品牌', 15);
  R.shop_dist = aggTop('店铺', 20);

  // ── size_dist:尺码按销售量降序 head(10) ──
  R.size_dist = has('尺码')
    ? sortValuesDesc(
        groupBy(col('尺码')!, idx).map(([k, ii]) => ({ k, v: gsum(ii, qty) })),
        (r) => r.v,
      ).slice(0, 10).map((r) => ({ 尺码: String(r.k), 数量: pyInt(r.v) }))
    : [];

  // ── sales_dist:销售人员 ──
  R.sales_dist = has('销售')
    ? sortValuesDesc(
        groupBy(col('销售')!, idx).map(([k, ii]) => ({
          k, 金额: gsum(ii, amt), 客户数: nunique(ii, custCol),
        })),
        (r) => r.金额,
      ).slice(0, 20).map((r) => ({ 销售: String(r.k), 金额: pyRound(r.金额), 客户数: r.客户数 }))
    : [];

  // ── product_top:货号 Top30 ──
  R.product_top = has('货号')
    ? sortValuesDesc(
        groupBy(col('货号')!, idx).map(([k, ii]) => ({
          k, 金额: gsum(ii, amt), 数量: gsum(ii, qty), 客户数: nunique(ii, custCol),
        })),
        (r) => r.金额,
      ).slice(0, 30).map((r) => ({
        货号: String(r.k), 金额: pyRound(r.金额), 数量: pyInt(r.数量), 客户数: r.客户数,
      }))
    : [];

  // ── monthly:to_period('M'),NaT → "NaT" 组(pandas astype(str) 行为) ──
  if (orderMs) {
    const monthKeys: (string | null)[] = new Array(rawRowCount).fill(null);
    for (const i of idx) {
      monthKeys[i] = Number.isNaN(orderMs[i]) ? 'NaT' : msToDateStr(orderMs[i]).slice(0, 7);
    }
    R.monthly = groupBy(monthKeys, idx).map(([k, ii]) => ({
      月份: String(k), 金额: pyRound(gsum(ii, amt)), 客户数: nunique(ii, custCol),
    }));
  } else {
    R.monthly = [];
  }

  // ── price_band:pd.cut 右闭区间,observed=True 按分类序 ──
  const BANDS: [number, string][] = [
    [50, '0-50'], [100, '50-100'], [150, '100-150'], [200, '150-200'],
    [300, '200-300'], [500, '300-500'], [Infinity, '500+'],
  ];
  const bandOf = (p: number) => BANDS.find(([hi]) => p <= hi)![1];
  const bandKeys: (string | null)[] = new Array(rawRowCount).fill(null);
  const bandIdx: number[] = [];
  for (const i of idx) {
    if (!Number.isNaN(price[i]) && price[i] > 0) {
      bandKeys[i] = bandOf(price[i]);
      bandIdx.push(i);
    }
  }
  const bandOrder = new Map(BANDS.map(([, l], i) => [l, i]));
  R.price_band = groupBy(bandKeys, bandIdx)
    .sort((a, b) => bandOrder.get(String(a[0]))! - bandOrder.get(String(b[0]))!)
    .map(([k, ii]) => ({ 价格带: String(k), 订单数: ii.length, 金额: pyRound(gsum(ii, amt)) }));

  // ── cat_price:分类均价(dropna 单价,不滤 >0),平均降序,round(1) ──
  if (has('分类')) {
    const priceIdx = idx.filter((i) => !Number.isNaN(price[i]));
    const g = groupBy(col('分类')!, priceIdx).map(([k, ii]) => {
      const ps = ii.map((i) => price[i]).sort((a, b) => a - b);
      return { k, 平均: gmean(ps), 中位数: quantile(ps, 0.5) };
    });
    R.cat_price = sortValuesDesc(g, (r) => r.平均).map((r) => ({
      分类: String(r.k), 平均: pyFloat(npRound(r.平均, 1)), 中位数: pyFloat(npRound(r.中位数, 1)),
    }));
  } else {
    R.cat_price = [];
  }

  // ── color_by_cat:[分类,颜色] 聚合 → 分类升序+金额降序 → 每分类 head(5) ──
  if (has('分类') && has('颜色')) {
    const catCol = col('分类')!, colorCol = col('颜色')!;
    const out: { 分类: string; 颜色: string; 数量: number; 金额: number }[] = [];
    for (const [cat, catIdx] of groupBy(catCol, idx)) {
      const rows = sortByDescStable(
        groupBy(colorCol, catIdx).map(([color, ii]) => ({
          color, 数量: gsum(ii, qty), 金额: gsum(ii, amt),
        })),
        (r) => r.金额,
      ).slice(0, 5);
      for (const r of rows) {
        out.push({ 分类: String(cat), 颜色: String(r.color), 数量: pyInt(r.数量), 金额: pyRound(r.金额) });
      }
    }
    R.color_by_cat = out;
  } else {
    R.color_by_cat = [];
  }

  // ── size_by_cat:[分类,尺码] 聚合,groupby 自然序(双键升序) ──
  if (has('分类') && has('尺码')) {
    const catCol = col('分类')!, sizeCol = col('尺码')!;
    const out: { 分类: string; 尺码: string; 数量: number }[] = [];
    for (const [cat, catIdx] of groupBy(catCol, idx)) {
      for (const [sz, ii] of groupBy(sizeCol, catIdx)) {
        out.push({ 分类: String(cat), 尺码: String(sz), 数量: pyInt(gsum(ii, qty)) });
      }
    }
    R.size_by_cat = out;
  } else {
    R.size_by_cat = [];
  }

  // ── customer_top + tier_summary:全客户聚合 → 分位数分层 ──
  const ct = sortValuesDesc(
    groupBy(custCol, idx).map(([k, ii]) => {
      const ps = ii.map((i) => price[i]).filter((p) => !Number.isNaN(p));
      return { k, 金额: gsum(ii, amt), 数量: gsum(ii, qty), 均价: gmean(ps), 订单数: ii.length };
    }),
    (r) => r.金额,
  );
  const amounts = ct.map((r) => r.金额).sort((a, b) => a - b);
  const q75 = quantile(amounts, 0.75);
  const q50 = quantile(amounts, 0.5);
  const q25 = quantile(amounts, 0.25);
  const tierOf = (x: number) =>
    x >= q75 ? 'VIP' : x >= q50 ? '高价值' : x >= q25 ? '中等' : '低价值';
  const ctTiered = ct.map((r) => ({ ...r, 等级: tierOf(r.金额) }));
  R.customer_top = ctTiered.slice(0, 50).map((r) => ({
    客户: String(r.k),
    金额: pyRound(r.金额),
    数量: pyInt(r.数量),
    均价: pyFloat(Number.isNaN(r.均价) ? 0.0 : npRound(r.均价, 1)),
    等级: r.等级,
  }));

  // tier_summary:groupby 等级(键升序:VIP < 中等 < 低价值 < 高价值)
  const tierGroups = new Map<string, { n: number; total: number; comp: number }>();
  for (const r of ctTiered) {
    const g = tierGroups.get(r.等级) ?? { n: 0, total: 0, comp: 0 };
    g.n += 1;
    const y = r.金额 - g.comp;
    const t = g.total + y;
    g.comp = t - g.total - y;
    g.total = t;
    tierGroups.set(r.等级, g);
  }
  const tierTotal = [...tierGroups.values()].reduce((a, b) => a + b.total, 0) || 1;
  R.tier_summary = [...tierGroups.entries()]
    .sort((a, b) => cmpKey(a[0], b[0]))
    .map(([k, g]) => ({
      等级: k, 客户数: g.n, 总金额: pyRound(g.total), 占比: pyFloat(npRound((g.total / tierTotal) * 100, 1)),
    }));

  // ── brand_customer:tops 客户 × 品牌(双键升序) ──
  const topIdx = idx.filter((i) => {
    const v = custCol[i];
    return v != null && topsSet.has((typeof v === 'number' ? v : String(v)) as Key);
  });
  if (has('品牌')) {
    const brandCol = col('品牌')!;
    const out: { 客户: string; 品牌: string; 金额: number }[] = [];
    for (const [cust, ci] of groupBy(custCol, topIdx)) {
      for (const [brand, ii] of groupBy(brandCol, ci)) {
        out.push({ 客户: String(cust), 品牌: String(brand), 金额: pyRound(gsum(ii, amt)) });
      }
    }
    R.brand_customer = out;
  } else {
    R.brand_customer = [];
  }

  // ── sensitivity:tops 客户 × 季节类型 ──
  {
    const seasonArr: (string | null)[] = new Array(rawRowCount).fill(null);
    for (const i of idx) seasonArr[i] = season[i];
    const out: { 客户: string; 类型: string; 金额: number }[] = [];
    for (const [cust, ci] of groupBy(custCol, topIdx)) {
      for (const [tp, ii] of groupBy(seasonArr, ci)) {
        out.push({ 客户: String(cust), 类型: String(tp), 金额: pyRound(gsum(ii, amt)) });
      }
    }
    R.sensitivity = out;
  }

  // ── profiles:tops(销售额降序)逐客户画像 ──
  const custIdxMap = new Map<Key, number[]>(crank.map((r) => [r.k, r.ii]));
  const profiles: Record<string, unknown>[] = [];
  for (const cust of tops) {
    const ci = custIdxMap.get(cust)!;
    const ps = ci.map((i) => price[i]).filter((p) => !Number.isNaN(p));
    const p: Record<string, unknown> = {
      客户: String(cust),
      金额: pyRound(psum(ci, amt)),
      数量: pyInt(psum(ci, qty)),
      均价: ps.length ? pyFloat(pyRound(pmean(ps), 1)) : 0,
    };
    for (const [name, key] of [['分类', '品类'], ['品牌', '品牌'], ['设计师品牌', '设计师']] as const) {
      if (has(name)) {
        p[key] = sortValuesDesc(
          groupBy(col(name)!, ci).map(([k, ii]) => ({ k, v: gsum(ii, amt) })),
          (r) => r.v,
        ).slice(0, 3).map((r) => String(r.k));
      }
    }
    if (has('颜色')) {
      p['颜色'] = sortValuesDesc(
        groupBy(col('颜色')!, ci).map(([k, ii]) => ({ k, v: gsum(ii, qty) })),
        (r) => r.v,
      ).slice(0, 5).map((r) => String(r.k));
    }
    if (has('尺码')) {
      const szs = sortValuesDesc(
        groupBy(col('尺码')!, ci).map(([k, ii]) => ({ k, v: gsum(ii, qty) })),
        (r) => r.v,
      );
      p['尺码'] = szs.length ? String(szs[0].k) : '-';
    }
    const nwList: number[] = [];
    for (const i of ci) if (season[i] === '当季新品') nwList.push(amt[i]);
    const nw = npPairwiseSum(nwList);
    const tt = psum(ci, amt);
    const pct = tt > 0 ? pyRound((nw / tt) * 100, 1) : 0;
    p['新品占比'] = tt > 0 ? pyFloat(pct) : 0;
    p['类型'] = pct > 60 ? '追新型' : pct < 40 ? '折扣型' : '均衡型';
    profiles.push(p);
  }
  R.profiles = profiles;

  return R;
}
