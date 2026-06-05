// preferenceExcel.ts — 客户偏好分析数据.xlsx,逐行照搬 _build_excel
// (preference_pipeline.py:573-791)。最多 23 个 sheet(缺列自动跳过),
// pandas to_excel 在这条管线里不带任何样式(实测基准),纯数据 + 单表头行;
// pivot sheet 的 A1 放索引名。边界:NaN → 空格,±inf → 字符串 'inf'/'-inf'。
import ExcelJS from 'exceljs';
import {
  gmean, groupBy, gsum, npPairwiseSum, npRound, nunique, pmean, psum, pyInt, pyRound,
  quantile, seasonType, sortByDescStable, sortValuesDesc,
} from './preferenceAnalyze';
import type { Key } from './preferenceAnalyze';
import { msToDateStr } from './reader';
import type { PreferenceData } from './preferenceReader';
import type { Cell } from '../types/excel';

type V = Cell | null;

/** 单元格值规整:NaN→空,±inf→'inf'/'-inf'(xlsxwriter 实测行为) */
function cellValue(v: V): ExcelJS.CellValue {
  if (v == null) return null;
  if (v === '') return null; // xlsxwriter 把空字符串写成空白单元格
  if (typeof v === 'number') {
    if (Number.isNaN(v)) return null;
    if (v === Infinity) return 'inf';
    if (v === -Infinity) return '-inf';
  }
  return v as ExcelJS.CellValue;
}

function addSheet(wb: ExcelJS.Workbook, name: string, header: V[], rows: V[][]) {
  const ws = wb.addWorksheet(name);
  ws.addRow(header.map(cellValue));
  for (const r of rows) ws.addRow(r.map(cellValue));
}

export function buildPreferenceExcelWorkbook(data: PreferenceData): ExcelJS.Workbook {
  const { qty, amt, orderMs, cols, rawRowCount } = data;
  const col = (name: string) => cols.get(name) ?? null;
  const has = (name: string) => cols.has(name);

  // d = dropna(销售量, 销售金额);单价同 _build_excel
  const idx: number[] = [];
  for (let i = 0; i < rawRowCount; i++) {
    if (!Number.isNaN(qty[i]) && !Number.isNaN(amt[i])) idx.push(i);
  }
  const price = new Float64Array(rawRowCount).fill(NaN);
  for (const i of idx) {
    const p = amt[i] / qty[i];
    if (Number.isFinite(p)) price[i] = p;
  }
  const custCol = col('客户名称')!;

  // tops = 全部客户按销售额降序(全量方案)
  const custGroups = groupBy(custCol, idx);
  const custSorted = sortValuesDesc(
    custGroups.map(([k, ii]) => ({ k, ii, total: gsum(ii, amt) })),
    (r) => r.total,
  );
  const tops = custSorted.map((r) => r.k);
  const custIdxMap = new Map<Key, number[]>(custGroups);
  // tops 过滤后的行集(全部客户 = 全部非空客户行)
  const topIdx = idx.filter((i) => custCol[i] != null);

  /** 嵌套 groupby([A,B]):返回 [aKey, [bKey, idx[]][]][](双键升序) */
  const groupBy2 = (a: (Cell | null)[], b: (Cell | null)[], indices: number[]) =>
    groupBy(a, indices).map(([ak, ai]) => [ak, groupBy(b, ai)] as const);

  /** pivot_table(index, columns, values=sum, fill_value=0) → {colKeys, rows}。
   *  pandas 先 groupby([A,B]) 丢掉任一键为 NaN 的行,所以 B 全空的 A 不会出现在 index 里 */
  function pivot(aCol: (Cell | null)[], bCol: (Cell | null)[], values: Float64Array, indices: number[]) {
    const colKeySet = new Map<Key, number>();
    const g2 = groupBy2(aCol, bCol, indices).filter(([, bs]) => bs.length > 0);
    for (const [, bs] of g2) for (const [bk] of bs) if (!colKeySet.has(bk)) colKeySet.set(bk, 0);
    const colKeys = [...colKeySet.keys()].sort((x, y) =>
      typeof x === 'number' && typeof y === 'number' ? x - y : String(x) < String(y) ? -1 : String(x) > String(y) ? 1 : 0);
    const pos = new Map(colKeys.map((k, i) => [k, i]));
    const rows = g2.map(([ak, bs]) => {
      const vals = new Array<number>(colKeys.length).fill(0);
      for (const [bk, ii] of bs) vals[pos.get(bk)!] = gsum(ii, values);
      return { key: ak, vals };
    });
    return { colKeys, rows };
  }

  const wb = new ExcelJS.Workbook();

  // ── 1-品类 ──
  if (has('分类')) {
    const catCol = col('分类')!;
    const cat = sortValuesDesc(
      groupBy(catCol, idx).map(([k, ii]) => {
        const ps = ii.map((i) => price[i]).filter((p) => !Number.isNaN(p));
        return { k, 总数量: gsum(ii, qty), 总金额: gsum(ii, amt), 客户数: nunique(ii, custCol), 平均单价: gmean(ps) };
      }),
      (r) => r.总金额,
    );
    const totalAmt = npPairwiseSum(cat.map((r) => r.总金额)) || 1;
    addSheet(wb, '1-品类偏好', ['分类', '总数量', '总金额', '客户数', '平均单价', '金额占比%'],
      cat.map((r) => [String(r.k), r.总数量, r.总金额, r.客户数, npRound(r.平均单价, 1),
        npRound((r.总金额 / totalAmt) * 100, 1)]));

    const pv = pivot(custCol, catCol, amt, topIdx);
    if (pv.rows.length > 0) {
      const withTotal = sortValuesDesc(
        pv.rows.map((r) => ({ ...r, 合计: r.vals.reduce((a, b) => a + b, 0) })),
        (r) => r.合计,
      );
      addSheet(wb, '1-品类_客户矩阵', ['客户名称', ...pv.colKeys.map(String), '合计'],
        withTotal.map((r) => [r.key, ...r.vals, r.合计]));
    }
  }

  // ── 2-颜色(全量) ──
  if (has('颜色')) {
    const colorCol = col('颜色')!;
    addSheet(wb, '2-颜色全量', ['颜色', '总数量', '总金额', '客户数'],
      sortValuesDesc(
        groupBy(colorCol, idx).map(([k, ii]) => ({ k, 总数量: gsum(ii, qty), 总金额: gsum(ii, amt), 客户数: nunique(ii, custCol) })),
        (r) => r.总金额,
      ).map((r) => [r.k, r.总数量, r.总金额, r.客户数]));
    if (has('分类')) {
      const rows: V[][] = [];
      for (const [cat, bs] of groupBy2(col('分类')!, colorCol, idx)) {
        for (const r of sortByDescStable(
          bs.map(([color, ii]) => ({ color, 数量: gsum(ii, qty), 金额: gsum(ii, amt) })),
          (r) => r.金额,
        )) rows.push([cat, r.color, r.数量, r.金额]);
      }
      addSheet(wb, '2-颜色_品类交叉_全量', ['分类', '颜色', '数量', '金额'], rows);
    }
    {
      const rows: V[][] = [];
      for (const [cust, bs] of groupBy2(custCol, colorCol, topIdx)) {
        for (const r of sortByDescStable(
          bs.map(([color, ii]) => ({ color, v: gsum(ii, qty) })),
          (r) => r.v,
        )) rows.push([cust, r.color, r.v]);
      }
      addSheet(wb, '2-颜色_客户全量', ['客户名称', '颜色', '销售量'], rows);
    }
  }

  // ── 3-码数 ──
  if (has('尺码')) {
    const sizeCol = col('尺码')!;
    const sz = sortValuesDesc(
      groupBy(sizeCol, idx).map(([k, ii]) => ({ k, 总数量: gsum(ii, qty), 总金额: gsum(ii, amt) })),
      (r) => r.总数量,
    );
    const totalQ = npPairwiseSum(sz.map((r) => r.总数量)) || 1;
    addSheet(wb, '3-码数偏好', ['尺码', '总数量', '总金额', '占比%'],
      sz.map((r) => [r.k, r.总数量, r.总金额, npRound((r.总数量 / totalQ) * 100, 1)]));
    if (has('分类')) {
      const pv = pivot(col('分类')!, sizeCol, qty, idx);
      addSheet(wb, '3-码数_品类交叉', ['分类', ...pv.colKeys.map(String)],
        pv.rows.map((r) => [r.key, ...r.vals]));
    }
    {
      const bySize = new Map(groupBy2(custCol, sizeCol, topIdx));
      const rows: V[][] = [];
      for (const cust of tops) {
        const bs = bySize.get(cust);
        if (!bs) continue;
        const sub = sortValuesDesc(
          bs.map(([k, ii]) => ({ k, v: gsum(ii, qty) })),
          (r) => r.v,
        );
        if (sub.length) {
          rows.push([cust, sub[0].k, pyInt(sub[0].v), sub.length > 1 ? sub[1].k : '-']);
        }
      }
      if (rows.length) addSheet(wb, '3-客户尺码档案', ['客户', '主力尺码', '数量', '次选'], rows);
    }
  }

  // ── 4-客单价 ──
  const priceIdx = idx.filter((i) => !Number.isNaN(price[i]));
  if (has('分类')) {
    addSheet(wb, '4-品类均价', ['分类', '平均单价', '中位数单价'],
      sortValuesDesc(
        groupBy(col('分类')!, priceIdx).map(([k, ii]) => {
          const ps = ii.map((i) => price[i]).sort((a, b) => a - b);
          return { k, 平均: gmean(ps), 中位: quantile(ps, 0.5) };
        }),
        (r) => r.平均,
      ).map((r) => [String(r.k), npRound(r.平均, 1), npRound(r.中位, 1)]));
  }
  {
    const BANDS: [number, string][] = [
      [50, '0-50'], [100, '50-100'], [150, '100-150'], [200, '150-200'],
      [300, '200-300'], [500, '300-500'], [Infinity, '500+'],
    ];
    const bandKeys: (string | null)[] = new Array(rawRowCount).fill(null);
    const bandIdx: number[] = [];
    for (const i of priceIdx) {
      if (price[i] > 0) {
        bandKeys[i] = BANDS.find(([hi]) => price[i] <= hi)![1];
        bandIdx.push(i);
      }
    }
    const order = new Map(BANDS.map(([, l], i) => [l, i]));
    addSheet(wb, '4-价格带分布', ['价格带', '订单数', '总金额'],
      groupBy(bandKeys, bandIdx)
        .sort((a, b) => order.get(String(a[0]))! - order.get(String(b[0]))!)
        .map(([k, ii]) => [String(k), ii.length, gsum(ii, amt)]));
  }
  if (has('分类')) {
    const posIdx = topIdx.filter((i) => !Number.isNaN(price[i]) && price[i] > 0);
    const rows: V[][] = [];
    for (const [cust, bs] of groupBy2(custCol, col('分类')!, posIdx)) {
      for (const [cat, ii] of bs) {
        const ps = ii.map((i) => price[i]);
        // pandas .round(1) 作用于所有数值列(平均单价/拿货量/金额)
        rows.push([cust, cat, npRound(gmean(ps), 1), npRound(gsum(ii, qty), 1), npRound(gsum(ii, amt), 1)]);
      }
    }
    addSheet(wb, '4-客户品类单价', ['客户名称', '分类', '平均单价', '拿货量', '金额'], rows);
  }

  // ── 5-品牌 / 设计师 ──
  if (has('品牌')) {
    const brandCol = col('品牌')!;
    const br = sortValuesDesc(
      groupBy(brandCol, idx).map(([k, ii]) => ({ k, 总数量: gsum(ii, qty), 总金额: gsum(ii, amt), 客户数: nunique(ii, custCol) })),
      (r) => r.总金额,
    );
    const totalAmt = npPairwiseSum(br.map((r) => r.总金额)) || 1;
    addSheet(wb, '5-品牌偏好', ['品牌', '总数量', '总金额', '客户数', '占比%'],
      br.map((r) => [r.k, r.总数量, r.总金额, r.客户数, npRound((r.总金额 / totalAmt) * 100, 1)]));
    const pv = pivot(custCol, brandCol, amt, topIdx);
    if (pv.rows.length > 0) {
      addSheet(wb, '5-品牌_客户矩阵', ['客户名称', ...pv.colKeys.map(String)],
        pv.rows.map((r) => [r.key, ...r.vals]));
    }
  }
  if (has('设计师品牌')) {
    const dCol = col('设计师品牌')!;
    addSheet(wb, '5-设计师排名', ['设计师品牌', '总金额', '客户数'],
      sortValuesDesc(
        groupBy(dCol, idx).map(([k, ii]) => ({ k, 总金额: gsum(ii, amt), 客户数: nunique(ii, custCol) })),
        (r) => r.总金额,
      ).map((r) => [r.k, r.总金额, r.客户数]));
    {
      const rows: V[][] = [];
      for (const [cust, bs] of groupBy2(custCol, dCol, topIdx)) {
        for (const r of sortByDescStable(
          bs.map(([k, ii]) => ({ k, v: gsum(ii, amt) })),
          (r) => r.v,
        )) rows.push([cust, r.k, r.v]);
      }
      addSheet(wb, '5-客户设计师全量', ['客户名称', '设计师品牌', '销售金额'], rows);
    }
  }

  // ── 6-敏感度 / 年份 ──
  const yearCol = col('年份');
  const seasonArr: (string | null)[] = new Array(rawRowCount).fill(null);
  for (const i of idx) seasonArr[i] = yearCol ? seasonType(yearCol[i]) : '未知';
  {
    const pv = pivot(custCol, seasonArr, amt, topIdx);
    if (pv.rows.length > 0) {
      const ncIdx = pv.colKeys.findIndex((k) => k === '当季新品');
      const nc = ncIdx >= 0 ? ncIdx : 0;
      const rows = sortValuesDesc(
        pv.rows.map((r) => ({ ...r, 合计: r.vals.reduce((a, b) => a + b, 0) })),
        (r) => r.合计,
      ).map((r) => {
        const pct = npRound((r.vals[nc] / r.合计) * 100, 1); // 0/0→NaN, x/0→±inf(照搬)
        const tp = pct > 60 ? '追新型' : pct < 40 ? '折扣型' : '均衡型'; // NaN 比较为 false → 均衡型
        return [r.key, ...r.vals, r.合计, pct, tp] as V[];
      });
      addSheet(wb, '6-价格敏感度',
        ['客户名称', ...pv.colKeys.map(String), '合计', '新品占比%', '类型'], rows);
    }
  }
  if (has('年份')) {
    addSheet(wb, '6-年份分布', ['年份', '总金额', '客户数'],
      sortValuesDesc(
        groupBy(yearCol!, idx).map(([k, ii]) => ({ k, 总金额: gsum(ii, amt), 客户数: nunique(ii, custCol) })),
        (r) => r.总金额,
      ).map((r) => [r.k, r.总金额, r.客户数]));
  }

  // ── 7-客户分层 ──
  {
    const cu = sortValuesDesc(
      custGroups.map(([k, ii]) => {
        const ps = ii.map((i) => price[i]).filter((p) => !Number.isNaN(p));
        return { k, 总金额: gsum(ii, amt), 总数量: gsum(ii, qty), 均单价: gmean(ps), 订单数: ii.length };
      }),
      (r) => r.总金额,
    );
    const amounts = cu.map((r) => r.总金额).sort((a, b) => a - b);
    const q75 = quantile(amounts, 0.75), q50 = quantile(amounts, 0.5), q25 = quantile(amounts, 0.25);
    addSheet(wb, '7-客户分层', ['客户名称', '总金额', '总数量', '均单价', '订单数', '等级'],
      cu.map((r) => [r.k, r.总金额, r.总数量, npRound(r.均单价, 1), r.订单数,
        r.总金额 >= q75 ? 'VIP' : r.总金额 >= q50 ? '高价值' : r.总金额 >= q25 ? '中等' : '低价值']));
  }

  // ── 8-字段 ──
  if (has('店铺')) {
    addSheet(wb, '8-店铺', ['店铺', '总金额', '客户数'],
      sortValuesDesc(
        groupBy(col('店铺')!, idx).map(([k, ii]) => ({ k, 总金额: gsum(ii, amt), 客户数: nunique(ii, custCol) })),
        (r) => r.总金额,
      ).map((r) => [r.k, r.总金额, r.客户数]));
  }
  if (has('销售')) {
    addSheet(wb, '8-销售人员', ['销售', '总金额', '客户数'],
      sortValuesDesc(
        groupBy(col('销售')!, idx).map(([k, ii]) => ({ k, 总金额: gsum(ii, amt), 客户数: nunique(ii, custCol) })),
        (r) => r.总金额,
      ).map((r) => [r.k, r.总金额, r.客户数]));
  }
  if (orderMs) {
    const monthKeys: (string | null)[] = new Array(rawRowCount).fill(null);
    for (const i of idx) {
      monthKeys[i] = Number.isNaN(orderMs[i]) ? 'NaT' : msToDateStr(orderMs[i]).slice(0, 7);
    }
    addSheet(wb, '8-月度趋势', ['月份', '总金额', '客户数'],
      groupBy(monthKeys, idx).map(([k, ii]) => [String(k), gsum(ii, amt), nunique(ii, custCol)]));
  }
  if (has('货号')) {
    addSheet(wb, '8-热销货号全量', ['货号', '总金额', '总数量', '客户数'],
      sortValuesDesc(
        groupBy(col('货号')!, idx).map(([k, ii]) => ({ k, 总金额: gsum(ii, amt), 总数量: gsum(ii, qty), 客户数: nunique(ii, custCol) })),
        (r) => r.总金额,
      ).map((r) => [r.k, r.总金额, r.总数量, r.客户数]));
  }

  // ── 9-客户画像汇总(全部客户) ──
  {
    const header: V[] = ['客户', '金额', '数量', '均价'];
    if (has('分类')) header.push('品类');
    if (has('品牌')) header.push('品牌');
    if (has('设计师品牌')) header.push('设计师');
    if (has('颜色')) header.push('颜色');
    if (has('尺码')) header.push('尺码');
    header.push('新品占比', '类型');

    const rows: V[][] = [];
    for (const cust of tops) {
      const ci = custIdxMap.get(cust)!;
      const ps = ci.map((i) => price[i]).filter((p) => !Number.isNaN(p));
      const row: V[] = [String(cust), pyRound(psum(ci, amt)), pyInt(psum(ci, qty)),
        ps.length ? pyRound(pmean(ps), 1) : 0];
      for (const name of ['分类', '品牌', '设计师品牌'] as const) {
        if (has(name)) {
          row.push(sortValuesDesc(
            groupBy(col(name)!, ci).map(([k, ii]) => ({ k, v: gsum(ii, amt) })),
            (r) => r.v,
          ).slice(0, 3).map((r) => String(r.k)).join('、'));
        }
      }
      if (has('颜色')) {
        row.push(sortValuesDesc(
          groupBy(col('颜色')!, ci).map(([k, ii]) => ({ k, v: gsum(ii, qty) })),
          (r) => r.v,
        ).slice(0, 5).map((r) => String(r.k)).join('、'));
      }
      if (has('尺码')) {
        const szs = sortValuesDesc(
          groupBy(col('尺码')!, ci).map(([k, ii]) => ({ k, v: gsum(ii, qty) })),
          (r) => r.v,
        );
        row.push(szs.length ? String(szs[0].k) : '-');
      }
      const nwList: number[] = [];
      for (const i of ci) if (seasonArr[i] === '当季新品') nwList.push(amt[i]);
      const nw = npPairwiseSum(nwList);
      const tt = psum(ci, amt);
      const pct = tt > 0 ? pyRound((nw / tt) * 100, 1) : 0;
      row.push(pct, pct > 60 ? '追新型' : pct < 40 ? '折扣型' : '均衡型');
      rows.push(row);
    }
    if (rows.length) addSheet(wb, '9-客户画像汇总', header, rows);
  }

  return wb;
}

export async function buildPreferenceExcel(data: PreferenceData): Promise<Uint8Array> {
  const wb = buildPreferenceExcelWorkbook(data);
  const buf = await wb.xlsx.writeBuffer();
  return new Uint8Array(buf as ArrayBuffer);
}
