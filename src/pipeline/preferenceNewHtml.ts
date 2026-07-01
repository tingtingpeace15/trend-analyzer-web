// preferenceNewHtml.ts — 新版客户偏好分析 HTML 报告。
// 新增链路使用，不替换旧版 buildPreferenceHtml。
import { pyJsonDumps } from './pyjson';
import { msToDateStr, pandasColumns, readWorkbook, sheetToRows, sheetWidth } from './reader';
import echartsSource from './echarts.min.js?raw';
import type { AnalyzeResult } from './preferenceAnalyze';
import type { PreferenceData } from './preferenceReader';
import type { Cell } from '../types/excel';

const PRICE_BANDS = [
  { label: '0-50', hi: 50 },
  { label: '50-100', hi: 100 },
  { label: '100-150', hi: 150 },
  { label: '150-200', hi: 200 },
  { label: '200-300', hi: 300 },
  { label: '300-500', hi: 500 },
  { label: '500+', hi: Infinity },
] as const;

const ACCEPTANCE_PRICE_BANDS = ['0-50', '50-80', '80-100', '100-150', '150-200', '200-300', '300-500', '500+'] as const;
const ACCEPTANCE_SEGMENTS = ['低价带', '主流价带', '高价带'] as const;

const INLINE_ECHARTS_SCRIPT = echartsSource.replace(/<\/script/gi, '<\\/script');

type PriceSegmentation = {
  lowCutoff: number;
  highCutoff: number;
};

export function loadPreferenceOrderIds(bytes: ArrayBuffer | Uint8Array): (Cell | null)[] {
  return loadAlignedPreviewColumnValues(bytes, ['销退单ID', '订单号', '订单ID', '单据编号']);
}

export function buildNewPreferenceHtml(data: PreferenceData, orderIds: (Cell | null)[], base: AnalyzeResult): string {
  const R = { ...base } as Record<string, unknown>;
  const priceSegmentation = buildPriceSegmentation(data);
  R.overview_analysis = buildOverviewAnalysis(data, orderIds);
  R.category_preference_analysis = buildCategoryPreferenceAnalysis(data, orderIds);
  R.color_preference_analysis = buildColorPreferenceAnalysis(data);
  R.size_preference_analysis = buildSizePreferenceAnalysis(data);
  R.price_analysis = buildCustomerCategoryPriceAnalysis(data, orderIds);
  R.price_acceptance_analysis = buildPriceAcceptanceAnalysis(data, orderIds, priceSegmentation);
  R.brand_style_analysis = buildBrandStyleAnalysis(data);
  R.season_preference_analysis = buildSeasonPreferenceAnalysis(data, orderIds);
  R.customer_visual_profiles = buildCustomerVisualProfiles(data, orderIds, priceSegmentation);
  return buildPreferenceContentHtml(R);
}

function normalizePreviewHeader(name: string): string {
  return name.trim().replace(/[:：]+$/, '');
}

function loadAlignedPreviewColumnValues(bytes: ArrayBuffer | Uint8Array, aliases: string[]): (Cell | null)[] {
  const wb = readWorkbook(bytes);
  type Frame = { columns: string[]; rows: Cell[][]; start: number };
  const frames: Frame[] = [];
  const sniff = ['客户名称', '销售', '下单时间', '货号'];

  for (const name of wb.SheetNames) {
    const ws = wb.Sheets[name];
    const rows = sheetToRows(ws);
    const width = sheetWidth(ws);
    let hrow: number | null = null;
    for (let i = 0; i < Math.min(3, rows.length); i++) {
      const joined = rows[i].map((v) => (v == null ? 'nan' : String(v))).join(' ');
      if (sniff.some((k) => joined.includes(k))) {
        hrow = i;
        break;
      }
    }
    const headerIdx = hrow ?? 0;
    const columns = pandasColumns(rows[headerIdx] ?? [], width);
    if (columns.length >= 10) frames.push({ columns, rows, start: headerIdx + 1 });
  }

  const baseCols = frames[0]?.columns;
  if (!baseCols) return [];
  for (let i = 1; i < frames.length; i++) {
    if (frames[i].columns.length >= baseCols.length) frames[i] = { ...frames[i], columns: baseCols };
  }

  const matchColumn = (name: string) => {
    const normalized = normalizePreviewHeader(name);
    return aliases.some((alias) => normalized === alias || normalized.includes(alias));
  };
  const values: (Cell | null)[] = [];
  for (const frame of frames) {
    const idx = frame.columns.findIndex(matchColumn);
    for (let r = frame.start; r < frame.rows.length; r++) values.push(idx >= 0 ? frame.rows[r][idx] ?? null : null);
  }
  return values;
}

function cellText(v: Cell | null): string | null {
  if (v == null) return null;
  const s = String(v).trim();
  return s ? s : null;
}

function round1(v: number): number {
  return Math.round(v * 10) / 10;
}

function percentile(sortedAsc: number[], q: number): number {
  if (sortedAsc.length === 0) return 0;
  const pos = (sortedAsc.length - 1) * q;
  const lo = Math.floor(pos);
  const hi = Math.ceil(pos);
  if (lo === hi) return sortedAsc[lo];
  return sortedAsc[lo] + (pos - lo) * (sortedAsc[hi] - sortedAsc[lo]);
}

function priceBand(price: number): string {
  return PRICE_BANDS.find((b) => price <= b.hi)?.label ?? '500+';
}

function acceptancePriceBand(price: number): string {
  if (price <= 50) return '0-50';
  if (price < 80) return '50-80';
  if (price <= 100) return '80-100';
  if (price <= 150) return '100-150';
  if (price <= 200) return '150-200';
  if (price <= 300) return '200-300';
  if (price <= 500) return '300-500';
  return '500+';
}

function buildPriceSegmentation(data: PreferenceData): PriceSegmentation {
  const values: { price: number; qty: number }[] = [];
  for (let i = 0; i < data.rawRowCount; i++) {
    const qty = data.qty[i];
    const amount = data.amt[i];
    if (!Number.isFinite(qty) || !Number.isFinite(amount) || qty <= 0 || amount <= 0) continue;
    const price = amount / qty;
    if (Number.isFinite(price) && price > 0) values.push({ price, qty });
  }
  const lowCutoff = round1(weightedPricePercentile(values, 1 / 3));
  const highCutoff = round1(weightedPricePercentile(values, 2 / 3));
  if (!Number.isFinite(lowCutoff) || !Number.isFinite(highCutoff) || lowCutoff <= 0 || highCutoff <= lowCutoff) {
    return { lowCutoff: 80, highCutoff: 150 };
  }
  return { lowCutoff, highCutoff };
}

function acceptanceSegment(price: number, segmentation: PriceSegmentation): string {
  if (price < segmentation.lowCutoff) return '低价带';
  if (price < segmentation.highCutoff) return '主流价带';
  return '高价带';
}

function priceTagClass(tag: string): number {
  return tag === '低价敏感型' ? 0
    : tag === '中价稳定型' ? 1
      : tag === '高价接受型' ? 2
        : tag === '价格波动型' ? 3
          : 4;
}

function orderKey(orderIds: (Cell | null)[], index: number, customer: string): string {
  const id = cellText(orderIds[index] ?? null);
  return id ? `${customer}\u0000${id}` : `${customer}\u0000__line_${index}`;
}

function weightedPricePercentile(values: { price: number; qty: number }[], q: number): number {
  const rows = values
    .filter((r) => Number.isFinite(r.price) && Number.isFinite(r.qty) && r.price > 0 && r.qty > 0)
    .sort((a, b) => a.price - b.price);
  const totalQty = rows.reduce((s, r) => s + r.qty, 0);
  if (rows.length === 0 || totalQty <= 0) return 0;
  const target = totalQty * q;
  let acc = 0;
  for (const r of rows) {
    acc += r.qty;
    if (acc >= target) return r.price;
  }
  return rows[rows.length - 1].price;
}

function styleOf(brand: string, year: string): string {
  const s = `${brand} ${year}`;
  if (s.includes('废板')) return '废板';
  if (s.includes('客户定制')) return '客户定制';
  if (s.includes('热卖')) return '热卖款';
  if (s.includes('订货会')) return '订货会';
  if (s.includes('旧') || s.includes('2023') || s.includes('2024')) return '往季/旧款';
  if (s.includes('2026') || s.includes('2025春') || s.includes('2025夏')) return '当季新品';
  return '常规款';
}

function addMetric(map: Map<string, { amount: number; qty: number; customers: Set<string> }>, key: string, amount: number, qty: number, customer: string) {
  let g = map.get(key);
  if (!g) {
    g = { amount: 0, qty: 0, customers: new Set<string>() };
    map.set(key, g);
  }
  g.amount += amount;
  g.qty += qty;
  g.customers.add(customer);
}

function hasPositiveNet(amount: number, qty: number): boolean {
  return Math.round(amount) > 0 && Math.round(qty) > 0;
}

function isPositiveSale(amount: number, qty: number): boolean {
  return amount > 0 && qty > 0;
}

function daysBetween(fromMs: number, toMs: number): number {
  if (!Number.isFinite(fromMs) || !Number.isFinite(toMs)) return 0;
  return Math.max(0, Math.round((toMs - fromMs) / (24 * 60 * 60 * 1000)));
}

function topEntries(map: Map<string, number>, n: number): string[] {
  return [...map.entries()]
    .sort((a, b) => b[1] - a[1])
    .slice(0, n)
    .map(([k]) => k);
}

function buildOverviewAnalysis(data: PreferenceData, orderIds: (Cell | null)[]) {
  const custCol = data.cols.get('客户名称');
  const shopCol = data.cols.get('店铺');
  const catCol = data.cols.get('分类');
  const brandCol = data.cols.get('品牌');
  const designerCol = data.cols.get('设计师品牌');
  const colorCol = data.cols.get('颜色');
  const sizeCol = data.cols.get('尺码');
  const productCol = data.cols.get('货号');

  const monthly = new Map<string, { amount: number; qty: number; customers: Set<string> }>();
  const shops = new Map<string, { amount: number; qty: number; customers: Set<string> }>();
  const colors = new Map<string, { amount: number; qty: number; customers: Set<string> }>();
  const sizes = new Map<string, { amount: number; qty: number; customers: Set<string> }>();
  const categories = new Map<string, { amount: number; qty: number; customers: Set<string> }>();
  const brands = new Map<string, { amount: number; qty: number; customers: Set<string> }>();
  const designers = new Map<string, { amount: number; qty: number; customers: Set<string> }>();
  const brandDesigners = new Map<string, { amount: number; qty: number; customers: Set<string> }>();
  const customers = new Map<string, { amount: number; qty: number; customers: Set<string> }>();
  const orderKeys = new Set<string>();
  const productCustomerOrders = new Map<string, Set<string>>();
  const productMetrics = new Map<string, { amount: number; qty: number; customers: Set<string> }>();

  for (let i = 0; i < data.rawRowCount; i++) {
    const amount = data.amt[i];
    const qty = data.qty[i];
    if (!Number.isFinite(amount) || !Number.isFinite(qty) || (amount === 0 && qty === 0)) continue;

    const customer = cellText(custCol?.[i] ?? null) ?? '未标记';
    const shop = cellText(shopCol?.[i] ?? null) ?? '未标记';
    const category = cellText(catCol?.[i] ?? null) ?? '未分类';
    const brand = cellText(brandCol?.[i] ?? null) ?? '未标记';
    const designer = cellText(designerCol?.[i] ?? null) ?? '未标记';
    const color = normalizeColorName(colorCol?.[i] ?? null);
    const size = normalizeSizeName(sizeCol?.[i] ?? null);
    const product = cellText(productCol?.[i] ?? null) ?? '未标记';
    const ms = data.orderMs?.[i];
    const month = ms != null && !Number.isNaN(ms) ? msToDateStr(ms).slice(0, 7) : '未标记';
    const ok = orderKey(orderIds, i, customer);

    orderKeys.add(ok);
    addMetric(monthly, month, amount, qty, customer);
    addMetric(shops, shop, amount, qty, customer);
    addMetric(colors, color, amount, qty, customer);
    addMetric(sizes, size, amount, qty, customer);
    addMetric(categories, category, amount, qty, customer);
    addMetric(brands, brand, amount, qty, customer);
    addMetric(designers, designer, amount, qty, customer);
    addMetric(brandDesigners, `${brand}\u0000${designer}`, amount, qty, customer);
    addMetric(customers, customer, amount, qty, customer);
    addMetric(productMetrics, product, amount, qty, customer);
    if (isPositiveSale(amount, qty)) {
      const repeatKey = `${product}\u0000${customer}`;
      let productOrders = productCustomerOrders.get(repeatKey);
      if (!productOrders) {
        productOrders = new Set<string>();
        productCustomerOrders.set(repeatKey, productOrders);
      }
      productOrders.add(ok);
    }
  }

  const metricRows = (map: Map<string, { amount: number; qty: number; customers: Set<string> }>, nameKey: string) => [...map.entries()]
    .map(([name, g]) => ({
      [nameKey]: name,
      name,
      amount: Math.round(g.amount),
      qty: Math.round(g.qty),
      customers: g.customers.size,
    }))
    .filter((r) => hasPositiveNet(Number(r.amount), Number(r.qty)))
    .sort((a, b) => b.amount - a.amount);

  const monthlySales = [...monthly.entries()]
    .map(([month, g]) => ({ month, amount: Math.round(g.amount), qty: Math.round(g.qty), customers: g.customers.size }))
    .sort((a, b) => (a.month === '未标记' ? 1 : b.month === '未标记' ? -1 : a.month.localeCompare(b.month)));
  const shopSales = metricRows(shops, 'shop');
  const colorRows = metricRows(colors, 'name').sort((a, b) => b.qty - a.qty || b.amount - a.amount);
  const sizeRows = metricRows(sizes, 'name').sort((a, b) => b.qty - a.qty || b.amount - a.amount);
  const categoryRows = metricRows(categories, 'name').sort((a, b) => b.qty - a.qty || b.amount - a.amount);
  const brandRows = metricRows(brands, 'name').sort((a, b) => b.qty - a.qty || b.amount - a.amount);
  const designerRows = metricRows(designers, 'name');
  const customerRows = metricRows(customers, 'name');
  const repeatCustomerCounts = new Map<string, number>();
  for (const [key, orders] of productCustomerOrders.entries()) {
    if (orders.size < 2) continue;
    const [product] = key.split('\u0000');
    repeatCustomerCounts.set(product, (repeatCustomerCounts.get(product) ?? 0) + 1);
  }
  const repeatProductRows = [...repeatCustomerCounts.entries()]
    .map(([product, repeat_customers]) => {
      const metric = productMetrics.get(product);
      return {
        product,
        repeat_customers,
        amount: Math.round(metric?.amount ?? 0),
        qty: Math.round(metric?.qty ?? 0),
        customers: metric?.customers.size ?? 0,
      };
    })
    .filter((r) => r.repeat_customers > 0 && hasPositiveNet(r.amount, r.qty))
    .sort((a, b) => b.repeat_customers - a.repeat_customers || b.amount - a.amount || b.qty - a.qty)
    .slice(0, 30);
  const makeItem = (label: string, row: { name: string; amount: number; qty: number; customers: number } | undefined, note: string, hex?: string) => ({
    label,
    name: row?.name ?? '-',
    amount: row?.amount ?? 0,
    qty: row?.qty ?? 0,
    customers: row?.customers ?? 0,
    note,
    hex,
  });

  return {
    order_count: orderKeys.size,
    detail_rows: data.rawRowCount,
    monthly_sales: monthlySales,
    shop_sales: shopSales,
    repeat_products: repeatProductRows,
    summary_items: [
      makeItem('热销颜色', colorRows[0], '按净销售量最高', colorHex(colorRows[0]?.name ?? '')),
      makeItem('热销尺码', sizeRows[0], '按净销售量最高'),
      makeItem('热销品类', categoryRows[0], '按净销售量最高'),
      makeItem('热销品牌', brandRows[0], '按净销售量最高'),
      makeItem('设计师品牌Top1', designerRows[0], '按净销售金额最高'),
      makeItem('贡献Top1客户', customerRows[0], '按净销售金额最高'),
    ],
  };
}

function buildBrandStyleAnalysis(data: PreferenceData) {
  const custCol = data.cols.get('客户名称');
  const shopCol = data.cols.get('店铺');
  const brandCol = data.cols.get('品牌');
  const designerCol = data.cols.get('设计师品牌');
  const catCol = data.cols.get('分类');
  const yearCol = data.cols.get('年份');
  if (!custCol || !brandCol || !designerCol) {
    return { brand_designer: [], brand_style: [], designer_brand: [], category_designer: [], customer_profiles: [], styles: [] };
  }

  const pair = new Map<string, { amount: number; qty: number; customers: Set<string> }>();
  const brandStyle = new Map<string, { amount: number; qty: number; customers: Set<string> }>();
  const designerBrand = new Map<string, { amount: number; qty: number; customers: Set<string> }>();
  const categoryDesigner = new Map<string, { amount: number; qty: number; customers: Set<string> }>();
  const shopBrand = new Map<string, { amount: number; qty: number; customers: Set<string> }>();
  const brandSummary = new Map<string, { amount: number; qty: number; customers: Set<string> }>();
  const designerSummary = new Map<string, { amount: number; qty: number; customers: Set<string> }>();
  const customerTotals = new Map<string, { qty: number; amount: number }>();
  const customerBrands = new Map<string, Map<string, { qty: number; amount: number }>>();
  const customerMap = new Map<string, {
    amount: number;
    qty: number;
    brands: Map<string, number>;
    designers: Map<string, number>;
    styles: Map<string, number>;
  }>();
  const brandAmount = new Map<string, number>();
  const designerAmount = new Map<string, number>();
  const styleSet = new Set<string>();

  for (let i = 0; i < data.rawRowCount; i++) {
    const amount = data.amt[i];
    const qty = data.qty[i];
    if (!Number.isFinite(amount) || !Number.isFinite(qty) || (amount === 0 && qty === 0)) continue;
    const customer = cellText(custCol[i]);
    const brand = cellText(brandCol[i]) ?? '未标记';
    const designer = cellText(designerCol[i]) ?? '未标记';
    if (!customer) continue;
    const shop = shopCol ? cellText(shopCol[i]) ?? '未标记' : '未标记';
    const category = catCol ? cellText(catCol[i]) ?? '未分类' : '未分类';
    const year = yearCol ? cellText(yearCol[i]) ?? '' : '';
    const style = styleOf(brand, year);
    styleSet.add(style);

    addMetric(pair, `${brand}\u0000${designer}`, amount, qty, customer);
    addMetric(brandStyle, `${brand}\u0000${style}`, amount, qty, customer);
    addMetric(designerBrand, `${designer}\u0000${brand}`, amount, qty, customer);
    addMetric(categoryDesigner, `${category}\u0000${designer}`, amount, qty, customer);
    addMetric(shopBrand, `${shop}\u0000${brand}`, amount, qty, customer);
    addMetric(brandSummary, brand, amount, qty, customer);
    addMetric(designerSummary, designer, amount, qty, customer);
    brandAmount.set(brand, (brandAmount.get(brand) ?? 0) + amount);
    designerAmount.set(designer, (designerAmount.get(designer) ?? 0) + amount);

    let customerTotal = customerTotals.get(customer);
    if (!customerTotal) {
      customerTotal = { qty: 0, amount: 0 };
      customerTotals.set(customer, customerTotal);
    }
    customerTotal.qty += qty;
    customerTotal.amount += amount;

    let customerBrandMap = customerBrands.get(customer);
    if (!customerBrandMap) {
      customerBrandMap = new Map<string, { qty: number; amount: number }>();
      customerBrands.set(customer, customerBrandMap);
    }
    const customerBrand = customerBrandMap.get(brand) ?? { qty: 0, amount: 0 };
    customerBrand.qty += qty;
    customerBrand.amount += amount;
    customerBrandMap.set(brand, customerBrand);

    let cg = customerMap.get(customer);
    if (!cg) {
      cg = { amount: 0, qty: 0, brands: new Map(), designers: new Map(), styles: new Map() };
      customerMap.set(customer, cg);
    }
    cg.amount += amount;
    cg.qty += qty;
    cg.brands.set(brand, (cg.brands.get(brand) ?? 0) + amount);
    cg.designers.set(designer, (cg.designers.get(designer) ?? 0) + amount);
    cg.styles.set(style, (cg.styles.get(style) ?? 0) + amount);
  }

  const topBrands = topEntries(brandAmount, 10);
  const topDesigners = topEntries(designerAmount, 12);
  const styles = ['当季新品', '订货会', '热卖款', '往季/旧款', '客户定制', '废板', '常规款']
    .filter((s) => styleSet.has(s));

  const brandSummaryRows = [...brandSummary.entries()]
    .map(([brand, g]) => ({ brand, amount: Math.round(g.amount), qty: Math.round(g.qty), customers: g.customers.size }))
    .filter((r) => hasPositiveNet(r.amount, r.qty))
    .sort((a, b) => b.amount - a.amount);

  const designerSummaryRows = [...designerSummary.entries()]
    .map(([designer, g]) => ({ designer, amount: Math.round(g.amount), qty: Math.round(g.qty), customers: g.customers.size }))
    .filter((r) => hasPositiveNet(r.amount, r.qty))
    .sort((a, b) => b.amount - a.amount);

  const brandDesigner = [...pair.entries()]
    .map(([key, g]) => {
      const [brand, designer] = key.split('\u0000');
      return { brand, designer, amount: Math.round(g.amount), qty: Math.round(g.qty), customers: g.customers.size };
    })
    .filter((r) => hasPositiveNet(r.amount, r.qty))
    .sort((a, b) => b.amount - a.amount);

  const categoryDesignerRows = [...categoryDesigner.entries()]
    .map(([key, g]) => {
      const [category, designer] = key.split('\u0000');
      return { category, designer, amount: Math.round(g.amount), qty: Math.round(g.qty), customers: g.customers.size };
    })
    .filter((r) => hasPositiveNet(r.amount, r.qty))
    .sort((a, b) => b.amount - a.amount);

  const shopBestBrandRows = [...shopBrand.entries()]
    .map(([key, g]) => {
      const [shop, brand] = key.split('\u0000');
      return { shop, brand, amount: Math.round(g.amount), qty: Math.round(g.qty), customers: g.customers.size };
    })
    .filter((r) => hasPositiveNet(r.amount, r.qty))
    .sort((a, b) => a.shop.localeCompare(b.shop) || b.amount - a.amount)
    .reduce<{ shop: string; brand: string; amount: number; qty: number; customers: number }[]>((rows, row) => {
      if (!rows.some((r) => r.shop === row.shop)) rows.push(row);
      return rows;
    }, [])
    .sort((a, b) => b.amount - a.amount);

  const brandStyleRows = topBrands.flatMap((brand) => styles.map((style) => {
    const g = brandStyle.get(`${brand}\u0000${style}`);
    return { brand, style, amount: Math.round(g?.amount ?? 0), qty: Math.round(g?.qty ?? 0), customers: g?.customers.size ?? 0 };
  })).filter((r) => hasPositiveNet(r.amount, r.qty));

  const designerBrandRows = topDesigners.flatMap((designer) => topBrands.map((brand) => {
    const g = designerBrand.get(`${designer}\u0000${brand}`);
    return { designer, brand, amount: Math.round(g?.amount ?? 0), qty: Math.round(g?.qty ?? 0), customers: g?.customers.size ?? 0 };
  })).filter((r) => hasPositiveNet(r.amount, r.qty));

  const customerProfiles = [...customerMap.entries()]
    .filter(([, g]) => hasPositiveNet(g.amount, g.qty))
    .map(([customer, g]) => {
      const topBrand = [...g.brands.entries()].sort((a, b) => b[1] - a[1])[0];
      const topDesigner = [...g.designers.entries()].sort((a, b) => b[1] - a[1])[0];
      const topStyle = [...g.styles.entries()].sort((a, b) => b[1] - a[1])[0];
      return {
        customer,
        amount: Math.round(g.amount),
        qty: Math.round(g.qty),
        avg_price: round1(g.qty ? g.amount / g.qty : 0),
        brands: topEntries(g.brands, 3),
        designers: topEntries(g.designers, 3),
        styles: topEntries(g.styles, 3),
        main_brand_share: topBrand ? round1((topBrand[1] / g.amount) * 100) : 0,
        main_designer_share: topDesigner ? round1((topDesigner[1] / g.amount) * 100) : 0,
        main_style_share: topStyle ? round1((topStyle[1] / g.amount) * 100) : 0,
      };
    })
    .sort((a, b) => b.amount - a.amount)
    .slice(0, 40);

  const topCustomerByAmount = [...customerTotals.entries()]
    .filter(([, total]) => total.amount > 0)
    .sort((a, b) => b[1].amount - a[1].amount || b[1].qty - a[1].qty)
    .slice(0, 20)
    .map(([customer]) => customer);
  const topCustomerBrandTotals = new Map<string, { qty: number; amount: number; customers: Set<string> }>();
  let topCustomerQty = 0;
  let topCustomerAmount = 0;
  for (const customer of topCustomerByAmount) {
    const customerTotal = customerTotals.get(customer);
    if (customerTotal) {
      topCustomerQty += customerTotal.qty;
      topCustomerAmount += customerTotal.amount;
    }
    for (const [brand, metric] of customerBrands.get(customer) ?? new Map()) {
      if (!hasPositiveNet(metric.amount, metric.qty)) continue;
      const row = topCustomerBrandTotals.get(brand) ?? { qty: 0, amount: 0, customers: new Set<string>() };
      row.qty += metric.qty;
      row.amount += metric.amount;
      row.customers.add(customer);
      topCustomerBrandTotals.set(brand, row);
    }
  }
  const topCustomerBrandRows = [...topCustomerBrandTotals.entries()]
    .filter(([, metric]) => hasPositiveNet(metric.amount, metric.qty))
    .sort((a, b) => b[1].qty - a[1].qty || b[1].amount - a[1].amount)
    .map(([brand, metric], index) => ({
      rank: index + 1,
      brand,
      qty: Math.round(metric.qty),
      amount: Math.round(metric.amount),
      customers: metric.customers.size,
      qty_share: topCustomerQty ? metric.qty / topCustomerQty : 0,
      amount_share: topCustomerAmount ? metric.amount / topCustomerAmount : 0,
    }));

  return {
    brand_designer: brandDesigner,
    brand_style: brandStyleRows,
    designer_brand: designerBrandRows,
    category_designer: categoryDesignerRows,
    shop_best_brands: shopBestBrandRows,
    customer_profiles: customerProfiles,
    top_customer_brands: topCustomerBrandRows,
    brand_summary: brandSummaryRows,
    designer_summary: designerSummaryRows,
    styles,
    brands: topBrands,
    designers: topDesigners,
  };
}

function buildSeasonPreferenceAnalysis(data: PreferenceData, orderIds: (Cell | null)[]) {
  const custCol = data.cols.get('客户名称');
  const yearCol = data.cols.get('年份');
  const catCol = data.cols.get('分类');
  const productCol = data.cols.get('货号');
  const brandCol = data.cols.get('品牌');
  const designerCol = data.cols.get('设计师品牌');
  if (!custCol || !yearCol) {
    return {
      customers: [],
      customer_type_summary: [],
      order_type_summary: [],
      booking_products: [],
      booking_categories: [],
      booking_brands: [],
      booking_designers: [],
      high_potential_cutoff: 0,
    };
  }

  type CustomerSeason = {
    customer: string;
    normalAmount: number;
    normalQty: number;
    bookingAmount: number;
    bookingQty: number;
    orderKeys: Set<string>;
  };
  type HotMetric = {
    amount: number;
    qty: number;
    customers: Set<string>;
    orders: Set<string>;
    lines: number;
  };

  const customers = new Map<string, CustomerSeason>();
  const productMap = new Map<string, HotMetric>();
  const categoryMap = new Map<string, HotMetric>();
  const brandMap = new Map<string, HotMetric>();
  const designerMap = new Map<string, HotMetric>();
  const summary = {
    normal: { amount: 0, qty: 0, customers: new Set<string>(), orders: new Set<string>(), lines: 0 },
    booking: { amount: 0, qty: 0, customers: new Set<string>(), orders: new Set<string>(), lines: 0 },
  };

  const addHot = (map: Map<string, HotMetric>, key: string, amount: number, qty: number, customer: string, order: string) => {
    const name = key || '未标记';
    let g = map.get(name);
    if (!g) {
      g = { amount: 0, qty: 0, customers: new Set<string>(), orders: new Set<string>(), lines: 0 };
      map.set(name, g);
    }
    g.amount += amount;
    g.qty += qty;
    g.customers.add(customer);
    g.orders.add(order);
    g.lines += 1;
  };

  for (let i = 0; i < data.rawRowCount; i++) {
    const amount = data.amt[i];
    const qty = data.qty[i];
    if (!Number.isFinite(amount) || !Number.isFinite(qty) || (amount === 0 && qty === 0)) continue;
    const customer = cellText(custCol[i]) ?? '未标记';
    const year = cellText(yearCol[i]) ?? '';
    const isBooking = year.includes('订货会');
    const key = orderKey(orderIds, i, customer);
    let c = customers.get(customer);
    if (!c) {
      c = { customer, normalAmount: 0, normalQty: 0, bookingAmount: 0, bookingQty: 0, orderKeys: new Set<string>() };
      customers.set(customer, c);
    }
    c.orderKeys.add(key);
    const bucket = isBooking ? summary.booking : summary.normal;
    bucket.amount += amount;
    bucket.qty += qty;
    bucket.customers.add(customer);
    bucket.orders.add(key);
    bucket.lines += 1;
    if (isBooking) {
      c.bookingAmount += amount;
      c.bookingQty += qty;
      addHot(productMap, cellText(productCol?.[i] ?? null) ?? '未标记', amount, qty, customer, key);
      addHot(categoryMap, cellText(catCol?.[i] ?? null) ?? '未分类', amount, qty, customer, key);
      addHot(brandMap, cellText(brandCol?.[i] ?? null) ?? '未标记', amount, qty, customer, key);
      addHot(designerMap, cellText(designerCol?.[i] ?? null) ?? '未标记', amount, qty, customer, key);
    } else {
      c.normalAmount += amount;
      c.normalQty += qty;
    }
  }

  const normalAmounts = [...customers.values()]
    .map((r) => r.normalAmount)
    .filter((v) => v > 0)
    .sort((a, b) => a - b);
  const highPotentialCutoff = percentile(normalAmounts, 0.75);
  const customerRows = [...customers.values()]
    .map((r) => {
      const totalAmount = r.normalAmount + r.bookingAmount;
      const totalQty = r.normalQty + r.bookingQty;
      const bookingAmountShare = totalAmount ? r.bookingAmount / totalAmount : 0;
      const bookingQtyShare = totalQty ? r.bookingQty / totalQty : 0;
      let type = '低参与客户';
      if (bookingAmountShare >= 0.6) type = '订货会型客户';
      else if (bookingAmountShare < 0.2 && r.normalAmount >= highPotentialCutoff && r.bookingAmount > 0) type = '高潜订货会客户';
      else if (bookingAmountShare < 0.2) type = '普通现货型客户';
      else if (r.normalAmount > 0 && r.bookingAmount > 0) type = '双轨客户';
      return {
        customer: r.customer,
        normal_amount: Math.round(r.normalAmount),
        normal_qty: Math.round(r.normalQty),
        booking_amount: Math.round(r.bookingAmount),
        booking_qty: Math.round(r.bookingQty),
        total_amount: Math.round(totalAmount),
        total_qty: Math.round(totalQty),
        booking_amount_share: round1(bookingAmountShare * 100),
        booking_qty_share: round1(bookingQtyShare * 100),
        orders: r.orderKeys.size,
        type,
      };
    })
    .filter((r) => hasPositiveNet(r.total_amount, r.total_qty))
    .sort((a, b) => b.total_amount - a.total_amount || b.total_qty - a.total_qty);

  const typeRows = [...customerRows.reduce((m, r) => {
    const g = m.get(r.type) ?? { type: r.type, customers: 0, amount: 0, qty: 0, booking_amount: 0, booking_qty: 0 };
    g.customers += 1;
    g.amount += r.total_amount;
    g.qty += r.total_qty;
    g.booking_amount += r.booking_amount;
    g.booking_qty += r.booking_qty;
    m.set(r.type, g);
    return m;
  }, new Map<string, { type: string; customers: number; amount: number; qty: number; booking_amount: number; booking_qty: number }>()).values()]
    .sort((a, b) => b.amount - a.amount);

  const orderTypeSummary = [
    { type: '普通款期成交', amount: Math.round(summary.normal.amount), qty: Math.round(summary.normal.qty), customers: summary.normal.customers.size, orders: summary.normal.orders.size, lines: summary.normal.lines },
    { type: '订货会成交', amount: Math.round(summary.booking.amount), qty: Math.round(summary.booking.qty), customers: summary.booking.customers.size, orders: summary.booking.orders.size, lines: summary.booking.lines },
  ];

  const hotRows = (map: Map<string, HotMetric>, keyName: string) => [...map.entries()]
    .map(([name, g]) => ({
      [keyName]: name,
      amount: Math.round(g.amount),
      qty: Math.round(g.qty),
      customers: g.customers.size,
      orders: g.orders.size,
      lines: g.lines,
      avg_qty_per_customer: g.customers.size ? round1(g.qty / g.customers.size) : 0,
    }))
    .filter((r) => hasPositiveNet(r.amount, r.qty))
    .sort((a, b) => b.qty - a.qty || b.amount - a.amount || b.customers - a.customers);

  return {
    customers: customerRows,
    customer_type_summary: typeRows,
    order_type_summary: orderTypeSummary,
    booking_products: hotRows(productMap, 'product'),
    booking_categories: hotRows(categoryMap, 'category'),
    booking_brands: hotRows(brandMap, 'brand'),
    booking_designers: hotRows(designerMap, 'designer'),
    high_potential_cutoff: Math.round(highPotentialCutoff),
  };
}

function addNumber(map: Map<string, number>, key: string, value: number) {
  map.set(key, (map.get(key) ?? 0) + value);
}

function metricRows(map: Map<string, number>) {
  return [...map.entries()]
    .sort((a, b) => b[1] - a[1])
    .filter(([, value]) => Math.round(value) > 0)
    .map(([name, value]) => ({ name, value: Math.round(value) }));
}

function colorHex(name: string): string {
  const s = name.toLowerCase();
  if (s.includes('图案') || s.includes('拼色')) return '#8e63b0';
  if (s.includes('未标记')) return '#bdbdbd';
  if (s.includes('黑')) return '#1f1f1f';
  if (s.includes('白')) return '#f7f2e8';
  if (s.includes('米') || s.includes('杏')) return '#eadcc8';
  if (s.includes('灰')) return '#9aa3b2';
  if (s.includes('蓝')) return '#4f83cc';
  if (s.includes('粉')) return '#f6a6c8';
  if (s.includes('绿')) return '#68a357';
  if (s.includes('紫')) return '#8e63b0';
  if (s.includes('黄') || s.includes('蛋')) return '#f2c94c';
  if (s.includes('红')) return '#d64545';
  if (s.includes('咖') || s.includes('棕')) return '#8a5a44';
  if (s.includes('卡其')) return '#b59b72';
  if (s.includes('橙')) return '#ef8a2f';
  return '#d9d9d9';
}

function normalizeColorName(value: Cell | null | undefined): string {
  const raw = cellText(value ?? null);
  if (!raw || raw.toLowerCase() === 'nan' || raw === '-' || raw === '无') return '未标记';
  const s = raw.replace(/\s+/g, '').replace(/[（）()【】\[\]]/g, '').toLowerCase();
  if (['其他', '其它', '其它色', '其他色', '杂色', '混色'].includes(s)) return '未标记';
  const colorTerms = ['黑', '白', '米', '杏', '灰', '蓝', '粉', '绿', '紫', '黄', '红', '咖', '棕', '卡其', '橙'];
  const explicitPattern = ['花色', '拼色', '撞色', '条纹', '格纹', '豹纹', '印花', '碎花', '波点', '千鸟格', '迷彩', '斑马', '渐变', '多色'];
  if (explicitPattern.some((k) => s.includes(k))) return '图案/拼色类';
  const colorHits = colorTerms.filter((k) => s.includes(k)).length;
  if ((/[\/+、,，]/.test(s) || /黑白|蓝白|红白|粉白|黄白|紫白|绿白/.test(s)) && colorHits >= 2) return '图案/拼色类';
  return raw;
}

function normalizeSizeName(value: Cell | null | undefined): string {
  const raw = cellText(value ?? null);
  if (!raw || raw.toLowerCase() === 'nan' || raw === '-' || raw === '无') return '未标记';
  return raw;
}

function clamp01(v: number): number {
  if (!Number.isFinite(v)) return 0;
  return Math.max(0, Math.min(1, v));
}

function monthKey(ms: number): string {
  if (!Number.isFinite(ms)) return '未标记月份';
  const d = new Date(ms);
  return `${d.getUTCFullYear()}-${String(d.getUTCMonth() + 1).padStart(2, '0')}`;
}

function colorStockTag(score: number, qtyScore: number, growthRate: number, comparable: boolean): string {
  if (score >= 65 && qtyScore >= 0.3) return '重点备货';
  if (comparable && growthRate >= 30 && score >= 35) return '趋势加量';
  if (score >= 45) return '常规备货';
  if (score >= 25) return '少量试水';
  return '控制备货';
}

function colorSalesGroup(
  qty: number,
  highQtyCutoff: number,
  midQtyCutoff: number,
  recentGrowthRate: number,
  recentGrowthQty: number,
  recentQty: number,
  hasRecentPeriodData: boolean,
): string {
  if (qty >= highQtyCutoff) return '稳定常备色';
  if (hasRecentPeriodData && qty >= midQtyCutoff && recentQty >= 100 && recentGrowthQty > 0 && recentGrowthRate >= 20) return '近期趋势色';
  return '控制备货色';
}

function colorSalesBand(qty: number, highQtyCutoff: number, midQtyCutoff: number): string {
  if (qty >= highQtyCutoff) return `高销量区：≥${Math.round(highQtyCutoff)}件`;
  if (qty >= midQtyCutoff) return `中销量区：${Math.round(midQtyCutoff)}-${Math.round(highQtyCutoff)}件`;
  return `低销量区：<${Math.round(midQtyCutoff)}件`;
}

function buildColorPreferenceAnalysis(data: PreferenceData) {
  const colorCol = data.cols.get('颜色');
  const catCol = data.cols.get('分类');
  const brandCol = data.cols.get('品牌');
  const custCol = data.cols.get('客户名称');
  if (!colorCol) {
    return {
      colors: [],
      category_colors: [],
      category_structure: [],
      top_color_category_distribution: [],
      top_color_brand_distribution: [],
      top_customer_colors: [],
      monthly_trend: [],
      months: [],
      trend_colors: [],
      has_prior_period_data: false,
    };
  }

  type MetricGroup = {
    color: string;
    category?: string;
    qty: number;
    amount: number;
    customers: Set<string>;
    categories: Map<string, number>;
    recentQty: number;
    priorQty: number;
    recent30Qty: number;
    previous30Qty: number;
    months: Map<string, number>;
  };

  type BreakdownMetric = {
    qty: number;
    amount: number;
    customers: Set<string>;
  };

  const newGroup = (color: string, category?: string): MetricGroup => ({
    color,
    category,
    qty: 0,
    amount: 0,
    customers: new Set<string>(),
    categories: new Map<string, number>(),
    recentQty: 0,
    priorQty: 0,
    recent30Qty: 0,
    previous30Qty: 0,
    months: new Map<string, number>(),
  });

  let maxMs = -Infinity;
  let minMs = Infinity;
  for (let i = 0; i < data.rawRowCount; i++) {
    const qty = data.qty[i], amount = data.amt[i];
    if (!Number.isFinite(qty) || !Number.isFinite(amount)) continue;
    const ms = data.orderMs?.[i];
    if (ms != null && Number.isFinite(ms)) {
      maxMs = Math.max(maxMs, ms);
      minMs = Math.min(minMs, ms);
    }
  }
  const dayMs = 24 * 60 * 60 * 1000;
  const hasDates = Number.isFinite(maxMs);
  const recentStart = hasDates ? maxMs - 90 * dayMs : Infinity;
  const priorStart = hasDates ? maxMs - 180 * dayMs : Infinity;
  const hasPriorPeriodData = hasDates && Number.isFinite(minMs) && minMs <= recentStart;
  const recent30Start = hasDates ? maxMs - 30 * dayMs : Infinity;
  const previous30Start = hasDates ? maxMs - 60 * dayMs : Infinity;
  const hasRecent30PeriodData = hasDates && Number.isFinite(minMs) && minMs <= recent30Start;

  const colors = new Map<string, MetricGroup>();
  const combos = new Map<string, MetricGroup>();
  const colorCategoryBreakdowns = new Map<string, Map<string, BreakdownMetric>>();
  const colorBrandBreakdowns = new Map<string, Map<string, BreakdownMetric>>();
  const categoryTotals = new Map<string, { qty: number; customers: Set<string> }>();
  const customerTotals = new Map<string, { qty: number; amount: number }>();
  const customerColors = new Map<string, Map<string, { qty: number; amount: number }>>();
  const allCustomers = new Set<string>();
  const allMonths = new Set<string>();

  const addToGroup = (g: MetricGroup, qty: number, amount: number, customer: string, category: string, ms: number | undefined) => {
    g.qty += qty;
    g.amount += amount;
    if (qty > 0) {
      g.customers.add(customer);
      g.categories.set(category, (g.categories.get(category) ?? 0) + qty);
    }
    if (hasDates && ms != null && Number.isFinite(ms)) {
      const mk = monthKey(ms);
      g.months.set(mk, (g.months.get(mk) ?? 0) + qty);
      if (ms > recentStart && ms <= maxMs) g.recentQty += qty;
      else if (ms > priorStart && ms <= recentStart) g.priorQty += qty;
      if (ms > recent30Start && ms <= maxMs) g.recent30Qty += qty;
      else if (ms > previous30Start && ms <= recent30Start) g.previous30Qty += qty;
    }
  };

  const addBreakdown = (map: Map<string, Map<string, BreakdownMetric>>, color: string, name: string, qty: number, amount: number, customer: string) => {
    let colorMap = map.get(color);
    if (!colorMap) {
      colorMap = new Map<string, BreakdownMetric>();
      map.set(color, colorMap);
    }
    const row = colorMap.get(name) ?? { qty: 0, amount: 0, customers: new Set<string>() };
    row.qty += qty;
    row.amount += amount;
    if (qty > 0) row.customers.add(customer);
    colorMap.set(name, row);
  };

  for (let i = 0; i < data.rawRowCount; i++) {
    const qty = data.qty[i], amount = data.amt[i];
    if (!Number.isFinite(qty) || !Number.isFinite(amount)) continue;
    const customer = custCol ? cellText(custCol[i]) ?? '未填写客户' : '未填写客户';
    const category = catCol ? cellText(catCol[i]) ?? '未分类' : '未分类';
    const brand = brandCol ? cellText(brandCol[i]) ?? '未标记品牌' : '未标记品牌';
    const color = normalizeColorName(colorCol[i]);
    const ms = data.orderMs?.[i];
    if (qty > 0) allCustomers.add(customer);
    if (hasDates && ms != null && Number.isFinite(ms)) allMonths.add(monthKey(ms));

    let customerTotal = customerTotals.get(customer);
    if (!customerTotal) {
      customerTotal = { qty: 0, amount: 0 };
      customerTotals.set(customer, customerTotal);
    }
    customerTotal.qty += qty;
    customerTotal.amount += amount;

    let colorMap = customerColors.get(customer);
    if (!colorMap) {
      colorMap = new Map<string, { qty: number; amount: number }>();
      customerColors.set(customer, colorMap);
    }
    const customerColor = colorMap.get(color) ?? { qty: 0, amount: 0 };
    customerColor.qty += qty;
    customerColor.amount += amount;
    colorMap.set(color, customerColor);

    let cg = colors.get(color);
    if (!cg) {
      cg = newGroup(color);
      colors.set(color, cg);
    }
    addToGroup(cg, qty, amount, customer, category, ms);

    const comboKey = `${category}\u0000${color}`;
    let combo = combos.get(comboKey);
    if (!combo) {
      combo = newGroup(color, category);
      combos.set(comboKey, combo);
    }
    addToGroup(combo, qty, amount, customer, category, ms);
    addBreakdown(colorCategoryBreakdowns, color, category, qty, amount, customer);
    addBreakdown(colorBrandBreakdowns, color, brand, qty, amount, customer);

    let ct = categoryTotals.get(category);
    if (!ct) {
      ct = { qty: 0, customers: new Set<string>() };
      categoryTotals.set(category, ct);
    }
    ct.qty += qty;
    if (qty > 0) ct.customers.add(customer);
  }

  const months = [...allMonths].sort();
  const stabilityOf = (g: MetricGroup) => {
    if (months.length <= 1 || g.qty <= 0) return 0.5;
    const values = months.map((m) => g.months.get(m) ?? 0);
    const mean = values.reduce((s, v) => s + v, 0) / values.length;
    if (mean <= 0) return 0;
    const sd = Math.sqrt(values.reduce((s, v) => s + (v - mean) ** 2, 0) / values.length);
    return clamp01(1 - sd / mean);
  };
  const growthRateOf = (g: MetricGroup) => {
    if (!hasPriorPeriodData) return 0;
    if (g.priorQty > 0) return ((g.recentQty - g.priorQty) / Math.abs(g.priorQty)) * 100;
    return g.recentQty > 0 ? 100 : 0;
  };
  const growthScoreOf = (g: MetricGroup) => {
    if (!hasPriorPeriodData) return 0.5;
    return clamp01((growthRateOf(g) + 100) / 200);
  };
  const recentGrowthRateOf = (g: MetricGroup) => {
    if (!hasRecent30PeriodData) return 0;
    if (g.previous30Qty > 0) return ((g.recent30Qty - g.previous30Qty) / Math.abs(g.previous30Qty)) * 100;
    return g.recent30Qty > 0 ? 100 : 0;
  };

  const positiveColors = [...colors.values()].filter((g) => hasPositiveNet(g.amount, g.qty));
  const maxColorQty = Math.max(1, ...positiveColors.map((g) => g.qty));
  const totalColorQty = positiveColors.reduce((s, g) => s + g.qty, 0);
  const totalCustomers = Math.max(1, allCustomers.size);
  const colorQtyValues = positiveColors.map((g) => g.qty).sort((a, b) => a - b);
  const highQtyCutoff = Math.max(percentile(colorQtyValues, 0.98), maxColorQty * 0.08);
  const midQtyCutoff = Math.max(percentile(colorQtyValues, 0.95), maxColorQty * 0.03);

  const colorRows = positiveColors.map((g) => {
    const qtyScore = clamp01(g.qty / maxColorQty);
    const coverageScore = clamp01(g.customers.size / totalCustomers);
    const growthRate = growthRateOf(g);
    const recentGrowthRate = recentGrowthRateOf(g);
    const recentGrowthQty = g.recent30Qty - g.previous30Qty;
    const stability = stabilityOf(g);
    const score = round1((qtyScore * 0.4 + coverageScore * 0.25 + growthScoreOf(g) * 0.2 + stability * 0.15) * 100);
    const tag = colorStockTag(score, qtyScore, growthRate, hasPriorPeriodData);
    const group = colorSalesGroup(g.qty, highQtyCutoff, midQtyCutoff, recentGrowthRate, recentGrowthQty, g.recent30Qty, hasRecent30PeriodData);
    return {
      color: g.color,
      hex: colorHex(g.color),
      qty: Math.round(g.qty),
      amount: Math.round(g.amount),
      customers: g.customers.size,
      customer_share: g.customers.size / totalCustomers,
      qty_share: totalColorQty ? g.qty / totalColorQty : 0,
      coverage_categories: [...g.categories.values()].filter((v) => v > 0).length,
      recent_qty: Math.round(g.recentQty),
      prior_qty: Math.round(g.priorQty),
      recent30_qty: Math.round(g.recent30Qty),
      previous30_qty: Math.round(g.previous30Qty),
      growth_rate: round1(growthRate),
      recent_growth_rate: round1(recentGrowthRate),
      recent_growth_qty: Math.round(recentGrowthQty),
      stability: round1(stability * 100),
      score,
      tag,
      group,
      qty_band: colorSalesBand(g.qty, highQtyCutoff, midQtyCutoff),
    };
  }).sort((a, b) => b.score - a.score || b.qty - a.qty);

  const positiveCombos = [...combos.values()].filter((g) => g.category && hasPositiveNet(g.amount, g.qty));
  const maxComboQty = Math.max(1, ...positiveCombos.map((g) => g.qty));
  const comboRows = positiveCombos.map((g) => {
    const categoryTotal = categoryTotals.get(g.category!) ?? { qty: 0, customers: new Set<string>() };
    const qtyScore = clamp01(g.qty / maxComboQty);
    const coverageDenom = Math.max(1, categoryTotal.customers.size);
    const coverageScore = clamp01(g.customers.size / coverageDenom);
    const growthRate = growthRateOf(g);
    const stability = stabilityOf(g);
    const score = round1((qtyScore * 0.4 + coverageScore * 0.25 + growthScoreOf(g) * 0.2 + stability * 0.15) * 100);
    const tag = colorStockTag(score, qtyScore, growthRate, hasPriorPeriodData);
    return {
      category: g.category!,
      color: g.color,
      hex: colorHex(g.color),
      label: `${g.category} / ${g.color}`,
      qty: Math.round(g.qty),
      amount: Math.round(g.amount),
      customers: g.customers.size,
      category_customer_share: g.customers.size / coverageDenom,
      category_qty_share: categoryTotal.qty ? g.qty / categoryTotal.qty : 0,
      recent_qty: Math.round(g.recentQty),
      prior_qty: Math.round(g.priorQty),
      growth_rate: round1(growthRate),
      stability: round1(stability * 100),
      score,
      tag,
    };
  }).sort((a, b) => b.qty - a.qty);

  const categoryStructure = [...categoryTotals.entries()]
    .sort((a, b) => b[1].qty - a[1].qty)
    .map(([category, total]) => {
      const rows = comboRows.filter((r) => r.category === category).sort((a, b) => b.qty - a.qty);
      const top = rows.slice(0, 1);
      return {
        category,
        total_qty: Math.round(total.qty),
        colors: top.map((r, i) => ({ rank: i + 1, color: r.color, qty: r.qty, share: total.qty ? r.qty / total.qty : 0, hex: r.hex })),
      };
    });

  const trendColors = [...colorRows].sort((a, b) => b.qty - a.qty).slice(0, 8).map((r) => r.color);
  const monthlyTrend = months.flatMap((m) => trendColors.map((color) => {
    const g = colors.get(color);
    return { month: m, color, qty: Math.round(g?.months.get(m) ?? 0) };
  }));

  const topCustomerColors = [...customerTotals.entries()]
    .filter(([, total]) => total.amount > 0)
    .sort((a, b) => b[1].amount - a[1].amount)
    .slice(0, 20)
    .map(([customer]) => customer);
  const topCustomerColorTotals = new Map<string, { qty: number; amount: number; customers: Set<string> }>();
  let topCustomerQty = 0;
  let topCustomerAmount = 0;
  for (const customer of topCustomerColors) {
    const customerTotal = customerTotals.get(customer);
    if (customerTotal) {
      topCustomerQty += customerTotal.qty;
      topCustomerAmount += customerTotal.amount;
    }
    for (const [color, metric] of customerColors.get(customer) ?? new Map()) {
      if (!hasPositiveNet(metric.amount, metric.qty)) continue;
      const row = topCustomerColorTotals.get(color) ?? { qty: 0, amount: 0, customers: new Set<string>() };
      row.qty += metric.qty;
      row.amount += metric.amount;
      row.customers.add(customer);
      topCustomerColorTotals.set(color, row);
    }
  }
  const topCustomerColorRows = [...topCustomerColorTotals.entries()]
    .filter(([, metric]) => hasPositiveNet(metric.amount, metric.qty))
    .sort((a, b) => b[1].qty - a[1].qty || b[1].amount - a[1].amount)
    .slice(0, 10)
    .map(([color, metric], index) => ({
      rank: index + 1,
      color,
      hex: colorHex(color),
      qty: Math.round(metric.qty),
      amount: Math.round(metric.amount),
      customers: metric.customers.size,
      qty_share: topCustomerQty ? metric.qty / topCustomerQty : 0,
      amount_share: topCustomerAmount ? metric.amount / topCustomerAmount : 0,
    }));

  const topQtyColors = [...colorRows].sort((a, b) => b.qty - a.qty || b.amount - a.amount).slice(0, 10);
  const buildTopColorBreakdown = (source: Map<string, Map<string, BreakdownMetric>>) => topQtyColors.map((colorRow, index) => {
    const positiveItems = [...(source.get(colorRow.color)?.entries() ?? [])]
      .map(([name, metric]) => ({
        name,
        qty: Math.round(metric.qty),
        amount: Math.round(metric.amount),
        customers: metric.customers.size,
      }))
      .filter((r) => hasPositiveNet(r.amount, r.qty))
      .sort((a, b) => b.qty - a.qty || b.amount - a.amount);
    const totalQty = positiveItems.reduce((sum, item) => sum + item.qty, 0) || Math.max(0, Number(colorRow.qty) || 0);
    const totalAmount = positiveItems.reduce((sum, item) => sum + item.amount, 0) || Math.max(0, Number(colorRow.amount) || 0);
    const rawItems = positiveItems.map((item) => ({
      ...item,
      share: totalQty ? item.qty / totalQty : 0,
    }));
    return {
      rank: index + 1,
      color: colorRow.color,
      hex: colorRow.hex,
      total_qty: totalQty,
      total_amount: totalAmount,
      customers: colorRow.customers,
      items: rawItems,
    };
  });

  return {
    recent_from: hasDates ? previewDateText(recentStart + dayMs) : '',
    recent_to: hasDates ? previewDateText(maxMs) : '',
    prior_from: hasDates ? previewDateText(priorStart + dayMs) : '',
    prior_to: hasDates ? previewDateText(recentStart) : '',
    recent30_from: hasDates ? previewDateText(recent30Start + dayMs) : '',
    recent30_to: hasDates ? previewDateText(maxMs) : '',
    previous30_from: hasDates ? previewDateText(previous30Start + dayMs) : '',
    previous30_to: hasDates ? previewDateText(recent30Start) : '',
    has_prior_period_data: hasPriorPeriodData,
    has_recent30_period_data: hasRecent30PeriodData,
    color_cutoffs: {
      high_qty: Math.round(highQtyCutoff),
      mid_qty: Math.round(midQtyCutoff),
      trend_growth_rate: 20,
    },
    colors: colorRows,
    category_colors: comboRows,
    category_structure: categoryStructure,
    top_color_category_distribution: buildTopColorBreakdown(colorCategoryBreakdowns),
    top_color_brand_distribution: buildTopColorBreakdown(colorBrandBreakdowns),
    top_customer_colors: topCustomerColorRows,
    months,
    trend_colors: trendColors,
    monthly_trend: monthlyTrend,
  };
}

function sizeSalesGroup(
  qty: number,
  highQtyCutoff: number,
  midQtyCutoff: number,
  recentGrowthRate: number,
  recentGrowthQty: number,
  recentQty: number,
  hasRecentPeriodData: boolean,
): string {
  if (qty >= highQtyCutoff) return '稳定常备码';
  if (hasRecentPeriodData && qty >= midQtyCutoff && recentQty >= 100 && recentGrowthQty > 0 && recentGrowthRate >= 20) return '近期趋势码';
  return '控制备货码';
}

function sizeSalesBand(qty: number, highQtyCutoff: number, midQtyCutoff: number): string {
  if (qty >= highQtyCutoff) return `高销量区：≥${Math.round(highQtyCutoff)}件`;
  if (qty >= midQtyCutoff) return `中销量区：${Math.round(midQtyCutoff)}-${Math.round(highQtyCutoff)}件`;
  return `低销量区：<${Math.round(midQtyCutoff)}件`;
}

function buildSizePreferenceAnalysis(data: PreferenceData) {
  const sizeCol = data.cols.get('尺码');
  const catCol = data.cols.get('分类');
  const custCol = data.cols.get('客户名称');
  if (!sizeCol) {
    return { sizes: [], category_sizes: [], category_structure: [], monthly_trend: [], months: [], trend_sizes: [], has_prior_period_data: false };
  }

  type SizeMetricGroup = {
    size: string;
    category?: string;
    qty: number;
    amount: number;
    customers: Set<string>;
    categories: Map<string, number>;
    recentQty: number;
    priorQty: number;
    recent30Qty: number;
    previous30Qty: number;
    months: Map<string, number>;
  };

  const newGroup = (size: string, category?: string): SizeMetricGroup => ({
    size,
    category,
    qty: 0,
    amount: 0,
    customers: new Set<string>(),
    categories: new Map<string, number>(),
    recentQty: 0,
    priorQty: 0,
    recent30Qty: 0,
    previous30Qty: 0,
    months: new Map<string, number>(),
  });

  let maxMs = -Infinity;
  let minMs = Infinity;
  for (let i = 0; i < data.rawRowCount; i++) {
    const qty = data.qty[i], amount = data.amt[i];
    if (!Number.isFinite(qty) || !Number.isFinite(amount)) continue;
    const ms = data.orderMs?.[i];
    if (ms != null && Number.isFinite(ms)) {
      maxMs = Math.max(maxMs, ms);
      minMs = Math.min(minMs, ms);
    }
  }
  const dayMs = 24 * 60 * 60 * 1000;
  const hasDates = Number.isFinite(maxMs);
  const recentStart = hasDates ? maxMs - 90 * dayMs : Infinity;
  const priorStart = hasDates ? maxMs - 180 * dayMs : Infinity;
  const hasPriorPeriodData = hasDates && Number.isFinite(minMs) && minMs <= recentStart;
  const recent30Start = hasDates ? maxMs - 30 * dayMs : Infinity;
  const previous30Start = hasDates ? maxMs - 60 * dayMs : Infinity;
  const hasRecent30PeriodData = hasDates && Number.isFinite(minMs) && minMs <= recent30Start;

  const sizes = new Map<string, SizeMetricGroup>();
  const combos = new Map<string, SizeMetricGroup>();
  const categoryTotals = new Map<string, { qty: number; amount: number; customers: Set<string> }>();
  const customerTotals = new Map<string, { qty: number; amount: number }>();
  const customerSizes = new Map<string, Map<string, { qty: number; amount: number }>>();
  const allCustomers = new Set<string>();
  const allMonths = new Set<string>();

  const addToGroup = (g: SizeMetricGroup, qty: number, amount: number, customer: string, category: string, ms: number | undefined) => {
    g.qty += qty;
    g.amount += amount;
    if (qty > 0) {
      g.customers.add(customer);
      g.categories.set(category, (g.categories.get(category) ?? 0) + qty);
    }
    if (hasDates && ms != null && Number.isFinite(ms)) {
      const mk = monthKey(ms);
      g.months.set(mk, (g.months.get(mk) ?? 0) + qty);
      if (ms > recentStart && ms <= maxMs) g.recentQty += qty;
      else if (ms > priorStart && ms <= recentStart) g.priorQty += qty;
      if (ms > recent30Start && ms <= maxMs) g.recent30Qty += qty;
      else if (ms > previous30Start && ms <= recent30Start) g.previous30Qty += qty;
    }
  };

  for (let i = 0; i < data.rawRowCount; i++) {
    const qty = data.qty[i], amount = data.amt[i];
    if (!Number.isFinite(qty) || !Number.isFinite(amount)) continue;
    const customer = custCol ? cellText(custCol[i]) ?? '未填写客户' : '未填写客户';
    const category = catCol ? cellText(catCol[i]) ?? '未分类' : '未分类';
    const size = normalizeSizeName(sizeCol[i]);
    const ms = data.orderMs?.[i];
    if (qty > 0) allCustomers.add(customer);
    if (hasDates && ms != null && Number.isFinite(ms)) allMonths.add(monthKey(ms));

    let customerTotal = customerTotals.get(customer);
    if (!customerTotal) {
      customerTotal = { qty: 0, amount: 0 };
      customerTotals.set(customer, customerTotal);
    }
    customerTotal.qty += qty;
    customerTotal.amount += amount;

    let sizeMap = customerSizes.get(customer);
    if (!sizeMap) {
      sizeMap = new Map<string, { qty: number; amount: number }>();
      customerSizes.set(customer, sizeMap);
    }
    const customerSize = sizeMap.get(size) ?? { qty: 0, amount: 0 };
    customerSize.qty += qty;
    customerSize.amount += amount;
    sizeMap.set(size, customerSize);

    let sg = sizes.get(size);
    if (!sg) {
      sg = newGroup(size);
      sizes.set(size, sg);
    }
    addToGroup(sg, qty, amount, customer, category, ms);

    const comboKey = `${category}\u0000${size}`;
    let combo = combos.get(comboKey);
    if (!combo) {
      combo = newGroup(size, category);
      combos.set(comboKey, combo);
    }
    addToGroup(combo, qty, amount, customer, category, ms);

    let ct = categoryTotals.get(category);
    if (!ct) {
      ct = { qty: 0, amount: 0, customers: new Set<string>() };
      categoryTotals.set(category, ct);
    }
    ct.qty += qty;
    ct.amount += amount;
    if (qty > 0) ct.customers.add(customer);
  }

  const months = [...allMonths].sort();
  const stabilityOf = (g: SizeMetricGroup) => {
    if (months.length <= 1 || g.qty <= 0) return 0.5;
    const values = months.map((m) => g.months.get(m) ?? 0);
    const mean = values.reduce((s, v) => s + v, 0) / values.length;
    if (mean <= 0) return 0;
    const sd = Math.sqrt(values.reduce((s, v) => s + (v - mean) ** 2, 0) / values.length);
    return clamp01(1 - sd / mean);
  };
  const growthRateOf = (g: SizeMetricGroup) => {
    if (!hasPriorPeriodData) return 0;
    if (g.priorQty > 0) return ((g.recentQty - g.priorQty) / Math.abs(g.priorQty)) * 100;
    return g.recentQty > 0 ? 100 : 0;
  };
  const growthScoreOf = (g: SizeMetricGroup) => {
    if (!hasPriorPeriodData) return 0.5;
    return clamp01((growthRateOf(g) + 100) / 200);
  };
  const recentGrowthRateOf = (g: SizeMetricGroup) => {
    if (!hasRecent30PeriodData) return 0;
    if (g.previous30Qty > 0) return ((g.recent30Qty - g.previous30Qty) / Math.abs(g.previous30Qty)) * 100;
    return g.recent30Qty > 0 ? 100 : 0;
  };

  const positiveSizes = [...sizes.values()].filter((g) => hasPositiveNet(g.amount, g.qty));
  const maxSizeQty = Math.max(1, ...positiveSizes.map((g) => g.qty));
  const totalSizeQty = positiveSizes.reduce((s, g) => s + g.qty, 0);
  const totalCustomers = Math.max(1, allCustomers.size);
  const sizeQtyValues = positiveSizes.map((g) => g.qty).sort((a, b) => a - b);
  const highQtyCutoff = Math.max(percentile(sizeQtyValues, 0.75), maxSizeQty * 0.25);
  const midQtyCutoff = Math.max(percentile(sizeQtyValues, 0.5), maxSizeQty * 0.08);

  const sizeRows = positiveSizes.map((g) => {
    const qtyScore = clamp01(g.qty / maxSizeQty);
    const coverageScore = clamp01(g.customers.size / totalCustomers);
    const growthRate = growthRateOf(g);
    const recentGrowthRate = recentGrowthRateOf(g);
    const recentGrowthQty = g.recent30Qty - g.previous30Qty;
    const stability = stabilityOf(g);
    const score = round1((qtyScore * 0.4 + coverageScore * 0.25 + growthScoreOf(g) * 0.2 + stability * 0.15) * 100);
    const tag = colorStockTag(score, qtyScore, growthRate, hasPriorPeriodData);
    const group = sizeSalesGroup(g.qty, highQtyCutoff, midQtyCutoff, recentGrowthRate, recentGrowthQty, g.recent30Qty, hasRecent30PeriodData);
    return {
      size: g.size,
      qty: Math.round(g.qty),
      amount: Math.round(g.amount),
      customers: g.customers.size,
      customer_share: g.customers.size / totalCustomers,
      qty_share: totalSizeQty ? g.qty / totalSizeQty : 0,
      coverage_categories: [...g.categories.values()].filter((v) => v > 0).length,
      recent_qty: Math.round(g.recentQty),
      prior_qty: Math.round(g.priorQty),
      recent30_qty: Math.round(g.recent30Qty),
      previous30_qty: Math.round(g.previous30Qty),
      growth_rate: round1(growthRate),
      recent_growth_rate: round1(recentGrowthRate),
      recent_growth_qty: Math.round(recentGrowthQty),
      stability: round1(stability * 100),
      score,
      tag,
      group,
      qty_band: sizeSalesBand(g.qty, highQtyCutoff, midQtyCutoff),
    };
  }).sort((a, b) => b.qty - a.qty);

  const allSizeNames = sizeRows.map((r) => r.size);
  const positiveCombos = [...combos.values()].filter((g) => g.category && hasPositiveNet(g.amount, g.qty));
  const maxComboQty = Math.max(1, ...positiveCombos.map((g) => g.qty));
  const comboRows = positiveCombos.map((g) => {
    const categoryTotal = categoryTotals.get(g.category!) ?? { qty: 0, amount: 0, customers: new Set<string>() };
    const qtyScore = clamp01(g.qty / maxComboQty);
    const coverageDenom = Math.max(1, categoryTotal.customers.size);
    const coverageScore = clamp01(g.customers.size / coverageDenom);
    const growthRate = growthRateOf(g);
    const stability = stabilityOf(g);
    const score = round1((qtyScore * 0.4 + coverageScore * 0.25 + growthScoreOf(g) * 0.2 + stability * 0.15) * 100);
    const tag = colorStockTag(score, qtyScore, growthRate, hasPriorPeriodData);
    return {
      category: g.category!,
      size: g.size,
      label: `${g.category} / ${g.size}`,
      qty: Math.round(g.qty),
      amount: Math.round(g.amount),
      customers: g.customers.size,
      category_customer_share: g.customers.size / coverageDenom,
      category_qty_share: categoryTotal.qty ? g.qty / categoryTotal.qty : 0,
      recent_qty: Math.round(g.recentQty),
      prior_qty: Math.round(g.priorQty),
      growth_rate: round1(growthRate),
      stability: round1(stability * 100),
      score,
      tag,
    };
  }).sort((a, b) => b.qty - a.qty);
  const comboMap = new Map(comboRows.map((r) => [`${r.category}\u0000${r.size}`, r]));

  const categoryStructure = [...categoryTotals.entries()]
    .sort((a, b) => b[1].qty - a[1].qty)
    .map(([category, total]) => ({
      category,
      total_qty: Math.round(total.qty),
      total_amount: Math.round(total.amount),
      sizes: allSizeNames.map((size) => {
        const row = comboMap.get(`${category}\u0000${size}`);
        const qty = row?.qty ?? 0;
        return {
          size,
          qty,
          amount: row?.amount ?? 0,
          share: total.qty ? qty / total.qty : 0,
        };
      }),
    }));

  const trendSizes = [...sizeRows].filter((r) => r.qty > 0).sort((a, b) => b.qty - a.qty).map((r) => r.size);
  const monthlyTrend = months.flatMap((m) => trendSizes.map((size) => {
    const g = sizes.get(size);
    return { month: m, size, qty: Math.round(g?.months.get(m) ?? 0) };
  }));

  const topCustomerByAmount = [...customerTotals.entries()]
    .filter(([, total]) => total.amount > 0)
    .sort((a, b) => b[1].amount - a[1].amount || b[1].qty - a[1].qty)
    .slice(0, 20)
    .map(([customer]) => customer);
  const topCustomerSizeTotals = new Map<string, { qty: number; amount: number; customers: Set<string> }>();
  let topCustomerQty = 0;
  let topCustomerAmount = 0;
  for (const customer of topCustomerByAmount) {
    const customerTotal = customerTotals.get(customer);
    if (customerTotal) {
      topCustomerQty += customerTotal.qty;
      topCustomerAmount += customerTotal.amount;
    }
    for (const [size, metric] of customerSizes.get(customer) ?? new Map()) {
      if (!hasPositiveNet(metric.amount, metric.qty)) continue;
      const row = topCustomerSizeTotals.get(size) ?? { qty: 0, amount: 0, customers: new Set<string>() };
      row.qty += metric.qty;
      row.amount += metric.amount;
      row.customers.add(customer);
      topCustomerSizeTotals.set(size, row);
    }
  }
  const topCustomerSizeRows = [...topCustomerSizeTotals.entries()]
    .filter(([, metric]) => hasPositiveNet(metric.amount, metric.qty))
    .sort((a, b) => b[1].qty - a[1].qty || b[1].amount - a[1].amount)
    .map(([size, metric], index) => ({
      rank: index + 1,
      size,
      qty: Math.round(metric.qty),
      amount: Math.round(metric.amount),
      customers: metric.customers.size,
      qty_share: topCustomerQty ? metric.qty / topCustomerQty : 0,
      amount_share: topCustomerAmount ? metric.amount / topCustomerAmount : 0,
    }));

  return {
    recent_from: hasDates ? previewDateText(recentStart + dayMs) : '',
    recent_to: hasDates ? previewDateText(maxMs) : '',
    prior_from: hasDates ? previewDateText(priorStart + dayMs) : '',
    prior_to: hasDates ? previewDateText(recentStart) : '',
    recent30_from: hasDates ? previewDateText(recent30Start + dayMs) : '',
    recent30_to: hasDates ? previewDateText(maxMs) : '',
    previous30_from: hasDates ? previewDateText(previous30Start + dayMs) : '',
    previous30_to: hasDates ? previewDateText(recent30Start) : '',
    has_prior_period_data: hasPriorPeriodData,
    has_recent30_period_data: hasRecent30PeriodData,
    size_cutoffs: {
      high_qty: Math.round(highQtyCutoff),
      mid_qty: Math.round(midQtyCutoff),
      trend_growth_rate: 20,
    },
    sizes: sizeRows,
    category_sizes: comboRows,
    category_structure: categoryStructure,
    top_customer_sizes: topCustomerSizeRows,
    months,
    trend_sizes: trendSizes,
    monthly_trend: monthlyTrend,
  };
}

function buildCustomerVisualProfiles(data: PreferenceData, orderIds: (Cell | null)[], segmentation: PriceSegmentation) {
  const custCol = data.cols.get('客户名称');
  if (!custCol) return { customers: [], profiles: [] };
  const catCol = data.cols.get('分类');
  const brandCol = data.cols.get('品牌');
  const designerCol = data.cols.get('设计师品牌');
  const colorCol = data.cols.get('颜色');
  const sizeCol = data.cols.get('尺码');
  const yearCol = data.cols.get('年份');
  const productCol = data.cols.get('货号');

  type CategoryPriceGroup = {
    amount: number;
    qty: number;
    priceAmount: number;
    priceQty: number;
    prices: { price: number; qty: number }[];
    bandAmounts: Map<string, number>;
    bandQty: Map<string, number>;
    orderKeys: Set<string>;
  };

  type ColorBreakdownMetric = {
    amount: number;
    qty: number;
  };

  type ColorBreakdownGroup = {
    amount: number;
    qty: number;
    categories: Map<string, ColorBreakdownMetric>;
    brands: Map<string, ColorBreakdownMetric>;
  };

  type ProfileGroup = {
    key: string;
    customer: string;
    brand: string | null;
    amount: number;
    qty: number;
    priceAmount: number;
    priceQty: number;
    orderKeys: Set<string>;
    categories: Map<string, number>;
    brands: Map<string, number>;
    designers: Map<string, number>;
    colors: Map<string, number>;
    sizes: Map<string, number>;
    priceBands: Map<string, number>;
    priceSegments: Map<string, number>;
    seasons: Map<string, number>;
    priceValues: { price: number; qty: number }[];
    categoryPrices: Map<string, CategoryPriceGroup>;
    colorDetails: Map<string, ColorBreakdownGroup>;
  };

  type RepeatOrderGroup = {
    groupKey: string;
    customer: string;
    product: string;
    category: string;
    orderKey: string;
    amount: number;
    qty: number;
    ms: number;
    lineIndex: number;
  };

  const groups = new Map<string, ProfileGroup>();
  const productOrders = new Map<string, RepeatOrderGroup>();
  const profileGroupKey = (customer: string, brand: string | null) => (brand ? `${customer}\u0000${brand}` : customer);
  const getGroup = (customer: string, brand: string | null = null) => {
    const key = profileGroupKey(customer, brand);
    let g = groups.get(key);
    if (!g) {
      g = {
        key,
        customer,
        brand,
        amount: 0,
        qty: 0,
        priceAmount: 0,
        priceQty: 0,
        orderKeys: new Set(),
        categories: new Map(),
        brands: new Map(),
        designers: new Map(),
        colors: new Map(),
        sizes: new Map(),
        priceBands: new Map(),
        priceSegments: new Map(),
        seasons: new Map(),
        priceValues: [],
        categoryPrices: new Map(),
        colorDetails: new Map(),
      };
      groups.set(key, g);
    }
    return g;
  };

  const addColorBreakdown = (
    g: ProfileGroup,
    color: string,
    category: string | null,
    brand: string | null,
    amount: number,
    qty: number,
  ) => {
    let cd = g.colorDetails.get(color);
    if (!cd) {
      cd = { amount: 0, qty: 0, categories: new Map(), brands: new Map() };
      g.colorDetails.set(color, cd);
    }
    cd.amount += amount;
    cd.qty += qty;
    const addTo = (map: Map<string, ColorBreakdownMetric>, name: string) => {
      const row = map.get(name) ?? { amount: 0, qty: 0 };
      row.amount += amount;
      row.qty += qty;
      map.set(name, row);
    };
    addTo(cd.categories, category || '未分类');
    addTo(cd.brands, brand || '未标记品牌');
  };

  for (let i = 0; i < data.rawRowCount; i++) {
    const amount = data.amt[i];
    const qty = data.qty[i];
    if (!Number.isFinite(amount) || !Number.isFinite(qty) || (amount === 0 && qty === 0)) continue;
    const customer = cellText(custCol[i]);
    if (!customer) continue;
    const price = amount > 0 && qty > 0 ? amount / qty : NaN;
    const hasEffectivePrice = Number.isFinite(price) && price > 0;
    const key = orderKey(orderIds, i, customer);
    const year = yearCol ? cellText(yearCol[i]) : null;
    const category = catCol ? cellText(catCol[i]) : null;
    const product = productCol ? cellText(productCol[i]) ?? '未标记' : '未标记';
    const brand = brandCol ? cellText(brandCol[i]) : null;
    const designer = designerCol ? cellText(designerCol[i]) : null;
    const color = colorCol ? normalizeColorName(colorCol[i]) : null;
    const size = sizeCol ? cellText(sizeCol[i]) : null;

    const targetGroups = [getGroup(customer)];
    if (brand) targetGroups.push(getGroup(customer, brand));

    for (const g of targetGroups) {
      g.amount += amount;
      g.qty += qty;
      g.orderKeys.add(key);
      if (hasEffectivePrice) {
        g.priceAmount += amount;
        g.priceQty += qty;
        g.priceValues.push({ price, qty });
        addNumber(g.priceBands, priceBand(price), amount);
        addNumber(g.priceSegments, acceptanceSegment(price, segmentation), amount);
      }
      addNumber(g.seasons, year || '未填写', amount);
      if (category) {
        addNumber(g.categories, category, amount);
        let cp = g.categoryPrices.get(category);
        if (!cp) {
          cp = {
            amount: 0,
            qty: 0,
            priceAmount: 0,
            priceQty: 0,
            prices: [],
            bandAmounts: new Map(),
            bandQty: new Map(),
            orderKeys: new Set(),
          };
          g.categoryPrices.set(category, cp);
        }
        cp.amount += amount;
        cp.qty += qty;
        if (hasEffectivePrice) {
          const band = priceBand(price);
          cp.priceAmount += amount;
          cp.priceQty += qty;
          cp.prices.push({ price, qty });
          addNumber(cp.bandAmounts, band, amount);
          addNumber(cp.bandQty, band, qty);
        }
        cp.orderKeys.add(key);
      }
      if (brand) addNumber(g.brands, brand, amount);
      if (designer) addNumber(g.designers, designer, amount);
      if (color) {
        addNumber(g.colors, color, qty);
        if (isPositiveSale(amount, qty)) addColorBreakdown(g, color, category, brand, amount, qty);
      }
      if (size) addNumber(g.sizes, size, qty);

      if (isPositiveSale(amount, qty)) {
        const productOrderKey = `${g.key}\u0000${product}\u0000${key}`;
        let po = productOrders.get(productOrderKey);
        if (!po) {
          po = {
            groupKey: g.key,
            customer,
            product,
            category: category || '未分类',
            orderKey: key,
            amount: 0,
            qty: 0,
            ms: Number.POSITIVE_INFINITY,
            lineIndex: i,
          };
          productOrders.set(productOrderKey, po);
        }
        po.amount += amount;
        po.qty += qty;
        const ms = data.orderMs?.[i];
        if (ms != null && Number.isFinite(ms)) po.ms = Math.min(po.ms, ms);
        po.lineIndex = Math.min(po.lineIndex, i);
      }
    }
  }

  const repeatByGroup = new Map<string, {
    products: Map<string, {
      product: string;
      category: string;
      orders: RepeatOrderGroup[];
    }>;
  }>();
  for (const order of productOrders.values()) {
    if (!hasPositiveNet(order.amount, order.qty)) continue;
    let groupRepeat = repeatByGroup.get(order.groupKey);
    if (!groupRepeat) {
      groupRepeat = { products: new Map() };
      repeatByGroup.set(order.groupKey, groupRepeat);
    }
    let productRepeat = groupRepeat.products.get(order.product);
    if (!productRepeat) {
      productRepeat = { product: order.product, category: order.category, orders: [] };
      groupRepeat.products.set(order.product, productRepeat);
    }
    productRepeat.orders.push(order);
  }

  const buildRepeatAnalysis = (key: string) => {
    const groupRepeat = repeatByGroup.get(key);
    const productRows = [...(groupRepeat?.products.values() ?? [])]
      .map((g) => {
        const orders = [...g.orders].sort((a, b) => {
          const am = Number.isFinite(a.ms) ? a.ms : Number.POSITIVE_INFINITY;
          const bm = Number.isFinite(b.ms) ? b.ms : Number.POSITIVE_INFINITY;
          return am - bm || a.lineIndex - b.lineIndex;
        });
        const replenishmentOrders = orders.slice(1);
        const intervals: number[] = [];
        for (let i = 1; i < orders.length; i++) {
          if (Number.isFinite(orders[i - 1].ms) && Number.isFinite(orders[i].ms)) intervals.push(daysBetween(orders[i - 1].ms, orders[i].ms));
        }
        const first = orders[0];
        const last = orders[orders.length - 1];
        const repeatAmount = replenishmentOrders.reduce((s, r) => s + r.amount, 0);
        const repeatQty = replenishmentOrders.reduce((s, r) => s + r.qty, 0);
        const totalAmount = orders.reduce((s, r) => s + r.amount, 0);
        const totalQty = orders.reduce((s, r) => s + r.qty, 0);
        const avgInterval = intervals.length ? intervals.reduce((s, v) => s + v, 0) / intervals.length : 0;
        const minInterval = intervals.length ? Math.min(...intervals) : 0;
        return {
          product: g.product,
          category: g.category,
          order_count: orders.length,
          repeat_count: Math.max(0, orders.length - 1),
          repeat_amount: Math.round(repeatAmount),
          repeat_qty: Math.round(repeatQty),
          total_amount: Math.round(totalAmount),
          total_qty: Math.round(totalQty),
          first_amount: Math.round(first?.amount ?? 0),
          first_qty: Math.round(first?.qty ?? 0),
          first_date: first && Number.isFinite(first.ms) ? previewDateText(first.ms) : '',
          last_date: last && Number.isFinite(last.ms) ? previewDateText(last.ms) : '',
          avg_interval_days: round1(avgInterval),
          min_interval_days: minInterval,
        };
      })
      .filter((r) => r.order_count >= 2 && hasPositiveNet(r.repeat_amount, r.repeat_qty))
      .sort((a, b) => b.repeat_qty - a.repeat_qty || b.repeat_amount - a.repeat_amount || b.order_count - a.order_count);
    const totalRepeatAmount = productRows.reduce((s, r) => s + r.repeat_amount, 0);
    const totalRepeatQty = productRows.reduce((s, r) => s + r.repeat_qty, 0);
    const intervalValues = productRows.map((r) => r.avg_interval_days).filter((v) => v > 0);
    const avgIntervalDays = intervalValues.length ? round1(intervalValues.reduce((s, v) => s + v, 0) / intervalValues.length) : 0;
    const categoryMap = new Map<string, { name: string; repeat_amount: number; repeat_qty: number; products: number }>();
    for (const row of productRows) {
      const g = categoryMap.get(row.category) ?? { name: row.category, repeat_amount: 0, repeat_qty: 0, products: 0 };
      g.repeat_amount += row.repeat_amount;
      g.repeat_qty += row.repeat_qty;
      g.products += 1;
      categoryMap.set(row.category, g);
    }
    const categories = [...categoryMap.values()]
      .map((r) => ({ ...r, repeat_amount: Math.round(r.repeat_amount), repeat_qty: Math.round(r.repeat_qty) }))
      .sort((a, b) => b.repeat_qty - a.repeat_qty || b.repeat_amount - a.repeat_amount);
    const intervalBuckets = [
      { name: '0-7天', min: 0, max: 7 },
      { name: '8-14天', min: 8, max: 14 },
      { name: '15-30天', min: 15, max: 30 },
      { name: '31-60天', min: 31, max: 60 },
      { name: '60天以上', min: 61, max: Infinity },
    ].map((bucket) => ({
      name: bucket.name,
      value: productRows.filter((r) => r.min_interval_days >= bucket.min && r.min_interval_days <= bucket.max).length,
    })).filter((r) => r.value > 0);
    return {
      has_repeat: productRows.length > 0,
      repeat_product_count: productRows.length,
      repeat_order_count: productRows.reduce((s, r) => s + r.repeat_count, 0),
      repeat_amount: Math.round(totalRepeatAmount),
      repeat_qty: Math.round(totalRepeatQty),
      avg_interval_days: avgIntervalDays,
      products: productRows.slice(0, 20),
      categories,
      interval_buckets: intervalBuckets,
    };
  };

  const toProfile = (g: ProfileGroup) => {
      const amount = Math.round(g.amount);
      const categories = metricRows(g.categories);
      const brands = metricRows(g.brands);
      const designers = metricRows(g.designers);
      const colors = metricRows(g.colors).map((r) => ({ ...r, hex: colorHex(r.name) }));
      const sizes = metricRows(g.sizes);
      const seasons = metricRows(g.seasons);
      const topCategory = categories[0]?.value ?? 0;
      const topBrand = brands[0]?.value ?? 0;
      const topColor = colors[0]?.value ?? 0;
      const topSeason = seasons[0]?.value ?? 0;
      const avgPrice = g.priceQty ? g.priceAmount / g.priceQty : 0;
      const priceP25 = weightedPricePercentile(g.priceValues, 0.25);
      const priceP75 = weightedPricePercentile(g.priceValues, 0.75);
      const priceSpread = avgPrice ? Math.min(100, ((priceP75 - priceP25) / avgPrice) * 100) : 0;
      const lowAmount = g.priceSegments.get('低价带') ?? 0;
      const mainAmount = g.priceSegments.get('主流价带') ?? 0;
      const highAmount = g.priceSegments.get('高价带') ?? 0;
      const segmentEntries = [
        { type: '低价型', segment: '低价带', amount: lowAmount },
        { type: '主流价型', segment: '主流价带', amount: mainAmount },
        { type: '高价型', segment: '高价带', amount: highAmount },
      ].sort((a, b) => b.amount - a.amount);
      const mainSegment = segmentEntries[0] ?? { type: '混合型', segment: '-', amount: 0 };
      const segmentTotal = lowAmount + mainAmount + highAmount;
      const mainShare = segmentTotal ? mainSegment.amount / segmentTotal : 0;
      const priceType = mainShare >= 0.6 ? mainSegment.type : '混合型';
      const priceTypeDesc = priceType === '低价型'
        ? '主要成交集中在低价带'
        : priceType === '主流价型'
          ? '主要成交集中在主流价带'
          : priceType === '高价型'
            ? '主要成交集中在高价带'
            : priceType === '混合型'
              ? '多个价格带都有成交'
              : '按实际成交价格带判断';
      const categoryPrices = [...g.categoryPrices.entries()]
        .map(([name, cp]) => {
          const avg = cp.priceQty ? cp.priceAmount / cp.priceQty : 0;
          const p25 = weightedPricePercentile(cp.prices, 0.25);
          const p75 = weightedPricePercentile(cp.prices, 0.75);
          const useQuantileRange = cp.prices.length >= 3 && p25 !== p75;
          const low = useQuantileRange ? p25 : avg * 0.85;
          const high = useQuantileRange ? p75 : avg * 1.15;
          const bandAmountTotal = PRICE_BANDS.reduce((sum, b) => sum + (cp.bandAmounts.get(b.label) ?? 0), 0);
          return {
            name,
            amount: Math.round(cp.amount),
            qty: Math.round(cp.qty),
            orders: cp.orderKeys.size,
            avg: round1(avg),
            low: Math.max(0, round1(low)),
            high: Math.max(0, round1(high)),
            price_bands: PRICE_BANDS.map((b) => ({
              name: b.label,
              value: Math.round(cp.bandAmounts.get(b.label) ?? 0),
              qty: Math.round(cp.bandQty.get(b.label) ?? 0),
              share: bandAmountTotal ? (cp.bandAmounts.get(b.label) ?? 0) / bandAmountTotal : 0,
            })).filter((r) => r.value > 0 || r.qty > 0),
          };
        })
        .filter((r) => hasPositiveNet(r.amount, r.qty))
        .sort((a, b) => b.amount - a.amount);
      const repeat = buildRepeatAnalysis(g.key);
      const colorDetails = [...g.colorDetails.entries()]
        .map(([color, cd]) => {
          const toRows = (map: Map<string, ColorBreakdownMetric>, totalQty: number) => [...map.entries()]
            .map(([name, metric]) => {
              const qty = Math.round(metric.qty);
              return {
                name,
                qty,
                amount: Math.round(metric.amount),
                share: totalQty ? qty / totalQty : 0,
              };
            })
            .filter((r) => r.qty > 0)
            .sort((a, b) => b.qty - a.qty || b.amount - a.amount)
            .slice(0, 8);
          const categoryQty = [...cd.categories.values()].reduce((sum, metric) => sum + Math.max(0, Math.round(metric.qty)), 0);
          const brandQty = [...cd.brands.values()].reduce((sum, metric) => sum + Math.max(0, Math.round(metric.qty)), 0);
          const qty = Math.max(categoryQty, brandQty, Math.round(cd.qty));
          return {
            color,
            hex: colorHex(color),
            qty,
            amount: Math.round(cd.amount),
            categories: toRows(cd.categories, categoryQty || qty),
            brands: toRows(cd.brands, brandQty || qty),
          };
        })
        .filter((r) => r.qty > 0)
        .sort((a, b) => b.qty - a.qty || b.amount - a.amount)
        .slice(0, 15);
      return {
        customer: g.customer,
        brand: g.brand,
        amount,
        qty: Math.round(g.qty),
        orders: g.orderKeys.size,
        avg_price: round1(avgPrice),
        price_low: round1(priceP25),
        price_high: round1(priceP75),
        price_type: priceType,
        price_type_desc: priceTypeDesc,
        price_type_share: round1(mainShare * 100),
        main_price_segment: mainSegment.segment,
        price_segments: [
          { name: '低价带', value: Math.round(lowAmount), share: round1(segmentTotal ? (lowAmount / segmentTotal) * 100 : 0) },
          { name: '主流价带', value: Math.round(mainAmount), share: round1(segmentTotal ? (mainAmount / segmentTotal) * 100 : 0) },
          { name: '高价带', value: Math.round(highAmount), share: round1(segmentTotal ? (highAmount / segmentTotal) * 100 : 0) },
        ].filter((r) => r.value > 0),
        categories,
        brands,
        designers,
        colors,
        color_details: colorDetails,
        sizes,
        category_prices: categoryPrices,
        price_bands: PRICE_BANDS.map((b) => ({ name: b.label, value: Math.round(g.priceBands.get(b.label) ?? 0) })).filter((r) => r.value > 0),
        seasons,
        repeat,
        radar: [
          { name: '品类集中', value: amount ? round1((topCategory / amount) * 100) : 0 },
          { name: '品牌集中', value: amount ? round1((topBrand / amount) * 100) : 0 },
          { name: '颜色集中', value: g.qty ? round1((topColor / g.qty) * 100) : 0 },
          { name: '年份季节集中', value: amount ? round1((topSeason / amount) * 100) : 0 },
          { name: '价格稳定', value: round1(Math.max(0, 100 - priceSpread)) },
          { name: '客单强度', value: round1(Math.min(100, avgPrice / 3)) },
        ],
      };
    };

  const profileRows = [...groups.values()]
    .filter((g) => hasPositiveNet(g.amount, g.qty))
    .map(toProfile)
    .sort((a, b) => b.amount - a.amount);
  const profiles = profileRows
    .filter((p) => !p.brand)
    .map((p) => {
      const brandProfiles = profileRows
        .filter((bp) => bp.customer === p.customer && !!bp.brand)
        .sort((a, b) => b.amount - a.amount);
      return {
        ...p,
        brand_options: brandProfiles.map((bp) => ({
          name: bp.brand || '',
          amount: bp.amount,
          qty: bp.qty,
          orders: bp.orders,
        })),
        brand_profiles: brandProfiles,
      };
    });

  return {
    customers: profiles.map((p) => p.customer),
    profiles,
  };
}

function previewDateText(ms: number): string {
  return Number.isFinite(ms) ? new Date(ms).toISOString().slice(0, 10) : '';
}

function buildCategoryPreferenceAnalysis(data: PreferenceData, orderIds: (Cell | null)[]) {
  const catCol = data.cols.get('分类');
  if (!catCol) {
    return {
      total_categories: 0,
      recent_window_days: 0,
      recent_from: '',
      recent_to: '',
      prior_from: '',
      prior_to: '',
      trend_period_type: '',
      trend_periods: [],
      category_trend: [],
      categories: [],
      shop_best_categories: [],
      cooccurrence: [],
    };
  }
  const custCol = data.cols.get('客户名称');
  const shopCol = data.cols.get('店铺');
  const dayMs = 24 * 60 * 60 * 1000;
  let minMs = Infinity;
  let maxMs = -Infinity;

  for (let i = 0; i < data.rawRowCount; i++) {
    const amount = data.amt[i];
    const qty = data.qty[i];
    const category = cellText(catCol[i]);
    if (!category || !Number.isFinite(amount) || !Number.isFinite(qty)) continue;
    const ms = data.orderMs?.[i];
    if (ms != null && Number.isFinite(ms)) {
      minMs = Math.min(minMs, ms);
      maxMs = Math.max(maxMs, ms);
    }
  }

  const hasDates = Number.isFinite(minMs) && Number.isFinite(maxMs) && maxMs >= minMs;
  const spanDays = hasDates ? Math.max(1, Math.ceil((maxMs - minMs) / dayMs) + 1) : 0;
  const windowDays = hasDates ? Math.max(7, Math.min(30, Math.floor(spanDays / 2) || 7)) : 0;
  const recentStart = hasDates ? maxMs - windowDays * dayMs : Infinity;
  const priorStart = hasDates ? recentStart - windowDays * dayMs : Infinity;
  const pad2 = (v: number) => String(v).padStart(2, '0');
  const dayStart = (ms: number) => {
    const d = new Date(ms);
    return Date.UTC(d.getUTCFullYear(), d.getUTCMonth(), d.getUTCDate());
  };
  const shortDate = (ms: number) => {
    const d = new Date(ms);
    return `${pad2(d.getUTCMonth() + 1)}-${pad2(d.getUTCDate())}`;
  };
  const dateKey = (ms: number) => previewDateText(dayStart(ms));
  const trendPeriodType = !hasDates ? '' : spanDays <= 21 ? '日' : spanDays <= 100 ? '周' : '月';
  const trendPeriods: { key: string; label: string; from: string; to: string }[] = [];
  const startDay = hasDates ? dayStart(minMs) : NaN;
  const endDay = hasDates ? dayStart(maxMs) : NaN;
  if (hasDates && trendPeriodType === '日') {
    for (let t = startDay; t <= endDay; t += dayMs) {
      trendPeriods.push({ key: dateKey(t), label: shortDate(t), from: previewDateText(t), to: previewDateText(t) });
    }
  } else if (hasDates && trendPeriodType === '周') {
    for (let t = startDay, i = 1; t <= endDay; t += 7 * dayMs, i++) {
      const end = Math.min(t + 6 * dayMs, endDay);
      trendPeriods.push({ key: `w${i}`, label: `${shortDate(t)}-${shortDate(end)}`, from: previewDateText(t), to: previewDateText(end) });
    }
  } else if (hasDates && trendPeriodType === '月') {
    const start = new Date(startDay);
    let y = start.getUTCFullYear();
    let m = start.getUTCMonth();
    for (;;) {
      const monthStart = Date.UTC(y, m, 1);
      if (monthStart > endDay) break;
      const monthEnd = Date.UTC(y, m + 1, 0);
      const key = `${y}-${pad2(m + 1)}`;
      trendPeriods.push({
        key,
        label: key,
        from: previewDateText(Math.max(monthStart, startDay)),
        to: previewDateText(Math.min(monthEnd, endDay)),
      });
      m += 1;
      if (m > 11) {
        m = 0;
        y += 1;
      }
    }
  }
  const periodKeyFor = (ms: number) => {
    if (!hasDates || !Number.isFinite(ms)) return '';
    const t = dayStart(ms);
    if (trendPeriodType === '日') return dateKey(t);
    if (trendPeriodType === '周') return `w${Math.floor((t - startDay) / (7 * dayMs)) + 1}`;
    const d = new Date(t);
    return `${d.getUTCFullYear()}-${pad2(d.getUTCMonth() + 1)}`;
  };

  type CategoryGroup = {
    category: string;
    amount: number;
    qty: number;
    lines: number;
    orders: Set<string>;
    customers: Set<string>;
    customerAmount: Map<string, number>;
    recentAmount: number;
    recentQty: number;
    priorAmount: number;
    priorQty: number;
  };

  const groups = new Map<string, CategoryGroup>();
  const shopCategory = new Map<string, { amount: number; qty: number; customers: Set<string> }>();
  const orderCats = new Map<string, Set<string>>();
  const trendMap = new Map<string, { amount: number; qty: number }>();
  const getGroup = (category: string) => {
    let g = groups.get(category);
    if (!g) {
      g = {
        category,
        amount: 0,
        qty: 0,
        lines: 0,
        orders: new Set<string>(),
        customers: new Set<string>(),
        customerAmount: new Map<string, number>(),
        recentAmount: 0,
        recentQty: 0,
        priorAmount: 0,
        priorQty: 0,
      };
      groups.set(category, g);
    }
    return g;
  };

  for (let i = 0; i < data.rawRowCount; i++) {
    const amount = data.amt[i];
    const qty = data.qty[i];
    if (!Number.isFinite(amount) || !Number.isFinite(qty)) continue;
    const category = cellText(catCol[i]);
    if (!category) continue;
    const customer = custCol ? cellText(custCol[i]) ?? '未填写客户' : '未填写客户';
    const shop = shopCol ? cellText(shopCol[i]) ?? '未标记' : '未标记';
    const key = orderKey(orderIds, i, customer);
    const g = getGroup(category);
    g.amount += amount;
    g.qty += qty;
    g.lines += 1;
    g.orders.add(key);
    g.customers.add(customer);
    g.customerAmount.set(customer, (g.customerAmount.get(customer) ?? 0) + amount);
    addMetric(shopCategory, `${shop}\u0000${category}`, amount, qty, customer);

    let cats = orderCats.get(key);
    if (!cats) {
      cats = new Set<string>();
      orderCats.set(key, cats);
    }
    cats.add(category);

    const ms = data.orderMs?.[i];
    if (hasDates && ms != null && Number.isFinite(ms)) {
      const pkey = periodKeyFor(ms);
      if (pkey) {
        const tkey = `${category}\u0000${pkey}`;
        const tg = trendMap.get(tkey) ?? { amount: 0, qty: 0 };
        tg.amount += amount;
        tg.qty += qty;
        trendMap.set(tkey, tg);
      }
      if (ms > recentStart && ms <= maxMs) {
        g.recentAmount += amount;
        g.recentQty += qty;
      } else if (ms > priorStart && ms <= recentStart) {
        g.priorAmount += amount;
        g.priorQty += qty;
      }
    }
  }

  const baseRows = [...groups.values()].map((g) => {
    const topCustomer = [...g.customerAmount.entries()].sort((a, b) => b[1] - a[1])[0];
    const growthAmount = g.recentAmount - g.priorAmount;
    const growthQty = g.recentQty - g.priorQty;
    const growthRate = g.priorAmount > 0 ? (growthAmount / g.priorAmount) * 100 : g.recentAmount > 0 ? 100 : 0;
    const trendAmounts = trendPeriods.map((p) => trendMap.get(`${g.category}\u0000${p.key}`)?.amount ?? 0);
    const half = Math.floor(trendAmounts.length / 2);
    const early = half > 0 ? trendAmounts.slice(0, half).reduce((s, v) => s + v, 0) / half : trendAmounts[0] ?? 0;
    const lateRows = half > 0 ? trendAmounts.slice(half) : trendAmounts.slice(1);
    const late = lateRows.length ? lateRows.reduce((s, v) => s + v, 0) / lateRows.length : early;
    const trendDelta = late - early;
    const trendRate = early !== 0 ? (trendDelta / Math.abs(early)) * 100 : late > 0 ? 100 : 0;
    let trendStatus = '周期不足';
    if (trendAmounts.length >= 2) {
      if (trendRate >= 15 && trendDelta > 0) trendStatus = '走势上升';
      else if (trendRate <= -15 && trendDelta < 0) trendStatus = '走势回落';
      else trendStatus = '走势平稳';
    }
    return {
      category: g.category,
      amount: Math.round(g.amount),
      qty: Math.round(g.qty),
      lines: g.lines,
      orders: g.orders.size,
      customers: g.customers.size,
      avg_price: round1(g.qty ? g.amount / g.qty : 0),
      top_customer: topCustomer?.[0] ?? '-',
      top_customer_amount: Math.round(topCustomer?.[1] ?? 0),
      top_customer_share: g.amount ? (topCustomer?.[1] ?? 0) / g.amount : 0,
      recent_amount: Math.round(g.recentAmount),
      recent_qty: Math.round(g.recentQty),
      prior_amount: Math.round(g.priorAmount),
      prior_qty: Math.round(g.priorQty),
      growth_amount: Math.round(growthAmount),
      growth_qty: Math.round(growthQty),
      growth_rate: round1(growthRate),
      trend_status: trendStatus,
      trend_delta: Math.round(trendDelta),
      trend_rate: round1(trendRate),
    };
  }).filter((r) => hasPositiveNet(r.amount, r.qty));

  const totalAmount = baseRows.reduce((s, r) => s + r.amount, 0);
  const totalQty = baseRows.reduce((s, r) => s + r.qty, 0);
  const allCustomers = new Set<string>();
  for (const g of groups.values()) {
    for (const customer of g.customers) allCustomers.add(customer);
  }
  const totalCustomers = allCustomers.size;
  const amountValues = baseRows.map((r) => r.amount).sort((a, b) => a - b);
  const customerValues = baseRows.map((r) => r.customers).sort((a, b) => a - b);
  const amountMedian = percentile(amountValues, 0.5);
  const customerMedian = percentile(customerValues, 0.5);
  const maxAmount = Math.max(1, ...baseRows.map((r) => r.amount));
  const maxQty = Math.max(1, ...baseRows.map((r) => r.qty));
  const maxCustomers = Math.max(1, ...baseRows.map((r) => r.customers));

  const pairCounts = new Map<string, number>();
  for (const cats of orderCats.values()) {
    const arr = [...cats].sort();
    if (arr.length < 2) continue;
    for (const from of arr) {
      for (const to of arr) {
        if (from === to) continue;
        const key = `${from}\u0000${to}`;
        pairCounts.set(key, (pairCounts.get(key) ?? 0) + 1);
      }
    }
  }

  const orderCountByCategory = new Map(baseRows.map((r) => [r.category, r.orders]));
  const cooccurrence = [...pairCounts.entries()]
    .map(([key, orders]) => {
      const [from, to] = key.split('\u0000');
      const fromOrders = orderCountByCategory.get(from) ?? 0;
      return {
        from,
        to,
        orders,
        rate: fromOrders ? orders / fromOrders : 0,
      };
    })
    .sort((a, b) => b.rate - a.rate || b.orders - a.orders);

  const companionMap = new Map<string, { category: string; orders: number; rate: number }>();
  for (const row of cooccurrence) {
    const cur = companionMap.get(row.from);
    if (!cur || row.rate > cur.rate || (row.rate === cur.rate && row.orders > cur.orders)) {
      companionMap.set(row.from, { category: row.to, orders: row.orders, rate: row.rate });
    }
  }

  const categories = baseRows.map((r) => {
    const amountScore = r.amount / maxAmount;
    const qtyScore = r.qty / maxQty;
    const customerScore = r.customers / maxCustomers;
    const trendScore = r.trend_status === '走势上升' ? 1 : r.trend_status === '走势平稳' ? 0.45 : 0;
    const stabilityScore = Math.max(0, 1 - r.top_customer_share);
    const amountStrong = r.amount >= amountMedian || amountScore >= 0.35;
    const customerStrong = r.customers >= customerMedian || customerScore >= 0.45;
    const concentrationHigh = r.top_customer_share >= 0.35;
    let categoryClass = '控量测试';
    if (amountStrong && customerStrong && !concentrationHigh) {
      categoryClass = '稳定主力';
    } else if (amountStrong && (!customerStrong || concentrationHigh)) {
      categoryClass = '少数客户拉动';
    } else if (!amountStrong && customerStrong) {
      categoryClass = '广覆盖连带';
    } else if (r.trend_status === '走势上升') {
      categoryClass = '潜力观察';
    }
    const companion = companionMap.get(r.category);
    let saleFocus = '控量测试';
    let saleAdvice = '少量陈列，优先选择卖点清晰、退换风险低的款。';
    if (categoryClass === '稳定主力') {
      saleFocus = '主推备货';
      saleAdvice = '作为日常主推入口，保证主销款深度，并用于带动搭配销售。';
    } else if (categoryClass === '少数客户拉动') {
      saleFocus = '定向推荐';
      saleAdvice = '优先推荐给高贡献客户，备货不要铺太散，适合做定向补货。';
    } else if (categoryClass === '广覆盖连带') {
      saleFocus = '连带搭配';
      saleAdvice = '覆盖客户多，适合做进店搭配和加购推荐，提高连带件数。';
    } else if (categoryClass === '潜力观察') {
      saleFocus = '测款放量';
      saleAdvice = '报表期内走势上升，可保留动销款并小幅增加相近款测试。';
    }
    return {
      ...r,
      amount_share: totalAmount ? r.amount / totalAmount : 0,
      qty_share: totalQty ? r.qty / totalQty : 0,
      customer_share: totalCustomers ? r.customers / totalCustomers : 0,
      top_customer_share: round1(r.top_customer_share * 100),
      class: categoryClass,
      recommend_score: round1(100 * (amountScore * 0.32 + qtyScore * 0.25 + customerScore * 0.25 + trendScore * 0.1 + stabilityScore * 0.08)),
      companion: companion?.category ?? '-',
      companion_orders: companion?.orders ?? 0,
      companion_rate: companion ? round1(companion.rate * 100) : 0,
      sale_focus: saleFocus,
      sale_advice: saleAdvice,
    };
  }).sort((a, b) => b.amount - a.amount);

  const categoryTrend = categories.flatMap((category) => trendPeriods.map((period) => {
    const row = trendMap.get(`${category.category}\u0000${period.key}`);
    return {
      category: category.category,
      period: period.key,
      label: period.label,
      amount: Math.round(row?.amount ?? 0),
      qty: Math.round(row?.qty ?? 0),
    };
  }));

  const shopBestCategories = [...shopCategory.entries()]
    .map(([key, g]) => {
      const [shop, category] = key.split('\u0000');
      return { shop, category, amount: Math.round(g.amount), qty: Math.round(g.qty), customers: g.customers.size };
    })
    .filter((r) => hasPositiveNet(r.amount, r.qty))
    .sort((a, b) => a.shop.localeCompare(b.shop) || b.amount - a.amount)
    .reduce<{ shop: string; category: string; amount: number; qty: number; customers: number }[]>((rows, row) => {
      if (!rows.some((r) => r.shop === row.shop)) rows.push(row);
      return rows;
    }, [])
    .sort((a, b) => b.amount - a.amount);

  return {
    total_categories: categories.length,
    recent_window_days: windowDays,
    recent_from: previewDateText(recentStart + dayMs),
    recent_to: previewDateText(maxMs),
    prior_from: previewDateText(priorStart + dayMs),
    prior_to: previewDateText(recentStart),
    trend_period_type: trendPeriodType,
    trend_periods: trendPeriods,
    category_trend: categoryTrend,
    categories,
    shop_best_categories: shopBestCategories,
    cooccurrence,
  };
}

function buildCustomerCategoryPriceAnalysis(data: PreferenceData, orderIds: (Cell | null)[]) {
  const custCol = data.cols.get('客户名称');
  const catCol = data.cols.get('分类');
  if (!custCol || !catCol) {
    return { categories: [], customers: [], matrix_customers: [], rows: [], labels: [] };
  }

  type Group = {
    customer: string;
    category: string;
    amount: number;
    qty: number;
    orderKeys: Set<string>;
    prices: number[];
    bandQty: Map<string, number>;
    bandAmount: Map<string, number>;
  };

  const catGroups = new Map<string, Group>();
  const groups = new Map<string, Group>();
  const customerAmount = new Map<string, number>();

  const getGroup = (map: Map<string, Group>, key: string, customer: string, category: string) => {
    let g = map.get(key);
    if (!g) {
      g = { customer, category, amount: 0, qty: 0, orderKeys: new Set(), prices: [], bandQty: new Map(), bandAmount: new Map() };
      map.set(key, g);
    }
    return g;
  };

  for (let i = 0; i < data.rawRowCount; i++) {
    const qty = data.qty[i];
    const amount = data.amt[i];
    if (!Number.isFinite(qty) || !Number.isFinite(amount) || qty <= 0 || amount <= 0) continue;
    const price = amount / qty;
    if (!Number.isFinite(price) || price <= 0) continue;

    const customer = cellText(custCol[i]);
    const category = cellText(catCol[i]);
    if (!customer || !category) continue;

    const band = priceBand(price);
    const key = orderKey(orderIds, i, customer);
    for (const g of [
      getGroup(groups, `${customer}\u0000${category}`, customer, category),
      getGroup(catGroups, category, '', category),
    ]) {
      g.amount += amount;
      g.qty += qty;
      g.orderKeys.add(key);
      g.prices.push(price);
      const bandQty = (g.bandQty.get(band) ?? 0) + Math.max(qty, 0);
      const bandAmount = (g.bandAmount.get(band) ?? 0) + Math.max(amount, 0);
      g.bandQty.set(band, bandQty);
      g.bandAmount.set(band, bandAmount);
    }
    customerAmount.set(customer, (customerAmount.get(customer) ?? 0) + amount);
  }

  const categoryBenchmarks = [...catGroups.values()]
    .map((g) => {
      const prices = [...g.prices].sort((a, b) => a - b);
      return {
        category: g.category,
        amount: Math.round(g.amount),
        qty: Math.round(g.qty),
        orders: g.orderKeys.size,
        weighted_avg: round1(g.qty ? g.amount / g.qty : 0),
        median: round1(percentile(prices, 0.5)),
        p25: round1(percentile(prices, 0.25)),
        p75: round1(percentile(prices, 0.75)),
      };
    })
    .sort((a, b) => b.amount - a.amount);
  const benchmarkMap = new Map(categoryBenchmarks.map((r) => [r.category, r]));

  const rows = [...groups.values()]
    .filter((g) => g.orderKeys.size > 0)
    .map((g) => {
      const prices = [...g.prices].sort((a, b) => a - b);
      const p25 = percentile(prices, 0.25);
      const p50 = percentile(prices, 0.5);
      const p75 = percentile(prices, 0.75);
      const weightedAvg = g.qty ? g.amount / g.qty : p50;
      const benchmark = benchmarkMap.get(g.category);
      const mainBand = [...g.bandAmount.entries()]
        .sort((a, b) => b[1] - a[1])[0]?.[0] ?? priceBand(weightedAvg);
      const orders = g.orderKeys.size;
      const enoughSample = g.qty >= 5 && orders >= 3;
      const volatility = p50 > 0 && (p75 - p25) / p50 >= 0.6;
      let tag = '样本不足';
      if (enoughSample && volatility) tag = '价格波动型';
      else if (enoughSample && benchmark && weightedAvg <= benchmark.p25) tag = '低价敏感型';
      else if (enoughSample && benchmark && weightedAvg > benchmark.p75) tag = '高价接受型';
      else if (enoughSample) tag = '中价稳定型';

      const recLow = enoughSample ? p25 : weightedAvg * 0.85;
      const recHigh = enoughSample ? p75 : weightedAvg * 1.15;

      return {
        customer: g.customer,
        category: g.category,
        amount: Math.round(g.amount),
        qty: Math.round(g.qty),
        orders,
        weighted_avg: round1(weightedAvg),
        median: round1(p50),
        p25: round1(p25),
        p75: round1(p75),
        main_band: mainBand,
        tag,
        tag_rank: priceTagClass(tag),
        recommend_low: round1(recLow),
        recommend_high: round1(recHigh),
      };
    })
    .sort((a, b) => b.amount - a.amount);

  const customers = [...customerAmount.entries()]
    .sort((a, b) => b[1] - a[1])
    .map(([customer]) => customer);

  return {
    categories: categoryBenchmarks,
    customers,
    matrix_customers: customers.slice(0, 30),
    rows,
    labels: ['低价敏感型', '中价稳定型', '高价接受型', '价格波动型', '样本不足'],
  };
}

function buildPriceAcceptanceAnalysis(data: PreferenceData, orderIds: (Cell | null)[], segmentation: PriceSegmentation) {
  const custCol = data.cols.get('客户名称');
  const catCol = data.cols.get('分类');

  type MetricGroup = {
    label: string;
    amount: number;
    qty: number;
    lines: number;
    orderKeys: Set<string>;
  };

  type CategoryGroup = {
    category: string;
    amount: number;
    qty: number;
    lines: number;
    orderKeys: Set<string>;
    customers: Set<string>;
    prices: { price: number; qty: number }[];
    lowAmount: number;
    lowQty: number;
    mainAmount: number;
    mainQty: number;
    highAmount: number;
    highQty: number;
  };

  const newMetricGroup = (label: string): MetricGroup => ({ label, amount: 0, qty: 0, lines: 0, orderKeys: new Set<string>() });
  const bandGroups = new Map<string, MetricGroup>(ACCEPTANCE_PRICE_BANDS.map((label) => [label, newMetricGroup(label)]));
  const segmentGroups = new Map<string, MetricGroup>(ACCEPTANCE_SEGMENTS.map((label) => [label, newMetricGroup(label)]));
  const categoryGroups = new Map<string, CategoryGroup>();
  const allPrices: { price: number; qty: number }[] = [];

  const getCategory = (category: string) => {
    let g = categoryGroups.get(category);
    if (!g) {
      g = {
        category,
        amount: 0,
        qty: 0,
        lines: 0,
        orderKeys: new Set<string>(),
        customers: new Set<string>(),
        prices: [],
        lowAmount: 0,
        lowQty: 0,
        mainAmount: 0,
        mainQty: 0,
        highAmount: 0,
        highQty: 0,
      };
      categoryGroups.set(category, g);
    }
    return g;
  };

  for (let i = 0; i < data.rawRowCount; i++) {
    const qty = data.qty[i];
    const amount = data.amt[i];
    if (!Number.isFinite(qty) || !Number.isFinite(amount) || qty <= 0 || amount <= 0) continue;
    const price = amount / qty;
    if (!Number.isFinite(price) || price <= 0) continue;

    const customer = custCol ? cellText(custCol[i]) ?? '未填写客户' : '未填写客户';
    const category = catCol ? cellText(catCol[i]) ?? '未分类' : '未分类';
    const key = orderKey(orderIds, i, customer);
    const band = acceptancePriceBand(price);
    const segment = acceptanceSegment(price, segmentation);

    allPrices.push({ price, qty });

    const bg = bandGroups.get(band)!;
    bg.amount += amount;
    bg.qty += qty;
    bg.lines += 1;
    bg.orderKeys.add(key);

    const sg = segmentGroups.get(segment)!;
    sg.amount += amount;
    sg.qty += qty;
    sg.lines += 1;
    sg.orderKeys.add(key);

    const cg = getCategory(category);
    cg.amount += amount;
    cg.qty += qty;
    cg.lines += 1;
    cg.orderKeys.add(key);
    cg.customers.add(customer);
    cg.prices.push({ price, qty });
    if (segment === '低价带') {
      cg.lowAmount += amount;
      cg.lowQty += qty;
    } else if (segment === '主流价带') {
      cg.mainAmount += amount;
      cg.mainQty += qty;
    } else {
      cg.highAmount += amount;
      cg.highQty += qty;
    }
  }

  const totalAmount = [...bandGroups.values()].reduce((s, g) => s + g.amount, 0);
  const totalQty = [...bandGroups.values()].reduce((s, g) => s + g.qty, 0);
  const overallP25 = weightedPricePercentile(allPrices, 0.25);
  const overallP50 = weightedPricePercentile(allPrices, 0.5);
  const overallP75 = weightedPricePercentile(allPrices, 0.75);
  const overallP90 = weightedPricePercentile(allPrices, 0.9);

  const toMetricRow = (g: MetricGroup) => ({
    name: g.label,
    amount: Math.round(g.amount),
    qty: Math.round(g.qty),
    orders: g.orderKeys.size,
    lines: g.lines,
    amount_share: totalAmount ? g.amount / totalAmount : 0,
    qty_share: totalQty ? g.qty / totalQty : 0,
  });

  const bands = ACCEPTANCE_PRICE_BANDS.map((label) => toMetricRow(bandGroups.get(label)!))
    .filter((r) => r.amount > 0 || r.qty > 0 || r.orders > 0 || r.lines > 0);
  const segments = ACCEPTANCE_SEGMENTS.map((label) => toMetricRow(segmentGroups.get(label)!))
    .filter((r) => r.amount > 0 || r.qty > 0 || r.orders > 0 || r.lines > 0);

  const categories = [...categoryGroups.values()]
    .map((g) => {
      const p25 = weightedPricePercentile(g.prices, 0.25);
      const p50 = weightedPricePercentile(g.prices, 0.5);
      const p75 = weightedPricePercentile(g.prices, 0.75);
      const p90 = weightedPricePercentile(g.prices, 0.9);
      const avg = g.qty ? g.amount / g.qty : 0;
      const highAmountShare = g.amount ? g.highAmount / g.amount : 0;
      const highQtyShare = g.qty ? g.highQty / g.qty : 0;
      const mainAmountShare = g.amount ? g.mainAmount / g.amount : 0;
      let acceptance = '低价为主';
      if (highAmountShare >= 0.5 || avg >= segmentation.highCutoff * 1.2) acceptance = '高价承接强';
      else if (highAmountShare >= 0.2 || p75 >= segmentation.highCutoff) acceptance = '有高价空间';
      else if (mainAmountShare >= 0.5) acceptance = '主流价带稳定';
      return {
        category: g.category,
        amount: Math.round(g.amount),
        qty: Math.round(g.qty),
        orders: g.orderKeys.size,
        customers: g.customers.size,
        avg_price: round1(avg),
        p25: round1(p25),
        p50: round1(p50),
        p75: round1(p75),
        p90: round1(p90),
        low_amount_share: g.amount ? g.lowAmount / g.amount : 0,
        low_amount: Math.round(g.lowAmount),
        low_qty: Math.round(g.lowQty),
        main_amount_share: mainAmountShare,
        main_amount: Math.round(g.mainAmount),
        main_qty: Math.round(g.mainQty),
        high_amount: Math.round(g.highAmount),
        high_qty: Math.round(g.highQty),
        high_amount_share: highAmountShare,
        high_qty_share: highQtyShare,
        acceptance,
      };
    })
    .sort((a, b) => b.amount - a.amount);

  const mostAcceptedBand = [...bands].sort((a, b) => b.qty - a.qty)[0];
  const highestAmountBand = [...bands].sort((a, b) => b.amount - a.amount)[0];
  const mainSegmentByAmount = [...segments].sort((a, b) => b.amount - a.amount)[0];
  const highSegment = segments.find((r) => r.name === '高价带');
  const highCategories = [...categories]
    .filter((r) => r.high_amount > 0)
    .sort((a, b) => b.high_amount - a.high_amount)
    .slice(0, 3)
    .map((r) => r.category);

  return {
    low_cutoff: segmentation.lowCutoff,
    high_cutoff: segmentation.highCutoff,
    total_amount: Math.round(totalAmount),
    total_qty: Math.round(totalQty),
    overall: {
      avg_price: round1(totalQty ? totalAmount / totalQty : 0),
      p25: round1(overallP25),
      p50: round1(overallP50),
      p75: round1(overallP75),
      p90: round1(overallP90),
    },
    kpi: {
      most_accepted_band: mostAcceptedBand?.name ?? '-',
      most_accepted_qty_share: mostAcceptedBand?.qty_share ?? 0,
      highest_amount_band: highestAmountBand?.name ?? '-',
      highest_amount_share: highestAmountBand?.amount_share ?? 0,
      main_segment: mainSegmentByAmount?.name ?? '-',
      main_segment_amount_share: mainSegmentByAmount?.amount_share ?? 0,
      high_amount_share: highSegment?.amount_share ?? 0,
      high_qty_share: highSegment?.qty_share ?? 0,
      recommend_low: round1(overallP25),
      recommend_high: round1(overallP75),
      high_categories: highCategories,
    },
    bands,
    segments,
    categories,
    recommendations: categories.map((r) => ({
      category: r.category,
      amount: r.amount,
      qty: r.qty,
      avg_price: r.avg_price,
      p25: r.p25,
      p50: r.p50,
      p75: r.p75,
      p90: r.p90,
      acceptance: r.acceptance,
    })),
  };
}

function buildPreferenceContentHtml(R: unknown): string {
  return `<!DOCTYPE html>
<html lang="zh-CN"><head><meta charset="UTF-8"><meta name="viewport" content="width=device-width,initial-scale=1">
<title>新客户偏好分析</title>
<script>${INLINE_ECHARTS_SCRIPT}</script>
<style>
*{margin:0;padding:0;box-sizing:border-box}
body{font-family:-apple-system,'Microsoft YaHei','Segoe UI',sans-serif;background:#f3f7ff;color:#10205f;font-size:14px}
.hdr{background:linear-gradient(135deg,#fbfdff,#eef5ff,#e8f1ff);color:#10205f;padding:28px 32px;text-align:center}
.hdr h1{font-size:28px;margin-bottom:6px;letter-spacing:2px}.hdr p{opacity:.7;font-size:13px}
.kpi{display:grid;grid-template-columns:repeat(auto-fit,minmax(150px,1fr));gap:12px;padding:18px 28px}
.kpi-card{background:#fff;border-radius:12px;padding:16px;text-align:center;box-shadow:0 2px 8px rgba(0,0,0,.06)}
.kpi-card .v{font-size:24px;font-weight:700;color:#10205f;margin:4px 0}.kpi-card .l{font-size:11px;color:#53627f;text-transform:uppercase;letter-spacing:1px}.kpi-card .s{font-size:11px;color:#7b87a3}
.tabs{display:flex;gap:0;padding:0 28px;border-bottom:2px solid #e0e0e0;background:#fff;overflow-x:auto;position:sticky;top:0;z-index:100}
.tab{padding:12px 18px;cursor:pointer;font-size:13px;font-weight:500;color:#666;border-bottom:3px solid transparent;white-space:nowrap;transition:.2s}
.tab:hover{color:#155cff;background:#f2f7ff}.tab.on{color:#155cff;border-bottom-color:#155cff;font-weight:600}
.body{padding:20px 28px}.sec{display:none}.sec.on{display:block}
.g{display:grid;grid-template-columns:repeat(auto-fit,minmax(480px,1fr));gap:16px;margin-bottom:16px}
.c{background:#fff;border-radius:12px;padding:20px;box-shadow:0 2px 8px rgba(0,0,0,.06)}
.c.full{grid-column:1/-1}.c h3{font-size:15px;margin-bottom:14px;color:#10205f;border-left:4px solid #155cff;padding-left:12px}
.ch{position:relative;height:360px}.ch.tall{height:480px}
.echart{width:100%;height:100%}
.tip{background:linear-gradient(135deg,#fffde7,#fff8e1);border-left:4px solid #ffd600;padding:12px 16px;margin-bottom:14px;border-radius:0 8px 8px 0;font-size:13px;line-height:1.7}
.tip b{color:#10205f}
table{width:100%;border-collapse:collapse;font-size:12px}
th{background:#f8f9fa;padding:8px 10px;text-align:left;font-weight:600;color:#555;border-bottom:2px solid #e0e0e0;position:sticky;top:0}
td{padding:7px 10px;border-bottom:1px solid #f0f0f0}tr:hover td{background:#f8f9ff}
.st{max-height:420px;overflow-y:auto}
.ov-summary{display:grid;grid-template-columns:repeat(2,minmax(0,1fr));gap:10px}
.ov-item{border:1px solid #eee;border-radius:10px;background:#fafafa;padding:11px 12px;min-height:92px}
.ov-item span{display:block;font-size:11px;color:#888;margin-bottom:6px}
.ov-item b{display:flex;align-items:center;gap:7px;color:#10205f;font-size:15px;line-height:1.25;min-height:36px}
.ov-item i{display:inline-block;width:12px;height:12px;border-radius:3px;border:1px solid rgba(0,0,0,.12);flex:0 0 12px}
.ov-item em{display:block;font-style:normal;font-size:12px;color:#555;margin-top:6px}
.ov-item small{display:block;color:#999;font-size:11px;margin-top:3px}
.tag{display:inline-block;padding:2px 8px;border-radius:10px;font-size:11px;font-weight:600}
.t-vip{background:#fce4ec;color:#c62828}.t-h{background:#e8f5e9;color:#2e7d32}.t-m{background:#fff3e0;color:#e65100}.t-l{background:#f5f5f5;color:#666}
.t-new{background:#e3f2fd;color:#1565c0}.t-old{background:#fff8e1;color:#f57f17}.t-bal{background:#f3e5f5;color:#6a1b9a}
.legend{display:flex;gap:10px;flex-wrap:wrap;margin:0 0 12px;font-size:11px;color:#666}
.legend span{display:inline-flex;align-items:center;gap:5px}.legend i{width:10px;height:10px;border-radius:3px;display:inline-block}
.heat{overflow:auto;max-height:520px}
.heat table{min-width:980px}.heat th:first-child,.heat td:first-child{position:sticky;left:0;background:#fff;z-index:2;min-width:108px}
.heat th:first-child{background:#f8f9fa;z-index:3}
.pcell{font-size:11px;line-height:1.35;text-align:center;min-width:70px;border-radius:6px;padding:5px 6px;display:block;font-weight:600}
.pcell small{display:block;font-size:10px;font-weight:500;opacity:.8;margin-top:2px}
.p-low{background:#eaf3ff;color:#155cff}.p-mid{background:#eafaf2;color:#079455}.p-high{background:#fff4e5;color:#b45309}.p-wave{background:#f3efff;color:#6d4aff}.p-lack{background:#f5f7fb;color:#777}
.pf-head{display:flex;gap:12px;align-items:end;flex-wrap:wrap;margin-bottom:16px}
.pf-head label{font-size:12px;color:#666;display:flex;flex-direction:column;gap:5px}
.pf-head select,.pf-head input{min-width:260px;height:34px;border:1px solid #ddd;border-radius:7px;background:#fff;padding:0 10px;color:#10205f}
.pf-brand-box{position:relative;min-width:260px}
.pf-brand-box input{width:100%;box-sizing:border-box}
.pf-brand-box.open .pf-brand-menu{display:block}
.pf-brand-menu{display:none;position:absolute;left:0;right:0;top:calc(100% + 5px);max-height:230px;overflow:auto;background:#fff;border:1px solid #dce7fb;border-radius:8px;box-shadow:0 12px 28px rgba(33,83,170,.16);z-index:500;padding:5px}
.pf-brand-option{display:block;width:100%;border:0;background:transparent;text-align:left;border-radius:6px;padding:8px 9px;color:#303642;font-size:12px;line-height:1.35;cursor:pointer}
.pf-brand-option:hover{background:#f2f7ff;color:#155cff}
.pf-brand-option b{display:block;font-size:12px;color:inherit}
.pf-brand-option span{display:block;margin-top:2px;color:#7a8292;font-size:11px}
.pf-brand-empty{padding:10px;color:#8a92a3;font-size:12px}
.pf-rec{display:block}
.pf-rec-title{display:none}
.pf-rec-row{display:flex;gap:8px;align-items:flex-start;margin-top:6px}
.pf-rec-row b{flex:0 0 74px;color:#666;font-size:12px;font-weight:600;line-height:24px}
.pf-rec-row span{display:inline-block;background:#f8f9fa;border:1px solid #eee;border-radius:14px;padding:3px 9px;margin:0 4px 4px 0;color:#555;font-size:12px}
.pf-mini{display:grid;grid-template-columns:repeat(auto-fit,minmax(140px,1fr));gap:10px;margin-bottom:16px}
.pf-mini div{background:#fff;border-radius:10px;padding:12px;box-shadow:0 2px 8px rgba(0,0,0,.05)}
.pf-mini b{display:block;font-size:18px;color:#10205f;margin-bottom:3px}.pf-mini span{font-size:11px;color:#888}
.pf-color-chart{height:300px;max-width:520px;width:100%;align-self:center}
.accept-note{display:grid;grid-template-columns:repeat(3,minmax(0,1fr));gap:8px;margin-top:12px}
.accept-note div{border:1px solid #eee;border-radius:8px;padding:9px 10px;background:#fafafa}
.accept-note b{display:flex;align-items:center;gap:6px;font-size:12px;color:#10205f;margin-bottom:4px}
.accept-note i{display:inline-block;width:10px;height:10px;border-radius:3px}
.accept-note span{display:block;font-size:11px;color:#666;line-height:1.45}
.range-board{margin-top:6px}
.range-note{display:flex;gap:12px;align-items:center;flex-wrap:wrap;background:#fafafa;border:1px solid #eee;border-radius:8px;padding:8px 10px;margin-bottom:10px;color:#666;font-size:12px;line-height:1.5}
.range-note b{color:#10205f;font-weight:700}
.range-note span{white-space:nowrap}
.range-scale{display:flex;justify-content:space-between;color:#888;font-size:11px;margin:0 124px 8px 150px}
.range-row{display:grid;grid-template-columns:140px minmax(220px,1fr) 114px;gap:10px;align-items:center;min-height:32px;margin:6px 0}
.range-cat{font-size:12px;color:#555;text-align:right;white-space:nowrap;overflow:hidden;text-overflow:ellipsis}
.range-track{height:12px;background:linear-gradient(90deg,#f8f9fa,#eef1f5);border:1px solid #e6e8ec;border-radius:999px;position:relative}
.range-bar{height:14px;position:absolute;top:-2px;border-radius:999px;background:#155cff;box-shadow:0 2px 5px rgba(21,92,255,.18)}
.range-dot{width:8px;height:8px;border-radius:50%;background:#0b4edc;position:absolute;top:1px;transform:translateX(-50%);box-shadow:0 0 0 2px #fff}
.range-value{font-size:12px;color:#10205f;font-weight:600;white-space:nowrap}
.range-value small{display:block;color:#888;font-weight:500;margin-top:2px}
.cat-legend{margin-top:10px}
.cat-note{display:flex;gap:12px;flex-wrap:wrap;background:#fafafa;border:1px solid #eee;border-radius:8px;padding:8px 10px;margin-top:10px;color:#666;font-size:12px;line-height:1.5}
.cat-note b{color:#10205f}
.cat-advice{display:grid;grid-template-columns:repeat(auto-fit,minmax(300px,1fr));gap:12px}
.cat-group{border:1px solid #e8e8e8;border-radius:10px;background:#fff;overflow:hidden}
.cat-group-head{display:flex;align-items:center;justify-content:space-between;gap:10px;padding:10px 12px;border-bottom:1px solid #eee;background:#fafafa}
.cat-group-title{display:flex;align-items:center;gap:8px;font-size:14px;color:#10205f;font-weight:700}
.cat-group-title i{display:inline-block;width:10px;height:10px;border-radius:3px}
.cat-group-count{font-size:11px;color:#777;background:#fff;border:1px solid #eee;border-radius:12px;padding:2px 7px}
.cat-group-desc{font-size:12px;color:#666;line-height:1.5;padding:9px 12px;border-bottom:1px solid #f1f1f1;min-height:54px}
.cat-list{padding:6px 10px 10px}
.cat-line{display:grid;grid-template-columns:minmax(120px,1fr) 62px 90px;gap:8px;align-items:center;padding:8px 0;border-bottom:1px solid #f5f5f5}
.cat-line:last-child{border-bottom:0}
.cat-line b{font-size:13px;color:#10205f;white-space:nowrap;overflow:hidden;text-overflow:ellipsis}
.cat-line span{font-size:11px;color:#777;white-space:nowrap}
.cat-line em{font-style:normal;font-size:11px;color:#555;background:#f8f9fa;border:1px solid #eee;border-radius:12px;padding:2px 7px;text-align:center;white-space:nowrap;overflow:hidden;text-overflow:ellipsis}
.cat-mini-bar{grid-column:1/-1;height:5px;background:#eef1f5;border-radius:999px;overflow:hidden}
.cat-mini-bar i{display:block;height:100%;border-radius:999px}
.cat-empty{font-size:12px;color:#999;padding:12px}
.color-all-wrap{height:560px}
.color-all-wrap .ch{height:560px}
.color-top3-wrap{height:520px;overflow-y:auto;overflow-x:hidden;border:1px solid #eee;border-radius:8px;background:#fff}
.color-top3-wrap .ch{height:560px;min-height:560px}
.customer-color-wrap{height:460px}
.customer-color-wrap .ch{height:460px}
.size-pie-layout{display:grid;grid-template-columns:minmax(240px,1fr) 150px;gap:14px;align-items:center}
.size-pie-layout .ch{height:360px}
.size-pie-legend{display:flex;flex-direction:column;gap:7px;max-height:360px;overflow-y:auto;font-size:12px;color:#333}
.size-pie-legend span{display:flex;align-items:center;gap:8px;line-height:1.2;white-space:nowrap}
.size-pie-legend i{width:24px;height:10px;border-radius:2px;display:inline-block;flex:0 0 24px}
.size-structure-wrap{height:620px;overflow-x:auto;overflow-y:hidden;border:1px solid #eee;border-radius:8px;background:#fff}
.size-structure-wrap .ch{height:600px;min-width:960px}
.size-top-board{display:grid;grid-template-columns:repeat(5,minmax(190px,1fr));gap:12px;overflow-x:auto;padding-bottom:4px}
.size-top-col{border:1px solid var(--line);border-radius:10px;background:linear-gradient(180deg,#fff,#f8fbff);overflow:hidden;min-width:190px}
.size-top-head{display:flex;align-items:center;justify-content:space-between;gap:8px;padding:10px 12px;border-bottom:1px solid var(--line);background:#f7fbff}
.size-top-head b{font-size:18px;color:var(--ink);line-height:1}
.size-top-head span{font-size:11px;color:#687083;white-space:nowrap}
.size-top-list{padding:8px 10px 10px;display:flex;flex-direction:column;gap:8px;max-height:520px;overflow:auto}
.size-top-row{display:grid;grid-template-columns:minmax(64px,1fr) 58px;gap:8px;align-items:center}
.size-top-row strong{font-size:12px;color:#4f5665;font-weight:600;white-space:nowrap;overflow:hidden;text-overflow:ellipsis}
.size-top-row em{font-style:normal;font-size:11px;color:#687083;text-align:right;white-space:nowrap}
.size-top-bar{grid-column:1/-1;height:8px;background:#edf3ff;border-radius:999px;overflow:hidden}
.size-top-bar i{display:block;height:100%;border-radius:999px;background:linear-gradient(90deg,var(--accent),#38bdf8)}
.size-top-empty{font-size:12px;color:#8a92a3;padding:14px 4px}
.brand-tall-wrap{height:680px;overflow-y:auto;overflow-x:hidden;border:1px solid #eee;border-radius:8px;background:#fff}
.brand-tall-wrap .ch{height:760px;min-height:760px}
.brand-designer-board{display:grid;grid-template-columns:repeat(auto-fit,minmax(320px,1fr));gap:14px}
.brand-designer-card{border:1px solid var(--line);border-radius:10px;background:linear-gradient(180deg,#fff,#f8fbff);padding:14px 14px 12px;box-shadow:0 6px 18px rgba(33,83,170,.05)}
.brand-designer-head{display:flex;justify-content:space-between;gap:12px;align-items:flex-start;margin-bottom:12px}
.brand-designer-head b{display:block;color:var(--ink);font-size:16px;line-height:1.25;white-space:nowrap;overflow:hidden;text-overflow:ellipsis;max-width:220px}
.brand-designer-head span{display:block;margin-top:5px;color:#687083;font-size:12px}
.brand-designer-head em{font-style:normal;color:var(--accent);background:#eef5ff;border:1px solid #d7e5ff;border-radius:999px;padding:3px 8px;font-size:11px;font-weight:700;white-space:nowrap}
.brand-designer-list{display:flex;flex-direction:column;gap:10px}
.brand-designer-row{display:grid;grid-template-columns:minmax(70px,1fr) auto;gap:8px;align-items:center}
.brand-designer-row strong{font-size:13px;color:#303642;white-space:nowrap;overflow:hidden;text-overflow:ellipsis}
.brand-designer-row span{font-size:12px;color:#10205f;font-weight:700;white-space:nowrap}
.brand-designer-row small{grid-column:1/-1;color:#7a8292;font-size:11px}
.brand-designer-bar{grid-column:1/-1;height:9px;background:#edf3ff;border-radius:999px;overflow:hidden}
.brand-designer-bar i{display:block;height:100%;border-radius:999px;background:linear-gradient(90deg,var(--accent),#38bdf8)}
.season-summary{display:grid;grid-template-columns:repeat(auto-fit,minmax(190px,1fr));gap:12px}
.season-summary-card{border:1px solid var(--line);border-radius:10px;background:linear-gradient(180deg,#fff,#f8fbff);padding:14px 15px;box-shadow:0 6px 18px rgba(33,83,170,.05)}
.season-summary-card span{display:block;color:#687083;font-size:12px;margin-bottom:8px}
.season-summary-card b{display:block;color:var(--ink);font-size:22px;line-height:1.2;margin-bottom:5px}
.season-summary-card em{font-style:normal;color:#4f5665;font-size:12px}
.season-type-layout{display:grid;grid-template-columns:minmax(340px,.9fr) minmax(320px,1fr);gap:20px;align-items:center}
.season-type-chart{height:390px}
.season-type-defs{display:flex;flex-direction:column;gap:0}
.season-type-def{display:grid;grid-template-columns:12px minmax(112px,auto) 1fr auto;gap:10px;align-items:start;padding:9px 0;border-bottom:1px solid var(--line-soft);font-size:12px;color:#4f5665;line-height:1.55}
.season-type-def:last-child{border-bottom:0}
.season-type-def i{width:10px;height:10px;border-radius:3px;margin-top:4px}
.season-type-def b{color:var(--ink);white-space:nowrap}
.season-type-def em{font-style:normal;color:#687083;white-space:nowrap}
@media(max-width:900px){.season-type-layout{grid-template-columns:1fr}.season-type-chart{height:340px}}

/* Premium preview theme */
:root{--bg:#f3f7ff;--panel:#fff;--panel-soft:#f8fbff;--ink:#10205f;--muted:#53627f;--line:#dce7fb;--line-soft:#edf3ff;--accent:#155cff;--accent-2:#0b4edc;--teal:#12b76a;--amber:#ff8a00;--purple:#6d4aff;--shadow:0 10px 26px rgba(33,83,170,.08)}
body{background:var(--bg);color:var(--ink);font-size:13px;letter-spacing:0}
.hdr{background:linear-gradient(180deg,#fbfdff,#eef5ff);color:var(--ink);padding:28px 34px 22px;text-align:left;border:1px solid var(--line);border-left:0;border-right:0;box-shadow:0 8px 24px rgba(33,83,170,.06)}
.hdr h1{position:relative;font-size:25px;line-height:1.25;margin:0 0 12px;letter-spacing:0;font-weight:800;color:#11247a}
.hdr h1:after{content:"";display:block;width:48px;height:4px;border-radius:999px;background:var(--accent);margin-top:12px}
.hdr p{opacity:.72;font-size:12px;letter-spacing:0;color:#50618a}
.kpi{grid-template-columns:repeat(auto-fit,minmax(176px,1fr));gap:14px;padding:18px 32px;background:#fff;border-bottom:1px solid var(--line)}
.kpi-card{position:relative;overflow:hidden;background:linear-gradient(180deg,#fff,#f8fbff);border:1px solid var(--line);border-radius:10px;padding:16px 17px;text-align:left;box-shadow:0 8px 20px rgba(33,83,170,.06)}
.kpi-card:before{content:"";position:absolute;left:0;right:0;top:0;height:3px;background:linear-gradient(90deg,var(--accent),#38bdf8)}
.kpi-card .l{font-size:11px;color:var(--muted);text-transform:none;letter-spacing:0;font-weight:600}
.kpi-card .v{font-size:25px;color:#10205f;line-height:1.15;margin:8px 0 5px;font-weight:800}
.kpi-card .s{font-size:11px;color:#8a92a3}
.kpi-top{display:flex;align-items:center;justify-content:space-between;gap:10px}
.kpi-top .lucide{width:18px;height:18px;color:var(--accent);stroke-width:2.1}
.tabs{gap:6px;padding:0 32px;border-bottom:1px solid var(--line);background:rgba(255,255,255,.96);box-shadow:0 1px 0 rgba(33,83,170,.04)}
.tab{padding:13px 14px 12px;border-bottom:2px solid transparent;color:#5f6675;font-size:13px;font-weight:600;letter-spacing:0;border-radius:6px 6px 0 0;display:flex;align-items:center;gap:6px}
.tab .lucide{width:14px;height:14px;stroke-width:2}
.tab:hover{color:var(--accent);background:#f2f7ff}
.tab.on{color:var(--accent);border-bottom-color:var(--accent);background:#f7fbff}
.body{padding:22px 32px 30px}
.g{grid-template-columns:repeat(auto-fit,minmax(520px,1fr));gap:18px;margin-bottom:18px}
.c{background:var(--panel);border:1px solid var(--line);border-radius:10px;padding:18px 20px 20px;box-shadow:var(--shadow)}
.c h3{display:flex;align-items:center;gap:10px;font-size:15px;line-height:1.3;margin-bottom:16px;color:var(--ink);border-left:0;padding-left:0;font-weight:740;letter-spacing:0}
.c h3:before{content:"";width:3px;height:18px;border-radius:2px;background:var(--accent);display:inline-block;flex:0 0 3px}
.card-title{display:flex;align-items:center;gap:8px;min-width:0;flex:1}
.card-title span{white-space:nowrap;overflow:hidden;text-overflow:ellipsis}
.card-title .lucide{width:16px;height:16px;color:var(--accent);stroke-width:2.1}
.help-btn{width:24px;height:24px;border:1px solid var(--line);border-radius:50%;background:#fff;color:#7a8292;display:inline-flex;align-items:center;justify-content:center;cursor:help;flex:0 0 24px;transition:.15s;position:relative}
.help-btn:hover{border-color:#b7ccff;color:var(--accent);background:#f2f7ff}
.help-btn .lucide{width:14px;height:14px;stroke-width:2}
.help-btn:focus-visible{outline:2px solid rgba(21,92,255,.35);outline-offset:2px}
.help-btn:hover:after,.help-btn:focus-visible:after{content:attr(data-tip);position:absolute;right:0;top:calc(100% + 8px);width:260px;background:#10205f;color:#fff;border-radius:8px;box-shadow:0 12px 28px rgba(16,32,95,.22);font-size:12px;font-weight:500;line-height:1.55;text-align:left;padding:8px 10px;z-index:300;white-space:normal}
.help-btn:hover:before,.help-btn:focus-visible:before{content:"";position:absolute;right:8px;top:calc(100% + 2px);border:6px solid transparent;border-bottom-color:#10205f;z-index:301}
.ch{height:360px}.ch.tall{height:500px}
.ch.large{height:460px}
.ch.wide{height:620px}
.fallback-chart{height:100%;overflow:auto;padding:4px 2px}
.fallback-row{display:grid;grid-template-columns:minmax(88px,150px) minmax(160px,1fr) auto;gap:10px;align-items:center;margin:7px 0;font-size:12px;color:#4f5665}
.fallback-row b{font-weight:600;white-space:nowrap;overflow:hidden;text-overflow:ellipsis}
.fallback-track{height:10px;background:#edf3ff;border-radius:999px;overflow:hidden}
.fallback-track i{display:block;height:100%;border-radius:999px;background:var(--accent)}
.fallback-note{font-size:12px;color:#687083;line-height:1.6;padding:12px;border:1px solid var(--line);border-radius:8px;background:#f8fbff}
.tip{background:linear-gradient(180deg,#fff,#f8fbff);border:1px solid var(--line);border-left:4px solid var(--accent);padding:12px 14px;margin-bottom:16px;border-radius:10px;color:#4f5665;box-shadow:0 5px 16px rgba(33,83,170,.05);line-height:1.65}
.tip b{color:var(--ink)}
table{font-size:12px}
th{background:#f7f8fa;color:#4b5260;border-bottom:1px solid var(--line);font-weight:700;padding:9px 10px}
td{padding:8px 10px;border-bottom:1px solid var(--line-soft);color:#303642}
tr:hover td{background:#f8fafc}
.st,.heat,.color-top3-wrap,.size-structure-wrap,.brand-tall-wrap{border-color:var(--line);border-radius:8px;background:#fff}
.legend{gap:12px;margin:2px 0 12px;color:#697184}
.legend i{width:11px;height:11px;border-radius:2px}
.ov-summary{gap:10px}
.ov-item{border:1px solid var(--line);border-radius:10px;background:linear-gradient(180deg,#fff,#f7fbff);padding:12px 13px;box-shadow:0 5px 14px rgba(33,83,170,.05)}
.ov-item span{color:#768092;font-weight:600}
.ov-item b{color:var(--ink);font-size:15px;font-weight:760}
.ov-item em{color:#3f4653;font-size:12px}
.ov-item small{color:#8b94a5}
.pf-head select,.pf-head input{border:1px solid var(--line);border-radius:6px;background:#fff;color:var(--ink)}
.pf-mini div{border:1px solid var(--line);border-radius:10px;background:linear-gradient(180deg,#fff,#f7fbff);box-shadow:none}
.pf-mini b{color:var(--ink)}
.pf-rec-row span,.cat-line em{background:#f7f8fa;border-color:var(--line);border-radius:999px;color:#4d5564}
.accept-note div,.range-note,.cat-note,.cat-group{border-color:var(--line);background:#fff;border-radius:8px}
.cat-group{box-shadow:0 6px 18px rgba(20,24,32,.04)}
.cat-group-head{background:#f7f8fa;border-bottom-color:var(--line)}
.cat-group-title{color:var(--ink)}
.cat-group-count{border-color:var(--line);background:#fff;color:#667085}
.range-track,.cat-mini-bar{background:#eef1f5;border-color:#e1e5eb}
.range-bar{background:var(--accent);box-shadow:0 3px 8px rgba(21,92,255,.16)}
.range-dot{background:var(--accent-2)}
.tag{border-radius:999px;font-weight:700}
.echart canvas{filter:saturate(.95)}
::-webkit-scrollbar{width:9px;height:9px}
::-webkit-scrollbar-track{background:#f1f3f6;border-radius:999px}
::-webkit-scrollbar-thumb{background:#c7ceda;border-radius:999px;border:2px solid #f1f3f6}
::-webkit-scrollbar-thumb:hover{background:#aeb7c6}
@media(max-width:760px){.pf-color-chart{height:260px}.accept-note{grid-template-columns:1fr}.range-note span{white-space:normal}.range-scale{display:none}.range-row{grid-template-columns:1fr}.range-cat{text-align:left}.range-value small{display:inline;margin-left:6px}.size-pie-layout{grid-template-columns:1fr}.size-pie-legend{max-height:none}}
</style></head><body>
<div class="hdr"><h1>新客户偏好分析</h1><p id="sub"></p></div>
<div class="kpi" id="kpi"></div>
<div class="tabs" id="tabs"></div>
<div class="body" id="bd"></div>
<script>
const D=${pyJsonDumps(R)};
const S=D.summary;
const P={amount:'#155cff',qty:'#12b76a',coverage:'#6d4aff',trend:'#ff8a00',muted:'#a6b1c6',soft:'#eaf1ff',ink:'#10205f'};
const C=['#155cff','#12b76a','#ff8a00','#6d4aff','#00a6ff','#10205f','#22c55e','#a855f7','#f59e0b','#64748b'];
var CH={},CFG={},OPT={};
function n(v){return Number(v)||0}
function fmt(v){v=n(v);var s=v<0?'-':'';v=Math.abs(v);return s+(v>=10000?(v/10000).toFixed(1)+'万':v>=1000?(v/1000).toFixed(1)+'千':v.toFixed?v.toFixed(0):v)}
function money(v){return '¥'+fmt(v)}
function pct(v,t){return (t?Math.round(n(v)/n(t)*1000)/10:0).toFixed(1)+'%'}
function positiveRows(rows,fields){return (rows||[]).filter(function(r){return fields.some(function(f){return n(r&&r[f])>0})})}
function valueRows(rows){return positiveRows(rows,['value','qty','amount','customers','金额','数量','客户数','orders','lines'])}
function esc(v){return String(v==null?'-':v).replace(/[&<>"']/g,function(m){return {'&':'&amp;','<':'&lt;','>':'&gt;','"':'&quot;',"'":'&#39;'}[m]})}
function axisTickFormatter(sc){return sc&&sc.ticks&&sc.ticks.callback?function(v){return sc.ticks.callback(v)}:undefined}
function legendOpt(cfg){var l=cfg.options&&cfg.options.plugins&&cfg.options.plugins.legend||{};if(l.display===false)return {show:false};if(l.position==='right')return {show:true,type:'scroll',orient:'vertical',right:0,top:36,bottom:12,itemWidth:20,itemHeight:8,textStyle:{color:'#687083',fontSize:11}};return {show:true,type:'scroll',bottom:0,left:'center',itemWidth:20,itemHeight:8,textStyle:{color:'#687083',fontSize:11}}}
function tooltipHtml(params,cfg){var ps=Array.isArray(params)?params:[params], cb=cfg.options&&cfg.options.plugins&&cfg.options.plugins.tooltip&&cfg.options.plugins.tooltip.callbacks||{}, labels=cfg.data.labels||[], dsets=cfg.data.datasets||[];var title=ps[0]&&ps[0].name!=null?ps[0].name:'';var html='<div style="font-weight:700;margin-bottom:6px;color:#fff">'+esc(title)+'</div>';var fake=[];ps.forEach(function(p){var ds=dsets[p.seriesIndex]||dsets[0]||{}, raw=Array.isArray(p.value)?p.value[p.value.length-1]:p.value, f={dataset:ds,datasetIndex:p.seriesIndex,dataIndex:p.dataIndex,raw:raw,label:labels[p.dataIndex]||p.name};fake.push(f);var txt=cb.label?cb.label(f):(p.seriesName?esc(p.seriesName)+': ':'')+(String(p.seriesName||'').includes('金额')||String(p.seriesName||'').includes('销售额')?money(raw):fmt(raw));if(Array.isArray(txt))txt=txt.join('<br>');html+='<div style="display:flex;align-items:center;gap:6px;line-height:1.7"><span style="display:inline-block;width:9px;height:9px;border-radius:2px;background:'+p.color+'"></span><span>'+esc(txt)+'</span></div>'});if(cb.afterBody){var extra=cb.afterBody(fake);if(extra){if(!Array.isArray(extra))extra=[extra];html+='<div style="margin-top:6px;padding-top:6px;border-top:1px solid rgba(255,255,255,.16);color:rgba(255,255,255,.82);line-height:1.6">'+extra.map(esc).join('<br>')+'</div>'}}return html}
function toEcharts(cfg){var type=cfg.type, labels=cfg.data.labels||[], dsets=cfg.data.datasets||[], opt=cfg.options||{}, plugins=opt.plugins||{}, tipOpt=plugins.tooltip||{}, isH=opt.indexAxis==='y', isPie=type==='pie'||type==='doughnut';var base={color:C,animationDuration:650,animationEasing:'cubicOut',tooltip:{trigger:tipOpt.trigger||(isPie?'item':'axis'),axisPointer:{type:type==='bar'?'shadow':'line'},backgroundColor:'rgba(16,32,95,.96)',borderWidth:0,padding:[9,11],textStyle:{color:'#fff',fontSize:12},formatter:tipOpt.formatter||function(p){return tooltipHtml(p,cfg)}} ,legend:legendOpt(cfg)};if(isPie){var ds=dsets[0]||{};return Object.assign(base,{series:[{name:ds.label||'',type:'pie',radius:type==='doughnut'?['48%','72%']:'70%',center:(plugins.legend&&plugins.legend.position)==='right'?['40%','50%']:['50%','48%'],avoidLabelOverlap:true,itemStyle:{borderColor:'#fff',borderWidth:2},label:{color:'#5f6675',fontSize:11},data:labels.map(function(l,i){return {name:l,value:(ds.data||[])[i],itemStyle:{color:Array.isArray(ds.backgroundColor)?ds.backgroundColor[i]:ds.backgroundColor}}})}]})}
var scales=opt.scales||{}, hasRight=dsets.some(function(ds){return ds.yAxisID==='y1'}), hasTop=dsets.some(function(ds){return ds.xAxisID==='x1'}), legendHidden=plugins.legend&&plugins.legend.display===false, grid=Object.assign({left:isH?64:hasRight?72:58,right:(isH&&hasTop)||(!isH&&hasRight)?100:36,top:hasTop||hasRight?54:28,bottom:legendHidden?44:70,containLabel:true},opt.grid||{});
function axisStyle(sc,name,pos){var isX=name==='x'||name==='x1';return {type:'value',position:pos,name:sc&&sc.title&&sc.title.display?sc.title.text:'',nameLocation:'middle',nameGap:isX?30:64,nameRotate:isX?0:(pos==='right'?-90:90),nameTextStyle:{color:'#687083',fontSize:11,padding:[0,0,0,0],align:'center'},min:sc&&sc.min!=null?sc.min:null,max:sc&&sc.max!=null?sc.max:null,axisLine:{show:false},axisTick:{show:false},axisLabel:{color:'#687083',fontSize:11,formatter:axisTickFormatter(sc)},splitLine:{lineStyle:{color:'rgba(16,32,95,.09)'}}}}
var xAxis,yAxis,series;
if(isH){xAxis=[axisStyle(scales.x,'x','bottom')];if(hasTop)xAxis.push(Object.assign(axisStyle(scales.x1,'x1','top'),{splitLine:{show:false}}));yAxis={type:'category',data:labels,inverse:true,axisLine:{show:false},axisTick:{show:false},axisLabel:{color:'#4f5665',fontSize:11,interval:0,width:118,overflow:'truncate',ellipsis:'...'},splitLine:{show:false}};series=dsets.map(function(ds,i){return {name:ds.label||'',type:type==='line'?'line':'bar',data:ds.data||[],xAxisIndex:ds.xAxisID==='x1'?1:0,stack:ds.stack,barMaxWidth:ds.barMaxWidth||18,itemStyle:{color:Array.isArray(ds.backgroundColor)?undefined:(ds.backgroundColor||C[i%C.length]),borderRadius:[0,4,4,0]},lineStyle:{color:ds.borderColor||ds.backgroundColor||C[i%C.length],width:2},areaStyle:ds.fill?{opacity:.12,color:ds.backgroundColor||ds.borderColor||C[i%C.length]}:undefined,smooth:!!ds.tension,symbolSize:type==='line'?6:0}})}
else{xAxis={type:'category',data:labels,axisLine:{lineStyle:{color:'rgba(16,32,95,.12)'}},axisTick:{show:false},axisLabel:{color:'#4f5665',fontSize:11,interval:0,rotate:labels.length>8?25:0},splitLine:{show:false}};yAxis=[axisStyle(scales.y,'y','left')];if(hasRight)yAxis.push(Object.assign(axisStyle(scales.y1,'y1','right'),{splitLine:{show:false}}));series=dsets.map(function(ds,i){var color=ds.borderColor||ds.backgroundColor||C[i%C.length];return {name:ds.label||'',type:type==='line'?'line':'bar',data:ds.data||[],yAxisIndex:ds.yAxisID==='y1'?1:0,stack:ds.stack,barMaxWidth:ds.barMaxWidth||18,itemStyle:{color:Array.isArray(ds.backgroundColor)?undefined:color,borderRadius:[4,4,0,0]},lineStyle:{color:color,width:2},areaStyle:ds.fill?{opacity:.12,color:ds.backgroundColor||color}:undefined,smooth:!!ds.tension,symbolSize:type==='line'?6:0}})}
series.forEach(function(s,i){var ds=dsets[i]||{};if(Array.isArray(ds.backgroundColor)){s.data=(ds.data||[]).map(function(v,j){return {value:v,itemStyle:{color:ds.backgroundColor[j]}}})}});
return Object.assign(base,{grid:grid,xAxis:xAxis,yAxis:yAxis,series:series})}
function chartReady(el){if(!el)return false;var r=el.getBoundingClientRect();return r.width>20&&r.height>20}
function drawCfg(id,cfg){var e=document.getElementById(id);if(!e||!window.echarts||!chartReady(e))return null;if(CH[id])CH[id].dispose();CH[id]=echarts.init(e,null,{renderer:'canvas'});CH[id].setOption(toEcharts(cfg),true);return CH[id]}
function drawOpt(id,entry){var e=document.getElementById(id);if(!e||!window.echarts||!chartReady(e))return null;if(CH[id])CH[id].dispose();CH[id]=echarts.init(e,null,{renderer:'canvas'});CH[id].setOption(entry.option,true);return CH[id]}
function resizeCharts(){Object.keys(CFG).forEach(function(k){try{if(!CH[k])drawCfg(k,CFG[k])}catch(e){}});Object.keys(OPT).forEach(function(k){try{if(!CH[k])drawOpt(k,OPT[k])}catch(e){}});Object.keys(CH).forEach(function(k){try{CH[k].resize()}catch(e){}})}
function fallbackChart(id,cfg){var e=document.getElementById(id);if(!e)return;var labels=(cfg.data&&cfg.data.labels)||[],ds=(cfg.data&&cfg.data.datasets&&cfg.data.datasets[0])||{},data=ds.data||[],max=Math.max.apply(null,data.map(n).concat([1]));var rows=labels.map(function(label,i){var val=n(data[i]),w=Math.max(1,Math.min(100,val/max*100));var color=Array.isArray(ds.backgroundColor)?ds.backgroundColor[i]:(ds.backgroundColor||P.amount);return '<div class="fallback-row" title="'+esc(label)+'"><b>'+esc(label)+'</b><div class="fallback-track"><i style="width:'+w+'%;background:'+color+'"></i></div><span>'+((ds.label||'').indexOf('金额')>=0?money(val):fmt(val))+'</span></div>'}).join('');e.innerHTML='<div class="fallback-chart">'+(rows||'<div class="fallback-note">当前图表暂无可展示数据。</div>')+'</div>'}
function mk(id,cfg){CFG[id]=cfg;delete OPT[id];var e=document.getElementById(id);if(!e)return null;if(CH[id]){CH[id].dispose();delete CH[id]}if(!window.echarts){fallbackChart(id,cfg);return null}var chart=drawCfg(id,cfg);setTimeout(resizeCharts,30);return chart}
function mkOption(id,option,fallback){OPT[id]={option:option,fallback:fallback};delete CFG[id];var e=document.getElementById(id);if(!e)return null;if(CH[id]){CH[id].dispose();delete CH[id]}if(!window.echarts){if(fallback)fallback();return null}var chart=drawOpt(id,OPT[id]);setTimeout(resizeCharts,30);return chart}
function iconSvg(name){var icons={
'badge-dollar-sign':'<path d="M7 3h10l3 3v12l-3 3H7l-3-3V6z"/><path d="M12 7v10"/><path d="M9.5 9.5c.8-.7 4.2-.8 5 .4.7 1-.2 2-2.6 2.1-2.4.1-3.3 1.1-2.5 2.1.9 1.2 4.2 1.1 5.1.3"/>',
'package-check':'<path d="M21 8l-9-5-9 5 9 5z"/><path d="M3 8v8l9 5 9-5V8"/><path d="M12 13v8"/><path d="M16 15l2 2 4-4"/>',
'users':'<path d="M16 21v-2a4 4 0 0 0-4-4H6a4 4 0 0 0-4 4v2"/><circle cx="9" cy="7" r="4"/><path d="M22 21v-2a4 4 0 0 0-3-3.9"/><path d="M16 3.1a4 4 0 0 1 0 7.8"/>',
'receipt':'<path d="M5 3v18l2-1 2 1 2-1 2 1 2-1 2 1 2-1V3z"/><path d="M8 7h8"/><path d="M8 11h8"/><path d="M8 15h5"/>',
'badge':'<path d="M7 3h10l4 4v10l-4 4H7l-4-4V7z"/><path d="M8 12h8"/><path d="M9 16h6"/>',
'layout-dashboard':'<rect x="3" y="3" width="7" height="8" rx="1"/><rect x="14" y="3" width="7" height="5" rx="1"/><rect x="14" y="12" width="7" height="9" rx="1"/><rect x="3" y="15" width="7" height="6" rx="1"/>',
'layers-3':'<path d="M12 2l9 5-9 5-9-5z"/><path d="M3 12l9 5 9-5"/><path d="M3 17l9 5 9-5"/>',
'palette':'<path d="M12 3a9 9 0 0 0 0 18h1.5a2.5 2.5 0 0 0 0-5H12a1.5 1.5 0 0 1 0-3h1a8 8 0 0 0 8-8 2 2 0 0 0-2-2z"/><circle cx="7.5" cy="10" r="1"/><circle cx="10" cy="6.5" r="1"/><circle cx="14" cy="6.5" r="1"/><circle cx="16.5" cy="10" r="1"/>',
'ruler':'<path d="M3 17l14-14 4 4L7 21z"/><path d="M14 6l2 2"/><path d="M11 9l2 2"/><path d="M8 12l2 2"/><path d="M5 15l2 2"/>',
'calendar-range':'<rect x="3" y="4" width="18" height="17" rx="2"/><path d="M16 2v4"/><path d="M8 2v4"/><path d="M3 10h18"/><path d="M7 14h4"/><path d="M13 17h4"/>',
'user-round-search':'<circle cx="10" cy="8" r="4"/><path d="M2 21a8 8 0 0 1 12-7"/><circle cx="17" cy="17" r="3"/><path d="M21 21l-2-2"/>',
'book-open':'<path d="M2 4h7a4 4 0 0 1 4 4v13a3 3 0 0 0-3-3H2z"/><path d="M22 4h-7a4 4 0 0 0-4 4v13a3 3 0 0 1 3-3h8z"/>',
'store':'<path d="M4 10h16l-1-6H5z"/><path d="M5 10v10h14V10"/><path d="M8 20v-6h4v6"/><path d="M15 14h2"/>',
'line-chart':'<path d="M3 3v18h18"/><path d="M7 15l4-4 3 3 5-7"/>',
'bar-chart-3':'<path d="M3 3v18h18"/><path d="M7 16V9"/><path d="M12 16V5"/><path d="M17 16v-3"/>',
'sparkles':'<path d="M12 3l1.6 4.4L18 9l-4.4 1.6L12 15l-1.6-4.4L6 9l4.4-1.6z"/><path d="M19 14l.8 2.2L22 17l-2.2.8L19 20l-.8-2.2L16 17l2.2-.8z"/><path d="M5 15l.8 2.2L8 18l-2.2.8L5 21l-.8-2.2L2 18l2.2-.8z"/>',
'info':'<circle cx="12" cy="12" r="9"/><path d="M12 10v6"/><path d="M12 7h.01"/>',
'chart-no-axes-combined':'<path d="M4 19V5"/><path d="M4 19h16"/><path d="M8 16v-4"/><path d="M12 16V8"/><path d="M16 16v-6"/><path d="M8 8l4-3 4 4 4-6"/>'
};var body=icons[name]||icons['chart-no-axes-combined'];return '<svg class="lucide" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round" aria-hidden="true">'+body+'</svg>'}
function hydrateUi(){document.querySelectorAll('[data-lucide]').forEach(function(el){var name=el.getAttribute('data-lucide')||'chart-no-axes-combined';el.innerHTML=iconSvg(name)})}
function sum(a,k){return a.reduce(function(x,r){return x+n(r[k])},0)}
function topRows(a,cnt,k){return a.slice().sort(function(x,y){return n(y[k])-n(x[k])}).slice(0,cnt)}
function iconForTitle(t){if(t.indexOf('字段')>=0||t.indexOf('计算说明')>=0)return 'book-open';if(t.indexOf('客户')>=0)return 'users';if(t.indexOf('颜色')>=0)return 'palette';if(t.indexOf('尺码')>=0||t.indexOf('码数')>=0)return 'ruler';if(t.indexOf('年份')>=0||t.indexOf('季节')>=0||t.indexOf('订货会')>=0)return 'calendar-range';if(t.indexOf('品牌')>=0||t.indexOf('设计师')>=0)return 'badge';if(t.indexOf('价格')>=0||t.indexOf('价位')>=0)return 'badge-dollar-sign';if(t.indexOf('店铺')>=0)return 'store';if(t.indexOf('趋势')>=0||t.indexOf('月度')>=0)return 'line-chart';if(t.indexOf('销量')>=0||t.indexOf('金额')>=0)return 'bar-chart-3';if(t.indexOf('推荐')>=0||t.indexOf('建议')>=0)return 'sparkles';return 'chart-no-axes-combined'}
function helpForTitle(t){if(t.indexOf('走势')>=0||t.indexOf('趋势')>=0)return '看当前报表时间段内的变化方向，不做跨报表或自然月强行环比。';if(t.indexOf('Top20客户')>=0)return '先按客户净销售金额筛出Top20客户，再汇总这些客户所有订单，不展示客户名称。';if(t.indexOf('年份')>=0||t.indexOf('季节')>=0||t.indexOf('订货会')>=0)return '年份字段包含“订货会”的记录归为订货会成交，其余归为普通款期成交。客户不拆分，只拆成交记录。';if(t.indexOf('品类 - 设计师品牌贡献值Top1')>=0)return '每个品类只展示贡献金额最高的设计师品牌，用于判断该品类主要由哪个设计师品牌创造销售价值。';if(t.indexOf('品牌 - 设计师贡献金额')>=0)return '按品牌分组展示设计师贡献金额，条形越长说明该设计师在该品牌下创造的成交金额越高。';if(t.indexOf('Top1')>=0)return '每个分组只展示销量或金额最高的第一项，用于快速判断主推方向。';if(t.indexOf('备货优先级')>=0)return '综合销量、客户覆盖、近期增长和稳定性生成，用于判断多备、常备或控量。';if(t.indexOf('价格')>=0||t.indexOf('价位')>=0)return '金额使用净销售金额，数量使用净销售量，单价按成交金额除以成交数量计算。';if(t.indexOf('客户覆盖')>=0)return '客户数按当前报表内发生过拿货的客户去重统计。';if(t.indexOf('颜色')>=0)return '颜色按源表颜色名称统计，空值归为未标记，图案类单独归类。';if(t.indexOf('尺码')>=0||t.indexOf('码数')>=0)return '尺码展示全部尺码，不因销量低而隐藏。';return '鼠标悬停图表可查看成交金额、数量、客户覆盖等明细。'}
function card(t,b,full){var help=helpForTitle(t);return '<div class="c'+(full?' full':'')+'"><h3><span class="card-title"><i data-lucide="'+iconForTitle(t)+'"></i><span>'+esc(t)+'</span></span><button class="help-btn" data-tip="'+esc(help)+'" aria-label="图表说明：'+esc(help)+'"><i data-lucide="info"></i></button></h3>'+b+'</div>'}
function chart(id,tall){return '<div class="ch'+(tall?' tall':'')+'"><div id="'+id+'" class="echart"></div></div>'}
function largeChart(id){return '<div class="ch large"><div id="'+id+'" class="echart"></div></div>'}
function wideChart(id){return '<div class="ch wide"><div id="'+id+'" class="echart"></div></div>'}
function scrollChart(id){return '<div class="color-all-wrap"><div class="ch"><div id="'+id+'" class="echart"></div></div></div>'}
function top3ScrollChart(id){return '<div class="color-top3-wrap"><div class="ch"><div id="'+id+'" class="echart"></div></div></div>'}
function customerColorScrollChart(id){return '<div class="customer-color-wrap"><div class="ch"><div id="'+id+'" class="echart"></div></div></div>'}
function sizePieChart(id){return '<div class="size-pie-layout"><div class="ch"><div id="'+id+'" class="echart"></div></div><div class="size-pie-legend" id="sizePieLegend"></div></div>'}
function sizeStructureChart(id){return '<div class="size-structure-wrap"><div class="ch"><div id="'+id+'" class="echart"></div></div></div>'}
function brandTallChart(id){return '<div class="brand-tall-wrap"><div class="ch"><div id="'+id+'" class="echart"></div></div></div>'}
function table(h,rows){return '<div class="st"><table><tr>'+h.map(function(x){return '<th>'+x+'</th>'}).join('')+'</tr>'+rows.join('')+'</table></div>'}
function tr(cells){return '<tr>'+cells.map(function(x){return '<td>'+x+'</td>'}).join('')+'</tr>'}
function overviewData(){return D.overview_analysis||{summary_items:[],monthly_sales:[],shop_sales:[],repeat_products:[],order_count:0,detail_rows:0}}
function overviewSummaryCard(){var items=positiveRows(overviewData().summary_items||[],['qty','amount','customers']);return '<div class="ov-summary">'+items.map(function(r){var sw=r.hex?'<i style="background:'+esc(r.hex)+'"></i>':'';return '<div class="ov-item"><span>'+esc(r.label)+'</span><b>'+sw+esc(r.name)+'</b><em>'+fmt(r.qty)+'件 / '+money(r.amount)+'</em><small>'+esc(r.note)+' · 覆盖 '+fmt(r.customers)+' 客户</small></div>'}).join('')+'</div>'}
function typeCls(v){return v==='追新型'?'t-new':v==='折扣型'?'t-old':'t-bal'}
function brandRows(){var by={};D.brand_customer.forEach(function(r){var c=r['客户'];(by[c]=by[c]||[]).push(r)});return Object.keys(by).map(function(c){var arr=topRows(by[c],3,'金额');var total=sum(by[c],'金额')||1;return {客户:c,品牌:arr.map(function(r){return r['品牌']+' '+money(r['金额'])}),集中度:pct(arr[0]?arr[0]['金额']:0,total)}})}
function sensitivityRows(){var by={};D.sensitivity.forEach(function(r){var c=r['客户'];var o=by[c]||(by[c]={客户:c,总:0});o[r['类型']]=n(o[r['类型']])+n(r['金额']);o.总+=n(r['金额'])});return Object.keys(by).map(function(k){var r=by[k];var np=pct(r['当季新品']||0,r.总);var order=pct(r['订货会']||0,r.总);var old=pct(r['往季/折扣']||0,r.总);var tp=n(r['当季新品'])/Math.max(n(r.总),1)*100>60?'追新型':n(r['当季新品'])/Math.max(n(r.总),1)*100<40?'折扣型':'均衡型';return {客户:k,总:r.总,当季:r['当季新品']||0,订货会:r['订货会']||0,折扣:r['往季/折扣']||0,新品占比:np,订货会占比:order,折扣占比:old,类型:tp}}).sort(function(a,b){return b.总-a.总})}
function catAdvice(r){var s=n(r['金额'])/Math.max(n(S.amount),1)*100;var cp=D.cat_price.find(function(x){return x['分类']===r['分类']});var p=cp?n(cp['平均']):0;if(s>=10)return '主推引流 + 稳定备货';if(p>=180)return '高客单重点陈列';if(n(r['客户数'])>=1000)return '广覆盖连带推荐';return '小众款控量测试'}
function catData(){return D.category_preference_analysis||{trend_period_type:'',trend_periods:[],category_trend:[],categories:[],cooccurrence:[]}}
function catClassColor(v){return v==='稳定主力'?P.qty:v==='少数客户拉动'?P.trend:v==='广覆盖连带'?P.coverage:v==='潜力观察'?'#00a6ff':P.muted}
function catClassLegend(){var items=['稳定主力','少数客户拉动','广覆盖连带','潜力观察','控量测试'];return '<div class="legend cat-legend">'+items.map(function(x){return '<span><i style="background:'+catClassColor(x)+'"></i>'+x+'</span>'}).join('')+'</div>'}
function catTooltip(r){return ['成交金额 '+money(r.amount),'成交数量 '+fmt(r.qty)+'件','客户覆盖 '+fmt(r.customers)+'人','所有客户占比 '+pct(r.customer_share,1),'订单数 '+fmt(r.orders),'常搭品类 '+esc(r.companion||'-')+' '+(r.companion_rate||0)+'%']}
function shopBestCategoryTooltip(r){return ['店铺 '+esc(r.shop||'-'),'Top1品类 '+esc(r.category||'-'),'成交金额 '+money(r.amount),'成交数量 '+fmt(r.qty)+'件','客户数 '+fmt(r.customers)+'人']}
function categoryCoverageNote(){return '<div class="cat-note"><span><b>绿条</b>表示买过该品类的客户数。</span><span><b>橙条</b>表示该品类覆盖了全部客户中的多少比例。</span><span>比例越高，说明这个品类越适合作为广泛推荐入口。</span></div>'}
function categoryTrendNote(){var ca=catData();return '<div class="cat-note"><span>只使用当前报表时间段内的数据。</span><span>当前按<b>'+esc(ca.trend_period_type||'-')+'</b>汇总，不做自然月环比或跨报表对比。</span><span>趋势图展示成交金额 Top10 主力品类，避免全部品类堆叠后无法阅读。</span></div>'}
function trendAmountMap(){var m={};(catData().category_trend||[]).forEach(function(r){m[r.category+'|'+r.period]=r});return m}
function categoryAdviceBoard(){var rows=positiveRows(catData().categories||[],['qty','amount','customers']).slice().sort(function(a,b){return n(b.recommend_score)-n(a.recommend_score)});var groups=[{name:'主推备货',desc:'作为日常主推入口，保证主销款深度，并用于带动搭配销售。',color:P.qty},{name:'连带搭配',desc:'覆盖客户多，适合做进店搭配和加购推荐，提高连带件数。',color:P.coverage},{name:'定向推荐',desc:'优先推荐给高贡献客户，备货不要铺太散，适合做定向补货。',color:P.trend},{name:'测款放量',desc:'报表期内走势上升，可保留动销款并小幅增加相近款测试。',color:'#00a6ff'},{name:'控量测试',desc:'少量陈列，优先选择卖点清晰、退换风险低的款。',color:P.muted}];return '<div class="cat-advice">'+groups.map(function(g){var items=rows.filter(function(r){return (r.sale_focus||'控量测试')===g.name});var list=items.length?items.map(function(r){var w=Math.max(0,Math.min(100,n(r.recommend_score)));return '<div class="cat-line" title="'+esc(r.category+'：推荐分 '+r.recommend_score+'；客户覆盖 '+fmt(r.customers)+'人；所有客户占比 '+pct(r.customer_share,1)+'；常搭 '+(r.companion||'-'))+'"><b>'+esc(r.category)+'</b><span>'+pct(r.customer_share,1)+'</span><em>'+esc(r.companion&&r.companion!=='-'?'搭 '+r.companion:'无常搭')+'</em><div class="cat-mini-bar"><i style="width:'+w+'%;background:'+g.color+'"></i></div></div>'}).join(''):'<div class="cat-empty">当前无品类归入此类</div>';return '<div class="cat-group"><div class="cat-group-head"><div class="cat-group-title"><i style="background:'+g.color+'"></i>'+esc(g.name)+'</div><span class="cat-group-count">'+items.length+'个品类</span></div><div class="cat-group-desc">'+esc(g.desc)+'</div><div class="cat-list">'+list+'</div></div>'}).join('')+'</div>'}
function colorData(){return D.color_preference_analysis||{colors:[],category_colors:[],category_structure:[],top_color_category_distribution:[],top_color_brand_distribution:[],top_customer_colors:[],monthly_trend:[],months:[],trend_colors:[],has_prior_period_data:false}}
function stockTagColor(v){return v==='重点备货'?P.qty:v==='常规备货'?'#00a6ff':v==='趋势加量'?P.amount:v==='少量试水'?P.trend:P.muted}
function stockGroupColor(v){return v==='稳定常备色'?P.qty:v==='近期趋势色'?P.amount:P.muted}
function stockGroupLegend(){var co=colorData(),cuts=co.color_cutoffs||{},high=cuts.high_qty?fmt(cuts.high_qty):'-',mid=cuts.mid_qty?fmt(cuts.mid_qty):'-',growth=cuts.trend_growth_rate||20;return '<div class="legend cat-legend"><span><i style="background:'+stockGroupColor('稳定常备色')+'"></i>稳定常备色：销量 ≥ '+high+'件，建议多备</span><span><i style="background:'+stockGroupColor('近期趋势色')+'"></i>近期趋势色：销量 ≥ '+mid+'件，近30天较前30天增长 ≥ '+growth+'%，可加量</span><span><i style="background:'+stockGroupColor('控制备货色')+'"></i>控制备货色：未进入主力销量区或近期未走高，少量备货</span></div>'}
function colorBar(r){return r.hex||'#d9d9d9'}
function colorTitle(r){var out=['净销售量 '+fmt(r.qty)+'件','净销售金额 '+money(r.amount),'客户数 '+fmt(r.customers)+'人'];if(r.coverage_categories!=null)out.push('覆盖品类 '+fmt(r.coverage_categories)+'类');return out.concat(['销量分层 '+esc(r.group||'-'),'销量区间 '+esc(r.qty_band||'-'),'近30天 '+fmt(r.recent30_qty)+'件 / 前30天 '+fmt(r.previous30_qty)+'件','近30天增长 '+(r.recent_growth_rate||0)+'%','备货分 '+(r.score||0),'稳定性 '+(r.stability||0)+'%'])}
function colorStructureRows(){var rows=[];(colorData().category_structure||[]).forEach(function(cat){(cat.colors||[]).forEach(function(r){rows.push({category:cat.category,rank:n(r.rank),color:r.color,qty:n(r.qty),share:n(r.share),label:cat.category+' / '+r.color})})});return rows}
function colorRankColor(rank){return n(rank)===1?P.qty:n(rank)===2?P.amount:'#00a6ff'}
function colorStructureTooltipRow(r){return ['品类 '+esc(r.category||'-'),'颜色 '+esc(r.color||'-'),'排名 Top'+(r.rank||'-'),'净销售量 '+fmt(r.qty)+'件','品类内占比 '+pct(r.share,1)]}
function customerColorTooltip(r){return ['颜色 '+esc(r.color||'-'),'排名 Top'+(r.rank||'-'),'净销售量 '+fmt(r.qty)+'件','净销售金额 '+money(r.amount),'覆盖销售额Top20客户 '+fmt(r.customers)+'人','销售额Top20客户内销量占比 '+pct(r.qty_share,1),'销售额Top20客户内金额占比 '+pct(r.amount_share,1)]}
function colorBreakdownPalette(i){var cs=['#155cff','#12b76a','#ff8a00','#6d4aff','#00a6ff','#10205f','#22c55e','#a855f7','#f59e0b','#64748b','#ec4899','#14b8a6'];return cs[i%cs.length]}
function colorBreakdownData(rawRows,limit){var rows=(rawRows||[]).filter(function(r){return n(r.total_qty)>0}).slice(0,10),totals={};rows.forEach(function(r){(r.items||[]).forEach(function(it){if(n(it.qty)>0)totals[it.name]=n(totals[it.name])+n(it.qty)})});var ordered=Object.keys(totals).sort(function(a,b){return n(totals[b])-n(totals[a])});var main=ordered.slice(0,limit||8),hasOther=ordered.length>main.length,names=hasOther?main.concat(['其他']):main;var mapped=rows.map(function(r){var by={},other={name:'其他',qty:0,amount:0,customers:0,share:0};(r.items||[]).forEach(function(it){if(main.indexOf(it.name)>=0)by[it.name]=it;else{other.qty+=n(it.qty);other.amount+=n(it.amount);other.customers+=n(it.customers)}});if(hasOther&&other.qty>0){other.share=n(r.total_qty)?other.qty/n(r.total_qty):0;by['其他']=other}return Object.assign({},r,{items:names.map(function(name){return Object.assign({name:name,qty:0,amount:0,customers:0,share:0},by[name]||{})})})});return {rows:mapped,names:names}}
function colorBreakdownItem(row,name){return (row.items||[]).find(function(x){return x.name===name})||{}}
function renderColorBreakdownChart(id,rawRows){var d=colorBreakdownData(rawRows,8),rows=d.rows,names=d.names;if(!rows.length||!names.length){mk(id,{type:'bar',data:{labels:[],datasets:[]},options:{responsive:true,maintainAspectRatio:false,plugins:{legend:{display:false}}}});return}mk(id,{type:'bar',data:{labels:rows.map(function(r){return r.color}),datasets:names.map(function(name,i){return {label:name,itemName:name,data:rows.map(function(r){return n(colorBreakdownItem(r,name).qty)}),backgroundColor:colorBreakdownPalette(i),stack:'color',barMaxWidth:28}})},options:{responsive:true,maintainAspectRatio:false,indexAxis:'y',grid:{left:74,right:46,top:44,bottom:86},plugins:{legend:{position:'bottom'},tooltip:{callbacks:{label:function(c){var row=rows[c.dataIndex]||{},item=colorBreakdownItem(row,c.dataset.itemName);return c.dataset.label+': '+fmt(c.raw)+'件 / 占该颜色 '+pct(c.raw,row.total_qty)+' / '+money(item.amount)},afterBody:function(ctx){var row=rows[ctx[0].dataIndex]||{};return ['颜色总销量 '+fmt(row.total_qty)+'件','颜色总金额 '+money(row.total_amount),'客户数 '+fmt(row.customers)+'人']}}}},scales:{x:{title:{display:true,text:'净销售量'},ticks:{callback:function(v){return fmt(v)}}},y:{ticks:{autoSkip:false,font:{size:11}}}}}})}
function sizeData(){return D.size_preference_analysis||{sizes:[],category_sizes:[],category_structure:[],top_customer_sizes:[],monthly_trend:[],months:[],trend_sizes:[],size_cutoffs:{},has_recent30_period_data:false}}
function sizeGroupColor(v){return v==='稳定常备码'?P.qty:v==='近期趋势码'?P.amount:P.muted}
function sizeStockGroupLegend(){var sp=sizeData(),cuts=sp.size_cutoffs||{},high=cuts.high_qty?fmt(cuts.high_qty):'-',mid=cuts.mid_qty?fmt(cuts.mid_qty):'-',growth=cuts.trend_growth_rate||20;return '<div class="legend cat-legend"><span><i style="background:'+sizeGroupColor('稳定常备码')+'"></i>稳定常备码：销量 ≥ '+high+'件，建议作为主备尺码</span><span><i style="background:'+sizeGroupColor('近期趋势码')+'"></i>近期趋势码：销量 ≥ '+mid+'件，近30天较前30天增长 ≥ '+growth+'%，可适当加量</span><span><i style="background:'+sizeGroupColor('控制备货码')+'"></i>控制备货码：未进入主力销量区或近期未走高，控制备货深度</span></div>'}
function sizeTitle(r){var out=['净销售量 '+fmt(r.qty)+'件','净销售金额 '+money(r.amount),'客户数 '+fmt(r.customers)+'人'];if(r.coverage_categories!=null)out.push('覆盖品类 '+fmt(r.coverage_categories)+'类');return out.concat(['销量分层 '+esc(r.group||'-'),'销量区间 '+esc(r.qty_band||'-'),'近30天 '+fmt(r.recent30_qty)+'件 / 前30天 '+fmt(r.previous30_qty)+'件','近30天增长 '+(r.recent_growth_rate||0)+'%','备货分 '+(r.score||0),'稳定性 '+(r.stability||0)+'%'])}
function sizePieLegend(rows){return rows.map(function(r,i){return '<span><i style="background:'+C[i%C.length]+'"></i>'+esc(r.size)+'</span>'}).join('')}
function sizeStructureTooltip(categoryRow,sizeRow){return ['品类 '+esc(categoryRow.category||'-'),'尺码 '+esc(sizeRow.size||'-'),'净销售量 '+fmt(sizeRow.qty)+'件','净销售金额 '+money(sizeRow.amount),'品类内占比 '+pct(sizeRow.share,1)]}
function sizeTopCategoryBoard(){
var sp=sizeData(), sizes=(sp.sizes||[]).slice().sort(function(a,b){return n(b.qty)-n(a.qty)}).slice(0,5), cats=sp.category_structure||[];
return '<div class="size-top-board">'+sizes.map(function(sizeRow,i){
  var rows=cats.map(function(cat){
    var row=(cat.sizes||[]).find(function(x){return x.size===sizeRow.size})||{};
    return {category:cat.category,qty:n(row.qty),amount:n(row.amount),share:n(row.share)};
  }).filter(function(r){return r.qty>0}).sort(function(a,b){return b.qty-a.qty||b.amount-a.amount});
  var max=Math.max.apply(null,rows.map(function(r){return r.qty}).concat([1]));
  var total=rows.reduce(function(s,r){return s+r.qty},0);
  var list=rows.length?rows.map(function(r){
    var w=Math.max(2,Math.min(100,r.qty/max*100));
    return '<div class="size-top-row" title="'+esc(sizeRow.size+' / '+r.category+'：'+fmt(r.qty)+'件，'+money(r.amount)+'，该尺码内占比 '+pct(r.qty,total))+'"><strong>'+esc(r.category)+'</strong><em>'+fmt(r.qty)+'件</em><div class="size-top-bar"><i style="width:'+w+'%;background:'+C[i%C.length]+'"></i></div></div>';
  }).join(''):'<div class="size-top-empty">该尺码暂无品类成交</div>';
  return '<div class="size-top-col"><div class="size-top-head"><b>'+esc(sizeRow.size)+'</b><span>'+fmt(sizeRow.qty)+'件 / '+money(sizeRow.amount)+'</span></div><div class="size-top-list">'+list+'</div></div>';
}).join('')+'</div>';
}
function customerSizeTooltip(r){return ['尺码 '+esc(r.size||'-'),'排名 Top'+(r.rank||'-'),'净销售量 '+fmt(r.qty)+'件','净销售金额 '+money(r.amount),'覆盖销售额Top20客户 '+fmt(r.customers)+'人','销售额Top20客户内销量占比 '+pct(r.qty_share,1),'销售额Top20客户内金额占比 '+pct(r.amount_share,1)]}
function sizeAdvice(i){return i<2?'主备尺码':i<5?'次备尺码':'少量补充'}
function priceAdvice(r,i){if(i===0)return '主成交带,保证款量';if(n(r['金额'])>=1000000)return '利润带,强化搭配销售';return '长尾带,谨慎备货'}
function priceTagCls(v){return v==='低价敏感型'?'p-low':v==='中价稳定型'?'p-mid':v==='高价接受型'?'p-high':v==='价格波动型'?'p-wave':'p-lack'}
function priceRange(a,b){return '¥'+n(a).toFixed(1)+'-'+n(b).toFixed(1)}
function priceTitle(r){return r.customer+' / '+r.category+'\\n成交金额 '+money(r.amount)+'\\n成交数量 '+fmt(r.qty)+'\\n订单数 '+r.orders+'\\n加权均价 ¥'+r.weighted_avg+'\\n中位数 ¥'+r.median+'\\n主力价格带 '+r.main_band+'\\n价格接受区间 '+priceRange(r.p25,r.p75)+'\\n标签 '+r.tag}
function priceTags(){return ['低价敏感型','中价稳定型','高价接受型','价格波动型','样本不足']}
function priceTagColor(v){return v==='低价敏感型'?'#60a5fa':v==='中价稳定型'?P.qty:v==='高价接受型'?P.trend:v==='价格波动型'?P.coverage:P.muted}
function priceLegend(){return '<div class="legend">'+priceTags().map(function(t){return '<span><i style="background:'+priceTagColor(t)+'"></i>'+t+'</span>'}).join('')+'</div>'}
function priceSensitivityByCategory(limit){var p=D.price_analysis||{categories:[],rows:[]};var cats=p.categories.slice(0,limit).map(function(r){return r.category});var by={};cats.forEach(function(c){by[c]={}});p.rows.forEach(function(r){if(by[r.category])by[r.category][r.tag]=n(by[r.category][r.tag])+1});return {cats:cats,datasets:priceTags().map(function(t){return {label:t,data:cats.map(function(c){return by[c][t]||0}),backgroundColor:priceTagColor(t),stack:'s'}})}}
function renderPriceHeatmap(){var p=D.price_analysis;if(!p)return;var el=document.getElementById('priceHeatmap');if(!el)return;var cats=p.categories.slice(0,10).map(function(r){return r.category});var customers=p.matrix_customers.slice(0,18);var map={};p.rows.forEach(function(r){map[r.customer+'|'+r.category]=r});var html=priceLegend()+'<div class="heat"><table><tr><th>客户</th>'+cats.map(function(x){return '<th>'+esc(x)+'</th>'}).join('')+'</tr>';customers.forEach(function(cu){html+='<tr><td><b>'+esc(cu)+'</b></td>'+cats.map(function(ca){var r=map[cu+'|'+ca];if(!r)return '<td><span class="pcell p-lack">-<small>无记录</small></span></td>';return '<td><span class="pcell '+priceTagCls(r.tag)+'" title="'+esc(priceTitle(r))+'">'+esc(r.main_band)+'<small>¥'+r.weighted_avg+'</small></span></td>'}).join('')+'</tr>'});html+='</table></div>';el.innerHTML=html}
function acceptData(){return D.price_acceptance_analysis||{low_cutoff:0,high_cutoff:0,overall:{},kpi:{},bands:[],segments:[],categories:[],recommendations:[]}}
function acceptPct(v){return (n(v)*100).toFixed(1)+'%'}
function percentHtml(v){var s=String(v==null?'0.0%':v).replace(/%/g,'');return '<span style="white-space:nowrap;font-family:Arial,Helvetica,sans-serif;letter-spacing:0;color:#fff"><span style="font-size:20px;font-weight:900;line-height:1.1">'+esc(s)+'</span><span style="font-size:20px;font-weight:900;margin-left:2px;line-height:1.1">%</span></span>'}
function acceptSegColor(v){return v==='低价带'?'#60a5fa':v==='主流价带'?P.qty:P.trend}
function acceptSegmentDesc(name){var ac=acceptData();if(name==='低价带')return '成交单价 < ¥'+ac.low_cutoff;if(name==='主流价带')return '¥'+ac.low_cutoff+' ≤ 成交单价 < ¥'+ac.high_cutoff;return '成交单价 ≥ ¥'+ac.high_cutoff}
function acceptSegmentNote(){var rows=positiveRows(acceptData().segments||[],['amount','qty','orders','lines']);return '<div class="accept-note">'+rows.map(function(r){return '<div><b><i style="background:'+acceptSegColor(r.name)+'"></i>'+esc(r.name)+'</b><span>'+acceptSegmentDesc(r.name)+'</span><span>金额占 '+acceptPct(r.amount_share)+' / 数量占 '+acceptPct(r.qty_share)+'</span></div>'}).join('')+'</div>'}
function acceptCategoryStructureNote(){var ac=acceptData();return '<div class="accept-note"><div><b><i style="background:'+acceptSegColor('低价带')+'"></i>低价带</b><span>'+acceptSegmentDesc('低价带')+'</span><span>看该品类是否主要靠低价成交拉动。</span></div><div><b><i style="background:'+acceptSegColor('主流价带')+'"></i>主流价带</b><span>'+acceptSegmentDesc('主流价带')+'</span><span>看该品类当前稳定成交的主价格区。</span></div><div><b><i style="background:'+acceptSegColor('高价带')+'"></i>高价带</b><span>'+acceptSegmentDesc('高价带')+'</span><span>看该品类能否承接更高单价货品。</span></div></div>'}
function acceptRangeData(limit){var rows=positiveRows(acceptData().recommendations||[],['amount','qty']).slice(0,limit);var vals=[];rows.forEach(function(r){vals.push(n(r.p25),n(r.p75),n(r.p50))});var min=Math.max(0,Math.floor(Math.min.apply(null,vals)/10)*10);var max=Math.ceil(Math.max.apply(null,vals)/10)*10;if(!Number.isFinite(min))min=0;if(!Number.isFinite(max)||max<=min)max=min+100;return {rows:rows,min:min,max:max,span:max-min}}
function acceptRangeLevel(row){var w=n(row.p75)-n(row.p25);if(w<=20)return '集中';if(w<=60)return '适中';return '跨度大'}
function acceptRangeBoard(limit){var ar=acceptRangeData(limit);function pos(v){return Math.max(0,Math.min(100,((n(v)-ar.min)/ar.span)*100))}var note='<div class="range-note"><span><b>左端</b>=建议价格下限 P25</span><span><b>右端</b>=建议价格上限 P75</span><span><b>蓝色区间越短</b>=价位越集中</span><span><b>白色圆点</b>=典型成交价 P50</span></div>';var scale='<div class="range-scale"><span>¥'+ar.min+'</span><span>成交单价</span><span>¥'+ar.max+'</span></div>';var rows=ar.rows.map(function(r){var left=pos(r.p25),right=pos(r.p75),mid=pos(r.p50),width=Math.max(1,right-left);return '<div class="range-row" title="'+esc(r.category+'：建议 ¥'+r.p25+'-'+r.p75+'；中位价 ¥'+r.p50+'；区间'+acceptRangeLevel(r))+'"><div class="range-cat">'+esc(r.category)+'</div><div class="range-track"><span class="range-bar" style="left:'+left+'%;width:'+width+'%"></span><span class="range-dot" style="left:'+mid+'%"></span></div><div class="range-value">¥'+r.p25+'-'+r.p75+'<small>'+acceptRangeLevel(r)+' / P50 ¥'+r.p50+'</small></div></div>'}).join('');return '<div class="range-board">'+note+scale+rows+'</div>'}
function brandSummaryRows(){var b=D.brand_style_analysis||{brand_summary:[]};return (b.brand_summary||[])}
function designerSummaryRows(){var b=D.brand_style_analysis||{designer_summary:[]};return (b.designer_summary||[])}
function customerBrandTooltip(r){return ['品牌 '+esc(r.brand||'-'),'排名 Top'+(r.rank||'-'),'净销售量 '+fmt(r.qty)+'件','净销售金额 '+money(r.amount),'覆盖销售额Top20客户 '+fmt(r.customers)+'人','销售额Top20客户内销量占比 '+pct(r.qty_share,1),'销售额Top20客户内金额占比 '+pct(r.amount_share,1)]}
function shopBestBrandTooltip(r){return ['店铺 '+esc(r.shop||'-'),'Top1品牌 '+esc(r.brand||'-'),'成交金额 '+money(r.amount),'成交数量 '+fmt(r.qty)+'件','客户数 '+fmt(r.customers)+'人']}
function categoryDesignerTopRows(){var b=D.brand_style_analysis||{category_designer:[]},rows=positiveRows(b.category_designer||[],['qty','amount','customers']),by={};rows.forEach(function(r){var c=r.category||'未分类',d=r.designer||'未标记';if(!by[c])by[c]={category:c,totalAmount:0,totalQty:0,items:{}};by[c].totalAmount+=n(r.amount);by[c].totalQty+=n(r.qty);if(!by[c].items[d])by[c].items[d]={designer:d,amount:0,qty:0};by[c].items[d].amount+=n(r.amount);by[c].items[d].qty+=n(r.qty)});return Object.keys(by).map(function(c){var g=by[c],items=Object.keys(g.items).map(function(d){return g.items[d]}).filter(function(x){return n(x.amount)>0||n(x.qty)>0}).sort(function(a,b){return b.amount-a.amount||b.qty-a.qty}).slice(0,1);return {category:c,totalAmount:g.totalAmount,totalQty:g.totalQty,items:items}}).filter(function(r){return n(r.totalAmount)>0||n(r.totalQty)>0}).sort(function(a,b){return b.totalAmount-a.totalAmount})}
function renderCategoryDesignerTopChart(id){
var el=document.getElementById(id);
if(!el)return;
var rows=categoryDesignerTopRows(), colors=[P.amount,P.qty,P.trend];
function fallback(){
  var max=Math.max.apply(null,rows.map(function(r){return n(r.totalAmount)}).concat([1]));
  el.innerHTML='<div class="fallback-chart">'+rows.map(function(r){
    var first=(r.items||[])[0]||{}, w=Math.max(1,Math.min(100,n(first.amount)/max*100));
    return '<div class="fallback-row" title="'+esc((r.category||'-')+' - '+(first.designer||'-')+'：贡献金额 '+money(first.amount||0)+' / '+fmt(first.qty||0)+'件')+'"><b>'+esc((r.category||'-')+' - '+(first.designer||'-'))+'</b><div class="fallback-track"><i style="width:'+w+'%;background:'+colors[0]+'"></i></div><span>'+money(first.amount||0)+'</span></div>';
  }).join('')+'</div>';
}
var option={
  color:colors,
  animationDuration:650,
  animationEasing:'cubicOut',
  grid:{left:70,right:150,top:30,bottom:56,containLabel:true},
  tooltip:{
    trigger:'axis',
    axisPointer:{type:'shadow'},
    backgroundColor:'rgba(16,32,95,.96)',
    borderWidth:0,
    padding:[9,11],
    textStyle:{color:'#fff',fontSize:12},
    formatter:function(params){
      var list=Array.isArray(params)?params:[params],row=rows[list[0].dataIndex]||{},it=(row.items||[])[0]||{};
      var html='<div style="font-weight:700;margin-bottom:6px">'+esc((row.category||'-')+' - '+(it.designer||'-'))+'</div>';
      html+='<div style="display:flex;gap:6px;align-items:center;line-height:1.7"><span style="display:inline-block;width:9px;height:9px;border-radius:2px;background:'+colors[0]+'"></span><span>贡献金额：'+money(it.amount)+' / '+fmt(it.qty)+'件</span></div>';
      return html;
    }
  },
  legend:{show:false},
  xAxis:{type:'value',name:'净销售金额',nameLocation:'middle',nameGap:30,axisLine:{show:false},axisTick:{show:false},axisLabel:{color:'#687083',fontSize:11,formatter:function(v){return money(v)}},splitLine:{lineStyle:{color:'rgba(16,32,95,.09)'}}},
  yAxis:{type:'category',data:rows.map(function(r){return r.category}),inverse:true,axisLine:{show:false},axisTick:{show:false},axisLabel:{color:'#4f5665',fontSize:11,interval:0,width:128,overflow:'truncate',ellipsis:'...'},splitLine:{show:false}},
  series:[0].map(function(rank){
    return {
      name:'Top'+(rank+1),
      type:'bar',
      barMaxWidth:14,
      barGap:'28%',
      data:rows.map(function(r){
        var it=(r.items||[])[rank]||{};
        return {value:n(it.amount),designer:it.designer||'',qty:n(it.qty)};
      }),
      itemStyle:{color:colors[rank],borderRadius:[0,4,4,0]},
      label:{show:true,position:'right',distance:6,width:118,overflow:'truncate',color:'#4f5665',fontSize:11,formatter:function(p){return p.data&&p.data.value?String(p.data.designer||''):''}}
    };
  })
};
mkOption(id,option,fallback);
}
function brandDesignerBoard(){var b=D.brand_style_analysis||{brand_summary:[],brand_designer:[]},by={};positiveRows(b.brand_designer||[],['qty','amount','customers']).forEach(function(r){var brand=r.brand||'未标记',designer=r.designer||'未标记';if(!by[brand])by[brand]={amount:0,qty:0,items:{}};by[brand].amount+=n(r.amount);by[brand].qty+=n(r.qty);if(!by[brand].items[designer])by[brand].items[designer]={designer:designer,amount:0,qty:0};by[brand].items[designer].amount+=n(r.amount);by[brand].items[designer].qty+=n(r.qty)});var brands=positiveRows(b.brand_summary||[],['qty','amount','customers']).slice().sort(function(a,b){return n(b.amount)-n(a.amount)||n(b.qty)-n(a.qty)});var palette=[P.amount,P.qty,P.trend];return '<div class="brand-designer-board">'+brands.map(function(br,i){var brand=br.brand||'未标记',g=by[brand]||{amount:n(br.amount),qty:n(br.qty),items:{}},total=n(br.amount)||n(g.amount)||1,totalQty=n(br.qty)||n(g.qty),items=Object.keys(g.items).map(function(k){return g.items[k]}).filter(function(x){return n(x.amount)>0||n(x.qty)>0}).sort(function(a,b){return n(b.amount)-n(a.amount)||n(b.qty)-n(a.qty)}).slice(0,3);var rows=items.length?items.map(function(it,j){var share=total?n(it.amount)/total:0,w=Math.max(2,Math.min(100,share*100));return '<div class="brand-designer-row" title="'+esc(brand+' / '+it.designer+'：'+money(it.amount)+'，'+fmt(it.qty)+'件，占该品牌 '+pct(it.amount,total))+'"><strong>'+esc(it.designer)+'</strong><span>'+money(it.amount)+'</span><small>'+fmt(it.qty)+'件 / 占该品牌 '+pct(it.amount,total)+'</small><div class="brand-designer-bar"><i style="width:'+w+'%;background:'+palette[j%palette.length]+'"></i></div></div>'}).join(''):'<div class="cat-empty">暂无设计师品牌数据</div>';return '<div class="brand-designer-card"><div class="brand-designer-head"><div><b>'+esc(brand)+'</b><span>'+fmt(totalQty)+'件 / '+money(total)+'</span></div><em>品牌Top'+(i+1)+'</em></div><div class="brand-designer-list">'+rows+'</div></div>'}).join('')+'</div>'}
function seasonData(){return D.season_preference_analysis||{customers:[],customer_type_summary:[],order_type_summary:[],booking_products:[],booking_categories:[],booking_brands:[],booking_designers:[],high_potential_cutoff:0}}
function seasonTypeColor(v){return v==='订货会型客户'?P.amount:v==='普通现货型客户'?P.qty:v==='双轨客户'?P.coverage:v==='高潜订货会客户'?P.trend:P.muted}
function seasonTypeOrder(){return ['订货会型客户','双轨客户','高潜订货会客户','普通现货型客户','低参与客户']}
function seasonTypeDef(v){return v==='订货会型客户'?'订货会金额占比 ≥ 60%，主要通过订货会成交。':v==='双轨客户'?'订货会金额占比在 20%-60%，普通款期和订货会都有成交。':v==='高潜订货会客户'?'普通成交金额较高，但订货会金额占比 < 20%，适合重点引导订货会。':v==='普通现货型客户'?'订货会金额占比 < 20%，主要依赖普通款期/现货成交。':'成交体量或订货会参与较低，暂不做强订货会判断。'}
function seasonTypeRows(includeZero){var raw=seasonData().customer_type_summary||[],by={};raw.forEach(function(r){by[r.type]=r});return seasonTypeOrder().map(function(t){var r=by[t]||{type:t,customers:0,amount:0,qty:0,booking_amount:0,booking_qty:0};return r}).filter(function(r){return includeZero||n(r.customers)>0})}
function seasonSummaryBoard(){var s=seasonData(),sum=s.order_type_summary||[],normal=sum.find(function(r){return r.type==='普通款期成交'})||{},booking=sum.find(function(r){return r.type==='订货会成交'})||{},totalAmount=n(normal.amount)+n(booking.amount),totalQty=n(normal.qty)+n(booking.qty);return '<div class="season-summary"><div class="season-summary-card"><span>普通款期成交</span><b>'+money(normal.amount)+'</b><em>'+fmt(normal.qty)+'件 / '+fmt(normal.customers)+'客户</em></div><div class="season-summary-card"><span>订货会成交</span><b>'+money(booking.amount)+'</b><em>'+fmt(booking.qty)+'件 / '+fmt(booking.customers)+'客户</em></div><div class="season-summary-card"><span>订货会金额占比</span><b>'+pct(booking.amount,totalAmount)+'</b><em>按净销售金额计算</em></div><div class="season-summary-card"><span>订货会数量占比</span><b>'+pct(booking.qty,totalQty)+'</b><em>按净销售量计算</em></div></div>'}
function seasonTypePieBoard(id){var rows=seasonTypeRows(true),active=seasonTypeRows(false);if(!active.length)return '<div class="cat-empty">暂无客户类型数据</div>';return '<div class="season-type-layout"><div class="season-type-chart"><div id="'+id+'" class="echart"></div></div><div class="season-type-defs">'+rows.map(function(r){return '<div class="season-type-def"><i style="background:'+seasonTypeColor(r.type)+'"></i><b>'+esc(r.type)+'</b><span>'+esc(seasonTypeDef(r.type))+'</span><em>'+fmt(r.customers)+'客户</em></div>'}).join('')+'</div></div>'}
function seasonCustomerTooltip(r){return ['客户 '+esc(r.customer||'-'),'客户类型 '+esc(r.type||'-'),'普通款期成交 '+money(r.normal_amount)+' / '+fmt(r.normal_qty)+'件','订货会成交 '+money(r.booking_amount)+' / '+fmt(r.booking_qty)+'件','客户总成交 '+money(r.total_amount)+' / '+fmt(r.total_qty)+'件','订货会金额占比 '+(r.booking_amount_share||0)+'%','订货会数量占比 '+(r.booking_qty_share||0)+'%','订单数 '+fmt(r.orders)+'笔']}
function bookingHotTooltip(r,nameKey){return ['名称 '+esc(r[nameKey]||'-'),'订货会成交数量 '+fmt(r.qty)+'件','订货会成交金额 '+money(r.amount),'成交客户数 '+fmt(r.customers)+'人','订单数 '+fmt(r.orders)+'笔','拿货记录数 '+fmt(r.lines)+'条','单客户平均拿货 '+fmt(r.avg_qty_per_customer)+'件']}
function seasonHotRows(key,limit){return positiveRows(seasonData()[key]||[],['qty','amount','customers','orders','lines']).slice().sort(function(a,b){return n(b.qty)-n(a.qty)||n(b.amount)-n(a.amount)||n(b.customers)-n(a.customers)}).slice(0,limit)}
var PF_BRAND='';
function pfData(){return D.customer_visual_profiles||{customers:[],profiles:[]}}
function pfBaseCurrent(){var p=pfData();var name=document.getElementById('pfCustomer')?.value||p.customers[0];return (p.profiles||[]).find(function(r){return r.customer===name})||p.profiles[0]}
function pfCurrent(){var base=pfBaseCurrent();if(!base)return null;if(!PF_BRAND)return base;return (base.brand_profiles||[]).find(function(r){return r.brand===PF_BRAND})||base}
function pfSelectOptions(){var p=pfData();return (p.customers||[]).map(function(c){return '<option value="'+esc(c)+'">'+esc(c)+'</option>'}).join('')}
function pfBrandRows(){return valueRows(pfBaseCurrent()?.brand_options||[])}
function pfFilteredBrandRows(showAll){var input=document.getElementById('pfBrandInput'),q=showAll?'':String(input?.value||'').trim().toLowerCase();var rows=pfBrandRows();return q?rows.filter(function(r){return String(r.name||'').toLowerCase().indexOf(q)>=0}):rows}
function pfBrandBoxOpen(open){var box=document.getElementById('pfBrandBox');if(box)box.classList.toggle('open',!!open)}
function renderPfBrandMenu(showAll){var menu=document.getElementById('pfBrandMenu');if(!menu)return;var rows=pfFilteredBrandRows(showAll);menu.innerHTML='<button type="button" class="pf-brand-option" data-brand-option data-brand=""><b>全部品牌</b><span>展示该客户全部品牌画像</span></button>'+rows.map(function(r){return '<button type="button" class="pf-brand-option" data-brand-option data-brand="'+esc(r.name)+'"><b>'+esc(r.name)+'</b><span>'+money(r.amount)+' / '+fmt(r.qty)+'件 / '+fmt(r.orders)+'笔订单</span></button>'}).join('')+(rows.length?'':'<div class="pf-brand-empty">当前客户没有匹配品牌</div>')}
function refreshPfBrandOptions(){var input=document.getElementById('pfBrandInput'),base=pfBaseCurrent();if(!input||!base)return;var names=(base.brand_options||[]).map(function(r){return r.name});if(names.indexOf(PF_BRAND)<0)PF_BRAND='';input.value=PF_BRAND;renderPfBrandMenu(true)}
function setPfBrand(brand){PF_BRAND=brand||'';var input=document.getElementById('pfBrandInput');if(input)input.value=PF_BRAND;pfBrandBoxOpen(false);renderCustomerProfile()}
function pfMetricCards(p){document.getElementById('pfMini').innerHTML='<div><b>'+money(p.amount)+'</b><span>成交金额</span></div><div><b>'+fmt(p.qty)+'</b><span>成交数量</span></div><div><b>'+p.orders+'</b><span>订单数</span></div><div><b>¥'+p.avg_price+'</b><span>商品成交均价</span></div><div><b>¥'+p.price_low+'-'+p.price_high+'</b><span>常买价格区间</span></div>'}
function pfNames(rows,cnt){return valueRows(rows||[]).slice(0,cnt).map(function(r){return r.name}).filter(Boolean)}
function pfRecommendation(p){var scoped=!!p.brand,cats=pfNames(p.categories,scoped?1:3),colors=pfNames(p.colors,scoped?1:3),sizes=pfNames(p.sizes,1),brands=scoped?[p.brand]:pfNames(p.brands,2),designers=pfNames(p.designers,1),r=p.repeat||{},hasRepeat=!!r.has_repeat;function chips(arr){return (arr.length?arr:['暂无']).map(function(x){return '<span>'+esc(x)+'</span>'}).join('')}var priceType=(p.price_type||'待判断')+' · '+(p.main_price_segment||'-')+' '+(p.price_type_share||0)+'%',repeatText=hasRepeat?'会重复拿同一商品':'暂无明显复拿',intervalText=(r.avg_interval_days||0)?(r.avg_interval_days+'天'):'-';document.getElementById('pfRecommend').innerHTML='<b class="pf-rec-title">客户推荐画像</b><div class="pf-rec-row"><b>客户类型</b><div><span title="'+esc(p.price_type_desc||'')+'">'+esc(priceType)+'</span></div></div><div class="pf-rec-row"><b>推荐品类</b><div>'+chips(cats)+'</div></div><div class="pf-rec-row"><b>推荐颜色</b><div>'+chips(colors)+'</div></div><div class="pf-rec-row"><b>推荐尺码</b><div>'+chips(sizes)+'</div></div><div class="pf-rec-row"><b>常买价位</b><div><span>¥'+p.price_low+'-'+p.price_high+'</span></div></div><div class="pf-rec-row"><b>主品牌</b><div>'+chips(brands)+'</div></div><div class="pf-rec-row"><b>设计师品牌</b><div>'+chips(designers)+'</div></div><div class="pf-rec-row"><b>复拿情况</b><div><span title="同一货号出现2笔及以上正向拿货订单才算复拿，不含退货">'+esc(repeatText)+' · '+fmt(r.repeat_product_count||0)+'个货号</span></div></div><div class="pf-rec-row"><b>平均补货间隔</b><div><span>'+esc(intervalText)+'</span></div></div>'}
function pfRepeatProductTooltip(row){return ['货号 '+esc(row.product||'-'),'品类 '+esc(row.category||'-'),'复拿金额 '+money(row.repeat_amount),'复拿数量 '+fmt(row.repeat_qty)+'件','复拿次数 '+fmt(row.repeat_count)+'次','首次 '+esc(row.first_date||'-')+' / 最近 '+esc(row.last_date||'-'),'平均间隔 '+(row.avg_interval_days||0)+'天']}
function renderPfRepeatCharts(p){var rep=p.repeat||{},products=(rep.products||[]).filter(function(r){return n(r.repeat_qty)>0||n(r.repeat_amount)>0}).slice(0,15);mk('pfRepeatProducts',{type:'bar',data:{labels:products.map(function(r){return r.product}),datasets:[{label:'复拿数量',data:products.map(function(r){return r.repeat_qty}),backgroundColor:P.qty,xAxisID:'x',barMaxWidth:22},{label:'复拿金额',data:products.map(function(r){return r.repeat_amount}),backgroundColor:P.amount,xAxisID:'x1',barMaxWidth:22}]},options:{responsive:true,maintainAspectRatio:false,indexAxis:'y',grid:{left:112,right:124,top:62,bottom:78},plugins:{legend:{position:'bottom'},tooltip:{callbacks:{label:function(c){return c.dataset.xAxisID==='x1'?'复拿金额: '+money(c.raw):'复拿数量: '+fmt(c.raw)+'件'},afterBody:function(ctx){return pfRepeatProductTooltip(products[ctx[0].dataIndex]||{})}}}},scales:{x:{position:'bottom',title:{display:true,text:'复拿数量'},ticks:{callback:function(v){return fmt(v)}}},x1:{position:'top',grid:{drawOnChartArea:false},title:{display:true,text:'复拿金额'},ticks:{callback:function(v){return money(v)}}},y:{ticks:{autoSkip:false,font:{size:10}}}}}})}
function pfColorPieTooltip(p,row,total){var d=(p.color_details||[]).find(function(x){return x.color===row.name})||{};function lines(title,items){items=(items||[]).filter(function(x){return n(x.qty)>0}).slice(0,5);if(!items.length)return '';return '<div style="margin-top:8px;padding-top:7px;border-top:1px solid rgba(255,255,255,.16)"><div style="font-weight:700;margin-bottom:4px">'+title+'</div>'+items.map(function(x){var share=n(x.share);return '<div style="display:flex;justify-content:space-between;gap:16px;line-height:1.65"><span>'+esc(x.name)+'</span><span>'+fmt(x.qty)+'件 / '+(share*100).toFixed(1)+'%</span></div>'}).join('')+'</div>'}return '<div style="font-weight:800;margin-bottom:6px">'+esc(row.name)+'</div><div style="line-height:1.75">拿货数量：'+fmt(row.value)+'件 ('+pct(row.value,total)+')</div><div style="line-height:1.75">拿货金额：'+money(d.amount||0)+'</div>'+lines('品类占比Top5',d.categories)+lines('品牌占比Top5',d.brands)}
function pfColorPie(p){var rows=valueRows(p.colors||[]).slice().sort(function(a,b){return n(b.value)-n(a.value)}).slice(0,15);mk('pfColorPie',{type:'pie',data:{labels:rows.map(function(r){return r.name}),datasets:[{label:'拿货数量',data:rows.map(function(r){return r.value}),backgroundColor:rows.map(function(r,i){return r.hex||C[i%C.length]}),borderColor:'#fff',borderWidth:1}]},options:{responsive:true,maintainAspectRatio:false,plugins:{legend:{display:false},tooltip:{trigger:'item',formatter:function(c){var total=(c.data&&c.seriesType==='pie'?rows.reduce(function(a,b){return a+n(b.value)},0):rows.reduce(function(a,b){return a+n(b.value)},0));return pfColorPieTooltip(p,rows[c.dataIndex]||{name:c.name,value:c.value},total)}}}}})}
function pfBar(id,rows,label,color){var rs=valueRows(rows||[]);mk(id,{type:'bar',data:{labels:rs.map(function(r){return r.name}),datasets:[{label:label,data:rs.map(function(r){return r.value}),backgroundColor:color}]},options:{responsive:true,maintainAspectRatio:false,indexAxis:'y',plugins:{legend:{display:false}},scales:{x:{ticks:{callback:function(v){return fmt(v)}}}}}})}
function pfPriceBandColor(i){return ['#60a5fa',P.qty,P.amount,P.coverage,'#00a6ff',P.trend,P.muted][i%7]}
function renderPfCategoryPriceChart(p){var rows=(p.category_prices||[]).filter(function(r){return n(r.amount)>0&&(r.price_bands||[]).some(function(b){return n(b.value)>0})}).sort(function(a,b){return n(b.amount)-n(a.amount)});var order=['0-50','50-100','100-150','150-200','200-300','300-500','500+'];var bandNames=order.filter(function(name){return rows.some(function(r){return (r.price_bands||[]).some(function(b){return b.name===name&&n(b.value)>0})})});if(!rows.length||!bandNames.length){var priceRows=valueRows(p.price_bands||[]);mk('pfPrice',{type:'bar',data:{labels:priceRows.map(function(r){return r.name}),datasets:[{label:'金额',data:priceRows.map(function(r){return r.value}),backgroundColor:'#00a6ff'}]},options:{responsive:true,maintainAspectRatio:false,plugins:{legend:{display:false}},scales:{y:{ticks:{callback:function(v){return money(v)}}}}}});return}function bandValue(row,name){var it=(row.price_bands||[]).find(function(x){return x.name===name})||{};return n(it.value)}function bandTotal(row){return (row.price_bands||[]).reduce(function(sum,b){return sum+n(b.value)},0)}mk('pfPrice',{type:'bar',data:{labels:rows.map(function(r){return r.name}),datasets:bandNames.map(function(name,i){return {label:name,bandName:name,data:rows.map(function(r){return bandValue(r,name)}),backgroundColor:pfPriceBandColor(i),stack:'price',barMaxWidth:28}})},options:{responsive:true,maintainAspectRatio:false,indexAxis:'y',grid:{left:92,right:56,top:44,bottom:86},plugins:{legend:{position:'bottom'},tooltip:{callbacks:{label:function(c){var row=rows[c.dataIndex]||{},total=bandTotal(row);return c.dataset.label+': '+money(c.raw)+' / 价格带内占 '+pct(c.raw,total)},afterBody:function(ctx){var row=rows[ctx[0].dataIndex]||{},total=bandTotal(row);return ['价格带金额合计 '+money(total),'品类净销售金额 '+money(row.amount),'品类净销售数量 '+fmt(row.qty)+'件','订单数 '+fmt(row.orders)+'笔','成交均价 ¥'+(row.avg||0),'常买价格区间 ¥'+(row.low||0)+'-'+(row.high||0)]}}}},scales:{x:{title:{display:true,text:'价格带成交金额'},ticks:{callback:function(v){return money(v)}}},y:{ticks:{autoSkip:false,font:{size:11}}}}}})}
function renderCustomerProfile(){var p=pfCurrent();if(!p)return;pfMetricCards(p);pfRecommendation(p);renderPfRepeatCharts(p);pfColorPie(p);pfBar('pfCat',p.categories,'金额',P.amount);var pfBrands=valueRows(p.brands||[]),pfDesigners=valueRows(p.designers||[]),pfBrandLabels=[...new Set(pfBrands.map(function(r){return r.name}).concat(pfDesigners.map(function(r){return r.name})))].filter(function(x){var br=pfBrands.find(function(a){return a.name===x}),de=pfDesigners.find(function(a){return a.name===x});return n(br&&br.value)>0||n(de&&de.value)>0});mk('pfBrand',{type:'bar',data:{labels:pfBrandLabels,datasets:[{label:'品牌',data:pfBrandLabels.map(function(x){var r=pfBrands.find(function(a){return a.name===x});return r?r.value:0}),backgroundColor:P.qty},{label:'设计师品牌',data:pfBrandLabels.map(function(x){var r=pfDesigners.find(function(a){return a.name===x});return r?r.value:0}),backgroundColor:P.amount}]},options:{responsive:true,maintainAspectRatio:false,indexAxis:'y',plugins:{legend:{position:'bottom'},tooltip:{trigger:'item',callbacks:{label:function(c){return '金额：'+money(c.raw)}}}},scales:{x:{ticks:{callback:function(v){return money(v)}}}}}});pfBar('pfSize',p.sizes,'销量',P.coverage);renderPfCategoryPriceChart(p);var seasonRows=valueRows(p.seasons||[]);mk('pfSeason',{type:'doughnut',data:{labels:seasonRows.map(function(r){return r.name}),datasets:[{data:seasonRows.map(function(r){return r.value}),backgroundColor:seasonRows.map(function(r,i){return C[i%C.length]})}]},options:{responsive:true,maintainAspectRatio:false,plugins:{legend:{position:'right'},tooltip:{callbacks:{label:function(c){return c.label+': '+money(c.raw)}}}}}});hydrateUi()}

document.getElementById('sub').textContent=S.date_from+' ~ '+S.date_to+' | '+S.records.toLocaleString()+'条交易 | '+S.customers.toLocaleString()+'位客户';
document.getElementById('kpi').innerHTML=
'<div class="kpi-card"><div class="kpi-top"><div class="l">总销售额</div><i data-lucide="badge-dollar-sign"></i></div><div class="v">'+money(S.amount)+'</div><div class="s">'+(overviewData().order_count||0).toLocaleString()+'笔订单</div></div>'
+'<div class="kpi-card"><div class="kpi-top"><div class="l">总数量</div><i data-lucide="package-check"></i></div><div class="v">'+fmt(S.qty)+'件</div><div class="s">'+S.products.toLocaleString()+'个货号</div></div>'
+'<div class="kpi-card"><div class="kpi-top"><div class="l">客户数</div><i data-lucide="users"></i></div><div class="v">'+S.customers.toLocaleString()+'</div><div class="s">'+S.shops+'家店铺</div></div>'
+'<div class="kpi-card"><div class="kpi-top"><div class="l">均单价</div><i data-lucide="receipt"></i></div><div class="v">¥'+S.avg_price+'</div><div class="s">'+S.categories+'品类/'+S.brands+'品牌</div></div>'
+'<div class="kpi-card"><div class="kpi-top"><div class="l">设计师</div><i data-lucide="badge"></i></div><div class="v">'+S.designers+'</div><div class="s">位</div></div>';

const TABS=[{id:'ov',n:'总览',ic:'layout-dashboard'},{id:'cat',n:'品类偏好',ic:'layers-3'},{id:'co',n:'颜色偏好',ic:'palette'},{id:'sz',n:'码数偏好',ic:'ruler'},{id:'br',n:'品牌/风格',ic:'badge'},{id:'ys',n:'普通款/订货会分析',ic:'calendar-range'},{id:'se',n:'价格接受度分析',ic:'badge-dollar-sign'},{id:'pf',n:'客户画像',ic:'user-round-search'}];
document.getElementById('tabs').innerHTML=TABS.map(function(t,i){return '<div class="tab'+(i===0?' on':'')+'" data-t="'+t.id+'"><i data-lucide="'+t.ic+'"></i>'+t.n+'</div>'}).join('');
document.getElementById('bd').innerHTML=TABS.map(function(t,i){return '<div class="sec'+(i===0?' on':'')+'" id="s-'+t.id+'"></div>'}).join('');
document.querySelectorAll('.tab').forEach(function(t){t.onclick=function(){document.querySelectorAll('.tab').forEach(function(x){x.classList.remove('on')});document.querySelectorAll('.sec').forEach(function(x){x.classList.remove('on')});t.classList.add('on');document.getElementById('s-'+t.dataset.t).classList.add('on');hydrateUi();[40,160,420].forEach(function(ms){setTimeout(resizeCharts,ms)})}});
window.addEventListener('resize',resizeCharts);

var OV=overviewData(), topCat=D.cat_dist[0]||{}, topCustomer=D.customer_top[0]||{}, topMonth=topRows(OV.monthly_sales||D.monthly,1,'amount')[0]||{}, topBrand=D.brand_dist[0]||{}, topDesigner=D.designer_dist[0]||{};
document.getElementById('s-ov').innerHTML='<div class="tip"><b>总览洞察：</b>本期核心品类为'+esc(topCat['分类'])+'，贡献'+money(topCat['金额'])+'；Top客户为'+esc(topCustomer['客户'])+'，贡献'+money(topCustomer['金额'])+'；销售峰值月份为'+esc(topMonth.month||topMonth['月份']||'-')+'。</div><div class="g">'
+card('经营热销总结',overviewSummaryCard())
+card('品类金额/客户覆盖 Top12',chart('c1',true))
+card('月度整体销量和销售额',chart('c3',false))
+card('店铺销售额与销量',chart('c1shop',true))
+card('商品重复拿货排行',largeChart('c1repeat'),true)
+card('Top15客户贡献',chart('c2',true))
+card('热销货号 Top15',chart('c4',false))
+'</div>';
var cats=positiveRows(D.cat_dist||[],['数量','金额','客户数']).slice(0,12);
mk('c1',{type:'bar',data:{labels:cats.map(function(r){return r['分类']}),datasets:[{label:'金额',data:cats.map(function(r){return r['金额']}),backgroundColor:P.amount,xAxisID:'x'},{label:'客户数',data:cats.map(function(r){return r['客户数']}),backgroundColor:P.qty,xAxisID:'x1'}]},options:{responsive:true,maintainAspectRatio:false,indexAxis:'y',plugins:{legend:{position:'bottom'},tooltip:{callbacks:{label:function(c){return c.dataset.xAxisID==='x1'?'客户数: '+fmt(c.raw)+'人':'金额: '+money(c.raw)},afterBody:function(ctx){var r=cats[ctx[0].dataIndex]||{};return '成交数量 '+fmt(r['数量'])+'件'}}}},scales:{x:{position:'bottom',title:{display:true,text:'销售额'},ticks:{callback:function(v){return money(v)}}},x1:{position:'top',grid:{drawOnChartArea:false},title:{display:false},ticks:{callback:function(v){return fmt(v)+'人'}}}}}});
var customerTopRows=positiveRows(D.customer_top||[],['金额','数量']).slice(0,15);
mk('c2',{type:'bar',data:{labels:customerTopRows.map(function(r){return r['客户']}),datasets:[{data:customerTopRows.map(function(r){return r['金额']}),backgroundColor:P.amount}]},options:{responsive:true,maintainAspectRatio:false,indexAxis:'y',plugins:{legend:{display:false}},scales:{x:{ticks:{callback:function(v){return money(v)}}}}}});
var monthlySales=positiveRows(OV.monthly_sales||[],['amount','qty','customers']);
mk('c3',{type:'line',data:{labels:monthlySales.map(function(r){return r.month}),datasets:[{label:'销售额',data:monthlySales.map(function(r){return r.amount}),borderColor:P.amount,backgroundColor:'rgba(21,92,255,.10)',fill:true,tension:.3,yAxisID:'y'},{label:'销量',data:monthlySales.map(function(r){return r.qty}),borderColor:P.qty,backgroundColor:'rgba(18,183,106,.08)',fill:false,tension:.3,yAxisID:'y1'}]},options:{responsive:true,maintainAspectRatio:false,plugins:{legend:{position:'bottom'},tooltip:{callbacks:{label:function(c){return c.dataset.yAxisID==='y1'?'销量: '+fmt(c.raw)+'件':'销售额: '+money(c.raw)}}}},scales:{y:{position:'left',title:{display:true,text:'销售额'},ticks:{callback:function(v){return money(v)}}},y1:{position:'right',grid:{drawOnChartArea:false},title:{display:true,text:'销量'},ticks:{callback:function(v){return fmt(v)+'件'}}}}}});
var shopSales=positiveRows(OV.shop_sales||D.shop_dist||[],['amount','qty','customers','金额','数量','客户数']);
mk('c1shop',{type:'bar',data:{labels:shopSales.map(function(r){return r.shop||r['店铺']}),datasets:[{label:'销售额',data:shopSales.map(function(r){return r.amount||r['金额']}),backgroundColor:P.amount,xAxisID:'x'},{label:'销量',data:shopSales.map(function(r){return r.qty||r['数量']}),backgroundColor:P.qty,xAxisID:'x1'}]},options:{responsive:true,maintainAspectRatio:false,indexAxis:'y',plugins:{legend:{position:'bottom'},tooltip:{callbacks:{label:function(c){return c.dataset.xAxisID==='x1'?'销量: '+fmt(c.raw)+'件':'销售额: '+money(c.raw)},afterBody:function(ctx){var r=shopSales[ctx[0].dataIndex]||{};return '客户数 '+fmt(r.customers||r['客户数'])+'人'}}}},scales:{x:{position:'bottom',title:{display:true,text:'销售额'},ticks:{callback:function(v){return money(v)}}},x1:{position:'top',grid:{drawOnChartArea:false},title:{display:true,text:'销量'},ticks:{callback:function(v){return fmt(v)+'件'}}}}}});
var repeatProducts=positiveRows(OV.repeat_products||[],['repeat_customers']).slice(0,30);
mk('c1repeat',{type:'bar',data:{labels:repeatProducts.map(function(r){return r.product}),datasets:[{label:'重复拿货客户数',data:repeatProducts.map(function(r){return r.repeat_customers}),backgroundColor:P.qty,barMaxWidth:22}]},options:{responsive:true,maintainAspectRatio:false,indexAxis:'y',plugins:{legend:{display:false},tooltip:{callbacks:{label:function(c){return '重复拿货客户数: '+fmt(c.raw)+'人'},afterBody:function(ctx){var r=repeatProducts[ctx[0].dataIndex]||{};return ['净销售金额 '+money(r.amount),'净销售量 '+fmt(r.qty)+'件','总拿货客户 '+fmt(r.customers)+'人']}}}},scales:{x:{title:{display:true,text:'商品重复拿货客户数'},ticks:{callback:function(v){return fmt(v)+'人'}}},y:{ticks:{autoSkip:false,font:{size:10}}}}}});
var productRows=positiveRows(D.product_top||[],['金额','数量','客户数']).slice(0,15);
mk('c4',{type:'bar',data:{labels:productRows.map(function(r){return r['货号']}),datasets:[{data:productRows.map(function(r){return r['金额']}),backgroundColor:C}]},options:{responsive:true,maintainAspectRatio:false,indexAxis:'y',plugins:{legend:{display:false}},scales:{x:{ticks:{callback:function(v){return money(v)}}}}}});

var CA=catData(), catAll=positiveRows(CA.categories||[],['qty','amount','customers']);
var catByQty=catAll.slice().sort(function(a,b){return n(b.qty)-n(a.qty)});
var catByCoverage=catAll.slice().sort(function(a,b){return n(b.customers)-n(a.customers)});
var catByScore=catAll.slice().sort(function(a,b){return n(b.recommend_score)-n(a.recommend_score)});
var catCompanion=catAll.slice().filter(function(r){return r.companion&&r.companion!=='-'&&n(r.companion_orders)>0&&n(r.companion_rate)>0}).sort(function(a,b){return n(b.companion_rate)-n(a.companion_rate)});
var shopBestCategories=positiveRows(CA.shop_best_categories||[],['amount','qty','customers']);
var catMostQty=catByQty[0]||{}, catMostAmount=catAll[0]||{}, catMostCover=catByCoverage[0]||{}, catBest=catByScore[0]||{};
document.getElementById('s-cat').innerHTML='<div class="tip"><b>品类偏好洞察：</b>本板块展示全部 '+catAll.length+' 个品类，不做 Top 截断。拿货最常见的是 '+esc(catMostQty.category||'-')+'，金额贡献最高的是 '+esc(catMostAmount.category||'-')+'，客户覆盖最广的是 '+esc(catMostCover.category||'-')+'；当前建议优先主推 '+esc(catBest.category||'-')+'。趋势只看当前报表时间段内的走势，不做跨月份强行对比。</div><div class="g">'
+card('品类拿货数量 / 成交金额',chart('c5',true))
+card('各店铺卖得最好的Top1品类',chart('c5shop',true))
+card('各品类客户消费占比',chart('c6',true)+categoryCoverageNote())
+card('主力品类成交走势（金额Top10）',chart('c6b',true)+categoryTrendNote(),true)
+card('常搭品类推荐',chart('c6d',true))
+card('重点推荐优先级',chart('c6c',true)+catClassLegend())
+card('品类售卖建议',categoryAdviceBoard(),true)
+'</div>';
mk('c5',{type:'bar',data:{labels:catByQty.map(function(r){return r.category}),datasets:[{label:'成交数量',data:catByQty.map(function(r){return r.qty}),backgroundColor:P.qty,xAxisID:'x'},{label:'成交金额',data:catByQty.map(function(r){return r.amount}),backgroundColor:P.amount,xAxisID:'x1'}]},options:{responsive:true,maintainAspectRatio:false,indexAxis:'y',plugins:{legend:{position:'bottom'},tooltip:{callbacks:{afterBody:function(ctx){var r=catByQty[ctx[0].dataIndex]||{};return catTooltip(r)}}}},scales:{x:{position:'bottom',title:{display:true,text:'成交数量'},ticks:{callback:function(v){return fmt(v)}}},x1:{position:'top',grid:{drawOnChartArea:false},title:{display:true,text:'成交金额'},ticks:{callback:function(v){return money(v)}}}}}});
mk('c5shop',{type:'bar',data:{labels:shopBestCategories.map(function(r){return r.shop+' / '+r.category}),datasets:[{label:'Top1品类成交金额',data:shopBestCategories.map(function(r){return r.amount}),backgroundColor:P.qty}]},options:{responsive:true,maintainAspectRatio:false,indexAxis:'y',plugins:{legend:{display:false},tooltip:{callbacks:{label:function(c){return '成交金额: '+money(c.raw)},afterBody:function(ctx){return shopBestCategoryTooltip(shopBestCategories[ctx[0].dataIndex]||{})}}}},scales:{x:{title:{display:true,text:'成交金额'},ticks:{callback:function(v){return money(v)}}},y:{ticks:{autoSkip:false,font:{size:10}}}}}});
mk('c6',{type:'bar',data:{labels:catByCoverage.map(function(r){return r.category}),datasets:[{label:'客户数',data:catByCoverage.map(function(r){return r.customers}),backgroundColor:P.qty,xAxisID:'x'},{label:'所有客户占比',data:catByCoverage.map(function(r){return Math.round(n(r.customer_share)*1000)/10}),backgroundColor:P.trend,xAxisID:'x1'}]},options:{responsive:true,maintainAspectRatio:false,indexAxis:'y',grid:{left:42,right:92},plugins:{legend:{position:'bottom'},tooltip:{callbacks:{label:function(c){return c.dataset.label+': '+(c.dataset.xAxisID==='x1'?c.raw+'%':fmt(c.raw)+'人')},afterBody:function(ctx){var r=catByCoverage[ctx[0].dataIndex]||{};return catTooltip(r)}}}},scales:{x:{position:'bottom',title:{display:true,text:'客户数'},ticks:{callback:function(v){return fmt(v)}}},x1:{position:'top',grid:{drawOnChartArea:false},min:0,max:100,title:{display:true,text:'客户占比'},ticks:{callback:function(v){return v+'%'}}}}}});
var trendPeriods=(CA.trend_periods||[]), trendMap=trendAmountMap();
var trendTopCats=catAll.slice().filter(function(r){return n(r.amount)>0||n(r.qty)>0}).sort(function(a,b){return n(b.amount)-n(a.amount)}).slice(0,10);
mk('c6b',{type:'line',data:{labels:trendPeriods.map(function(p){return p.label}),datasets:trendTopCats.map(function(r,i){return {label:r.category,data:trendPeriods.map(function(p){var row=trendMap[r.category+'|'+p.key];return row?row.amount:0}),borderColor:C[i%C.length],backgroundColor:C[i%C.length],tension:.35,fill:false}})},options:{responsive:true,maintainAspectRatio:false,plugins:{legend:{position:'bottom'},tooltip:{callbacks:{label:function(c){return c.dataset.label+': '+money(c.raw)}}}},scales:{x:{title:{display:true,text:'报表内按'+(CA.trend_period_type||'-')+'汇总'}},y:{ticks:{callback:function(v){return money(v)}}}}}});
mk('c6d',{type:'bar',data:{labels:catCompanion.map(function(r){return r.category+' → '+r.companion}),datasets:[{label:'共购比例',data:catCompanion.map(function(r){return r.companion_rate}),backgroundColor:P.amount}]},options:{responsive:true,maintainAspectRatio:false,indexAxis:'y',plugins:{legend:{display:false},tooltip:{callbacks:{label:function(c){var r=catCompanion[c.dataIndex]||{};return '共购比例: '+c.raw+'%'},afterBody:function(ctx){var r=catCompanion[ctx[0].dataIndex]||{};return ['共购订单 '+fmt(r.companion_orders)+' 单','读法：买 '+esc(r.category)+' 时，常一起买 '+esc(r.companion)]}}}},scales:{x:{min:0,max:100,title:{display:true,text:'同订单共购比例'},ticks:{callback:function(v){return v+'%'}}}}}});
mk('c6c',{type:'bar',data:{labels:catByScore.map(function(r){return r.category}),datasets:[{label:'推荐分',data:catByScore.map(function(r){return r.recommend_score}),backgroundColor:catByScore.map(function(r){return catClassColor(r.class)})}]},options:{responsive:true,maintainAspectRatio:false,indexAxis:'y',plugins:{legend:{display:false},tooltip:{callbacks:{label:function(c){var r=catByScore[c.dataIndex]||{};return '推荐分: '+c.raw+' / '+esc(r.class)},afterBody:function(ctx){var r=catByScore[ctx[0].dataIndex]||{};return catTooltip(r)}}}},scales:{x:{min:0,max:100,title:{display:true,text:'推荐优先级评分'},ticks:{callback:function(v){return v}}}}}});

var CO=colorData(), coColors=positiveRows(CO.colors||[],['qty','amount','customers']);
var colorByGroup=['稳定常备色','近期趋势色','控制备货色'].flatMap(function(g){return coColors.filter(function(r){return r.group===g&&n(r.qty)>0}).sort(function(a,b){return g==='近期趋势色'?(n(b.recent_growth_qty)-n(a.recent_growth_qty)||n(b.qty)-n(a.qty)):n(b.qty)-n(a.qty)}).slice(0,g==='控制备货色'?6:8).map(function(r){return Object.assign({},r,{groupLabel:g})})});
var structureRows=valueRows(colorStructureRows());
var trendMap={}, colorTrendTotals={};(CO.monthly_trend||[]).forEach(function(r){trendMap[r.color+'|'+r.month]=r.qty;colorTrendTotals[r.color]=n(colorTrendTotals[r.color])+n(r.qty)});
var mainColor=coColors.slice().sort(function(a,b){return n(b.qty)-n(a.qty)})[0]||{}, trendColor=coColors.filter(function(r){return r.group==='近期趋势色'}).sort(function(a,b){return n(b.recent_growth_qty)-n(a.recent_growth_qty)||n(b.qty)-n(a.qty)})[0]||{};
document.getElementById('s-co').innerHTML='<div class="tip"><b>颜色备货洞察：</b>颜色按源表颜色名称分别统计，米白、杏色、米白色等分开计算；空值归为“未标记”，花色/拼色/条纹/格纹/豹纹等归为“图案/拼色类”。本板块只分析颜色，并结合品类给出备货优先级。全局销量最高为 '+esc(mainColor.color||'-')+'，净销售量 '+fmt(mainColor.qty)+'件；近期趋势代表色为 '+esc(trendColor.color||'-')+'。销量分层按当前报表所有颜色的净销售量区间计算，近期趋势使用 '+esc(CO.recent30_from||'-')+' ~ '+esc(CO.recent30_to||'-')+' 对比 '+esc(CO.previous30_from||'-')+' ~ '+esc(CO.previous30_to||'-')+'。</div><div class="g">'
+card('颜色销量Top30',scrollChart('c7all'),true)
+card('销售额Top20客户颜色销量Top10',customerColorScrollChart('c7cust'),true)
+card('销量Top10颜色的品类分布',chart('c7catDist',true),true)
+card('销量Top10颜色的品牌分布',chart('c7brandDist',true),true)
+card('颜色备货建议',chart('c7b',true)+stockGroupLegend())
+card('各品类卖得最好的Top1颜色',top3ScrollChart('c7c'),true)
+card('颜色月度趋势',chart('c7e',false),true)
+'</div>';
var colorAll=coColors.slice().sort(function(a,b){return n(b.qty)-n(a.qty)}).slice(0,30);
mk('c7all',{type:'bar',data:{labels:colorAll.map(function(r){return r.color}),datasets:[{label:'净销售量',data:colorAll.map(function(r){return r.qty}),backgroundColor:P.qty,borderRadius:3}]},options:{responsive:true,maintainAspectRatio:false,indexAxis:'y',plugins:{legend:{display:false},tooltip:{callbacks:{label:function(c){return '净销售量: '+fmt(c.raw)+'件'},afterBody:function(ctx){return colorTitle(colorAll[ctx[0].dataIndex]||{})}}}},scales:{x:{title:{display:true,text:'净销售量'},ticks:{callback:function(v){return fmt(v)}}}}}});
var customerColorRows=positiveRows(CO.top_customer_colors||[],['qty','amount','customers']);
mk('c7cust',{type:'bar',data:{labels:customerColorRows.map(function(r){return r.color}),datasets:[{label:'净销售量',data:customerColorRows.map(function(r){return r.qty}),backgroundColor:P.qty,xAxisID:'x'},{label:'净销售金额',data:customerColorRows.map(function(r){return r.amount}),backgroundColor:P.amount,xAxisID:'x1'}]},options:{responsive:true,maintainAspectRatio:false,indexAxis:'y',plugins:{legend:{position:'bottom'},tooltip:{callbacks:{label:function(c){return c.dataset.label+': '+(c.dataset.xAxisID==='x1'?money(c.raw):fmt(c.raw)+'件')},afterBody:function(ctx){return customerColorTooltip(customerColorRows[ctx[0].dataIndex]||{})}}}},scales:{x:{position:'bottom',title:{display:true,text:'净销售量'},ticks:{callback:function(v){return fmt(v)}}},x1:{position:'top',grid:{drawOnChartArea:false},title:{display:true,text:'净销售金额'},ticks:{callback:function(v){return money(v)}}}}}});
renderColorBreakdownChart('c7catDist',CO.top_color_category_distribution||[]);
renderColorBreakdownChart('c7brandDist',CO.top_color_brand_distribution||[]);
mk('c7b',{type:'bar',data:{labels:colorByGroup.map(function(r){return r.color}),datasets:[{label:'净销售量',data:colorByGroup.map(function(r){return r.qty}),backgroundColor:colorByGroup.map(function(r){return stockGroupColor(r.groupLabel)})}]},options:{responsive:true,maintainAspectRatio:false,indexAxis:'y',plugins:{legend:{display:false},tooltip:{callbacks:{label:function(c){return '净销售量: '+fmt(c.raw)+'件'},afterBody:function(ctx){var r=colorByGroup[ctx[0].dataIndex]||{};return ['分组 '+esc(r.groupLabel||'-')].concat(colorTitle(r))}}}},scales:{x:{title:{display:true,text:'净销售量'},ticks:{callback:function(v){return fmt(v)}}}}}});
mk('c7c',{type:'bar',data:{labels:structureRows.map(function(r){return r.label}),datasets:[{label:'净销售量',data:structureRows.map(function(r){return r.qty}),backgroundColor:structureRows.map(function(r){return colorRankColor(r.rank)}),borderRadius:3,barPercentage:.72,categoryPercentage:.78}]},options:{responsive:true,maintainAspectRatio:false,indexAxis:'y',plugins:{legend:{display:false},tooltip:{callbacks:{label:function(c){return '净销售量: '+fmt(c.raw)+'件'},afterBody:function(ctx){return colorStructureTooltipRow(structureRows[ctx[0].dataIndex]||{})}}}},scales:{x:{title:{display:true,text:'净销售量'},ticks:{callback:function(v){return fmt(v)}}},y:{title:{display:true,text:'品类 / 颜色'},ticks:{autoSkip:false,font:{size:10}}}}}});
var trendColorNames=(CO.trend_colors||[]).filter(function(color){return n(colorTrendTotals[color])>0});
mk('c7e',{type:'line',data:{labels:(CO.months||[]),datasets:trendColorNames.map(function(color,i){var r=coColors.find(function(x){return x.color===color})||{};return {label:color,data:(CO.months||[]).map(function(m){return trendMap[color+'|'+m]||0}),borderColor:r.hex||C[i%C.length],backgroundColor:r.hex||C[i%C.length],tension:.3,fill:false}})},options:{responsive:true,maintainAspectRatio:false,plugins:{legend:{position:'bottom'}},scales:{y:{title:{display:true,text:'净销售量'},ticks:{callback:function(v){return fmt(v)}}}}}});

var SP=sizeData(), sizeAll=positiveRows(SP.sizes||[],['qty','amount','customers']).slice().sort(function(a,b){return n(b.qty)-n(a.qty)}), sizeByScore=sizeAll.slice().sort(function(a,b){return n(b.score)-n(a.score)||n(b.qty)-n(a.qty)});
var sizeByGroup=['稳定常备码','近期趋势码','控制备货码'].flatMap(function(g){return sizeAll.filter(function(r){return r.group===g&&n(r.qty)>0}).sort(function(a,b){return g==='近期趋势码'?(n(b.recent_growth_qty)-n(a.recent_growth_qty)||n(b.qty)-n(a.qty)):n(b.qty)-n(a.qty)}).map(function(r){return Object.assign({},r,{groupLabel:g})})});
var sizeTrendMap={}, sizeTrendTotals={};(SP.monthly_trend||[]).forEach(function(r){sizeTrendMap[r.size+'|'+r.month]=r.qty;sizeTrendTotals[r.size]=n(sizeTrendTotals[r.size])+n(r.qty)});
var mainSize=sizeAll[0]||{}, trendSize=(SP.sizes||[]).filter(function(r){return r.group==='近期趋势码'}).sort(function(a,b){return n(b.recent_growth_qty)-n(a.recent_growth_qty)||n(b.qty)-n(a.qty)})[0]||{};
var sizeNames=sizeAll.map(function(r){return r.size});
var sizeCategories=(SP.category_structure||[]);
var topCustomerSizeRows=positiveRows(SP.top_customer_sizes||[],['qty','amount','customers']);
document.getElementById('s-sz').innerHTML='<div class="tip"><b>码数备货洞察：</b>本板块展示全部 '+sizeAll.length+' 个尺码，不做 Top 截断。全局销量最高尺码为 '+esc(mainSize.size||'-')+'，净销售量 '+fmt(mainSize.qty)+'件；近期趋势尺码为 '+esc(trendSize.size||'-')+'。当前重点查看尺码销量结构、核心客户尺码偏好、尺码分层备货建议，以及 Top5 主力尺码由哪些品类贡献。</div><div class="g">'
+card('尺码销量结构',sizePieChart('c8'))
+card('销售额Top20客户尺码销量 / 金额',largeChart('c8cust'))
+card('尺码备货建议',chart('c8c',false)+sizeStockGroupLegend())
+card('码数销量Top5品类分布',sizeTopCategoryBoard(),true)
+card('尺码月度趋势',chart('c9b',false),true)
+'</div>';
document.getElementById('sizePieLegend').innerHTML=sizePieLegend(sizeAll);
mk('c8',{type:'pie',data:{labels:sizeAll.map(function(r){return r.size}),datasets:[{label:'净销售量',data:sizeAll.map(function(r){return r.qty}),backgroundColor:sizeAll.map(function(r,i){return C[i%C.length]}),borderColor:'#fff',borderWidth:1}]},options:{responsive:true,maintainAspectRatio:false,plugins:{legend:{display:false},tooltip:{callbacks:{label:function(c){var total=c.dataset.data.reduce(function(a,b){return a+n(b)},0);return c.label+': '+fmt(c.raw)+'件 ('+pct(c.raw,total)+')'},afterBody:function(ctx){return sizeTitle(sizeAll[ctx[0].dataIndex]||{})}}}}}});
mk('c8cust',{type:'bar',data:{labels:topCustomerSizeRows.map(function(r){return r.size}),datasets:[{label:'净销售量',data:topCustomerSizeRows.map(function(r){return r.qty}),backgroundColor:P.qty,xAxisID:'x',barMaxWidth:22},{label:'净销售金额',data:topCustomerSizeRows.map(function(r){return r.amount}),backgroundColor:P.amount,xAxisID:'x1',barMaxWidth:22}]},options:{responsive:true,maintainAspectRatio:false,indexAxis:'y',grid:{left:76,right:116,top:66,bottom:78},plugins:{legend:{position:'bottom'},tooltip:{callbacks:{label:function(c){return c.dataset.label+': '+(c.dataset.xAxisID==='x1'?money(c.raw):fmt(c.raw)+'件')},afterBody:function(ctx){return customerSizeTooltip(topCustomerSizeRows[ctx[0].dataIndex]||{})}}}},scales:{x:{position:'bottom',title:{display:true,text:'净销售量'},ticks:{callback:function(v){return fmt(v)}}},x1:{position:'top',grid:{drawOnChartArea:false},title:{display:true,text:'净销售金额'},ticks:{callback:function(v){return money(v)}}}}}});
mk('c8c',{type:'bar',data:{labels:sizeByGroup.map(function(r){return r.size}),datasets:[{label:'净销售量',data:sizeByGroup.map(function(r){return r.qty}),backgroundColor:sizeByGroup.map(function(r){return sizeGroupColor(r.groupLabel)}),borderRadius:3}]},options:{responsive:true,maintainAspectRatio:false,indexAxis:'y',plugins:{legend:{display:false},tooltip:{callbacks:{label:function(c){return '净销售量: '+fmt(c.raw)+'件'},afterBody:function(ctx){var r=sizeByGroup[ctx[0].dataIndex]||{};return ['分组 '+esc(r.groupLabel||'-')].concat(sizeTitle(r))}}}},scales:{x:{title:{display:true,text:'净销售量'},ticks:{callback:function(v){return fmt(v)}}}}}});
var trendSizeNames=(SP.trend_sizes||[]).filter(function(size){return n(sizeTrendTotals[size])>0});
mk('c9b',{type:'line',data:{labels:(SP.months||[]),datasets:trendSizeNames.map(function(size,i){return {label:size,data:(SP.months||[]).map(function(m){return sizeTrendMap[size+'|'+m]||0}),borderColor:C[i%C.length],backgroundColor:C[i%C.length],tension:.3,fill:false}})},options:{responsive:true,maintainAspectRatio:false,plugins:{legend:{position:'bottom'}},scales:{y:{title:{display:true,text:'净销售量'},ticks:{callback:function(v){return fmt(v)}}}}}});

var BS=D.brand_style_analysis||{brand_summary:[],designer_summary:[],brand_designer:[],category_designer:[],shop_best_brands:[]};
var brs=positiveRows(brandSummaryRows(),['qty','amount','customers']), drs=positiveRows(designerSummaryRows(),['qty','amount','customers']);
var topCustomerBrandRows=positiveRows(BS.top_customer_brands||[],['qty','amount','customers']);
var shopBestBrands=positiveRows(BS.shop_best_brands||[],['qty','amount','customers']);
document.getElementById('s-br').innerHTML='<div class="tip"><b>品牌/风格洞察：</b>本页只保留标准可视化图表：品牌金额/销量、设计师品牌金额/销量、销售额Top20客户品牌偏好、品类 - 设计师品牌贡献值Top1、品牌 - 设计师贡献金额。</div><div class="g">'
+card('品牌金额与销量',chart('c11',true))
+card('销售额Top20客户品牌偏好',wideChart('c11cust'),true)
+card('设计师品牌金额与销量',brandTallChart('c12'))
+card('各店铺卖得最好的Top1品牌',chart('c11shop',true))
+card('品类 - 设计师品牌贡献值Top1',chart('c16b',true),true)
+card('品牌 - 设计师贡献金额',brandDesignerBoard(),true)
+'</div>';
mk('c11',{type:'bar',data:{labels:brs.map(function(r){return r.brand}),datasets:[{label:'金额',data:brs.map(function(r){return r.amount}),backgroundColor:P.amount},{label:'销量',data:brs.map(function(r){return r.qty}),backgroundColor:P.qty,xAxisID:'x1'}]},options:{responsive:true,maintainAspectRatio:false,indexAxis:'y',plugins:{legend:{position:'bottom'}},scales:{x:{ticks:{callback:function(v){return money(v)}}},x1:{position:'top',grid:{drawOnChartArea:false},ticks:{callback:function(v){return fmt(v)}}}}}});
mk('c11cust',{type:'bar',data:{labels:topCustomerBrandRows.map(function(r){return r.brand}),datasets:[{label:'净销售量',data:topCustomerBrandRows.map(function(r){return r.qty}),backgroundColor:P.qty,xAxisID:'x',barMaxWidth:24},{label:'净销售金额',data:topCustomerBrandRows.map(function(r){return r.amount}),backgroundColor:P.amount,xAxisID:'x1',barMaxWidth:24}]},options:{responsive:true,maintainAspectRatio:false,indexAxis:'y',grid:{left:86,right:126,top:70,bottom:84},plugins:{legend:{position:'bottom'},tooltip:{callbacks:{label:function(c){return c.dataset.label+': '+(c.dataset.xAxisID==='x1'?money(c.raw):fmt(c.raw)+'件')},afterBody:function(ctx){return customerBrandTooltip(topCustomerBrandRows[ctx[0].dataIndex]||{})}}}},scales:{x:{position:'bottom',title:{display:true,text:'净销售量'},ticks:{callback:function(v){return fmt(v)}}},x1:{position:'top',grid:{drawOnChartArea:false},title:{display:true,text:'净销售金额'},ticks:{callback:function(v){return money(v)}}},y:{ticks:{autoSkip:false,font:{size:12}}}}}});
mk('c11shop',{type:'bar',data:{labels:shopBestBrands.map(function(r){return r.shop+' / '+r.brand}),datasets:[{label:'Top1品牌成交金额',data:shopBestBrands.map(function(r){return r.amount}),backgroundColor:P.qty}]},options:{responsive:true,maintainAspectRatio:false,indexAxis:'y',plugins:{legend:{display:false},tooltip:{callbacks:{label:function(c){return '成交金额: '+money(c.raw)},afterBody:function(ctx){return shopBestBrandTooltip(shopBestBrands[ctx[0].dataIndex]||{})}}}},scales:{x:{title:{display:true,text:'成交金额'},ticks:{callback:function(v){return money(v)}}},y:{ticks:{autoSkip:false,font:{size:10}}}}}});
mk('c12',{type:'bar',data:{labels:drs.map(function(r){return r.designer}),datasets:[{label:'金额',data:drs.map(function(r){return r.amount}),backgroundColor:P.amount},{label:'销量',data:drs.map(function(r){return r.qty}),backgroundColor:P.qty,xAxisID:'x1'}]},options:{responsive:true,maintainAspectRatio:false,indexAxis:'y',plugins:{legend:{position:'bottom'}},scales:{x:{ticks:{callback:function(v){return money(v)}}},x1:{position:'top',grid:{drawOnChartArea:false},ticks:{callback:function(v){return fmt(v)}}},y:{ticks:{autoSkip:false,font:{size:10}}}}}});
renderCategoryDesignerTopChart('c16b');

var YS=seasonData(), ysCustomers=positiveRows(YS.customers||[],['total_amount','total_qty']).slice().sort(function(a,b){return n(b.total_amount)-n(a.total_amount)||n(b.total_qty)-n(a.total_qty)}).slice(0,30), ysTypeRows=seasonTypeRows(false);
var ysProducts=seasonHotRows('booking_products',30), ysCats=seasonHotRows('booking_categories',16), ysBrands=seasonHotRows('booking_brands',16), ysDesigners=seasonHotRows('booking_designers',20);
var ysBooking=(YS.order_type_summary||[]).find(function(r){return r.type==='订货会成交'})||{}, ysTotalAmount=(YS.order_type_summary||[]).reduce(function(s,r){return s+n(r.amount)},0), ysTopProduct=ysProducts[0]||{};
document.getElementById('s-ys').innerHTML='<div class="tip"><b>普通款/订货会分析：</b>按【年份】字段判断，包含“订货会”的记录归为订货会成交，其余归为普通款期成交。同一个客户不拆成两个客户，只把成交记录拆成两类。当前订货会成交金额占比 '+pct(ysBooking.amount,ysTotalAmount)+'；订货会爆款按全局商品维度统计，重点看货号销量、金额、成交客户数、订单数和拿货记录数，避免只被少数大客户拉动。</div><div class="g">'
+card('普通款期 / 订货会成交概览',seasonSummaryBoard(),true)
+card('客户订货会类型分布',seasonTypePieBoard('ysType'),true)
+card('客户订货会参与结构（总成交金额Top30）',largeChart('ys1'),true)
+card('订货会爆款货号Top30',largeChart('ys2'),true)
+card('订货会品类表现',chart('ys3',true))
+card('订货会品牌表现',chart('ys4',true))
+card('订货会设计师品牌表现',chart('ys5',true))
+'</div>';
mk('ysType',{type:'pie',data:{labels:ysTypeRows.map(function(r){return r.type}),datasets:[{label:'客户数',data:ysTypeRows.map(function(r){return r.customers}),backgroundColor:ysTypeRows.map(function(r){return seasonTypeColor(r.type)})}]},options:{responsive:true,maintainAspectRatio:false,plugins:{legend:{position:'bottom'},tooltip:{callbacks:{label:function(c){var total=c.dataset.data.reduce(function(a,b){return a+n(b)},0);var r=ysTypeRows[c.dataIndex]||{};return c.label+': '+fmt(c.raw)+'客户 ('+pct(c.raw,total)+')'},afterBody:function(ctx){var r=ysTypeRows[ctx[0].dataIndex]||{};return ['判断规则 '+seasonTypeDef(r.type),'总成交 '+money(r.amount)+' / '+fmt(r.qty)+'件','订货会成交 '+money(r.booking_amount)+' / '+fmt(r.booking_qty)+'件']}}}}}});
mk('ys1',{type:'bar',data:{labels:ysCustomers.map(function(r){return r.customer}),datasets:[{label:'普通款期成交',data:ysCustomers.map(function(r){return r.normal_amount}),backgroundColor:P.qty,stack:'s'},{label:'订货会成交',data:ysCustomers.map(function(r){return r.booking_amount}),backgroundColor:P.amount,stack:'s'}]},options:{responsive:true,maintainAspectRatio:false,plugins:{legend:{position:'bottom'},tooltip:{callbacks:{label:function(c){return c.dataset.label+': '+money(c.raw)},afterBody:function(ctx){return seasonCustomerTooltip(ysCustomers[ctx[0].dataIndex]||{})}}}},scales:{x:{ticks:{autoSkip:false,font:{size:10}}},y:{title:{display:true,text:'净销售金额'},ticks:{callback:function(v){return money(v)}}}}}});
mk('ys2',{type:'bar',data:{labels:ysProducts.map(function(r){return r.product}),datasets:[{label:'订货会成交数量',data:ysProducts.map(function(r){return r.qty}),backgroundColor:P.qty,xAxisID:'x'},{label:'订货会成交金额',data:ysProducts.map(function(r){return r.amount}),backgroundColor:P.amount,xAxisID:'x1'}]},options:{responsive:true,maintainAspectRatio:false,indexAxis:'y',plugins:{legend:{position:'bottom'},tooltip:{callbacks:{label:function(c){return c.dataset.xAxisID==='x1'?'成交金额: '+money(c.raw):'成交数量: '+fmt(c.raw)+'件'},afterBody:function(ctx){return bookingHotTooltip(ysProducts[ctx[0].dataIndex]||{},'product')}}}},scales:{x:{position:'bottom',title:{display:true,text:'订货会成交数量'},ticks:{callback:function(v){return fmt(v)}}},x1:{position:'top',grid:{drawOnChartArea:false},title:{display:true,text:'订货会成交金额'},ticks:{callback:function(v){return money(v)}}}}}});
mk('ys3',{type:'bar',data:{labels:ysCats.map(function(r){return r.category}),datasets:[{label:'订货会成交数量',data:ysCats.map(function(r){return r.qty}),backgroundColor:P.qty,xAxisID:'x'},{label:'订货会成交金额',data:ysCats.map(function(r){return r.amount}),backgroundColor:P.amount,xAxisID:'x1'}]},options:{responsive:true,maintainAspectRatio:false,indexAxis:'y',plugins:{legend:{position:'bottom'},tooltip:{callbacks:{label:function(c){return c.dataset.xAxisID==='x1'?'成交金额: '+money(c.raw):'成交数量: '+fmt(c.raw)+'件'},afterBody:function(ctx){return bookingHotTooltip(ysCats[ctx[0].dataIndex]||{},'category')}}}},scales:{x:{position:'bottom',title:{display:true,text:'订货会成交数量'},ticks:{callback:function(v){return fmt(v)}}},x1:{position:'top',grid:{drawOnChartArea:false},title:{display:true,text:'订货会成交金额'},ticks:{callback:function(v){return money(v)}}}}}});
mk('ys4',{type:'bar',data:{labels:ysBrands.map(function(r){return r.brand}),datasets:[{label:'订货会成交数量',data:ysBrands.map(function(r){return r.qty}),backgroundColor:P.qty,xAxisID:'x'},{label:'订货会成交金额',data:ysBrands.map(function(r){return r.amount}),backgroundColor:P.amount,xAxisID:'x1'}]},options:{responsive:true,maintainAspectRatio:false,indexAxis:'y',plugins:{legend:{position:'bottom'},tooltip:{callbacks:{label:function(c){return c.dataset.xAxisID==='x1'?'成交金额: '+money(c.raw):'成交数量: '+fmt(c.raw)+'件'},afterBody:function(ctx){return bookingHotTooltip(ysBrands[ctx[0].dataIndex]||{},'brand')}}}},scales:{x:{position:'bottom',title:{display:true,text:'订货会成交数量'},ticks:{callback:function(v){return fmt(v)}}},x1:{position:'top',grid:{drawOnChartArea:false},title:{display:true,text:'订货会成交金额'},ticks:{callback:function(v){return money(v)}}}}}});
mk('ys5',{type:'bar',data:{labels:ysDesigners.map(function(r){return r.designer}),datasets:[{label:'订货会成交数量',data:ysDesigners.map(function(r){return r.qty}),backgroundColor:P.qty,xAxisID:'x'},{label:'订货会成交金额',data:ysDesigners.map(function(r){return r.amount}),backgroundColor:P.amount,xAxisID:'x1'}]},options:{responsive:true,maintainAspectRatio:false,indexAxis:'y',plugins:{legend:{position:'bottom'},tooltip:{callbacks:{label:function(c){return c.dataset.xAxisID==='x1'?'成交金额: '+money(c.raw):'成交数量: '+fmt(c.raw)+'件'},afterBody:function(ctx){return bookingHotTooltip(ysDesigners[ctx[0].dataIndex]||{},'designer')}}}},scales:{x:{position:'bottom',title:{display:true,text:'订货会成交数量'},ticks:{callback:function(v){return fmt(v)}}},x1:{position:'top',grid:{drawOnChartArea:false},title:{display:true,text:'订货会成交金额'},ticks:{callback:function(v){return money(v)}}}}}});

var AC=acceptData(), ACK=AC.kpi||{}, ACO=AC.overall||{};
var highCats=(ACK.high_categories||[]).join(' / ')||'-';
var acBands=positiveRows(AC.bands||[],['amount','qty','orders','lines']);
var acSegments=positiveRows(AC.segments||[],['amount','qty','orders','lines']);
document.getElementById('s-se').innerHTML='<div class="tip"><b>价格接受度洞察：</b>低价带 / 主流价带 / 高价带边界按本次导入成交单价的销量加权 P33 / P67 自动计算：低价带 < ¥'+AC.low_cutoff+'，主流价带 ¥'+AC.low_cutoff+'-'+AC.high_cutoff+'，高价带 ≥ ¥'+AC.high_cutoff+'。当前最能接受的价格段是 '+esc(ACK.most_accepted_band||'-')+'，按数量占 '+acceptPct(ACK.most_accepted_qty_share)+'；金额最高价格段是 '+esc(ACK.highest_amount_band||'-')+'，按金额占 '+acceptPct(ACK.highest_amount_share)+'。高价带金额占 '+acceptPct(ACK.high_amount_share)+'，主要由 '+esc(highCats)+' 承接。整体推荐价位可参考 ¥'+(ACK.recommend_low||ACO.p25||0)+'-'+(ACK.recommend_high||ACO.p75||0)+'。</div><div class="g">'
+card('整体价格段成交结构',chart('c13',true))
+card('低价 / 主流 / 高价成交占比',chart('c13b',false)+acceptSegmentNote())
+card('品类价格带成交结构',chart('c13c',true)+acceptCategoryStructureNote())
+card('品类推荐价位条',acceptRangeBoard(18),true)
+'</div>';
mk('c13',{type:'bar',data:{labels:acBands.map(function(r){return r.name}),datasets:[{label:'成交数量',data:acBands.map(function(r){return r.qty}),backgroundColor:P.qty},{label:'成交金额',data:acBands.map(function(r){return r.amount}),backgroundColor:P.amount,yAxisID:'y1'}]},options:{responsive:true,maintainAspectRatio:false,plugins:{legend:{position:'bottom'},tooltip:{callbacks:{label:function(c){var r=acBands[c.dataIndex]||{};return c.dataset.label+': '+(c.dataset.yAxisID==='y1'?money(c.raw):fmt(c.raw)+'件')+' / 数量占比 '+acceptPct(r.qty_share)+' / 金额占比 '+acceptPct(r.amount_share)}}}},scales:{y:{title:{display:true,text:'成交数量'},ticks:{callback:function(v){return fmt(v)}}},y1:{position:'right',grid:{drawOnChartArea:false},title:{display:true,text:'成交金额'},ticks:{callback:function(v){return money(v)}}}}}});
mk('c13b',{type:'doughnut',data:{labels:acSegments.map(function(r){return r.name}),datasets:[{label:'金额占比',data:acSegments.map(function(r){return r.amount}),backgroundColor:acSegments.map(function(r){return acceptSegColor(r.name)}),borderColor:'#fff',borderWidth:1}]},options:{responsive:true,maintainAspectRatio:false,plugins:{legend:{position:'right'},tooltip:{formatter:function(c){var total=(c.dataset&&c.dataset.data?c.dataset.data:[]).reduce(function(a,b){return a+n(b)},0),r=acSegments[c.dataIndex]||{},amountShare=pct(c.value,total),qtyShare=acceptPct(r.qty_share),color=acceptSegColor(r.name||c.name);return '<div style="min-width:188px"><div style="display:flex;align-items:center;gap:7px;font-weight:800;margin-bottom:8px;color:#fff"><i style="display:inline-block;width:10px;height:10px;border-radius:2px;background:'+color+'"></i>'+esc(c.name)+'</div><div style="color:rgba(255,255,255,.72);font-size:12px">成交金额</div><div style="font-size:16px;font-weight:800;line-height:1.35;color:#fff">'+money(c.value)+'</div><div style="display:grid;grid-template-columns:1fr 1fr;gap:8px;margin-top:9px"><div style="background:rgba(255,255,255,.12);border:1px solid rgba(255,255,255,.16);border-radius:7px;padding:7px 8px"><span style="display:block;color:rgba(255,255,255,.72);font-size:11px">金额占比</span><b style="display:block;color:#fff;line-height:1.2;margin-top:2px">'+percentHtml(amountShare)+'</b></div><div style="background:rgba(255,255,255,.12);border:1px solid rgba(255,255,255,.16);border-radius:7px;padding:7px 8px"><span style="display:block;color:rgba(255,255,255,.72);font-size:11px">数量占比</span><b style="display:block;color:#fff;line-height:1.2;margin-top:2px">'+percentHtml(qtyShare)+'</b></div></div></div>'}}}}}});
var acCats=positiveRows(AC.categories||[],['amount','qty','orders','customers']).slice().sort(function(a,b){return n(b.amount)-n(a.amount)}).slice(0,16);
var acSegmentSeries=[{label:'低价带',field:'low_amount',share:'low_amount_share'},{label:'主流价带',field:'main_amount',share:'main_amount_share'},{label:'高价带',field:'high_amount',share:'high_amount_share'}].filter(function(s){return acCats.some(function(r){return n(r[s.field])>0})});
mk('c13c',{type:'bar',data:{labels:acCats.map(function(r){return r.category}),datasets:acSegmentSeries.map(function(s){return {label:s.label,data:acCats.map(function(r){return r[s.field]}),backgroundColor:acceptSegColor(s.label),stack:'price',shareField:s.share}})},options:{responsive:true,maintainAspectRatio:false,indexAxis:'y',plugins:{legend:{position:'bottom'},tooltip:{callbacks:{label:function(c){var r=acCats[c.dataIndex]||{},share=c.dataset.shareField?r[c.dataset.shareField]:0;return c.dataset.label+': '+money(c.raw)+' / 品类内金额占 '+acceptPct(share)}}}},scales:{x:{stacked:true,ticks:{callback:function(v){return money(v)}},title:{display:true,text:'各品类低价 / 主流 / 高价成交金额'}},y:{stacked:true}}}});

document.getElementById('s-pf').innerHTML='<div class="tip"><b>客户画像：</b>选择客户后查看其品类、品牌设计师、颜色、尺码、价格区间和年份季节；品牌默认全部品牌，也可在选定客户后进一步搜索该客户拿过货的品牌，画像卡会按该客户在该品牌下的成交重新计算。订单数按销退单ID去重计算；颜色饼图展示该客户净拿货数量Top15颜色，鼠标悬停颜色后可查看该颜色下的品类占比Top5和品牌占比Top5；价格区间按品类拆分展示，便于判断不同品类的主成交价带。</div><div class="pf-head"><label>客户<select id="pfCustomer">'+pfSelectOptions()+'</select></label><label>品牌<div class="pf-brand-box" id="pfBrandBox"><input id="pfBrandInput" type="search" autocomplete="off" placeholder="全部品牌 / 搜索该客户拿过的品牌"><div class="pf-brand-menu" id="pfBrandMenu"></div></div></label></div><div class="pf-mini" id="pfMini"></div><div class="g">'
+card('客户推荐画像卡','<div class="pf-rec" id="pfRecommend"></div>')
+card('复拿商品排行',largeChart('pfRepeatProducts'),true)
+card('全部品类偏好金额',chart('pfCat',false))
+card('全部品牌 / 设计师偏好',chart('pfBrand',true))
+card('颜色拿货数量占比Top15','<div class="ch pf-color-chart"><div id="pfColorPie" class="echart"></div></div>')
+card('全部尺码偏好',chart('pfSize',false))
+card('品类价格区间偏好',chart('pfPrice',true),true)
+card('年份季节',chart('pfSeason',false))
+'</div>';
refreshPfBrandOptions();
document.getElementById('pfCustomer').onchange=function(){PF_BRAND='';refreshPfBrandOptions();renderCustomerProfile()};
document.getElementById('pfBrandInput').onfocus=function(){renderPfBrandMenu(true);pfBrandBoxOpen(true)};
document.getElementById('pfBrandInput').onclick=function(){renderPfBrandMenu(true);pfBrandBoxOpen(true)};
document.getElementById('pfBrandInput').oninput=function(){PF_BRAND='';renderPfBrandMenu(false);renderCustomerProfile();pfBrandBoxOpen(true)};
document.getElementById('pfBrandInput').onkeydown=function(e){if(e.key==='Enter'){e.preventDefault();var rows=pfFilteredBrandRows(false);setPfBrand(rows[0]?.name||'')}else if(e.key==='Escape'){pfBrandBoxOpen(false)}};
document.getElementById('pfBrandMenu').onclick=function(e){var btn=e.target.closest('[data-brand-option]');if(!btn)return;setPfBrand(btn.getAttribute('data-brand')||'')};
document.addEventListener('click',function(e){var box=document.getElementById('pfBrandBox');if(box&&!box.contains(e.target))pfBrandBoxOpen(false)});
renderCustomerProfile();
hydrateUi();
</script></body></html>`;
}
