// 生成 JS 版输出 xlsx(M7 全字段 diff 用)。Node 里没有 OffscreenCanvas,
// 嵌图用 1×1 占位 PNG——diff 不比图片字节(Canvas vs matplotlib 必然不同),
// 比的是锚点/数量;图片内容靠浏览器目检 + 几何测试保障。
//
// 用法: npx vite-node scripts/gen_js_output.ts
// 输出: baseline/js/商品销售趋势.xlsx
import { mkdirSync, readFileSync, writeFileSync } from 'node:fs';
import { resolve } from 'node:path';
import { readSales, readZhixiao } from '../src/pipeline/reader';
import { aggregate } from '../src/pipeline/aggregator';
import { writeExcel } from '../src/pipeline/writer';
import type { ItemImages } from '../src/pipeline/writer';

const ROOT = resolve(import.meta.dirname, '..');
const BASE_DIR = resolve(ROOT, '..', '趋势分析工具 -最终app版本');
const OUT_DIR = resolve(ROOT, 'baseline', 'js');

const TINY_PNG = new Uint8Array(Buffer.from(
  'iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJAAAADUlEQVR42mNk+M9QDwADhgGAWjR9awAAAABJRU5ErkJggg==',
  'base64',
));

const t0 = Date.now();
const zhixiao = readZhixiao(new Uint8Array(readFileSync(
  resolve(BASE_DIR, '滞销商品【按销售】_njfayxbugwDqARB.xlsx'))));
const sales = readSales(new Uint8Array(readFileSync(
  resolve(BASE_DIR, '各商品客户拿货历史_5.12.xlsx'))));
const agg = aggregate(zhixiao, sales);

// 零销量款共用同一个 Uint8Array 引用(同 Worker 的 placeholder 行为 → 媒体去重一致)
const placeholder = TINY_PNG;
const unique = new Uint8Array(TINY_PNG); // 有销量款:每款独立字节(模拟真实独立 PNG)
const images = new Map<number, ItemImages>();
agg.items.forEach((item, i) => {
  const png = item.销售量 > 0 ? new Uint8Array(unique) : placeholder;
  images.set(i, { sm: png, dt: png });
});

const bytes = await writeExcel(agg, images, (text, kind, step) => {
  console.log(`[${step}][${kind}] ${text}`);
});
mkdirSync(OUT_DIR, { recursive: true });
writeFileSync(resolve(OUT_DIR, '商品销售趋势.xlsx'), bytes);
console.log(`js output done in ${((Date.now() - t0) / 1000).toFixed(1)}s, ${(bytes.length / 1024 / 1024).toFixed(2)} MB`);
