// trend.worker.ts — 商品趋势分析专用 Worker。
// 只打包趋势分析依赖,避免商品趋势初始化时加载客户偏好分析/ECharts 代码。
import { readSales, readZhixiao } from '../pipeline/reader';
import { aggregate } from '../pipeline/aggregator';
import { buildDetailScene, buildSmallScene, sceneToPng } from '../pipeline/chart';
import { splitPinnedFields, writeExcel } from '../pipeline/writer';
import { PipelineError } from '../pipeline/errors';
import type {
  LogKind,
  MainToWorkerMessage,
  WorkerToMainMessage,
} from '../types/pipeline';

function post(msg: WorkerToMainMessage, transfer: Transferable[] = []) {
  (self as unknown as Worker).postMessage(msg, transfer);
}

let runStart = 0;
let curStep = 1;

function log(text: string, kind: LogKind = 'normal', step = 1) {
  curStep = step;
  post({ type: 'log', text, kind, step, t: (performance.now() - runStart) / 1000 });
}

function banner(msg: string, step: number) {
  const rule = '='.repeat(44);
  log(rule, 'rule', step);
  log(` ${msg}`, 'header', step);
  log(rule, 'rule', step);
}

const fmtN = (n: number) => n.toLocaleString('en-US');

async function runPool(total: number, concurrency: number, task: (i: number) => Promise<void>) {
  let next = 0;
  const lane = async () => {
    while (next < total) {
      const i = next++;
      await task(i);
    }
  };
  await Promise.all(Array.from({ length: Math.min(concurrency, total) }, lane));
}

self.onmessage = async (e: MessageEvent<MainToWorkerMessage>) => {
  const msg = e.data;
  if (msg.type !== 'run') return;
  runStart = performance.now();
  curStep = 1;

  try {
    if (msg.mode !== 'trend') {
      throw new PipelineError('分析模式错误:当前 Worker 只支持商品趋势分析');
    }

    const zhixiaoFile = msg.files.find((f) => f.role === '滞销表');
    const salesFile = msg.files.find((f) => f.role === '销售明细');
    if (!zhixiaoFile || !salesFile) {
      throw new PipelineError('缺少输入文件:需要滞销商品 + 各商品客户拿货历史各一份');
    }

    log('✓ 分析引擎已就绪', 'normal', 1);

    banner('步骤 1：识别输入文件', 1);
    log(`  滞销商品：           ${zhixiaoFile.name}`, 'normal', 1);
    log(`  各商品客户拿货历史:  ${salesFile.name}`, 'normal', 1);

    banner('步骤 2：聚合销售数据（取净销售量/净销售金额）', 2);

    const zhixiao = readZhixiao(zhixiaoFile.buffer);
    log(`  滞销商品: ${fmtN(zhixiao.rows.length)} 行 × ${zhixiao.columns.length + 1} 列`, 'normal', 2);

    const sales = readSales(salesFile.buffer, salesFile.name);
    log(
      `  拿货历史: ${fmtN(sales.rawRowCount)} 行 × ${sales.columns.length} 列（${sales.sheetNames.length} 个 sheet）`,
      'normal', 2,
    );

    if (sales.records.length === 0) {
      throw new PipelineError(
        '拿货历史清洗后没有有效记录',
        '请检查「下单时间」列是否为日期格式、「货号」列是否有值。',
      );
    }

    const agg = aggregate(zhixiao, sales);
    log(
      `  时间窗（来自拿货历史的下单时间）: ${agg.startDate} ~ ${agg.endDate} （${agg.windowDays} 天）`,
      'normal', 2,
    );
    if (agg.extraFields.length > 0) {
      const pyList = `[${agg.extraFields.map((f) => `'${f}'`).join(', ')}]`;
      log(`  滞销表透传 ${agg.extraFields.length} 个额外字段: ${pyList}`, 'milestone', 2);
      const { pinnedAfterQty } = splitPinnedFields(agg.extraFields);
      if (pinnedAfterQty.length > 0) {
        log(`  钉位字段(置于销售量之后): [${pinnedAfterQty.map((f) => `'${f}'`).join(', ')}]`, 'milestone', 2);
      } else {
        log('  钉位字段: 未命中(透传里没有 A价,全部按默认位置排)', 'normal', 2);
      }
    }
    log(`  总货号数（来自滞销商品）: ${agg.itemsTotal}`, 'milestone', 2);
    log(`    有销量款: ${agg.itemsWithSales}`, 'normal', 2);
    log(`    零销量款（趋势列显示“无销售”,不嵌图）: ${agg.itemsTotal - agg.itemsWithSales}`, 'normal', 2);

    banner('步骤 3：生成每款销售趋势图', 3);
    log('  使用渲染器: OffscreenCanvas（Python 版为 matplotlib Agg）', 'normal', 3);

    const totalN = agg.items.length;
    const PROGRESS_EVERY = Math.max(1, Math.floor(totalN / 8));
    const t0 = performance.now();
    const isZeroArr = agg.items.map((item) => item.销售量 === 0);

    const images = new Map<number, { sm: Uint8Array; dt: Uint8Array }>();
    let doneCount = 0;
    await runPool(totalN, 8, async (i) => {
      if (!isZeroArr[i]) {
        const values = agg.dailyByItem.get(agg.items[i].货号)!;
        const [sm, dt] = await Promise.all([
          sceneToPng(buildSmallScene(values)),
          sceneToPng(buildDetailScene(values, agg.dateRange)),
        ]);
        images.set(i, { sm, dt });
      }
      doneCount += 1;
      if (doneCount % PROGRESS_EVERY === 0 || doneCount === totalN) {
        log(`  已生成 ${doneCount}/${totalN}`, 'progress', 3);
      }
    });
    const chartSecs = ((performance.now() - t0) / 1000).toFixed(1);
    log(`  画图耗时 ${chartSecs}s（${totalN} 款,零销量不嵌图,8 路并行）`, 'normal', 3);

    const firstSalesIdx = agg.items.findIndex((it) => it.销售量 > 0);
    const previews: { label: string; png: ArrayBuffer }[] = [];
    if (firstSalesIdx >= 0) {
      const im = images.get(firstSalesIdx)!;
      previews.push({ label: `${agg.items[firstSalesIdx].货号}(小图)`, png: im.sm.slice().buffer });
      previews.push({ label: `${agg.items[firstSalesIdx].货号}(详情图)`, png: im.dt.slice().buffer });
    }
    post({ type: 'preview', images: previews }, previews.map((p) => p.png));

    banner('步骤 4：组装 Excel（3 个 sheet + 嵌图）', 4);
    const xlsx = await writeExcel(agg, images, log);
    log(`  打包  ~ ${(xlsx.byteLength / 1024 / 1024).toFixed(2)} MB`, 'normal', 4);

    banner('步骤 5：完成', 5);
    log('✅ 完成', 'done', 5);

    const outBuffer = xlsx.buffer as ArrayBuffer;
    post(
      {
        type: 'done',
        mode: 'trend',
        filename: '商品销售趋势.xlsx',
        buffer: outBuffer,
        summary: {
          items: agg.itemsTotal,
          itemsWithSales: agg.itemsWithSales,
          windowFrom: agg.startDate,
          windowTo: agg.endDate,
          windowDays: agg.windowDays,
          sizeBytes: xlsx.byteLength,
        },
      },
      [outBuffer],
    );
  } catch (err) {
    const isPipe = err instanceof PipelineError;
    const message = isPipe ? err.message : `内部错误：${err instanceof Error ? err.message : err}`;
    const hint = isPipe ? err.hint : '';
    log(`  ✗ ${message}`, 'error', curStep);
    post({ type: 'error', message, hint, step: curStep });
  }
};
