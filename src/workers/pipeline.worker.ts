// pipeline.worker.ts — 后台跑完整分析 pipeline 的 Web Worker。
// 日志文案/格式/步骤号/时间戳对齐 Python 版(backend/main.py 的 SSE + pipeline.py 的 log),
// 方便 A/B 对照。流程:读 Excel → 聚合 → Canvas 画图(并行编码)→ ExcelJS 组装。
import { readSales, readZhixiao } from '../pipeline/reader';
import { loadPreference } from '../pipeline/preferenceReader';
import { analyzePreference, TOP_N } from '../pipeline/preferenceAnalyze';
import { buildNewPreferenceHtml, loadPreferenceOrderIds } from '../pipeline/preferenceNewHtml';
import { buildPreferenceExcel } from '../pipeline/preferenceExcel';
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
let curStep = 1; // 同 main.py 的 job.step:跟随最近一条日志

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

/** 并发跑任务(PNG 编码是异步的,适度并行能明显提速) */
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
    if (msg.mode === 'preference') {
      await runPreference(msg.files);
      return;
    }
    const zhixiaoFile = msg.files.find((f) => f.role === '滞销表');
    const salesFile = msg.files.find((f) => f.role === '销售明细');
    if (!zhixiaoFile || !salesFile) {
      throw new PipelineError('缺少输入文件:需要滞销商品 + 各商品客户拿货历史各一份');
    }

    log('✓ 分析引擎已就绪', 'normal', 1);

    // ---------- 步骤 1:识别 ----------
    // 以下日志文案(含全角标点)逐字符对齐 pipeline.py,A/B 对照时能直接 diff
    banner('步骤 1：识别输入文件', 1);
    log(`  滞销商品：           ${zhixiaoFile.name}`, 'normal', 1);
    log(`  各商品客户拿货历史:  ${salesFile.name}`, 'normal', 1);

    // ---------- 步骤 2:读取 + 聚合 ----------
    banner('步骤 2：聚合销售数据（取净销售量/净销售金额）', 2);

    const zhixiao = readZhixiao(zhixiaoFile.buffer);
    // Python 版此处 df_zi 已加内部列 货号_k,所以列数 +1(对照用,保持一致)
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
      // 钉位判定结果直接进日志:排查"配置了不生效"时一眼定位是表头问题还是代码问题
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

    // ---------- 步骤 3:生成 PNG(对应 pipeline.py:283-434)----------
    banner('步骤 3：生成每款销售趋势图', 3);
    log('  使用渲染器: OffscreenCanvas（Python 版为 matplotlib Agg）', 'normal', 3);

    const totalN = agg.items.length;
    const PROGRESS_EVERY = Math.max(1, Math.floor(totalN / 8));
    const t0 = performance.now();

    // 优化 A:零销量款不生成 PNG,Excel 趋势列直接显示“无销售”
    const isZeroArr = agg.items.map((item) => item.销售量 === 0);

    /** 货号原始顺序 idx → {sm, dt} PNG 字节(写 Excel 时按 idx 取) */
    const images = new Map<number, { sm: Uint8Array; dt: Uint8Array }>();
    let doneCount = 0;
    await runPool(totalN, 8, async (i) => {
      if (isZeroArr[i]) {
        // 零销量行不写图片对象,减少 Windows Excel/WPS 渲染压力
      } else {
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

    // 样图预览:浏览器里目检 Canvas 复刻效果(第一个有销量款)
    const firstSalesIdx = agg.items.findIndex((it) => it.销售量 > 0);
    const previews: { label: string; png: ArrayBuffer }[] = [];
    if (firstSalesIdx >= 0) {
      const im = images.get(firstSalesIdx)!;
      previews.push({ label: `${agg.items[firstSalesIdx].货号}(小图)`, png: im.sm.slice().buffer });
      previews.push({ label: `${agg.items[firstSalesIdx].货号}(详情图)`, png: im.dt.slice().buffer });
    }
    post({ type: 'preview', images: previews }, previews.map((p) => p.png));

    // ---------- 步骤 4:组装 Excel(对应 pipeline.py:436-755)----------
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
    // 对齐 main.py:日志打 "✗ {e}",error 对象带 {message, hint, step}
    const isPipe = err instanceof PipelineError;
    const message = isPipe ? err.message : `内部错误：${err instanceof Error ? err.message : err}`;
    const hint = isPipe ? err.hint : '';
    log(`  ✗ ${message}`, 'error', curStep);
    post({ type: 'error', message, hint, step: curStep });
  }
};

// ── 客户偏好分析(对应 backend/preference_pipeline.py)────────────────────
// M9 是 stub:验证 tab/Worker/Result 全链路。M10 接 reader,M11 聚合,M12 出新版 html+xlsx。
async function runPreference(files: { role: string; name: string; buffer: ArrayBuffer }[]) {
  const input = files.find((f) => f.role === '拿货历史');
  if (!input) {
    throw new PipelineError('缺少输入文件:需要一份「各商品客户拿货历史」');
  }

  banner('客户偏好分析', 1);
  log(`  输入：${input.name}`, 'normal', 1);
  const data = loadPreference(input.buffer);
  // 日志格式对齐 preference_pipeline.py:111
  log(`  读到 ${fmtN(data.rawRowCount)} 条原始记录（${data.sheetNames.length} 个 sheet）`, 'normal', 1);
  if (data.droppedConflicts.length > 0) {
    log(`  丢弃毛值冲突列: ${data.droppedConflicts.join(', ')}（取净值）`, 'normal', 1);
  }
  log(`  规范化后 ${data.columns.length} 列: ${data.columns.join(' / ')}`, 'normal', 1);
  banner('聚合与画像', 2);
  const R = analyzePreference(data);
  const sm = R.summary as Record<string, number | string>;
  // 日志格式对齐 preference_pipeline.py:222-227
  log(
    `  有效交易 ${fmtN(sm.records as number)} 条 | ` +
    `客户 ${fmtN(sm.customers as number)} 位 | ` +
    `销售额 ¥${((sm.amount as number) / 10000).toFixed(0)}万`,
    'milestone', 2,
  );
  // ---------- 步骤 3:生成新客户偏好分析 ----------
  banner('生成新客户偏好分析', 3);
  const orderIds = loadPreferenceOrderIds(input.buffer);
  const newHtmlBytes = new TextEncoder().encode(buildNewPreferenceHtml(data, orderIds, R)).buffer as ArrayBuffer;
  log(`  ✓ 新客户偏好分析已生成（${(newHtmlBytes.byteLength / 1024).toFixed(0)} KB）`, 'done', 3);

  // ---------- 步骤 4:生成 Excel 数据表 ----------
  banner('生成 Excel 数据表', 4);
  const xlsxU8 = await buildPreferenceExcel(data);
  const kb = xlsxU8.byteLength / 1024;
  const sizeStr = kb < 1024 ? `${kb.toFixed(0)} KB` : `${(kb / 1024).toFixed(1)} MB`;
  log(`  ✓ Excel 已生成（${sizeStr}，含约 20 个 sheet）`, 'done', 4);

  banner('完成', 5);
  log(`  全部 ${fmtN(sm.customers as number)} 位客户的画像可在 Excel「9-客户画像汇总」查看`, 'milestone', 5);
  log(`  网页可视化仍只展示前 ${TOP_N} 大客户（聚焦决策）`, 'normal', 5);
  log('✅ 完成', 'done', 5);

  const xlsxBuf = xlsxU8.buffer as ArrayBuffer;
  post(
    {
      type: 'done',
      mode: 'preference',
      newHtml: { filename: '新客户偏好分析.html', buffer: newHtmlBytes },
      xlsx: { filename: '客户偏好分析数据.xlsx', buffer: xlsxBuf },
      summary: {
        records: sm.records as number,
        customers: sm.customers as number,
        amount: sm.amount as number,
        dateFrom: sm.date_from as string,
        dateTo: sm.date_to as string,
      },
    },
    [newHtmlBytes, xlsxBuf],
  );
}
