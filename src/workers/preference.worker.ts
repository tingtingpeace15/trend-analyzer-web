// preference.worker.ts — 客户偏好分析专用 Worker。
// 与商品趋势 Worker 拆分,避免趋势分析初始化时加载 ECharts/偏好看板代码。
import { loadPreference } from '../pipeline/preferenceReader';
import { analyzePreference, TOP_N } from '../pipeline/preferenceAnalyze';
import { buildNewPreferenceHtml, loadPreferenceOrderIds } from '../pipeline/preferenceNewHtml';
import { buildPreferenceExcel } from '../pipeline/preferenceExcel';
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

self.onmessage = async (e: MessageEvent<MainToWorkerMessage>) => {
  const msg = e.data;
  if (msg.type !== 'run') return;
  runStart = performance.now();
  curStep = 1;

  try {
    if (msg.mode !== 'preference') {
      throw new PipelineError('分析模式错误:当前 Worker 只支持客户偏好分析');
    }

    const input = msg.files.find((f) => f.role === '拿货历史');
    if (!input) {
      throw new PipelineError('缺少输入文件:需要一份「各商品客户拿货历史」');
    }

    banner('客户偏好分析', 1);
    log(`  输入：${input.name}`, 'normal', 1);
    const data = loadPreference(input.buffer);
    log(`  读到 ${fmtN(data.rawRowCount)} 条原始记录（${data.sheetNames.length} 个 sheet）`, 'normal', 1);
    if (data.droppedConflicts.length > 0) {
      log(`  丢弃毛值冲突列: ${data.droppedConflicts.join(', ')}（取净值）`, 'normal', 1);
    }
    log(`  规范化后 ${data.columns.length} 列: ${data.columns.join(' / ')}`, 'normal', 1);

    banner('聚合与画像', 2);
    const R = analyzePreference(data);
    const sm = R.summary as Record<string, number | string>;
    log(
      `  有效交易 ${fmtN(sm.records as number)} 条 | ` +
      `客户 ${fmtN(sm.customers as number)} 位 | ` +
      `销售额 ¥${((sm.amount as number) / 10000).toFixed(0)}万`,
      'milestone', 2,
    );

    banner('生成新客户偏好分析', 3);
    const orderIds = loadPreferenceOrderIds(input.buffer);
    const newHtmlBytes = new TextEncoder().encode(buildNewPreferenceHtml(data, orderIds, R)).buffer as ArrayBuffer;
    log(`  ✓ 新客户偏好分析已生成（${(newHtmlBytes.byteLength / 1024).toFixed(0)} KB）`, 'done', 3);

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
  } catch (err) {
    const isPipe = err instanceof PipelineError;
    const message = isPipe ? err.message : `内部错误：${err instanceof Error ? err.message : err}`;
    const hint = isPipe ? err.hint : '';
    log(`  ✗ ${message}`, 'error', curStep);
    post({ type: 'error', message, hint, step: curStep });
  }
};
