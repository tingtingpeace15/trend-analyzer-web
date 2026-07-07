// api.ts — 主线程跟分析 Worker 的通信封装。
// 旧版 api.jsx 是 fetch + SSE;网页版没有后端,这里换成 Worker 消息,但对 UI 暴露同样的"任务 + 日志流"心智模型。
import type {
  AnalysisMode,
  DoneMessage,
  FileRole,
  LogMessage,
  WorkerToMainMessage,
} from './types/pipeline';

/** done 消息去掉 type 字段后的结果载荷(趋势/偏好两种形态) */
export type JobDone = Omit<Extract<DoneMessage, { mode: 'trend' }>, 'type'>
  | Omit<Extract<DoneMessage, { mode: 'preference' }>, 'type'>;

export interface JobInputFile {
  role: FileRole;
  file: File;
}

export interface PreviewImage {
  label: string;
  png: ArrayBuffer;
}

/** 对齐 Python 版 main.py 的 job.error 结构 */
export interface JobError {
  message: string;
  hint: string;
  step: number;
}

export interface JobCallbacks {
  onLog: (log: LogMessage) => void;
  onPreview?: (images: PreviewImage[]) => void;
  onDone: (done: JobDone) => void;
  onError: (error: JobError) => void;
}

export interface JobHandle {
  cancel: () => void;
}

function createWorker(mode: AnalysisMode): Worker {
  if (mode === 'preference') {
    return new Worker(
      new URL('./workers/preference.worker.ts', import.meta.url),
      { type: 'module' },
    );
  }

  return new Worker(
    new URL('./workers/trend.worker.ts', import.meta.url),
    { type: 'module' },
  );
}

// 趋势/偏好依赖差异很大,拆成两个 Worker:商品趋势不再加载偏好看板/ECharts。
// 池子按模式各常备一个预热实例,任务正常结束后归还复用;cancel/error 才销毁。
const pooled: Partial<Record<AnalysisMode, Worker>> = {};

/** 页面空闲时调用:提前下载+编译 Worker,点「开始分析」即刻可用 */
export function prewarmWorker(mode: AnalysisMode = 'trend'): void {
  if (!pooled[mode]) pooled[mode] = createWorker(mode);
}

/**
 * 启动一次分析:读文件 → 以 transferable ArrayBuffer 传给 Worker → 把 Worker 消息转回调。
 * cancel() 直接 terminate Worker(没有后端任务要清理)。
 */
export function startJob(mode: AnalysisMode, inputs: JobInputFile[], cb: JobCallbacks): JobHandle {
  const worker = pooled[mode] ?? createWorker(mode);
  delete pooled[mode];
  let alive = true;

  worker.onmessage = (e: MessageEvent<WorkerToMainMessage>) => {
    if (!alive) return;
    const msg = e.data;
    switch (msg.type) {
      case 'log':
        cb.onLog(msg);
        break;
      case 'preview':
        cb.onPreview?.(msg.images);
        break;
      case 'done':
        alive = false;
        // 任务正常完成:Worker 归还池子复用(每次 run 自身会重置状态)
        worker.onmessage = null;
        worker.onerror = null;
        if (!pooled[mode]) pooled[mode] = worker;
        cb.onDone(msg);
        break;
      case 'error':
        cb.onError({ message: msg.message, hint: msg.hint, step: msg.step });
        worker.terminate();
        alive = false;
        break;
    }
  };
  worker.onerror = (e) => {
    if (!alive) return;
    cb.onError({ message: e.message || 'Worker 异常', hint: '', step: 1 });
    worker.terminate();
    alive = false;
  };

  // 异步读文件字节再发给 Worker;buffer 用 transfer 移交所有权,避免大文件拷贝
  (async () => {
    try {
      const files = await Promise.all(
        inputs.map(async ({ role, file }) => ({
          role,
          name: file.name,
          buffer: await file.arrayBuffer(),
        })),
      );
      if (!alive) return;
      worker.postMessage(
        { type: 'run', mode, files },
        files.map((f) => f.buffer),
      );
    } catch (err) {
      if (!alive) return;
      cb.onError({ message: err instanceof Error ? err.message : String(err), hint: '', step: 1 });
      worker.terminate();
      alive = false;
    }
  })();

  return {
    cancel: () => {
      alive = false;
      worker.terminate();
    },
  };
}
