// api.ts — 主线程跟 pipeline Worker 的通信封装。
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

function createWorker(): Worker {
  return new Worker(
    new URL('./workers/pipeline.worker.ts', import.meta.url),
    { type: 'module' },
  );
}

// Worker bundle 有 ~1.3MB(SheetJS+ExcelJS),按需创建时"下载+编译"会让
// 用户点了开始后干等好几秒。池子里常备一个预热好的实例:页面加载即预热,
// 任务正常结束后归还复用;cancel/error 才销毁。
let pooled: Worker | null = null;

/** 页面空闲时调用:提前下载+编译 Worker,点「开始分析」即刻可用 */
export function prewarmWorker(): void {
  if (!pooled) pooled = createWorker();
}

/**
 * 启动一次分析:读文件 → 以 transferable ArrayBuffer 传给 Worker → 把 Worker 消息转回调。
 * cancel() 直接 terminate Worker(没有后端任务要清理)。
 */
export function startJob(mode: AnalysisMode, inputs: JobInputFile[], cb: JobCallbacks): JobHandle {
  const worker = pooled ?? createWorker();
  pooled = null;
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
        if (!pooled) pooled = worker;
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
