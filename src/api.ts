// api.ts — 主线程跟 pipeline Worker 的通信封装。
// 旧版 api.jsx 是 fetch + SSE;网页版没有后端,这里换成 Worker 消息,但对 UI 暴露同样的"任务 + 日志流"心智模型。
import type {
  FileRole,
  JobSummary,
  LogMessage,
  WorkerToMainMessage,
} from './types/pipeline';

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
  onDone: (filename: string, buffer: ArrayBuffer, summary: JobSummary) => void;
  onError: (error: JobError) => void;
}

export interface JobHandle {
  cancel: () => void;
}

/**
 * 启动一次分析:读文件 → 以 transferable ArrayBuffer 传给 Worker → 把 Worker 消息转回调。
 * cancel() 直接 terminate Worker(没有后端任务要清理)。
 */
export function startJob(inputs: JobInputFile[], cb: JobCallbacks): JobHandle {
  const worker = new Worker(
    new URL('./workers/pipeline.worker.ts', import.meta.url),
    { type: 'module' },
  );
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
        cb.onDone(msg.filename, msg.buffer, msg.summary);
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
        { type: 'run', files },
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
