// 主线程 ↔ Web Worker 的消息协议。
// M1 只有「收到文件确认」;M2+ 逐步扩成完整 pipeline 进度/结果。

/** 文件角色,沿用旧版 mockApi.jsx 的命名 */
export type FileRole = '滞销表' | '销售明细';

/** 主线程 → Worker:启动分析任务 */
export interface RunJobMessage {
  type: 'run';
  files: WorkerFileInput[];
}

export interface WorkerFileInput {
  role: FileRole;
  name: string;
  /** Excel 原始字节,用 transferable 传入 Worker 避免拷贝 */
  buffer: ArrayBuffer;
}

export type MainToWorkerMessage = RunJobMessage;

/** Worker → 主线程:进度日志(对齐 Python 版 SSE 日志的 kind 分类) */
export type LogKind = 'normal' | 'rule' | 'header' | 'milestone' | 'progress' | 'done' | 'error';

export interface LogMessage {
  type: 'log';
  text: string;
  kind: LogKind;
  /** 对应 Python 版的步骤号(1=识别输入 2=聚合 3=画图 4=写 Excel 5=完成) */
  step: number;
  /** 任务启动以来的秒数(对应 Python 版 SSE 的 t 字段) */
  t: number;
}

/** Worker → 主线程:样图预览(M4 起;浏览器里目检 Canvas 图 vs Python 图) */
export interface PreviewMessage {
  type: 'preview';
  images: { label: string; png: ArrayBuffer }[];
}

/** 分析结果摘要(Result 页展示,对应 Python 版 run_pipeline 返回值) */
export interface JobSummary {
  items: number;
  itemsWithSales: number;
  windowFrom: string;
  windowTo: string;
  windowDays: number;
  sizeBytes: number;
}

/** Worker → 主线程:任务完成,携带结果 xlsx 字节 */
export interface DoneMessage {
  type: 'done';
  filename: string;
  buffer: ArrayBuffer;
  summary: JobSummary;
}

/** 对应 main.py 的 job.error = {message, hint, step} */
export interface ErrorMessage {
  type: 'error';
  message: string;
  hint: string;
  step: number;
}

export type WorkerToMainMessage = LogMessage | PreviewMessage | DoneMessage | ErrorMessage;
