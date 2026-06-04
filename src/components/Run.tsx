// Run.tsx — 处理中页面:步骤指示器 + 流式终端日志,复刻旧版 Run.jsx。
// 数据来源从 SSE 换成了 Worker 消息,UI 结构/配色/动画原样保留。
import { Fragment, useEffect, useMemo, useRef } from 'react';
import { IconAlert, IconRefresh } from './icons';
import type { LogMessage } from '../types/pipeline';
import type { JobError, PreviewImage } from '../api';

const STEPS = [
  { n: 1, label: '识别文件' },
  { n: 2, label: '聚合销售' },
  { n: 3, label: '生成趋势图' },
  { n: 4, label: '组装 Excel' },
  { n: 5, label: '完成' },
];

function StepProgress({ current, errored, errorStep, subProgress }: {
  current: number;
  errored: boolean;
  errorStep: number;
  subProgress: string;
}) {
  return (
    <div className="w-full">
      <div className="flex items-start gap-0">
        {STEPS.map((s, idx) => {
          const isDone = s.n < current && !(errored && s.n === errorStep);
          const isCur = current === s.n && !errored;
          const isErr = errored && s.n === errorStep;
          return (
            <Fragment key={s.n}>
              <div className="flex flex-col items-center min-w-0 shrink-0" style={{ width: 96 }}>
                <div className="relative h-6 flex items-center justify-center">
                  {isCur && (
                    <span className="absolute inset-0 m-auto w-3.5 h-3.5 rounded-full animate-pulsering" />
                  )}
                  <span
                    className={`relative w-3.5 h-3.5 rounded-full flex items-center justify-center transition-colors ${
                      isErr || isDone || isCur ? 'bg-brand' : 'bg-white border border-[#DADADA]'
                    }`}
                  >
                    {isErr && (
                      <svg width="9" height="9" viewBox="0 0 24 24" stroke="white" strokeWidth="3" strokeLinecap="round" fill="none">
                        <path d="M6 6l12 12M18 6L6 18" />
                      </svg>
                    )}
                  </span>
                </div>
                <div
                  className={`mt-2 text-[12px] font-medium tracking-wide whitespace-nowrap ${
                    isErr ? 'text-brand' : isDone || isCur ? 'text-ink' : 'text-ink3'
                  }`}
                >
                  {s.label}
                </div>
                <div className="mt-1 text-[10px] font-mono text-ink3 h-3">
                  {s.n === 3 && isCur && subProgress ? subProgress : ''}
                </div>
              </div>
              {idx < STEPS.length - 1 && (
                <div className="flex-1 h-px mt-3 relative min-w-[24px]">
                  <div className="absolute inset-0 bg-line" />
                  <div
                    className={`absolute inset-y-0 left-0 transition-all duration-500 ${
                      isErr ? 'bg-line' : s.n < current ? 'bg-brand w-full' : 'w-0'
                    }`}
                  />
                </div>
              )}
            </Fragment>
          );
        })}
      </div>
    </div>
  );
}

const KIND_CLASS: Record<LogMessage['kind'], string> = {
  normal: 'text-term-text',
  rule: 'text-term-dim',
  header: 'text-brand font-semibold',
  milestone: 'text-brand',
  progress: 'text-term-text',
  done: 'text-[#34D399] font-semibold',
  error: 'text-brand font-semibold',
};

function LogLine({ line }: { line: LogMessage }) {
  return (
    <div className="flex gap-3 animate-logline leading-[1.6]">
      <span className="text-term-dim font-mono text-[12px] w-12 shrink-0 text-right tabular-nums">
        +{line.t.toFixed(1)}s
      </span>
      <span className={`font-mono text-[13px] whitespace-pre ${KIND_CLASS[line.kind]}`}>{line.text}</span>
    </div>
  );
}

function ErrorCard({ error, onRetry }: { error: JobError; onRetry: () => void }) {
  return (
    <div className="mt-3 rounded-md border border-brand/40 bg-[#1B0F0E] p-4 animate-fadeup">
      <div className="flex items-start gap-3">
        <div className="text-brand shrink-0 mt-0.5"><IconAlert size={18} /></div>
        <div className="flex-1">
          <div className="text-[13px] font-semibold text-brand">处理失败</div>
          <div className="mt-1 text-[12px] text-[#E5C7C3] leading-relaxed font-mono">
            {error.message}
            {error.hint ? `  ${error.hint}` : ''}
          </div>
          <div className="mt-3">
            <button
              onClick={onRetry}
              className="h-8 px-3 rounded-btn bg-brand text-white text-[12px] font-medium hover:bg-brand-deep inline-flex items-center gap-1.5"
            >
              <IconRefresh size={13} />
              重新上传
            </button>
          </div>
        </div>
      </div>
    </div>
  );
}

export default function Run({ logs, previews, error, onBack }: {
  logs: LogMessage[];
  previews: PreviewImage[];
  error: JobError | null;
  onBack: () => void;
}) {
  const termRef = useRef<HTMLDivElement>(null);
  useEffect(() => {
    if (termRef.current) termRef.current.scrollTop = termRef.current.scrollHeight;
  }, [logs.length, error]);

  // 当前步骤 = 最近一条日志的步骤;done 日志(步骤 5)代表全部完成
  const current = logs.length > 0 ? logs[logs.length - 1].step : 1;
  const finished = logs.some((l) => l.kind === 'done' && l.step === 5);
  // 子进度:步骤 3 的 "已生成 X/Y"
  const subProgress = useMemo(() => {
    for (let i = logs.length - 1; i >= 0; i--) {
      const m = /已生成 (\d+)\/(\d+)/.exec(logs[i].text);
      if (m) return `${m[1]}/${m[2]}`;
      if (logs[i].step !== 3) break;
    }
    return '';
  }, [logs]);

  // ArrayBuffer → blob URL,unmount/更新时回收
  const previewUrls = useMemo(
    () => previews.map((p) => ({
      label: p.label,
      url: URL.createObjectURL(new Blob([p.png], { type: 'image/png' })),
    })),
    [previews],
  );
  useEffect(() => () => previewUrls.forEach((p) => URL.revokeObjectURL(p.url)), [previewUrls]);

  return (
    <div className="max-w-[1080px] mx-auto px-6 py-10 animate-fadeup">
      {/* step indicator */}
      <div className="bg-white border border-line rounded-card px-6 py-5 overflow-x-auto">
        <StepProgress
          current={finished ? 5 : current}
          errored={!!error}
          errorStep={error?.step ?? 1}
          subProgress={subProgress}
        />
      </div>

      {/* terminal */}
      <div className="mt-5 rounded-card border border-line bg-term overflow-hidden">
        <div className="h-9 px-4 flex items-center justify-between border-b border-term-line">
          <div className="flex items-center gap-2">
            <span className="w-2.5 h-2.5 rounded-full bg-[#3a3a3a]" />
            <span className="w-2.5 h-2.5 rounded-full bg-[#3a3a3a]" />
            <span className="w-2.5 h-2.5 rounded-full bg-[#3a3a3a]" />
            <span className="ml-3 text-[11px] font-mono text-term-dim">analyze.log</span>
          </div>
          <div className="text-[11px] font-mono text-term-dim">{logs.length} lines</div>
        </div>
        <div
          ref={termRef}
          className="px-5 py-4 overflow-y-auto"
          style={{ height: 'min(60vh, 540px)' }}
        >
          {logs.map((l, i) => (
            <LogLine key={i} line={l} />
          ))}
          {!error && !finished && (
            <div className="flex gap-3 leading-[1.6]">
              <span className="w-12 shrink-0" />
              <span className="font-mono text-[13px] text-brand animate-caret">▍</span>
            </div>
          )}
          {error && <ErrorCard error={error} onRetry={onBack} />}
        </div>
      </div>

      {/* 趋势图样张(Canvas 渲染目检) */}
      {previewUrls.length > 0 && (
        <div className="mt-5">
          <div className="text-[11px] text-ink3 font-mono mb-2 uppercase tracking-wide">趋势图样张(Canvas 渲染)</div>
          <div className="space-y-3">
            {previewUrls.map((p) => (
              <div key={p.label} className="rounded-card border border-line bg-white p-3 animate-fadeup">
                <div className="text-[11px] text-ink3 font-mono mb-2">{p.label}</div>
                <img src={p.url} alt={p.label} className="max-w-full" />
              </div>
            ))}
          </div>
        </div>
      )}

      {/* footer */}
      <div className="mt-5 flex items-center justify-between">
        <div className="text-[11px] font-mono text-ink3">
          {error ? '已停止' : finished ? '✓ 完成 · 即将跳转' : `进行中 · 步骤 ${current} / 5`}
        </div>
        {!error && !finished && (
          <button onClick={onBack} className="text-[12px] text-ink3 hover:text-ink transition-colors">
            取消
          </button>
        )}
      </div>
    </div>
  );
}
