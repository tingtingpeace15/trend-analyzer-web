// App.tsx — 顶层路由:upload → run → result。布局壳复刻旧版 App.jsx。
import { useEffect, useRef, useState } from 'react';
import Upload from './components/Upload';
import Run from './components/Run';
import Result from './components/Result';
import PreferenceResult from './components/PreferenceResult';
import { IconLogo } from './components/icons';
import { prewarmWorker, startJob } from './api';
import type { JobDone, JobError, JobHandle, JobInputFile, PreviewImage } from './api';
import type { AnalysisMode, LogMessage } from './types/pipeline';

type Stage = 'upload' | 'run' | 'result';

export default function App() {
  const [stage, setStage] = useState<Stage>('upload');
  const [mode, setMode] = useState<AnalysisMode>('trend');
  const [logs, setLogs] = useState<LogMessage[]>([]);
  const [previews, setPreviews] = useState<PreviewImage[]>([]);
  const [result, setResult] = useState<JobDone | null>(null);
  const [error, setError] = useState<JobError | null>(null);
  const jobRef = useRef<JobHandle | null>(null);

  // 页面加载后空闲预热 Worker(1.3MB bundle 提前下载+编译,消掉点开始后的长空白)
  useEffect(() => {
    const t = setTimeout(prewarmWorker, 300);
    return () => clearTimeout(t);
  }, []);

  const handleStart = (jobMode: AnalysisMode, inputs: JobInputFile[]) => {
    setMode(jobMode);
    // 即时反馈:Worker 首条日志到来前,终端不留白
    setLogs([{ type: 'log', text: '正在初始化分析引擎…', kind: 'normal', step: 1, t: 0 }]);
    setPreviews([]);
    setResult(null);
    setError(null);
    setStage('run');
    jobRef.current = startJob(jobMode, inputs, {
      onLog: (log) => setLogs((prev) => [...prev, log]),
      onPreview: (images) => setPreviews(images),
      onDone: (done) => {
        setResult(done);
        // 同旧版:让「完成」状态在步骤条上停留一瞬再跳转
        setTimeout(() => setStage('result'), 600);
      },
      onError: (err) => setError(err),
    });
  };

  const handleBack = () => {
    jobRef.current?.cancel();
    jobRef.current = null;
    setStage('upload');
  };

  return (
    <div className="min-h-full w-full flex flex-col">
      <header className="h-14 border-b border-line bg-white/80 backdrop-blur-sm sticky top-0 z-10">
        <div className="max-w-[960px] mx-auto h-full px-6 flex items-center justify-between">
          <div className="flex items-center gap-2">
            <IconLogo size={18} />
            <span className="text-[13px] text-ink font-medium">分析工具</span>
            <span className="text-[11px] text-ink3 font-mono ml-2 hidden sm:inline">web v0.2</span>
          </div>
        </div>
      </header>

      <main className="flex-1">
        {stage === 'upload' && <Upload onStart={handleStart} />}
        {stage === 'run' && (
          <Run mode={mode} logs={logs} previews={previews} error={error} onBack={handleBack} />
        )}
        {stage === 'result' && result?.mode === 'trend' && (
          <Result
            filename={result.filename}
            buffer={result.buffer}
            summary={result.summary}
            onAgain={handleBack}
          />
        )}
        {stage === 'result' && result?.mode === 'preference' && (
          <PreferenceResult
            html={result.html}
            newHtml={result.newHtml}
            xlsx={result.xlsx}
            summary={result.summary}
            onAgain={handleBack}
          />
        )}
      </main>
    </div>
  );
}
