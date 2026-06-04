// App.tsx — 顶层路由:upload → run(→ result,M5 起)。布局壳复刻旧版 App.jsx。
import { useRef, useState } from 'react';
import Upload from './components/Upload';
import Run from './components/Run';
import Result from './components/Result';
import { IconLogo } from './components/icons';
import { startJob } from './api';
import type { JobError, JobHandle, JobInputFile, PreviewImage } from './api';
import type { JobSummary, LogMessage } from './types/pipeline';

type Stage = 'upload' | 'run' | 'result';

interface JobResult {
  filename: string;
  buffer: ArrayBuffer;
  summary: JobSummary;
}

export default function App() {
  const [stage, setStage] = useState<Stage>('upload');
  const [logs, setLogs] = useState<LogMessage[]>([]);
  const [previews, setPreviews] = useState<PreviewImage[]>([]);
  const [result, setResult] = useState<JobResult | null>(null);
  const [error, setError] = useState<JobError | null>(null);
  const jobRef = useRef<JobHandle | null>(null);

  const handleStart = (inputs: JobInputFile[]) => {
    setLogs([]);
    setPreviews([]);
    setResult(null);
    setError(null);
    setStage('run');
    jobRef.current = startJob(inputs, {
      onLog: (log) => setLogs((prev) => [...prev, log]),
      onPreview: (images) => setPreviews(images),
      onDone: (filename, buffer, summary) => {
        setResult({ filename, buffer, summary });
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
            <span className="text-[13px] text-ink font-medium">商品趋势分析</span>
            <span className="text-[11px] text-ink3 font-mono ml-2 hidden sm:inline">web v0.1</span>
          </div>
        </div>
      </header>

      <main className="flex-1">
        {stage === 'upload' && <Upload onStart={handleStart} />}
        {stage === 'run' && <Run logs={logs} previews={previews} error={error} onBack={handleBack} />}
        {stage === 'result' && result && (
          <Result
            filename={result.filename}
            buffer={result.buffer}
            summary={result.summary}
            onAgain={handleBack}
          />
        )}
      </main>
    </div>
  );
}
