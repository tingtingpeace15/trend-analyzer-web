// Upload.tsx — 上传页,UI 风格复刻旧版 src/Upload.jsx。
// 两个分析互不干扰:切换 mode 时清空文件队列,避免误把上一个 mode 的文件提交到新 mode。
import { useCallback, useRef, useState } from 'react';
import type { DragEvent, ChangeEvent, ReactNode } from 'react';
import { IconArrowRight, IconClose, IconUpload, IconWarn } from './icons';
import { classifyFile, classifyPreferenceFile, formatBytes } from '../utils/files';
import type { AnalysisMode, FileRole } from '../types/pipeline';

interface FileEntry {
  id: string;
  name: string;
  size: string;
  role: FileRole | null;
  raw: File;
}

function RoleTag({ role }: { role: string }) {
  return (
    <span className="inline-flex items-center px-2 h-6 rounded-md bg-brand-soft text-brand-deep text-[11px] font-semibold tracking-wide whitespace-nowrap">
      {role}
    </span>
  );
}

function FileRow({ file, onRemove }: { file: FileEntry; onRemove: () => void }) {
  return (
    <div className="group flex items-center gap-3 px-3.5 h-14 rounded-btn border border-line bg-white hover:border-ink3 transition-colors animate-fadeup">
      <div className="shrink-0">
        {file.role && <RoleTag role={file.role} />}
      </div>
      <div className="flex-1 min-w-0">
        <div className="font-mono text-[13px] text-ink truncate">{file.name}</div>
      </div>
      <div className="shrink-0 font-mono text-[12px] text-ink3 tabular-nums">{file.size}</div>
      <button
        onClick={onRemove}
        className="shrink-0 w-7 h-7 rounded-md text-ink3 hover:text-ink hover:bg-[#F4F4F4] inline-flex items-center justify-center transition-colors"
        aria-label="移除"
      >
        <IconClose size={14} />
      </button>
    </div>
  );
}

function WarnCard({ name, hint, onIgnore }: { name: string; hint: ReactNode; onIgnore: () => void }) {
  return (
    <div className="border border-warn/40 bg-warn-soft/60 rounded-card p-3.5 flex items-start gap-3 animate-fadeup">
      <div className="text-warn shrink-0 mt-0.5"><IconWarn size={18} /></div>
      <div className="flex-1 text-[13px] text-[#7A4A09] leading-relaxed">
        无法识别 <span className="font-mono font-medium text-[#5A3700]">「{name}」</span>
        {hint}
      </div>
      <button onClick={onIgnore} className="shrink-0 text-[12px] text-[#7A4A09] hover:text-ink underline-offset-2 hover:underline">
        忽略此文件
      </button>
    </div>
  );
}

function DupCard({ role, files, onPick }: { role: string; files: FileEntry[]; onPick: (id: string) => void }) {
  return (
    <div className="border border-warn/40 bg-warn-soft/60 rounded-card p-3.5 animate-fadeup">
      <div className="flex items-start gap-3">
        <div className="text-warn shrink-0 mt-0.5"><IconWarn size={18} /></div>
        <div className="flex-1">
          <div className="text-[13px] text-[#7A4A09] font-medium">
            两个文件都被识别为「{role}」— 请选择保留哪一份
          </div>
          <div className="mt-2.5 space-y-1.5">
            {files.map((f) => (
              <div key={f.id} className="flex items-center gap-3 px-3 h-10 rounded-md bg-white border border-line">
                <div className="font-mono text-[12px] text-ink truncate flex-1">{f.name}</div>
                <div className="font-mono text-[11px] text-ink3 tabular-nums">{f.size}</div>
                <button
                  onClick={() => onPick(f.id)}
                  className="text-[12px] px-2.5 h-7 rounded-md bg-ink text-white hover:bg-black"
                >
                  保留
                </button>
              </div>
            ))}
          </div>
        </div>
      </div>
    </div>
  );
}

function Dropzone({ hasAny, multiple, idleTitle, idleHint, onFiles }: {
  hasAny: boolean;
  multiple: boolean;
  idleTitle: string;
  idleHint: string;
  onFiles: (files: File[]) => void;
}) {
  const [drag, setDrag] = useState(false);
  const inputRef = useRef<HTMLInputElement>(null);

  const handleDrop = (e: DragEvent) => {
    e.preventDefault();
    setDrag(false);
    onFiles(Array.from(e.dataTransfer.files || []));
  };
  const handlePick = (e: ChangeEvent<HTMLInputElement>) => {
    onFiles(Array.from(e.target.files || []));
    e.target.value = '';
  };

  return (
    <div
      onClick={() => inputRef.current?.click()}
      onDragOver={(e) => { e.preventDefault(); setDrag(true); }}
      onDragLeave={() => setDrag(false)}
      onDrop={handleDrop}
      className={`relative cursor-pointer rounded-card border-2 border-dashed transition-all duration-200 ${
        drag
          ? 'border-brand bg-brand-soft/40'
          : 'border-[#DCDCDC] bg-[#FBFBFB] hover:border-ink3 hover:bg-white'
      } ${hasAny ? 'py-10 px-6' : 'py-16 px-6'}`}
    >
      <input
        ref={inputRef}
        type="file"
        multiple={multiple}
        accept=".xlsx,.xls"
        className="hidden"
        onChange={handlePick}
      />
      <div className="flex flex-col items-center text-center">
        <div className={`w-12 h-12 rounded-xl border border-line bg-white flex items-center justify-center mb-4 transition-colors ${drag ? 'text-brand border-brand' : 'text-ink2'}`}>
          <IconUpload size={20} />
        </div>
        <div className="text-[15px] font-medium text-ink">
          {drag ? '松开放下文件' : (hasAny ? '再拖一份进来' : idleTitle)}
        </div>
        <div className="mt-1 text-[12px] text-ink3">{idleHint}</div>
      </div>
    </div>
  );
}

const TABS: { id: AnalysisMode; label: string; desc: string }[] = [
  { id: 'trend', label: '商品趋势分析', desc: '滞销商品 + 各商品客户拿货历史 → 含趋势图的报表' },
  { id: 'preference', label: '客户偏好分析', desc: '各商品客户拿货历史 → 客户分层 + 偏好画像' },
];

function AnalysisTabs({ value, onChange }: { value: AnalysisMode; onChange: (m: AnalysisMode) => void }) {
  return (
    <div className="grid grid-cols-2 gap-3">
      {TABS.map((t) => {
        const on = value === t.id;
        return (
          <button
            key={t.id}
            onClick={() => onChange(t.id)}
            className={`text-left rounded-card border px-4 py-3 transition-all ${
              on
                ? 'border-brand bg-brand-soft/40 ring-2 ring-brand/20'
                : 'border-line bg-white hover:border-ink3'
            }`}
          >
            <div className={`text-[14px] font-semibold ${on ? 'text-brand-deep' : 'text-ink'}`}>
              {t.label}
            </div>
            <div className="mt-1 text-[12px] text-ink2 leading-relaxed">{t.desc}</div>
          </button>
        );
      })}
    </div>
  );
}

const HELPER_CARDS: Record<AnalysisMode, { n: string; title: string; desc: string }[]> = {
  trend: [
    { n: '01', title: '上传两份 Excel', desc: '滞销商品 + 各商品客户拿货历史,按文件名识别' },
    { n: '02', title: '浏览器本地分析', desc: '数据不离开你的电脑,无需安装环境' },
    { n: '03', title: '下载报表', desc: '3 个 sheet · 含嵌入式趋势图' },
  ],
  preference: [
    { n: '01', title: '上传一份 Excel', desc: '「各商品客户拿货历史」按文件名识别' },
    { n: '02', title: '聚合客户偏好', desc: '品类 / 颜色 / 尺码 / 价格 / 品牌' },
    { n: '03', title: '下载 Excel + 网页', desc: '约 20 个 sheet · 9 Tab 交互式报告' },
  ],
};

export default function Upload({ onStart }: {
  onStart: (mode: AnalysisMode, files: { role: FileRole; file: File }[]) => void;
}) {
  const [mode, setMode] = useState<AnalysisMode>('trend');
  const [files, setFiles] = useState<FileEntry[]>([]);
  const [unrec, setUnrec] = useState<FileEntry[]>([]);

  const isPref = mode === 'preference';

  const switchMode = (next: AnalysisMode) => {
    if (next === mode) return;
    setMode(next);
    setFiles([]);
    setUnrec([]);
  };

  const addFiles = useCallback((rawList: File[]) => {
    const classifier = isPref ? classifyPreferenceFile : classifyFile;
    let next = [...files];
    const nextUnrec = [...unrec];
    for (const f of rawList) {
      const role = classifier(f.name);
      const entry: FileEntry = {
        id: `${f.name}_${f.size}_${Math.random().toString(36).slice(2, 6)}`,
        name: f.name,
        size: formatBytes(f.size),
        role,
        raw: f,
      };
      if (role) next.push(entry);
      else nextUnrec.push(entry);
    }
    // 偏好分析最多保留 1 个识别成功的文件——后来的覆盖前面的
    if (isPref && next.length > 1) next = next.slice(-1);
    setFiles(next);
    setUnrec(nextUnrec);
  }, [files, unrec, isPref]);

  const removeFile = (id: string) => setFiles(files.filter((f) => f.id !== id));
  const ignoreUnrec = (id: string) => setUnrec(unrec.filter((f) => f.id !== id));

  // 趋势分析:按 role 检测重复
  const byRole = files.reduce<Record<string, FileEntry[]>>((acc, f) => {
    const key = f.role ?? '';
    (acc[key] = acc[key] || []).push(f);
    return acc;
  }, {});
  const duplicates = isPref ? [] : Object.entries(byRole).filter(([, arr]) => arr.length > 1);
  const pickDup = (role: string, keepId: string) => {
    setFiles(files.filter((f) => f.role !== role || f.id === keepId));
  };

  // 两种模式各自的就绪判断
  let ready = false;
  let missing: string | null = null;
  if (isPref) {
    ready = files.length === 1;
    missing = ready ? null : '拿货历史明细';
  } else {
    const haveSlow = byRole['滞销表']?.length === 1;
    const haveSales = byRole['销售明细']?.length === 1;
    ready = !!haveSlow && !!haveSales && duplicates.length === 0;
    missing = !haveSlow ? '滞销商品' : (!haveSales ? '各商品客户拿货历史' : null);
  }

  const heroTitle = isPref ? '上传一份 Excel' : '上传两份 Excel';
  const heroSub = isPref
    ? '文件名需包含 拿货历史 / 客户偏好 / 拿货 · 数据全程在浏览器本地处理'
    : '程序按文件名自动识别 · 数据全程在浏览器本地处理';
  const dropIdleTitle = isPref ? '拖入「各商品客户拿货历史」Excel' : '拖入两个 Excel 文件';
  const dropIdleHint = isPref
    ? '或点击选择 · 文件名带 拿货历史 / 客户偏好 即可'
    : '或点击选择 · 自动识别滞销商品 + 各商品客户拿货历史';
  const unrecHint: ReactNode = isPref ? (
    <> — 文件名需包含 <span className="font-mono">拿货历史</span> 或 <span className="font-mono">客户偏好</span></>
  ) : (
    <> — 文件名需包含 <span className="font-mono">滞销</span> 或 <span className="font-mono">各商品客户拿货历史</span>(也兼容"销售明细 / 销售单明细 / 销售流水")</>
  );

  const submit = () => {
    if (!ready) return;
    onStart(mode, files.map((f) => ({ role: f.role!, file: f.raw })));
  };

  return (
    <div className="max-w-[760px] mx-auto px-6 py-10 sm:py-12 animate-fadeup">
      <div className="mb-3">
        <div className="text-[11px] text-ink3 font-mono mb-2 uppercase tracking-wide">分析类型</div>
        <AnalysisTabs value={mode} onChange={switchMode} />
      </div>

      <h1 className="text-[26px] sm:text-[30px] font-semibold tracking-tight text-ink leading-tight">
        {heroTitle}
      </h1>
      <p className="mt-2 text-[13px] text-ink2">{heroSub}</p>

      <div className="mt-7">
        <Dropzone
          hasAny={files.length + unrec.length > 0}
          multiple={!isPref}
          idleTitle={dropIdleTitle}
          idleHint={dropIdleHint}
          onFiles={addFiles}
        />
      </div>

      {(files.length > 0 || unrec.length > 0) && (
        <div className="mt-4 space-y-2.5">
          {files.map((f) => (
            <FileRow key={f.id} file={f} onRemove={() => removeFile(f.id)} />
          ))}
          {unrec.map((f) => (
            <WarnCard key={f.id} name={f.name} hint={unrecHint} onIgnore={() => ignoreUnrec(f.id)} />
          ))}
          {duplicates.map(([role, arr]) => (
            <DupCard key={role} role={role} files={arr} onPick={(keepId) => pickDup(role, keepId)} />
          ))}
        </div>
      )}

      <div className="mt-9 flex items-end justify-between gap-4">
        <div className="text-[12px] text-ink3 font-mono leading-relaxed">
          {isPref ? '预计 10~30 秒' : '预计 30~90 秒(2000 款约 1 分钟)'}
        </div>
        <div className="relative group">
          <button
            onClick={submit}
            disabled={!ready}
            className={`h-11 px-6 rounded-btn text-[14px] font-medium inline-flex items-center gap-2 transition-all ${
              ready
                ? 'bg-brand text-white hover:bg-brand-deep hover:shadow-btn active:translate-y-[0.5px]'
                : 'bg-[#F1F1F1] text-ink3 cursor-not-allowed'
            }`}
          >
            开始分析
            <IconArrowRight size={14} />
          </button>
          {!ready && missing && (
            <div className="absolute right-0 top-full mt-2 hidden group-hover:block">
              <div className="bg-ink text-white text-[11px] px-2.5 py-1.5 rounded-md whitespace-nowrap">
                还需要一份「{missing}」
              </div>
            </div>
          )}
        </div>
      </div>

      <div className="mt-12 grid grid-cols-1 sm:grid-cols-3 gap-3 text-[12px] text-ink2">
        {HELPER_CARDS[mode].map((s) => (
          <div key={s.n} className="rounded-card border border-line bg-white p-4">
            <div className="font-mono text-[11px] text-ink3 mb-1">{s.n}</div>
            <div className="text-[13px] font-medium text-ink">{s.title}</div>
            <div className="mt-1 text-[12px] text-ink2 leading-relaxed">{s.desc}</div>
          </div>
        ))}
      </div>
    </div>
  );
}
