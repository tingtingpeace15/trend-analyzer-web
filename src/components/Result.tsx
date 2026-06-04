// Result.tsx — 下载结果页,复刻旧版 Result.jsx 的趋势分析视图(简化:无 jobId/后端)。
import { useState } from 'react';
import {
  IconArrowRight, IconChart, IconCheck, IconDownload, IconGrid, IconLayers,
  IconRefresh, IconSpark,
} from './icons';
import { formatBytes } from '../utils/files';
import type { JobSummary } from '../types/pipeline';

const SHEETS = [
  { key: 's1', icon: 'chart', name: '商品销售趋势', desc: '每款一行,小趋势图 + 销售汇总,可筛可排' },
  { key: 's2', icon: 'layers', name: '款趋势明细图', desc: '大图版,逐日走势看得更清楚' },
  { key: 's3', icon: 'grid', name: '款日销量明细', desc: '货号 × 日期的销量矩阵,含合计列' },
] as const;

function SheetCard({ icon, name, desc, idx }: { icon: string; name: string; desc: string; idx: number }) {
  const Icon = icon === 'chart' ? IconChart : icon === 'layers' ? IconLayers : IconGrid;
  return (
    <div className="flex items-start gap-3 py-4 border-b border-line last:border-b-0">
      <div className="w-9 h-9 rounded-md bg-[#FAFAFA] border border-line flex items-center justify-center shrink-0 text-ink2">
        <Icon size={16} />
      </div>
      <div className="flex-1 min-w-0">
        <div className="flex items-center gap-2">
          <div className="font-mono text-[11px] text-ink3">sheet {idx}</div>
          <div className="font-semibold text-[14px] text-ink truncate">{name}</div>
        </div>
        <div className="mt-1 text-[12.5px] text-ink2 leading-relaxed">{desc}</div>
      </div>
    </div>
  );
}

export default function Result({ filename, buffer, summary, onAgain }: {
  filename: string;
  buffer: ArrayBuffer;
  summary: JobSummary;
  onAgain: () => void;
}) {
  const [downloaded, setDownloaded] = useState(false);

  const triggerDownload = () => {
    setDownloaded(true);
    const blob = new Blob([buffer], {
      type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
    });
    const url = URL.createObjectURL(blob);
    const a = document.createElement('a');
    a.href = url;
    a.download = filename;
    document.body.appendChild(a);
    a.click();
    a.remove();
    setTimeout(() => URL.revokeObjectURL(url), 1000);
  };

  return (
    <div className="max-w-[960px] mx-auto px-6 py-12 sm:py-16">
      <div className="text-center animate-fadeup">
        <div className="inline-flex w-12 h-12 rounded-full bg-[#ECFDF5] border border-[#A7F3D0] items-center justify-center text-ok">
          <IconCheck size={22} />
        </div>
        <h1 className="mt-5 text-[30px] sm:text-[34px] font-semibold tracking-tight text-ink">
          分析完成
        </h1>
        <div className="mt-2 text-[13px] text-ink2 font-mono">
          <span>分析了 <span className="text-ink font-medium">{summary.items.toLocaleString()}</span> 款</span>
          <span className="mx-2 text-ink3">·</span>
          <span>时间窗 <span className="text-ink font-medium">{summary.windowDays}</span> 天</span>
          <span className="mx-2 text-ink3">·</span>
          <span><span className="text-ink font-medium">{summary.itemsWithSales.toLocaleString()}</span> 款有销量</span>
        </div>
      </div>

      <div className="mt-12 grid grid-cols-1 lg:grid-cols-5 gap-5 animate-fadeup">
        <div className="lg:col-span-3 space-y-3">
          <button
            onClick={triggerDownload}
            className="group w-full rounded-card bg-brand hover:bg-brand-deep transition-all px-7 py-7 text-left text-white hover:shadow-btn active:translate-y-[0.5px]"
          >
            <div className="flex items-center gap-4">
              <div className="w-10 h-10 rounded-md bg-white/15 flex items-center justify-center shrink-0">
                <IconDownload size={20} />
              </div>
              <div className="flex-1 min-w-0">
                <div className="text-[16px] font-semibold tracking-tight truncate">
                  下载 {filename}
                </div>
                <div className="mt-0.5 text-[12px] text-white/80 font-mono">
                  {formatBytes(summary.sizeBytes)} · .xlsx · 3 sheet · 含嵌入式趋势图
                </div>
              </div>
              <IconArrowRight size={18} />
            </div>
          </button>
          {downloaded && (
            <div className="text-[12px] text-ok font-mono pl-1 animate-fadeup">
              ✓ 已开始下载 — 也可重复点击
            </div>
          )}

          <div className="flex items-center gap-3">
            <button
              onClick={onAgain}
              className="h-10 px-4 rounded-btn border border-brand text-brand text-[13px] font-medium hover:bg-brand-soft transition-colors inline-flex items-center gap-1.5"
            >
              <IconRefresh size={14} />
              再分析一份
            </button>
            <div className="text-[11px] font-mono text-ink3">
              {summary.windowFrom} → {summary.windowTo} · 浏览器本地生成
            </div>
          </div>
        </div>

        <aside className="lg:col-span-2">
          <div className="rounded-card border border-line bg-white p-5">
            <div className="flex items-center gap-2 mb-1">
              <IconSpark size={14} />
              <div className="text-[14px] font-semibold text-ink">这份 Excel 里有什么</div>
            </div>
            <div className="text-[12px] text-ink3 mb-3">3 个 sheet · 全部嵌图</div>
            <div>
              {SHEETS.map((s, i) => (
                <SheetCard key={s.key} idx={i + 1} icon={s.icon} name={s.name} desc={s.desc} />
              ))}
            </div>
          </div>
          <div className="mt-3 text-[11px] text-ink3 font-mono leading-relaxed">
            数字以红色 <span className="text-brand">#E74C3C</span> 标记 ·
            嵌图为每款独立 PNG · 可直接发邮件 / 打印
          </div>
        </aside>
      </div>
    </div>
  );
}
