// PreferenceResult.tsx — 偏好分析下载页,复刻旧版 Result.jsx 的偏好视图。
// 新增「新标签页预览」:html 报告用 blob URL 直接打开,不用先下载。
import { useEffect, useMemo, useState } from 'react';
import {
  IconArrowRight, IconChart, IconCheck, IconDownload, IconGrid, IconLayers,
  IconRefresh, IconSpark,
} from './icons';
import { formatBytes } from '../utils/files';
import type { FilePayload, PreferenceSummary } from '../types/pipeline';

const SHEETS = [
  { key: 's1', icon: 'chart', name: '客户分层 + Top50', desc: 'VIP / 高价值 / 中等 / 低价值,附核心客户排名' },
  { key: 's2', icon: 'layers', name: '品类 / 颜色 / 尺码偏好', desc: '每个客户主推哪些品类、哪种配色、什么尺码' },
  { key: 's3', icon: 'grid', name: '价格敏感度 + 客户画像', desc: '追新型 / 折扣型 / 均衡型,全部客户的完整偏好档案' },
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
          <div className="font-mono text-[11px] text-ink3">{idx}</div>
          <div className="font-semibold text-[14px] text-ink truncate">{name}</div>
        </div>
        <div className="mt-1 text-[12.5px] text-ink2 leading-relaxed">{desc}</div>
      </div>
    </div>
  );
}

function fmtAmount(n: number): string {
  if (!n) return '0';
  if (n >= 10000) return `¥${(n / 10000).toFixed(1)}万`;
  return `¥${n.toLocaleString()}`;
}

function download(buffer: ArrayBuffer, filename: string, mime: string) {
  const url = URL.createObjectURL(new Blob([buffer], { type: mime }));
  const a = document.createElement('a');
  a.href = url;
  a.download = filename;
  document.body.appendChild(a);
  a.click();
  a.remove();
  setTimeout(() => URL.revokeObjectURL(url), 1000);
}

export default function PreferenceResult({ html, newHtml, xlsx, summary, onAgain }: {
  html: FilePayload;
  newHtml: FilePayload;
  xlsx: FilePayload;
  summary: PreferenceSummary;
  onAgain: () => void;
}) {
  const [downloaded, setDownloaded] = useState(false);

  // html 报告的预览 URL(组件存活期间有效)
  const htmlUrl = useMemo(
    () => URL.createObjectURL(new Blob([html.buffer], { type: 'text/html' })),
    [html],
  );
  useEffect(() => () => URL.revokeObjectURL(htmlUrl), [htmlUrl]);
  const newHtmlUrl = useMemo(
    () => URL.createObjectURL(new Blob([newHtml.buffer], { type: 'text/html' })),
    [newHtml],
  );
  useEffect(() => () => URL.revokeObjectURL(newHtmlUrl), [newHtmlUrl]);

  return (
    <div className="max-w-[960px] mx-auto px-6 py-12 sm:py-16">
      <div className="text-center animate-fadeup">
        <div className="inline-flex w-12 h-12 rounded-full bg-[#ECFDF5] border border-[#A7F3D0] items-center justify-center text-ok">
          <IconCheck size={22} />
        </div>
        <h1 className="mt-5 text-[30px] sm:text-[34px] font-semibold tracking-tight text-ink">
          偏好分析完成
        </h1>
        <div className="mt-2 text-[13px] text-ink2 font-mono">
          <span><span className="text-ink font-medium">{summary.records.toLocaleString()}</span> 条交易</span>
          <span className="mx-2 text-ink3">·</span>
          <span><span className="text-ink font-medium">{summary.customers.toLocaleString()}</span> 位客户</span>
          <span className="mx-2 text-ink3">·</span>
          <span>销售额 <span className="text-ink font-medium">{fmtAmount(summary.amount)}</span></span>
          {summary.dateFrom && summary.dateTo && (
            <>
              <span className="mx-2 text-ink3">·</span>
              <span>{summary.dateFrom} → {summary.dateTo}</span>
            </>
          )}
        </div>
      </div>

      <div className="mt-12 grid grid-cols-1 lg:grid-cols-5 gap-5 animate-fadeup">
        <div className="lg:col-span-3 space-y-3">
          {/* 主推 — Excel */}
          <button
            onClick={() => {
              setDownloaded(true);
              download(xlsx.buffer, xlsx.filename,
                'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
            }}
            className="group w-full rounded-card bg-brand hover:bg-brand-deep transition-all px-7 py-7 text-left text-white hover:shadow-btn active:translate-y-[0.5px]"
          >
            <div className="flex items-center gap-4">
              <div className="w-10 h-10 rounded-md bg-white/15 flex items-center justify-center shrink-0">
                <IconDownload size={20} />
              </div>
              <div className="flex-1 min-w-0">
                <div className="text-[16px] font-semibold tracking-tight truncate">
                  下载 {xlsx.filename}
                </div>
                <div className="mt-0.5 text-[12px] text-white/80 font-mono">
                  {formatBytes(xlsx.buffer.byteLength)} · .xlsx · 完整客户偏好数据
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

          {/* 副 — 网页报告:预览 + 下载 */}
          <div className="w-full rounded-card border border-line bg-white px-5 py-4">
            <div className="flex items-center gap-3">
              <div className="w-9 h-9 rounded-md bg-[#FAFAFA] border border-line flex items-center justify-center shrink-0 text-ink2">
                <IconChart size={16} />
              </div>
              <div className="flex-1 min-w-0">
                <div className="text-[14px] font-semibold text-ink">网页报告</div>
                <div className="mt-0.5 text-[12px] text-ink3 font-mono truncate">
                  {html.filename} · {formatBytes(html.buffer.byteLength)} · 9 个 Tab 交互式图表
                </div>
              </div>
              <a
                href={htmlUrl}
                target="_blank"
                rel="noreferrer"
                className="shrink-0 h-9 px-3.5 rounded-btn border border-brand text-brand text-[13px] font-medium hover:bg-brand-soft transition-colors inline-flex items-center"
              >
                新标签页预览
              </a>
              <button
                onClick={() => download(html.buffer, html.filename, 'text/html')}
                className="shrink-0 h-9 px-3.5 rounded-btn bg-ink text-white text-[13px] font-medium hover:bg-black transition-colors"
              >
                下载
              </button>
            </div>
          </div>

          {/* 新增 — 新客户偏好分析:预览 + 下载 */}
          <div className="w-full rounded-card border border-line bg-white px-5 py-4">
            <div className="flex items-center gap-3">
              <div className="w-9 h-9 rounded-md bg-[#FAFAFA] border border-line flex items-center justify-center shrink-0 text-ink2">
                <IconChart size={16} />
              </div>
              <div className="flex-1 min-w-0">
                <div className="text-[14px] font-semibold text-ink">新客户偏好分析</div>
                <div className="mt-0.5 text-[12px] text-ink3 font-mono truncate">
                  {newHtml.filename} · {formatBytes(newHtml.buffer.byteLength)} · 新版客户偏好看板
                </div>
              </div>
              <a
                href={newHtmlUrl}
                target="_blank"
                rel="noreferrer"
                className="shrink-0 h-9 px-3.5 rounded-btn border border-brand text-brand text-[13px] font-medium hover:bg-brand-soft transition-colors inline-flex items-center"
              >
                新标签页预览
              </a>
              <button
                onClick={() => download(newHtml.buffer, newHtml.filename, 'text/html')}
                className="shrink-0 h-9 px-3.5 rounded-btn bg-ink text-white text-[13px] font-medium hover:bg-black transition-colors"
              >
                下载
              </button>
            </div>
          </div>

          <div className="flex items-center gap-3 pt-1">
            <button
              onClick={onAgain}
              className="h-10 px-4 rounded-btn border border-brand text-brand text-[13px] font-medium hover:bg-brand-soft transition-colors inline-flex items-center gap-1.5"
            >
              <IconRefresh size={14} />
              再分析一份
            </button>
            <div className="text-[11px] font-mono text-ink3">浏览器本地生成 · 数据不上传</div>
          </div>
        </div>

        <aside className="lg:col-span-2">
          <div className="rounded-card border border-line bg-white p-5">
            <div className="flex items-center gap-2 mb-1">
              <IconSpark size={14} />
              <div className="text-[14px] font-semibold text-ink">这份分析里有什么</div>
            </div>
            <div className="text-[12px] text-ink3 mb-3">Excel · 网页 · 双格式</div>
            <div>
              {SHEETS.map((s, i) => (
                <SheetCard key={s.key} idx={i + 1} icon={s.icon} name={s.name} desc={s.desc} />
              ))}
            </div>
          </div>
          <div className="mt-3 text-[11px] text-ink3 font-mono leading-relaxed">
            Excel sheet 数随输入字段而定(缺列自动跳过)·
            网页报告含 9 个 Tab,图表需联网加载 Chart.js
          </div>
        </aside>
      </div>
    </div>
  );
}
