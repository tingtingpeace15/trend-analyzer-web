# trend-analyzer-web — 项目结构 + 维护笔记

商品趋势分析工具的**纯前端网页版**。0 后端,Excel 在浏览器本地解析/分析/生成,最终部署 GitHub Pages。

完整需求见基准项目里的 `WEB_VERSION_BRIEF.md`(`../趋势分析工具 -最终app版本/`,**只读参考,不要改**)。

## 进度

- [x] **M1** 项目骨架 + 上传 UI + Web Worker 通信(2026-06-04)
- [x] **M2** SheetJS 读 Excel + 字段嗅探(2026-06-04;9 项 Vitest 对照 pandas 全过)
- [x] **M3** 聚合算法(2026-06-04;透视/透传/base 拼装,7 项对照全过,合计 16 项)
- [x] **M4** Canvas 画图(2026-06-04;真实尺寸 407×115 / 713×206,brief 里 605/1109 是过时数字;几何 8 项对照全过,合计 24 项)
- [x] **M5** ExcelJS 写 3 个 sheet + 嵌图(2026-06-04;排序/单元格/样式/锚点对照 9 项全过,合计 33 项;端到端可下载)
- [x] **M6** 进度日志 + 性能优化 + 错误处理(2026-06-04;Run 页复刻旧版,PNG 8 路并行,结构化错误)
- [x] **M7** 全字段 diff Python 版输出(2026-06-04;**逐 sheet/单元格/样式/锚点全部一致**)
- [x] **M8** GitHub Pages 部署(2026-06-04;https://tingtingpeace15.github.io/trend-analyzer-web/ ,推 main 自动部署)

### 第二条业务线:客户偏好分析(对照 backend/preference_pipeline.py)

- [x] **M9** UI 双 tab + 偏好 Worker 通道 + 偏好 Result 页(2026-06-05;stub 走通)
- [x] **M10** 读取 + 列规范化(2026-06-05;7 项对照 pandas 全过,合计 41 项)
- [x] **M11** _analyze 复刻(2026-06-05;R 全部 19 个 key 逐位一致,含 numpy introsort 移植;合计 61 项)
- [x] **M12** HTML 报告 + Excel 23 sheet(2026-06-05;html 逐字节一致,xlsx 全字段一致仅平局重排;合计 62 项测试)
- [x] **M13** 端到端全字段 diff + 部署上线(2026-06-05)

## 结构

```
src/
├── main.tsx / App.tsx          # 顶层路由:upload → run → result(M5)
├── components/
│   ├── Upload.tsx              # 上传页,UI 复刻旧版 src/Upload.jsx
│   ├── Run.tsx                 # 进度日志终端(M1 简化版,M6 完整化)
│   └── icons.tsx               # SVG 图标,移植自旧版 icons.jsx
├── types/pipeline.ts           # 主线程 ↔ Worker 消息协议
├── types/excel.ts              # 读取阶段数据结构(ZhixiaoTable / SalesTable)
├── pipeline/
│   ├── errors.ts               # PipelineError(msg + hint,对应 Python 版)
│   ├── reader.ts               # SheetJS 读 Excel + 表头嗅探(照搬 pipeline.py:116-172)
│   ├── aggregator.ts           # 透视/聚合/字段透传(照搬 pipeline.py:174-282)
│   ├── chart.ts                # Canvas 复刻 matplotlib 趋势图(pipeline.py:289-389);
│   │                           #   几何(纯函数,Node 可测)与渲染(OffscreenCanvas)分离;
│   │                           #   几何常量由 matplotlib 真实输出反推校准,勿凭感觉改
│   └── writer.ts               # ExcelJS 写 3 个 sheet(pipeline.py:440-755);
│                               #   含 pandas str() 语义(float64 列整数带 ".0")、
│                               #   TwoCellAnchor 嵌图、批注、稳定多列排序
├── utils/files.ts              # classifyFile / formatBytes(规则同旧版 mockApi.jsx)
├── api.ts                      # startJob():读文件 → transferable 传 Worker → 回调
├── workers/pipeline.worker.ts  # 分析 pipeline(M2 已接 reader;M3+ 接 aggregator/chart/writer)
└── __tests__/
    ├── golden/reader.golden.json  # pandas 跑真实数据的基准(scripts/gen_golden_reader.py 生成)
    └── reader.test.ts             # JS reader vs pandas 逐项对照
```

## 黄金基准测试法(M2 起的工作流)

每个阶段先用 `scripts/gen_golden_*.py` 拿 pandas/Python 版跑真实客户数据,把关键
中间结果写成 JSON 基准,再写 Vitest 逐项对照 JS 实现。**测试不过修 JS,不改基准**。
Python 脚本只读基准目录(`../趋势分析工具 -最终app版本/`),绝不写入。

## M7 全字段 diff(回归验证流程)

改了 reader/aggregator/writer 之后,跑这三步确认跟 Python 版仍逐字段一致:

```bash
python3 scripts/gen_baseline_python.py   # Python 版真跑 → baseline/python/(~50s)
npx vite-node scripts/gen_js_output.ts   # JS 版(Node,嵌图用占位)→ baseline/js/
python3 scripts/diff_xlsx.py             # 逐 sheet/单元格/样式/锚点 diff,0 = 一致
```

约定不比的两项:图片字节(Canvas vs matplotlib 光栅化必然不同,由 chart.test.ts
的几何对照 + 浏览器目检保障)、批注外形尺寸(ExcelJS 不支持设置)。

## 偏好分析的回归验证流程(M13)

```bash
python3 scripts/gen_baseline_preference.py        # Python 真跑 → baseline/python-pref/(~55s)
npx vite-node scripts/gen_js_preference.ts        # JS 版 → baseline/js-pref/(~8s)
npx vitest run src/__tests__/preferenceHtml.test.ts   # html 逐字节对比(需先有基准)
python3 scripts/diff_xlsx.py baseline/python-pref/客户偏好分析数据.xlsx baseline/js-pref/客户偏好分析数据.xlsx
```

**平局重排**:排序值相等的行,行序允许不同(diff 自动归类为可接受)。原因:numpy 2.x
的 SIMD argsort 平局顺序跨平台/版本不稳定,Python 版自己换台机器跑都会变。

## pandas 数值语义速查(改聚合代码前必读)

| Python 写法 | 语义 | JS 对应(preferenceAnalyze.ts) |
|---|---|---|
| `groupby().sum()/mean()` | Kahan 补偿求和 | `gsum` / `gmean` |
| `Series.sum()/mean()`(无 bottleneck) | numpy pairwise(8 累加器+128 分块) | `psum` / `pmean` / `npPairwiseSum` |
| 内建 `round(x, d)` | 对 double 精确十进制 half-even | `pyRound`(BigInt 精确) |
| `.round(d)`(pandas/np) | 乘 10^d 后 half-even(带乘法误差) | `npRound` |
| `sort_values(ascending=False)` | 反转输入→升序 argsort→再反转 | `sortValuesDesc`(内含 numpy introsort 移植) |
| 多列 sort_values | 稳定 lexsort | `sortByDescStable` |
| `int(x)` | 向零截断 | `pyInt` |

## 维护笔记

- Tailwind **v3**(不是 v4),色卡/动画整份 copy 自旧版 `tailwind.config.js`,改风格先看那边。
- Worker 用 Vite 原生 `new Worker(new URL(...), { type: 'module' })`,文件 buffer 用 transferable 移交,大文件零拷贝。
- 文件分类按文件名正则(滞销 / 拿货历史 / 销售明细…),UI 和 Worker 共用 `utils/files.ts`,改关键词只改一处。
- 第一版**严格复刻 Python 版输出**:不加智能时间窗、不拉长横轴、不用原生 sparkline(用户明确拒绝过)。
- 滞销表未知字段**自动透传**是核心需求,M3/M5 实现时参考 `pipeline.py` 的 `KNOWN_ZHIXIAO_FIELDS` + `extra_fields`。
