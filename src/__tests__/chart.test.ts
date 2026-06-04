// chart.test.ts — JS 图表几何 vs matplotlib 黄金参数(charts.golden.json)。
// 像素级逐位对比做不到(字体光栅化不同),对照的是几何:画布尺寸、ylim、
// axes 框、数据→像素仿射探针、刻度位置/标签。容差:尺寸 ±2px、axes/仿射 ±3px。
import { readFileSync } from 'node:fs';
import { resolve } from 'node:path';
import { describe, expect, it } from 'vitest';
import { readSales, readZhixiao } from '../pipeline/reader';
import { aggregate } from '../pipeline/aggregator';
import { buildDetailScene, buildSmallScene } from '../pipeline/chart';
import type { ChartScene } from '../pipeline/chart';
import readerGolden from './golden/reader.golden.json';
import golden from './golden/charts.golden.json';

const BASE_DIR = resolve(process.cwd(), '..', '趋势分析工具 -最终app版本');

function load(name: string): Uint8Array {
  return new Uint8Array(readFileSync(resolve(BASE_DIR, name)));
}

const zhixiao = readZhixiao(load(readerGolden.zhixiao.file));
const sales = readSales(load(readerGolden.sales.file), readerGolden.sales.file);
const agg = aggregate(zhixiao, sales);

function valuesFor(sample: (typeof golden.samples)[number]): number[] {
  if (sample.货号 == null) return new Array(golden.nDays).fill(0);
  const row = agg.dailyByItem.get(sample.货号);
  expect(row, `透视行 ${sample.货号}`).toBeDefined();
  return Array.from(row!);
}

interface GoldenChart {
  pngSize: number[];
  ylim: number[];
  axesPx: { x0: number; y0: number; w: number; h: number };
  probe: { p0: { data: number[]; px: number[] }; pEnd: { data: number[]; px: number[] } };
  tickPos?: number[];
  tickLabels?: string[];
  yticks?: number[];
  ytickLabels?: string[];
}

function dataToPx(scene: ChartScene, xd: number, yd: number): [number, number] {
  const { ax, ylim, xlim } = scene.meta;
  return [
    ax.x0 + ((xd - xlim[0]) / (xlim[1] - xlim[0])) * ax.w,
    ax.y0 + ((ylim[1] - yd) / (ylim[1] - ylim[0])) * ax.h,
  ];
}

function checkScene(scene: ChartScene, g: GoldenChart, label: string) {
  // 画布尺寸 ±2px(matplotlib tight crop 受字体度量影响有 ±1px 浮动)
  expect(Math.abs(scene.width - g.pngSize[0]), `${label} width`).toBeLessThanOrEqual(2);
  expect(Math.abs(scene.height - g.pngSize[1]), `${label} height`).toBeLessThanOrEqual(2);
  // ylim 精确(纯公式)
  expect(scene.meta.ylim[0], `${label} ylim0`).toBeCloseTo(g.ylim[0], 6);
  expect(scene.meta.ylim[1], `${label} ylim1`).toBeCloseTo(g.ylim[1], 6);
  // axes 框 ±3px
  expect(Math.abs(scene.meta.ax.x0 - g.axesPx.x0), `${label} ax.x0`).toBeLessThanOrEqual(3);
  expect(Math.abs(scene.meta.ax.y0 - g.axesPx.y0), `${label} ax.y0`).toBeLessThanOrEqual(3);
  expect(Math.abs(scene.meta.ax.w - g.axesPx.w), `${label} ax.w`).toBeLessThanOrEqual(3);
  expect(Math.abs(scene.meta.ax.h - g.axesPx.h), `${label} ax.h`).toBeLessThanOrEqual(3);
  // 仿射探针 ±3px
  for (const [name, probe] of Object.entries(g.probe)) {
    const [px, py] = dataToPx(scene, probe.data[0], probe.data[1]);
    expect(Math.abs(px - probe.px[0]), `${label} ${name}.x`).toBeLessThanOrEqual(3);
    expect(Math.abs(py - probe.px[1]), `${label} ${name}.y`).toBeLessThanOrEqual(3);
  }
}

describe('图表几何 vs matplotlib', () => {
  for (const sample of golden.samples) {
    it(`${sample.id} 小图几何`, () => {
      checkScene(buildSmallScene(valuesFor(sample)), sample.small, `${sample.id}/small`);
    });

    it(`${sample.id} 详情图几何 + 刻度`, () => {
      const scene = buildDetailScene(valuesFor(sample), agg.dateRange);
      const g = sample.detail as GoldenChart;
      checkScene(scene, g, `${sample.id}/detail`);
      // x 刻度位置/标签精确(纯整数公式)
      expect(scene.meta.tickPos, 'tickPos').toEqual(g.tickPos);
      // y 刻度数值精确(locator 公式);标签做 unicode minus 归一化后对比
      expect(scene.meta.yticks, 'yticks').toEqual(g.yticks);
      const norm = (s: string) => s.replace('−', '-');
      expect(scene.meta.ytickLabels!.map(norm), 'ytickLabels').toEqual(g.ytickLabels!.map(norm));
    });
  }
});
