import { describe, expect, it } from 'vitest';
import { buildNewPreferenceHtml } from '../pipeline/preferenceNewHtml';
import type { PreferenceData } from '../pipeline/preferenceReader';
import type { Cell } from '../types/excel';

function reportData(html: string) {
  const match = html.match(/<script type="application\/json" id="reportData">(.*?)<\/script>/s);
  expect(match).not.toBeNull();
  return JSON.parse(match![1]);
}

describe('销售分析全表明细口径', () => {
  it('保留净成交非正客户的明细，但不把该客户加入客户画像列表', () => {
    const data: PreferenceData = {
      sheetNames: ['明细'],
      sheetMeta: [],
      rawRowCount: 4,
      columns: ['客户名称', '分类', '品牌', '设计师品牌', '颜色', '尺码', '货号', '年份', '销售量', '销售金额'],
      droppedConflicts: [],
      cols: new Map<string, (Cell | null)[]>([
        ['客户名称', ['客户A', '净退货客户', '客户C', null]],
        ['分类', ['上衣', '上衣', '裤装', '上衣']],
        ['品牌', ['品牌A', '品牌A', '品牌B', '品牌A']],
        ['设计师品牌', ['设计师A', '设计师B', '设计师C', '设计师A']],
        ['颜色', ['黑色', '黑色', '杂色', '黄色']],
        ['尺码', ['M', 'M', 'L', 'L']],
        ['货号', ['A1', 'A1', 'C1', 'D1']],
        ['年份', ['2026秋', '2026秋', '2026秋', '2026秋']],
      ]),
      qty: new Float64Array([10, -2, 10, 1]),
      amt: new Float64Array([100, -20, 200, 10]),
      orderMs: new Float64Array([1, 2, 3, 4]),
    };

    const report = reportData(buildNewPreferenceHtml(data, ['1', '2', '3', '4'], { summary: {} }));
    const visual = report.customer_visual_profiles;
    expect(visual.customers).toEqual(['客户C', '客户A']);
    expect(visual.profiles.map((profile: { customer: string }) => profile.customer)).toContain('净退货客户');
    expect(visual.profiles.map((profile: { customer: string }) => profile.customer)).toContain('');

    const rawColors = visual.detail_row_dicts.raw_colors as string[];
    const colorQty = new Map<string, number>();
    for (const profile of visual.profiles) {
      for (const row of profile.detail_rows) {
        const color = rawColors[row[12]];
        colorQty.set(color, (colorQty.get(color) ?? 0) + row[1]);
      }
    }
    expect(colorQty.get('黑色')).toBe(8);
    expect(colorQty.get('杂色')).toBe(10);
    expect([...colorQty.values()].filter((qty) => qty > 0).reduce((sum, qty) => sum + qty, 0)).toBe(19);

    const negativeDesigner = report.brand_style_analysis.category_share_details
      .find((row: { designer: string }) => row.designer === '设计师B');
    expect(negativeDesigner).toMatchObject({ category: '上衣', brand: '品牌A', amount: -20, qty: -2 });
  });
});
