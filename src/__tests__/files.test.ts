// files.test.ts — 上传文件名识别规则(2026-11-16 用户指定的关键词)
import { describe, expect, it } from 'vitest';
import { classifyFile, classifyPreferenceFile } from '../utils/files';

const HISTORY_NAMES = [
  '各商品客户拿货历史.xlsx',
  '商品客户拿货历史_1116.xlsx',
  '客户拿货历史.xls',
  '拿货历史(10月).xlsx',
  '各商品客户拿货历史【含散客】_pMNksfRcPwcpwhd.xlsx',
  '商品客户拿货·含散客.xlsx',
  '各商品客户拿货·含散客_导出.xlsx',
  // 间隔点的其他写法 / 没有点
  '商品客户拿货・含散客.xlsx',
  '商品客户拿货•含散客.xlsx',
  '商品客户拿货含散客.xlsx',
];

describe('商品趋势分析 classifyFile', () => {
  it('含「滞销」→ 滞销表', () => {
    expect(classifyFile('滞销商品.xlsx')).toBe('滞销表');
    expect(classifyFile('2026年11月滞销款.xls')).toBe('滞销表');
  });

  it.each(HISTORY_NAMES)('%s → 销售明细', (name) => {
    expect(classifyFile(name)).toBe('销售明细');
  });

  it('旧关键词仍兼容', () => {
    expect(classifyFile('销售明细.xlsx')).toBe('销售明细');
    expect(classifyFile('销售单明细.xlsx')).toBe('销售明细');
    expect(classifyFile('销售流水.xlsx')).toBe('销售明细');
  });

  it('同时含「滞销」和「拿货历史」时算滞销表', () => {
    expect(classifyFile('滞销商品客户拿货历史.xlsx')).toBe('滞销表');
  });

  it('不含关键词 → 无法识别', () => {
    expect(classifyFile('库存表.xlsx')).toBeNull();
  });
});

describe('客户偏好分析 classifyPreferenceFile', () => {
  it.each(HISTORY_NAMES)('%s → 拿货历史', (name) => {
    expect(classifyPreferenceFile(name)).toBe('拿货历史');
  });

  it('旧关键词仍兼容', () => {
    expect(classifyPreferenceFile('客户偏好.xlsx')).toBe('拿货历史');
    expect(classifyPreferenceFile('拿货明细.xlsx')).toBe('拿货历史');
  });

  it('不含关键词 → 无法识别', () => {
    expect(classifyPreferenceFile('滞销商品.xlsx')).toBeNull();
    expect(classifyPreferenceFile('库存表.xlsx')).toBeNull();
  });
});
