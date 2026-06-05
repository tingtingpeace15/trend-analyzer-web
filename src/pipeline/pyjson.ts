// pyjson.ts — Python json.dumps(obj, ensure_ascii=False) 的字节级复刻。
// html 报告里嵌的是这个 JSON,M13 做字符串 diff,所以三处细节必须一致:
//   1. 默认分隔符是 ", " 和 ": "(JS JSON.stringify 是紧凑的)
//   2. Python float 即使整数值也带小数点(0.0 / 100.0),int 不带 —— 用 PyFloat 标记
//   3. 字符串转义规则与 JSON.stringify 相同(双引号/反斜杠/控制符)

/** 标记"Python 里是 float"的数值(json 序列化成 0.0 而不是 0) */
export class PyFloat {
  constructor(public readonly v: number) {}
}
export const pyFloat = (v: number) => new PyFloat(v);

/** 取数值(PyFloat 或普通 number) */
export const numOf = (v: unknown): number => (v instanceof PyFloat ? v.v : (v as number));

function formatFloat(v: number): string {
  // Python repr(float):整数值带 .0,其余最短表示(与 JS String 一致)
  if (Number.isInteger(v)) return `${v}.0`;
  return String(v);
}

export function pyJsonDumps(o: unknown): string {
  if (o === null || o === undefined) return 'null';
  if (o instanceof PyFloat) return formatFloat(o.v);
  switch (typeof o) {
    case 'number':
      return Number.isInteger(o) ? String(o) : formatFloat(o);
    case 'string':
      return JSON.stringify(o);
    case 'boolean':
      return o ? 'true' : 'false';
  }
  if (Array.isArray(o)) return `[${o.map(pyJsonDumps).join(', ')}]`;
  return `{${Object.entries(o as Record<string, unknown>)
    .map(([k, v]) => `${JSON.stringify(k)}: ${pyJsonDumps(v)}`)
    .join(', ')}}`;
}
