// errors.ts — 对应 pipeline.py 的 PipelineError(msg + hint)
export class PipelineError extends Error {
  hint: string;
  constructor(msg: string, hint = '') {
    super(msg);
    this.name = 'PipelineError';
    this.hint = hint;
  }
}
