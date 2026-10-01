import { ApiError } from './shared.js';
// JSON을 객체로 펼치지 않고 문법만 검증합니다. 큰 문서도 메모리를 청크 크기로 제한합니다.
export class JsonValidator {
  constructor() { this.stack = []; this.root = 'value'; this.mode = ''; this.token = ''; this.escape = false; this.unicode = 0; this.key = false; this.values = 0; }
  bad() { throw new ApiError(400, '올바른 JSON 문서가 아닙니다.', 'INVALID_JSON'); }
  state() { return this.stack.length ? this.stack[this.stack.length - 1].state : this.root; }
  complete() {
    if (++this.values > 3000000) throw new ApiError(413, '문서 구조가 너무 복잡합니다.', 'DOCUMENT_COMPLEXITY');
    if (this.stack.length) this.stack[this.stack.length - 1].state = 'comma'; else this.root = 'end';
  }
  finishToken() {
    if (this.mode === 'number' && !/^-?(?:0|[1-9]\d*)(?:\.\d+)?(?:[eE][+-]?\d+)?$/.test(this.token)) this.bad();
    if (this.mode === 'literal' && !['true', 'false', 'null'].includes(this.token)) this.bad();
    this.mode = ''; this.token = ''; this.complete();
  }
  write(text) {
    for (let i = 0; i < text.length; i++) {
      const c = text[i];
      if (this.mode === 'string') {
        if (this.unicode) { if (!/[0-9a-fA-F]/.test(c)) this.bad(); this.unicode--; continue; }
        if (this.escape) { this.escape = false; if (c === 'u') this.unicode = 4; else if (!'"\\/bfnrt'.includes(c)) this.bad(); continue; }
        if (c === '\\') { this.escape = true; continue; }
        if (c === '"') { this.mode = ''; if (this.key) this.stack[this.stack.length - 1].state = 'colon'; else this.complete(); continue; }
        if (c.charCodeAt(0) < 32) this.bad();
        continue;
      }
      if (this.mode === 'number' || this.mode === 'literal') {
        if (/[,\]}\s]/.test(c)) { this.finishToken(); i--; continue; }
        this.token += c; if (this.token.length > 128) this.bad(); continue;
      }
      if (c === ' ' || c === '\n' || c === '\r' || c === '\t') continue;
      const state = this.state(), frame = this.stack[this.stack.length - 1];
      if (state === 'end') this.bad();
      if (state === 'colon') { if (c !== ':') this.bad(); frame.state = 'value'; continue; }
      if (state === 'comma') {
        if (c === ',') { frame.state = frame.kind === '{' ? 'key' : 'value'; continue; }
        if (c === (frame.kind === '{' ? '}' : ']')) { this.stack.pop(); this.complete(); continue; }
        this.bad();
      }
      if (state === 'key' || state === 'keyOrEnd') {
        if (state === 'keyOrEnd' && c === '}') { this.stack.pop(); this.complete(); continue; }
        if (c !== '"') this.bad(); this.mode = 'string'; this.key = true; continue;
      }
      if (state === 'valueOrEnd' && c === ']') { this.stack.pop(); this.complete(); continue; }
      if (c === '{' || c === '[') {
        if (this.stack.length >= 128) throw new ApiError(413, '문서의 중첩 단계가 너무 깊습니다.', 'DOCUMENT_COMPLEXITY');
        this.stack.push({ kind: c, state: c === '{' ? 'keyOrEnd' : 'valueOrEnd' }); continue;
      }
      if (c === '"') { this.mode = 'string'; this.key = false; continue; }
      if (c === '-' || /[0-9]/.test(c)) { this.mode = 'number'; this.token = c; continue; }
      if ('tfn'.includes(c)) { this.mode = 'literal'; this.token = c; continue; }
      this.bad();
    }
  }
  finish() {
    if (this.mode === 'number' || this.mode === 'literal') this.finishToken();
    if (this.mode || this.escape || this.unicode || this.stack.length || this.root !== 'end') this.bad();
  }
}
