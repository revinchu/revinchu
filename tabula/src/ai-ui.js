import { el, openDialog } from './ui.js';
import { server } from './storage.js';
import { AI_PROVIDERS } from './ai-providers.js';
import { aiChanges } from './ai-cells.js';

const keys = new Map();
const field = (label, input) => el('label', { class: 'ai-field' }, el('span', {}, label), input);
const stopKeys = e => e.stopPropagation();
function available(capability, retry, connect) {
  if (!server.capabilities[capability]) { openDialog({ title: 'AI 연결', body: el('p', {}, 'AI 연결은 온라인 보관함을 제공하는 공개 위셀에서 사용할 수 있습니다.', el('a', { href: 'https://wixel-3.wizx.workers.dev/', target: '_blank', rel: 'noopener' }, ' 공개 위셀 열기')), buttons: [{ label: '닫기' }] }); return false; }
  if (!server.connected) { connect(retry); return false; } return true;
}
export function parseAiAnswer(text) {
  const clean = text.trim().replace(/^```(?:json)?\s*/i, '').replace(/\s*```$/, '');
  let result; try { result = JSON.parse(clean); } catch { return { answer: text, changes: [] }; }
  if (!result || typeof result.answer !== 'string') return { answer: text, changes: [] };
  return { answer: result.answer, changes: result.changes?.length ? aiChanges(result.changes) : [] };
}
export function aiAssistantDialog(options) {
  if (!available('ai', () => aiAssistantDialog(options), options.connect)) return;
  let control = null, closed = false, proposal = null, captured = null;
  const provider = el('select', { 'aria-label': 'AI 서비스' }, Object.entries(AI_PROVIDERS).map(([id, p]) => el('option', { value: id }, p.label)));
  const model = el('input', { value: AI_PROVIDERS.openai.model, 'aria-label': '모델 ID', spellcheck: 'false' });
  const key = el('input', { type: 'password', autocomplete: 'off', value: keys.get('openai') ?? '', 'aria-label': 'AI API 키', placeholder: '선택한 서비스에서 발급한 API 키' });
  const include = el('input', { type: 'checkbox' });
  const prompt = el('textarea', { rows: 4, maxLength: 16000, 'aria-label': 'AI 요청', placeholder: '예: 선택한 광고 지표를 요약하고 ROAS 계산 수식을 제안해 줘.' });
  const status = el('p', { role: 'status', class: 'muted' });
  const answer = el('div', { class: 'ai-answer', 'aria-live': 'polite' });
  const preview = el('div', { class: 'ai-preview' });
  const apply = el('button', { class: 'btn', disabled: true, onclick: () => {
    try { options.apply(captured, proposal); apply.disabled = true; proposal = null; status.textContent = '셀 변경을 적용했습니다. Ctrl+Z로 실행 취소할 수 있습니다.'; }
    catch (e) { status.textContent = e.message; }
  } }, '변경안 적용');
  const cancel = el('button', { class: 'btn', disabled: true, onclick: () => control?.abort() }, '요청 중지');
  const send = el('button', { class: 'btn primary', onclick: async () => {
    if (control) return;
    proposal = null; apply.disabled = true; preview.replaceChildren(); answer.textContent = '';
    try {
      captured = options.capture(include.checked); keys.set(provider.value, key.value.trim());
      control = new AbortController(); send.disabled = true; cancel.disabled = false; status.textContent = 'AI가 답변을 작성하고 있습니다…';
      const response = await server.request('ai/chat', { method: 'POST', signal: control.signal, headers: { 'X-Wixel-AI-Key': key.value.trim() }, body: JSON.stringify({ provider: provider.value, model: model.value.trim(), prompt: prompt.value, context: captured.context }) });
      if (closed) return;
      const result = parseAiAnswer(response.text); answer.textContent = result.answer;
      for (const url of response.citations ?? []) answer.append(el('p', {}, el('a', { href: url, target: '_blank', rel: 'noopener noreferrer' }, url)));
      proposal = result.changes;
      if (proposal.length) {
        preview.append(el('p', {}, `${captured.sheetName} 시트의 변경안 ${proposal.length}개`), el('table', {},
          el('thead', {}, el('tr', {}, ['셀', '현재 입력', '변경할 입력'].map(x => el('th', {}, x)))),
          el('tbody', {}, proposal.map(c => el('tr', {}, el('td', {}, c.cell), el('td', {}, options.before(captured, c)), el('td', {}, c.input))))));
        apply.disabled = false;
      }
      status.textContent = proposal.length ? '변경할 셀과 입력을 확인한 뒤 적용하세요.' : '답변을 받았습니다.';
    } catch (e) { if (!closed) status.textContent = e.name === 'AbortError' ? 'AI 요청을 중지했습니다.' : e.message; }
    finally { control = null; send.disabled = false; cancel.disabled = true; }
  } }, 'AI에게 요청');
  provider.addEventListener('change', () => { model.value = AI_PROVIDERS[provider.value].model; key.value = keys.get(provider.value) ?? ''; });
  const body = el('div', { class: 'ai-panel', onkeydown: stopKeys },
    el('p', { class: 'muted' }, '서비스별 API 키가 필요하며 API 사용료는 해당 서비스에서 청구됩니다. 키는 이 탭 메모리에만 유지되고 요청을 중계할 때 사용됩니다.'),
    field('AI 서비스', provider), field('모델 ID', model), el('p', { class: 'muted' }, '모델 ID는 해당 계정에서 사용할 수 있는 모델로 바꿀 수 있습니다.'), field('API 키', key),
    el('button', { class: 'btn', onclick: () => { keys.clear(); key.value = ''; status.textContent = '이 탭에 입력한 모든 AI 키를 지웠습니다.'; } }, '입력한 키 모두 지우기'),
    field('요청', prompt), el('label', {}, include, ' 선택한 셀을 함께 전송 (최대 500셀, 입력·저장된 값 포함)'),
    el('p', { class: 'muted' }, '선택을 켜면 요청한 시트 자료가 선택한 AI 서비스로 전송됩니다. 전체 파일은 전송하지 않습니다.'),
    el('div', { class: 'ai-actions' }, send, cancel, apply), status, answer, preview);
  openDialog({ title: 'AI 도우미', width: 760, body, buttons: [{ label: '닫기' }], onClose: () => { closed = true; control?.abort(); key.value = ''; } });
}
export function aiConnectionsDialog(options) {
  if (!available('mcp', () => aiConnectionsDialog(options), options.connect)) return;
  let closed = false, busy = false;
  const endpoint = new URL('api/mcp', location.href).href;
  const url = el('input', { value: endpoint, readOnly: true, 'aria-label': 'MCP 서버 주소' });
  const docs = el('select', { 'aria-label': '연결할 온라인 문서' });
  const label = el('input', { value: '내 AI 연결', maxLength: 80, 'aria-label': '연결 이름' });
  const write = el('input', { type: 'checkbox' });
  const result = el('textarea', { rows: 3, readOnly: true, hidden: true, 'aria-label': 'MCP 연결 토큰' });
  const status = el('p', { role: 'status' });
  const list = el('div', { class: 'ai-connections' });
  const refresh = async () => {
    try {
      const [files, connections] = await Promise.all([server.list(), server.request('ai/connections')]); if (closed) return;
      docs.replaceChildren(...files.map(f => el('option', { value: f.name }, f.name)));
      list.replaceChildren(...connections.map(c => el('div', { class: 'ai-connection' },
        el('span', {}, `${c.label} · ${c.document} · ${c.write ? '읽기·편집' : '읽기 전용'} · ${new Date(c.expires).toLocaleDateString('ko-KR')} 만료`),
        el('button', { class: 'btn', onclick: async e => {
          e.currentTarget.disabled = true;
          try { await server.request('ai/connections/' + c.id, { method: 'DELETE' }); result.value = ''; result.hidden = true; await refresh(); status.textContent = '연결을 해제했습니다. 해당 토큰으로 문서를 열 수 없습니다.'; }
          catch (error) { status.textContent = error.message; e.currentTarget.disabled = false; }
        } }, '연결 해제'))));
      if (!files.length) status.textContent = '먼저 파일 → 저장 위치에서 온라인 개인 보관함에 문서를 저장하세요.';
    } catch (e) { if (!closed) status.textContent = e.message; }
  };
  const create = el('button', { class: 'btn primary', onclick: async () => {
    if (busy) return; busy = true; create.disabled = true;
    try {
      if (!docs.value) throw new Error('연결할 온라인 문서를 선택하세요.');
      const response = await server.request('ai/connections', { method: 'POST', body: JSON.stringify({ document: docs.value, label: label.value, write: write.checked }) });
      if (closed) return;
      result.value = response.token; result.hidden = false; await refresh();
      status.textContent = '토큰은 지금 한 번만 표시됩니다. MCP 클라이언트의 Authorization 헤더에 Bearer 뒤에 붙여 넣으세요. 30일 후 만료됩니다.';
    } catch (e) { if (!closed) status.textContent = e.message; }
    finally { busy = false; create.disabled = false; }
  } }, '수동 연결 토큰 발급');
  const body = el('div', { class: 'ai-panel', onkeydown: stopKeys },
    el('h3', {}, 'ChatGPT·Claude 등에서 위셀 연결'),
    el('p', {}, 'AI 앱의 사용자 지정 MCP 서버에 아래 주소를 추가하고 OAuth 인증을 선택하세요. 클라이언트 등록은 자동 등록(DCR)을 사용합니다. 위셀 연결 화면에서 허용할 문서와 권한을 확인할 수 있습니다.'),
    field('MCP 서버 주소', url), el('button', { class: 'btn', onclick: () => options.copy(endpoint, 'MCP 주소를 복사했습니다.', url) }, '주소 복사'),
    el('p', { class: 'muted' }, '외부 앱의 MCP 지원 여부는 앱·플랜에 따라 다릅니다. Grok·DeepSeek 등을 포함한 다섯 서비스는 위셀의 AI 도우미에서도 API로 사용할 수 있습니다. MCP는 온라인 저장본에 연결되며 이 탭의 미저장 내용은 포함하지 않습니다.'),
    el('p', { class: 'muted' }, 'MCP: 문서당 압축 해제 후 8MiB·10만 셀, 한 번에 읽기 1,000셀·편집 200셀. 수식은 저장된 값을 읽습니다. 보호·피벗 시트는 AI 편집을 제한합니다.'),
    el('details', {}, el('summary', {}, '토큰을 직접 입력하는 MCP 클라이언트'), field('연결 이름', label), field('온라인 문서', docs), el('label', {}, write, ' 이 문서의 셀 편집 허용'), create, result,
      el('button', { class: 'btn', onclick: () => { if (result.value) options.copy(result.value, '연결 토큰을 복사했습니다.', result); } }, '토큰 복사')),
    status, el('h3', {}, '허용한 연결'), list, el('button', { class: 'btn', onclick: refresh }, '목록 새로 고침'));
  openDialog({ title: '외부 AI · MCP 연결', width: 760, body, buttons: [{ label: '닫기' }], onClose: () => { closed = true; result.value = ''; } });
  void refresh();
}
