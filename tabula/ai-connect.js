const $ = id => document.getElementById(id), request = new URL(location.href).searchParams.get('request');
let pending = null, busy = false, loadedKey = '';
try { $('key').value = JSON.parse(localStorage.getItem('wixel.connection.v3'))?.key ?? ''; } catch { /* 사용자가 복구키를 입력할 수 있습니다. */ }
async function api(path, options = {}) {
  const response = await fetch('/api/' + path, { ...options, cache: 'no-store', headers: { 'Content-Type': 'application/json', 'X-Wixel-Vault': $('key').value.trim() } });
  const data = await response.json(); if (!response.ok) throw new Error(data.error || '연결 요청에 실패했습니다.'); return data;
}
async function load() {
  $('allow').disabled = true; $('documents').replaceChildren(); loadedKey = '';
  try {
    const key = $('key').value.trim(); if (!/^[A-Za-z0-9_-]{43}$/.test(key)) throw new Error('올바른 보관함 복구키를 입력하세요.');
    const docs = await api('files'); if ($('key').value.trim() !== key) return;
    for (const d of docs) { const option = document.createElement('option'); option.value = option.textContent = d.name; $('documents').append(option); }
    loadedKey = key; $('allow').disabled = !pending || !docs.length;
    $('status').textContent = docs.length ? '' : '이 보관함에 문서가 없습니다. 위셀에서 문서를 온라인 보관함에 먼저 저장하세요.';
  } catch (e) { $('status').textContent = e.message; }
}
$('key').addEventListener('input', () => { loadedKey = ''; $('allow').disabled = true; });
$('load').addEventListener('click', load);
async function finish(deny) {
  if (busy || !pending) return;
  if (!deny && (!$('documents').value || loadedKey !== $('key').value.trim())) { $('status').textContent = '보관함 문서를 먼저 불러오세요.'; return; }
  if (!deny && pending.write && !$('write').checked) { $('status').textContent = '앱이 요청한 편집 권한을 확인하거나 취소하세요.'; return; }
  busy = true; $('allow').disabled = true; $('deny').disabled = true;
  try {
    const result = await api(deny ? 'oauth/deny' : 'ai/consent', { method: 'POST', body: JSON.stringify({ request, document: $('documents').value, write: $('write').checked }) });
    $('key').value = ''; location.replace(result.redirect);
  } catch (e) { $('status').textContent = e.message; busy = false; $('deny').disabled = false; }
}
$('consent').addEventListener('submit', e => { e.preventDefault(); void finish(false); });
$('deny').addEventListener('click', () => finish(true));
try {
  pending = await api('oauth/pending?request=' + encodeURIComponent(request ?? ''));
  $('client').textContent = `${pending.label} 앱이 문서 ${pending.write ? '읽기·편집' : '읽기'} 권한을 요청했습니다.`;
  $('destination').textContent = '돌아갈 주소: ' + pending.redirect;
  $('write').parentElement.hidden = !pending.write;
  if ($('key').value) await load();
} catch (e) { $('status').textContent = e.message; $('load').disabled = true; }
