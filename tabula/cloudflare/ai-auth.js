import { ApiError, json, randomId, sha256, validateOrigin, validateName } from './shared.js';
import { boundedBytes } from './proxy.js';

const DAY = 86400000, SCOPES = ['document:read', 'document:write'];
export const MCP_PATH = '/api/mcp';
export async function aiBody(request, limit = 131072) {
  if (request.headers.has('Content-Encoding')) throw new ApiError(415, '압축하지 않은 요청을 전송하세요.');
  const bytes = await boundedBytes(new Response(request.body, { headers: request.headers }), limit, AbortSignal.timeout(10000));
  try { return JSON.parse(new TextDecoder('utf-8', { fatal: true }).decode(bytes)); }
  catch { throw new ApiError(400, '올바른 JSON 요청을 전송하세요.'); }
}
function state(env, type, id) { return env.AI_STATE.getByName(`${type}:${id}`); }
function required(value, test, message) { if (typeof value !== 'string' || !test.test(value)) throw new ApiError(400, message, 'invalid_request'); return value; }
function opaque(value) { return required(value, /^[A-Za-z0-9_-]{43}$/, '연결 식별자를 확인하세요.'); }
function cookie(request) { return /(?:^|;\s*)wixel_ai_consent=([A-Za-z0-9_-]{43})(?:;|$)/.exec(request.headers.get('Cookie') ?? '')?.[1] ?? ''; }
export function validRedirect(value) {
  try {
    const u = new URL(value);
    return typeof value === 'string' && value.length < 2048 && !u.username && !u.password && !u.hash &&
      (u.protocol === 'https:' || u.protocol === 'http:' && ['localhost', '127.0.0.1', '[::1]'].includes(u.hostname));
  } catch { return false; }
}
function oauthError(error, status = 400) { return json({ error }, status); }
function scopeValue(value = 'document:read') {
  const scopes = value.split(' ').filter(Boolean);
  if (!scopes.length || scopes.some(s => !SCOPES.includes(s))) throw new ApiError(400, '지원하는 문서 권한을 선택하세요.', 'invalid_scope');
  return [...new Set(['document:read', ...scopes])].join(' ');
}
export function oauthMetadata(origin) {
  return { issuer: origin, authorization_endpoint: origin + '/api/oauth/authorize', token_endpoint: origin + '/api/oauth/token',
    registration_endpoint: origin + '/api/oauth/register', revocation_endpoint: origin + '/api/oauth/revoke',
    response_types_supported: ['code'], grant_types_supported: ['authorization_code', 'refresh_token'], token_endpoint_auth_methods_supported: ['none'],
    code_challenge_methods_supported: ['S256'], scopes_supported: SCOPES, authorization_response_iss_parameter_supported: true };
}
export function protectedMetadata(origin) { return { resource: origin + MCP_PATH, authorization_servers: [origin], scopes_supported: SCOPES, bearer_methods_supported: ['header'], resource_name: '위셀 문서' }; }
function tokenParts(token) { const match = /^([A-Za-z0-9_-]{43})\.([A-Za-z0-9_-]{43})$/.exec(token ?? ''); return match ? { id: match[1], secret: match[2] } : null; }
export async function authenticateMcp(request, env) {
  const p = tokenParts(/^Bearer (.+)$/i.exec(request.headers.get('Authorization') ?? '')?.[1]);
  return p ? state(env, 'grant', p.id).authenticate(await sha256(p.secret)) : null;
}
export function mcpChallenge(origin) { return json({ error: 'AI 연결 인증이 필요합니다.' }, 401, { 'WWW-Authenticate': `Bearer resource_metadata="${origin}/.well-known/oauth-protected-resource/api/mcp"` }); }
async function grant(env, owner, vault, { document, label, write, client = '', resource }) {
  document = validateName(document);
  const docs = await (await vault.list()).json();
  if (!docs.some(d => d.name === document)) throw new ApiError(404, '연결할 온라인 문서를 먼저 저장하세요.');
  const id = randomId(), access = randomId(), refresh = randomId(), now = Date.now();
  const record = { kind: 'grant', id, owner, document, label: String(label || 'MCP 연결').slice(0, 80), write: write === true, client, resource,
    accessHash: await sha256(access), refreshHash: client ? await sha256(refresh) : null,
    accessExpires: now + (client ? 3600000 : 30 * DAY), created: now, expires: now + 30 * DAY };
  const summary = { id, document, label: record.label, write: record.write, created: now, expires: record.expires };
  if (!await vault.addAiGrant(summary)) throw new ApiError(409, 'AI 연결은 최대 20개입니다. 기존 연결을 해제하세요.');
  try { if (!await state(env, 'grant', id).create(record)) throw new Error('collision'); }
  catch (e) { await vault.removeAiGrant(id); throw e; }
  return { record, access: id + '.' + access, refresh: id + '.' + refresh };
}
export async function manageAi(request, env, owner, vault) {
  const u = new URL(request.url), path = u.pathname;
  if (path === '/api/ai/connections' && request.method === 'GET') return json(await vault.aiGrants());
  if (path === '/api/ai/connections' && request.method === 'POST') {
    const data = await aiBody(request, 8192), g = await grant(env, owner, vault, { ...data, client: '', resource: u.origin + MCP_PATH });
    return json({ token: g.access, expires: g.record.expires, id: g.record.id }, 201);
  }
  if (path.startsWith('/api/ai/connections/') && request.method === 'DELETE') {
    const id = opaque(path.slice('/api/ai/connections/'.length));
    await state(env, 'grant', id).revoke(owner); await vault.removeAiGrant(id); return json({ ok: true });
  }
  if (path === '/api/ai/consent' && request.method === 'POST') {
    const data = await aiBody(request, 8192), id = opaque(data.request), csrf = cookie(request);
    if (!csrf) throw new ApiError(403, '연결 화면을 다시 여세요.');
    const pending = await state(env, 'pending', id).consume({ csrfHash: await sha256(csrf) });
    if (!pending) throw new ApiError(400, '연결 요청이 만료되었거나 이미 사용되었습니다. AI 앱에서 다시 연결하세요.');
    const redirect = new URL(pending.redirect);
    redirect.searchParams.set('state', pending.state); redirect.searchParams.set('iss', u.origin);
    if (data.deny) { redirect.searchParams.set('error', 'access_denied'); return json({ redirect: redirect.href }); }
    if (pending.scope.includes('document:write') && data.write !== true) throw new ApiError(400, '편집 권한을 확인하거나 연결을 취소하세요.');
    const g = await grant(env, owner, vault, { document: data.document, label: pending.label, write: pending.scope.includes('document:write'), client: pending.client, resource: pending.resource });
    const code = randomId();
    try {
      await state(env, 'code', await sha256(code)).create({ ...pending, access: g.access, refresh: g.refresh, scope: pending.scope, expires: Date.now() + 120000 });
    } catch (e) { await state(env, 'grant', g.record.id).revoke(owner); await vault.removeAiGrant(g.record.id); throw e; }
    redirect.searchParams.set('code', code); return json({ redirect: redirect.href });
  }
  return null;
}
export async function oauthRoute(request, env) {
  const u = new URL(request.url), path = u.pathname, resource = u.origin + MCP_PATH;
  if (request.method === 'GET' && path === '/.well-known/oauth-authorization-server') return json(oauthMetadata(u.origin));
  if (request.method === 'GET' && ['/.well-known/oauth-protected-resource', '/.well-known/oauth-protected-resource/api/mcp'].includes(path)) return json(protectedMetadata(u.origin));
  if (path === '/api/oauth/register' && request.method === 'POST') {
    const data = await aiBody(request, 16384);
    if (!data || !Array.isArray(data.redirect_uris) || !data.redirect_uris.length || data.redirect_uris.length > 5 || !data.redirect_uris.every(validRedirect) || data.token_endpoint_auth_method && data.token_endpoint_auth_method !== 'none') return oauthError('invalid_client_metadata');
    if (data.grant_types !== undefined && (!Array.isArray(data.grant_types) || data.grant_types.some(x => !['authorization_code', 'refresh_token'].includes(x))) || data.response_types !== undefined && (!Array.isArray(data.response_types) || data.response_types.some(x => x !== 'code'))) return oauthError('invalid_client_metadata');
    const id = randomId(), now = Date.now();
    const record = { kind: 'client', client_id: id, redirect_uris: data.redirect_uris, client_name: String(data.client_name || 'MCP 앱').slice(0, 80), token_endpoint_auth_method: 'none', grant_types: ['authorization_code', 'refresh_token'], response_types: ['code'], expires: now + 365 * DAY };
    await state(env, 'client', id).create(record);
    const { kind, expires, ...publicData } = record; return json({ ...publicData, client_id_issued_at: Math.floor(now / 1000) }, 201);
  }
  if (path === '/api/oauth/authorize' && request.method === 'GET') {
    const p = u.searchParams, client = opaque(p.get('client_id')), registration = await state(env, 'client', client).read();
    const redirect = p.get('redirect_uri');
    if (!registration || !registration.redirect_uris.includes(redirect)) return oauthError('invalid_client');
    if (p.get('response_type') !== 'code' || p.get('code_challenge_method') !== 'S256' || p.get('resource') !== resource) return oauthError('invalid_request');
    const challenge = required(p.get('code_challenge'), /^[A-Za-z0-9_-]{43}$/, 'PKCE S256 값을 확인하세요.');
    const requestState = p.get('state') ?? ''; if (requestState.length > 2048) return oauthError('invalid_request');
    const id = randomId(), csrf = randomId();
    await state(env, 'pending', id).create({ kind: 'pending', client, redirect, state: requestState, resource, challenge, scope: scopeValue(p.get('scope') ?? undefined), label: registration.client_name, csrfHash: await sha256(csrf), expires: Date.now() + 600000 });
    return new Response(null, { status: 302, headers: { Location: '/ai-connect.html?request=' + id, 'Cache-Control': 'no-store', 'Set-Cookie': `wixel_ai_consent=${csrf}; Path=/api/; HttpOnly; Secure; SameSite=Lax; Max-Age=600` } });
  }
  if (path === '/api/oauth/pending' && request.method === 'GET') {
    validateOrigin(request); const r = await state(env, 'pending', opaque(u.searchParams.get('request'))).read();
    if (!r || !cookie(request) || r.csrfHash !== await sha256(cookie(request))) throw new ApiError(403, '연결 요청이 만료되었습니다. AI 앱에서 다시 연결하세요.');
    return json({ label: r.label, redirect: r.redirect, write: r.scope.includes('document:write') });
  }
  if (path === '/api/oauth/deny' && request.method === 'POST') {
    validateOrigin(request); const data = await aiBody(request, 8192), csrf = cookie(request);
    if (!csrf) throw new ApiError(403, '연결 요청을 확인할 수 없습니다.');
    const r = await state(env, 'pending', opaque(data.request)).consume({ csrfHash: await sha256(csrf) });
    if (!r) throw new ApiError(400, '연결 요청이 만료되었습니다.');
    const redirect = new URL(r.redirect); redirect.searchParams.set('state', r.state); redirect.searchParams.set('error', 'access_denied'); redirect.searchParams.set('iss', u.origin);
    return json({ redirect: redirect.href });
  }
  if (['/api/oauth/token', '/api/oauth/revoke'].includes(path) && request.method === 'POST') {
    if (!(request.headers.get('Content-Type') ?? '').startsWith('application/x-www-form-urlencoded')) return oauthError('invalid_request');
    const bytes = await boundedBytes(new Response(request.body), 16384, AbortSignal.timeout(10000)), p = new URLSearchParams(new TextDecoder().decode(bytes));
    if (path.endsWith('/revoke')) {
      const token = tokenParts(p.get('token')); if (token) await state(env, 'grant', token.id).revoke(null, await sha256(token.secret)); return json({});
    }
    if (p.get('resource') !== resource) return oauthError('invalid_target');
    if (p.get('grant_type') === 'authorization_code') {
      const code = p.get('code'), verifier = p.get('code_verifier');
      if (!/^[A-Za-z0-9_-]{43}$/.test(code ?? '') || !/^[A-Za-z0-9._~-]{43,128}$/.test(verifier ?? '')) return oauthError('invalid_grant');
      const hash = await sha256(verifier); let raw = ''; for (let i = 0; i < hash.length; i += 2) raw += String.fromCharCode(parseInt(hash.slice(i, i + 2), 16));
      const challenge = btoa(raw).replace(/\+/g, '-').replace(/\//g, '_').replace(/=+$/, '');
      const r = await state(env, 'code', await sha256(code)).consume({ client: p.get('client_id'), redirect: p.get('redirect_uri'), resource, challenge });
      if (!r) return oauthError('invalid_grant');
      return json({ access_token: r.access, refresh_token: r.refresh, token_type: 'Bearer', expires_in: 3600, scope: r.scope });
    }
    if (p.get('grant_type') === 'refresh_token') {
      const old = tokenParts(p.get('refresh_token')); if (!old) return oauthError('invalid_grant');
      const access = randomId(), refresh = randomId();
      const r = await state(env, 'grant', old.id).rotate(await sha256(old.secret), p.get('client_id'), resource, { accessHash: await sha256(access), refreshHash: await sha256(refresh), accessExpires: Date.now() + 3600000 });
      if (!r) return oauthError('invalid_grant');
      return json({ access_token: old.id + '.' + access, refresh_token: old.id + '.' + refresh, token_type: 'Bearer', expires_in: 3600, scope: 'document:read' + (r.write ? ' document:write' : '') });
    }
    return oauthError('unsupported_grant_type');
  }
  return null;
}
