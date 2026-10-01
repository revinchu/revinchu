import test from 'node:test';
import assert from 'node:assert/strict';
import { DatabaseSync } from 'node:sqlite';
import { VaultStore } from './store.js';
import { JsonValidator } from './json-stream.js';
import { LIMITS, parseRevision, randomId, vaultKey, validateOrigin } from './shared.js';
import { publicUrl, fetchPublicText, boundedBytes } from './proxy.js';

function fixture(limits) {
  const db = new DatabaseSync(':memory:');
  const storage = { sql: { exec(query, ...bindings) {
    const statement = db.prepare(query), values = bindings.map(x => x instanceof ArrayBuffer ? new Uint8Array(x) : x);
    return { toArray() { return statement.all(...values); } };
  } }, transactionSync(fn) {
    db.exec('BEGIN'); try { const result = fn(); db.exec('COMMIT'); return result; } catch (e) { db.exec('ROLLBACK'); throw e; }
  } };
  // Workerd exec executes eagerly; node:sqlite statements require run/all.
  storage.sql.exec = (query, ...bindings) => {
    const statement = db.prepare(query), values = bindings.map(x => x instanceof ArrayBuffer ? new Uint8Array(x) : x);
    const rows = statement.columns().length ? statement.all(...values) : (statement.run(...values), []);
    return { toArray: () => rows };
  };
  return { store: new VaultStore(storage, limits), db, storage };
}
function body(text, step = 8192) {
  const bytes = new TextEncoder().encode(text); let offset = 0;
  return new ReadableStream({ pull(controller) {
    if (offset >= bytes.length) return controller.close();
    controller.enqueue(bytes.slice(offset, offset += step));
  } });
}
async function read(store, name) { return new Response(store.read(name).body).text(); }
test('streaming JSON grammar agrees with JSON.parse across UTF-8/chunk boundaries', () => {
  const valid = [null, true, false, 0, -3.2e44, '', 'a\\b"\n한글😀', [], {}, { a: [1, { b: false }, null], c: '\u0000' }];
  for (const item of valid) {
    const text = JSON.stringify(item), parser = new JsonValidator();
    for (const char of text) parser.write(char); parser.finish();
  }
  for (const text of ['', '[1,]', '{"a":1,}', '{"a" 1}', '{"a":}', 'true false', '01', '1.', '-.', 'nul', '["\n"]', '"\\x"', '{]', '[]x', '[truefalse]', '[1\u00a0]']) {
    const parser = new JsonValidator(); assert.throws(() => { for (const char of text) parser.write(char); parser.finish(); }, undefined, text);
  }
});
test('key is canonical 256-bit base64url and revision/origin fail closed', () => {
  const key = randomId(); assert.equal(key.length, 43);
  assert.equal(vaultKey(new Request('https://wixel.test', { headers: { 'X-Wixel-Vault': key } })), key);
  assert.throws(() => vaultKey(new Request('https://wixel.test')), { status: 401 });
  assert.throws(() => parseRevision(new Request('https://wixel.test')), { status: 428 });
  assert.equal(parseRevision(new Request('https://wixel.test', { headers: { 'If-Match': '"0"' } })), 0);
  assert.throws(() => parseRevision(new Request('https://wixel.test', { headers: { 'If-Match': '*' } })), { status: 400 });
  assert.throws(() => validateOrigin(new Request('https://wixel.test', { headers: { Origin: 'https://evil.test' } })), { status: 403 });
});
test('SQLite multi-chunk Unicode round trip and stale update stays atomic', async () => {
  const { store, db } = fixture(LIMITS); try {
    const text = JSON.stringify({ value: '한글😀'.repeat(250000) });
    const saved = await store.put('테스트', 0, body(text, 65537));
    assert.ok(store.find('테스트').chunks >= 2); assert.equal(await read(store, '테스트'), text);
    await assert.rejects(store.put('테스트', 0, body('{}')), { status: 412 });
    assert.equal(store.find('테스트').revision, saved.revision);
    await assert.rejects(store.put('테스트', saved.revision, body('{"broken":')), { status: 400 });
    assert.equal(await read(store, '테스트'), text);
    assert.equal(db.prepare('SELECT count(*) AS n FROM staging').get().n, 0);
  } finally { db.close(); }
});
test('invalid UTF-8 and oversized upload leave no partially stored document', async () => {
  const { store, db } = fixture({ ...LIMITS, maxDocumentBytes: 32 }); try {
    await assert.rejects(store.put('bad', 0, new Response(new Uint8Array([34, 0xff, 34])).body), { status: 400 });
    await assert.rejects(store.put('big', 0, body(JSON.stringify('x'.repeat(100)))), { status: 413 });
    assert.equal(store.list().length, 0); assert.equal(db.prepare('SELECT count(*) AS n FROM staging').get().n, 0);
    assert.equal((await store.put('ok', 0, body('{}'))).revision, 1);
  } finally { db.close(); }
});
test('count, byte quota and replace quota use actual committed document sizes', async () => {
  const { store, db } = fixture({ ...LIMITS, maxDocuments: 2, maxVaultBytes: 10 }); try {
    const first = await store.put('a', 0, body('"1234"'));
    await store.put('b', 0, body('{}'));
    await assert.rejects(store.put('c', 0, body('{}')), { status: 413 });
    await assert.rejects(store.put('a', first.revision, body('"12345678"')), { status: 413 });
    assert.equal(await read(store, 'a'), '"1234"');
    await store.put('a', first.revision, body('{}'));
    assert.equal(store.usage().bytes, 4);
  } finally { db.close(); }
});
test('delete/recreate never reuses revision and an open reader keeps old snapshot', async () => {
  const { store, db } = fixture(LIMITS); try {
    const text = JSON.stringify('x'.repeat(2300000)), first = await store.put('a', 0, body(text));
    const open = store.read('a').body;
    store.remove('a', first.revision);
    const second = await store.put('a', 0, body('{"new":true}'));
    assert.ok(second.revision > first.revision);
    await assert.rejects(store.put('a', first.revision, body('{}')), { status: 412 });
    assert.equal(await new Response(open).text(), text);
    assert.equal(db.prepare('SELECT count(*) AS n FROM chunks WHERE revision=?').get(first.revision).n, 0);
  } finally { db.close(); }
});
test('concurrent upload is rejected and final CAS prevents delete during upload', async () => {
  const { store, db } = fixture(LIMITS); try {
    const first = await store.put('a', 0, body('{}'));
    let finish;
    const stream = new ReadableStream({ start(controller) { controller.enqueue(new TextEncoder().encode('{"next":')); finish = () => { controller.enqueue(new TextEncoder().encode('true}')); controller.close(); }; } });
    const pending = store.put('a', first.revision, stream);
    await assert.rejects(store.put('b', 0, body('{}')), { status: 429 });
    store.remove('a', first.revision); finish();
    await assert.rejects(pending, { status: 412 }); assert.equal(store.find('a'), null);
  } finally { db.close(); }
});
test('backup streams consistent documents when current data changes', async () => {
  const { store, db } = fixture(LIMITS); try {
    const first = await store.put('a', 0, body('{"old":true}')); await store.put('b', 0, body('[1,2,3]'));
    const backup = store.backup();
    await store.put('a', first.revision, body('{"new":true}'));
    const result = await new Response(backup).json();
    assert.equal(result.format, 'wixel-vault-backup'); assert.deepEqual(result.documents.find(x => x.name === 'a').data, { old: true });
    assert.deepEqual(result.documents.find(x => x.name === 'b').data, [1, 2, 3]);
  } finally { db.close(); }
});
test('proxy rejects IP, private schemes, nondefault ports, credentials and self host', () => {
  for (const value of ['http://google.com', 'https://127.1', 'https://[::1]', 'https://2130706433', 'https://name.local', 'https://metadata.google.internal', 'https://user:pass@google.com', 'https://google.com:8443', 'https://wixel.workers.dev']) {
    assert.throws(() => publicUrl(value, 'wixel.workers.dev'), undefined, value);
  }
  assert.equal(publicUrl('https://docs.google.com/spreadsheets/d/id/gviz/tq', 'wixel.workers.dev').protocol, 'https:');
});
test('proxy checks every redirect and never forwards request secrets', async () => {
  let count = 0;
  await assert.rejects(fetchPublicText('https://docs.google.com/a', 'wixel.workers.dev', async (url, options) => {
    count++; assert.equal(options.redirect, 'manual'); assert.equal(options.headers['X-Wixel-Vault'], undefined);
    return new Response(null, { status: 302, headers: { Location: 'https://127.0.0.1/private' } });
  }), { status: 403 }); assert.equal(count, 1);
  count = 0;
  await assert.rejects(fetchPublicText('https://docs.google.com/a', 'wixel.workers.dev', async () => {
    count++; return new Response(null, { status: 302, headers: { Location: '/again' } });
  }), { status: 502 }); assert.equal(count, 6);
});
test('proxy bounds actual streamed bytes even without a Content-Length', async () => {
  await assert.rejects(boundedBytes(new Response(body('x'.repeat(100), 17)), 32), { status: 413 });
  const response = await fetchPublicText('https://docs.google.com/a', 'wixel.workers.dev', async () =>
    new Response('<script>test</script>', { headers: { 'Content-Type': 'text/html', 'Set-Cookie': 'secret=yes' } }));
  assert.equal(response.headers.get('Content-Type'), 'text/plain; charset=utf-8');
  assert.equal(response.headers.get('Set-Cookie'), null); assert.equal(await response.text(), '<script>test</script>');
});

test('slow upload times out and releases the upload lock without replacing data', async () => {
  const { store, db } = fixture({ ...LIMITS, uploadTimeoutMs: 10 }); try {
    const saved = await store.put('a', 0, body('{}'));
    const slow = new ReadableStream({ start(controller) { controller.enqueue(new TextEncoder().encode('{"pending":')); } });
    await assert.rejects(store.put('a', saved.revision, slow), { status: 408 });
    assert.equal(await read(store, 'a'), '{}');
    assert.equal(db.prepare('SELECT count(*) AS n FROM staging').get().n, 0);
    assert.ok((await store.put('b', 0, body('{}'))).revision > saved.revision);
  } finally { db.close(); }
});
test('publication quota releases a slot after revoke', () => {
  const { store, db } = fixture({ ...LIMITS, maxPublications: 2 }); try {
    store.reservePublication('one'); store.reservePublication('two');
    assert.throws(() => store.reservePublication('three'), { status: 413 });
    store.releasePublication('one'); store.reservePublication('three');
    assert.equal(store.publications().length, 2);
  } finally { db.close(); }
});

test('history records prior content and restore creates a new CAS-protected revision', async () => {
  const { store, db } = fixture(LIMITS); try {
    const first = await store.put('a', 0, body('{"a":1}'));
    const second = await store.put('a', first.revision, body('{"a":2}'));
    assert.equal(store.versions('a').currentRevision, second.revision);
    assert.equal(store.versions('a').versions[0].revision, first.revision);
    assert.equal(await new Response(store.readVersion('a', first.revision).body).text(), '{"a":1}');
    assert.throws(() => store.restore('a', first.revision, first.revision), { status: 412 });
    const restored = store.restore('a', first.revision, second.revision);
    assert.ok(restored.revision > second.revision); assert.equal(await read(store, 'a'), '{"a":1}');
    assert.equal(store.versions('a').versions[0].revision, second.revision);
  } finally { db.close(); }
});

test('bounded history retains open streams through pruning and releases canceled readers', async () => {
  const { store, db } = fixture({ ...LIMITS, maxVersions: 1, maxHistoryBytes: 3 * 1024 * 1024 }); try {
    const text = JSON.stringify('x'.repeat(2300000));
    const a = await store.put('a', 0, body(text)), b = await store.put('a', a.revision, body('{"b":1}'));
    const old = store.readVersion('a', a.revision).body;
    const c = await store.put('a', b.revision, body('{"c":1}'));
    assert.deepEqual(store.versions('a').versions.map(x => x.revision), [b.revision]);
    assert.equal(await new Response(old).text(), text);
    assert.equal(db.prepare('SELECT COUNT(*) AS n FROM chunks WHERE revision=?').get(a.revision).n, 0);
    const canceled = store.readVersion('a', b.revision).body;
    await store.put('a', c.revision, body('{"d":1}')); await canceled.cancel();
    assert.equal(db.prepare('SELECT COUNT(*) AS n FROM chunks WHERE revision=?').get(b.revision).n, 0);
    assert.equal(store.readers.size, 0);
  } finally { db.close(); }
});

test('vault-wide history bytes are bounded without charging current-document quota', async () => {
  const { store, db } = fixture({ ...LIMITS, maxHistoryBytes: 8, maxVaultBytes: 20 }); try {
    const a = await store.put('a', 0, body('"1111"')); await store.put('a', a.revision, body('{}'));
    const b = await store.put('b', 0, body('"2222"')); await store.put('b', b.revision, body('{}'));
    assert.equal(store.usage().bytes, 4); assert.equal(store.usage().historyBytes, 6);
    assert.deepEqual(store.versions('a').versions, []); assert.equal(store.versions('b').versions.length, 1);
  } finally { db.close(); }
});

test('additive schema reopening preserves current documents and historical chunks', async () => {
  const { store, db, storage } = fixture(LIMITS); try {
    const first = await store.put('legacy', 0, body('{"legacy":true}'));
    db.exec('DROP TABLE versions; DROP TABLE imports; DROP TABLE import_items; DROP TABLE import_chunks');
    const migrated = new VaultStore(storage, LIMITS);
    assert.equal(await read(migrated, 'legacy'), '{"legacy":true}');
    await migrated.put('legacy', first.revision, body('{}'));
    const restarted = new VaultStore(storage, LIMITS);
    assert.equal(await new Response(restarted.readVersion('legacy', first.revision).body).text(), '{"legacy":true}');
  } finally { db.close(); }
});

test('deletion removes every history entry while an in-flight historical reader finishes', async () => {
  const { store, db } = fixture(LIMITS); try {
    const a = await store.put('a', 0, body('{"old":1}')), b = await store.put('a', a.revision, body('{}'));
    const open = store.readVersion('a', a.revision).body;
    store.remove('a', b.revision);
    assert.throws(() => store.readVersion('a', a.revision), { status: 404 });
    assert.equal(await new Response(open).text(), '{"old":1}');
    assert.equal(store.usage().historyBytes, 0); assert.equal(db.prepare('SELECT COUNT(*) AS n FROM chunks').get().n, 0);
  } finally { db.close(); }
});

function manifest(documents, mode = 'create') { return { format: 'wixel-vault-backup', version: 1, mode, documents }; }
test('import rejects malformed manifests, duplicate names, implicit overwrite and incomplete uploads', async () => {
  const { store, db } = fixture(LIMITS); try {
    const a = await store.put('a', 0, body('{"old":true}'));
    for (const bad of [null, {}, manifest([]), manifest([{ name: 'a', expectedRevision: a.revision }]), manifest([{ name: 'a', expectedRevision: -1 }]), manifest([{ name: 'b', expectedRevision: 0 }, { name: ' b ', expectedRevision: 0 }]), manifest([{ name: 'bad\u0001', expectedRevision: 0 }])]) assert.throws(() => store.beginImport(bad), { status: 400 });
    assert.throws(() => store.beginImport(manifest([{ name: 'a', expectedRevision: 0 }])), { status: 412 });
    const session = store.beginImport(manifest([{ name: 'a', expectedRevision: a.revision }, { name: 'b', expectedRevision: 0 }], 'replace'));
    assert.throws(() => store.beginImport(manifest([{ name: 'c', expectedRevision: 0 }])), { status: 409 });
    await assert.rejects(store.stageImport(session.id, 'a', a.revision, body('{"broken":')), { status: 400 });
    await store.stageImport(session.id, 'b', 0, body('{"b":1}'));
    assert.throws(() => store.commitImport(session.id), { status: 409 });
    assert.equal(await read(store, 'a'), '{"old":true}'); assert.equal(store.find('b'), null);
    store.cancelImport(session.id); assert.equal(db.prepare('SELECT COUNT(*) AS n FROM import_chunks').get().n, 0);
  } finally { db.close(); }
});

test('multi-document import commits atomically after restart, creates history and keeps unrelated data', async () => {
  const { store, db, storage } = fixture(LIMITS); try {
    const first = await store.put('a', 0, body('{"old":1}')); await store.put('untouched', 0, body('true'));
    const session = store.beginImport(manifest([{ name: 'a', expectedRevision: first.revision }, { name: 'b', expectedRevision: 0 }], 'replace'));
    await store.stageImport(session.id, 'a', first.revision, body('{"new":2}'));
    await store.stageImport(session.id, 'b', 0, body('[1,2,3]'));
    const restarted = new VaultStore(storage, LIMITS), result = restarted.commitImport(session.id);
    assert.equal(result.documents.length, 2); assert.equal(await read(restarted, 'a'), '{"new":2}'); assert.equal(await read(restarted, 'b'), '[1,2,3]'); assert.equal(await read(restarted, 'untouched'), 'true');
    assert.equal(restarted.versions('a').versions[0].revision, first.revision);
    assert.equal(db.prepare('SELECT COUNT(*) AS n FROM import_chunks').get().n, 0);
  } finally { db.close(); }
});

test('import final CAS checks every document and leaves no partial replacements', async () => {
  const { store, db } = fixture(LIMITS); try {
    const a = await store.put('a', 0, body('1')), b = await store.put('b', 0, body('2'));
    const session = store.beginImport(manifest([{ name: 'a', expectedRevision: a.revision }, { name: 'b', expectedRevision: b.revision }], 'replace'));
    await store.stageImport(session.id, 'a', a.revision, body('11')); await store.stageImport(session.id, 'b', b.revision, body('22'));
    await store.put('b', b.revision, body('3'));
    assert.throws(() => store.commitImport(session.id), { status: 412 });
    assert.equal(await read(store, 'a'), '1'); assert.equal(await read(store, 'b'), '3'); store.cancelImport(session.id);
  } finally { db.close(); }
});

test('import SQL failure rolls back documents, revisions, history and all chunk changes', async () => {
  const { store, db, storage } = fixture(LIMITS); try {
    const a = await store.put('a', 0, body('1')), b = await store.put('b', 0, body('2'));
    const session = store.beginImport(manifest([{ name: 'a', expectedRevision: a.revision }, { name: 'b', expectedRevision: b.revision }], 'replace'));
    await store.stageImport(session.id, 'a', a.revision, body('11')); await store.stageImport(session.id, 'b', b.revision, body('22'));
    const original = storage.sql.exec;
    storage.sql.exec = (sql, ...args) => { if (sql.startsWith('INSERT INTO documents') && args[0] === 'b') throw new Error('synthetic disk error'); return original(sql, ...args); };
    assert.throws(() => store.commitImport(session.id), /synthetic disk error/); storage.sql.exec = original;
    assert.equal(await read(store, 'a'), '1'); assert.equal(await read(store, 'b'), '2'); assert.equal(store.usage().historyBytes, 0); assert.equal(Number(store.meta('sequence')), b.revision);
    assert.equal(store.commitImport(session.id).documents.length, 2);
  } finally { db.close(); }
});

test('import quota, invalid UTF-8 and deep JSON fail without altering any current documents', async () => {
  const { store, db } = fixture({ ...LIMITS, maxVaultBytes: 10 }); try {
    await store.put('old', 0, body('"1234"'));
    const session = store.beginImport(manifest([{ name: 'a', expectedRevision: 0 }]));
    await assert.rejects(store.stageImport(session.id, 'a', 0, new Response(new Uint8Array([34, 255, 34])).body), { status: 400 });
    await assert.rejects(store.stageImport(session.id, 'a', 0, body('['.repeat(129) + '0' + ']'.repeat(129))), { status: 413 });
    await store.stageImport(session.id, 'a', 0, body('"1234"'));
    assert.throws(() => store.commitImport(session.id), { status: 413 }); assert.equal(store.find('a'), null); assert.equal(await read(store, 'old'), '"1234"');
    store.cancelImport(session.id);
  } finally { db.close(); }
});

test('expired imports reclaim staged chunks after restart and session ids are vault-local', async () => {
  const { store, db, storage } = fixture(LIMITS), other = fixture(LIMITS); try {
    const session = store.beginImport(manifest([{ name: '__proto__', expectedRevision: 0 }]));
    await store.stageImport(session.id, '__proto__', 0, body('{"__proto__":{"polluted":true}}'));
    assert.throws(() => other.store.commitImport(session.id), { status: 404 }); assert.equal({}.polluted, undefined);
    db.prepare('UPDATE imports SET expires=0').run(); const restarted = new VaultStore(storage, LIMITS);
    assert.throws(() => restarted.commitImport(session.id), { status: 404 }); assert.equal(db.prepare('SELECT COUNT(*) AS n FROM import_chunks').get().n, 0);
    assert.equal(restarted.list().length, 0);
  } finally { db.close(); other.db.close(); }
});
