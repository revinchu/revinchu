import { ApiError, LIMITS, randomId, validateName } from './shared.js';
import { JsonValidator } from './json-stream.js';

const CHUNK = 1024 * 1024;
const encoder = new TextEncoder();

export class VaultStore {
  constructor(storage, limits = LIMITS) {
    this.storage = storage; this.sql = storage.sql; this.limits = { ...LIMITS, ...limits }; this.uploading = false; this.readers = new Map();
    this.sql.exec('CREATE TABLE IF NOT EXISTS meta (key TEXT PRIMARY KEY, value TEXT NOT NULL)');
    this.sql.exec('CREATE TABLE IF NOT EXISTS documents (name TEXT PRIMARY KEY, revision INTEGER NOT NULL, modified INTEGER NOT NULL, size INTEGER NOT NULL, chunks INTEGER NOT NULL)');
    this.sql.exec('CREATE TABLE IF NOT EXISTS chunks (name TEXT NOT NULL, revision INTEGER NOT NULL, n INTEGER NOT NULL, data BLOB NOT NULL, PRIMARY KEY(name,revision,n))');
    this.sql.exec('CREATE TABLE IF NOT EXISTS staging (upload TEXT NOT NULL, n INTEGER NOT NULL, data BLOB NOT NULL, PRIMARY KEY(upload,n))');
    this.sql.exec('CREATE TABLE IF NOT EXISTS publications (id TEXT PRIMARY KEY, modified INTEGER NOT NULL)');
    // Additive schema migration: existing current documents and their revision numbers stay intact.
    this.sql.exec('CREATE TABLE IF NOT EXISTS versions (name TEXT NOT NULL, revision INTEGER NOT NULL, modified INTEGER NOT NULL, size INTEGER NOT NULL, chunks INTEGER NOT NULL, PRIMARY KEY(name,revision))');
    this.sql.exec('CREATE TABLE IF NOT EXISTS imports (id TEXT PRIMARY KEY, expires INTEGER NOT NULL)');
    this.sql.exec('CREATE TABLE IF NOT EXISTS import_items (id TEXT NOT NULL, name TEXT NOT NULL, expected INTEGER NOT NULL, size INTEGER, chunks INTEGER, PRIMARY KEY(id,name))');
    this.sql.exec('CREATE TABLE IF NOT EXISTS import_chunks (id TEXT NOT NULL, name TEXT NOT NULL, n INTEGER NOT NULL, data BLOB NOT NULL, PRIMARY KEY(id,name,n))');
    // 이전 인스턴스에서 완료하지 못한 업로드/이미 교체된 문서만 정리합니다.
    this.sql.exec('DELETE FROM staging');
    this.sql.exec('DELETE FROM chunks WHERE NOT EXISTS (SELECT 1 FROM documents d WHERE d.name=chunks.name AND d.revision=chunks.revision) AND NOT EXISTS (SELECT 1 FROM versions v WHERE v.name=chunks.name AND v.revision=chunks.revision)');
    this.expireImports();
  }
  meta(key) { return this.sql.exec('SELECT value FROM meta WHERE key=?', key).toArray()[0]?.value ?? null; }
  setMeta(key, value) { this.sql.exec('INSERT INTO meta(key,value) VALUES(?,?) ON CONFLICT(key) DO UPDATE SET value=excluded.value', key, String(value)); }
  list() { return this.sql.exec('SELECT name,revision,modified,size,chunks FROM documents ORDER BY modified DESC,name').toArray(); }
  find(name) { return this.sql.exec('SELECT name,revision,modified,size,chunks FROM documents WHERE name=?', name).toArray()[0] ?? null; }
  usage() {
    const row = this.sql.exec('SELECT COUNT(*) AS documents,COALESCE(SUM(size),0) AS bytes FROM documents').toArray()[0];
    return { ...row, historyBytes: this.sql.exec('SELECT COALESCE(SUM(size),0) AS n FROM versions').toArray()[0].n, publications: this.sql.exec('SELECT COUNT(*) AS n FROM publications').toArray()[0].n, ...this.limits };
  }
  assertRevision(name, expected) {
    const current = this.find(name), revision = current?.revision ?? 0;
    if (revision !== expected) throw new ApiError(412, '다른 기기에서 문서가 변경되었습니다. 최신 문서를 확인한 뒤 다시 저장하세요.', 'REVISION_CONFLICT', { currentRevision: revision });
    return current;
  }
  key(name, revision) { return JSON.stringify([name, revision]); }
  retain(meta) {
    const key = this.key(meta.name, meta.revision); this.readers.set(key, (this.readers.get(key) ?? 0) + 1);
  }
  release(meta) {
    const key = this.key(meta.name, meta.revision), count = (this.readers.get(key) ?? 1) - 1;
    if (count > 0) this.readers.set(key, count); else { this.readers.delete(key); this.clean(meta.name, meta.revision); }
  }
  clean(name, revision) {
    if (!this.readers.has(this.key(name, revision)) && this.find(name)?.revision !== revision && !this.version(name, revision)) this.sql.exec('DELETE FROM chunks WHERE name=? AND revision=?', name, revision);
  }
  version(name, revision) { return this.sql.exec('SELECT name,revision,modified,size,chunks FROM versions WHERE name=? AND revision=?', name, revision).toArray()[0] ?? null; }
  versions(name) {
    name = validateName(name); const current = this.find(name);
    if (!current) throw new ApiError(404, '문서가 없습니다.', 'NOT_FOUND');
    return { currentRevision: current.revision, versions: this.sql.exec('SELECT revision,modified,size FROM versions WHERE name=? ORDER BY revision DESC', name).toArray(), maxVersions: this.limits.maxVersions, maxHistoryBytes: this.limits.maxHistoryBytes };
  }
  readVersion(name, revision) {
    name = validateName(name); const meta = this.version(name, revision);
    if (!meta) throw new ApiError(404, '보관된 버전이 없습니다.', 'VERSION_NOT_FOUND');
    return { meta, body: this.stream(meta) };
  }
  archive(meta) {
    if (meta && this.limits.maxVersions > 0 && this.limits.maxHistoryBytes > 0) this.sql.exec('INSERT INTO versions(name,revision,modified,size,chunks) VALUES(?,?,?,?,?)', meta.name, meta.revision, meta.modified, meta.size, meta.chunks);
  }
  pruneVersions() {
    const entries = this.sql.exec('SELECT name,revision,size FROM versions ORDER BY revision DESC').toArray(), counts = new Map();
    let bytes = 0;
    // Newest revisions win both the per-document count and the vault-wide byte budget.
    for (const item of entries) {
      const count = counts.get(item.name) ?? 0;
      if (count < this.limits.maxVersions && bytes + item.size <= this.limits.maxHistoryBytes) { counts.set(item.name, count + 1); bytes += item.size; }
      else { this.sql.exec('DELETE FROM versions WHERE name=? AND revision=?', item.name, item.revision); this.clean(item.name, item.revision); }
    }
  }
  nextRevision() {
    const revision = Number(this.meta('sequence') ?? 0) + 1;
    if (!Number.isSafeInteger(revision)) throw new ApiError(507, '문서 버전 한도에 도달했습니다.', 'REVISION_EXHAUSTED');
    this.setMeta('sequence', revision); return revision;
  }
  install(name, expected, size, chunks, copy) {
    const old = this.assertRevision(name, expected), revision = this.nextRevision(), modified = Date.now();
    this.archive(old); copy(revision);
    this.sql.exec('INSERT INTO documents(name,revision,modified,size,chunks) VALUES(?,?,?,?,?) ON CONFLICT(name) DO UPDATE SET revision=excluded.revision,modified=excluded.modified,size=excluded.size,chunks=excluded.chunks', name, revision, modified, size, chunks);
    if (old) this.clean(name, old.revision);
    return { name, revision, modified, size };
  }
  restore(name, revision, expected) {
    name = validateName(name);
    return this.storage.transactionSync(() => {
      const current = this.assertRevision(name, expected), source = this.version(name, revision);
      if (!current || !source) throw new ApiError(404, '복원할 문서 버전이 없습니다.', 'VERSION_NOT_FOUND');
      if (this.usage().bytes - current.size + source.size > this.limits.maxVaultBytes) throw new ApiError(413, '보관함의 저장 용량을 초과했습니다.', 'VAULT_QUOTA');
      const result = this.install(name, expected, source.size, source.chunks, (next) => this.sql.exec('INSERT INTO chunks(name,revision,n,data) SELECT name,?,n,data FROM chunks WHERE name=? AND revision=?', next, name, revision));
      this.pruneVersions(); return { ok: true, ...result };
    });
  }
  chunk(meta, n) {
    const row = this.sql.exec('SELECT data FROM chunks WHERE name=? AND revision=? AND n=?', meta.name, meta.revision, n).toArray()[0];
    if (!row) throw new ApiError(500, '문서 조각을 읽지 못했습니다. 백업을 확인하세요.', 'STORAGE_INCOMPLETE');
    return new Uint8Array(row.data);
  }
  stream(meta) {
    this.retain(meta);
    let n = 0, closed = false;
    const finish = () => { if (!closed) { closed = true; this.release(meta); } };
    return new ReadableStream({
      pull: (controller) => {
        try {
          if (n < meta.chunks) controller.enqueue(this.chunk(meta, n++));
          else { controller.close(); finish(); }
        } catch (error) { controller.error(error); finish(); }
      },
      cancel: finish,
    });
  }
  read(name) {
    const meta = this.find(validateName(name));
    if (!meta) throw new ApiError(404, '문서가 없습니다.', 'NOT_FOUND');
    return { meta, body: this.stream(meta) };
  }
  async put(name, expected, body) {
    name = validateName(name); this.assertRevision(name, expected);
    const initial = this.usage();
    if (!this.find(name) && initial.documents >= this.limits.maxDocuments) throw new ApiError(413, '보관함의 문서 개수 한도를 초과했습니다.', 'VAULT_QUOTA');
    return this.stage(body, (upload, size, n) => {
      const old = this.assertRevision(name, expected), usage = this.usage();
      if ((!old && usage.documents >= this.limits.maxDocuments) || usage.bytes - (old?.size ?? 0) + size > this.limits.maxVaultBytes) throw new ApiError(413, '보관함의 저장 용량을 초과했습니다.', 'VAULT_QUOTA');
      const result = this.install(name, expected, size, n, (revision) => this.sql.exec('INSERT INTO chunks(name,revision,n,data) SELECT ?,?,n,data FROM staging WHERE upload=?', name, revision, upload));
      this.pruneVersions(); return { ok: true, ...result };
    });
  }
  async stage(body, commit) {
    if (!body) throw new ApiError(400, '저장할 문서가 없습니다.', 'EMPTY_BODY');
    if (this.uploading) throw new ApiError(429, '이 보관함에서 다른 문서를 저장 중입니다. 잠시 후 다시 저장하세요.', 'UPLOAD_BUSY');
    const upload = crypto.randomUUID(), reader = body.getReader(), parser = new JsonValidator(), decoder = new TextDecoder('utf-8', { fatal: true });
    this.uploading = true;
    let timedOut = false;
    const timer = setTimeout(() => { timedOut = true; reader.cancel().catch(() => {}); }, this.limits.uploadTimeoutMs ?? 30000);
    let size = 0, n = 0, used = 0, buffer = new Uint8Array(CHUNK), complete = false;
    const flush = () => {
      if (!used) return;
      this.sql.exec('INSERT INTO staging(upload,n,data) VALUES(?,?,?)', upload, n++, buffer.slice(0, used).buffer);
      used = 0;
    };
    try {
      for (;;) {
        const item = await reader.read();
        if (timedOut) throw new ApiError(408, '문서 전송 제한 시간 30초를 초과했습니다.', 'UPLOAD_TIMEOUT');
        if (item.done) break;
        const value = item.value;
        if (!(value instanceof Uint8Array)) throw new ApiError(400, '문서 전송 형식이 올바르지 않습니다.', 'INVALID_BODY');
        size += value.byteLength;
        if (size > this.limits.maxDocumentBytes) throw new ApiError(413, '문서 하나는 최대 20MiB까지 저장할 수 있습니다.', 'DOCUMENT_TOO_LARGE');
        try { parser.write(decoder.decode(value, { stream: true })); } catch (error) {
          if (error instanceof ApiError) throw error; throw new ApiError(400, 'UTF-8 JSON 문서만 저장할 수 있습니다.', 'INVALID_JSON');
        }
        for (let offset = 0; offset < value.length;) {
          const take = Math.min(CHUNK - used, value.length - offset);
          buffer.set(value.subarray(offset, offset + take), used); used += take; offset += take;
          if (used === CHUNK) flush();
        }
      }
      try { parser.write(decoder.decode()); parser.finish(); } catch (error) {
        if (error instanceof ApiError) throw error; throw new ApiError(400, 'UTF-8 JSON 문서만 저장할 수 있습니다.', 'INVALID_JSON');
      }
      flush();
      const result = this.storage.transactionSync(() => {
        const result = commit(upload, size, n);
        this.sql.exec('DELETE FROM staging WHERE upload=?', upload);
        return result;
      });
      complete = true; return result;
    } catch (error) {
      if (timedOut) throw new ApiError(408, '문서 전송 제한 시간을 초과했습니다.', 'UPLOAD_TIMEOUT');
      throw error;
    } finally {
      if (!complete) { await reader.cancel().catch(() => {}); this.sql.exec('DELETE FROM staging WHERE upload=?', upload); }
      clearTimeout(timer); reader.releaseLock(); this.uploading = false;
    }
  }
  remove(name, expected) {
    name = validateName(name);
    return this.storage.transactionSync(() => {
      const old = this.assertRevision(name, expected);
      if (!old) throw new ApiError(404, '문서가 없습니다.', 'NOT_FOUND');
      const versions = this.sql.exec('SELECT revision FROM versions WHERE name=?', name).toArray();
      this.sql.exec('DELETE FROM versions WHERE name=?', name);
      this.sql.exec('DELETE FROM documents WHERE name=?', name); this.clean(name, old.revision);
      for (const v of versions) this.clean(name, v.revision);
      return { ok: true };
    });
  }
  dropImport(id) {
    this.sql.exec('DELETE FROM import_chunks WHERE id=?', id);
    this.sql.exec('DELETE FROM import_items WHERE id=?', id);
    this.sql.exec('DELETE FROM imports WHERE id=?', id);
  }
  expireImports() {
    for (const row of this.sql.exec('SELECT id FROM imports WHERE expires<=?', Date.now()).toArray()) this.dropImport(row.id);
  }
  requireImport(id) {
    this.expireImports();
    const session = this.sql.exec('SELECT id,expires FROM imports WHERE id=?', id).toArray()[0];
    if (!session) throw new ApiError(404, '백업 복원 준비가 없거나 만료되었습니다. 다시 시작하세요.', 'IMPORT_NOT_FOUND');
    return session;
  }
  beginImport(manifest) {
    this.expireImports();
    if (this.uploading || this.sql.exec('SELECT id FROM imports').toArray().length) throw new ApiError(409, '이미 백업 복원을 준비 중입니다. 취소하거나 30분 후 다시 시도하세요.', 'IMPORT_BUSY');
    if (!manifest || manifest.format !== 'wixel-vault-backup' || manifest.version !== 1 || !['create', 'replace'].includes(manifest.mode) || !Array.isArray(manifest.documents) || !manifest.documents.length || manifest.documents.length > this.limits.maxDocuments) throw new ApiError(400, 'WIXEL 보관함 백업 형식과 문서 목록을 확인하세요.', 'INVALID_BACKUP');
    const seen = new Set(), items = [];
    for (const item of manifest.documents) {
      const name = validateName(item?.name), expected = item?.expectedRevision;
      if (seen.has(name) || !Number.isSafeInteger(expected) || expected < 0 || (manifest.mode === 'create' && expected !== 0)) throw new ApiError(400, '중복 이름이나 잘못된 복원 버전이 있습니다.', 'INVALID_BACKUP');
      seen.add(name); this.assertRevision(name, expected); items.push({ name, expected });
    }
    if (this.usage().documents + items.filter((x) => !this.find(x.name)).length > this.limits.maxDocuments) throw new ApiError(413, '보관함의 문서 개수 한도를 초과했습니다.', 'VAULT_QUOTA');
    const id = randomId(), expires = Date.now() + this.limits.importTimeoutMs;
    return this.storage.transactionSync(() => {
      this.sql.exec('INSERT INTO imports(id,expires) VALUES(?,?)', id, expires);
      for (const item of items) this.sql.exec('INSERT INTO import_items(id,name,expected) VALUES(?,?,?)', id, item.name, item.expected);
      return { id, expires, documents: items.length };
    });
  }
  async stageImport(id, name, expected, body) {
    name = validateName(name); this.requireImport(id);
    const item = this.sql.exec('SELECT expected FROM import_items WHERE id=? AND name=?', id, name).toArray()[0];
    if (!item || item.expected !== expected) throw new ApiError(400, '복원 준비 목록의 이름과 버전이 일치하지 않습니다.', 'IMPORT_MANIFEST_MISMATCH');
    this.assertRevision(name, expected);
    return this.stage(body, (upload, size, chunks) => {
      this.requireImport(id); this.assertRevision(name, expected);
      const staged = this.sql.exec('SELECT COALESCE(SUM(size),0) AS bytes FROM import_items WHERE id=? AND name<>?', id, name).toArray()[0].bytes;
      if (staged + size > this.limits.maxVaultBytes) throw new ApiError(413, '백업 복원 데이터의 용량 한도를 초과했습니다.', 'VAULT_QUOTA');
      this.sql.exec('DELETE FROM import_chunks WHERE id=? AND name=?', id, name);
      this.sql.exec('INSERT INTO import_chunks(id,name,n,data) SELECT ?,?,n,data FROM staging WHERE upload=?', id, name, upload);
      this.sql.exec('UPDATE import_items SET size=?,chunks=? WHERE id=? AND name=?', size, chunks, id, name);
      return { ok: true, name, size };
    });
  }
  commitImport(id) {
    this.requireImport(id);
    if (this.uploading) throw new ApiError(429, '문서 전송이 끝난 뒤 복원하세요.', 'UPLOAD_BUSY');
    return this.storage.transactionSync(() => {
      const items = this.sql.exec('SELECT name,expected,size,chunks FROM import_items WHERE id=? ORDER BY name', id).toArray();
      if (!items.length || items.some((x) => x.size === null || !x.chunks)) throw new ApiError(409, '모든 백업 문서의 전송과 검증이 끝나지 않았습니다.', 'IMPORT_INCOMPLETE');
      const usage = this.usage(); let bytes = usage.bytes, count = usage.documents;
      // Check every CAS and quota before changing even the first document.
      for (const item of items) { const old = this.assertRevision(item.name, item.expected); bytes += item.size - (old?.size ?? 0); if (!old) count++; }
      if (bytes > this.limits.maxVaultBytes || count > this.limits.maxDocuments) throw new ApiError(413, '복원 후 보관함 용량 또는 문서 개수 한도를 초과합니다.', 'VAULT_QUOTA');
      const documents = items.map((item) => this.install(item.name, item.expected, item.size, item.chunks, (revision) => this.sql.exec('INSERT INTO chunks(name,revision,n,data) SELECT ?,?,n,data FROM import_chunks WHERE id=? AND name=?', item.name, revision, id, item.name)));
      this.pruneVersions(); this.dropImport(id);
      return { ok: true, documents };
    });
  }
  cancelImport(id) { this.requireImport(id); return this.storage.transactionSync(() => { this.dropImport(id); return { ok: true }; }); }
  reservePublication(id) {
    if (this.usage().publications >= this.limits.maxPublications) throw new ApiError(413, '읽기 전용 게시 링크는 최대 20개까지 만들 수 있습니다.', 'PUBLISH_QUOTA');
    this.sql.exec('INSERT INTO publications(id,modified) VALUES(?,?)', id, Date.now()); return { ok: true };
  }
  releasePublication(id) { this.sql.exec('DELETE FROM publications WHERE id=?', id); return { ok: true }; }
  publications() { return this.sql.exec('SELECT id,modified FROM publications ORDER BY modified DESC').toArray(); }
  backup() {
    const docs = this.list(); for (const meta of docs) this.retain(meta);
    let closed = false;
    const finish = () => { if (!closed) { closed = true; for (const meta of docs) this.release(meta); } };
    const self = this;
    function* pieces() {
      yield encoder.encode('{"format":"wixel-vault-backup","version":1,"exportedAt":' + JSON.stringify(new Date().toISOString()) + ',"documents":[');
      for (let i = 0; i < docs.length; i++) {
        const meta = docs[i];
        yield encoder.encode((i ? ',' : '') + '{"name":' + JSON.stringify(meta.name) + ',"revision":' + meta.revision + ',"modified":' + meta.modified + ',"data":');
        for (let n = 0; n < meta.chunks; n++) yield self.chunk(meta, n);
        yield encoder.encode('}');
      }
      yield encoder.encode(']}');
    }
    const iterator = pieces();
    return new ReadableStream({
      pull(controller) {
        try { const item = iterator.next(); if (item.done) { controller.close(); finish(); } else controller.enqueue(item.value); }
        catch (error) { controller.error(error); finish(); }
      },
      cancel() { iterator.return(); finish(); },
    });
  }
}
