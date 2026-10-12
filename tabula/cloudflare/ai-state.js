import { sameHash } from './shared.js';

// One opaque object per client, pending consent, authorization code or grant.
// Atomic storage transactions keep code redemption and refresh rotation single-use.
export class AiStateStore {
  constructor(storage) { this.storage = storage; }
  async create(record) {
    return this.storage.transaction(async tx => {
      if (await tx.get('record')) return false;
      await tx.put('record', record); await tx.setAlarm(record.expires); return true;
    });
  }
  async read() {
    const record = await this.storage.get('record');
    return record && record.expires > Date.now() ? record : null;
  }
  async consume(expected) {
    return this.storage.transaction(async tx => {
      const record = await tx.get('record');
      if (!record || record.expires <= Date.now() || !Object.entries(expected).every(([key, value]) => key.endsWith('Hash') ? sameHash(record[key], value) : record[key] === value)) return null;
      await tx.delete('record'); return record;
    });
  }
  async authenticate(hash) {
    const r = await this.read();
    return r?.kind === 'grant' && r.accessExpires > Date.now() && sameHash(r.accessHash, hash) ? r : null;
  }
  async rotate(hash, client, resource, next) {
    return this.storage.transaction(async tx => {
      const r = await tx.get('record');
      if (!r || r.kind !== 'grant' || r.expires <= Date.now() || r.client !== client || r.resource !== resource || !sameHash(r.refreshHash, hash)) return null;
      const updated = { ...r, ...next }; await tx.put('record', updated); return updated;
    });
  }
  async revoke(owner, hash = null) {
    return this.storage.transaction(async tx => {
      const r = await tx.get('record');
      if (!r || r.kind !== 'grant' || (owner ? r.owner !== owner : !sameHash(r.accessHash, hash) && !sameHash(r.refreshHash, hash))) return false;
      await tx.delete('record'); return true;
    });
  }
}
