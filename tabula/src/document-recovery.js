/** A small journal of the visible document, separate from the last complete snapshot.
 * The journal contains identity/status only; it never replaces or deletes workbook data.
 */
const markerVersion = 1;
const textId = value => typeof value === 'string' && value.length > 0 && value.length <= 240;
const validVersion = value => Number.isSafeInteger(value) && value >= 0;
const sameOpen = (a, b) => !!a && !!b && a.docId === b.docId && a.sessionId === b.sessionId && a.openId === b.openId;

export function parseRecoveryMarker(raw) {
  try {
    const value = typeof raw === 'string' ? JSON.parse(raw) : raw;
    if (!value || value.v !== markerVersion || !textId(value.docId) || !textId(value.sessionId) || !textId(value.openId)
      || typeof value.docName !== 'string' || !validVersion(value.version) || !Number.isFinite(value.openedAt)
      || !['pending', 'saved', 'failed'].includes(value.phase)
      || (value.generation != null && !textId(value.generation))) return null;
    return { v: markerVersion, docId: value.docId, docName: value.docName.slice(0, 512), sessionId: value.sessionId,
      openId: value.openId, version: value.version, openedAt: value.openedAt, phase: value.phase,
      generation: value.generation ?? null };
  } catch { return null; }
}

/** Call only after a requested document has successfully become the visible workbook. */
export function createRecoveryMarker({ docId, docName, sessionId, openId, version = 0, openedAt = Date.now() }) {
  const marker = parseRecoveryMarker({ v: markerVersion, docId, docName, sessionId, openId, version, openedAt, phase: 'pending', generation: null });
  if (!marker) throw new TypeError('복구 문서 식별 정보가 올바르지 않습니다.');
  return marker;
}

/** Persist this with the complete snapshot, using its actual captured workbook version. */
export function recoveryCheckpoint(marker) {
  const value = parseRecoveryMarker(marker);
  return value ? { docId: value.docId, sessionId: value.sessionId, openId: value.openId, version: value.version } : null;
}

export function markRecoveryPending(marker, version) {
  const value = parseRecoveryMarker(marker);
  if (!value || !validVersion(version) || version < value.version) return value;
  return { ...value, version, phase: 'pending' };
}

export function markRecoverySaved(marker, checkpoint, generation = null) {
  const value = parseRecoveryMarker(marker);
  if (!value || !sameOpen(value, checkpoint) || checkpoint.version !== value.version
    || (generation != null && !textId(generation))) return value;
  return { ...value, phase: 'saved', generation };
}

export function markRecoveryFailed(marker, checkpoint) {
  const value = parseRecoveryMarker(marker);
  if (!value || !sameOpen(value, checkpoint) || checkpoint.version !== value.version) return value;
  return { ...value, phase: 'failed' };
}

/** Other tabs and late saves cannot take ownership of a newer document's journal.
 * A successful explicit open may replace ownership; ordinary pending/save updates may not.
 */
export function canUpdateRecoveryMarker(current, candidate) {
  const before = parseRecoveryMarker(current), after = parseRecoveryMarker(candidate);
  return sameOpen(before, after) && after.version >= before.version;
}

/** Per-tab identity wins over another tab's most recent global document. */
export function selectRecoveryMarker(sessionMarker, globalMarker) {
  return parseRecoveryMarker(sessionMarker) ?? parseRecoveryMarker(globalMarker);
}

/** Decide before loading a snapshot. A matching name alone is never evidence.
 * stored is the descriptor of the actual complete payload/IDB manifest, not its old pointer.
 * Persisted checkpoints allow recovery if the app stopped after commit but before journal update.
 */
export function recoveryDecision(marker, stored) {
  const latest = parseRecoveryMarker(marker);
  if (!latest) return { action: stored ? 'restore' : 'empty', reason: stored ? 'legacy-snapshot' : 'no-snapshot' };
  if (!stored) return { action: 'prompt', reason: 'missing-snapshot', latest };
  if (stored.docId !== latest.docId) return { action: 'prompt', reason: 'different-document', latest };
  const checkpoint = stored.recovery;
  if (sameOpen(latest, checkpoint) && validVersion(checkpoint.version) && checkpoint.version >= latest.version) {
    return { action: 'restore', reason: 'committed-checkpoint', latest };
  }
  if (latest.phase === 'saved' && latest.generation != null && latest.generation === stored.generation) {
    return { action: 'restore', reason: 'committed-generation', latest };
  }
  return { action: 'prompt', reason: latest.phase === 'saved' ? 'unverified-snapshot' : 'unsaved-changes', latest };
}
