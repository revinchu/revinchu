import test from 'node:test';
import assert from 'node:assert/strict';
import { createRecoveryMarker, parseRecoveryMarker, recoveryCheckpoint, markRecoveryPending,
  markRecoverySaved, markRecoveryFailed, canUpdateRecoveryMarker, selectRecoveryMarker, recoveryDecision } from '../src/document-recovery.js';

const marker = (overrides = {}) => createRecoveryMarker({ docId:'doc-b', docName:'새 문서', sessionId:'tab-1', openId:'open-b', openedAt:100, ...overrides });
const stored = (state, extra = {}) => ({ docId:state.docId, docName:state.docName, generation:'generation-b', recovery:recoveryCheckpoint(state), ...extra });

test('successful open journal is pending and contains no workbook payload', () => {
  const state = marker();
  assert.equal(state.phase, 'pending');
  assert.equal(state.generation, null);
  assert.equal(state.version, 0);
  assert.deepEqual(parseRecoveryMarker(JSON.stringify(state)), state);
  assert.deepEqual(Object.keys(state).sort(), ['v','docId','docName','sessionId','openId','version','openedAt','phase','generation'].sort());
});

test('invalid or future journals fail safely to legacy handling', () => {
  for (const input of [null,undefined,'{',[],{}, { ...marker(), v:2 }, { ...marker(), docId:'' },
    { ...marker(), version:-1 }, { ...marker(), version:1.5 }, { ...marker(), phase:'unknown' },
    { ...marker(), generation:{} }, { ...marker(), openedAt:NaN }]) assert.equal(parseRecoveryMarker(input), null);
  assert.throws(() => marker({ docId:'' }), TypeError);
  assert.equal(recoveryDecision('invalid', { docId:'legacy' }).action, 'restore');
});

test('opening a new document cannot silently restore a different previous document', () => {
  const latest = marker();
  assert.equal(recoveryDecision(latest, { docId:'doc-a', docName:'이전 문서' }).action, 'prompt');
  assert.equal(recoveryDecision(latest, { docId:'doc-a' }).reason, 'different-document');
  assert.equal(recoveryDecision(latest, null).reason, 'missing-snapshot');
});

test('identical file names do not authorize restoration of a different document', () => {
  const latest = marker();
  assert.equal(recoveryDecision(latest, { docId:'doc-a', docName:latest.docName }).action, 'prompt');
});

test('committed snapshot is recognized before journal saved-state update', () => {
  const latest = marker({ version:4 });
  assert.equal(recoveryDecision(latest, stored(latest)).reason, 'committed-checkpoint');
  const saved = markRecoverySaved(latest, recoveryCheckpoint(latest), 'generation-b');
  assert.equal(saved.phase, 'saved');
  assert.equal(recoveryDecision(saved, stored(latest)).action, 'restore');
});

test('new edits remain pending when an older save completes', () => {
  const first = marker({ version:4 }), checkpoint = recoveryCheckpoint(first);
  const latest = markRecoveryPending(first, 5);
  assert.deepEqual(markRecoverySaved(latest, checkpoint, 'old-generation'), latest);
  assert.equal(recoveryDecision(latest, stored(first)).reason, 'unsaved-changes');
  assert.equal(recoveryDecision(latest, stored(latest)).action, 'restore');
});

test('a same-file unsaved edit warns instead of presenting an older generation as latest', () => {
  const first = marker(), snapshot = stored(first);
  const saved = markRecoverySaved(first, recoveryCheckpoint(first), snapshot.generation);
  const latest = markRecoveryPending(saved, 1);
  assert.equal(recoveryDecision(saved, snapshot).action, 'restore');
  assert.equal(recoveryDecision(latest, snapshot).action, 'prompt');
  assert.equal(latest.generation, snapshot.generation);
});

test('a failed current save is recorded but late errors do not corrupt newer state', () => {
  const first = marker(), checkpoint = recoveryCheckpoint(first);
  assert.equal(markRecoveryFailed(first, checkpoint).phase, 'failed');
  const latest = markRecoveryPending(first, 2);
  assert.deepEqual(markRecoveryFailed(latest, checkpoint), latest);
  assert.deepEqual(markRecoveryFailed(first, { ...checkpoint, openId:'old-open' }), first);
});

test('IDB commit survives later pointer or journal write failure', () => {
  const latest = marker({ version:3 });
  const failed = markRecoveryFailed(latest, recoveryCheckpoint(latest));
  assert.equal(failed.phase, 'failed');
  assert.equal(recoveryDecision(failed, stored(latest)).action, 'restore');
});

test('a completed snapshot newer than a journal checkpoint is safe to restore', () => {
  const latest = marker({ version:2 }), newer = markRecoveryPending(latest, 3);
  assert.equal(recoveryDecision(latest, stored(newer)).action, 'restore');
});

test('generation identity recovers complete snapshots when optional checkpoint metadata is absent', () => {
  const pending = marker();
  const saved = markRecoverySaved(pending, recoveryCheckpoint(pending), 'same-generation');
  assert.equal(recoveryDecision(saved, { docId:saved.docId, generation:'same-generation' }).reason, 'committed-generation');
  assert.equal(recoveryDecision(saved, { docId:saved.docId, generation:'older-generation' }).action, 'prompt');
  assert.equal(recoveryDecision(saved, { docId:'different', generation:'same-generation' }).action, 'prompt');
  assert.equal(recoveryDecision(saved, { docId:saved.docId }).action, 'prompt');
});

test('small snapshots require checkpoints when they have no generation', () => {
  const pending = marker(), saved = markRecoverySaved(pending, recoveryCheckpoint(pending));
  assert.equal(recoveryDecision(saved, { docId:saved.docId }).action, 'prompt');
  assert.equal(recoveryDecision(saved, { docId:saved.docId, recovery:recoveryCheckpoint(saved) }).action, 'restore');
});

test('reopening the same document creates a distinct open identity', () => {
  const latest = marker({ openId:'second-open' }), previous = marker();
  assert.equal(recoveryDecision(latest, stored(previous)).action, 'prompt');
  assert.deepEqual(markRecoverySaved(latest, recoveryCheckpoint(previous), 'old'), latest);
});

test('another tab checkpoint cannot authorize same-document auto recovery', () => {
  const latest = marker(), other = marker({ sessionId:'tab-2' });
  assert.equal(recoveryDecision(latest, stored(other)).action, 'prompt');
  assert.deepEqual(markRecoverySaved(latest, recoveryCheckpoint(other), 'other'), latest);
});

test('per-tab marker wins when a background tab becomes the global last document', () => {
  const own = marker(), global = marker({ docId:'doc-c', sessionId:'tab-2', openId:'open-c' });
  assert.deepEqual(selectRecoveryMarker(JSON.stringify(own), JSON.stringify(global)), own);
  assert.deepEqual(selectRecoveryMarker(null, JSON.stringify(global)), global);
  assert.deepEqual(selectRecoveryMarker('broken', JSON.stringify(global)), global);
});

test('late background and previous-document saves cannot take global journal ownership', () => {
  const current = marker(), ownNext = markRecoveryPending(current, 2);
  assert.equal(canUpdateRecoveryMarker(current, ownNext), true);
  assert.equal(canUpdateRecoveryMarker(ownNext, current), false);
  for (const stale of [marker({ docId:'previous' }), marker({ sessionId:'tab-2' }), marker({ openId:'previous-open' })]) {
    assert.equal(canUpdateRecoveryMarker(current, stale), false);
    assert.equal(canUpdateRecoveryMarker(stale, current), false);
  }
  assert.equal(canUpdateRecoveryMarker(null, current), false);
});

test('legacy documents without a journal still load, absent storage opens blank', () => {
  assert.deepEqual(recoveryDecision(null, { docName:'legacy' }), { action:'restore', reason:'legacy-snapshot' });
  assert.deepEqual(recoveryDecision(null, null), { action:'empty', reason:'no-snapshot' });
});

test('protocol does not mutate journal, checkpoint, or snapshot inputs', () => {
  const state = Object.freeze(marker()), checkpoint = Object.freeze(recoveryCheckpoint(state));
  const snapshot = Object.freeze(stored(state));
  const latest = markRecoveryPending(state, 1);
  markRecoverySaved(state, checkpoint, 'g'); markRecoveryFailed(state, checkpoint);
  recoveryDecision(state, snapshot); selectRecoveryMarker(state, latest); canUpdateRecoveryMarker(state, latest);
  assert.equal(state.phase, 'pending'); assert.equal(state.version, 0); assert.equal(checkpoint.version, 0);
  assert.deepEqual(snapshot, stored(state));
});

test('version updates are monotonic and malformed checkpoints never grant restoration', () => {
  const latest = marker({ version:4 });
  for (const version of [-1,3,Infinity,4.5,undefined]) assert.deepEqual(markRecoveryPending(latest, version), latest);
  for (const version of [-1,Infinity,4.5,'4',null,undefined]) {
    assert.equal(recoveryDecision(latest, { docId:latest.docId, recovery:{ ...recoveryCheckpoint(latest), version } }).action, 'prompt');
  }
});
