import { test } from 'node:test';
import assert from 'node:assert/strict';
import { releaseUpdateTarget, watchReleaseUpdate } from '../src/release-update.js';

const oldAsset = 'wixel-0123456789abcdef.js', newAsset = 'wixel-fedcba9876543210.js';
const doc = (...sources) => ({ querySelectorAll: () => sources.map(src => ({ getAttribute: () => src })) });
const drain = () => new Promise(resolve => setImmediate(resolve));
function fixture(fetchImpl = async () => ({ ok: true, json: async () => ({ asset: oldAsset }) })) {
  const window = new EventTarget(); window.location = { href: 'https://wixel.test/app/' };
  let time = 0, id = 0, notices = 0;
  const timers = new Map(), calls = [];
  const dispose = watchReleaseUpdate({ window, document: doc(oldAsset), now: () => time,
    fetch: (...args) => { calls.push(args); return fetchImpl(...args); }, onUpdate: () => { notices++; },
    setTimeout: (fn, delay) => { const key = ++id; timers.set(key, { at: time + delay, fn }); return key; }, clearTimeout: key => timers.delete(key),
  });
  return { calls, dispose, get notices() { return notices; }, focus: async () => { window.dispatchEvent(new Event('focus')); await drain(); },
    async advance(ms) {
      time += ms;
      for (const [key, timer] of [...timers]) if (timer.at <= time) { timers.delete(key); timer.fn(); }
      await drain();
    },
  };
}
test('배포 확인은 동일 출처의 해시 모듈만: 소스·인라인·file·외부 URL 제외', () => {
  assert.deepEqual(releaseUpdateTarget(doc(oldAsset), 'https://wixel.test/app/'), { asset: oldAsset, url: 'https://wixel.test/app/version.json' });
  for (const src of ['src/app.js', 'https://foreign.test/' + oldAsset, oldAsset + '?x=1', 'wixel-untrusted.js']) assert.equal(releaseUpdateTarget(doc(src), 'https://wixel.test/'), null);
  assert.equal(releaseUpdateTarget(doc(), 'https://wixel.test/'), null);
  assert.equal(releaseUpdateTarget(doc(oldAsset), 'file:///D:/app/index.html'), null);
});
test('최초 15초와 focus 5분 간격: 요청은 문서·쿠키·referrer 없는 GET', async () => {
  const f = fixture(); await f.focus(); await f.advance(14999); assert.equal(f.calls.length, 0);
  await f.advance(1); assert.equal(f.calls.length, 1); await f.focus(); assert.equal(f.calls.length, 1);
  const [url, options] = f.calls[0]; assert.equal(url, 'https://wixel.test/app/version.json');
  assert.equal(options.method, 'GET'); assert.equal(options.credentials, 'omit'); assert.equal(options.cache, 'no-store');
  assert.equal(options.referrerPolicy, 'no-referrer'); assert.equal(options.redirect, 'error'); assert.equal(options.body, undefined);
  await f.advance(299999); await f.focus(); assert.equal(f.calls.length, 1);
  await f.advance(1); await f.focus(); assert.equal(f.calls.length, 2); assert.equal(f.notices, 0); f.dispose();
});
test('다른 유효 버전은 한 번만 안내하며 사용자가 결정할 때까지 아무 작업도 수행하지 않음', async () => {
  const f = fixture(async () => ({ ok: true, json: async () => ({ asset: newAsset }) }));
  await f.advance(15000); assert.equal(f.notices, 1);
  await f.advance(300000); await f.focus(); assert.equal(f.calls.length, 1); assert.equal(f.notices, 1); f.dispose();
});
test('잘못된 응답·HTTP 오류·오프라인은 안내 없이 다음 허용 focus에서 재시도', async () => {
  const replies = [() => { throw Error('offline'); }, () => ({ ok: false }), () => ({ ok: true, json: async () => ({ asset: 'https://elsewhere.test/a.js' }) }), () => ({ ok: true, json: async () => ({ asset: newAsset }) })];
  const f = fixture(async () => replies.shift()()); await f.advance(15000); assert.equal(f.notices, 0);
  for (let i = 0; i < 3; i++) { await f.advance(300000); await f.focus(); assert.equal(f.notices, i === 2 ? 1 : 0); }
  assert.equal(f.calls.length, 4); f.dispose();
});
test('중복 focus 요청은 합치고, 확인 종료 뒤 도착한 응답은 무시', async () => {
  let complete;
  const f = fixture(() => new Promise(resolve => { complete = resolve; }));
  await f.advance(15000); await f.advance(300000); await f.focus(); assert.equal(f.calls.length, 1);
  f.dispose(); complete({ ok: true, json: async () => ({ asset: newAsset }) }); await drain();
  assert.equal(f.notices, 0); assert.equal(f.calls[0][1].signal.aborted, true);
});
test('지연된 요청은 10초 후 중단하고 비배포본은 타이머조차 만들지 않음', async () => {
  const f = fixture((url, { signal }) => new Promise((resolve, reject) => signal.addEventListener('abort', () => reject(Error('timeout')))));
  await f.advance(15000); await f.advance(10000); assert.equal(f.calls[0][1].signal.aborted, true); assert.equal(f.notices, 0); f.dispose();
  let timers = 0;
  watchReleaseUpdate({ document: doc('src/app.js'), window: { location: { href: 'https://wixel.test/' } }, setTimeout: () => { timers++; } })();
  assert.equal(timers, 0);
});
