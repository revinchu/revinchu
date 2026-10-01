// 오래 열린 배포 탭의 새 버전 안내. 문서나 저장소는 읽지 않으며 새로고침하지 않는다.
const RELEASE_ASSET = /^wixel-[a-f0-9]{16}\.js$/;
const FIRST_CHECK_MS = 15_000;
const CHECK_INTERVAL_MS = 5 * 60_000;

/** 외부 JS를 사용하는 동일 출처 배포본만 확인한다. 소스/단일 HTML 실행은 제외. */
export function releaseUpdateTarget(document, href) {
  try {
    const page = new URL(href);
    if (!/^https?:$/.test(page.protocol)) return null;
    for (const script of document.querySelectorAll('script[type="module"][src]')) {
      const url = new URL(script.getAttribute('src'), page);
      const asset = url.pathname.split('/').at(-1);
      if (url.origin !== page.origin || url.username || url.password || url.search || url.hash || !RELEASE_ASSET.test(asset)) continue;
      return { asset, url: new URL('version.json', url).href };
    }
  } catch { /* 잘못된 주소는 배포 확인 대상이 아님 */ }
  return null;
}

/** 최초 15초 뒤, 이후 창 focus 때 최대 5분에 한 번. 실패는 다음 focus에서 재시도한다. */
export function watchReleaseUpdate({
  onUpdate, document = globalThis.document, window = globalThis.window,
  fetch = globalThis.fetch, now = () => performance.now(),
  setTimeout = globalThis.setTimeout, clearTimeout = globalThis.clearTimeout,
} = {}) {
  const target = releaseUpdateTarget(document, window?.location?.href);
  if (!target || typeof fetch !== 'function') return () => {};
  const firstAt = now() + FIRST_CHECK_MS;
  let lastCheck = -Infinity, pending = false, stopped = false, notified = false, controller;
  const check = async () => {
    const time = now();
    if (stopped || notified || pending || time < firstAt || time - lastCheck < CHECK_INTERVAL_MS) return;
    lastCheck = time;
    pending = true;
    controller = new AbortController();
    const timeout = setTimeout(() => controller?.abort(), 10_000);
    try {
      const response = await fetch(target.url, {
        method: 'GET', cache: 'no-store', credentials: 'omit', redirect: 'error',
        referrerPolicy: 'no-referrer', signal: controller.signal,
      });
      if (!response.ok || stopped) return;
      const manifest = await response.json();
      if (stopped || !RELEASE_ASSET.test(manifest?.asset) || manifest.asset === target.asset) return;
      notified = true;
      onUpdate?.();
    } catch { /* 오프라인·이전 배포의 manifest 없음은 조용히 무시 */ }
    finally { clearTimeout(timeout); controller = null; pending = false; }
  };
  const firstTimer = setTimeout(check, FIRST_CHECK_MS);
  window.addEventListener('focus', check);
  return () => {
    stopped = true;
    clearTimeout(firstTimer);
    controller?.abort();
    window.removeEventListener('focus', check);
  };
}
