import { test } from 'node:test';
import assert from 'node:assert/strict';
import { createServer, request as httpRequest } from 'node:http';
import { Readable } from 'node:stream';
import { EventEmitter } from 'node:events';
import { gzipSync } from 'node:zlib';
import { isPublicAddress, isLoopbackAddress, parsePublicUrl, resolvePublicTarget, requestPinned, fetchPublicResource } from '../server/network.js';

const publicLookup = async () => [{ address: '93.184.216.34', family: 4 }];
function response(status = 200, chunks = ['정상 응답'], headers = {}) {
  const stream = Readable.from(chunks.map((chunk) => Buffer.from(chunk)));
  stream.statusCode = status; stream.headers = headers; return stream;
}
const statusIs = (status) => (error) => error.status === status;

test('외부 중계: 사설·메타데이터·특수 IPv4/IPv6와 전환 주소를 거부한다', () => {
  for (const address of [
    '0.0.0.0', '0.1.2.3', '10.1.2.3', '100.64.0.1', '100.127.255.254', '100.100.100.200',
    '127.0.0.1', '127.44.55.66', '168.63.129.16', '169.254.169.254', '172.16.0.1', '172.31.255.255', '192.168.1.1',
    '192.0.0.192', '192.0.2.1', '192.88.99.1', '198.18.1.1', '198.19.1.1', '198.51.100.1',
    '203.0.113.1', '224.0.0.1', '239.1.2.3', '240.0.0.1', '255.255.255.255',
    '::', '::1', '0:0:0:0:0:0:0:1', '::ffff:127.0.0.1', '::ffff:7f00:1', '::ffff:8.8.8.8',
    '64:ff9b::a00:1', '64:ff9b:1::a00:1', '2001::1', '2001:2::1', '2001:db8::1',
    '2002:7f00:1::', '3fff::1', 'fc00::1', 'fd00:ec2::254', 'fe80::1', 'febf::1', 'ff02::1',
    'fe80::1%eth0', 'not-an-ip', '',
  ]) assert.equal(isPublicAddress(address), false, address);
  for (const address of ['1.1.1.1', '8.8.8.8', '93.184.216.34', '100.63.255.255', '172.15.0.1', '172.32.0.1', '2001:4860:4860::8888', '2606:4700:4700::1111']) {
    assert.equal(isPublicAddress(address), true, address);
  }
});

test('외부 중계: URL 정규화 뒤에도 내부 IP 표현·사용자 정보·포트 우회를 거부한다', () => {
  for (const url of [
    'http://2130706433/', 'http://0x7f000001/', 'http://0177.0.0.1/', 'http://127.1/',
    'http://%31%32%37.0.0.1/', 'http://[::ffff:127.0.0.1]/', 'http://localhost./',
    'http://a.localhost/', 'http://metadata.google.internal/', 'http://router.home.arpa/',
    'file:///etc/passwd', 'ftp://example.com/', 'http://u:p@example.com/',
    'http://example.com:8080/', 'https://example.com:8443/', 'http://example.com:0/', 'bad-url',
  ]) assert.throws(() => parsePublicUrl(url), statusIs(400), url);
  for (const url of ['http://example.com/', 'https://example.com/', 'http://example.com:443/', 'https://example.com:80/']) {
    assert.ok(parsePublicUrl(url) instanceof URL);
  }
});

test('외부 중계: DNS의 모든 주소와 family를 검사하고 IP 리터럴은 재해석하지 않는다', async () => {
  let calls = 0;
  const resolved = await resolvePublicTarget('https://public.example/path', async (host, options) => {
    calls++; assert.equal(host, 'public.example'); assert.deepEqual(options, { all: true, verbatim: true });
    return [{ address: '2606:4700:4700::1111', family: 6 }, { address: '1.1.1.1', family: 4 }];
  });
  assert.equal(calls, 1); assert.equal(resolved.address, '1.1.1.1');
  for (const addresses of [[], [{ address: '1.1.1.1', family: 4 }, { address: '10.0.0.1', family: 4 }],
    [{ address: '8.8.8.8', family: 6 }], [{ address: '::ffff:a00:1', family: 6 }]]) {
    await assert.rejects(resolvePublicTarget('http://public.example/', async () => addresses), statusIs(400));
  }
  await resolvePublicTarget('https://1.1.1.1/', () => { throw new Error('IP를 다시 조회하면 안 됨'); });
  assert.equal(isLoopbackAddress('127.3.2.1'), true);
  assert.equal(isLoopbackAddress('::ffff:127.0.0.1'), true);
  assert.equal(isLoopbackAddress('::1'), true);
  assert.equal(isLoopbackAddress('0.0.0.0'), false);
  assert.equal(isLoopbackAddress('localhost'), false);
});

test('외부 중계: 실제 HTTP 소켓도 검증한 IP만 사용하고 원래 Host를 유지한다', async (t) => {
  let seenHost = '', dnsCalls = 0, pinnedCalls = 0;
  const upstream = createServer((req, res) => { seenHost = req.headers.host; res.end('실제 소켓 응답'); });
  await new Promise((resolve) => upstream.listen(0, '127.0.0.1', resolve));
  t.after(() => new Promise((resolve) => { upstream.close(resolve); upstream.closeAllConnections?.(); }));
  const lookup = async () => (++dnsCalls === 1 ? [{ address: '93.184.216.34', family: 4 }] : [{ address: '127.0.0.1', family: 4 }]);
  const result = await fetchPublicResource('http://public.example/data', {
    lookup,
    request: (target, options) => requestPinned(target, { ...options, requestHttp: (requestOptions, done) => {
      assert.equal(requestOptions.hostname, 'public.example');
      assert.equal(requestOptions.agent, false);
      // 테스트만 검증된 공개 IP를 로컬 fixture로 매핑합니다. 운영 코드에는 이 매핑이 없습니다.
      return httpRequest({ ...requestOptions, port: upstream.address().port, lookup(host, opts, callback) {
        requestOptions.lookup(host, opts, (error, address, family) => {
          assert.ifError(error); assert.equal(address, '93.184.216.34'); assert.equal(family, 4); pinnedCalls++;
          callback(null, '127.0.0.1', 4);
        });
      } }, done);
    } }),
  });
  assert.equal(result.body.toString(), '실제 소켓 응답');
  assert.equal(dnsCalls, 1); assert.equal(pinnedCalls, 1); assert.equal(seenHost, 'public.example');
});

test('외부 중계: HTTPS의 원래 SNI와 인증서 검증을 유지한다', async () => {
  const target = await resolvePublicTarget('https://public.example/a?q=1', publicLookup);
  const result = await requestPinned(target, { requestHttps(options, done) {
    assert.equal(options.hostname, 'public.example'); assert.equal(options.servername, 'public.example');
    assert.equal(options.rejectUnauthorized, true); assert.equal(options.agent, false);
    assert.equal(options.path, '/a?q=1'); assert.equal(options.headers.Host, 'public.example');
    options.lookup('public.example', { all: true }, (error, addresses) => {
      assert.ifError(error); assert.deepEqual(addresses, [{ address: '93.184.216.34', family: 4 }]);
    });
    const req = new EventEmitter(); req.end = () => done(response()); return req;
  } });
  result.destroy();
});

test('외부 중계: 리디렉션은 매 단계 검사하며 최대 다섯 번만 따른다', async () => {
  const visited = [];
  const result = await fetchPublicResource('https://public.example/0', {
    lookup: publicLookup, request: async (target) => {
      visited.push(target.url.pathname);
      return visited.length <= 5 ? response(302, [], { location: '/' + visited.length }) : response();
    },
  });
  assert.equal(result.status, 200); assert.deepEqual(visited, ['/0', '/1', '/2', '/3', '/4', '/5']);
  let calls = 0;
  await assert.rejects(fetchPublicResource('https://public.example/', {
    lookup: publicLookup, request: async () => { calls++; return response(307, [], { location: '/again' }); },
  }), statusIs(502));
  assert.equal(calls, 6);
  for (const destination of ['http://127.0.0.1/', 'http://169.254.169.254/', 'http://public.example:8080/']) {
    let requests = 0;
    await assert.rejects(fetchPublicResource('https://public.example/', {
      lookup: publicLookup, request: async () => { requests++; return response(302, [], { location: destination }); },
    }), statusIs(400));
    assert.equal(requests, 1, destination);
  }
  let dnsCalls = 0, requests = 0;
  await assert.rejects(fetchPublicResource('http://public.example/', {
    lookup: async () => [{ address: ++dnsCalls === 1 ? '93.184.216.34' : '10.0.0.1', family: 4 }],
    request: async () => { requests++; return response(301, [], { location: '/same-host' }); },
  }), statusIs(400));
  assert.equal(requests, 1); assert.equal(dnsCalls, 2);
});

test('외부 중계: 서명 헤더 요청의 리디렉션을 금지해 키가 다른 호스트로 전송되지 않는다', async () => {
  let calls = 0;
  await assert.rejects(fetchPublicResource('https://public.example/', {
    lookup: publicLookup, maxRedirects: 0, headers: { 'X-API-KEY': 'test-only' },
    request: async (_target, options) => {
      calls++; assert.equal(options.headers['X-API-KEY'], 'test-only');
      return response(302, [], { location: 'https://other.example/' });
    },
  }), statusIs(502));
  assert.equal(calls, 1);
});

test('외부 중계: Content-Length 없는 본문과 압축 해제 결과도 스트리밍 크기 제한을 받는다', async () => {
  for (const makeResponse of [
    () => response(200, ['a'], { 'content-length': '101' }),
    () => response(200, [Buffer.alloc(60), Buffer.alloc(41)]),
    () => response(200, [gzipSync(Buffer.alloc(4096, 'a'))], { 'content-encoding': 'gzip' }),
  ]) {
    const stream = makeResponse();
    await assert.rejects(fetchPublicResource('http://public.example/', {
      lookup: publicLookup, maxBytes: 100, request: async () => stream,
    }), statusIs(413));
    assert.equal(stream.destroyed, true);
  }
  const normal = await fetchPublicResource('http://public.example/', {
    lookup: publicLookup, request: async () => response(200, [gzipSync('한국어')], { 'content-encoding': 'gzip' }),
  });
  assert.equal(normal.body.toString(), '한국어');
});

test('외부 중계: DNS와 끝나지 않는 응답 본문에 같은 전체 시간 제한을 적용한다', async () => {
  let requests = 0;
  await assert.rejects(fetchPublicResource('http://public.example/', {
    lookup: () => new Promise(() => {}), timeoutMs: 20,
    request: async () => { requests++; return response(); },
  }), statusIs(504));
  assert.equal(requests, 0);
  const stream = new Readable({ read() {} }); stream.statusCode = 200; stream.headers = {};
  await assert.rejects(fetchPublicResource('http://public.example/', {
    lookup: publicLookup, timeoutMs: 20, request: async () => stream,
  }), statusIs(504));
  assert.equal(stream.destroyed, true);
});
