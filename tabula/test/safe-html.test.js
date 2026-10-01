import { test } from 'node:test';
import assert from 'node:assert/strict';
import { safeUrl } from '../src/safe-html.js';

test('HTML 주소는 목적별 프로토콜과 제어문자를 검사한다', () => {
  for (const value of ['javascript:alert(1)', 'JaVaScRiPt:alert(1)', 'java\nscript:alert(1)', 'vbscript:msgbox(1)', 'file:///D:/secret', 'data:text/html,<script>alert(1)</script>', 'https://user:pass@example.com/', '\\evil.test/path']) {
    assert.equal(safeUrl(value), null, value);
    assert.equal(safeUrl(value, 'image'), null, value);
  }
  for (const value of ['https://example.com/a?x=1', 'http://example.com/', '#Sheet1!A1', '/docs/1', 'mailto:hello@example.com', 'tel:+821012345678']) assert.equal(safeUrl(value), value);
  assert.equal(safeUrl('blob:https://example.com/123', 'link'), null);
  assert.equal(safeUrl('blob:https://example.com/123', 'image'), 'blob:https://example.com/123');
  assert.equal(safeUrl('data:image/png;base64,AA==', 'image'), 'data:image/png;base64,AA==');
  assert.equal(safeUrl('data:image/svg+xml,%3Csvg/%3E', 'image'), 'data:image/svg+xml,%3Csvg/%3E');
  assert.equal(safeUrl('data:image/png;base64,AA=='), null);
  assert.equal(safeUrl('https://example.com/image.svg#x', 'reference'), null);
  assert.equal(safeUrl('#gradient', 'reference'), '#gradient');
});
