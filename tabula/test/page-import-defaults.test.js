import test from 'node:test';
import assert from 'node:assert/strict';
import { pageFromXml, pageXml, normPage } from '../src/page.js';
import { parseXml } from '../src/xml.js';

test('imported pageSetup without paperSize uses Excel Letter and preserves explicit A4', () => {
  assert.equal(normPage().paper,9);
  const source=pageFromXml({pageSetup:parseXml('<pageSetup orientation="portrait"/>')});
  assert.equal(source.paper,1);
  const written=pageXml(source,String);
  assert.match(written.setup,/paperSize="1"/);
  assert.equal(pageFromXml({pageSetup:parseXml(written.setup)}).paper,1);
  assert.equal(normPage(pageFromXml({pageSetup:parseXml('<pageSetup paperSize="9" orientation="portrait"/>')})).paper,9);
  assert.equal(pageFromXml({}),null);
});
