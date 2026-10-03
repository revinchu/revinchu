import { parseXml, allText } from './xml.js';
import { readPhonetic } from './phonetic.js';

import { scanXmlChildren } from './xml-stream.js';

/** 공유 문자열과 윗주를 항목 단위로 읽고 주기적으로 진행 상태를 돌려줍니다. */
export function* readSharedStrings(bytes, fonts = [], chunkSize = 1 << 20) {
  const strings = [], phonetics = [];
  if (!bytes) return { strings, phonetics };
  for (const xml of scanXmlChildren(bytes, 'si', Math.max(64, chunkSize))) {
    const node = parseXml(xml), text = allText(node), index = strings.length;
    strings.push(text);
    const phonetic = readPhonetic(node, text, fonts);
    if (phonetic) phonetics[index] = phonetic;
    if (strings.length % 2048 === 0) yield strings.length;
  }
  return { strings, phonetics };
}
