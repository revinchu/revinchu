// MS-XLSB 2.4.711/710/73/74/86/195: binary value metadata -> the existing
// OOXML picture reader. This only maps XLRICHVALUE; other metadata stays unmapped.
const utf16 = new TextDecoder('utf-16le');
const escape = text => text.replace(/[&<>"]/g, c => ({'&':'&amp;','<':'&lt;','>':'&gt;','"':'&quot;'}[c]));
const invalid = () => { throw new Error('XLSB 그림 값 메타데이터의 형식이 올바르지 않습니다.'); };
function* records(bytes) {
  let p = 0;
  while (p < bytes.length) {
    let type = bytes[p++];
    if (type & 128) { if (p >= bytes.length) invalid(); type = (type & 127) | ((bytes[p++] & 127) << 7); }
    let length = 0, ended = false;
    for (let shift = 0; shift < 28; shift += 7) {
      if (p >= bytes.length) invalid();
      const n = bytes[p++]; length += (n & 127) * 2 ** shift;
      if (!(n & 128)) { ended = true; break; }
    }
    if (!ended || p + length > bytes.length) invalid();
    yield { type, data: bytes.subarray(p, p + length) }; p += length;
  }
}
/** Preserve metadata block indexes, including unknown types and empty blocks. */
export function xlsbRichMetadataXml(bytes) {
  const types = [], groups = [], values = [];
  let group = null, block = -1, pair = null, valueCount = null, seenValues = 0;
  for (const { type, data } of records(bytes)) {
    const view = new DataView(data.buffer, data.byteOffset, data.byteLength);
    const u32 = offset => { if (offset + 4 > data.length) invalid(); return view.getUint32(offset, true); };
    const string = offset => { const n = u32(offset); if (n > 65535 || offset + 4 + n * 2 > data.length) invalid(); return utf16.decode(data.subarray(offset + 4, offset + 4 + n * 2)); };
    const index = () => { if (data.length !== 8 || u32(0) !== 0) invalid(); return u32(4); };
    if (type === 335) types.push(string(8));
    else if (type === 339) { group = { name: string(4), count: u32(0), blocks: [] }; groups.push(group); block = -1; }
    else if (type === 340) {
      if (group?.name === 'XLRICHVALUE' && (pair || group.blocks.length !== group.count)) invalid();
      group = null; block = -1;
    } else if (type === 52 && group) {
      if (pair) invalid(); block = group.blocks.length; group.blocks.push(null);
    } else if (type === 5002 && group?.name === 'XLRICHVALUE' && block >= 0) {
      if (pair || group.blocks[block] !== null || (data.length !== 0 && data.length !== 8)) invalid();
      pair = { index: data.length ? index() : null };
    } else if (type === 5003 && group?.name === 'XLRICHVALUE' && block >= 0) {
      // The specification puts FRTBlank+irv on Begin. Native Excel also writes
      // the exact payload on End with an empty Begin. Accept only those pairs.
      if (!pair || (pair.index === null ? data.length !== 8 : data.length !== 0)) invalid();
      group.blocks[block] = pair.index === null ? index() : pair.index; pair = null;
    } else if (type === 53) { if (pair) invalid(); block = -1; }
    else if (type === 337) { valueCount = u32(4) === 0 ? u32(0) : null; seenValues = 0; }
    else if (type === 338) { if (valueCount !== null && seenValues !== valueCount) invalid(); valueCount = null; }
    else if (type === 51 && valueCount !== null) {
      const count = u32(0); if (!count || count * 8 !== data.length - 4) invalid();
      const refs = [];
      for (let i = 0; i < count; i++) refs.push({ type: u32(4 + i * 8), index: u32(8 + i * 8) });
      values.push(refs); seenValues++;
    }
  }
  if (pair) invalid();
  const rich = groups.find(g => g.name === 'XLRICHVALUE'), blocks = rich?.blocks ?? [];
  const valueXml = values.map(refs => {
    const known = refs.filter(r => types[r.type - 1] === 'XLRICHVALUE');
    if (known.some(r => r.index >= blocks.length)) invalid();
    return `<bk>${known.map(r => `<rc t="${r.type}" v="${r.index}"/>`).join('')}</bk>`;
  }).join('');
  const futureXml = rich ? `<futureMetadata name="XLRICHVALUE" count="${blocks.length}">${blocks.map(i => i === null ? '<bk/>' : `<bk><extLst><ext uri="{3e2802c4-a4d2-4d8b-9148-e3be6c30e623}"><xlrd:rvb i="${i}"/></ext></extLst></bk>`).join('')}</futureMetadata>` : '';
  return `<metadata xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:xlrd="http://schemas.microsoft.com/office/spreadsheetml/2017/richdata"><metadataTypes count="${types.length}">${types.map(name => `<metadataType name="${escape(name)}"/>`).join('')}</metadataTypes>${futureXml}<valueMetadata count="${values.length}">${valueXml}</valueMetadata></metadata>`;
}
