"""Create a diagnostic-only XLSX copy with object/pivot/table references removed.
The source is untouched; worksheet prefixes/sheetData, styles and shared strings
are preserved byte-for-byte. This is NOT a product export or a repaired workbook.
"""
import argparse
import json
import hashlib
from pathlib import Path
import posixpath
import re
import xml.etree.ElementTree as ET
import zipfile
import zlib

MAIN = 'http://schemas.openxmlformats.org/spreadsheetml/2006/main'
ET.register_namespace('', MAIN)
ET.register_namespace('r', 'http://schemas.openxmlformats.org/officeDocument/2006/relationships')
REMOVE_TAGS = {'drawing','legacyDrawing','legacyDrawingHF','picture','oleObjects','controls','tableParts','pivotTableParts','extLst'}
DROP_PREFIXES = ('xl/drawings/', 'xl/charts/', 'xl/slicers/', 'xl/slicerCaches/', 'xl/pivotTables/', 'xl/pivotCache/', 'xl/tables/')
CHUNK = 1 << 20


def local(tag):
    return tag.rsplit('}', 1)[-1]


def rewrite_sheet(src, dst, remove_tags=REMOVE_TAGS):
    prefix = b''
    while True:
        block = src.read(CHUNK)
        if not block:
            raise ValueError('Missing sheetData')
        prefix += block
        match = re.search(br'<sheetData(?:\s[^>]*|/)?>', prefix)
        if match:
            break
        if len(prefix) > 16 << 20:
            raise ValueError('Diagnostic worksheet prefix limit')
    self_closing = prefix[match.start():match.end()].endswith(b'/>' )
    header = prefix[:match.end()]
    if self_closing: header = header[:-2] + b'>'
    opening = re.search(br'<worksheet\b([^>]*)>', header).group(1)
    dst.write(header)
    buffer = (b'</sheetData>' if self_closing else b'') + prefix[match.end():]
    crc, length = 0, 0
    closing = b'</sheetData>'
    while True:
        at = buffer.find(closing)
        if at >= 0:
            data = buffer[:at]
            dst.write(data); crc = zlib.crc32(data, crc); length += len(data)
            tail = buffer[at + len(closing):]
            while True:
                block = src.read(CHUNK)
                if not block:
                    break
                tail += block
                if len(tail) > 32 << 20:
                    raise ValueError('Diagnostic worksheet tail limit')
            break
        end = max(0, len(buffer) - len(closing) + 1)
        data = buffer[:end]
        dst.write(data); crc = zlib.crc32(data, crc); length += len(data)
        buffer = buffer[end:]
        block = src.read(CHUNK)
        if not block:
            raise ValueError('Missing sheetData closing tag')
        buffer += block
    if not tail.rstrip().endswith(b'</worksheet>'):
        raise ValueError('Unexpected worksheet end')
    body = tail.rstrip()[:-len(b'</worksheet>')]
    parsed = ET.fromstring(b'<diagnostic' + opening + b'>' + body + b'</diagnostic>')
    removed = []
    dst.write(closing)
    for child in parsed:
        if local(child.tag) in remove_tags:
            removed.append(local(child.tag))
        else:
            dst.write(ET.tostring(child, encoding='utf-8'))
    dst.write(b'</worksheet>')
    return {'sheetDataBytes': length, 'sheetDataCrc32': f'{crc & 0xffffffff:08x}', 'prefixBytes': len(header), 'removedTags': removed}


def make_copy(source, output, keep_features='none'):
    source, output = Path(source).resolve(), Path(output).resolve()
    if source == output or output.exists():
        raise ValueError('Output must be new and different')
    before = source.stat()
    drop_prefixes = DROP_PREFIXES
    remove_tags = REMOVE_TAGS
    if keep_features == 'pivots-tables':
        drop_prefixes = ('xl/drawings/', 'xl/charts/', 'xl/slicers/', 'xl/slicerCaches/')
        remove_tags = REMOVE_TAGS - {'tableParts', 'pivotTableParts'}
    elif keep_features in {'cache-tables', 'cache1-tables', 'cache23-tables'}:
        drop_prefixes = ('xl/drawings/', 'xl/charts/', 'xl/slicers/', 'xl/slicerCaches/', 'xl/pivotTables/')
        remove_tags = REMOVE_TAGS - {'tableParts'}
    elif keep_features == 'drawings-tables':
        drop_prefixes = ('xl/pivotTables/', 'xl/pivotCache/', 'xl/slicers/', 'xl/slicerCaches/')
        remove_tags = {'pivotTableParts', 'extLst'}
    report = {'diagnosticOnly': True, 'keepFeatures': keep_features, 'removedParts': [], 'changedParts': [], 'sheets': []}
    with zipfile.ZipFile(source) as zin:
        names = set(zin.namelist())
        removed = {name for name in names if name.startswith(drop_prefixes)}
        if keep_features in {'cache1-tables', 'cache23-tables'}:
            for name in names:
                match = re.fullmatch(r'xl/pivotCache/(?:_rels/)?pivotCache(?:Definition|Records)(\d+)\.xml(?:\.rels)?', name)
                if match and ((match.group(1) != '1') if keep_features == 'cache1-tables' else (match.group(1) == '1')): removed.add(name)
        wb_rel_root = ET.fromstring(zin.read('xl/_rels/workbook.xml.rels'))
        removed_cache_rel_ids = {e.get('Id') for e in wb_rel_root if posixpath.normpath('xl/'+e.get('Target','')).lstrip('/') in removed}
        report['removedParts'] = sorted(removed)
        report['sourceParts'] = [{'part':e.filename,'bytes':e.file_size,'crc32':f'{e.CRC:08x}'} for e in zin.infolist()]
        with zipfile.ZipFile(output, 'x', compression=zipfile.ZIP_DEFLATED, compresslevel=6, allowZip64=False) as zout:
            for old in zin.infolist():
                name = old.filename
                if name in removed:
                    continue
                if re.fullmatch(r'xl/worksheets/sheet\d+\.xml', name):
                    with zin.open(old) as src, zout.open(name, 'w', force_zip64=False) as dst:
                        details = rewrite_sheet(src, dst, remove_tags)
                    report['sheets'].append({'part': name, **details}); report['changedParts'].append(name)
                    continue
                if name == 'xl/workbook.xml':
                    tree = ET.fromstring(zin.read(old))
                    for child in list(tree):
                        if local(child.tag) == 'extLst' or (local(child.tag) == 'pivotCaches' and keep_features not in {'pivots-tables', 'cache-tables', 'cache1-tables', 'cache23-tables'}):
                            tree.remove(child)
                        elif local(child.tag) == 'pivotCaches':
                            for cache in list(child):
                                if cache.get('{http://schemas.openxmlformats.org/officeDocument/2006/relationships}id') in removed_cache_rel_ids: child.remove(cache)
                    data = ET.tostring(tree, encoding='utf-8', xml_declaration=True)
                    report['changedParts'].append(name)
                elif keep_features == 'drawings-tables' and re.fullmatch(r'xl/drawings/drawing\d+\.xml', name):
                    tree = ET.fromstring(zin.read(old))
                    for anchor in list(tree):
                        if any(local(e.tag) == 'slicer' for e in anchor.iter()): tree.remove(anchor)
                    data = ET.tostring(tree, encoding='utf-8', xml_declaration=True)
                    report['changedParts'].append(name)
                elif name == '[Content_Types].xml':
                    tree = ET.fromstring(zin.read(old))
                    for child in list(tree):
                        if child.get('PartName', '').lstrip('/') in removed:
                            tree.remove(child)
                    data = ET.tostring(tree, encoding='utf-8', xml_declaration=True)
                    report['changedParts'].append(name)
                elif name.endswith('.rels'):
                    tree = ET.fromstring(zin.read(old)); base = posixpath.dirname(posixpath.dirname(name)); changed = False
                    for child in list(tree):
                        target = posixpath.normpath(posixpath.join(base, child.get('Target', ''))).lstrip('/')
                        if child.get('TargetMode') != 'External' and target in removed:
                            tree.remove(child); changed = True
                    data = ET.tostring(tree, encoding='utf-8', xml_declaration=True) if changed else zin.read(old)
                    if changed: report['changedParts'].append(name)
                else:
                    with zin.open(old) as src, zout.open(name, 'w', force_zip64=False) as dst:
                        while True:
                            data = src.read(CHUNK)
                            if not data: break
                            dst.write(data)
                    continue
                zout.writestr(name, data)
    after = source.stat()
    if (before.st_size, before.st_mtime_ns) != (after.st_size, after.st_mtime_ns):
        raise ValueError('Source metadata changed')
    report.update(sourceUnchanged=True, sourceBytes=before.st_size, outputBytes=output.stat().st_size)
    with zipfile.ZipFile(source) as zin, zipfile.ZipFile(output) as zout:
        report['stylesAndStringsUnchanged'] = all((zin.getinfo(n).CRC, zin.getinfo(n).file_size) == (zout.getinfo(n).CRC, zout.getinfo(n).file_size) for n in ['xl/styles.xml','xl/sharedStrings.xml'])
    class Sink:
        def write(self, data): return len(data)
    with zipfile.ZipFile(output) as zout:
        for item in report['sheets']:
            with zout.open(item['part']) as src: check = rewrite_sheet(src, Sink())
            if (check['sheetDataBytes'],check['sheetDataCrc32']) != (item['sheetDataBytes'],item['sheetDataCrc32']): raise ValueError('sheetData changed')
            item['outputSheetDataCrc32'] = check['sheetDataCrc32']; item['sheetDataVerified'] = True
    with source.open('rb') as src: report['sourceSha256'] = hashlib.file_digest(src,'sha256').hexdigest()
    report['allSheetDataVerified'] = True
    return report


if __name__ == '__main__':
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument('source'); parser.add_argument('output'); parser.add_argument('report')
    parser.add_argument('--keep-features', choices=['none','pivots-tables','drawings-tables','cache-tables','cache1-tables','cache23-tables'], default='none')
    args = parser.parse_args()
    result = make_copy(args.source, args.output, args.keep_features)
    Path(args.report).write_text(json.dumps(result, indent=2), encoding='utf-8')
    print(json.dumps({'outputBytes':result['outputBytes'],'removedParts':len(result['removedParts']),'sheets':len(result['sheets']),'sourceUnchanged':result['sourceUnchanged'],'stylesAndStringsUnchanged':result['stylesAndStringsUnchanged']}))
