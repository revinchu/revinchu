#!/usr/bin/env python3
"""Read-only, bounded-memory OOXML container audit. Never prints XML content."""
import argparse
import json
import os
import time
import zipfile
import zlib
from xml.parsers import expat

CHUNK_BYTES = 64 * 1024

def audit_container(source):
    before = os.stat(source)
    started = time.monotonic()
    result = {'scope': 'ZIP CRC, expanded byte lengths and XML namespace well-formedness only; no OOXML schema or Excel semantic validation', 'chunkBytes': CHUNK_BYTES, 'fileBytes': before.st_size, 'entries': 0, 'xmlEntries': 0, 'expandedBytes': 0, 'errors': [], 'warnings': [], 'entryChecks': []}
    with zipfile.ZipFile(source, 'r') as archive:
        seen = set()
        for info in archive.infolist():
            if info.is_dir():
                continue
            name = info.filename
            if name in seen:
                result['errors'].append({'entry': name, 'kind': 'duplicate-entry'})
            seen.add(name)
            is_xml = name.endswith(('.xml', '.rels'))
            parser = expat.ParserCreate(namespace_separator='}') if is_xml else None
            if parser:
                parser.SetParamEntityParsing(expat.XML_PARAM_ENTITY_PARSING_NEVER)
                parser.ExternalEntityRefHandler = lambda *args: 0
                def reject_doctype(*args):
                    raise ValueError('DTD is not permitted in the audit parser')
                parser.StartDoctypeDeclHandler = reject_doctype
            size = crc = 0
            xml_error = None
            read_error = None
            try:
                with archive.open(info, 'r') as stream:
                    while True:
                        chunk = stream.read(CHUNK_BYTES)
                        if not chunk:
                            break
                        size += len(chunk)
                        crc = zlib.crc32(chunk, crc)
                        if parser and xml_error is None:
                            try:
                                parser.Parse(chunk, False)
                            except expat.ExpatError as error:
                                xml_error = {'kind': 'xml-syntax', 'code': error.code, 'message': expat.ErrorString(error.code), 'byteOffset': parser.ErrorByteIndex, 'line': error.lineno, 'column': error.offset}
                            except ValueError:
                                xml_error = {'kind': 'xml-doctype', 'byteOffset': parser.CurrentByteIndex}
                    if parser and xml_error is None:
                        try:
                            parser.Parse(b'', True)
                        except expat.ExpatError as error:
                            xml_error = {'kind': 'xml-syntax', 'code': error.code, 'message': expat.ErrorString(error.code), 'byteOffset': parser.ErrorByteIndex, 'line': error.lineno, 'column': error.offset}
            except (zipfile.BadZipFile, RuntimeError, OSError, zlib.error) as error:
                read_error = {'kind': 'zip-read', 'errorType': type(error).__name__, 'bytesRead': size}
            check = {'entry': name, 'compressedBytes': info.compress_size, 'expectedBytes': info.file_size, 'actualBytes': size, 'expectedCrc32': f'{info.CRC:08x}', 'actualCrc32': f'{crc & 0xffffffff:08x}', 'xml': is_xml}
            if size != info.file_size:
                result['errors'].append({'entry': name, 'kind': 'expanded-size', 'expected': info.file_size, 'actual': size})
            if (crc & 0xffffffff) != info.CRC:
                result['errors'].append({'entry': name, 'kind': 'crc32', 'expected': check['expectedCrc32'], 'actual': check['actualCrc32']})
            for error in (read_error, xml_error):
                if error:
                    result['errors'].append({'entry': name, **error})
            check['ok'] = read_error is None and xml_error is None and size == info.file_size and (crc & 0xffffffff) == info.CRC
            result['entryChecks'].append(check)
            result['entries'] += 1
            result['xmlEntries'] += int(is_xml)
            result['expandedBytes'] += size
    after = os.stat(source)
    result['sourceUnchanged'] = before.st_size == after.st_size and before.st_mtime_ns == after.st_mtime_ns
    if not result['sourceUnchanged']:
        result['errors'].append({'kind': 'source-changed-during-audit'})
    result['elapsedMs'] = round((time.monotonic() - started) * 1000)
    result['ok'] = not result['errors']
    return result

def main():
    cli = argparse.ArgumentParser(description=__doc__)
    cli.add_argument('source')
    cli.add_argument('--out', required=True)
    args = cli.parse_args()
    if os.path.normcase(os.path.realpath(args.source)) == os.path.normcase(os.path.realpath(args.out)):
        cli.error('결과 경로는 입력 파일과 달라야 합니다.')
    report = audit_container(args.source)
    with open(args.out, 'w', encoding='utf-8') as stream:
        json.dump(report, stream, ensure_ascii=False, indent=2)
    print(json.dumps({key: report[key] for key in ['ok', 'entries', 'xmlEntries', 'expandedBytes', 'errors', 'elapsedMs', 'sourceUnchanged']}, ensure_ascii=False))
    return 0 if report['ok'] else 1

if __name__ == '__main__':
    raise SystemExit(main())
