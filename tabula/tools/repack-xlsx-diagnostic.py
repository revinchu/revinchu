"""Repackage a diagnostic COPY with Python's ZIP writer, without changing part bytes.
Never edits the source or overwrites an existing output. No XML or workbook edits.
Usage: python tools/repack-xlsx-diagnostic.py source.xlsx output.xlsx
"""
import argparse
import json
from pathlib import Path
import zipfile


def repack(source, output):
    source, output = Path(source).resolve(), Path(output).resolve()
    if source == output or output.exists():
        raise ValueError('Output must be a new, different file')
    before = source.stat()
    expected = []
    with zipfile.ZipFile(source, 'r') as zin, zipfile.ZipFile(output, 'x', compression=zipfile.ZIP_DEFLATED, compresslevel=6, allowZip64=False) as zout:
        zout.comment = zin.comment
        for old in zin.infolist():
            entry = zipfile.ZipInfo(old.filename, old.date_time)
            entry.compress_type = zipfile.ZIP_DEFLATED
            entry.external_attr, entry.internal_attr = old.external_attr, old.internal_attr
            entry.comment = old.comment
            entry.file_size = old.file_size
            with zin.open(old, 'r') as src, zout.open(entry, 'w', force_zip64=False) as dst:
                while True:
                    block = src.read(1 << 20)
                    if not block:
                        break
                    dst.write(block)
            expected.append((old.filename, old.CRC, old.file_size))
    with zipfile.ZipFile(output) as check:
        actual = [(entry.filename, entry.CRC, entry.file_size) for entry in check.infolist()]
        if actual != expected:
            raise ValueError('Repack changed entry names, CRC or uncompressed sizes')
    after = source.stat()
    if (before.st_size, before.st_mtime_ns) != (after.st_size, after.st_mtime_ns):
        raise ValueError('Source metadata changed during diagnostic repack')
    return {'entries': len(expected), 'sourceBytes': before.st_size, 'outputBytes': output.stat().st_size, 'partsUnchanged': True, 'sourceUnchanged': True}


if __name__ == '__main__':
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument('source')
    parser.add_argument('output')
    args = parser.parse_args()
    print(json.dumps(repack(args.source, args.output)))
