#!/usr/bin/env python3
"""Verify the full local transfer archive using SHA-256 (standard library only)."""
import hashlib
import json
from pathlib import Path
import sys

ROOT = Path(__file__).resolve().parent

def main():
    manifest = ROOT/'TRANSFER_MANIFEST.json'
    if not manifest.is_file():
        print('TRANSFER_MANIFEST.json is absent. It is included in the complete personal archive, not the GitHub working copy.')
        return 1
    data = json.loads(manifest.read_text(encoding='utf-8'))
    failures = []
    for entry in data['files']:
        path = ROOT/entry['path']
        if not path.is_file():
            failures.append(entry['path']+' (missing)')
            continue
        digest = hashlib.sha256(path.read_bytes()).hexdigest()
        if digest != entry['sha256']:
            failures.append(entry['path']+' (changed)')
    if failures:
        print('Files missing or changed:', len(failures))
        print('\n'.join(failures))
        return 1
    print('PASS:', len(data['files']), 'files match the transfer manifest.')
    return 0

if __name__ == '__main__':
    sys.exit(main())
