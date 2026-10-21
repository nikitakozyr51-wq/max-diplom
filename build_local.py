#!/usr/bin/env python3
"""Build the thesis using paths relative to this project, on Windows/macOS/Linux."""
import argparse
import importlib.util
import json
from pathlib import Path
import subprocess
import sys

ROOT = Path(__file__).resolve().parent

def main():
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument('--check', action='store_true', help='Check prerequisites without building')
    parser.add_argument('--input', type=Path, default=ROOT/'research'/'manuscript.json')
    parser.add_argument('--output', type=Path, default=ROOT/'docs'/'Working_Draft_Zhurilov_RoRo')
    args = parser.parse_args()
    if sys.version_info < (3, 10):
        parser.exit(1, 'Python 3.10 or newer is required.\n')
    missing = [name for name in ('docx', 'fitz', 'PIL', 'lxml') if importlib.util.find_spec(name) is None]
    if missing:
        parser.exit(1, 'Missing modules: '+', '.join(missing)+'. Run: python -m pip install -r requirements.txt\n')
    script = ROOT/'tools'/'build_thesis.py'
    spec = importlib.util.spec_from_file_location('thesis_builder', script)
    builder = importlib.util.module_from_spec(spec)
    spec.loader.exec_module(builder)
    data = json.loads(args.input.read_text(encoding='utf-8'))
    for asset in (builder.TEMPLATE, builder.PLAN):
        if not asset.is_file():
            parser.exit(1, 'Missing file: '+str(asset)+'\n')
    executable = builder.find_soffice()
    print('Manuscript:', args.input.resolve(), flush=True)
    print('Blocks:', len(data['blocks']), '| References:', len(data.get('references', [])), flush=True)
    print('LibreOffice:', executable, flush=True)
    if args.check:
        print('Prerequisites OK. No files changed.')
        return
    subprocess.run([sys.executable, str(script), str(args.input.resolve()), str(args.output.resolve())], check=True)

if __name__ == '__main__':
    main()

