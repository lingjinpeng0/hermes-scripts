#!/usr/bin/env python3
"""Run mineru OCR on one chapter PDF, capturing real-time progress."""

import subprocess
import sys
import os

chapter_num = sys.argv[1]
base_dir = r'C:\Users\Rei\Downloads\学习\渔业资源学'
output_dir = r'C:\Users\Rei\.hermes\project-data\渔业资源学\ocr'

files = {
    '1': '1.绪论.pdf',
    '2': '2.鱼类资源的生物学基础.pdf',
    '3': '3.鱼类资源数量的变动.pdf',
    '4': '4.渔业资源评估.pdf',
    '5': '5.渔业资源保护与管理.pdf',
}

pdf_path = os.path.join(base_dir, files[chapter_num])
out_path = os.path.join(output_dir, f'ch{chapter_num}')

print(f"=== OCR Chapter {chapter_num}: {files[chapter_num]} ===")
print(f"Input: {pdf_path}")
print(f"Output: {out_path}")
print()

env = os.environ.copy()
env['HTTP_PROXY'] = '127.0.0.1:7897'
env['HTTPS_PROXY'] = '127.0.0.1:7897'
env['PYTHONPATH'] = ''
env['HF_HUB_DISABLE_SYMLINKS'] = '1'

proc = subprocess.Popen(
    ['mineru', '-p', pdf_path, '-o', out_path, '-b', 'pipeline', '-m', 'ocr'],
    stdout=subprocess.PIPE,
    stderr=subprocess.STDOUT,
    text=True,
    bufsize=1,
    env=env
)

for line in proc.stdout:
    print(line, end='', flush=True)

proc.wait()
print(f"\n=== Exit code: {proc.returncode} ===")
sys.exit(proc.returncode)
