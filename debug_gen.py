#!/usr/bin/env python3
"""调试gen_biostatistics.py - 追踪每章处理"""
import os, re
from docx import Document
from docx.shared import Pt
from docx.oxml.ns import qn
from docx.oxml import OxmlElement

CONTENT_PATH = os.path.expanduser("~/.hermes/project-data/生物统计/extracted_content_v3.md")

with open(CONTENT_PATH, 'r', encoding='utf-8') as f:
    content = f.read()

# 按章节分割
sections = re.split(r'^# (第.+章.+)$', content, flags=re.MULTILINE)
print(f'len(sections) = {len(sections)}')
print(f'sections[:4] = {[s[:50] for s in sections[:4]]}')

chapters = {}
i = 1
while i < len(sections):
    title = sections[i].strip()
    body = sections[i+1].strip() if i+1 < len(sections) else ""
    chapters[title] = body
    print(f'章节: "{title}" → {len(body)}字')
    i += 2

print(f'\n总章节数: {len(chapters)}')
print(f'章节标题列表: {list(chapters.keys())}')

# 检查第0章前的内容（前言）
if sections:
    front = sections[0]
    print(f'\n前言部分: {len(front)}字, 前80字: {front[:80]}')
