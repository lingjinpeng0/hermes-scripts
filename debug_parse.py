#!/usr/bin/env python3
"""调试：追踪 parse_md_to_doc 的内容损失"""
import os, re

CONTENT_PATH = os.path.expanduser('~/.hermes/project-data/生物统计/extracted_content_v2.md')
with open(CONTENT_PATH, 'r', encoding='utf-8') as f:
    content = f.read()

sections = re.split(r'^# (第.+章.+)$', content, flags=re.MULTILINE)
chapters = {}
i = 1
while i < len(sections):
    title = sections[i].strip()
    body = sections[i+1].strip() if i+1 < len(sections) else ''
    chapters[title] = body
    i += 2

def strip_md(text):
    text = re.sub(r'\*\*(.+?)\*\*', r'\1', text)
    text = re.sub(r'\*(.+?)\*', r'\1', text)
    text = re.sub(r'`(.+?)`', r'\1', text)
    return text

def trace_parse(md_text):
    lines = md_text.split('\n')
    processed = 0
    stats = {'##_body': 0, '###_head': 0, '####_head': 0, 'tables': 0, 'lists': 0, 'paras': 0, 'empty': 0, 'lines_total': len(lines)}
    
    i = 0
    while i < len(lines):
        line = lines[i].strip()
        if not line:
            stats['empty'] += 1
            i += 1
            continue
        if line == '---':
            i += 1
            continue
        if line.startswith('## '):
            text = line[3:].strip()
            processed += len(text)
            stats['##_body'] += 1
            i += 1
            continue
        if line.startswith('### '):
            text = line[4:].strip()
            processed += len(text)
            stats['###_head'] += 1
            i += 1
            continue
        if line.startswith('#### '):
            text = line[5:].strip()
            processed += len(text)
            stats['####_head'] += 1
            i += 1
            continue
        if '|' in line and i+1 < len(lines) and lines[i+1].strip().startswith('|---'):
            headers = [c.strip() for c in line.split('|') if c.strip()]
            i += 2
            rows = []
            while i < len(lines) and '|' in lines[i]:
                cells = [c.strip() for c in lines[i].split('|') if c.strip()]
                if cells:
                    rows.append(cells)
                i += 1
            if headers and rows:
                stats['tables'] += 1
                for row in rows:
                    for cell in row:
                        processed += len(cell)
            continue
        if line.startswith('- '):
            text = line[2:].strip()
            processed += len(text) + 2
            stats['lists'] += 1
            i += 1
            continue
        if re.match(r'^\d+\.', line):
            text = strip_md(line)
            processed += len(text)
            stats['lists'] += 1
            i += 1
            continue
        if line.startswith('> '):
            text = line[2:].strip()
            processed += len(text)
            i += 1
            continue
        para_lines = [line]
        i += 1
        while i < len(lines):
            next_line = lines[i].strip()
            if not next_line or next_line.startswith('#') or next_line.startswith('- ') or next_line.startswith('>') or next_line.startswith('|') or next_line == '---':
                break
            para_lines.append(next_line)
            i += 1
        text = ' '.join(para_lines)
        text = strip_md(text)
        processed += len(text)
        stats['paras'] += 1
    return processed, stats

total_src = 0
total_proc = 0
for ch_title, md_text in chapters.items():
    p, stats = trace_parse(md_text)
    total_src += len(md_text)
    total_proc += p
    loss = len(md_text) - p
    print(f'{ch_title}:')
    print(f'  源字: {len(md_text):>6} → 处理后: {p:>6} (损失: {loss:>6}, 损失率: {loss/len(md_text)*100:.1f}%)')
    print(f'  线数: {stats["lines_total"]} | ##: {stats["##_body"]} | ###: {stats["###_head"]} | 表格: {stats["tables"]} | 列表: {stats["lists"]} | 段落: {stats["paras"]} | 空: {stats["empty"]}')

print(f'\n总计: 源字 {total_src} → 处理后 {total_proc} (损失 {total_src - total_proc}, 损失率 {(total_src-total_proc)/total_src*100:.1f}%)')
