#!/usr/bin/env python3
"""模拟 parse_md_to_doc 并追踪输出"""
import re, os

CONTENT_PATH = os.path.expanduser("~/.hermes/project-data/生物统计/extracted_content_v3.md")

with open(CONTENT_PATH, 'r', encoding='utf-8') as f:
    content = f.read()

sections = re.split(r'^# (第.+章.+)$', content, flags=re.MULTILINE)
chapters = {}
i = 1
while i < len(sections):
    title = sections[i].strip()
    body = sections[i+1].strip() if i+1 < len(sections) else ""
    chapters[title] = body
    i += 2

def parse_debug(md_text):
    """返回处理后的字符数和行数统计"""
    lines = md_text.split('\n')
    output_chars = 0
    i = 0
    line_types = {}
    
    while i < len(lines):
        line = lines[i].strip()
        
        if not line:
            line_types['empty'] = line_types.get('empty', 0) + 1
            i += 1
            continue
        if line == '---':
            line_types['hr'] = line_types.get('hr', 0) + 1
            i += 1
            continue
        if line.startswith('## '):
            text = line[3:].strip()
            output_chars += len(text)
            line_types['h2_body'] = line_types.get('h2_body', 0) + 1
            i += 1
            continue
        if line.startswith('### '):
            text = line[4:].strip()
            output_chars += len(text)
            line_types['h3_head'] = line_types.get('h3_head', 0) + 1
            i += 1
            continue
        if line.startswith('#### '):
            text = line[5:].strip()
            output_chars += len(text)
            line_types['h4_body'] = line_types.get('h4_body', 0) + 1
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
                line_types['table'] = line_types.get('table', 0) + 1
                for row in rows:
                    for cell in row:
                        output_chars += len(cell)
            continue
        if line.startswith('- '):
            text = line[2:].strip()
            text = re.sub(r'\*\*(.+?)\*\*', r'\1', text)
            output_chars += len(text) + 2
            line_types['list'] = line_types.get('list', 0) + 1
            i += 1
            continue
        if re.match(r'^\d+\.', line):
            text = re.sub(r'\*\*(.+?)\*\*', r'\1', line)
            output_chars += len(text)
            line_types['olist'] = line_types.get('olist', 0) + 1
            i += 1
            continue
        if line.startswith('> '):
            text = line[2:].strip()
            output_chars += len(text)
            line_types['quote'] = line_types.get('quote', 0) + 1
            i += 1
            continue
        
        # 普通段落
        para_lines = [line]
        i += 1
        while i < len(lines):
            next_line = lines[i].strip()
            if not next_line or next_line.startswith('#') or next_line.startswith('- ') or next_line.startswith('>') or next_line.startswith('|') or next_line == '---' or re.match(r'^\d+\.', next_line):
                break
            para_lines.append(next_line)
            i += 1
        
        text = ' '.join(para_lines)
        text = re.sub(r'\*\*(.+?)\*\*', r'\1', text)
        output_chars += len(text)
        line_types['para'] = line_types.get('para', 0) + 1
    
    return output_chars, line_types

total_in = 0
total_out = 0
for title, body in chapters.items():
    out_chars, types = parse_debug(body)
    total_in += len(body)
    total_out += out_chars
    loss = len(body) - out_chars
    pct = loss / len(body) * 100
    print(title + ":")
    print("  输入 {:>6}字 -> 输出 {:>6}字 (损失 {:>6}, {:.1f}%)".format(len(body), out_chars, loss, pct))
    h2b = types.get("h2_body", 0)
    h3h = types.get("h3_head", 0)
    h4b = types.get("h4_body", 0)
    tbl = types.get("table", 0)
    lst = types.get("list", 0)
    olst = types.get("olist", 0)
    prs = types.get("para", 0)
    print("  类型: h2_body={} h3_head={} h4_body={} table={} list={} olist={} para={}".format(h2b, h3h, h4b, tbl, lst, olst, prs))

print("\n总计: 输入 {}字 -> 输出 {}字".format(total_in, total_out))
