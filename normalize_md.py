#!/usr/bin/env python3
"""标准化 extracted_content_v2.md 为一致 markdown 格式"""
import re, os

CONTENT_PATH = os.path.expanduser('~/.hermes/project-data/生物统计/extracted_content_v2.md')
OUTPUT_PATH = os.path.expanduser('~/.hermes/project-data/生物统计/extracted_content_v3.md')

with open(CONTENT_PATH, 'r', encoding='utf-8') as f:
    content = f.read()

# 按章节分割
sections = re.split(r'^(# 第.+章.+)$', content, flags=re.MULTILINE)

normalized_sections = []
for section in sections:
    section = section.strip()
    if not section:
        continue
    
    # 如果是章节标题行 (# 第X章 XXX)，保留
    if re.match(r'^# 第.+章.+$', section):
        normalized_sections.append(section)
        continue
    
    # 如果是内容正文，进行标准化
    lines = section.split('\n')
    out_lines = []
    
    i = 0
    while i < len(lines):
        line = lines[i]
        stripped = line.strip()
        prev_out = out_lines[-1] if out_lines else ''
        
        # 空行：跳过较长的连续空行
        if not stripped:
            # 只保留一个空行
            if out_lines and out_lines[-1] != '':
                out_lines.append('')
            i += 1
            continue
        
        # 已有 ## / ### 标题的行：保留
        if stripped.startswith('## ') or stripped.startswith('### ') or stripped.startswith('#### '):
            # 确保前面有空行
            if out_lines and out_lines[-1] != '':
                out_lines.append('')
            out_lines.append(stripped)
            i += 1
            continue
        
        # 纯 # 开头的非章节标题（如 # 水产生物统计）→ 转为 ##
        if stripped.startswith('# ') and not re.match(r'^# 第.+章.+$', stripped):
            if out_lines and out_lines[-1] != '':
                out_lines.append('')
            out_lines.append(f'## {stripped[2:]}')
            i += 1
            continue
        
        # 数字标题如 "1.1 基本概念" 或 "1.1.1 试验指标" → 转为 ###
        m = re.match(r'^(\d+\.\d+(?:\.\d+)?)\s+(.+)$', stripped)
        if m and len(stripped) < 50:
            if out_lines and out_lines[-1] != '':
                out_lines.append('')
            # 用 ### 处理
            out_lines.append(f'### {stripped}')
            i += 1
            continue
        
        # 单数字标题如 "§1 方差分析的基本原理" → ##
        m = re.match(r'^§(\d+)\s+(.+)$', stripped)
        if m and len(stripped) < 60:
            if out_lines and out_lines[-1] != '':
                out_lines.append('')
            out_lines.append(f'## §{m.group(1)} {m.group(2)}')
            i += 1
            continue
        
        # "引 言", "本 章 概 览" 等 → ##
        if re.match(r'^[本引][章节]\s*(概览|言|重点|内容|难点)', stripped) or stripped in ['引 言', '本 章 概 览', '引 言', '复习思考题']:
            if out_lines and out_lines[-1] != '':
                out_lines.append('')
            out_lines.append(f'### {stripped}')
            i += 1
            continue
        
        # • / – / ⚫ 开头的行 → 转换为 markdown 列表
        if stripped.startswith('•'):
            # 去多余的符号
            text = re.sub(r'^•\s*', '', stripped)
            out_lines.append(f'- {text}')
            i += 1
            continue
        
        if stripped.startswith('–') or stripped.startswith('—'):
            text = re.sub(r'^[–—]\s*', '', stripped)
            out_lines.append(f'  - {text}')
            i += 1
            continue
        
        if stripped.startswith('⚫'):
            text = re.sub(r'^⚫\s*', '', stripped)
            out_lines.append(f'  - {text}')
            i += 1
            continue
        
        # 普通内容行：如果前一行也是普通内容且末尾不是句号/问号/冒号/分号，合并
        if out_lines and not out_lines[-1].startswith('#') and not out_lines[-1].startswith('-') and not out_lines[-1].startswith('  -') and not out_lines[-1].startswith('|') and out_lines[-1] != '':
            prev = out_lines[-1]
            # 如果前一行以句号/问号/感叹号/冒号/分号结尾，或者当前行以大写字母开头，不合并
            if re.search(r'[。！？；：:]$', prev) or re.match(r'^[A-Za-z]', stripped):
                out_lines.append(stripped)
            else:
                # 合并到最后一行
                out_lines[-1] = prev + stripped
        else:
            out_lines.append(stripped)
        
        i += 1
    
    normalized_sections.append('\n'.join(out_lines))

result = '\n\n'.join(normalized_sections)
with open(OUTPUT_PATH, 'w', encoding='utf-8') as f:
    f.write(result)

print(f'写入 {OUTPUT_PATH}')
print(f'字符: {len(content)} → {len(result)}')

# 验证章节
nchapters = re.findall(r'^# 第.+章.+$', result, re.MULTILINE)
print(f'章节数: {len(nchapters)}')

# 检查5章结构
ch5_start = result.find('# 第5章 方差分析')
if ch5_start >= 0:
    ch5_end = result.find('# 第6章', ch5_start)
    ch5_text = result[ch5_start:ch5_end] if ch5_end > 0 else result[ch5_start:]
    
    lines5 = ch5_text.split('\n')
    h3 = sum(1 for l in lines5 if l.strip().startswith('###'))
    h2 = sum(1 for l in lines5 if l.strip().startswith('##'))
    bullets = sum(1 for l in lines5 if l.strip().startswith('-') or l.strip().startswith('  -'))
    empty = sum(1 for l in lines5 if not l.strip())
    print(f'\n第5章标准化后:')
    print(f'  ##: {h2} | ###: {h3} | 列表: {bullets} | 空行: {empty}')
