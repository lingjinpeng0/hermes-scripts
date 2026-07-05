"""Check ch5 formula rendering in the docx."""
from docx import Document

doc = Document(r'C:\Users\Rei\Downloads\学习\生物统计\生物统计课件整理.docx')

in_ch5 = False
for i, p in enumerate(doc.paragraphs):
    t = p.text
    if '5. 方差分析' in t:
        in_ch5 = True
    if in_ch5 and '本章习题' in t and p.style and p.style.name == 'Heading 3':
        break
    if in_ch5 and any(sym in t for sym in ['SST', 'SSt', 'SSe', '\u2211', 'CT', 'F=', 'df', 'St\u00b2', 'Se\u00b2']):
        print(f'[{i:3d}] {t[:120]}')
