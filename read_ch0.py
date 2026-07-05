"""Read ch0 content from the generated docx, paragraph by paragraph."""
from docx import Document

doc = Document(r'C:\Users\Rei\Downloads\学习\生物统计\生物统计课件整理.docx')

# Read from start through end of ch0
print("=== 第0章 完整通读 ===")
for i, p in enumerate(doc.paragraphs):
    if p.text.strip():
        style = p.style.name if p.style else 'None'
        print(f'[{i:3d}] {p.text}')
print()
