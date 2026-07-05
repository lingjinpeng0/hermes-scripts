#!/usr/bin/env python3
import sys
import fitz
import easyocr

pdf_path = sys.argv[1]
out_path = pdf_path.replace('.pdf', '_ocr.txt')

doc = fitz.open(pdf_path)
print(f"总页数: {doc.page_count}", flush=True)

reader = easyocr.Reader(['ch_sim', 'en'], gpu=False, verbose=False)

all_text = []
for i in range(doc.page_count):
    page = doc[i]
    pix = page.get_pixmap(dpi=200)
    import numpy as np
    img = np.frombuffer(pix.samples, dtype=np.uint8).reshape(pix.height, pix.width, pix.n)
    if pix.n == 4:
        img = img[:,:,:3]
    
    result = reader.readtext(img)
    if result:
        lines = [item[1] for item in result]
        all_text.append(f"--- 第{i+1}页 ---\n" + "\n".join(lines))
    else:
        all_text.append(f"--- 第{i+1}页 ---\n(无文字)")
    
    if (i+1) % 5 == 0:
        print(f"已处理 {i+1}/{doc.page_count} 页...", flush=True)

doc.close()
output = "\n\n".join(all_text)
with open(out_path, 'w', encoding='utf-8') as f:
    f.write(output)
print(f"OCR完成! 输出: {out_path}", flush=True)
print(f"总字数: {len(output)}", flush=True)
