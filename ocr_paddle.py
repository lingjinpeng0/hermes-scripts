#!/usr/bin/env python3
import os, sys
os.environ['DNNL_DEFAULT_FPMATH_MODE'] = 'BF16'

import fitz
from paddleocr import PaddleOCR

pdf_path = sys.argv[1]
out_path = pdf_path.replace('.pdf', '_ocr.txt')

doc = fitz.open(pdf_path)
print(f"总页数: {doc.page_count}", flush=True)

ocr = PaddleOCR(use_angle_cls=True, lang='ch', show_log=False)

import numpy as np

all_text = []
for i in range(doc.page_count):
    page = doc[i]
    pix = page.get_pixmap(dpi=200)
    # 转numpy array
    img = np.frombuffer(pix.samples, dtype=np.uint8).reshape(pix.height, pix.width, pix.n)
    if pix.n == 4:  # RGBA → RGB
        img = img[:,:,:3]
    
    result = ocr.ocr(img, cls=True)
    if result and result[0]:
        lines = []
        for line in result[0]:
            text = line[1][0]
            lines.append(text)
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
