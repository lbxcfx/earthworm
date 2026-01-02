from docx import Document
import os

book_dir = r"e:\earthworm\book"

for filename in sorted(os.listdir(book_dir)):
    if filename.endswith('.docx'):
        print(f"\n{'='*60}")
        print(f"文件: {filename}")
        print('='*60)
        
        doc = Document(os.path.join(book_dir, filename))
        
        # 读取前100段
        for i, para in enumerate(doc.paragraphs[:100]):
            text = para.text.strip()
            if text:
                print(text)
