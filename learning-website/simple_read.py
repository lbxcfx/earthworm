from docx import Document
import os

book_dir = r"e:\earthworm\book"

# 只处理Unit 1
filename = "《英语》（新标准）小学修订版五年级上册Unit 1教学设计.docx"
filepath = os.path.join(book_dir, filename)

doc = Document(filepath)

print("=" * 60)
print("UNIT 1 Content")
print("=" * 60)

# 获取所有文本
all_text = []
for para in doc.paragraphs:
    if para.text.strip():
        all_text.append(para.text.strip())

# 打印前50行
for i, text in enumerate(all_text[:50]):
    print(f"{i+1}: {text[:100]}")
