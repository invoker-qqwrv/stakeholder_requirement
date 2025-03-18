import fitz
import re
import openpyxl

def extract_chapters_to_excel(pdf_path, output_excel):
    doc = fitz.open(pdf_path)
    text = ""

    # 读取所有页面文本
    for page in doc:
        text += page.get_text("text") + "\n"

    # 匹配章节标题（如 "1 Introduction", "2 Methods"）
    pattern = re.compile(r"(?P<title>^\d+\s+.*)", re.MULTILINE)

    # 查找所有匹配的章节标题
    matches = list(pattern.finditer(text))

    chapters = []
    for i in range(len(matches)):
        title = matches[i].group("title").strip()  # 获取章节标题
        start = matches[i].end()  # 章节起始位置
        end = matches[i + 1].start() if i + 1 < len(matches) else len(text)  # 下一个章节的开始

        # 提取章节内容，并合并标题 + 内容
        content = text[start:end].strip()
        full_chapter = f"{title}\n{content}"  # 标题和正文合并
        chapters.append([full_chapter])  # 作为单元格内容

    # 创建 Excel 工作簿
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Extracted Chapters"

    # 写入数据，每个章节独占一行
    for chapter in chapters:
        ws.append(chapter)

    # 保存 Excel
    wb.save(output_excel)
    print(f"提取完成，已保存到 {output_excel}")

# 运行代码（替换成你的 PDF 文件路径）
pdf_file = "汽车整车信息安全技术要求.pdf"
output_excel = "output_1.xlsx"
extract_chapters_to_excel(pdf_file, output_excel)

