# import fitz
# import re
#
# def extract_sections(pdf_path):
#     doc = fitz.open(pdf_path)
#     text = ""
#
#     # 读取所有页面文本
#     for page in doc:
#         text += page.get_text("text") + "\n"
#
#     # 通过正则表达式匹配小标题（适用于 "1. xxx", "1.1 xxx", "Conclusion" 等）
#     pattern = re.compile(r"(?P<title>\b\d+(\.\d+)*\s+.*|Conclusion|References)\n", re.MULTILINE)
#
#     # 用正则找到所有小标题的位置
#     matches = list(pattern.finditer(text))
#
#     sections = {}
#     for i in range(len(matches)):
#         title = matches[i].group("title").strip()
#         start = matches[i].end()  # 获取标题后的位置
#         end = matches[i + 1].start() if i + 1 < len(matches) else len(text)  # 下一个标题的开始
#
#         # 提取该章节内容
#         content = text[start:end].strip()
#         sections[title] = content
#
#     # 输出每个章节的内容
#     for title, content in sections.items():
#         print(f"\n=== {title} ===\n{content}\n{'-'*40}")
#
# # 运行代码（替换成你的 PDF 文件路径）
# pdf_file = "汽车整车信息安全技术要求.pdf"
# extract_sections(pdf_file)




import fitz
import re
import openpyxl

def extract_sections_to_excel(pdf_path, output_excel):
    doc = fitz.open(pdf_path)
    text = ""

    # 读取所有页面文本
    for page in doc:
        text += page.get_text("text") + "\n"

    # 通过正则表达式匹配小标题（适用于 "1. xxx", "1.1 xxx", "Conclusion" 等）
    pattern = re.compile(r"(?P<title>\b\d+(\.\d+)*\s+.*|Conclusion|References)\n", re.MULTILINE)

    # 用正则找到所有小标题的位置
    matches = list(pattern.finditer(text))

    sections = []
    for i in range(len(matches)):
        title = matches[i].group("title").strip()
        start = matches[i].end()  # 获取标题后的位置
        end = matches[i + 1].start() if i + 1 < len(matches) else len(text)  # 下一个标题的开始

        # 提取该章节内容，并合并标题 + 内容
        content = text[start:end].strip()
        full_section = f"{title}\n{content}"  # 标题和内容合并
        sections.append([full_section])  # 作为一个单独的 Excel 单元格内容

    # 创建 Excel 工作簿
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Extracted Sections"

    # 写入数据，每个小标题和内容放在一个单元格
    for section in sections:
        ws.append(section)

    # 保存 Excel
    wb.save(output_excel)
    print(f"提取完成，已保存到 {output_excel}")

# 运行代码（替换成你的 PDF 文件路径）
pdf_file = "汽车整车信息安全技术要求.pdf"
output_excel = "output.xlsx"
extract_sections_to_excel(pdf_file, output_excel)
