# -*- coding: utf-8 -*-
"""
===================================================================================
Copyright (c) 2025, Yuxuan Zhou & Xiaolong Wang
Author: Yuxuan Zhou & Xiaolong Wang
All rights reserved.
===================================================================================
"""

from DocumentProcessor import DocumentProcessor
from RequirementExtractor import RequirementExtractor
from ExcelProcessor import ExcelProcessor

if __name__ == "__main__":
    #需要进行提取的文件的路径
    pdf_file = "汽车整车信息安全技术要求.pdf"

    #这个不用管，仅用于暂存中间数据
    output_excel = "output_2.xlsx"

    #输出到该excel表格中。请确保路径和文件名正确
    final_output_excel = "顾客要求跟踪矩阵_1.xlsx"

    #这里换成自己的api_key（具体咋换请查readme文档）。api_url不用换。默认使用deepseek-v3模型。也可升级r1。
    api_key = "sk-nwhqbtdjhizqhagulvtgcymarqpxiswaegwjszxyryojktzs"
    api_url = "https://api.siliconflow.cn/v1/chat/completions"

    #以下是顾客要求跟踪矩阵中我们要填的内容
    #file_source对应输入文档的名字
    file_source = "《汽车整车信息安全技术要求》"

    #文档版本
    file_version = "V1.0"

    #分析该文档的人的名字
    analyst = "马牛逼"

    # pdf文件预处理啊，提取pdf中的文字到excel中
    print("📄 正在提取 PDF 章节...")
    doc_processor = DocumentProcessor(pdf_file, output_excel, api_key, api_url)
    doc_processor.extract_chapters_to_excel()

    # 遍历excel，并喂入api进行需求提取和解析
    print("📝 正在处理 Excel 提取需求...")
    # with tqdm(total=100, desc="提取进度") as pbar:
    #     for _ in range(10):
    #         doc_processor.process_excel(output_excel)
    #         pbar.update(10)
    doc_processor.process_excel(output_excel)
    # 拆分提取内容，将章节号和相关方需求对应
    print("🔍 拆分章节号和需求内容...")
    extractor = RequirementExtractor(output_excel)
    #     with tqdm(total=100, desc="拆分进度") as pbar:
    #     extractor.split_titles_and_text()
    #     pbar.update(100)
    extractor.split_titles_and_text()
    # 填写顾客要求跟踪矩阵
    print("📊 填写顾客要求根总矩阵...")
    processor = ExcelProcessor(output_excel, final_output_excel, file_source, file_version, analyst)
    # with tqdm(total=100, desc="数据复制进度") as pbar:
    #     processor.copy_column_to_excel()
    #     pbar.update(100)
    processor.copy_column_to_excel()
    print("✅ Done! 顾客要求根总矩阵已生成！")
