import os
import shutil
import win32com.client
import pythoncom

def convert_to_pdf(docx_file, pdf_file):
    # 初始化 COM 环境
    pythoncom.CoInitialize()
    try:
        # 尝试调用 Dispatch("Word.Application")
        word = win32com.client.Dispatch("Word.Application")
        doc = word.Documents.Open(docx_file)
        # 将文件另存为 PDF 格式
        doc.SaveAs(pdf_file, FileFormat=17)
        doc.Close()
        print(f"转换成功: {docx_file} -> {pdf_file}")
        # 移动 PDF 文件到指定目录
        save_dir = "../reports"
        print(f"文件已保存到本地: {os.path.join(save_dir, os.path.basename(pdf_file))}")
    except Exception as e:
        print(f"转换失败: {docx_file} -> {pdf_file}. 错误信息: {e}")
    finally:
        # 释放 COM 环境资源
        pythoncom.CoUninitialize()

# # 示例用法
# if __name__ == "__main__":
#     input_docx = "example.docx"
#     output_pdf = "example.pdf"
#     convert_to_pdf(os.path.abspath(input_docx), os.path.abspath(output_pdf))
