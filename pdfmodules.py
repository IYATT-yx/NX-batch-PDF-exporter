import os
import sys
from typing import Callable

# 插入三方库路径
BASE_DIR = os.path.dirname(os.path.abspath(__file__))
pypdf_LIB_PATH = os.path.join(BASE_DIR, 'lib', 'pypdf-6.7.5')

if pypdf_LIB_PATH not in sys.path:
    sys.path.insert(0, pypdf_LIB_PATH)

from pypdf import PdfWriter

class PdfModules:
    @staticmethod
    def mergePdfs(outputFilename: str, inputFiles: list[str], writeMsg: Callable):
        """
        合并 PDF

        Args:
            outputFilename (str): 输出文件名
            inputFiles (list[str]): 输入文件列表
        """
        number = len(inputFiles)
        writer = PdfWriter()
        for idx, pdf in enumerate(inputFiles):
            writer.append(pdf)
            writeMsg(f"正在合并 PDF 文件，已合并 {idx + 1} / {number} 个文件")
        with open(outputFilename, "wb") as outputFileObj:
            writer.write(outputFileObj)
        writer.close()
        writeMsg(f"PDF 文件合并完成，共合并 {number} 个文件")
        writeMsg(f"合并后的 PDF 文件保存在 {outputFilename}")