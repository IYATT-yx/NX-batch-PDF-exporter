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
        total = len(inputFiles)
        writer = PdfWriter()
        
        writeMsg("=" * 50)
        writeMsg(f"📚 开始合并任务 | 共计: {total} 个 PDF 文件")
        writeMsg("=" * 50)

        try:
            for idx, pdf in enumerate(inputFiles, 1):
                # 只获取文件名用于简洁显示
                short_name = os.path.basename(pdf)
                
                # 合并操作
                writer.append(pdf)
                
                # 打印进度条风格的消息
                writeMsg(f"[{idx}/{total}] 正在压入: {short_name}")

            # 写入文件
            writeMsg(f"\n[➔] 正在生成合并后的文件...")
            with open(outputFilename, "wb") as outputFileObj:
                writer.write(outputFileObj)
            writer.close()

            writeMsg("\n" + "=" * 50)
            writeMsg(f"✨ 合并成功！", 'success') # 成功级别
            writeMsg(f"   ✓ 最终文件: {os.path.basename(outputFilename)}", 'success')
            writeMsg(f"   📍 完整路径: {outputFilename}")
            writeMsg("=" * 50)

        except Exception as e:
            writeMsg("\n" + "=" * 50)
            writeMsg(f"❌ 合并过程中发生错误:", 'error') # 错误级别
            writeMsg(f"   {str(e)}", 'error')
            writeMsg("=" * 50)