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
        msg: str = _('📚 开始合并任务 | 共计: {total} 个 PDF 文件')
        writeMsg(msg.format(total=total))
        writeMsg("=" * 50)

        try:
            for idx, pdf in enumerate(inputFiles, 1):
                short_name = os.path.basename(pdf)
                writer.append(pdf)
                msg: str = _('[{idx}/{total}] 正在压入: {short_name}')
                writeMsg(msg.format(idx=idx, total=total, short_name=short_name))

            # 写入文件
            writeMsg(_('\n[➔] 正在生成合并后的文件...'))
            with open(outputFilename, "wb") as outputFileObj:
                writer.write(outputFileObj)
            writer.close()

            writeMsg("\n" + "=" * 50)
            writeMsg(_('✨ 合并成功！'), 'success') # 成功级别
            msg: str = _('   ✓ 最终文件: {outputFilename}')
            writeMsg(msg.format(outputFilename=os.path.basename(outputFilename)), 'success')
            msg: str = _('   📍 完整路径: {outputFilename}')
            writeMsg(msg.format(outputFilename=outputFilename))
            writeMsg("=" * 50)

        except Exception as e:
            writeMsg("\n" + "=" * 50)
            writeMsg(_('❌ 合并过程中发生错误:'), 'error') # 错误级别
            writeMsg(f"   {str(e)}", 'error')
            writeMsg("=" * 50)