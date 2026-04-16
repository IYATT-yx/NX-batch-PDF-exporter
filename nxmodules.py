from tkinter import filedialog
import tkinter as tk
import NXOpen

from typing import Callable
import os
from datetime import datetime

class NxModules:
    @staticmethod
    def selectPrts() -> list[str]:
        """
        选择 NX 文件

        Retunrs:
            list[str]: NX 文件路径列表
        """
        prts = filedialog.askopenfilenames(
            title='选择 NX 文件',
            filetypes=[('NX 文件', '*.prt')]
        )
        return prts
    
    @staticmethod
    def getOpenedPrts() -> list[str]:
        """
        获取已打开的 NX 文件

        Returns:
            list[str]: NX 文件路径列表
        """
        session = NXOpen.Session.GetSession()
        return [prt.FullPath for prt in session.Parts]
    
    @staticmethod
    def getPrtsFromSelectedFolder(recursive: bool = False) -> list[str]:
        """
        从指定文件夹获取 NX 文件

        Args:
            recursive (bool, optional): 是否递归获取子文件夹中的文件。默认不递归。

        Returns:
            list[str]: NX 文件路径列表
        """
        folder = filedialog.askdirectory(title='选择包含 NX 文件的文件夹')
        prts = []

        if recursive:
            for root, dirs, files in os.walk(folder):
                for file in files:
                    if file.lower().endswith('.prt'):
                        prts.append(os.path.join(root, file))
        else:
            for file in os.listdir(folder):
                if file.lower().endswith('.prt'):
                    prts.append(os.path.join(folder, file))

        return prts
    
    @staticmethod
    def getWorkPrt() -> list[str]:
        """
        获取当前工作部件

        Returns:
            list[str]: NX 文件路径列表
        """
        try:
            workPrt = NXOpen.Session.GetSession().Parts.Work
        except NXOpen.NXException as e:
            NxModules.printMsg(f'获取当前工作部件失败: {e}')
            return []

        if workPrt is None:
            NxModules.printMsg('当前没有工作部件')
            return []

        return [workPrt.FullPath]
    
    @staticmethod
    def foreachPrt(func, prtList: list[str], prefixName: str, suffixName: str, folder: str, writeMsg: Callable):
        """
        对每个 NX 文件执行指定函数

        Args:
            func (function): 要执行的函数
            prtList (list[str]): NX 文件路径列表
            prefixName (str): 文件名前缀
            suffixName (str): 文件名后缀
            folder (str): 文件保存路径
            writeMsg (function): 用于输出消息的函数
        """
        session = NXOpen.Session.GetSession()
        total = len(prtList)
        counter = 0
        timestamp = datetime.now().strftime("_%Y%m%d_%H%M%S")
        
        writeMsg("=" * 50)
        writeMsg(f"🚀 开始批量任务 | 共计: {total} 个文件") # 默认 info
        writeMsg("=" * 50)

        for index, prtPath in enumerate(prtList, 1):
            short_name = os.path.basename(prtPath)
            writeMsg(f"\n[{index}/{total}] 正在处理: {short_name}")
            
            closePart = True
            try:
                part, status = session.Parts.OpenActiveDisplay(prtPath, NXOpen.DisplayPartOption.AllowAdditional)
            except NXOpen.NXException as e:
                if '文件已存在' in str(e):
                    closePart = False
                    for p in session.Parts:
                        if p and p.FullPath == prtPath:
                            part = p
                            session.Parts.SetActiveDisplay(part, NXOpen.DisplayPartOption.AllowAdditional, NXOpen.PartDisplayPartWorkPartOption.UseLast)
                            break
                    writeMsg(f"  [!] 提示: {short_name} 已经在内存中", 'warn') # 警告色
                else:
                    writeMsg(f"  [✘] 错误: 无法打开文件 - {e}", 'error') # 错误色
                    continue

            # 执行导出函数
            if func(part, prefixName, suffixName + timestamp, folder, writeMsg):
                counter += 1

            if closePart:
                part.Close(NXOpen.BasePart.CloseWholeTree.FalseValue, NXOpen.BasePart.CloseModified.CloseModified, None)
                status.Dispose()
                writeMsg(f"  [➔] 资源释放: {short_name} 已关闭")

        writeMsg("\n" + "=" * 50)
        # 根据完成情况输出不同颜色
        if counter == total:
            writeMsg(f"✨ 任务完成！成功导出: {counter} / {total}", 'success')
        else:
            writeMsg(f"✨ 任务完成！成功导出: {counter} / {total}", 'warn')
        writeMsg("=" * 50)

    @staticmethod
    def exportPdf(part: NXOpen.Part, prefixName: str, suffixName: str, folder: str, writeMsg: Callable):
        """
        导出 PDF

        Args:
            part (NXOpen.Part): NX 部件
            prefixName (str): 文件名前缀
            suffixName (str): 文件名后缀
            folder (str): 文件保存路径
            writeMsg (function): 用于输出消息的函数
        
        Returns:
            bool: 导出是否成功
        """
        try:
            pdfbuilder = part.PlotManager.CreatePrintPdfbuilder()
            target_folder = folder if folder else os.path.dirname(part.FullPath)
            filename = os.path.join(target_folder, f"{prefixName}{part.Leaf}{suffixName}.pdf")
            
            sheets = [sheet for sheet in part.DrawingSheets]
            if not sheets:
                writeMsg(f"  [!] 跳过: 该部件未发现任何图纸页", 'warn') # 警告色
                return False
            
            for sheet in sheets:
                sheet.Open()
            
            pdfbuilder.Filename = filename
            pdfbuilder.SourceBuilder.SetSheets(sheets)
            
            pdfbuilder.Commit()
            pdfbuilder.Destroy()

            writeMsg(f"  [✓] PDF 导出成功", 'success') # 成功色
            writeMsg(f"      📍 路径: {filename}")
            return True
        except Exception as e:
            writeMsg(f"  [✘] PDF 导出异常: {str(e)}", 'error') # 错误色
            return False
            
    @staticmethod
    def exportDwg(part: NXOpen.Part, prefixName: str, suffixName: str, folder: str, writeMsg: Callable):
        """
        导出 AutoCAD DWG 图纸

        Args:
            part (NXOpen.Part): NX 部件
            prefixName (str): 文件名前缀
            suffixName (str): 文件名后缀
            folder (str): 文件保存路径
            writeMsg (function): 用于输出消息的函数
        
        Returns:
            bool: 导出是否成功
        """
        try:
            theSession = NXOpen.Session.GetSession()
            sheets = list(part.DrawingSheets)
            
            if not sheets:
                writeMsg(f"  [!] 跳过: 该部件未发现任何图纸页", 'warn') # 警告色
                return False

            dxfdwgCreator = theSession.DexManager.CreateDxfdwgCreator()
            
            target_folder = folder if folder else os.path.dirname(part.FullPath)
            filename = os.path.join(target_folder, f"{prefixName}{part.Leaf}{suffixName}.dwg")
            
            dxfdwgCreator.InputFile = part.FullPath
            dxfdwgCreator.OutputFile = filename
            dxfdwgCreator.ExportData = NXOpen.DxfdwgCreator.ExportDataOption.Drawing
            dxfdwgCreator.AutoCADRevision = NXOpen.DxfdwgCreator.AutoCADRevisionOptions.R2004
            dxfdwgCreator.OutputFileType = NXOpen.DxfdwgCreator.OutputFileTypeOption.Dwg
            dxfdwgCreator.DrawingList = ",".join([f'"{s.Name}"' for s in sheets])

            dxfdwgCreator.Commit()
            dxfdwgCreator.Destroy()

            writeMsg(f"  [✓] DWG 导出成功 ({len(sheets)} 张图纸)", 'success') # 成功色
            writeMsg(f"      📍 路径: {filename}")
            return True 

        except Exception as e:
            writeMsg(f"  [✘] DWG 导出异常: {str(e)}", 'error') # 错误色
            return False

    @staticmethod
    def setFileReadOnly(filePath: str, readonly: bool = True):
        """
        设置文件只读属性

        Args:
            filePath(str): 文件路径
            readonly(bool): True 设置只读，False 取消只读
        """
        if readonly:
            os.chmod(filePath, 0o444)
        else:
            os.chmod(filePath, 0o777)