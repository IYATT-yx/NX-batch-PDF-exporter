from tkinter import filedialog
import tkinter as tk
import NXOpen

from typing import Callable
import os

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
    def printMsg(msg: str):
        """
        在列表窗口中打印消息
        """
        session = NXOpen.Session.GetSession()
        lw = session.ListingWindow
        if not lw.IsOpen():
            lw.Open()
        lw.WriteLine(msg)
    
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
        part: NXOpen.Part = None
        counter = 0
        for prtPath in prtList:
            closePart: bool = True
            try:
                part, status = session.Parts.OpenActiveDisplay(prtPath, NXOpen.DisplayPartOption.AllowAdditional)
            except NXOpen.NXException as e:
                if '文件已存在' in str(e):
                    closePart = False
                    part = session.Parts.Display
                    writeMsg(f'{part.Name} 是已打开的文件')
                else:
                    writeMsg(f'打开文件 {prtPath} 失败: {e}')
                    continue
            else:
                writeMsg(f'成功打开文件 {prtPath}')
            if func(part, prefixName, suffixName, folder, writeMsg):
                counter += 1

            if closePart:
                part.Close(NXOpen.BasePart.CloseWholeTree.TrueValue, NXOpen.BasePart.CloseModified.CloseModified, None)
                status.Dispose()
                writeMsg(f'⭙ 关闭文件 {prtPath}')

        writeMsg(f'✅ 共导出 {counter} 个文件')

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
            if folder == '':
                folder = os.path.dirname(part.FullPath)
            filename = os.path.join(folder, prefixName + part.Name + suffixName + '.pdf')
            pdfbuilder.Filename = filename
            sheets = [sheet for sheet in part.DrawingSheets]
            if len(sheets) == 0:
                writeMsg(f'❗{part.Name} 没有图纸')
                return
            pdfbuilder.SourceBuilder.SetSheets(sheets)
            writeMsg(f'开始导出 {filename}')
            pdfbuilder.Commit()
            pdfbuilder.Destroy()
        except NXOpen.NXException as e:
            writeMsg(f'❌ 导出部件 {part.FullPath} 失败: {e}')
            return False
        else:
            writeMsg(f'✓ 导出 PDF 成功: {filename}，共 {len(sheets)} 张图纸')
            return True
            
            
