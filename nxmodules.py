"""
file: nxmodules.py
description: NX 相关调用功能模块
author: IYATT-yx
repository: https://github.com/IYATT-yx/NX-batch-PDF-exporter
copyright:  Copyright (c) 2026 IYATT-yx.
            Licensed under the MIT License. See LICENSE file in the project root for full license information.
"""
from tkinter import filedialog
import tkinter as tk
import NXOpen

from typing import Callable
import os
from datetime import datetime

class NxModules:
    ImageResolution = {
        _('高'): NXOpen.PrintPDFBuilder.ImageResolutionOption.High,
        _('中'): NXOpen.PrintPDFBuilder.ImageResolutionOption.Medium,
        _('低'): NXOpen.PrintPDFBuilder.ImageResolutionOption.Low,
        _('草稿'): NXOpen.PrintPDFBuilder.ImageResolutionOption.Draft
    }

    Color = {
        _('按显示'): NXOpen.PrintPDFBuilder.Color.AsDisplayed,
        _('部件颜色'): NXOpen.PrintPDFBuilder.Color.PartColors,
        _('定制调色板'): NXOpen.PrintPDFBuilder.Color.CustomPalette,
        _('白纸黑字'): NXOpen.PrintPDFBuilder.Color.BlackOnWhite,
        _('原有颜色'): NXOpen.PrintPDFBuilder.Color.LegacyColors,
        _('按宽度定色'): NXOpen.PrintPDFBuilder.Color.ColorsByWidth
    }

    Width = {
        _('标准宽度'): NXOpen.PrintPDFBuilder.Width.StandardWidths,
        _('单线宽度'): NXOpen.PrintPDFBuilder.Width.SingleWidth,
        _('定制宽度'): NXOpen.PrintPDFBuilder.Width.CustomThreeWidths,
        _('定制调色板'): NXOpen.PrintPDFBuilder.Width.CustomPalette
    }

    SizeOption = {
        _('实际大小'): NXOpen.PrintPDFBuilder.SizeOption.FullScale,
        _('比例因子'): NXOpen.PrintPDFBuilder.SizeOption.ScaleFactor,
        _('尺寸'): NXOpen.PrintPDFBuilder.SizeOption.Dimension
    }

    OutputTextOption = {
        _('文本'): NXOpen.PrintPDFBuilder.OutputTextOption.Text,
        _('多段线'): NXOpen.PrintPDFBuilder.OutputTextOption.Polylines
    }

    UnitsOption = {
        _('毫米'): NXOpen.PrintPDFBuilder.UnitsOption.Metric,
        _('英寸'): NXOpen.PrintPDFBuilder.UnitsOption.English
    }

    class PdfSettings:
        def __init__(self, color, width, size, scale: float, units, xDimension: float, yDimension: float, outputText, addWatermark: bool, watermark: str, rasterImages: bool, shadedGeometry: bool, customSymbolsInForeground: bool, imageResolution):
            '''
            PDF 导出设置参数

            Args:
                color: 颜色
                width: 宽度
                size: 大小
                scale (float): 比例因子
                units: 单位
                xDimension (float): X 尺寸
                yDimension (float): Y 尺寸
                outputText: 输出文本
                addWatermark (bool): 是否添加水印
                watermark (str): 水印文本
                rasterImages (bool): 在图纸中导出光栅图像
                shadedGeometry (bool): 将着色的几何体导出为线框
                customSymbolsInForeground (bool): 在前景中导出定制符号
                imageResolution: 图像分辨率
            '''
            self.color = color
            self.width = width
            self.size = size
            self.scale = scale
            self.units = units
            self.xDimension = xDimension
            self.yDimension = yDimension
            self.outputText = outputText
            self.addWatermark = addWatermark
            if addWatermark:
                self.watermark = watermark
            else:
                self.watermark = ''
            self.rasterImages = rasterImages
            self.shadedGeometry = shadedGeometry
            self.customSymbolsInForeground = customSymbolsInForeground
            self.imageResolution = imageResolution

        def get(self):
            '''
            读取 PDF 导出设置参数

            Returns:
                tuple: PDF 导出设置参数( color, width, size, scale, units, xDimension, yDimension, outputText, addWatermark, watermark, rasterImages, shadedGeometry, customSymbolsInForeground, imageResolution )
            '''
            return self.color, self.width, self.size, self.scale, self.units, self.xDimension, self.yDimension, self.outputText, self.addWatermark, self.watermark, self.rasterImages, self.shadedGeometry, self.customSymbolsInForeground, self.imageResolution

    @staticmethod
    def selectPrts() -> list[str]:
        """
        选择 NX 文件

        Retunrs:
            list[str]: NX 文件路径列表
        """
        prts = filedialog.askopenfilenames(
            title=_('选择 NX 文件'),
            filetypes=[(_('NX 文件'), '*.prt')]
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
        folder = filedialog.askdirectory(title=_('选择包含 NX 文件的文件夹'))
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
            msg: str = _('获取当前工作部件失败: {e}')
            NxModules.printMsg(msg.format(e=str(e)))
            return []

        if workPrt is None:
            NxModules.printMsg(_('当前没有工作部件'))
            return []

        return [workPrt.FullPath]
    
    @staticmethod
    def foreachPrt(func, prtList: list[str], prefixName: str, suffixName: str, folder: str, writeMsg: Callable, pdfSettings: PdfSettings=None):
        """
        对每个 NX 文件执行指定函数

        Args:
            func (function): 要执行的函数
            prtList (list[str]): NX 文件路径列表
            prefixName (str): 文件名前缀
            suffixName (str): 文件名后缀
            folder (str): 文件保存路径
            writeMsg (function): 用于输出消息的函数
            pdfSettings (PdfSettings): PDF 导出设置
        """
        session = NXOpen.Session.GetSession()
        total = len(prtList)
        counter = 0
        
        writeMsg("=" * 50)
        msg: str = _('🚀 开始批量任务 | 共计: {total} 个文件')
        writeMsg(msg.format(total=total))
        writeMsg("=" * 50)

        for index, prtPath in enumerate(prtList, 1):
            short_name = os.path.basename(prtPath)
            msg: str = _('\n[{index}/{total}] 正在处理: {short_name}')
            writeMsg(msg.format(index=index, total=total, short_name=short_name))
            
            closePart = True
            try:
                part, status = session.Parts.OpenActiveDisplay(prtPath, NXOpen.DisplayPartOption.AllowAdditional)
            except NXOpen.NXException as e:
                if _('文件已存在') in str(e):
                    closePart = False
                    for p in session.Parts:
                        if p and p.FullPath == prtPath:
                            part = p
                            session.Parts.SetActiveDisplay(part, NXOpen.DisplayPartOption.AllowAdditional, NXOpen.PartDisplayPartWorkPartOption.UseLast)
                            break
                    msg: str = _('  [!] 提示: {short_name} 已经在内存中')
                    writeMsg(msg.format(short_name=short_name), 'warn') # 警告色
                else:
                    msg: str = _('  [✘] 错误: 无法打开文件 - {e}')
                    writeMsg(msg.format(e=e), 'error') # 错误色
                    continue

            # 执行导出函数
            if pdfSettings is not None:
                if func(part, prefixName, suffixName, folder, writeMsg, pdfSettings):
                    counter += 1
            else:
                if func(part, prefixName, suffixName, folder, writeMsg):
                    counter += 1

            if closePart:
                part.Close(NXOpen.BasePart.CloseWholeTree.FalseValue, NXOpen.BasePart.CloseModified.CloseModified, None)
                status.Dispose()
                msg: str = _('  [➔] 资源释放: {short_name} 已关闭')
                writeMsg(msg.format(short_name=short_name))

        writeMsg("\n" + "=" * 50)
        # 根据完成情况输出不同颜色
        if counter == total:
            level = 'success'
        else:
            level = 'warn'
        msg = _('✨ 任务完成！成功导出: {counter} / {total}')
        writeMsg(msg.format(counter=counter, total=total), level)
        writeMsg("=" * 50)

    @staticmethod
    def exportPdf(part: NXOpen.Part, prefixName: str, suffixName: str, folder: str, writeMsg: Callable, pdfSettings: PdfSettings):
        """
        导出 PDF

        Args:
            part (NXOpen.Part): NX 部件
            prefixName (str): 文件名前缀
            suffixName (str): 文件名后缀
            folder (str): 文件保存路径
            writeMsg (function): 用于输出消息的函数
            pdfSettings (PdfSettings): PDF 导出设置
        
        Returns:
            bool: 导出是否成功
        """
        try:
            pdfbuilder = part.PlotManager.CreatePrintPdfbuilder()
            target_folder = folder if folder else os.path.dirname(part.FullPath)
            filename = os.path.join(target_folder, f"{prefixName}{part.Leaf}{suffixName}.pdf")
            
            sheets = [sheet for sheet in part.DrawingSheets]
            if not sheets:
                writeMsg(_('  [!] 跳过: 该部件未发现任何图纸页'), 'warn') # 警告色
                return False
            
            for sheet in sheets:
                sheet.Open()

            pdfSettingsObj = pdfSettings.get()

            pdfbuilder.Colors = pdfSettingsObj[0]
            pdfbuilder.Widths = pdfSettingsObj[1]
            pdfbuilder.Size = pdfSettingsObj[2]
            pdfbuilder.Scale = pdfSettingsObj[3]
            pdfbuilder.Units = pdfSettingsObj[4]
            pdfbuilder.XDimension = pdfSettingsObj[5]
            pdfbuilder.YDimension = pdfSettingsObj[6]
            pdfbuilder.OutputText = pdfSettingsObj[7]
            pdfbuilder.AddWatermark = pdfSettingsObj[8]
            pdfbuilder.Watermark = pdfSettingsObj[9]
            pdfbuilder.RasterImages = pdfSettingsObj[10]
            pdfbuilder.ShadedGeometry = pdfSettingsObj[11]
            pdfbuilder.CustomSymbolsInForeground = pdfSettingsObj[12]
            pdfbuilder.ImageResolution = pdfSettingsObj[13]
                                                        
            pdfbuilder.Filename = filename # 输出文件名
            pdfbuilder.SourceBuilder.SetSheets(sheets) # 图纸

            pdfbuilder.Commit()
            pdfbuilder.Destroy()

            writeMsg(_('  [✓] PDF 导出成功'), 'success') # 成功色
            msg: str = _('      📍 路径: {filename}')
            writeMsg(msg.format(filename=filename))
            return True
        except Exception as e:
            msg: str = _('  [✘] PDF 导出异常: {e}')
            writeMsg(msg.format(e=str(e)), 'error') # 错误色
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
                writeMsg(_('  [!] 跳过: 该部件未发现任何图纸页'), 'warn') # 警告色
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

            msg: str = _('  [✓] DWG 导出成功 ({len_of_sheets} 张图纸)')
            writeMsg(msg.format(len_of_sheets=len(sheets)), 'success') # 成功色
            writeMsg(f"      📍 路径: {filename}")
            return True 

        except Exception as e:
            msg: str = _('  [✘] DWG 导出异常: {e}')
            writeMsg(msg.format(e=str(e)), 'error') # 错误色
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
