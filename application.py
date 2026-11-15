"""
file: application.py
description: 图形界面
author: IYATT-yx
repository: https://github.com/IYATT-yx/NX-batch-PDF-exporter
copyright:  Copyright (c) 2026 IYATT-yx.
            Licensed under the MIT License. See LICENSE file in the project root for full license information.
"""
from listshow import ListShow
from nxmodules import NxModules
import version
from pdfmodules import PdfModules

import tkinter as tk
from tkinter import ttk, filedialog, messagebox
import os
from enum import Enum, auto

programName = _('Siemens NX/Designcenter 批量 PDF 图纸导出工具')
baseDir = os.path.dirname(os.path.abspath(__file__))
tclDir = os.path.join(baseDir, 'tcl', 'tcl8.6')
os.environ["TCL_LIBRARY"] = tclDir

class Application(tk.Frame):
    class DrawingClass(Enum):
        PDF = auto()
        DWG = auto()

    def __init__(self, root: tk.Tk):
        super().__init__(root)
        self.root = root
        self.pack(fill=tk.BOTH, expand=True)

        self.msgText = None
        self.pdfSettingObj = None

        self.createWidgets()
        msg: str = _('{programName} {version}\nIYATT-yx iyatt@iyatt.com\n本工具为开源项目，发布地址：https://github.com/IYATT-yx/NX-batch-PDF-exporter')
        self.writeMsg(msg.format(programName=programName, version=version.V))

    def createWidgets(self):
        """
        创建组件
        """
        # 顶部：Notebook 区域（设置权重大，优先拉伸）
        self.notebook = ttk.Notebook(self)
        self.notebook.pack(side=tk.TOP, expand=True, fill=tk.BOTH)

        self.exportPdfTab = ttk.Frame(self.notebook)
        self.createExportPdfTab(self.exportPdfTab)
        self.notebook.add(self.exportPdfTab, text=_('导出图纸'))

        self.mergePdfTab = ttk.Frame(self.notebook)
        self.createMergePdfTab(self.mergePdfTab)
        self.notebook.add(self.mergePdfTab, text=_('合并PDF'))

        # 底部：公用日志输出框（固定在最下方）
        self.msgText = tk.Text(self, wrap=tk.CHAR, height=16, state=tk.DISABLED)
        self.msgText.pack(side=tk.BOTTOM, fill=tk.X, padx=5, pady=5)
        self._config_text_tags(self.msgText)

    def createExportPdfTab(self, parent: tk.Frame):
        """
        创建导出 PDF 页面

        Args:
            parent (tk.Frame): 父组件
        """
        parent.rowconfigure([0,1,2,3,4,5,6], weight=1)
        parent.columnconfigure(0, weight=1)

        self.prtListShow = ListShow(parent)
        self.prtListShow.grid(row=0, column=0, rowspan=8, sticky=tk.NSEW)

        tk.Button(parent, text=_('选择文件'), command=self.onGetSelectedPrts).grid(row=0, column=1, sticky=tk.NSEW)
        tk.Button(parent, text=_('选择文件夹'), command=self.onGetPrtsFromSelectedFolder).grid(row=1, column=1, sticky=tk.NSEW)
        tk.Button(parent, text=_('所有打开文件'), command=self.onGetOpenedPrts).grid(row=2, column=1, sticky=tk.NSEW)
        tk.Button(parent, text=_('工作部件'), command=self.onWorkPrt).grid(row=3, column=1, sticky=tk.NSEW)
        tk.Button(parent, text=_('清空显示'), command=self.prtListShow.clear).grid(row=4, column=1, sticky=tk.NSEW)
        ttk.Separator(parent, orient=tk.HORIZONTAL).grid(row=5, column=1, columnspan=2, sticky=tk.EW, pady=5)

        tk.Button(parent, text=_('设置选中项'), command=self.onSetSelectedPrtStatus).grid(row=6, column=1, sticky=tk.NSEW)
        tk.Button(parent, text=_('设置所有'), command=self.onSetAllPrtStatus).grid(row=7, column=1, sticky=tk.NSEW)
        
        self.recursiveValue = tk.BooleanVar(parent, value=False)
        tk.Checkbutton(parent, text=_('递归'), variable=self.recursiveValue).grid(row=1, column=2, sticky=tk.W)
        self.selectedPrtStatusValue = tk.BooleanVar(parent, value=True)
        tk.Checkbutton(parent, text=_('导出'), variable=self.selectedPrtStatusValue).grid(row=6, column=2, sticky=tk.W)
        self.allPrtStatusValue = tk.BooleanVar(parent, value=True)
        tk.Checkbutton(parent, text=_('导出'), variable=self.allPrtStatusValue).grid(row=7, column=2, sticky=tk.W)

        ttk.Separator(parent, orient=tk.HORIZONTAL).grid(row=8, column=0, columnspan=3, sticky=tk.EW, pady=5)

        self.exportFolderValue = tk.StringVar(parent, value='')
        tk.Entry(parent, textvariable=self.exportFolderValue, bd=3).grid(row=9, column=0, sticky=tk.NSEW)
        tk.Button(parent, text=_('选择导出文件夹'), command=self.onSelectExportFolder).grid(row=9, column=1, sticky=tk.NSEW)
        tk.Label(parent, text=_('（可选）')).grid(row=9, column=2, sticky=tk.W)

        self.prefixNameValue = tk.StringVar(parent, value='')
        tk.Entry(parent, textvariable=self.prefixNameValue,  bd=3).grid(row=10, column=0, sticky=tk.NSEW)
        tk.Label(parent, text=_('<--- 导出文件名前缀（可选）')).grid(row=10, column=1, columnspan=2, sticky=tk.W)

        self.suffixNameValue = tk.StringVar(parent, value='')
        tk.Entry(parent, textvariable=self.suffixNameValue, bd=3).grid(row=11, column=0, sticky=tk.NSEW)
        tk.Label(parent, text=_('<--- 导出文件名后缀（可选）')).grid(row=11, column=1, columnspan=2, sticky=tk.W)

        exportFormatBox = tk.LabelFrame(parent, text=_('导出格式'), padx=10, pady=10)
        exportFormatBox.grid(row=12, column=0, columnspan=3, sticky=tk.NSEW)
        for c in range(3):
            exportFormatBox.columnconfigure(c, weight=1)
        self.drawingClassValue = tk.IntVar(parent, self.DrawingClass.PDF.value)
        tk.Radiobutton(exportFormatBox, text='PDF', variable=self.drawingClassValue, value=self.DrawingClass.PDF.value, command=self.onChangeDrawingClass).grid(row=0, column=1, sticky=tk.W)
        tk.Radiobutton(exportFormatBox, text='DWG', variable=self.drawingClassValue, value=self.DrawingClass.DWG.value, command=self.onChangeDrawingClass).grid(row=0, column=2, sticky=tk.W)

        # PDF 导出设置
        self.pdfSettingBox = tk.LabelFrame(parent, text=_('PDF 导出设置'), padx=10, pady=10)
        self.pdfSettingBox.grid(row=13, column=0, columnspan=6, sticky=tk.NSEW)
        for c in range(6):
            self.pdfSettingBox.columnconfigure(c, weight=1)

        # 打印属性 - 颜色
        tk.Label(self.pdfSettingBox, text=_('颜色：')).grid(row=0, column=0, sticky=tk.E)
        self.colorCombox = ttk.Combobox(self.pdfSettingBox, values=list(NxModules.Color.keys()), state='readonly')
        self.colorCombox.grid(row=0, column=1, sticky=tk.W)
        self.colorCombox.set(list(NxModules.Color.keys())[0]) # 按显示

        # 打印属性 - 宽度
        tk.Label(self.pdfSettingBox, text=_('宽度：')).grid(row=0, column=2, sticky=tk.E)
        self.widthCombox = ttk.Combobox(self.pdfSettingBox, values=list(NxModules.Width.keys()), state='readonly')
        self.widthCombox.grid(row=0, column=3, sticky=tk.W)
        self.widthCombox.set(list(NxModules.Width.keys())[0]) # 标准宽度

        # 打印属性 - 大小
        tk.Label(self.pdfSettingBox, text=_('大小：')).grid(row=0, column=4, sticky=tk.E)
        self.sizeOptionCombox = ttk.Combobox(self.pdfSettingBox, values=list(NxModules.SizeOption.keys()), state='readonly')
        self.sizeOptionCombox.grid(row=0, column=5, sticky=tk.W)
        self.sizeOptionCombox.bind("<<ComboboxSelected>>", self.onChangeSizeOption)
        self.sizeOptionCombox.set(list(NxModules.SizeOption.keys())[1]) # 比例因子

        # 打印属性 - 比例因子
        self.scaleFactorFrame = tk.LabelFrame(self.pdfSettingBox, text=_('比例因子'))
        self.scaleFactorFrame.grid(row=1, column=0, columnspan=6, sticky=tk.NSEW)
        for c in range(6):
            self.scaleFactorFrame.columnconfigure(c, weight=1)
        tk.Label(self.scaleFactorFrame, text=_('比例因子：')).grid(row=0, column=4, sticky=tk.E)
        self.scaleFactorValue = tk.DoubleVar(parent, value='3.0')
        tk.Entry(self.scaleFactorFrame, textvariable=self.scaleFactorValue, bd=3).grid(row=0, column=5, sticky=tk.W)

        # 打印属性 - 尺寸
        self.dimensionFrame = tk.LabelFrame(self.pdfSettingBox, text=_('尺寸'))
        self.dimensionFrame.grid(row=2, column=0, columnspan=6, sticky=tk.NSEW)
        for c in range(6):
            self.dimensionFrame.columnconfigure(c, weight=1)
        tk.Label(self.dimensionFrame, text=_('单位：')).grid(row=0, column=0, sticky=tk.E)
        self.unitsOptionCombox = ttk.Combobox(self.dimensionFrame, values=list(NxModules.UnitsOption.keys()), state='readonly')
        self.unitsOptionCombox.grid(row=0, column=1, sticky=tk.W)
        self.unitsOptionCombox.bind("<<ComboboxSelected>>", self.onChangeDimensionUnits)
        self.lastDimensionUnits = list(NxModules.UnitsOption.keys())[1] # 英寸
        self.unitsOptionCombox.set(self.lastDimensionUnits)

        tk.Label(self.dimensionFrame, text=_('X尺寸：')).grid(row=0, column=2, sticky=tk.E)
        self.dimensionXValue = tk.DoubleVar(parent, value='8.5')
        tk.Entry(self.dimensionFrame, textvariable=self.dimensionXValue, bd=3).grid(row=0, column=3, sticky=tk.W)

        tk.Label(self.dimensionFrame, text=_('Y尺寸：')).grid(row=0, column=4, sticky=tk.E)
        self.dimensionYValue = tk.DoubleVar(parent, value='11.0')
        tk.Entry(self.dimensionFrame, textvariable=self.dimensionYValue, bd=3).grid(row=0, column=5, sticky=tk.W)

        # 设置
        self.settingBox = tk.LabelFrame(self.pdfSettingBox, text=_('设置'), padx=10, pady=10)
        self.settingBox.grid(row=3, column=0, columnspan=6, sticky=tk.NSEW)
        for c in range(6):
            self.settingBox.columnconfigure(c, weight=1)

        tk.Label(self.settingBox, text=_('输出文本：')).grid(row=0, column=0, sticky=tk.E)
        self.outputTextOptionCombox = ttk.Combobox(self.settingBox, values=list(NxModules.OutputTextOption.keys()), state='readonly')
        self.outputTextOptionCombox.grid(row=0, column=1, sticky=tk.W)
        self.outputTextOptionCombox.set(list(NxModules.OutputTextOption.keys())[1]) # 多段线

        self.addWatermarkBoolValue = tk.BooleanVar(parent, value=False)
        tk.Checkbutton(self.settingBox, text=_('添加水印'), variable=self.addWatermarkBoolValue, command=self.onChangeAddWatermark).grid(row=0, column=2, sticky=tk.E)
        self.watermarkTextValue = tk.StringVar(parent, value='')
        self.watermarkEntry = tk.Entry(self.settingBox, textvariable=self.watermarkTextValue, bd=3)
        self.watermarkEntry.grid(row=0, column=3, columnspan=3, sticky=tk.NSEW)

        # 设置 - 图纸页
        self.paperSettingBox = tk.LabelFrame(self.settingBox, text=_('图纸页'), padx=10, pady=10)
        self.paperSettingBox.grid(row=1, column=0, columnspan=6, sticky=tk.NSEW)
        for c in range(6):
            self.paperSettingBox.columnconfigure(c, weight=1)

        self.rasterImagesBoolValue = tk.BooleanVar(parent, value=True)
        tk.Checkbutton(self.paperSettingBox, text=_('在图纸上导出光栅图像'), variable=self.rasterImagesBoolValue, command=self.onChangeRaster).grid(row=0, column=0, sticky=tk.W)

        self.shadedGeometryBoolValue = tk.BooleanVar(parent, value=False)
        self.shadedGeometryCheckBtn = tk.Checkbutton(self.paperSettingBox, text=_('将着色的几何体导出为线框'), variable=self.shadedGeometryBoolValue)
        self.shadedGeometryCheckBtn.grid(row=0, column=1, sticky=tk.W)

        self.customSymbolsInForegroundBoolValue = tk.BooleanVar(parent, value=False)
        tk.Checkbutton(self.paperSettingBox, text=_('在前景中导出定制符号'), variable=self.customSymbolsInForegroundBoolValue).grid(row=0, column=2, sticky=tk.W)

        tk.Label(self.paperSettingBox, text=_('图像分辨率：')).grid(row=0, column=3, sticky=tk.E)
        self.imageResolutionCombox = ttk.Combobox(self.paperSettingBox, values=list(NxModules.ImageResolution.keys()), state='readonly')
        self.imageResolutionCombox.grid(row=0, column=4, sticky=tk.W)
        self.imageResolutionCombox.set(list(NxModules.ImageResolution.keys())[0]) # 高

        self.exportButton = tk.Button(parent, text=_('导出'), command=self.onExport, bd=3)
        self.exportButton.grid(row=14, column=0, columnspan=3, sticky=tk.NSEW)

        # 初始化选项状态
        self.onChangeSizeOption()
        self.onChangeAddWatermark()

    def onChangeRaster(self):
        """切换光栅图像"""
        if self.rasterImagesBoolValue.get():
            self.shadedGeometryCheckBtn.config(state=tk.NORMAL)
        else:
            self.shadedGeometryBoolValue.set(True)
            self.shadedGeometryCheckBtn.config(state=tk.DISABLED)

    def onChangeAddWatermark(self):
        """切换添加水印"""
        if self.addWatermarkBoolValue.get():
            self.watermarkEntry.config(state=tk.NORMAL)
        else:
            self.watermarkEntry.config(state=tk.DISABLED)

    def onChangeDrawingClass(self):
        """切换导出格式"""
        if self.drawingClassValue.get() == self.DrawingClass.PDF.value:
            self.pdfSettingBox.grid()
        else:
            self.pdfSettingBox.grid_remove()

    def onChangeSizeOption(self, event=None):
        """切换尺寸选项"""
        sizeOptionValue = self.sizeOptionCombox.get()
        if sizeOptionValue == list(NxModules.SizeOption.keys())[1]:
            self.scaleFactorFrame.grid()
            self.dimensionFrame.grid_remove()
        elif sizeOptionValue == list(NxModules.SizeOption.keys())[2]:
            self.scaleFactorFrame.grid_remove()
            self.dimensionFrame.grid()
        else:
            self.scaleFactorFrame.grid_remove()
            self.dimensionFrame.grid_remove()

    def onChangeDimensionUnits(self, event=None):
        """切换尺寸单位"""
        try:
            x = self.dimensionXValue.get()
            y = self.dimensionYValue.get()
        except tk.TclError:
            messagebox.showerror(_('类型错误'), _('尺寸必须为数值类型'))
            return

        currUnitsValue = self.unitsOptionCombox.get()
        if currUnitsValue != self.lastDimensionUnits:
            precision = 4 # 保留小数点后四位
            if currUnitsValue == list(NxModules.UnitsOption.keys())[0]: # 毫米
                self.dimensionXValue.set(round(x * 25.4, precision))
                self.dimensionYValue.set(round(y * 25.4, precision))
            else:
                self.dimensionXValue.set(round(x / 25.4, precision))
                self.dimensionYValue.set(round(y / 25.4, precision))
            self.lastDimensionUnits = currUnitsValue

    def createMergePdfTab(self, parent: tk.Frame):
        """
        创建合并 PDF 页面

        Args:
            parent (tk.Frame): 父组件
        """
        parent.rowconfigure([0,1,2,3,4,5,6], weight=1)
        parent.columnconfigure(0, weight=1)

        self.pdfListShow = ListShow(parent, heading2=_('合并'), heading3=_('PDF 图纸'))
        self.pdfListShow.grid(row=0, column=0, rowspan=4, sticky=tk.NSEW)

        tk.Button(parent, text=_('打开文件'), command=self.onOpenPdfs).grid(row=0, column=1, sticky=tk.EW)
        tk.Button(parent, text=_('清空显示'), command=self.pdfListShow.clear).grid(row=1, column=1, sticky=tk.EW)
        tk.Button(parent, text=_('设置选中项'), command=self.onSetSelectedPdfStatus).grid(row=2, column=1, sticky=tk.EW)
        tk.Button(parent, text=_('设置所有'), command=self.onSetAllPdfStatus).grid(row=3, column=1, sticky=tk.EW)


        self.selectedPdfStatusValue = tk.BooleanVar(parent, value=True)
        tk.Checkbutton(parent, text=_('合并'), variable=self.selectedPdfStatusValue).grid(row=2, column=2, sticky=tk.W)

        self.allPdfStatusValue = tk.BooleanVar(parent, value=True)
        tk.Checkbutton(parent, text=_('合并'), variable=self.allPdfStatusValue).grid(row=3, column=2, sticky=tk.W)

        self.savePdfNameValue = tk.StringVar(parent, value='')
        tk.Entry(parent, textvariable=self.savePdfNameValue, bd=3).grid(row=4, column=0, sticky=tk.EW)

        tk.Button(parent, text=_('设置合并后的 PDF 文件名'), command=self.onSaveMergedPdf).grid(row=4, column=1, sticky=tk.EW)
        tk.Button(parent, text=_('合并'), command=self.onMergePdf).grid(row=5, column=0, columnspan=3, sticky=tk.NSEW)

    def _config_text_tags(self, text_widget: tk.Text):
        """配置 Text 组件的颜色标签"""
        text_widget.tag_config('info', foreground='black')      # 普通消息
        text_widget.tag_config('success', foreground='green')   # 成功消息
        text_widget.tag_config('warn', foreground='orange')     # 警告消息
        text_widget.tag_config('error', foreground='red')       # 错误消息

    def onMergePdf(self):
        pdfList = self.pdfListShow.getPrtList()
        savePdf = self.savePdfNameValue.get()
        if savePdf == '':
            self.writeMsg(_('请设置合并后的 PDF 文件名'))
            return
        
        self.exportButton.config(state=tk.DISABLED)
        PdfModules.mergePdfs(savePdf, pdfList, self.writeMsg)
        # NX 2506 中执行时有异常，必须执行 update，否则按钮禁用状态下的点击也会留到本轮执行结束后继续触发
        self.update()
        self.exportButton.config(state=tk.NORMAL)

    def onOpenPdfs(self):
        """
        打开要合并的 PDF 文件
        """
        pdfFiles = filedialog.askopenfilenames(
            title=_('选择要合并的 PDF 文件'),
            filetypes=[(_('PDF 文件'), '*.pdf')]
        )
        self.pdfListShow.insertPrtList(pdfFiles)

    def onSetSelectedPdfStatus(self):
        """
        设置 PDF 图纸选中项的状态
        """
        self.pdfListShow.setSelectedStatus(self.selectedPdfStatusValue.get())

    def onSetAllPdfStatus(self):
        """
        设置所有 PDF 图纸的状态
        """
        self.pdfListShow.setAllStatus(self.allPdfStatusValue.get())

    def onSaveMergedPdf(self):
        """
        保存合并的 PDF
        """
        saveFilename = filedialog.asksaveasfilename(
            title=_('设置合并后的 PDF 文件名'),
            filetypes=[(_('PDF 文件'), '*.pdf')],
            defaultextension='.pdf'
        )
        if saveFilename == '':
            return
        saveFilename = os.path.normpath(saveFilename)
        self.savePdfNameValue.set(saveFilename)

    def onGetSelectedPrts(self):
        """
        获取选择的文件
        """
        prtList = NxModules.selectPrts()
        self.prtListShow.insertPrtList(prtList)

    def onGetPrtsFromSelectedFolder(self):
        """
        获取选择的文件夹中的文件
        """
        prtList = NxModules.getPrtsFromSelectedFolder(self.recursiveValue.get())
        self.prtListShow.insertPrtList(prtList)

    def onGetOpenedPrts(self):
        """
        获取已打开的文件
        """
        prtList = NxModules.getOpenedPrts()
        self.prtListShow.insertPrtList(prtList)

    def onWorkPrt(self):
        """
        获取工作部件
        """
        prtList = NxModules.getWorkPrt()
        self.prtListShow.insertPrtList(prtList)

    def onSetSelectedPrtStatus(self):
        """
        设置 NX 图纸选中项的状态
        """
        self.prtListShow.setSelectedStatus(self.selectedPrtStatusValue.get())

    def onSetAllPrtStatus(self):
        """
        设置所有 NX 图纸的状态
        """
        self.prtListShow.setAllStatus(self.allPrtStatusValue.get())

    def onSelectExportFolder(self):
        """
        选择导出文件夹
        """
        folder = filedialog.askdirectory(title='选择导出文件夹')
        if folder == '':
            return
        folder = os.path.normpath(folder)
        self.exportFolderValue.set(folder)

    def writeMsg(self, msg: str, level: str = 'info'):
        """
        写消息

        Args:
            msg (str): 消息内容
            level (str): 消息级别 ('info', 'success', 'warn', 'error')
        """
        if self.msgText is None:
            return
            
        self.msgText.config(state=tk.NORMAL)
        # 在 insert 时传入对应的 tag 名字
        self.msgText.insert(tk.END, msg + '\n', level) 
        self.msgText.config(state=tk.DISABLED)
        self.msgText.see(tk.END)
        self.msgText.update_idletasks()

    def extraPdfSetting(self) -> bool:
        '''提取设置'''
        try:
            color_key = self.colorCombox.get()
            width_key = self.widthCombox.get()
            size_key = self.sizeOptionCombox.get()
            units_key = self.unitsOptionCombox.get()
            output_key = self.outputTextOptionCombox.get()
            res_key = self.imageResolutionCombox.get()

            if not all([color_key, width_key, size_key, units_key, output_key, res_key]):
                self.writeMsg(_('存在未选择的 PDF 设置项！'), 'error')
                return False

            color = NxModules.Color[color_key]
            width = NxModules.Width[width_key]
            size = NxModules.SizeOption[size_key]

            scale = self.scaleFactorValue.get()
            units = NxModules.UnitsOption[units_key]
            xDimension = self.dimensionXValue.get()
            yDimension = self.dimensionYValue.get()

            outputText = NxModules.OutputTextOption[output_key]
            addWatermark = self.addWatermarkBoolValue.get()
            watermark = self.watermarkTextValue.get()
            rasterImages = self.rasterImagesBoolValue.get()
            shadedGeometry = self.shadedGeometryBoolValue.get()
            customSymbolsInForeground = self.customSymbolsInForegroundBoolValue.get()
            imageResolution = NxModules.ImageResolution[res_key]

            self.pdfSettingObj = NxModules.PdfSettings(
                color, width, size, scale, units, xDimension, yDimension, 
                outputText, addWatermark, watermark, rasterImages, 
                shadedGeometry, customSymbolsInForeground, imageResolution
            )
            return True
        except (tk.TclError, KeyError, Exception) as e:
            self.writeMsg(_('提取 PDF 设置异常: {e}').format(e=str(e)), 'error')
            return False

    def onExport(self):
        """"
        导出
        """
        prtList = self.prtListShow.getPrtList()
        exportFolder = self.exportFolderValue.get().strip()
        if exportFolder != '':
            os.makedirs(exportFolder, exist_ok=True)
        self.exportButton.config(state=tk.DISABLED)
        try:
            match self.drawingClassValue.get():
                case self.DrawingClass.PDF.value:
                    if not self.extraPdfSetting():
                        self.writeMsg(_('提取PDF设置失败，停止导出！'), 'error')
                        return
                    if self.pdfSettingObj is None:
                        self.writeMsg(_('PDF设置对象未初始化，停止导出！'), 'error')
                        return
                    NxModules.foreachPrt(
                        NxModules.exportPdf, prtList, 
                        self.prefixNameValue.get(), self.suffixNameValue.get(), 
                        exportFolder, self.writeMsg, self.pdfSettingObj
                    )
                case self.DrawingClass.DWG.value:
                    NxModules.foreachPrt(
                        NxModules.exportDwg, prtList, 
                        self.prefixNameValue.get(), self.suffixNameValue.get(), 
                        exportFolder, self.writeMsg
                    )
        except Exception as e:
            self.writeMsg(_('导出过程发生未捕获异常: {e}').format(e=str(e)), 'error')
        finally:
            self.update()
            self.exportButton.config(state=tk.NORMAL)

def run():
    root = tk.Tk()
    width = 900
    height = 900
    screenwidth = root.winfo_screenwidth()
    screenheight = root.winfo_screenheight()
    root.geometry(f'{width}x{height}+{int((screenwidth - width) / 2)}+{int((screenheight - height) / 2)}')
    root.attributes('-topmost', True)
    root.title(f'{programName}')
    root.iconbitmap(os.path.join(os.path.dirname(__file__), 'icon.ico'))
    app = Application(root)
    app.mainloop()