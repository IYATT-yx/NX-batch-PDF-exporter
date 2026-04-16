from listshow import ListShow
from nxmodules import NxModules
import version
from pdfmodules import PdfModules

import tkinter as tk
from tkinter import ttk, filedialog
import os
from enum import Enum, auto

programName = 'Siemens NX 批量 PDF 图纸导出工具'
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

        self.createWidgets()
        self.writeMsg(f'{programName} {version.V}\nIYATT-yx iyatt@iyatt.com\n本工具为开源项目，发布地址：https://github.com/IYATT-yx/NX-batch-PDF-exporter')


    def createWidgets(self):
        """
        创建组件
        """
        self.notebook = ttk.Notebook(self)
        self.notebook.pack(expand=True, fill=tk.BOTH)

        self.exportPdfTab = ttk.Frame(self.notebook)
        self.createExportPdfTab(self.exportPdfTab)
        self.notebook.add(self.exportPdfTab, text='导出图纸')

        self.mergePdfTab = ttk.Frame(self.notebook)
        self.createMergePdfTab(self.mergePdfTab)
        self.notebook.add(self.mergePdfTab, text='合并PDF')

        self.notebook.bind('<<NotebookTabChanged>>', self.onTabChanged)

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

        tk.Button(parent, text='选择文件', command=self.onGetSelectedPrts).grid(row=0, column=1, sticky=tk.NSEW)
        tk.Button(parent, text='选择文件夹', command=self.onGetPrtsFromSelectedFolder).grid(row=1, column=1, sticky=tk.NSEW)
        tk.Button(parent, text='所有打开文件', command=self.onGetOpenedPrts).grid(row=2, column=1, sticky=tk.NSEW)
        tk.Button(parent, text='工作部件', command=self.onWorkPrt).grid(row=3, column=1, sticky=tk.NSEW)
        tk.Button(parent, text='清空显示', command=self.prtListShow.clear).grid(row=4, column=1, sticky=tk.NSEW)
        ttk.Separator(parent, orient=tk.HORIZONTAL).grid(row=5, column=1, columnspan=2, sticky=tk.EW, pady=5)

        tk.Button(parent, text='设置选中项', command=self.onSetSelectedPrtStatus).grid(row=6, column=1, sticky=tk.NSEW)
        tk.Button(parent, text='设置所有', command=self.onSetAllPrtStatus).grid(row=7, column=1, sticky=tk.NSEW)
        
        self.recursiveValue = tk.BooleanVar(parent, value=False)
        tk.Checkbutton(parent, text='递归', variable=self.recursiveValue).grid(row=1, column=2, sticky=tk.W)
        self.selectedPrtStatusValue = tk.BooleanVar(parent, value=True)
        tk.Checkbutton(parent, text='导出', variable=self.selectedPrtStatusValue).grid(row=6, column=2, sticky=tk.W)
        self.allPrtStatusValue = tk.BooleanVar(parent, value=True)
        tk.Checkbutton(parent, text='导出', variable=self.allPrtStatusValue).grid(row=7, column=2, sticky=tk.W)

        ttk.Separator(parent, orient=tk.HORIZONTAL).grid(row=8, column=0, columnspan=3, sticky=tk.EW, pady=5)

        self.exportFolderValue = tk.StringVar(parent, value='')
        tk.Entry(parent, textvariable=self.exportFolderValue, bd=3).grid(row=9, column=0, sticky=tk.NSEW)
        tk.Button(parent, text='选择导出文件夹', command=self.onSelectExportFolder).grid(row=9, column=1, sticky=tk.NSEW)
        tk.Label(parent, text='（可选）').grid(row=9, column=2, sticky=tk.W)

        self.prefixNameValue = tk.StringVar(parent, value='')
        tk.Entry(parent, textvariable=self.prefixNameValue,  bd=3).grid(row=10, column=0, sticky=tk.NSEW)
        tk.Label(parent, text='<--- 导出文件名前缀（可选）').grid(row=10, column=1, columnspan=2, sticky=tk.W)

        self.suffixNameValue = tk.StringVar(parent, value='')
        tk.Entry(parent, textvariable=self.suffixNameValue, bd=3).grid(row=11, column=0, sticky=tk.NSEW)
        tk.Label(parent, text='<--- 导出文件名后缀（可选）').grid(row=11, column=1, columnspan=2, sticky=tk.W)

        self.drawingClassValue = tk.IntVar(parent, self.DrawingClass.PDF.value)
        tk.Label(parent, text='导出格式：').grid(row=12, column=0, sticky=tk.E)
        tk.Radiobutton(parent, text='PDF', variable=self.drawingClassValue, value=self.DrawingClass.PDF.value).grid(row=12, column=1, sticky=tk.W)
        tk.Radiobutton(parent, text='DWG', variable=self.drawingClassValue, value=self.DrawingClass.DWG.value).grid(row=12, column=2, sticky=tk.W)

        self.exportButton = tk.Button(parent, text='导出', command=self.onExport, bd=3)
        self.exportButton.grid(row=13, column=0, columnspan=3, sticky=tk.NSEW)

        self.prtMsgText = tk.Text(parent, wrap=tk.CHAR, height=10, state=tk.DISABLED)
        self.prtMsgText.grid(row=14, column=0, columnspan=3, sticky=tk.NSEW)
        self._config_text_tags(self.prtMsgText)
        self.msgText = self.prtMsgText

    def createMergePdfTab(self, parent: tk.Frame):
        """
        创建合并 PDF 页面

        Args:
            parent (tk.Frame): 父组件
        """
        parent.rowconfigure([0,1,2,3,4,5,6], weight=1)
        parent.columnconfigure(0, weight=1)

        self.pdfListShow = ListShow(parent, heading2='合并', heading3='PDF 图纸')
        self.pdfListShow.grid(row=0, column=0, rowspan=4, sticky=tk.NSEW)

        tk.Button(parent, text='打开文件', command=self.onOpenPdfs).grid(row=0, column=1, sticky=tk.EW)
        tk.Button(parent, text='清空显示', command=self.pdfListShow.clear).grid(row=1, column=1, sticky=tk.EW)
        tk.Button(parent, text='设置选中项', command=self.onSetSelectedPdfStatus).grid(row=2, column=1, sticky=tk.EW)
        tk.Button(parent, text='设置所有', command=self.onSetAllPdfStatus).grid(row=3, column=1, sticky=tk.EW)


        self.selectedPdfStatusValue = tk.BooleanVar(parent, value=True)
        tk.Checkbutton(parent, text='合并', variable=self.selectedPdfStatusValue).grid(row=2, column=2, sticky=tk.W)

        self.allPdfStatusValue = tk.BooleanVar(parent, value=True)
        tk.Checkbutton(parent, text='合并', variable=self.allPdfStatusValue).grid(row=3, column=2, sticky=tk.W)

        self.savePdfNameValue = tk.StringVar(parent, value='')
        tk.Entry(parent, textvariable=self.savePdfNameValue, bd=3).grid(row=4, column=0, sticky=tk.EW)

        tk.Button(parent, text='设置合并后的 PDF 文件名', command=self.onSaveMergedPdf).grid(row=4, column=1, sticky=tk.EW)
        tk.Button(parent, text='合并', command=self.onMergePdf).grid(row=5, column=0, columnspan=3, sticky=tk.NSEW)

        self.pdfMsgText = tk.Text(parent, wrap=tk.CHAR, height=10, state=tk.DISABLED)
        self.pdfMsgText.grid(row=6, column=0, columnspan=3, sticky=tk.NSEW)
        self._config_text_tags(self.pdfMsgText)

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
            self.writeMsg('请设置合并后的 PDF 文件名')
            return
        
        self.exportButton.config(state=tk.DISABLED)
        PdfModules.mergePdfs(savePdf, pdfList, self.writeMsg)
        # NX 2506 中执行时有异常，必须执行 update，否则按钮禁用状态下的点击也会留到本轮执行结束后继续触发
        self.update()
        self.exportButton.config(state=tk.NORMAL)

    def onTabChanged(self, event: tk.Event):
        """
        切换标签页事件处理，根据当前标签页设置消息框

        Args:
            event (tk.Event): 事件对象
        """
        currentTabId = self.notebook.select()
        if not currentTabId:
            return

        currentFrame = self.notebook.nametowidget(currentTabId)
        match currentFrame:
            case self.exportPdfTab:
                self.msgText = self.prtMsgText
            case self.mergePdfTab:
                self.msgText = self.pdfMsgText
    
    def onOpenPdfs(self):
        """
        打开要合并的 PDF 文件
        """
        pdfFiles = filedialog.askopenfilenames(
            title='选择要合并的 PDF 文件',
            filetypes=[('PDF 文件', '*.pdf')]
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
            title='设置合并后的 PDF 文件名',
            filetypes=[('PDF 文件', '*.pdf')],
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

    def onExport(self):
        """"
        导出
        """
        prtList = self.prtListShow.getPrtList()
        exportFolder = self.exportFolderValue.get().strip()
        if exportFolder != '':
            os.makedirs(exportFolder, exist_ok=True)
        self.exportButton.config(state=tk.DISABLED)
        match self.drawingClassValue.get():
            case self.DrawingClass.PDF.value:
                NxModules.foreachPrt(NxModules.exportPdf, prtList, self.prefixNameValue.get(), self.suffixNameValue.get(), exportFolder, self.writeMsg)
            case self.DrawingClass.DWG.value:
                NxModules.foreachPrt(NxModules.exportDwg, prtList, self.prefixNameValue.get(), self.suffixNameValue.get(), exportFolder, self.writeMsg)
        # NX 2506 中执行时有异常，必须执行 update，否则按钮禁用状态下的点击也会留到本轮执行结束后继续触发
        self.update()
        self.exportButton.config(state=tk.NORMAL)

def run():
    root = tk.Tk()
    width = 900
    height = 600
    screenwidth = root.winfo_screenwidth()
    screenheight = root.winfo_screenheight()
    root.geometry(f'{width}x{height}+{int((screenwidth - width) / 2)}+{int((screenheight - height) / 2)}')
    root.attributes('-topmost', True)
    root.title(f'{programName}')
    root.iconbitmap(os.path.join(os.path.dirname(__file__), 'icon.ico'))
    app = Application(root)
    app.mainloop()