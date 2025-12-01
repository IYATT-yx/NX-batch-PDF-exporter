from prtlistshow import PrtListShow
from nxmodules import NxModules
import version

import tkinter as tk
from tkinter import ttk, filedialog
import os

programName = 'Siemens NX 批量 PDF 图纸导出工具'
baseDir = os.path.dirname(os.path.abspath(__file__))
tclDir = os.path.join(baseDir, 'tcl', 'tcl8.6')
os.environ["TCL_LIBRARY"] = tclDir

class Application(tk.Frame):
    def __init__(self, root):
        super().__init__(root)
        self.root = root
        self.pack(fill=tk.BOTH, expand=True)

        self.createWidgets()
        self.writeMsg(f'{programName} {version.V}\nIYATT-yx iyatt@iyatt.com')

        self.rowconfigure([0,1,2,3,4,5,6], weight=1)
        self.columnconfigure(0, weight=1)

    def createWidgets(self):
        self.prtListShow = PrtListShow(self)
        self.prtListShow.grid(row=0, column=0, rowspan=7, sticky=tk.NSEW)

        tk.Button(self, text='选择文件', command=self.onGetSelectedPrts).grid(row=0, column=1, sticky=tk.NSEW)
        tk.Button(self, text='选择文件夹', command=self.onGetPrtsFromSelectedFolder).grid(row=1, column=1, sticky=tk.NSEW)
        tk.Button(self, text='显示文件', command=self.onGetOpenedPrts).grid(row=2, column=1, sticky=tk.NSEW)
        tk.Button(self, text='清空显示', command=self.prtListShow.clear).grid(row=3, column=1, sticky=tk.NSEW)
        ttk.Separator(self, orient=tk.HORIZONTAL).grid(row=4, column=1, columnspan=2, sticky=tk.EW, pady=5)

        tk.Button(self, text='设置选中项', command=self.onSetSelectedStatus).grid(row=5, column=1, sticky=tk.NSEW)
        tk.Button(self, text='设置所有', command=self.onSetAllStatus).grid(row=6, column=1, sticky=tk.NSEW)
        
        self.recursiveValue = tk.BooleanVar(self, value=False)
        tk.Checkbutton(self, text='递归', variable=self.recursiveValue).grid(row=1, column=2, sticky=tk.W)
        self.selectedStatusValue = tk.BooleanVar(self, value=True)
        tk.Checkbutton(self, text='导出', variable=self.selectedStatusValue).grid(row=5, column=2, sticky=tk.W)
        self.allStatusValue = tk.BooleanVar(self, value=True)
        tk.Checkbutton(self, text='导出', variable=self.allStatusValue).grid(row=6, column=2, sticky=tk.W)

        ttk.Separator(self, orient=tk.HORIZONTAL).grid(row=7, column=0, columnspan=3, sticky=tk.EW, pady=5)

        self.exportFolderValue = tk.StringVar(self, value='')
        tk.Entry(self, textvariable=self.exportFolderValue, bd=3).grid(row=8, column=0, sticky=tk.NSEW)
        tk.Button(self, text='选择导出文件夹', command=self.onSelectExportFolder).grid(row=8, column=1, sticky=tk.NSEW)
        tk.Label(self, text='（可选）').grid(row=8, column=2, sticky=tk.W)

        self.prefixNameValue = tk.StringVar(self, value='')
        tk.Entry(self, textvariable=self.prefixNameValue,  bd=3).grid(row=9, column=0, sticky=tk.NSEW)
        tk.Label(self, text='<--- 导出PDF的前缀（可选）').grid(row=9, column=1, columnspan=2, sticky=tk.W)

        self.suffixNameValue = tk.StringVar(self, value='_PDF')
        tk.Entry(self, textvariable=self.suffixNameValue, bd=3).grid(row=10, column=0, sticky=tk.NSEW)
        tk.Label(self, text='<--- 导出PDF的后缀（可选）').grid(row=10, column=1, columnspan=2, sticky=tk.W)

        tk.Button(self, text='导出', command=self.onExport, bd=3).grid(row=11, column=0, columnspan=3, sticky=tk.NSEW)

        self.msgText = tk.Text(self, wrap=tk.CHAR, height=10, state=tk.DISABLED)
        self.msgText.grid(row=12, column=0, columnspan=3, sticky=tk.NSEW)


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

    def onSetSelectedStatus(self):
        """
        设置选中项的状态
        """
        self.prtListShow.setSelectedStatus(self.selectedStatusValue.get())

    def onSetAllStatus(self):
        """
        设置所有项的状态
        """
        self.prtListShow.setAllStatus(self.allStatusValue.get())

    def onSelectExportFolder(self):
        """
        选择导出文件夹
        """
        folder = filedialog.askdirectory(title='选择导出文件夹')
        if folder == '':
            return
        folder = os.path.normpath(folder)
        self.exportFolderValue.set(folder)

    def writeMsg(self, msg: str):
        """
        写消息

        Args:
            msg (str): 消息
        """
        self.msgText.config(state=tk.NORMAL)
        self.msgText.insert(tk.END, msg + '\n')
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
        NxModules.foreachPrt(NxModules.exportPdf, prtList, self.prefixNameValue.get(), self.suffixNameValue.get(), exportFolder, self.writeMsg)        

def run():
    root = tk.Tk()
    width = 900
    height = 490
    screenwidth = root.winfo_screenwidth()
    screenheight = root.winfo_screenheight()
    root.geometry(f'{width}x{height}+{int((screenwidth - width) / 2)}+{int((screenheight - height) / 2)}')
    root.attributes('-topmost', True)
    root.title(f'{programName}')
    root.iconbitmap(os.path.join(os.path.dirname(__file__), 'icon.ico'))
    app = Application(root)
    app.mainloop()