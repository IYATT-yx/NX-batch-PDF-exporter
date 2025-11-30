"""
图纸列表显示控件
"""
import tkinter as tk
from tkinter import ttk

import os

class PrtListShow(tk.Frame):
    def __init__(self, master, defaultStatus = '是'):
        super().__init__(master)
        self.master = master
        self.number = 1
        self.defaultStatus = defaultStatus
        self.createWidgets()

    def createWidgets(self):
        self.tree = ttk.Treeview(self, columns=('number', 'status', 'filePath'), show='headings')
        self.tree.heading('number', text='序号')
        self.tree.heading('status', text='导出')
        self.tree.heading('filePath', text='图纸')
        self.tree.column('number', anchor=tk.CENTER, width=50, stretch=False)
        self.tree.column('status', anchor=tk.CENTER, width=50, stretch=False)
        self.tree.column('filePath', anchor=tk.W, width=200, stretch=True)

        vsb = ttk.Scrollbar(self, orient=tk.VERTICAL, command=self.tree.yview)
        self.tree.configure(yscrollcommand=vsb.set)

        self.tree.grid(row=0, column=0, sticky=tk.NSEW)
        vsb.grid(row=0, column=1, sticky=tk.NS)

        self.rowconfigure(0, weight=1)
        self.columnconfigure(0, weight=1)

        self.tree.bind('<Button-1>', self.toggleCheck)

    def insertPrtList(self, prtList: list[str]):
        """
        插入图纸列表

        Args:
            prtList (list[str]): 图纸列表
        """
        prtList = [os.path.normpath(prt) for prt in prtList if os.path.exists(prt)]
        prtList = list(set(prtList)) # 去重
        prtList = sorted(prtList)
        for prt in prtList:
            try:
                self.tree.insert('', tk.END, iid=prt, values=(f'{self.number}', self.defaultStatus, prt))
                self.number += 1
            except tk.TclError as e:
                if 'already exists' in str(e): # 跳过重复文件
                    continue
                raise

    def clear(self):
        """
        清空图纸列表显示
        """
        self.tree.delete(*self.tree.get_children())
        self.number = 1

    def toggleCheck(self, event):
        rowId = self.tree.identify_row(event.y)
        # 点击空白处时忽略
        if rowId == '':
            return
        status = self.tree.set(rowId)['status']
        if status == '是':
            self.tree.set(rowId, 'status', '否')
        else:
            self.tree.set(rowId, 'status', '是')

    def setAllStatus(self, status: bool = True):
        """
        设置所有图纸的导出状态

        Args:
            status (bool, optional): 导出状态. Defaults to True.
        """
        for rowId in self.tree.get_children():
            self.tree.set(rowId, 'status', '是' if status else '否')

    def setSelectedStatus(self, status: bool = True):
        """
        设置选中图纸的导出状态

        Args:
            status (bool, optional): 导出状态. Defaults to True.
        """
        for rowId in self.tree.selection():
            self.tree.set(rowId, 'status', '是' if status else '否')

    def getPrtList(self) -> list[str]:
        """
        获取需要导出的图纸列表

        Returns:
            list[str]: 需要导出的图纸列表
        """
        return [self.tree.set(rowId)['filePath'] for rowId in self.tree.get_children() if self.tree.set(rowId)['status'] == '是']
        


