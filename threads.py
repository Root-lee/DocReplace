# -*- coding: utf-8 -*-
from PyQt4 import QtCore, QtGui
import time,os

#继承 QThread 类
class BigWorkThread(QtCore.QThread):
    """docstring for BigWorkThread"""
    def __init__(self, line = []):
        super(BigWorkThread, self).__init__(None)
        self.errorlist = ""
        self.doctype = line[0]
        self.dir = line[1]
        self.OldStr = line[2]
        self.NewStr = line[3]
        self.filetype = self.getfiletype()
    
	#子进程函数
    def run(self):
        
        self.getallfile(self.dir)

    def getallfile(self,path):
        import win32com ,os
        from win32com.client import Dispatch, constants
        #判断是用EXCEL还是WORD程序打开
        if ((self.filetype == ".xls") | (self.filetype == ".xlsx")):
            w = win32com.client.Dispatch('Excel.Application')
        else:
            w = win32com.client.Dispatch('Word.Application')
        count=0
        filenum=0.0
        w.Visible = 0
        w.DisplayAlerts = 0
		#遍历文件夹所有特定格式文件
        for directory, dirnames, filenames in os.walk(path):
            for filename in filenames:
                if filename.endswith(self.filetype):
                    filenum = filenum + 1
                    self.emit(QtCore.SIGNAL("finddocfile"),filenum)
        #对各个文件处理
        for directory, dirnames, filenames in os.walk(path):
            directoryname = directory.split(path)[-1]
            filenameout=path+"\\..\\NewDocFiles"+directoryname+"\\"+filename
                    
            for filename in filenames:
                
                if filename.endswith(self.filetype):
                    try:
                        filenamein=directory+"\\"+filename
                        directoryname = directory.split(path)[-1]

                        if ((self.filetype == ".xls") | (self.filetype == ".xlsx")):  #用EXCEL处理
                            workbook = w.Workbooks.Open(filenamein)
                            w.Cells.Select
                            w.Selection.Replace( self.OldStr, self.NewStr)

                            workbook.Save()
                            workbook.Close(SaveChanges=True)
                        else:    
                            doc = w.Documents.Open( FileName = filenamein )  #用WORD处理
                            # 正文文字替换
                            w.Selection.Find.ClearFormatting()
                            w.Selection.Find.Replacement.ClearFormatting()
                            w.Selection.Find.Execute(self.OldStr, False, False, False, False, False, True, 1,
                            True, self.NewStr, 2)
                            # 页眉文字替换
                            w.ActiveDocument.Sections[0].Headers[0].Range.Find.ClearFormatting()
                            w.ActiveDocument.Sections[0].Headers[0].Range.Find.Replacement.ClearFormatting()
                            w.ActiveDocument.Sections[0].Headers[0].Range.Find.Execute(self.OldStr, False,
                            False, False, False, False, True, 1, False, self.NewStr, 2)
                            # 关闭
                            #doc.SaveAs(filenameout)
                            doc.Save()
                            doc.Close()
					 
                        count=count+1
                        j = int(count/filenum*100)
                        self.emit(QtCore.SIGNAL("where"),j)
                        self.emit(QtCore.SIGNAL("showtxt"),filename)
                    except:
                       
						self.errorlist = self.errorlist + "\n    error:"+directory+"\\"+filename
						count =count + 1
                else:
                    pass
        w.Quit()

        self.emit(QtCore.SIGNAL("finish_show"),self.errorlist)
    def getfiletype(self):
        doctypeinput = self.doctype
        if doctypeinput == 1:
            doctype='.doc'
        elif doctypeinput == 2:
            doctype='.docx'
        elif doctypeinput == 3:
            doctype='.xls'
        elif doctypeinput == 4:
            doctype='.xlsx'
        elif doctypeinput == 5:
            doctype='.txt'
        else :
            doctype='.docfhhfd'
        return doctype


