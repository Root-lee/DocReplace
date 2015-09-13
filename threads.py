# -*- coding: utf-8 -*-
from PyQt4 import QtCore, QtGui
#import win32com ,os
#from win32com.client import Dispatch, constants
import time,os

#继承 QThread 类
class BigWorkThread(QtCore.QThread):
    """docstring for BigWorkThread"""
    def __init__(self, line = []):
        super(BigWorkThread, self).__init__(None)

        self.doctype = line[0]
        self.dir = line[1]
        self.OldStr = line[2]
        self.NewStr = line[3]
        self.filetype = self.getfiletype()


        
    def run(self):
        
        



        self.getallfile(self.dir)

        

    def getallfile(self,path):
        import win32com ,os
        from win32com.client import Dispatch, constants
        w = win32com.client.Dispatch('Word.Application')
        # w = win32com.client.DispatchEx('Word.Application')

        w.Visible = 0
        w.DisplayAlerts = 0
        for directory, dirnames, filenames in os.walk(path):
       #     for dirname in dirnames:
        #        #self.getallfile(directory+"\\"+dirname)
        #        print dirname+"-----"
            if not os.path.isdir(path+"\\..\\NewDocFiles"):
                os.mkdir(path+"\\..\\NewDocFiles")
            for filename in filenames:
                
      #          print 'Directory',dirnames
                
                if filename.endswith(self.filetype):
                    print ' File',filename
                    filenamein=directory+"\\"+filename
                    directoryname = directory.split(path)[-1]
                    filenameout=path+"\\..\\NewDocFiles"+directoryname+"\\"+filename
                    print filenameout
                    if not os.path.isdir(path+"\\..\\NewDocFiles"+directoryname):
                        os.mkdir(path+"\\..\\NewDocFiles"+directoryname)
                    try:
                        doc = w.Documents.Open( FileName = filenamein )
                    except:
                        continue
                
                
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
                    doc.SaveAs(filenameout)
                    doc.Close()
                else:
                    pass
        w.Quit()
        print "ssssssssssss"
    def getfiletype(self):
        doctypeinput = self.doctype
        if doctypeinput == 1:
            doctype='.doc'
        elif doctypeinput == 2:
            doctype='.docx'
        else :
            doctype='.docfhhfd'
        return doctype


