# -*- coding: cp936 -*-
import win32com ,os
from win32com.client import Dispatch, constants
w = win32com.client.Dispatch('Word.Application')
# w = win32com.client.DispatchEx('Word.Application')

w.Visible = 0
w.DisplayAlerts = 0

current_path = os.getcwd()
files = os.listdir(current_path+"\\..")
print "请选择需要替换文档的格式：[1]:.doc    [2]: .docx"
doctypeinput=input()
if doctypeinput == 1:
    doctype='.doc'
elif doctypeinput == 2:
    doctype='.docx'
else :
    doctype='.docfhhfd'
OldStr=raw_input("被替换的词：")
NewStr=raw_input("把 "+OldStr+" 替换为：")
for filename in files:
    if filename.endswith(doctype):
    # 打开新的文件
        
        filenamein=current_path+"\\..\\"+filename
        filenameout=current_path+"\\..\\new\\"+filename
        if not os.path.isdir(current_path+"\\..\\new"):
            os.mkdir(current_path+"\\..\\new")
        print filename+"替换成功！"
        doc = w.Documents.Open( FileName = filenamein )

        
        
        # 正文文字替换
        w.Selection.Find.ClearFormatting()
        w.Selection.Find.Replacement.ClearFormatting()
        w.Selection.Find.Execute(OldStr, False, False, False, False, False, True, 1,
        True, NewStr, 2)
        # 页眉文字替换
        w.ActiveDocument.Sections[0].Headers[0].Range.Find.ClearFormatting()
        w.ActiveDocument.Sections[0].Headers[0].Range.Find.Replacement.ClearFormatting()
        w.ActiveDocument.Sections[0].Headers[0].Range.Find.Execute(OldStr, False,
        False, False, False, False, True, 1, False, NewStr, 2)


        # 关闭
        doc.SaveAs(filenameout)
        doc.Close()
#w.Documents.Close()
w.Quit()
