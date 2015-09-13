# -*- coding: utf-8 -*-

# Form implementation generated from reading ui file 'wordreplace.ui'
#
# Created by: PyQt4 UI code generator 4.11.4
#
# WARNING! All changes made in this file will be lost!

from PyQt4 import QtCore, QtGui
import sys ,os 
from threads import BigWorkThread

try:
    _fromUtf8 = QtCore.QString.fromUtf8
except AttributeError:
    def _fromUtf8(s):
        return s

try:
    _encoding = QtGui.QApplication.UnicodeUTF8
    def _translate(context, text, disambig):
        return QtGui.QApplication.translate(context, text, disambig, _encoding)
except AttributeError:
    def _translate(context, text, disambig):
        return QtGui.QApplication.translate(context, text, disambig)

class Ui_Dialog(QtGui.QWidget):
    def  __init__(self, parent = None):
        QtGui.QWidget.__init__(self, parent)
        self.txt = u"- - - - - - - - - - - - - - - - - - - - -\n- Word文档批量替换工具1.0   ---Root lee \n- - - - - - - - - - - - - - - - - - - - -\n"
        self.setupUi(self)
        self.retranslateUi(self)
        self.setWindowIcon(QtGui.QIcon('Word.png'))
    def setupUi(self, Dialog):
        Dialog.setObjectName(_fromUtf8("Dialog"))
        Dialog.resize(399, 348)
        icon = QtGui.QIcon()
        icon.addPixmap(QtGui.QPixmap(_fromUtf8(":/新前缀/Word.png")), QtGui.QIcon.Normal, QtGui.QIcon.Off)
        Dialog.setWindowIcon(icon)
        self.radioButton = QtGui.QRadioButton(Dialog)
        self.radioButton.setGeometry(QtCore.QRect(20, 30, 371, 16))
        self.radioButton.setObjectName(_fromUtf8("radioButton"))
        self.radioButton_2 = QtGui.QRadioButton(Dialog)
        self.radioButton_2.setGeometry(QtCore.QRect(20, 50, 351, 16))
        self.radioButton_2.setObjectName(_fromUtf8("radioButton_2"))

        
        
        
        self.progressBar = QtGui.QProgressBar(Dialog)
        self.progressBar.setGeometry(QtCore.QRect(10, 190, 391, 23))
        self.progressBar.setProperty("value", 0)
        self.progressBar.setObjectName(_fromUtf8("progressBar"))
        self.progressBar.hide()
        self.line = QtGui.QFrame(Dialog)
        self.line.setGeometry(QtCore.QRect(0, 70, 401, 16))
        self.line.setFrameShape(QtGui.QFrame.HLine)
        self.line.setFrameShadow(QtGui.QFrame.Sunken)
        self.line.setObjectName(_fromUtf8("line"))
        self.line_2 = QtGui.QFrame(Dialog)
        self.line_2.setGeometry(QtCore.QRect(0, 170, 401, 20))
        self.line_2.setFrameShape(QtGui.QFrame.HLine)
        self.line_2.setFrameShadow(QtGui.QFrame.Sunken)
        self.line_2.setObjectName(_fromUtf8("line_2"))
        self.textEdit = QtGui.QTextEdit(Dialog)
        self.textEdit.setGeometry(QtCore.QRect(10, 220, 371, 111))
        self.textEdit.setObjectName(_fromUtf8("textEdit"))
        self.line_3 = QtGui.QFrame(Dialog)
        self.line_3.setGeometry(QtCore.QRect(0, 120, 401, 16))
        self.line_3.setFrameShape(QtGui.QFrame.HLine)
        self.line_3.setFrameShadow(QtGui.QFrame.Sunken)
        self.line_3.setObjectName(_fromUtf8("line_3"))
        self.label_3 = QtGui.QLabel(Dialog)
        self.label_3.setGeometry(QtCore.QRect(0, 10, 161, 16))
        self.label_3.setObjectName(_fromUtf8("label_3"))
        self.lineEdit = QtGui.QLineEdit(Dialog)
        self.lineEdit.setGeometry(QtCore.QRect(20, 90, 291, 20))
        self.lineEdit.setObjectName(_fromUtf8("lineEdit"))
        self.pushButton = QtGui.QPushButton(Dialog)
        self.pushButton.setGeometry(QtCore.QRect(320, 90, 75, 23))
        self.pushButton.setObjectName(_fromUtf8("pushButton"))
        self.label = QtGui.QLabel(Dialog)
        self.label.setGeometry(QtCore.QRect(10, 136, 54, 31))
        self.label.setObjectName(_fromUtf8("label"))
        self.lineEdit_2 = QtGui.QLineEdit(Dialog)
        self.lineEdit_2.setGeometry(QtCore.QRect(30, 140, 113, 20))
        self.lineEdit_2.setObjectName(_fromUtf8("lineEdit_2"))
        self.label_2 = QtGui.QLabel(Dialog)
        self.label_2.setGeometry(QtCore.QRect(150, 130, 54, 41))
        self.label_2.setObjectName(_fromUtf8("label_2"))
        self.lineEdit_3 = QtGui.QLineEdit(Dialog)
        self.lineEdit_3.setGeometry(QtCore.QRect(190, 140, 113, 20))
        self.lineEdit_3.setObjectName(_fromUtf8("lineEdit_3"))
        self.pushButton_2 = QtGui.QPushButton(Dialog)
        self.pushButton_2.setGeometry(QtCore.QRect(320, 140, 75, 23))
        self.pushButton_2.setObjectName(_fromUtf8("pushButton_2"))
        self.label_4 = QtGui.QLabel(Dialog)
        self.label_4.setGeometry(QtCore.QRect(323, 140, 61, 21))
        self.label_4.setObjectName(_fromUtf8("label_4"))

        self.retranslateUi(Dialog)
        QtCore.QObject.connect(self.pushButton_2, QtCore.SIGNAL(_fromUtf8("clicked()")), self.progressBar.show)
        QtCore.QObject.connect(self.pushButton_2, QtCore.SIGNAL(_fromUtf8("clicked()")), self.label_4.show)
        QtCore.QObject.connect(self.pushButton_2, QtCore.SIGNAL(_fromUtf8("clicked()")), self.pushButton_2.hide)
        QtCore.QObject.connect(self.pushButton_2, QtCore.SIGNAL(_fromUtf8("clicked()")), self.subprocess)
        QtCore.QObject.connect(self.pushButton, QtCore.SIGNAL(_fromUtf8("clicked()")), self.showDialog)
        #QtCore.QObject.connect(self.radioButton, QtCore.SIGNAL(_fromUtf8("clicked()")), self.bwThread,setdoctype)
     #   QtCore.QObject.connect(self.radioButton_2, QtCore.SIGNAL(_fromUtf8("clicked()")), self.bwThread , setdoctype2)
        QtCore.QMetaObject.connectSlotsByName(Dialog)

    def retranslateUi(self, Dialog):
        Dialog.setWindowTitle(_translate("Dialog", "word文档批量替换软件", None))
        self.radioButton.setText(_translate("Dialog", ".doc（适用于Office2003及以下版本创建的文档）", None))
        self.radioButton_2.setText(_translate("Dialog", ".docx（适用于Office2007及以上版本创建的文档）", None))
        self.textEdit.setHtml(_translate("Dialog", "<!DOCTYPE HTML PUBLIC \"-//W3C//DTD HTML 4.0//EN\" \"http://www.w3.org/TR/REC-html40/strict.dtd\">\n"
"<html><head><meta name=\"qrichtext\" content=\"1\" /><style type=\"text/css\">\n"
"p, li { white-space: pre-wrap; }\n"
"</style></head><body style=\" font-family:\'SimSun\'; font-size:9pt; font-weight:400; font-style:normal;\">\n"
"<p style=\" margin-top:0px; margin-bottom:0px; margin-left:0px; margin-right:0px; -qt-block-indent:0; text-indent:0px;\">Word文档批量替换软件 1.0   </p>\n"
"<p style=\" margin-top:0px; margin-bottom:0px; margin-left:0px; margin-right:0px; -qt-block-indent:0; text-indent:0px;\">         ---Powered by Root lee</p></body></html>", None))
        self.textEdit.setText(self.txt)
        self.label_3.setText(_translate("Dialog", "   选择文件类型：", None))
        self.pushButton.setText(_translate("Dialog", "选择目录", None))
        self.label.setText(_translate("Dialog", "将", None))
        self.label_2.setText(_translate("Dialog", "替换成", None))
        self.pushButton_2.setText(_translate("Dialog", "确定", None))
        self.label_4.setText(_translate("Dialog", "正在替换..", None))
        self.label_4.hide()
        self.lineEdit_2.setText(u'')
        self.lineEdit_3.setText(u'')
        
    def subprocess(self):
        #from threads import BigWorkThread
        #新建对象

       # print str(self.lineEdit.text().toLocal8Bit())
        #line=[self.lineEdit.text(),self.lineEdit_2.text(),self.lineEdit_3.text()]
        if self.radioButton.isChecked():
            line=[1,str(self.lineEdit.text().toLocal8Bit()),str(self.lineEdit_2.text().toLocal8Bit()),str(self.lineEdit_3.text().toLocal8Bit())]
        elif self.radioButton_2.isChecked():
            line=[2,str(self.lineEdit.text().toLocal8Bit()),str(self.lineEdit_2.text().toLocal8Bit()),str(self.lineEdit_3.text().toLocal8Bit())]
        else:
            line=[]
        #print line[1]
        self.bwThread = BigWorkThread(line)
        #self.changecheck()
        #self.bwThread.whichischeck = 1
        #开始执行run()函数里的内容
        self.connect(self.bwThread,QtCore.SIGNAL("where"),self.update)
        self.connect(self.bwThread,QtCore.SIGNAL("showtxt"),self.showtxt)
        self.connect(self.bwThread,QtCore.SIGNAL("finish_show"),self.finish_show)
        self.connect(self.bwThread,QtCore.SIGNAL("finddocfile"),self.finddocfile)
        self.bwThread.start()
    def update(self,where):
        self.progressBar.setProperty("value", where)
    def finddocfile(self,i):
        self.txt = u"一共发现 %d 个文件..\n" %i
        self.textEdit.setText(self.txt)
    def finish_show(self,errorlist):
       
        
        self.txt = self.txt + u"- - - - - - - - - - - \n批量替换结束！\n"
        if errorlist:
            self.txt = self.txt + u"以下个文件替换时发生错误：\n"
        self.txt = self.txt +errorlist.decode('gbk')
        self.textEdit.setText(self.txt)
        self.label_4.setText(_translate("Dialog", "替换成功！", None))
    def showtxt(self,what):
 
        self.txt = self.txt + what.decode('gbk')+u"替换成功！\n"
        self.textEdit.setText(self.txt)
    def showDialog(self):

        filename = QtGui.QFileDialog.getExistingDirectory(self, 'Open file','/home')
        
        self.lineEdit.setText("%s" %filename)

app = QtGui.QApplication(sys.argv)
qb = Ui_Dialog()
qb.show()
sys.exit(app.exec_())
