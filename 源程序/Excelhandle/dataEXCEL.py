# -*- coding: utf-8 -*-

# Form implementation generated from reading ui file 'dataEXCEL.ui'
#
# Created by: PyQt5 UI code generator 5.11.2
#
# WARNING! All changes made in this file will be lost!


from PyQt5 import QtCore, QtGui
from PyQt5 import QtWidgets
from PyQt5.QtWidgets import qApp, QFileDialog

import sys
import pandas as pd
import os
import glob
import numpy as np
import matplotlib.pyplot as plt

root = ""
fileNum = 0
myrow = 0


def resource_path(relative_path):
    """ Get absolute path to resource, works for dev and for PyInstaller """
    try:
        # PyInstaller creates a temp folder and stores path in _MEIPASS
        base_path = sys._MEIPASS
    except Exception:
        base_path = os.path.abspath(".")

    return os.path.join(base_path, relative_path)

# 自定义函数SaveExcel用于保存数据到Excel


def SaveExcel(df, isChecked):
    # 将提取后的数据保存到Excel
    if isChecked:
        writer = pd.ExcelWriter("mycell.xlsx", engine="openpyxl")
    else:
        global temproot
        writer = pd.ExcelWriter(os.path.join(
            temproot, "mycell.xlsx"), engine="openpyxl")
    df.to_excel(writer, "sheet1", index=False)
    writer.close()


class Ui_MainWindow(QtWidgets.QWidget):
    def setupUi(self, MainWindow):
        MainWindow.setObjectName("MainWindow")
        MainWindow.resize(838, 596)
        self.centralwidget = QtWidgets.QWidget(MainWindow)
        self.centralwidget.setObjectName("centralwidget")

        # 创建主布局
        main_layout = QtWidgets.QVBoxLayout(self.centralwidget)

        # 创建上半部分水平布局（包含列表和文本显示区域）
        top_layout = QtWidgets.QHBoxLayout()

        # 左侧列表区域
        self.list1 = QtWidgets.QListView()
        self.list1.setObjectName("list1")
        self.list1.setFixedWidth(170)

        # 右侧文本显示区域
        self.textEdit = QtWidgets.QTextEdit()
        self.textEdit.setObjectName("textEdit")
        self.textEdit.setHorizontalScrollBarPolicy(QtCore.Qt.ScrollBarAlwaysOn)

        top_layout.addWidget(self.list1)
        top_layout.addWidget(self.textEdit, 4)

        # 创建下半部分布局
        bottom_layout = QtWidgets.QHBoxLayout()

        # 左侧单选按钮区域
        radio_layout = QtWidgets.QVBoxLayout()
        self.label = QtWidgets.QLabel()
        self.label.setEnabled(False)
        font = QtGui.QFont()
        font.setPointSize(10)
        font.setBold(True)
        font.setWeight(75)
        self.label.setFont(font)
        self.label.setObjectName("label")

        self.rButton1 = QtWidgets.QRadioButton()
        font = QtGui.QFont()
        font.setBold(True)
        font.setWeight(75)
        self.rButton1.setFont(font)
        self.rButton1.setObjectName("rButton1")

        self.rButton2 = QtWidgets.QRadioButton()
        font = QtGui.QFont()
        font.setBold(True)
        font.setWeight(75)
        self.rButton2.setFont(font)
        self.rButton2.setCheckable(True)
        self.rButton2.setChecked(True)
        self.rButton2.setObjectName("rButton2")

        radio_layout.addWidget(self.label)
        radio_layout.addWidget(self.rButton1)
        radio_layout.addWidget(self.rButton2)
        radio_layout.addStretch()

        # 右侧路径选择区域
        path_layout = QtWidgets.QHBoxLayout()
        self.text1 = QtWidgets.QTextEdit()
        sizePolicy = QtWidgets.QSizePolicy(
            QtWidgets.QSizePolicy.Expanding, QtWidgets.QSizePolicy.Fixed)
        sizePolicy.setHeightForWidth(
            self.text1.sizePolicy().hasHeightForWidth())
        self.text1.setSizePolicy(sizePolicy)
        self.text1.setMaximumHeight(30)
        self.text1.setObjectName("text1")

        self.viewButton = QtWidgets.QPushButton()
        self.viewButton.setObjectName("viewButton")
        self.viewButton.setMaximumWidth(100)
        self.viewButton.setMaximumHeight(30)

        path_layout.addWidget(self.text1)
        path_layout.addWidget(self.viewButton)

        bottom_layout.addLayout(radio_layout, 1)
        bottom_layout.addLayout(path_layout, 4)

        main_layout.addLayout(top_layout, 7)
        main_layout.addLayout(bottom_layout, 1)

        MainWindow.setCentralWidget(self.centralwidget)

        self.menubar = QtWidgets.QMenuBar(MainWindow)
        self.menubar.setGeometry(QtCore.QRect(0, 0, 838, 23))
        self.menubar.setObjectName("menubar")
        MainWindow.setMenuBar(self.menubar)
        self.statusbar = QtWidgets.QStatusBar(MainWindow)
        self.statusbar.setObjectName("statusbar")
        MainWindow.setStatusBar(self.statusbar)
        self.toolBar = QtWidgets.QToolBar(MainWindow)
        self.toolBar.setEnabled(True)
        self.toolBar.setSizePolicy(QtWidgets.QSizePolicy(
            QtWidgets.QSizePolicy.Preferred, QtWidgets.QSizePolicy.Fixed))
        self.toolBar.setIconSize(QtCore.QSize(48, 48))
        self.toolBar.setToolButtonStyle(QtCore.Qt.ToolButtonTextUnderIcon)
        self.toolBar.setFloatable(False)
        self.toolBar.setObjectName("toolBar")
        MainWindow.addToolBar(QtCore.Qt.TopToolBarArea, self.toolBar)

        self.button1 = QtWidgets.QAction(MainWindow)
        icon = QtGui.QIcon.fromTheme("导入EXCEL")
        icon.addPixmap(QtGui.QPixmap(resource_path(
            "image/图标-01.png")), QtGui.QIcon.Normal, QtGui.QIcon.Off)
        self.button1.setIcon(icon)
        self.button1.setObjectName("button1")
        self.button2 = QtWidgets.QAction(MainWindow)
        icon = QtGui.QIcon()
        icon.addPixmap(QtGui.QPixmap(resource_path(
            "image/图标-02.png")), QtGui.QIcon.Normal, QtGui.QIcon.Off)
        self.button2.setIcon(icon)
        self.button2.setObjectName("button2")
        self.button3 = QtWidgets.QAction(MainWindow)
        icon1 = QtGui.QIcon()
        icon1.addPixmap(QtGui.QPixmap(resource_path(
            "image/图标-03.png")), QtGui.QIcon.Normal, QtGui.QIcon.Off)
        self.button3.setIcon(icon1)
        self.button3.setObjectName("button3")
        self.button4 = QtWidgets.QAction(MainWindow)
        icon2 = QtGui.QIcon()
        icon2.addPixmap(QtGui.QPixmap(resource_path(
            "image/图标-04.png")), QtGui.QIcon.Normal, QtGui.QIcon.Off)
        self.button4.setIcon(icon2)
        self.button4.setObjectName("button4")
        self.button5 = QtWidgets.QAction(MainWindow)
        icon3 = QtGui.QIcon()
        icon3.addPixmap(QtGui.QPixmap(resource_path(
            "image/图标-05.png")), QtGui.QIcon.Normal, QtGui.QIcon.Off)
        self.button5.setIcon(icon3)
        self.button5.setObjectName("button5")
        self.button6 = QtWidgets.QAction(MainWindow)
        icon4 = QtGui.QIcon()
        icon4.addPixmap(QtGui.QPixmap(resource_path(
            "image/图标-06.png")), QtGui.QIcon.Normal, QtGui.QIcon.Off)
        self.button6.setIcon(icon4)
        self.button6.setObjectName("button6")
        self.button7 = QtWidgets.QAction(MainWindow)
        icon5 = QtGui.QIcon()
        icon5.addPixmap(QtGui.QPixmap(resource_path(
            "image/图标-07.png")), QtGui.QIcon.Normal, QtGui.QIcon.Off)
        self.button7.setIcon(icon5)
        self.button7.setObjectName("button7")

        self.toolBar.addSeparator()
        self.toolBar.addAction(self.button1)
        self.toolBar.addAction(self.button2)
        self.toolBar.addSeparator()
        self.toolBar.addAction(self.button3)
        self.toolBar.addAction(self.button4)
        self.toolBar.addAction(self.button5)
        self.toolBar.addAction(self.button6)
        self.toolBar.addSeparator()
        self.toolBar.addAction(self.button7)
        self.toolBar.addSeparator()

        self.button7.triggered.connect(qApp.quit)
        self.button1.triggered.connect(self.click1)
        self.button2.triggered.connect(self.click2)
        self.button3.triggered.connect(self.click3)
        self.button4.triggered.connect(self.click4)
        self.button5.triggered.connect(self.click5)
        self.button6.triggered.connect(self.click6)
        self.viewButton.clicked.connect(self.viewButton_click)
        self.list1.clicked.connect(self.clicked)

        pd.set_option("display.max_columns", None)
        pd.set_option("max_colwidth", 200)

        self.retranslateUi(MainWindow)
        QtCore.QMetaObject.connectSlotsByName(MainWindow)

        # 用来保存文件列表
        self.file_list = []
    def retranslateUi(self, MainWindow):
        _translate = QtCore.QCoreApplication.translate
        # MainWindow.setWindowTitle(_translate("MainWindow", "Excel数据分析师"))
        # self.viewButton.setText(_translate("MainWindow", "浏览"))
        # self.rButton1.setText(_translate("MainWindow", "自定义文件夹"))
        # self.label.setText(_translate("MainWindow", "输出选项"))
        # self.rButton2.setText(_translate("MainWindow", "保存在原文件夹内"))
        # self.toolBar.setWindowTitle(_translate("MainWindow", "toolBar"))
        # self.button1.setText(_translate("MainWindow", "导入EXCEL"))
        # self.button2.setText(_translate("MainWindow", "提取列数据"))
        # self.button3.setText(_translate("MainWindow", "定向筛选"))
        # self.button3.setToolTip(_translate("MainWindow", "定向筛选"))
        # self.button4.setText(_translate("MainWindow", "多表合并"))
        # self.button5.setText(_translate("MainWindow", "多表统计排行"))
        # self.button5.setToolTip(_translate("MainWindow", "多表统计排行"))
        # self.button6.setText(_translate("MainWindow", "生成图表"))
        # self.button6.setToolTip(_translate("MainWindow", "生成图表"))
        # self.button7.setText(_translate("MainWindow", "退出"))
        # self.button7.setToolTip(_translate("MainWindow", "退出"))

        MainWindow.setWindowTitle(_translate("MainWindow", "Excel Data Analyst"))
        self.viewButton.setText(_translate("MainWindow", "Browse"))
        self.rButton1.setText(_translate("MainWindow", "Custom Folder"))
        self.label.setText(_translate("MainWindow", "Output Options"))
        self.rButton2.setText(_translate("MainWindow", "Save in Original Folder"))
        self.toolBar.setWindowTitle(_translate("MainWindow", "toolBar"))
        self.button1.setText(_translate("MainWindow", "Import Excel"))
        self.button2.setText(_translate("MainWindow", "Extract Columns"))
        self.button3.setText(_translate("MainWindow", "Targeted Filter"))
        self.button3.setToolTip(_translate("MainWindow", "Targeted Filter"))
        self.button4.setText(_translate("MainWindow", "Merge Tables"))
        self.button5.setText(_translate("MainWindow", "Multi-table Statistics"))
        self.button5.setToolTip(_translate("MainWindow", "Multi-table Statistics"))
        self.button6.setText(_translate("MainWindow", "Generate Chart"))
        self.button6.setToolTip(_translate("MainWindow", "Generate Chart"))
        self.button7.setText(_translate("MainWindow", "Exit"))
        self.button7.setToolTip(_translate("MainWindow", "Exit"))

    def click1(self):
        global root
        root = QFileDialog.getExistingDirectory(self, "选择文件夹", "/")
        mylist = []
        for dirpath, dirnames, filenames in os.walk(root):
            for filepath in filenames:
                mylist.append(os.path.join(filepath))
        self.model = QtCore.QStringListModel()
        self.model.setStringList(mylist)
        self.list1.setModel(self.model)

        # 保存文件名列表
        self.file_list = mylist

    def clicked(self, qModelIndex):
        global root, myrow
        myrow = qModelIndex.row()
        a = os.path.join(root, str(self.file_list[qModelIndex.row()]))
        df = pd.DataFrame(pd.read_excel(a))
        self.textEdit.setText(str(df))

    def click2(self):
        global root, myrow
        a = os.path.join(root, str(self.file_list[myrow]))
        df = pd.DataFrame(pd.read_excel(a))
        df1 = df[["买家会员名", "收货人姓名", "联系手机", "宝贝标题"]]
        self.textEdit.setText(str(df1))
        SaveExcel(df1, self.rButton2.isChecked())

    def click3(self):
        global root
        filearray = glob.glob(os.path.join(root, "*.xls"))
        res = pd.read_excel(filearray[0])
        for i in range(1, len(filearray)):
            A = pd.read_excel(filearray[i])
            res = pd.concat([res, A], ignore_index=False, sort=True)
        df1 = res[["买家会员名", "收货人姓名", "联系手机", "宝贝标题"]]
        df2 = df1.loc[df1["宝贝标题"] == "零基础学Python"]
        self.textEdit.setText(str(df2))
        SaveExcel(df2, self.rButton2.isChecked())

    def click4(self):
        global root
        filearray = glob.glob(os.path.join(root, "*.xls"))
        res = pd.read_excel(filearray[0])
        for i in range(1, len(filearray)):
            A = pd.read_excel(filearray[i])
            res = pd.concat([res, A], ignore_index=False, sort=True)
        self.textEdit.setText(str(res.index))
        SaveExcel(res, self.rButton2.isChecked())

    def click5(self):
        global root
        filearray = glob.glob(os.path.join(root, "*.xls"))
        res = pd.read_excel(filearray[0])
        for i in range(1, len(filearray)):
            A = pd.read_excel(filearray[i])
            res = pd.concat([res, A], ignore_index=False, sort=True)
        df = res.groupby(["宝贝标题"])["宝贝总数量"].sum().reset_index()
        df1 = df.sort_values(by="宝贝总数量", ascending=False)
        self.textEdit.setText(str(df1))
        SaveExcel(df1, self.rButton2.isChecked())

    def click6(self):
        global root
        filearray = glob.glob(os.path.join(root, "*.xls"))
        res = pd.read_excel(filearray[0])
        for i in range(1, len(filearray)):
            A = pd.read_excel(filearray[i])
            res = pd.concat([res, A], ignore_index=False, sort=True)
        df = res[(res.类别 == "全彩系列")]
        df1 = df.groupby(["图书编号"])["买家实际支付金额"].sum().reset_index()
        df1 = df1.set_index("图书编号")
        df1 = df1["买家实际支付金额"].copy()
        df2 = df1.sort_values(ascending=False)
        SaveExcel(df2, self.rButton2.isChecked())
        plt.rc("font", family="SimHei", size=10)
        plt.figure("贡献度分析")
        df2.plot(kind="bar")
        plt.ylabel("销售收入（元）")
        p = 1.0 * df2.cumsum() / df2.sum()
        p.plot(color="r", secondary_y=True, style="-o", linewidth=0.5)
        plt.annotate(
            format(p.iloc[9], ".4%"),
            xy=(9, p.iloc[9]),
            xytext=(9 * 0.9, p.iloc[9] * 0.9),
            arrowprops=dict(arrowstyle="->", connectionstyle="arc3,rad=.1"),
        )
        plt.ylabel("收入（比例）")
        plt.show()

    def viewButton_click(self):
        global temproot
        temproot = QFileDialog.getExistingDirectory(self, "选择文件夹", "/")
        self.text1.setText(temproot)


# 定义载入主窗体的方法
def show_MainWindow():
    app = QtWidgets.QApplication(sys.argv)
    MainWindow = QtWidgets.QMainWindow()
    ui = Ui_MainWindow()
    ui.setupUi(MainWindow)
    MainWindow.show()
    sys.exit(app.exec_())


if __name__ == "__main__":
    show_MainWindow()
    path = root