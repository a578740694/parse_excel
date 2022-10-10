from datetime import datetime
import logging
import os.path
import sys
import threading
import time
import warnings
from concurrent import futures
from concurrent.futures import ThreadPoolExecutor

import pandas as pd
from PyQt6 import QtWidgets, QtGui, QtCore
from PyQt6.QtCore import Qt
from PyQt6.QtWidgets import QApplication, QHBoxLayout, QTableWidget, QHeaderView, \
    QAbstractItemView, QMessageBox, QProgressDialog
from PyQt6.QtWidgets import (QWidget, QLabel, QGridLayout, QLineEdit, QPushButton,
                             QCheckBox)


class GridLayout(QWidget):

    __directory = None

    __sheet = None

    __filterFlag = False

    __isSuccess = False

    def __init__(self):
        super().__init__()
        warnings.filterwarnings('ignore')

        #初始化界面
        self.initUI()

    def initUI(self):

        #resize()方法调整窗口的大小。宽,高
        self.resize(500, 300)
        #move()方法移动窗口在屏幕上的位置到x = 300，y = 300坐标。
        self.move(600, 300)
        #设置窗口的标题
        self.setWindowTitle('Excel两区域对比工具')

        layout = QGridLayout()
        layout.setContentsMargins(40,40,40,40)
        #layout.setAlignment(QtCore.Qt.AlignRight|QtCore.Qt.AlignVCenter)
        #layout.setSpacing(0)
        #layout.setVerticalSpacing(0)
        #layout.setHorizontalSpacing(0)

        self.button = QPushButton("选择Excel文件:")
        self.button.clicked.connect(self.chooseFile)
        self.button.setFixedSize(95, 30)
        layout.addWidget(self.button, 1, 0)

        self.qLabel0 = QLabel()
        layout.addWidget(self.qLabel0,1,1,1,4)

        qLabel1 = QLabel("Excel 工作表:")
        layout.addWidget(qLabel1,2,0)

        self.sheetNum = QLineEdit()
        self.sheetNum.setFixedSize(250, 20)
        self.sheetNum.textChanged.connect(self.textChange)
        layout.addWidget(self.sheetNum,2,1,1,4)

        qLabel4 = QLabel("是否需要筛选内容:")
        layout.addWidget(qLabel4,3,0)

        self.buttonConfirm = QPushButton("筛选")
        self.buttonConfirm.clicked.connect(self.filterClick)
        layout.addWidget(self.buttonConfirm,3,1)

        self.qLabel5 = QLabel("已筛选")
        self.qLabel5.setStyleSheet("color:green")
        self.qLabel5.setHidden(True)
        layout.addWidget(self.qLabel5,3,3,1,2)

        self.buttonConfirm = QPushButton("确认")
        self.buttonConfirm.clicked.connect(self.buttonClick)
        layout.addWidget(self.buttonConfirm,4,2)

        self.setLayout(layout)
        #显示在屏幕上
        self.show()

    def chooseFile(self):
        self.__directory = None
        self.__sheet = None
        self.__filterFlag = False
        self.__isSuccess = False

        self.__directory = QtWidgets.QFileDialog.getOpenFileName(self, "选取文件","./", "Excel Files (*.xls | *.xlsx)")
        self.qLabel0.setText(os.path.split(self.__directory[0])[1])

    def textChange(self):
        self.__sheet = None
        self.__filterFlag = False
        self.__isSuccess = False
        self.qLabel5.setHidden(True)

    def filterClick(self):
        if self.check():
            self.child1 = ChildWin1(self, self.__sheet, '筛选')

    def signal(self):
        self.__filterFlag = True
        self.qLabel5.setHidden(False)

    def buttonClick(self):
        try:
            if self.check():
                if self.__filterFlag:
                    filterData = self.child1.child2.retData()

                    for k,v in filterData.items():
                        self.__sheet = self.__sheet[self.__sheet[k].isin(v)]

                self.child3 = ChildWin1(self, self.__sheet, '比对')
                self.child3.signal.connect(self.comparison)
        except Exception as e:
            logging.exception(e)

    def run(self, col):
        return self.__sheet.groupby(col)[col].count()

    def work(self):
        comparison = self.child1.retData()

        #使用线程池
        with ThreadPoolExecutor(max_workers=5) as executor:
            all_task = [executor.submit(self.run, (value)) for value in comparison]
            done_iter = futures.as_completed(all_task)

            try:
                counts = []
                for done in done_iter:
                    result = done.result()
                    counts.append(result)

                diff = []
                for k, v in counts[0].items():
                    if (k not in counts[1] or counts[1][k] != v) and k not in diff:
                        diff.append(k)

                for k, v in counts[1].items():
                    if (k not in counts[0] or counts[0][k] != v) and k not in diff:
                        diff.append(k)

                time_tup = datetime.now().strftime('%Y%m%d%H%M%S%f')[:-3]
                pd.DataFrame({'差异': diff}).to_excel('差异数据_' + time_tup + '.xlsx')

                self.__isSuccess = True
            except Exception as e:
                self.progress.close()
                logging.exception(e)

    def comparison(self):
        try:
            thread_obj = threading.Thread(target=self.work)
            thread_obj.start()

            self.progress = QProgressDialog(self)
            self.progress.setWindowTitle("请稍等")
            self.progress.setLabelText("正在比对...")
            self.progress.setCancelButtonText("取消")
            self.progress.setMinimumDuration(2000)
            self.progress.setWindowModality(Qt.WindowModality.WindowModal)
            self.progress.setRange(0,100)

            i = 0
            while(True):
                if self.__isSuccess:
                    break

                time.sleep(0.2)
                self.progress.setValue(i)
                if self.progress.wasCanceled():
                    QMessageBox.warning(self,"提示","操作失败")
                    break

                if i < 99:
                    i = i + 1

            if self.__isSuccess:
                self.progress.setValue(100)
                self.progress.close()
                QMessageBox.information(self,"提示","操作成功")
        except Exception as e:
            logging.exception(e)

    def check(self):
        sheetNum = self.sheetNum.text().strip()

        if self.__directory is None:
            QMessageBox.warning(self, "提示", "请选择文件！！")
            return False

        if len(sheetNum) == 0:
            QMessageBox.warning(self, "提示", "请输入工作表！！")
            return False

        if self.__sheet is None:
            try:
                #解析需要筛选的列和值
                self.__sheet = pd.read_excel(self.__directory[0], sheet_name=int(sheetNum) - 1, dtype=str)

            except Exception as e:
                logging.exception(e)
                QMessageBox.warning(self, "提示", "所选工作表无数据！！")
                return False

        return True

    def closeEvent(self, a0: QtGui.QCloseEvent) -> None:
        sys.exit(0)

class ChildWin1(QtWidgets.QDialog):

    signal = QtCore.pyqtSignal()

    __main = None

    __sheet = None

    __filterBox = {}

    __title = None

    def __init__(self, main, sheet, title):
        self.__sheet = sheet
        self.__main = main
        self.__title = title

        try:
            super().__init__()
            self.setWindowTitle("选择需要" + title + "的列")
            self.resize(400, 200)

            layout = QGridLayout()

            x = 1
            y = 0
            skip = 0
            for i, v in enumerate(self.__sheet.columns.values):
                if v.startswith("Unnamed"):
                    skip = skip + 1
                    continue

                i = i - skip
                self.__filterBox[i] = QCheckBox(v)
                self.__filterBox[i].setCheckState(Qt.CheckState.Unchecked)

                if i != 0 and i % 4 == 0:
                    x = x + 1
                    y = 0

                layout.addWidget(self.__filterBox[i], x, y)
                y = y + 1

            self.buttonConfirm = QPushButton("确认")
            self.buttonConfirm.clicked.connect(self.submit)
            layout.addWidget(self.buttonConfirm,x + 2,0)

            self.setLayout(layout)
            self.show()
        except Exception as e:
            logging.exception(e)

    def submit(self):
        if self.__title == "比对":
            count = 0
            for k,v in self.__filterBox.items():
                if v.isChecked():
                    count = count + 1

            if count != 2:
                QtWidgets.QMessageBox.warning(self, "提示", "比对的列只能选择两列！！")
                return

            self.close()
            self.signal.emit()
        elif self.__title == "筛选":
            self.close()
            self.child2 = ChildWin2(self.__sheet, self.__filterBox)
            self.child2.signal.connect(self.__main.signal)
        else:
            raise Exception("类型错误")

    def retData(self):
        comparison = []
        for k,v in self.__filterBox.items():
            if v.isChecked():
                comparison.append(v.text())

        return comparison

class ChildWin2(QtWidgets.QDialog):

    signal = QtCore.pyqtSignal()

    __filterText = []

    __selected = {}

    def __init__(self, sheet, filterBox):
        try:
            self.__filterText = []
            self.__selected = {}

            for i, value in enumerate(filterBox.values()):
                if value.isChecked():
                    self.__filterText.append(value.text())

            super().__init__()
            self.setWindowTitle("筛选")
            self.resize(500, 300)

            layout = QHBoxLayout()

            self.buttonConfirm = QPushButton("确认")
            self.buttonConfirm.clicked.connect(self.submit)
            layout.addWidget(self.buttonConfirm)

            #表格对象
            self.tableWidget = QTableWidget()
            self.tableWidget.setSelectionMode(QAbstractItemView.SelectionMode.MultiSelection)
            self.tableWidget.setSelectionBehavior(QAbstractItemView.SelectionBehavior.SelectItems)
            self.tableWidget.horizontalHeader().setSectionResizeMode(QHeaderView.ResizeMode.Interactive)
            self.tableWidget.setColumnCount(len(self.__filterText))

            rowCount = 0
            groups = []
            for v in self.__filterText:
                group = sheet[v].unique()
                groups.append(group)

                if rowCount < len(group):
                    rowCount = len(group)

            self.tableWidget.setRowCount(rowCount)
            for i in range(rowCount):
                for j in range(0, len(groups)):
                    value = '' if len(groups[j]) <= i else str(groups[j][i])

                    self.tableWidget.setItem(i, j, QtWidgets.QTableWidgetItem(value))

                    if value == '':
                        item = self.tableWidget.item(i,j)
                        item.setFlags(item.flags() & ~Qt.ItemFlag.ItemIsSelectable)

            #设置表格字段
            self.tableWidget.setHorizontalHeaderLabels(self.__filterText)
            self.tableWidget.setEditTriggers(QtWidgets.QAbstractItemView.EditTrigger.NoEditTriggers)
            layout.addWidget(self.tableWidget)

            self.setLayout(layout)
            self.show()
        except Exception as e:
            logging.exception(e)

    def submit(self):
        self.__selected = {}
        selected = self.tableWidget.selectionModel()
        indexs = selected.selectedIndexes()

        for index in indexs:
            if index.data() == None or len(index.data().strip()) == 0:
                break

            if self.__filterText[index.column()] not in self.__selected:
                self.__selected[self.__filterText[index.column()]] = []

            self.__selected[self.__filterText[index.column()]].append(index.data())

        self.close()
        self.signal.emit()

    def retData(self):
        return self.__selected

if __name__ == '__main__':
    #每一pyqt5应用程序必须创建一个应用程序对象。sys.argv参数是一个列表，从命令行输入参数。
    app = QApplication(sys.argv)
    #QWidget部件是pyqt5所有用户界面对象的基类。他为QWidget提供默认构造函数。默认构造函数没有父类。
    w = GridLayout()
    sys.exit(app.exec())