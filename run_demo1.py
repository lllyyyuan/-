from PyQt5.QtCore import pyqtSlot
from PyQt5.QtWidgets import QWidget, QMessageBox, QSpinBox, QApplication,QFileDialog, QTableWidgetItem, QDialog, QPushButton, QTableWidget

from PyQt5.QtGui import QImage, QPixmap
import sys
import pandas as pd
from demo import Ui_Form
from PyQt5.QtCore import QTimer
from PyQt5.QtCore import Qt
from PyQt5.uic import loadUi
from openpyxl.drawing.image import Image as xlImage
from PyQt5.QtGui import QImage


class myForm(QWidget, Ui_Form):
    def __init__(self):
        super(myForm, self).__init__()
        #loadUi("demo.ui", self)
        self.setupUi(self)
        self.file_path = ""
        self.df = pd.DataFrame()
        self.table_widget.clicked.connect(self.show_text)  #绑定单击tablewidget事件
        self.pushButton_1.clicked.connect(self.Manual_execution)  #手动执行按钮绑定事件
        self.pushButton_2.clicked.connect(self.Automatic_execution)  #自动执行按钮绑定事件
        if self.table_widget.rowCount() == 0:
            self.pushButton_1.setEnabled(False)
            self.pushButton_2.setEnabled(False)


    def browse_file(self):
        """
        Open a file dialog to browse Excel files.
        """
        file_name, _ = QFileDialog.getOpenFileName(self, "Open Excel File", "", "Excel Files (*.xlsx)")
        if file_name:
            self.file_path = file_name
            self.file_path_line_edit.setText(self.file_path)
            #把选中的文件给了lineedit输入框
    def read_excel(self):
        """
        Read the Excel file using Pandas and display the data in the table widget.
        """
        if not self.file_path:
            QMessageBox.critical(self, "Error", "Please select an Excel file to read.")
            return
        try:
            self.df = pd.read_excel(self.file_path)
            self.table_widget.setRowCount(self.df.shape[0])
            self.table_widget.setColumnCount(self.df.shape[1])
            self.table_widget.setHorizontalHeaderLabels(self.df.columns)
            for i in range(self.df.shape[0]):
                for j in range(self.df.shape[1]):
                    self.table_widget.setItem(i, j, QTableWidgetItem(str(self.df.iloc[i, j])))
        except Exception as e:
            QMessageBox.critical(self, "Error", "An error occurred while reading the Excel file.")
            #QMessageBox.critical(self, "Error", "An error occurred while reading the Excel file.\n\n{e}")

    def show_text(self):
         self.pushButton_1.setEnabled(True)
         self.pushButton_2.setEnabled(True)
         row = self.table_widget.currentRow()               #获取表格的当前行赋给row
         content_1 = self.table_widget.item(row, 0).text()  # 获取选中行的第1列数据
         self.textEdit_2.setText(content_1)  # 显示内容
         content_2 = self.table_widget.item(row, 1).text()  # 获取选中行的第2列数据
         self.textEdit_3.setText(content_2)  # 显示内容
         content_3 = self.table_widget.item(row, 2).text()  # 获取选中行的第3列数据
         self.textEdit_4.setText(content_3)  # 显示内容
         content_4 = self.table_widget.item(row, 3).text()  # 获取选中行的第4列数据
         self.textEdit_5.setText(content_4)  # 显示内容

         path = self.table_widget.item(row, 4).text()   #预期输出
         # 加载并显示图片
         pixmap = QPixmap(path)
         self.label_1.setPixmap(pixmap)
         # 调整 QLabel 的尺寸
         #self.label_1.resize(pixmap.width(), pixmap.height())
         self.label_8.clear()  #清空实际结果

    def Manual_execution(self):    #手动执行

        # 清空标签中的图片
        #self.label_8.clear()
        # 创建一个定时器，用于在每次循环后等待1秒钟
        self.timer = QTimer()
        self.timer.setInterval(1000)
        self.timer.timeout.connect(self.show_next_image)

        # 显示第一张图像
        self.image_index = 0
        # self.show_next_image()

        # 启动定时器
        self.timer.start()

    def show_next_image(self):
        row = self.table_widget.currentRow()  # 获取表格的当前行赋给row
        path = self.table_widget.item(row, 4).text()  # 获取选中行的第4列路径
        self.label_8.setText("加载中...")   # 显示“加载中...”文本
        self.label_8.setPixmap(QPixmap(path))  # 显示图片
        # 清除“加载中...”文本
        self.label_8.setText("")

        # 加载下一张图像
        #pixmap = QPixmap('D:\Pycharm\image\IMG{}.jpeg'.format(self.image_index))
        # 显示图像
        #self.label_8.setPixmap(pixmap)

        # 增加图片索引
        self.image_index += 1

        # 如果所有图像都已经显示完毕，则停止定时器
        if self.image_index == self.spinBox.value():
            self.timer.stop()
            QMessageBox.information(self, "提示", "获取了 {} 张图片".format(self.image_index))

    # 此处的Automatic_execution为设置的触发槽名字
    def Automatic_execution(self):      #自动执行

        if self.pushButton_2.text() == '自动执行':
            self.pushButton_2.setText('暂停')
            # 执行自动执行的代码

            self.timer = QTimer()
            self.timer.timeout.connect(self.onTimeout)
            self.timer.start(1000)  # 1秒执行一次
            self.case_index = 0
            self.current_loop = 0  # 当前已经执行的循环次数

        else:
            self.pushButton_2.setText('自动执行')
        # 执行暂停的代码
            self.timer.stop()
            QMessageBox.information(self, "提示", "自动执行暂停，用例总数:{}".format(self.case_index))


    def onTimeout(self):
        # 获取当前选中单元格的行号和列号
        current_row = self.table_widget.currentRow()
        current_column = self.table_widget.currentColumn()
        self.show_text()
        path = self.table_widget.item(current_row, 4).text()  # 获取选中行的第4列路径
        self.label_8.setPixmap(QPixmap(path))  # 显示图片
        # 增加case索引
        self.case_index += 1

        if current_row < self.table_widget.rowCount() - 1:
            self.table_widget.setCurrentCell(current_row + 1, current_column)
            # 如果当前选中的单元格不是最后一行，则将选中状态移动到下一行
        else:
            #self.current_loop = 0  # 重置已经执行的循环次数
            self.current_loop += 1  #循环次数加1
            print(self.current_loop)
            if self.current_loop >= self.spinBox_2.value():
                self.timer.stop()
                self.pushButton_2.setEnabled(True)
                self.pushButton_2.setText('自动执行')
                QMessageBox.information(self, "提示", "自动执行结束\n执行用例总数{}条\n共循环执行了{}遍".format(self.case_index,self.current_loop))
            else:
                self.table_widget.setCurrentCell(0, current_column)  # 跳到第一行






if __name__ == "__main__":
    app = QApplication(sys.argv)
    form = myForm()
    form.show()
    sys.exit(app.exec_())