import sys

from PyQt5 import QtWidgets, uic
from wb_an import WbAnalits
from ozon_an import OzonAnalits
from table_analit import analits_month, analits_wb
from month_an import MonthAnalits


class HomeScreen(QtWidgets.QMainWindow):
    def __init__(self):
        super().__init__()
        self.initUI()

    def initUI(self):
        uic.loadUi('uis/home2.ui', self)
        self.pushButton.clicked.connect(self.anal_wb_week)
        self.pushButton_10.clicked.connect(self.anal_ozon_week)
        self.pushButton_3.clicked.connect(self.openWB)
        self.pushButton_4.clicked.connect(self.openOzon)
        self.pushButton_2.clicked.connect(self.exit)
        self.pushButton_5.clicked.connect(self.to_month_res)
        self.routsOzon = []
        self.routsWB = []

    def anal_wb_week(self):
        route = QtWidgets.QFileDialog.getOpenFileName(self, 'Open File', '/', "Excel (*.xls *.xlsx)")[0]
        if route:
            try:
                analits_wb(route)
            except:
                self.label_6.setText('<p style="color: rgb(250, 55, 55);">Ошибка данных</p>')
                self.pushButton.setText('Открыть файлы ❌')

            else:
                self.pushButton.setText('Открыть файлы')
                self.label_6.setText('')
                self.st = WbAnalits(route, self)
                self.st.move(self.x() - 230, self.y() - 150)
                self.st.show()
                self.hide()

    def anal_ozon_week(self):
        route = QtWidgets.QFileDialog.getOpenFileName(self, 'Open File', '/', "Excel (*.xls *.xlsx)")[0]
        if route:
            self.st = OzonAnalits(route, self)
            self.st.move(self.x() - 230, self.y() - 150)
            self.st.show()
            self.hide()

    def openWB(self):
        self.routsWB = QtWidgets.QFileDialog.getOpenFileNames(self, 'Open File', 'results_wb/', "Excel (*.xls *.xlsx)")[
            0]
        try:
            a = 5 / len(self.routsWB)
            for route in self.routsWB:
                analits_month(route)
        except:
            self.label_4.setText('<p style="color: rgb(250, 55, 55);">Ошибка данных</p>')
            self.pushButton_3.setText('Открыть файлы ❌')
            self.pushButton_5.setEnabled(False)
        else:
            self.label_4.setText(f'Выбрано {len(self.routsWB)} файла(-ов)')
            self.pushButton_3.setText('Открыть файлы ✅')
            if self.pushButton_4.text() == 'Открыть файлы ✅':
                self.pushButton_5.setEnabled(True)

    def openOzon(self):
        self.routsOzon = \
            QtWidgets.QFileDialog.getOpenFileNames(self, 'Open File', 'results_ozon/', "Excel (*.xls *.xlsx)")[0]
        try:
            a = 5 / len(self.routsOzon)
            for route in self.routsOzon:
                analits_month(route)
        except:
            self.label_5.setText('<p style="color: rgb(250, 55, 55);">Ошибка данных</p>')
            self.pushButton_4.setText('Открыть файлы ❌')
            self.pushButton_5.setEnabled(False)
        else:
            self.label_5.setText(f'Выбрано {len(self.routsOzon)} файла(-ов)')
            self.pushButton_4.setText('Открыть файлы ✅')
            if self.pushButton_3.text() == 'Открыть файлы ✅':
                self.pushButton_5.setEnabled(True)

    def to_month_res(self):
        self.st = MonthAnalits(self.routsWB, self.routsOzon, self)
        self.st.move(self.x() - 230, self.y() - 150)
        self.st.show()
        self.label_4.setText('')
        self.label_5.setText('')
        self.pushButton_3.setText('Открыть файлы')
        self.pushButton_4.setText('Открыть файлы')
        self.pushButton_5.setEnabled(False)
        self.routsWB = []
        self.routsOzon = []
        self.hide()

    def exit(self):
        exit(sys.argv)
