import sys

from PyQt5 import QtWidgets, uic
from wb_an import WbAnalits
from ozon_an import OzonAnalits


class HomeScreen(QtWidgets.QMainWindow):
    def __init__(self):
        super().__init__()
        self.initUI()

    def initUI(self):
        uic.loadUi('uis/home.ui', self)
        self.pushButton.clicked.connect(self.anal_wb_week)
        self.pushButton_9.clicked.connect(self.anal_wb_week)
        self.pushButton_10.clicked.connect(self.anal_ozon_week)
        self.pushButton_14.clicked.connect(self.anal_wb_week)
        self.pushButton_2.clicked.connect(self.exit)

    def anal_wb_week(self):
        route = QtWidgets.QFileDialog.getOpenFileName(self, 'Open File', '/', "Excel (*.xls *.xlsx)")[0]
        if route:
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

    def exit(self):
        exit(sys.argv)
