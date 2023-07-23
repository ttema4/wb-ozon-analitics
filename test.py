import sys
from PyQt5.QtWidgets import QApplication, QMainWindow, QTableWidget, QTableWidgetItem
from PyQt5.QtCore import QSize


class MainWindow(QMainWindow):
    def __init__(self):
        super().__init__()
        self.tableWidget = QTableWidget(self)

        self.tableWidget.setColumnCount(4)
        self.tableWidget.setRowCount(1)
        self.tableWidget.setMinimumSize(QSize(400, 600))

        self.tableWidget.setItem(0, 1, QTableWidgetItem('rnejnre'))
        self.tableWidget.setItem(0, 2, QTableWidgetItem('dlmfke'))
        self.tableWidget.setItem(0, 3, QTableWidgetItem('keji'))



if __name__ == "__main__":
    app = QApplication(sys.argv)
    mw = MainWindow()
    mw.show()
    sys.exit(app.exec_())
