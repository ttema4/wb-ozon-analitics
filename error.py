from PyQt5 import QtWidgets, uic


class ErrorWidget(QtWidgets.QMainWindow):
    def __init__(self, parent=None):
        super().__init__()
        self.parent = parent

    def initUI(self):
        uic.loadUi('home.ui', self)
        self.pushButton.clicked.connect(self.back)

    def back(self):
        self.parent.move(self.x() - 80, self.y() - 50)
        self.parent.show()
        self.hide()
