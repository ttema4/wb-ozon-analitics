from PyQt5.QtWidgets import QApplication
from PyQt5 import QtGui
from home import HomeScreen
import sys

app = QApplication(sys.argv)
app.setWindowIcon(QtGui.QIcon('analizer.ico'))
ex = HomeScreen()
ex.setWindowIcon(QtGui.QIcon('analizer.ico'))
ex.show()
sys.exit(app.exec())
