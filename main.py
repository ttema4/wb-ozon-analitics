from PyQt5.QtWidgets import QApplication
from home import HomeScreen
import sys

app = QApplication(sys.argv)
ex = HomeScreen()
ex.show()
sys.exit(app.exec())
