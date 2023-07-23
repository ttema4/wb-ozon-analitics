from PyQt5.Qt import *
from PyQt5 import QtWidgets, uic
from table_analit import analits_month
from another_funcs import be_nums


class MonthAnalits(QtWidgets.QMainWindow):
    def __init__(self, routsWB, rountOzon, parent=None):
        super().__init__()
        self.routsWB = routsWB
        self.routsOzon = rountOzon
        self.parent = parent
        self.resWB = []
        for el in self.routsWB:
            self.resWB.append(analits_month(el))
        self.resWB.append([sum([self.resWB[j][i] for j in range(len(self.resWB))]) for i in range(len(self.resWB[0]))])
        self.resOzon = []
        for el in self.routsOzon:
            self.resOzon.append(analits_month(el))
        self.resOzon.append(
            [sum([self.resOzon[j][i] for j in range(len(self.resOzon))]) for i in range(len(self.resOzon[0]))])
        self.initUI()

    def initUI(self):
        uic.loadUi('uis/month.ui', self)
        self.pushButton.clicked.connect(self.back)
        self.pushButton_2.clicked.connect(self.copy_clipboard)
        nums = [0, 1, 3, 6, 9]
        for i in range(len(nums)):
            self.tableWidget.setItem(nums[i], 0, QTableWidgetItem(be_nums(self.resOzon[-1][i])))
            self.tableWidget.setItem(nums[i], 1, QTableWidgetItem(be_nums(self.resWB[-1][i])))
            self.tableWidget.setItem(nums[i], 2, QTableWidgetItem(be_nums(self.resOzon[-1][i] + self.resWB[-1][i])))

        self.tableWidget.setItem(2, 0, QTableWidgetItem(be_nums(self.resOzon[-1][1] / self.resOzon[-1][0])))
        self.tableWidget.setItem(2, 1, QTableWidgetItem(be_nums(self.resWB[-1][1] / self.resWB[-1][0])))
        self.tableWidget.setItem(2, 2, QTableWidgetItem(
            be_nums((self.resOzon[-1][1] + self.resWB[-1][1]) / (self.resOzon[-1][0] + self.resWB[-1][0]))))

        self.tableWidget.setItem(4, 0, QTableWidgetItem(be_nums(self.resOzon[-1][2] / self.resOzon[-1][0])))
        self.tableWidget.setItem(4, 1, QTableWidgetItem(be_nums(self.resWB[-1][2] / self.resWB[-1][0])))
        self.tableWidget.setItem(4, 2, QTableWidgetItem(
            be_nums((self.resOzon[-1][2] + self.resWB[-1][2]) / (self.resOzon[-1][0] + self.resWB[-1][0]))))

        self.tableWidget.setItem(5, 0, QTableWidgetItem(
            be_nums(round((self.resOzon[-1][1] - self.resOzon[-1][2]) / self.resOzon[-1][1] * 100)) + '%'))
        self.tableWidget.setItem(5, 1, QTableWidgetItem(
            be_nums(round((self.resWB[-1][1] - self.resWB[-1][2]) / self.resWB[-1][1] * 100)) + '%'))
        self.tableWidget.setItem(5, 2, QTableWidgetItem(
            be_nums(round(((self.resOzon[-1][1] + self.resWB[-1][1]) - (self.resOzon[-1][2] + self.resWB[-1][2])) / (
                    self.resOzon[-1][1] + self.resWB[-1][1]) * 100)) + '%'))

        self.tableWidget.setItem(7, 0, QTableWidgetItem(
            be_nums(round(self.resOzon[-1][3] / self.resOzon[-1][1] * 100)) + '%'))
        self.tableWidget.setItem(7, 1,
                                 QTableWidgetItem(be_nums(round(self.resWB[-1][3] / self.resWB[-1][1] * 100)) + '%'))
        self.tableWidget.setItem(7, 2, QTableWidgetItem(
            be_nums(round(
                (self.resOzon[-1][3] + self.resWB[-1][3]) / (self.resOzon[-1][1] + self.resWB[-1][1]) * 100)) + '%'))

        self.tableWidget.setItem(8, 0, QTableWidgetItem(be_nums(self.resOzon[-1][3] / self.resOzon[-1][0])))
        self.tableWidget.setItem(8, 1, QTableWidgetItem(be_nums(self.resWB[-1][3] / self.resWB[-1][0])))
        self.tableWidget.setItem(8, 2, QTableWidgetItem(
            be_nums((self.resOzon[-1][3] + self.resWB[-1][3]) / (self.resOzon[-1][0] + self.resWB[-1][0]))))

        self.tableWidget.setItem(10, 0, QTableWidgetItem(be_nums(self.resOzon[-1][4] / self.resOzon[-1][0])))
        self.tableWidget.setItem(10, 1, QTableWidgetItem(be_nums(self.resWB[-1][4] / self.resWB[-1][0])))
        self.tableWidget.setItem(10, 2, QTableWidgetItem(
            be_nums((self.resOzon[-1][4] + self.resWB[-1][4]) / (self.resOzon[-1][0] + self.resWB[-1][0]))))
        piece = [self.resOzon[-1][4] / self.resOzon[-1][0], self.resWB[-1][4] / self.resWB[-1][0],
                 (self.resOzon[-1][4] + self.resWB[-1][4]) / (self.resOzon[-1][0] + self.resWB[-1][0])]

        self.tableWidget.setItem(11, 0, QTableWidgetItem(be_nums(self.resOzon[-1][5] / len(self.routsOzon))))
        self.tableWidget.setItem(11, 1, QTableWidgetItem(be_nums(self.resWB[-1][5] / len(self.routsWB))))
        self.tableWidget.setItem(11, 2, QTableWidgetItem(
            be_nums((self.resOzon[-1][5] + self.resWB[-1][5]) / (len(self.routsOzon) + len(self.routsWB)))))
        sb = [self.resOzon[-1][5] / len(self.routsOzon), self.resWB[-1][5] / len(self.routsWB),
              (self.resOzon[-1][5] + self.resWB[-1][5]) / (len(self.routsOzon) + len(self.routsWB))]

        self.tableWidget.setItem(12, 0, QTableWidgetItem(be_nums(piece[0] - sb[0])))
        self.tableWidget.setItem(12, 1, QTableWidgetItem(be_nums(piece[1] - sb[1])))
        self.tableWidget.setItem(12, 2, QTableWidgetItem(
            be_nums(piece[2] - sb[2])))

        self.tableWidget.setItem(13, 0, QTableWidgetItem(
            be_nums((piece[0] - sb[0]) / (self.resOzon[-1][1] / self.resOzon[-1][0]) * 100) + '%'))
        self.tableWidget.setItem(13, 1,
                                 QTableWidgetItem(
                                     be_nums((piece[1] - sb[1]) / (self.resWB[-1][1] / self.resWB[-1][0]) * 100) + '%'))
        self.tableWidget.setItem(13, 2, QTableWidgetItem(
            be_nums((piece[2] - sb[2]) / ((self.resOzon[-1][1] + self.resWB[-1][1]) / (
                    self.resOzon[-1][0] + self.resWB[-1][0])) * 100) + '%'))

        self.tableWidget.setItem(14, 0, QTableWidgetItem(be_nums((piece[0] - sb[0]) * self.resOzon[-1][0])))
        self.tableWidget.setItem(14, 1, QTableWidgetItem(be_nums((piece[1] - sb[1]) * self.resWB[-1][0])))
        self.tableWidget.setItem(14, 2, QTableWidgetItem(
            be_nums((piece[2] - sb[2]) * (self.resOzon[-1][0] + self.resWB[-1][0]))))

    def copy_clipboard(self):
        line_to_copy = ''
        for i in range(15):
            for j in range(3):
                line_to_copy += self.tableWidget.item(i, j).text()
                if j != 2:
                    line_to_copy += '\t'
            if i != 14:
                line_to_copy += '\n'
        QApplication.clipboard().setText(line_to_copy)
        self.label_2.setText('Текст скопирован!')

    def back(self):
        self.parent.move(self.x() + 230, self.y() + 150)
        self.parent.show()
        self.hide()
