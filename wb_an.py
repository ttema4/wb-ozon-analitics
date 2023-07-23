import xlwt
from PyQt5.Qt import *
from PyQt5 import QtWidgets, uic
from PyQt5.QtChart import QChart, QChartView, QPieSeries
from table_analit import analits_wb, get_sheet_name
from another_funcs import be_nums


class WbAnalits(QtWidgets.QMainWindow):
    def __init__(self, route, parent=None):
        super().__init__()
        self.parent = parent
        try:
            self.data = analits_wb(route)
        except:
            uic.loadUi('uis/error3.ui', self)
            self.pushButton.clicked.connect(self.back)

        else:
            self.sheet_name = get_sheet_name(route)
            self.name_file = ''.join(route.split('/')[-1].split('.')[:-1])
            self.upd_data = [0] * 4
            self.sr_doh = 0
            self.sr_pr = 0
            self.tov_min_rek = 0
            self.initUI()

    def initUI(self):
        uic.loadUi('uis/wb_an2.ui', self)
        self.pushButton_2.clicked.connect(self.download)
        self.pushButton_3.clicked.connect(self.back)
        self.lineEdit.textEdited.connect(self.upd)
        self.lineEdit_2.textEdited.connect(self.upd)
        self.lineEdit_3.textEdited.connect(self.upd)

        stats = ''
        stats += f"· Продажи: <b>{str(be_nums(self.data[-1][1]))}шт.</b><br>"
        stats += f"· Возвраты: <b>{be_nums(self.data[-1][3])}шт.</b><br>"
        stats += f"· Доплаты: <b>{be_nums(self.data[-1][7])}руб.</b><br>"
        self.label_6.setText(stats)

        series = QPieSeries()
        series.setHoleSize(0.40)
        all_purs = self.data[-1][1]

        for x in self.data[:-1]:
            series.append(str(x[0]) + ' ' + str(x[1]) + '(' + str(round(x[1] / all_purs * 100, 1)) + '%' + ')', x[1])

        chart = QChart()
        chart.addSeries(series)

        chart.setAnimationOptions(QChart.SeriesAnimations)
        chart.setTheme(QChart.ChartThemeLight)

        chartview = QChartView(chart)
        chartview.setRenderHint(QPainter.Antialiasing)
        grid = QtWidgets.QGridLayout(self.widget)
        grid.addWidget(chartview, 1, 1)

        series.hovered.connect(self.handle_hovered)

    def handle_hovered(self, slice, state):
        if state:
            slice.setExploded(True)
            slice.setLabelVisible(True)
        else:
            slice.setExploded(False)
            slice.setLabelVisible(False)

    def back(self):
        self.parent.move(self.x() + 230, self.y() + 150)
        self.parent.show()
        self.hide()

    def upd(self):
        self.pushButton_2.setEnabled(True)

        self.storage_sum = int(self.lineEdit.text()) if self.lineEdit.text() else 0
        self.extra_paids = int(self.lineEdit_2.text()) if self.lineEdit_2.text() else 0
        self.selfcount = int(self.lineEdit_3.text()) if self.lineEdit_3.text() else 0

        # data = id, кол-во продаж, кол-во дост, кол-во возвр, сумма со скидками, сумма к переводу, затраты на дост, доплаты, итого сумма (к пер - дост)

        stats = ''
        stats += f"· Хранение - {str(round(self.storage_sum / self.data[-1][4] * 100))}%<br>"
        stats += f"· Реклама - {str(round(self.extra_paids / self.data[-1][4] * 100))}%<br>"
        self.label_7.setText(stats)

        stats = ''
        stats += f"· Сумма с учётом скидки: <b>{be_nums(self.data[-1][4])}руб.</b><br>"
        stats += f"<i>&nbsp;&nbsp;&nbsp;&nbsp;– Цена за 1шт: </i><b>{be_nums(round(self.data[-1][4] / self.data[-1][1]))}руб.</b><br>"

        stats += f"· Сумма коммисий МП: <b>{be_nums(self.data[-1][4] - self.data[-1][5])}руб.</b> | {str(round((self.data[-1][4] - self.data[-1][5]) / self.data[-1][4] * 100))}%<br>"
        stats += f"<i>&nbsp;&nbsp;&nbsp;&nbsp;– Цена за 1шт: </i><b>{be_nums(round(self.data[-1][5] / self.data[-1][1]))}руб.</b><br>"

        stats += f"· Сумма затрат ПослМиля: <b>{be_nums(self.data[-1][6])}руб.</b> | {str(round(self.data[-1][6] / self.data[-1][4] * 100))}%<br>"
        stats += f"<i>&nbsp;&nbsp;&nbsp;&nbsp;– Цена за 1шт: </i><b>{be_nums(round((self.data[-1][5] - self.data[-1][6]) / self.data[-1][1]))}руб.</b><br>"

        self.result = self.data[-1][5] - self.data[-1][6] - self.storage_sum + self.data[-1][7]
        stats += f"· Итого без рекламы: <b>{be_nums(self.result)}руб.</b> | {str(round(self.result / self.data[-1][4] * 100))}%<br>"
        stats += f"<i>&nbsp;&nbsp;&nbsp;&nbsp;– Цена за 1шт: </i><b>{be_nums(round(self.result / self.data[-1][1]))}руб.</b><br>"
        stats += '<br>'

        stats += f"· Цена рекламы/1 товар: <b>{be_nums(round(self.extra_paids / self.data[-1][1]))}руб.</b><br>"
        stats += '<br>'

        one_piece = round((self.result - self.extra_paids) / self.data[-1][1])
        stats += f"· Итого к перечислению: <b>{be_nums(self.result - self.extra_paids)}руб.</b> | {str(round((self.result - self.extra_paids) / self.data[-1][4] * 100))}%<br>"
        stats += f"<i>&nbsp;&nbsp;&nbsp;&nbsp;– Цена за 1шт: </i><b>{be_nums(one_piece)}руб.</b><br>"
        stats += f"<i>&nbsp;&nbsp;&nbsp;&nbsp;– Себестоимость: </i><b>{be_nums(self.selfcount)}руб.</b><br>"
        stats += f"<i>&nbsp;&nbsp;&nbsp;&nbsp;– Доход за 1шт: </i><b>{be_nums(one_piece - self.selfcount)}руб.</b><br>"
        stats += '<br>'

        stats += f"<b>· Итоговый доход: {be_nums((one_piece - self.selfcount) * self.data[-1][1])}руб.</b> | {str(round(((one_piece - self.selfcount) * self.data[-1][1]) / self.data[-1][4] * 100))}%<br>"

        self.label_8.setText(stats)

    def download(self):
        book = xlwt.Workbook(encoding='utf-8')
        sheet1 = book.add_sheet('Sheet 1')

        headers = ['Продажи, шт.', 'Продажи на площадке, руб', 'К перечислению после вычета МП, руб.',
                   'Реклама МП, руб.',
                   'К перечислению после вычета рекламы, руб.', 'Себестоимость, 1 шт.']
        for i in range(len(headers)):
            sheet1.write(i, 0, headers[i], style=xlwt.easyxf('font: name Calibri, bold on'))
        # data = id, кол-во продаж, кол-во дост, кол-во возвр, сумма со скидками, сумма к переводу, затраты на дост, доплаты, итого сумма (к пер - дост)
        sheet1.write(0, 1, self.data[-1][1])
        sheet1.write(1, 1, self.data[-1][4])
        sheet1.write(2, 1, self.data[-1][5])
        sheet1.write(3, 1, self.extra_paids)
        sheet1.write(4, 1, self.result - self.extra_paids)
        sheet1.write(5, 1, self.selfcount)

        book.save('results_wb/' + self.name_file + '_wb_week.xls')
