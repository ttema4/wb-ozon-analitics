import xlwt
from PyQt5.Qt import *
from PyQt5 import QtWidgets, uic
from PyQt5.QtChart import QChart, QChartView, QPieSeries
from table_analit import analits_ozon, get_sheet_name
from another_funcs import be_nums


class OzonAnalits(QtWidgets.QMainWindow):
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
            self.upd_data = [0] * 4
            self.sr_doh = 0
            self.sr_pr = 0
            self.tov_min_rek = 0
            self.initUI()

    def initUI(self):
        uic.loadUi('uis/wb_an2.ui', self)
        self.pushButton.clicked.connect(self.upd)
        self.pushButton_2.clicked.connect(self.download)
        self.pushButton_3.clicked.connect(self.back)
        points = ['Всего товаров: ', 'Всего продаж: ', 'Всего доставок: ', 'Всего возвратов: ',
                  'Сумма с учетом скидки: ', 'Сумма к перечислению: ', 'Затраты на доставку: ', 'Сумма доплат: ',
                  'Итоговая сумма: ']
        stats = ''
        for i in range(len(points)):
            stats += f"<b>·</b> {points[i]} <b>{be_nums(self.data[-1][i])}</b><br>"
        stats += '<br>'

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
        self.upd_data[0] = self.lineEdit.text() if self.lineEdit.text() != '' else 0
        self.upd_data[3] = self.lineEdit_2.text() if self.lineEdit_2.text() != '' else 0
        self.upd_data[2] = self.lineEdit_3.text() if self.lineEdit_3.text() != '' else 0
        self.upd_data[1] = self.lineEdit_4.text() if self.lineEdit_3.text() != '' else 0

        points = ['Всего товаров: ', 'Всего продаж: ', 'Всего доставок: ', 'Всего возвратов: ',
                  'Сумма с учетом скидки: ', 'Сумма к перечислению: ', 'Затраты на доставку: ', 'Сумма доплат: ',
                  'Итоговая сумма: ']
        stats = ''
        for i in range(len(points) - 1):
            stats += f"<b>·</b> {points[i]} <b>{be_nums(self.data[-1][i])}</b><br>"
        stats += '<br>'
        stats += f"<b>·</b> {'Итог за вычетом доп. расх.: '} <b>{be_nums(self.data[-1][-1] - int(self.upd_data[0]) - int(self.upd_data[3]))}</b><br>"
        stats += f"<b>·</b> {'Средняя цена товара: '} <b>{be_nums(self.data[-1][4] / self.data[-1][1])}</b><br>"
        stats += f"<b>·</b> {'Средний доход: '} <b>{be_nums((self.data[-1][-1] - int(self.upd_data[0]) - int(self.upd_data[3])) / self.data[-1][1])}</b><br>"
        self.sr_doh = round((self.data[-1][-1] - int(self.upd_data[0]) - int(self.upd_data[3])) / self.data[-1][1], 1)
        stats += f"<b>·</b> {'Средняя прибыль: '} <b>{be_nums(self.sr_doh - int(self.upd_data[1]))}</b><br>"
        self.sr_pr = round(self.sr_doh - int(self.upd_data[1]), 2)
        stats += f"<b>· {'Доходность продукта: '} {str(round(self.sr_pr / (self.data[-1][4] / self.data[-1][1]), 2) * 100)}%</b><br>"
        stats += '<br>'
        stats += f"<b>·</b> {'Цена рекламы/1 товар: '} <b>{be_nums(int(self.upd_data[2]) / self.data[-1][1])}</b><br>"
        self.tov_min_rek = self.sr_pr - (int(self.upd_data[2]) / self.data[-1][1])
        stats += f"<b>· {'Доходность чистая: '} {str(round(self.tov_min_rek / (self.data[-1][4] / self.data[-1][1]), 2) * 100)}%</b><br>"
        stats += '<br>'
        stats += f"<b>· {'Итог недели чистыми: '} {be_nums(self.tov_min_rek * self.data[-1][1])} руб.< /b > < br >"

        self.label_6.setText(stats)

    def download(self):
        route = QFileDialog.getExistingDirectory(self, 'Место сохранения', '/')
        book = xlwt.Workbook(encoding='utf-8')
        sheet1 = book.add_sheet('Sheet 1')
        headers = ['id товара', 'Всего продаж, шт.', 'Всего доставок, шт.', 'Всего возвратов, шт.',
                   'Общ. стоимость с учетом скидки, руб.', 'Сумма к перечислению, руб.', 'Затраты на доставку, руб.',
                   'Сумма доплат, руб.']
        for i in range(len(headers)):
            sheet1.write(0, i, headers[i], style=xlwt.easyxf('font: name Calibri, bold on'))
        for i in range(len(self.data)):
            for j in range(len(headers)):
                if i == len(self.data) - 1 and j == 0:
                    sheet1.write(i + 1, j, 'Итог:', style=xlwt.easyxf(
                        'font: name Calibri, bold on; pattern: pattern solid, fore_colour light_orange;'))
                elif j == 0 or i == len(self.data) - 1:
                    sheet1.write(i + 1, j, self.data[i][j], style=xlwt.easyxf(
                        'font: name Calibri, bold on'))
                elif i % 2 == 1:
                    sheet1.write(i + 1, j, self.data[i][j],
                                 style=xlwt.easyxf(
                                     'font: name Calibri; pattern: pattern solid, fore_colour gray25;'))
                else:
                    sheet1.write(i + 1, j, self.data[i][j], style=xlwt.easyxf('font: name Calibri'))
        headers = ['Итог за вычетом доп. расх, руб.', 'Средняя цена товара, руб.', 'Средний доход, руб.',
                   'Средняя прибыль, руб.',
                   'Доходность продукта, %', 'Цена рекламы/1 товар, руб.', 'Доходность чистая, %',
                   'Итог недели чистыми, руб.']
        for j in range(9, 9 + len(headers)):
            sheet1.write(0, j, headers[j - 9], style=xlwt.easyxf('font: name Calibri, bold on'))

        sheet1.write(1, 9, be_nums(self.data[-1][-1] - int(self.upd_data[0]) - int(self.upd_data[3])))
        sheet1.write(1, 10, be_nums(self.data[-1][4] / self.data[-1][1]))
        sheet1.write(1, 11,
                     be_nums((self.data[-1][-1] - int(self.upd_data[0]) - int(self.upd_data[3])) / self.data[-1][1]))
        sheet1.write(1, 12, be_nums(self.sr_doh - int(self.upd_data[1])))
        sheet1.write(1, 13, str(round(self.sr_pr / (self.data[-1][4] / self.data[-1][1]), 2) * 100), style=xlwt.easyxf(
            'font: name Calibri, bold on; pattern: pattern solid, fore_colour yellow;'))
        sheet1.write(1, 14, be_nums(int(self.upd_data[2]) / self.data[-1][1]))
        sheet1.write(1, 15, str(round(self.tov_min_rek / (self.data[-1][4] / self.data[-1][1]), 2) * 100),
                     style=xlwt.easyxf(
                         'font: name Calibri, bold on; pattern: pattern solid, fore_colour yellow;'))
        sheet1.write(1, 16, be_nums(self.tov_min_rek * self.data[-1][1]), style=xlwt.easyxf(
            'font: name Calibri, bold on; pattern: pattern solid, fore_colour light_green;'))

        book.save(route + '/' + self.sheet_name + '_week_wb.xls')
