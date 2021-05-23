#! python3
import xlsxwriter
from pathlib import Path
import pandas as pd
import numpy as np


class Data:
    def __init__(self):
        self.temp_av = []
        self.hum = []
        self.qq = []
        self.wind = []
        self.rows = []
        self.columns = []
        self.wb = []
        self.heat = []

    def reading(self):
        # reading excel file
        # directory below should contains only cleaned xlsx files
        p = Path('C:/Users/kinan/Desktop/Bioinformatyka/Modern Data Analytics/cleaned')
        files = list(p.glob('*.xlsx'))
        files.sort  # sort in alphabetical order
        self.hum = pd.read_excel(files[0]).drop(['FR'], axis=1)
        self.qq = pd.read_excel(files[1])
        self.temp_av = pd.read_excel(files[2]).drop(['FR'], axis=1)
        self.wind = pd.read_excel(files[3]).drop(['FR'], axis=1)
        self.rows = self.qq.iloc[:, 0]
        self.columns = list(self.qq.columns)

    def calculation(self):
        # formula: 0.735×Ta+0.0374×RH+0.00292×Ta×RH +7.619×SR−4.557×SR2−0.0572×WS−4.064
        for i in range(1, len(self.columns)):
            for j in range(0, len(self.rows)):
                sr = self.qq.iloc[j, i] * 0.001
                ws = self.wind.iloc[j, i] * 0.1
                ta = self.temp_av.iloc[j, i]
                rh = self.wind.iloc[j, i]
                self.wb.append(0.735*ta + 0.0374*rh + 0.00292*ta*rh + 7.619*sr - 4.557*(sr**2) - 0.0572*ws - 4.064)

    def write_to_excel(self):
        wb = xlsxwriter.Workbook('wet_bulb2.xlsx')
        sheet = wb.add_worksheet()
        cnt = 0

        for i in range(1, len(self.columns)):
            cnt_row = 0
            sheet.write(0, i, self.columns[i])
            for j in range(0, len(self.rows)):
                if 3 < int(self.rows[j][5:7]) < 10:
                    print(j+1, i, self.wb[cnt])
                    sheet.write(cnt_row+1, i, self.wb[cnt])
                    if i == 1:
                        sheet.write(cnt_row+1, 0, self.rows[j])
                    cnt_row += 1
                cnt += 1
        wb.close()

    def percentile_90(self):
        wb = pd.read_excel('wet_bulb_quarters.xlsx') # read the we_bulb file from current working directory

        de_90 = np.percentile(wb['DE'], 90)
        es_90 = np.percentile(wb['ES'], 90)
        nl_90 = np.percentile(wb['NL'], 90)
        per = [de_90, es_90, nl_90]  # percentiles list
        print(de_90, es_90, nl_90)

        for i in range(1, 4):
            cnt_h = 0  # number of heatwaves
            in_row = False
            to_break = 0  # counter for the days after heatwave
            start = []  # date of first day of heatwave
            stop = []  # dat of last day of heatwave
            first = True  # first day of heatwave

            for j in range (0, len(wb.iloc[:, i])):
                if wb.iloc[j, i] > per[i-1] and wb.iloc[j-1, i] > per[i-1]:
                    if first:
                        start.append(wb.iloc[j, 0])
                        first = False
                    in_row = True  # if True we are within heatwave
                    to_break = 0  # in heatwave, number after heatwave = 0
                else:
                    # counting days with temp below percentile value
                    if in_row:
                        to_break += 1

                    # if we have more then 7 days after high temperatures it is the end of heatwave
                    if to_break > 7:
                        to_break = 0
                        in_row = False  # end of heatwave
                        cnt_h += 1
                        first = True
                        stop.append(wb.iloc[j, 0])
                        #print("HEAT:"+ wb.iloc[j, 0])

            print(cnt_h)
            print(start)
            print(stop)


wet_bulb = Data()
#wet_bulb.reading()
#wet_bulb.calculation()
#wet_bulb.write_to_excel()
wet_bulb.percentile_90()

