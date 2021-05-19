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
        p = Path('C:/Users/your/directory/cleaned')
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
        wb = xlsxwriter.Workbook('wet_bulb.xlsx')
        sheet = wb.add_worksheet()
        print(len(self.wb))
        cnt = 0

        for i in range(1, len(self.columns)):
            sheet.write(0, i, self.columns[i])
            for j in range(0, len(self.rows)):
                if i == 1:
                    sheet.write(j+1, 0, self.rows[j])
                sheet.write(j+1, i, self.wb[cnt])
                cnt += 1

        wb.close()

    def percentile_90(self):
        wb = pd.read_excel('wet_bulb.xlsx') # read the we_bulb file from current working directory
        de_90 = np.percentile(wb['DE'], 90)
        es_90 = np.percentile(wb['ES'], 90)
        nl_90 = np.percentile(wb['NL'], 90)
        per = [de_90, es_90, nl_90]  # percentiles list
        print(de_90, es_90, nl_90)

        for i in range(1, 4):
            cnt = 0
            for j in range(0, len(wb.iloc[:, i])):
                # if percentile from day x and day before x is above percentile value add 1 to counter
                if wb.iloc[j, i] > per[i-1] and wb.iloc[j-1, i] > per[i-1]:
                    # print(wb.iloc[j, 0])
                    cnt += 1
            self.heat.append(cnt)
            print(cnt)




wet_bulb = Data()
#wet_bulb.reading()
#wet_bulb.calculation()
#wet_bulb.write_to_excel()
wet_bulb.percentile_90()

