#! python3
import re
import pandas as pd
import xlsxwriter
from sklearn.impute import SimpleImputer
from pathlib import Path


class Info:
    def __init__(self):
        self.final = ['DE', 'ES', 'NL']
        self.name = ''
        self.df = []
        self.rows = []
        self.columns = []

    def temp(self):
        # this function change temperature value from Fahrenheit to Celsius
        for i in range(0, len(self.df.columns)):
            for j in range(0, len(self.df[i])):
                self.df[i][j] = (self.df[i][j])*0.1

    def cleaning(self, df):
        # this function replace or remove NA and return pandas data frame
        df = df.iloc[:, :-1]  # removing last column(it's empty) and first row (header)
        df_final = []
        # deleting every row where more then 365 records are missing
        for i in range(0, len(df)):
            if df.loc[[i]].isna().sum().sum() > 365:
                df = df.drop([i])
        for i in range(0, len(df)):
            if str(df.iloc[i, 0][:2]) in self.final:
                df_final.append(df.iloc[i])
        df = pd.DataFrame(df_final).T  # transpose
        df.columns = df.iloc[0, ]  # first row will be header
        df = df.iloc[1:, :]  # remove first row

        imp = SimpleImputer(strategy="mean")  # replacing NA by mean from the entire column
        imp.fit(df)

        self.columns = list(df.columns)
        self.rows = list(df.index)
        df = pd.DataFrame(imp.transform(df)).T
        self.df = df

    def write_to_excel(self):
        wb = xlsxwriter.Workbook(self.name)
        sheet = wb.add_worksheet()
        print(self.columns)
        print(self.df)

        for i in range(0, len(self.df.columns)):
            sheet.write(i+1, 0, self.rows[i])
            if i < len(self.columns):
                sheet.write(0, i+1, self.columns[i])
            for j in range(0, len(self.df[i])):
                sheet.write(i+1, j + 1, self.df[i][j])  # writing variables'''
        wb.close()


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
        self.final = []

    def reading(self):
        # reading excel file
        # directory below should contains only cleaned xlsx files
        p = Path('C:/Users/kinan/Desktop/Bioinformatyka/Modern Data Analytics/raw/map')
        files = list(p.glob('*.xlsx'))
        files.sort  # sort in alphabetical order
        self.hum = pd.read_excel(files[0]).sort_index(axis=1)
        self.qq = pd.read_excel(files[1]).sort_index(axis=1)
        self.temp_av = pd.read_excel(files[2]).sort_index(axis=1)
        self.wind = pd.read_excel(files[3]).sort_index(axis=1)

        # Converting the arrays into sets
        s1 = set(self.hum.columns)
        s2 = set(self.qq.columns)
        s3 = set(self.temp_av.columns)
        s4 = set(self.wind.columns)
        # Compering sets to find mutual stations for every file
        set1_2 = s1.intersection(s2)
        set3_4 = s3.intersection(s4)
        self.final = set1_2.intersection(set3_4)

        self.hum = self.hum.drop(columns=[col for col in self.hum if col not in self.final])
        self.qq = self.qq.drop(columns=[col for col in self.qq if col not in self.final])
        self.temp_av = self.temp_av.drop(columns=[col for col in self.temp_av if col not in self.final])
        self.wind = self.wind.drop(columns=[col for col in self.wind if col not in self.final])
        self.rows = self.qq.iloc[:, -1]
        print(self.rows)
        self.columns = list(self.qq.columns)

    def calculation(self):
        # formula: 0.735×Ta+0.0374×RH+0.00292×Ta×RH +7.619×SR−4.557×SR2−0.0572×WS−4.064
        for i in range(1, len(self.columns)-1):
            for j in range(0, len(self.rows)):

                sr = self.qq.iloc[j, i] * 0.001
                ws = self.wind.iloc[j, i] * 0.1
                ta = self.temp_av.iloc[j, i]
                rh = self.wind.iloc[j, i]
                self.wb.append(0.735 * ta + 0.0374 * rh + 0.00292 * ta * rh + 7.619 * sr - 4.557 * (
                            sr ** 2) - 0.0572 * ws - 4.064)

    def write_to_excel(self):
        wb = xlsxwriter.Workbook('wet_bulb_map.xlsx')
        sheet = wb.add_worksheet()
        cnt = 0

        for i in range(1, len(self.columns)-1):
            cnt_row = 0
            sheet.write(0, i, self.columns[i])
            for j in range(0, len(self.rows)):
                if 3 < int(self.rows[j][5:7]) < 10:
                    #print(j + 1, i, self.wb[cnt])
                    sheet.write(cnt_row + 1, i, self.wb[cnt])
                    if i == 1:
                        sheet.write(cnt_row + 1, 0, self.rows[j])
                    cnt_row += 1
                cnt += 1
        wb.close()

        # writing down also list of stations
        final = list(self.final)
        wb = xlsxwriter.Workbook('stations.xlsx')
        sheet = wb.add_worksheet()
        for i in range(0, len(final)):
            sheet.write(i, 0, final[i])
        wb.close()


class Map:
    def __init__(self):
        self.file = 'C:/Users/kinan/Desktop/Bioinformatyka/Modern Data Analytics/qq/sources.txt'
        self.names = []
        self.numbers = []
        self.rows = []
        self.columns = []
        self.df = []

        # reading stattion names from previously prepared stations.xlsx file
        names = pd.read_excel('stations.xlsx', header=None)
        names = names[0].to_list()  # convert to list
        names.remove('Unnamed: 0')

        for i in range(0, len(names)):
            i = names[i].split(' ')
            self.names.append(i[1])
            self.numbers.append(i[2])

    def df_prepare(self, text):
        """
          Converting text from .txt file into data frame.
          Step 1: transforming spaces into tabs
          Step 2: Creating data frame (columns are split by ',')
          """
        # this part is meant to transform block of spaces in .txt file into tabs
        pattern = '\\s*\\s'
        patter_regex = re.compile(pattern)
        for i in range(0, len(text)):
            text[i] = patter_regex.sub('', str(text[i]))

        # this part is meant to transform string text into data frame
        df = pd.DataFrame(text)
        df = df[0].str.split(',', expand=True)
        header = df.iloc[0, ]  # first row will be the header
        df.columns = header
        df = df.iloc[2:, :-1]  # last column and first row are unnecessary

        return df

    def long_lati(self):
        with open(Path(self.file), encoding='utf8') as file:
            text = file.readlines()[23:]  # first 23 lines contains unnecessary description
        df = self.df_prepare(text)
        df_station = pd.DataFrame(columns=['Name', 'Longitude', 'Latitude'])
        for i, row in df.iterrows():
            if row['STAID'] in self.numbers and row['SOUNAME'] in self.names and row['ELEI'] == 'QQ6':
                lat = row['LAT']
                lon = row['LON']
                info = row['CN'] + ' ' + row['SOUNAME'] + ' ' + row['STAID']
                df_station = df_station.append({'Name': info, 'Longitude': lon, 'Latitude': lat}, ignore_index=True)
        df_station = df_station.set_index('Name')
        df_station = df_station.sort_index(axis=0)

        self.df = df_station
        self.rows = list(df_station.index)
        self.columns = list(df_station.columns)

    def write_to_excel(self):
        wb = xlsxwriter.Workbook('location.xlsx')
        sheet = wb.add_worksheet()

        for i in range(0, len(self.df.columns)):
            sheet.write(0, i+1, self.columns[i])
             #   sheet.write(0, i+1, self.columns[i])
            for j in range(0, len(self.rows)):
                if i == 0:
                    sheet.write(j+1, i, self.rows[j])
                sheet.write(j+1, i+1, self.df.iloc[j, i])  # writing variables'''
        wb.close()



def prepocessing():
    pattern = 'temp'
    pattern_regex = re.compile(pattern)
    info = Info()
    p = Path('C:/Users/kinan/Desktop/Bioinformatyka/Modern Data Analytics/raw')
    for textFilePathObj in p.glob('*.xlsx'):
        file = pd.read_excel(textFilePathObj)
        data = info.cleaning(file)
        result = pattern_regex.search(str(textFilePathObj))
        if result != None:
            info.temp()
        info.name = (str(textFilePathObj)[:-5]+'_cl.xlsx')
        info.write_to_excel()
#prepocessing()
#wet_bulb = Data()
#wet_bulb.reading()
#wet_bulb.calculation()
#wet_bulb.write_to_excel()
map = Map()
map.long_lati()
map.write_to_excel()


