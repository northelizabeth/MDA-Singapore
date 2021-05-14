#! python3
import re
import pandas as pd
import xlsxwriter
from sklearn.impute import SimpleImputer
from pathlib import Path


class Info:
    def __init__(self):
        self.final = ['DE', 'ES', 'NL', 'FR']
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

        self.columns = list(df.columns.str.replace(' .*', ''))
        self.rows = list(df.index)
        df = pd.DataFrame(imp.transform(df)).T
        df.insert(0, 'ISO', self.columns)
        self.columns = list(df.groupby('ISO').groups.keys())
        self.df = df.groupby('ISO').mean()

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


pattern = 'temp'
pattern_regex = re.compile(pattern)
info = Info()
p = Path('C:/Users/kinan/Desktop/Bioinformatyka/Modern Data Analytics/raw/temp')
for textFilePathObj in p.glob('*.xlsx'):
    file = pd.read_excel(textFilePathObj)
    data = info.cleaning(file)
    result = pattern_regex.search(str(textFilePathObj))
    if result != None:
        info.temp()
    info.name = (str(textFilePathObj)[:-5]+'_cl.xlsx')
    info.write_to_excel()


