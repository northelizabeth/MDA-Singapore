#! python3
import xlsxwriter
import re
import os
import pandas as pd
import datetime
from pathlib import Path
from datetime import timedelta

"""
This script only extracts data. It does not clean them. It works only for blended data, but before running the script you
have to prepare source.txt file. From source.txt file you have to remove additional ',' in:
CADIZ 
BALLYHAISE CAVAN (x2)
PLYMOUTH MOUNT BATTEN
NOTTINGHAM WATNALL
OAK PARK CARLOW 
"""


class Data:

    def __init__(self):
        self.cnt = 0  # this is counter for how many files has been already proceed, just to see the progress
        self.files = 'C:/Users/kinan/Desktop/Bioinformatyka/Modern Data Analytics/temp_min'  # path to folder with files INSERT YOUR OWN ONE HERE
        self.excel = 'C:/Users/kinan/Desktop/temp_min.xlsx'  #  path to directory where xlsx fil will be created # INSERT YOUR OWN ONE HERE WITH .xlxs EXTENTION
        self.pre = ''
        self.date = []  # contains all dates
        self.stations = []  # contains stations' info (ISO code, localization, id)
        self.all_data = []  # contains data from every file

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

    def source_file(self):
        """
        Creating lists with STATID and localization info
        """

        with open(Path(self.files)/'sources.txt', encoding='utf8') as file:
            text = file.readlines()[23:]  # first 23 lines contains unnecessary description
        df = self.df_prepare(text)

        # creating lists that contain stations id and where are they located
        stations = []
        country = []
        for i, row in df.iterrows():
            # print(i, 'STAID:', row['STAID'])
            if (int(row['START']) <= 19961231 and int(row['STOP']) > 20191231) and (row['STAID']) not in stations:
                stations.append(row['STAID'])
                country.append(str(row['CN']) + ' ' + str(row['SOUNAME']) + ' ' + str(row['STAID']))
        self.stations = country

        return stations, country

    def create_file(self):
        """
        Function to create Excel file and fill first row with dates
        from 1996-01-01 to 2019-12-30
        """
        end_date = datetime.date(2020, 1, 1)
        begin_date = datetime.date(1996, 1, 1)
        date = [begin_date]
        i = 0
        # calculating dates day by day from 1996-01-01 to 2020-01-01
        while date[i] < end_date:
            i += 1
            date.append(date[0]+timedelta(days=i))
        self.date = date

    def file_to_excel(self, file):
        """
        Function to open file, pull out data, save data as list
        """

        with open(Path(self.files)/file, encoding="utf8") as file:
            text = file.readlines()[20:]
        df = self.df_prepare(text)

        # if for some reason file is not matching the description in source.txt, because the measurement
        # starts after 1996-01-01 it is skipped and function returns None value
        if int(df['DATE'].iloc[0]) > 19960101:
            print('DELETING: ', self.stations[self.cnt]) # deleting station from list
            del self.stations[self.cnt]
            return False

        self.cnt += 1
        print(self.cnt)
        data = []
        for i, row in df.iterrows():
            if 19960101 <= int(row['DATE']) <= 20191231:  # choosing only records within particular time period
                if row[str(self.pre)] == '-9999':
                    data.append('NA')  # transforming -9999 into NA
                else:
                    data.append(int(row[self.pre]))
        self.all_data.append(data)

    def write_to_excel(self, countries):
        wb = xlsxwriter.Workbook(self.excel)
        sheet = wb.add_worksheet()
        print(len(self.all_data))
        print(len(countries))
        for i in range(0, len(self.date)):
            sheet.write(0, i + 1, str(self.date[i]))  # writing date in first row
        cnt = 1
        for i in range(0, len(self.all_data)):
            sheet.write(cnt, 0, countries[cnt-1])  # writing station's info in first column
            for j in range(0, len(self.all_data[i])):
                sheet.write(cnt, j + 1, self.all_data[i][j])  # writing variables
            cnt += 1
        wb.close()


info = Data()
files = os.listdir(info.files)  # list of files in directory
pre = files[0] # zip file name
info.pre = str(pre[-6:-4].upper())  # last two letters of zip file are prefix to every txt file
pre = str(pre[-6:-4].upper())+'_STAID'
print(info.pre)
info.create_file()
names, place = info.source_file()
for i in range(0, len(names)):
    names[i] = pre + str(0)*(6-len(names[i])) + str(names[i]) + '.txt'

for filenames in files:
    if filenames in names:
        print(filenames)
        info.file_to_excel(filenames)

info.write_to_excel(place)
