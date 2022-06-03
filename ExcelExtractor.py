from itertools import dropwhile
from operator import index
from turtle import right
import pandas as pd
import numpy as np
import glob

# to do list
# 1. convert the daily excel to dataframe.
# 2. process the dataframe.
# 3. put the dataframe into the consolidated excel.

fact_data = pd.DataFrame()

f = r"C:\Users\Yusuf\Documents\My Project\Factory Work Plan\ExcelExtractor\sources\Production Line Arrangement of 2022.xlsx"

# process excel shifts of CCC4


def day_CCC4(filepath):
    max_row = 0
    max_col = 0

    delimiter = []

    xl = pd.ExcelFile(filepath)

    # night shift
    df = xl.parse(sheet_name=0, usecols="A:R")

    #print(df)

    #print(df.dropna(how='all', axis=1))

    df.columns = ['1', '2', '3', '4', '5', '6', '7', '8',
                  '9', '10', '11', '12', '13', '14', '15', '16', '17', '18']


    max_row = df.shape[0]
    max_col = df.shape[1]

    #print(max_col, " ", max_row)

    df2 = df.iloc[max_row-39:max_row-12, max_col-8:max_col]

    df3 = df2.dropna(how='all', axis=1)
    df4 = df3.dropna(how='all')
    df5 = df4.reset_index(drop=True)

    #print(df2)

    delimiter = df5[df5['11'] == 'Total HC:'].index.values

    fir_table, sec_table = df5.iloc[:delimiter[0]+1], df5.iloc[delimiter[0]+1:]

    return fir_table.reset_index(drop=True), sec_table.reset_index(drop=True)


def night_CCC4(filepath):
    max_row = 0
    max_col = 0

    xl = pd.ExcelFile(filepath)

    # night shift
    df = xl.parse(2)

    df.columns = ['1', '2', '3', '4', '5', '6', '7', '8',
                  '9', '10', '11', '12', '13', '14', '15', '16', '17', '18', '19']


    max_row = df.shape[0]
    max_col = df.shape[1]

    df2 = df.iloc[max_row-31:max_row-10, max_col-8:max_col]
    df3 = df2.dropna(how='all', axis=1)
    df4 = df3.dropna(how='all')
    df5 = df4.reset_index(drop=True)
    #print(df5)

    delimiter = df5[df5['12'] == 'Total HC:'].index.values

    fir_table, sec_table = df5.iloc[:delimiter[0]+1], df5.iloc[delimiter[0]+1:]

    return print(fir_table.reset_index(drop=True), sec_table.reset_index(drop=True))

# ----------------------------------------------
# process excel shifts of CCC2


def day_CCC2(filepath):
    max_row = 0
    max_col = 0

    xl = pd.ExcelFile(filepath)

    df = xl.parse(sheet_name=0, usecols="A:R")

    df.columns = ['1', '2', '3', '4', '5', '6', '7', '8',
                  '9', '10', '11', '12', '13', '14', '15', '16', '17', '18']

    max_row = df.shape[0]
    max_col = df.shape[1]

    df2 = df.iloc[max_row-39:max_row-12, max_col-18:max_col-8]
    df3 = df2.dropna(how='all', axis=1)
    df4 = df3.dropna(how='all')
    df5 = df4.reset_index(drop=True)

    # print(df5)

    delimiter = df5[df5['2'] == 'Total HC:'].index.values

    fir_table, sec_table = df5.iloc[:delimiter[0]+1], df5.iloc[delimiter[0]+1:]

    return fir_table.reset_index(drop=True), sec_table.reset_index(drop=True)


def night_CCC2(filepath):
    max_row = 0
    max_col = 0

    xl = pd.ExcelFile(filepath)
    df = xl.parse(2)

    df.columns = ['1', '2', '3', '4', '5', '6', '7', '8',
                  '9', '10', '11', '12', '13', '14', '15', '16', '17', '18', '19']

    max_row = df.shape[0]
    max_col = df.shape[1]

    df2 = df.iloc[max_row-31:max_row-10, max_col-18:max_col-9]
    df3 = df2.dropna(how='all', axis=1)
    df4 = df3.dropna(how='all')
    df5 = df4.reset_index(drop=True)

    # print(df5)

    delimiter = df5[df5['3'] == 'Total HC:'].index.values

    fir_table, sec_table = df5.iloc[:delimiter[0]+1], df5.iloc[delimiter[0]+1:]

    return fir_table.reset_index(drop=True), sec_table.reset_index(drop=True)

# ----------------------------------------------
# process dataframe shifts of CCC4


def day_CCC4Df(df):
    # On duty time
    global fName

    rightDf = df


    if not rightDf[rightDf['11'].str.contains("Next Day Shift")].empty:

        rightDfclean = rightDf.dropna(how='all', axis=1)

        rightDfclean.columns = ['Line', 'Start Time', 'End Time', 'UPH']

        fNameDf = rightDfclean['Line'].str.extract(r'(CCC[2-4])')
        dateDf = rightDfclean['Line'].str.extract(
            r'([A-Z][a-z][a-z][.,-][0-3][0-9])')
        isNight = rightDfclean[rightDfclean['Line'].str.contains("Day")].empty

        fName = fNameDf.loc[0].iat[0]
        date = dateDf.loc[0].iat[0]
        shift = ''

        shift = 'start'

        #df1 = rightDf.drop(columns=['4'])

        df2 = rightDfclean.drop([0, 1, (rightDfclean.shape[0])-1])

    # Means that the DF is on off duty time time or end shift.
    # process df

    elif not rightDf[rightDf['11'].str.contains("Today")].empty:

        rightDfclean = rightDf.dropna(how='all', axis=1)

        rightDfclean.columns = ['Line', 'OT', 'HC', 'End shift']

        fNameDf = rightDfclean['Line'].str.extract(r'(CCC[2-4])')
        dateDf = rightDfclean['Line'].str.extract(
            r'([A-Z][a-z][a-z][.,-][0-3][0-9])')
        isNight = rightDfclean[rightDfclean['Line'].str.contains("Day")].empty

        fName = fNameDf.loc[0].iat[0]
        date = dateDf.loc[0].iat[0]
        shift = ''

        shift = 'end'

        df2 = rightDfclean.drop([0, 1, (rightDfclean.shape[0])-1])

    return df2.reset_index(drop=True), fName, date, shift, isNight


def night_CCC4Df(df):
    # On duty time
    global fName

    rightDf = df

    print(rightDf)


    if not rightDf[rightDf['12'].str.contains("Next Night-Shift")].empty:

        rightDfclean = rightDf.dropna(how='all', axis=1)

        rightDfclean.columns = ['Line', 'Start Time', 'End Time', '4', 'UPH']

        fNameDf = rightDfclean['Line'].str.extract(r'(CCC[2-4])')
        dateDf = rightDfclean['Line'].str.extract(
            r'([A-Z][a-z][a-z][.,-][0-3][0-9])')
        isNight = rightDfclean[rightDfclean['Line'].str.contains("Day")].empty

        fName = fNameDf.loc[0].iat[0]
        date = dateDf.loc[0].iat[0]
        shift = ''

        shift = 'start'

        df1 = rightDfclean.drop(columns=['4'])

        df2 = df1.drop([0, 1, (df1.shape[0])-1])

    # Means that the DF is on off duty time time or end shift.
    # process df

    elif not rightDf[rightDf['12'].str.contains("Today")].empty:

        rightDfclean = rightDf.dropna(how='all', axis=1)

        rightDfclean.columns = ['Line', 'OT', 'HC', 'End shift']

        fNameDf = rightDfclean['Line'].str.extract(r'(CCC[2-4])')
        dateDf = rightDfclean['Line'].str.extract(
            r'([A-Z][a-z][a-z][.,-][0-3][0-9])')
        isNight = rightDfclean[rightDfclean['Line'].str.contains("Day")].empty

        fName = fNameDf.loc[0].iat[0]
        date = dateDf.loc[0].iat[0]
        shift = ''

        shift = 'end'

        df2 = rightDfclean.drop([0, 1, (rightDfclean.shape[0])-1])

    return df2.reset_index(drop=True), fName, date, shift, isNight

# ----------------------------------------------
# process dataframe shifts of CCC2


def day_CCC2Df(df):
    # On duty time
    fName = ''

    #This is CCC2
    leftDf = df

    if not leftDf[leftDf['2'].str.contains("Next Day Shift")].empty:

        leftDfclean = leftDf.dropna(how='all', axis=1)

        leftDfclean.columns = ['Line', 'Start Time', 'End Time', 'UPH']

        fNameDf = leftDfclean['Line'].str.extract(r'(CCC[2-4])')
        dateDf = leftDfclean['Line'].str.extract(
            r'([A-Z][a-z][a-z][.,-][0-3][0-9])')
        isNight = leftDfclean[leftDfclean['Line'].str.contains("Day")].empty

        fName = fNameDf.loc[0].iat[0]
        date = dateDf.loc[0].iat[0]
        shift = ''

        shift = 'start'

        df2 = leftDfclean.drop([0, 1, (leftDfclean.shape[0])-1])

    elif not leftDf[leftDf['2'].str.contains("Today")].empty:

        leftDfclean = leftDf.dropna(how='all', axis=1)

        leftDfclean.columns = ['Line', 'OT', 'HC', 'End shift']

        fNameDf = leftDfclean['Line'].str.extract(r'(CCC[2-4])')
        dateDf = leftDfclean['Line'].str.extract(
            r'([A-Z][a-z][a-z][.,-][0-3][0-9])')
        isNight = leftDfclean[leftDfclean['Line'].str.contains("Day")].empty

        fName = fNameDf.loc[0].iat[0]
        date = dateDf.loc[0].iat[0]
        shift = ''

        shift = 'end'

        df2 = leftDfclean.drop([0, 1, (leftDfclean.shape[0])-1])

    return df2.reset_index(drop=True), fName, date, shift, isNight


def night_CCC2Df(df):
    # On duty time
    fName = ''

    # This is Night CCC2
    leftDf = df

    if not leftDf[leftDf['3'].str.contains("Next Night-Shift")].empty:

        leftDfclean = leftDf.dropna(how='all', axis=1)

        leftDfclean.columns = ['Line', 'Start Time', 'End Time', '4', 'UPH']

        fNameDf = leftDfclean['Line'].str.extract(r'(CCC[2-4])')
        dateDf = leftDfclean['Line'].str.extract(
            r'([A-Z][a-z][a-z][.,-][0-3][0-9])')
        isNight = leftDfclean[leftDfclean['Line'].str.contains("Day")].empty

        fName = fNameDf.loc[0].iat[0]
        date = dateDf.loc[0].iat[0]
        shift = ''

        shift = 'start'

        df1 = leftDfclean.drop(columns=['4'])

        df2 = df1.drop([0, 1, (df1.shape[0])-1])

    elif not leftDf[leftDf['3'].str.contains("Today")].empty:

        leftDf1 = leftDf

        if not leftDf[leftDf['3'].str.contains("K8")].empty:
            leftDf1 = leftDf.drop(columns=['6'])

        leftDfclean = leftDf1.dropna(how='all', axis=1)

        leftDfclean.columns = ['Line', 'OT', 'HC', 'End shift']

        fNameDf = leftDfclean['Line'].str.extract(r'(CCC[2-4])')
        dateDf = leftDfclean['Line'].str.extract(
            r'([A-Z][a-z][a-z][.,-][0-3][0-9])')
        isNight = leftDfclean[leftDfclean['Line'].str.contains("Day")].empty

        fName = fNameDf.loc[0].iat[0]
        date = dateDf.loc[0].iat[0]
        shift = ''

        shift = 'end'

        df2 = leftDfclean.drop([0, 1, (leftDfclean.shape[0])-1])

    else:
        print("Try Again")

    return df2.reset_index(drop=True), fName, date, shift, isNight


# day_CCC4(f)
# removed duplicate
# df6 = CCC4Df()['Line'].drop_duplicates()
# df7 = df6.dropna(how='all')

# s = df7.sort_index()
# df8 = s.to_frame().T

# print(filterDf())
# print(night_CCC2(f))


# print(night_CCC2(f)[1])
# night_CCC2Df(night_CCC2(f)[1])
# print(CCC2Df())
