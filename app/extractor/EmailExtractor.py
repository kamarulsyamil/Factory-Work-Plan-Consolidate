from cmath import nan
import os
from re import L
import re
from matplotlib.pyplot import axis
import numpy as np
from numpy import AxisError, NaN, expand_dims
import pandas as pd
from scipy.fftpack import shift
#from sqlalchemy import true
import win32com.client as client
import datetime
import tabula

today = datetime.date.today()

# date regex [0-3][0-9]-[A-Z][a-z][a-z]


def getTableEmail():

    msg = []
    APCC_msg = []

    # create instance of Outlook
    outlook = client.Dispatch('Outlook.Application')

    # get the inbox
    namespace = outlook.GetNameSpace('MAPI')
    inbox = namespace.GetDefaultFolder(6)

    # the email I want to download a file from

    # get only mail items from the inbox (other items can exists and will return an error if you try get the subject line of a non-mail item)
    mail_items = [item for item in inbox.Items if item.Class == 43]

    lastWeekDateTime = datetime.datetime.now() - datetime.timedelta(days=8)

    last_week_mail = inbox.Items.Restrict(
        "[ReceivedTime] >= '"+lastWeekDateTime.strftime('%m/%d/%Y %H:%M %p')+"'")

    c = 0
    for message in last_week_mail:
        # print(message.subject)
        # print(message.ReceivedTime)

        APCC_msg.append(message)

        c = c+1

    # filter to the target email
    filtered = [
        item for item in mail_items if item.Unread and item.Senton.date() == today or (item.Unread and 'EMFP Overtime' in item.Subject)]

    #print("The count of meesgaes unread from past week ==", c)

    if len(filtered) == 0:
        print("No filtered email(s)")
        return

    n = 0
    # get the first item if it exists (assuming the there is only one item to get)
    while n < len(filtered):

        if len(filtered) != 0:
            target_email = filtered[n]
            n += 1

            msg.append(target_email)

        elif len(filtered) != 0:
            print("No Email")
    print(n)

    # if there is APCC workplan, download it and save it to source.
    # while n < len(APCC_filter):

    #     if len(APCC_filter) != 0:
    #         target_email = APCC_filter[n]
    #         n += 1

    #         APCC_msg.append(target_email)

    #     elif len(APCC_filter) != 0:
    #         print("No Email")

    return msg, APCC_msg

# APCC


def APCClogic(df):
    date = today.strftime('%#d-%b')
    #date = '30-May'

    first_shift = []
    second_shift = []

    #df, factName = getTableEmail(email_dir)

    df['group_no'] = df.isnull().all(axis=1).cumsum()

    # dictionary
    d = {i: df.loc[df.group_no == i, ['Date', 'Line', 'Frontend', 'Backend']]
         for i in range(0, df.group_no.iat[-1])}

    # k v = key value
    new_d = {k: v for (k, v) in d.items() if not v.empty}

    result = []
    for k, v in new_d.items():

        # select dataframe by
        if date in v.to_string():
            df = v.dropna(how='all')
            result.append(df)

    first_shift = result[0]['Frontend']
    second_shift = result[0]['Backend']

    return first_shift, second_shift, date

    # print(df)

    # BRH dont have shift times so maybe just put the hours?


def BRHlogic(df, factName):

    notebook_df = df.loc[df['LOB'] == 'NOTEBOOK']
    desktop_df = df.loc[df['LOB'] == 'DESKTOP']
    server_df = df.loc[df['LOB'] == 'Server']
    aio_df = df.loc[df['LOB'] == 'AIO']

    nb_hrs1 = pd.to_numeric(notebook_df['HRS1']).max()
    nb_hrs2 = pd.to_numeric(notebook_df['HRS2']).max()
    nb_UPH1 = pd.to_numeric(notebook_df['UPH1']).sum()
    nb_UPH2 = pd.to_numeric(notebook_df['UPH2']).sum()

    dt_hrs1 = pd.to_numeric(desktop_df['HRS1']).max()
    dt_hrs2 = pd.to_numeric(desktop_df['HRS2']).max()
    dt_UPH1 = pd.to_numeric(desktop_df['UPH1']).sum()
    dt_UPH2 = pd.to_numeric(desktop_df['UPH2']).sum()

    server_hrs1 = pd.to_numeric(server_df['HRS1']).max()
    server_hrs2 = pd.to_numeric(server_df['HRS2']).max()
    server_UPH1 = pd.to_numeric(server_df['UPH1']).sum()
    server_UPH2 = pd.to_numeric(server_df['UPH2']).sum()

    aio_hrs1 = pd.to_numeric(aio_df['HRS1']).max()
    aio_hrs2 = pd.to_numeric(aio_df['HRS2']).max()
    aio_UPH1 = pd.to_numeric(aio_df['UPH1']).sum()
    aio_UPH2 = pd.to_numeric(aio_df['UPH2']).sum()

    first_hrs = [nb_hrs1, dt_hrs1, server_hrs1, aio_hrs1]
    second_hrs = [nb_hrs2, dt_hrs2, server_hrs2, aio_hrs2]

    first_UPH = [nb_UPH1, dt_UPH1, server_UPH1, aio_UPH1]
    second_UPH = [nb_UPH2, dt_UPH2, server_UPH2, aio_UPH2]

    return first_hrs, second_hrs, first_UPH, second_UPH

    # ICC


def ICClogic(df, factName):

    df1 = []

    df2 = []

    #df, factoryname = getTableEmail(ICCemail)

    #shift_time = df.loc[1]['BACK END'].replace(" ", "")

    # print(df)

    for i, row in df.iterrows():
        data = str(row['FRONT END']).replace(" ", "")

        if len(data) > 12:
            chunk, chunk_size = len(data), len(data)//2

            split_data = [data[i:i+chunk_size]
                          for i in range(0, chunk, chunk_size)]
            df1.append(split_data)

            front_df = pd.DataFrame(
                df1, columns=["first_shift", "second_shift"])
        else:
            chunk, chunk_size = len(data), len(data)//1

            split_data = [data[i:i+chunk_size]
                          for i in range(0, chunk, chunk_size)]
            df1.append(split_data)

            front_df = pd.DataFrame(df1, columns=["first_shift"])

    # Backend array
    for i, row in df.iterrows():
        back_data = str(row['BACK END']).replace(" ", "")

        if len(back_data) > 12:
            chunk, chunk_size = len(back_data), len(back_data)//2

            split_data = [back_data[i:i+chunk_size]
                          for i in range(0, chunk, chunk_size)]
            df2.append(split_data)

            back_df = pd.DataFrame(
                df2, columns=["first_shift", "second_shift"])
        else:
            chunk, chunk_size = len(back_data), len(back_data)//1

            split_data = [back_data[i:i+chunk_size]
                          for i in range(0, chunk, chunk_size)]
            df2.append(split_data)

            back_df = pd.DataFrame(
                df2, columns=["first_shift"])

    return front_df, back_df


def EMFPlogic(f):
    #print("EMFP Logic")
    df = tabula.read_pdf(
        f, stream=True, pages='all')

    try:
        for i in df:

            # print(i)
            if today.strftime('%B') in i.values:
                #print("Month :", today.strftime('%B'))

                df1 = i.dropna(how='all', thresh=4, axis=1)
                df2 = df1.dropna(how='all', thresh=5)
                df3 = df2.reset_index(drop=True)
                df4 = df3.iloc[-2:, :]

                df5 = pd.DataFrame()
                templist = []
                for col in df4:
                    templist.append(df4[col].str.split(expand=True))

                df5 = pd.concat(templist, axis=1)
                df5.columns = np.arange(len(df5.columns))

                return df5[today.day]

    except Exception as e:
        print("No data for", today.strftime('%#d %B'))


# getTableEmail()
# BRHlogic()
# EMFPlogic()
