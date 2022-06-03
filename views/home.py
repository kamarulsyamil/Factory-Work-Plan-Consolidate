import pandas as pd
from requests import head
import streamlit as st
import json
import utils as utl


def load_view():
    st.header('GPP Workplan (Shift times)')

    # LOAD EXCEL
    # read config file

    with open(r"app\configuration\tool_config.json") as config_file:
        config = json.load(config_file)
        file_dir = config['file_dir']

    excel_file = file_dir['main_excel']

    sheet_name = 'Workplans'

    df = pd.read_excel(excel_file, sheet_name=sheet_name,
                       usecols='B:K', header=7)
    df2 = df.fillna('')

    datedf = pd.read_excel(excel_file, sheet_name=sheet_name,
                           usecols='B:K', header=None)

    date = datedf.iloc[4][5]

    st.subheader('Updated on :' + date)
    # st.write(date)

    # st.write(df2.astype(str))

    option = st.selectbox(
        'Choose factory',
        ('ALL', 'CCC4', 'CCC2', 'CCC6', 'APCC', 'ICC', 'EMFP', 'BRH1'))

    st.write('Selected:', option)

    if option == 'ALL':
        st.write(datedf.fillna('').astype(str))

    elif option == 'CCC4':
        st.write('Date: ', datedf.iloc[6][9])

        st.write(df2.iloc[0:7, :].astype(str))

    elif option == 'CCC2':
        st.write('Date: ', datedf.iloc[15][9])

        st.write(df2.iloc[9:15, :].astype(str))

    elif option == 'CCC6':
        st.write('Date: ', datedf.iloc[23][9])

        st.write(df2.iloc[17:18, :].astype(str))

    elif option == 'APCC':
        st.write('Date: ', datedf.iloc[31][9])

        st.write(df2.iloc[25:31, :].astype(str))

    elif option == 'ICC':
        st.write('Date: ', datedf.iloc[39][9])

        st.write(df2.iloc[33:39, :].astype(str))

    elif option == 'EMFP':
        st.write('Date: ', datedf.iloc[47][9])

        st.write(df2.iloc[41:42, :].astype(str))
        st.write(df2.iloc[57:58, :].astype(str))

    elif option == 'BRH1':
        st.write('Date: ', datedf.iloc[55][9])

        st.write(df2.iloc[49:56, :].astype(str))

    with open(file_dir['main_excel'], 'rb') as file:
        btn = st.download_button(
            label="Download workplan",
            data=file,
            file_name=file_dir['main_excel']
        )

    # if st.button('Say hello'):
    #     st.write('hi')

        # main()
