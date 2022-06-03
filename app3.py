import streamlit as st
import pandas as pd
import altair as alt

from urllib.error import URLError

st.set_page_config(page_title='Dell Factory Consolidate View',
                   page_icon='Dell_Logo_Blue_rgb.png')

st.markdown('<link rel="stylesheet" href="https://maxcdn.bootstrapcdn.com/bootstrap/4.0.0/css/bootstrap.min.css" integrity="sha384-Gn5384xqQ1aoWXA+058RXPxPg6fy4IWvTNh0E263XmFcJlSAwiGgFAW/dAiS6JXm" crossorigin="anonymous">', unsafe_allow_html=True)


st.markdown("""
<nav class="navbar fixed-top navbar-expand-lg navbar-dark" style="background-color: #3498DB; font:serif;">
  <a class="navbar-brand" style="max-width: 500px;
  margin: auto;">
  <img  src ="media/47d6a03cf55bdc885d96a614db437d6c9e03a03651f46b79dc178db4.png" alt="" width="30" height="30">
     Dell Global Factory Workplan
  </a>
  <button class="navbar-toggler" type="button" data-toggle="collapse" data-target="#navbarNav" aria-controls="navbarNav" aria-expanded="false" aria-label="Toggle navigation">
    <span class="navbar-toggler-icon"></span>
  </button>
  <div class="collapse navbar-collapse" id="navbarNav">
  </div>
</nav>
""", unsafe_allow_html=True)

hide_st_style = """ 
                <style>
                #MainMenu {visibility: hidden;}
                header {visibility: hidden;}
                footer {visibility: hidden;}
                </style>
                """

st.markdown(hide_st_style, unsafe_allow_html=True)


excel_file = r'Consolidated Factory Workplan.xlsx'
sheet_name = 'Workplans'
df = pd.read_excel(excel_file,
                   sheet_name=sheet_name,
                   usecols='B:K',
                   skiprows=(0, 1, 2, 3, 4, 5),
                   header=None)

df1 = df.fillna('')
time = pd.read_excel(excel_file,
                     sheet_name=sheet_name,
                     usecols='F:F',
                     nrows=1,
                     header=3)

time1 = time.iloc[0][0]
st.header('Consolidate View')
st.subheader('Update On : ' + time1)

ccc4 = df1.iloc[0:9, :].astype(str)
ccc2 = df1.iloc[9:18, :].astype(str)
ccc6 = df1.iloc[17:21, :].astype(str)
apcc = df1.iloc[25:33, :].astype(str)
emfp = df1.iloc[41:44, :].astype(str)
brh1 = df1.iloc[49:58, :].astype(str)
icc = df1.iloc[33:41, :].astype(str)

try:
    factory = st.multiselect(
        'Choose the factory', ['All', 'BRH1', 'EMFP',
                               'CCC4', 'APCC', 'CCC2', 'CCC6', 'ICC']
    )

    # if not factory:
    #     st.error("Please select at least one country.")

    if factory:

        if 'All' in factory:
            st.write(df1.astype(str))

        if 'BRH1' in factory:
            D1 = pd.read_excel(excel_file,
                               sheet_name=sheet_name,
                               usecols='J:J',
                               nrows=1,
                               header=54)
            t1 = pd.read_excel(excel_file,
                               sheet_name=sheet_name,
                               usecols='D:D',
                               nrows=1,
                               header=54)

            DD1 = D1.iloc[0][0]
            tt1 = t1.iloc[0][0]
            st.write('Factory: ' + tt1)
            st.write('Date: ' + DD1)
            st.write(brh1)

        if 'CCC2' in factory:
            D2 = pd.read_excel(excel_file,
                               sheet_name=sheet_name,
                               usecols='J:J',
                               nrows=1,
                               header=14)
            t2 = pd.read_excel(excel_file,
                               sheet_name=sheet_name,
                               usecols='D:D',
                               nrows=1,
                               header=14)

            DD2 = D2.iloc[0][0]
            tt2 = t2.iloc[0][0]
            st.write('Date: ' + tt2)
            st.write('Date: ' + DD2)
            st.write(ccc2)

        if 'CCC4' in factory:
            D3 = pd.read_excel(excel_file,
                               sheet_name=sheet_name,
                               usecols='J:J',
                               nrows=1,
                               header=5)
            t3 = pd.read_excel(excel_file,
                               sheet_name=sheet_name,
                               usecols='D:D',
                               nrows=1,
                               header=5)

            DD3 = D3.iloc[0][0]
            tt3 = t3.iloc[0][0]
            st.write('Date: ' + tt3)
            st.write('Date: ' + DD3)
            st.write(ccc4)

        if 'CCC6' in factory:
            D4 = pd.read_excel(excel_file,
                               sheet_name=sheet_name,
                               usecols='J:J',
                               nrows=1,
                               header=22)
            t4 = pd.read_excel(excel_file,
                               sheet_name=sheet_name,
                               usecols='D:D',
                               nrows=1,
                               header=22)

            DD4 = D4.iloc[0][0]
            tt4 = t4.iloc[0][0]
            st.write('Date: ' + tt4)
            st.write('Date: ' + DD4)
            st.write(ccc6)

        if 'APCC' in factory:
            D5 = pd.read_excel(excel_file,
                               sheet_name=sheet_name,
                               usecols='J:J',
                               nrows=1,
                               header=30)
            t5 = pd.read_excel(excel_file,
                               sheet_name=sheet_name,
                               usecols='D:D',
                               nrows=1,
                               header=30)

            DD5 = D5.iloc[0][0]
            tt5 = t5.iloc[0][0]
            st.write('Date: ' + tt5)
            st.write('Date: ' + DD5)
            st.write(apcc)

        if 'ICC' in factory:
            D6 = pd.read_excel(excel_file,
                               sheet_name=sheet_name,
                               usecols='J:J',
                               nrows=1,
                               header=38)
            t6 = pd.read_excel(excel_file,
                               sheet_name=sheet_name,
                               usecols='D:D',
                               nrows=1,
                               header=38)

            DD6 = D6.iloc[0][0]
            tt6 = t6.iloc[0][0]
            st.write('Date: ' + tt6)
            st.write('Date: ' + DD6)
            st.write(icc)

        if 'EMFP' in factory:
            D7 = pd.read_excel(excel_file,
                               sheet_name=sheet_name,
                               usecols='J:J',
                               nrows=1,
                               header=46)
            t7 = pd.read_excel(excel_file,
                               sheet_name=sheet_name,
                               usecols='D:D',
                               nrows=1,
                               header=46)

            DD7 = D7.iloc[0][0]
            tt7 = t7.iloc[0][0]
            st.write('Date: ' + tt7)
            st.write('Date: ' + DD7)
            st.write(emfp)

        # data = df.loc[factory]
        # data /= 1000000.0
        # st.write("### Gross Agricultural Production ($B)", data.sort_index())

        # data = data.T.reset_index()
        # data = pd.melt(data, id_vars=["index"]).rename(
        # columns={"index": "year", "value": "Gross Agricultural Product ($B)"}
        # )
    else:
        st.error("Please select at least one factory.")


except URLError as e:
    st.error(
        """
**This demo requires internet access.**



Connection error: %s
"""
        % e.reason
    )

st.write('To download the full page of consolidate view, click download button below :')
btn = st.download_button(
    label='Download File',
    data=excel_file,
    file_name=excel_file,
)

st.markdown("""
<script src="https://code.jquery.com/jquery-3.2.1.slim.min.js" integrity="sha384-KJ3o2DKtIkvYIK3UENzmM7KCkRr/rE9/Qpg6aAZGJwFDMVNA/GpGFF93hXpG5KkN" crossorigin="anonymous"></script>
<script src="https://cdnjs.cloudflare.com/ajax/libs/popper.js/1.12.9/umd/popper.min.js" integrity="sha384-ApNbgh9B+Y1QKtv3Rn7W3mgPxhU9K/ScQsAP7hUibX39j7fakFPskvXusvfa0b4Q" crossorigin="anonymous"></script>
<script src="https://maxcdn.bootstrapcdn.com/bootstrap/4.0.0/js/bootstrap.min.js" integrity="sha384-JZR6Spejh4U02d8jOt6vLEHfe/JQGiRRSQQxSfFWpi1MquVdAyjUar5+76PVCmYl" crossorigin="anonymous"></script>
""", unsafe_allow_html=True)
