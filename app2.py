from tokenize import group
from nbformat import write
import pandas as pd
from plotly import data
import streamlit as st
from streamlit_option_menu import option_menu


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


st.header('Consolidate View')

group = option_menu(
    menu_title="Factory",
    options=['ALL', 'BRH1', 'EMFP', 'CCC4', 'APCC', 'CCC2', 'CCC6', 'ICC'],
    menu_icon="building",
    icons=['bank2', 'bank2', 'bank2', 'bank2',
           'bank2', 'bank2', 'bank2', 'bank2', ],
    styles={
        "menu_icon": {"color": "#3498DB"},
        "nav-link-selected": {"background-color": "#3498DB"}
    },
    orientation="horizontal"
)

hide_st_style = """ 
                <style>
                #MainMenu {visibility: hidden;}
                header {visibility: hidden;}
                footer {visibility: hidden;}
                </style>
                """

st.markdown(hide_st_style, unsafe_allow_html=True)

# -- LOAD DATAFRAME
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

st.subheader('Update On : ' + time1)

# st.write(df1.astype(str))
# df2 = df1.astype(str)

col1, col2, col3 = st.columns(3)


# group = st.selectbox('Choose the factory',('All','BRH1','EMFP','CCC4','APCC','CCC2','CCC6','ICC'))
st.write('You have selected', group)

ccc4 = df1.iloc[0:9, :].astype(str)
ccc2 = df1.iloc[9:18, :].astype(str)
ccc6 = df1.iloc[17:21, :].astype(str)
apcc = df1.iloc[25:33, :].astype(str)
emfp = df1.iloc[41:44, :].astype(str)
brh1 = df1.iloc[49:58, :].astype(str)
icc = df1.iloc[33:41, :].astype(str)

if group == 'ALL':
    st.write(df1.astype(str))

elif group == 'BRH1':
    D1 = pd.read_excel(excel_file,
                       sheet_name=sheet_name,
                       usecols='J:J',
                       nrows=1,
                       header=54)

    DD1 = D1.iloc[0][0]
    st.write('Date: ' + DD1)
    st.write(brh1)

elif group == 'CCC2':
    D2 = pd.read_excel(excel_file,
                       sheet_name=sheet_name,
                       usecols='J:J',
                       nrows=1,
                       header=14)

    DD2 = D2.iloc[0][0]
    st.write('Date: ' + DD2)
    st.write(ccc2)

elif group == 'CCC4':
    D3 = pd.read_excel(excel_file,
                       sheet_name=sheet_name,
                       usecols='J:J',
                       nrows=1,
                       header=5)

    DD3 = D3.iloc[0][0]
    st.write('Date: ' + DD3)
    st.write(ccc4)

elif group == 'CCC6':
    D4 = pd.read_excel(excel_file,
                       sheet_name=sheet_name,
                       usecols='J:J',
                       nrows=1,
                       header=22)

    DD4 = D4.iloc[0][0]
    st.write('Date: ' + DD4)
    st.write(ccc6)

elif group == 'APCC':
    D5 = pd.read_excel(excel_file,
                       sheet_name=sheet_name,
                       usecols='J:J',
                       nrows=1,
                       header=30)

    DD5 = D5.iloc[0][0]
    st.write('Date: ' + DD5)
    st.write(apcc)

elif group == 'ICC':
    D6 = pd.read_excel(excel_file,
                       sheet_name=sheet_name,
                       usecols='J:J',
                       nrows=1,
                       header=38)

    DD6 = D6.iloc[0][0]
    st.write('Date: ' + DD6)
    st.write(icc)

elif group == 'EMFP':
    D7 = pd.read_excel(excel_file,
                       sheet_name=sheet_name,
                       usecols='J:J',
                       nrows=1,
                       header=46)

    DD7 = D7.iloc[0][0]
    st.write('Date: ' + DD7)
    st.write(emfp)


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
