import pandas as pd
from requests import head
import streamlit as st
import json
import utils as utl
from views import home, about


st.set_page_config(
    layout="wide", page_title='Factories Workplans Consolidated', page_icon="assets\images\settings.png")

st.set_option('deprecation.showPyplotGlobalUse', False)
utl.inject_custom_css()
utl.navbar_component()


def navigation():
    route = utl.get_current_route()
    if route == "home":
        home.load_view()
    elif route == "about":
        about.load_view()
    elif route == None:
        home.load_view()


navigation()
