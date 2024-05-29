
import streamlit as st
import ui as ui
import data_processing 
import pandas as pd
import numpy as np
import re
from datetime import datetime
import openpyxl
from openpyxl.styles import Font, Alignment, Border, Side
from openpyxl.styles.fills import PatternFill
from openpyxl import Workbook
from openpyxl import load_workbook
import os

pd.set_option('display.precision', 2)

st.set_page_config(
    page_title="APP - Auto Forms",
    layout="wide"
)

def main():
    pages = {
        "Gerar formulários": ui.page_gera_forms
    }

    with st.sidebar:
        st.title("FGV IBRE - SPDO")
        page = st.radio("Menu", tuple(pages.keys()))
        st.markdown('---')

    pages[page]()

if __name__ == "__main__":
    main()