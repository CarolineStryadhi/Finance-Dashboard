import pandas as pd
import streamlit as st
import numpy as np

st.header("PROFIT & LOSS STATEMENT", divider= True)
col1, col2, col3 = st.columns(3)
col1.metric("nama", 78)
col1.write("col1")
col1.write("col1")

col2.write("col2")
col2.write("col2")
col2.write("col2")

col3.write("col3")
col3.write("col3")
col3.write("col3")

excel_file = 'C:/Users/norma/OneDrive/Pictures/folder tugas 2/Finance Dashboard Template.xlsx' #path
file = pd.read_excel(excel_file, engine='openpyxl')
st.write(file)