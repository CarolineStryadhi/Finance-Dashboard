import pandas as pd
import streamlit as st
import numpy as np
import matplotlib.pyplot as plt

excel_file = 'C:/Users/norma/OneDrive/Pictures/folder tugas 2/Finance Dashboard Template.xlsx' #path
file = pd.read_excel(excel_file, engine='openpyxl')
st.write(file)

#logic
Income = file.iloc[0]
Income['Indicator Name'] = pd.to_numeric(Income['Indicator Name'], errors='coerce')
Income_oldyear = Income.iloc[1:13].sum()
Income_curryear = Income.iloc[14:].sum()
Income_change = Income_curryear - Income_oldyear
st.write(Income_oldyear)
st.write(Income_curryear)

Cost_of_Goods_Sold = file.iloc[1] #file.iloc utk mendapat data ke subset n-row
st.write(Cost_of_Goods_Sold)

Gross_Profit = file.iloc[2]
st.write(Gross_Profit)
Total_Operating_Expenses   = file.iloc[3]
st.write(Total_Operating_Expenses)

Operating_Profit = file.iloc[4]
st.write(Operating_Profit)

Taxes = file.iloc[5]
st.write(Taxes)

Net_Profit = file.iloc[6]
st.write(Net_Profit)

Net_Profit_Margin = file.iloc[7]
st.write(Net_Profit_Margin)

Expenses = file.iloc[8]
st.write(Expenses)

Cash_at_EOM = file.iloc[9]
st.write(Cash_at_EOM)

Quick_Ratio = file.iloc[10]
st.write(Quick_Ratio)

Current_Ratio = file.iloc[11]
st.write(Current_Ratio)

Accounts_Receivable = file.iloc[12]
st.write(Accounts_Receivable)

Accounts_Payable = file.iloc[13]
st.write(Accounts_Payable)

#tampilan UI
st.header("PROFIT & LOSS STATEMENT", divider= True)
col1, col2, col3 = st.columns(3)
col1.metric("revenue", Income_curryear, Income_change)
col1.line_chart(Income, x = None, y = None)
col1.write("col1")
col1.write("col1")

col2.write("col2")
col2.write("col2")
col2.write("col2")

col3.write("col3")
col3.write("col3")
col3.write("col3")
