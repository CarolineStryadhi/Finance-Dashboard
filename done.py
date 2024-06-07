import pandas as pd
import streamlit as st
import numpy as np
import matplotlib.pyplot as plt

excel_file = 'C:/Users/norma/OneDrive/Pictures/folder tugas 2/Finance Dashboard Template.xlsx' #path
file = pd.read_excel(excel_file, engine='openpyxl')
st.write(file)

income = file.iloc[0,:]
cost_of_good_sold = file.iloc[1,:]
gross_profit = file.iloc[3,:]
Total_operating_Expenses = file.iloc[4,:]
operating_profit = file.iloc[5,:]
taxes = file.iloc[6,:]
net_profit = file.iloc[7, :]
expenses = file.iloc[8,:]
cash_at_eom = file.iloc[9,:]
quick_ratio = file.iloc[10,:]
curr_ratio = file.iloc[11, :]
account_receivable = file.iloc[12, :]
account_payable = file.iloc[13,:]


c1,c2,c3 = st.columns(3)
with c1:
    #calculate income revenue
    income_last_year = income[1:13].astype(float).sum()
    income_this_year = income[13:].astype(float).sum()
    curr_change_percent = round(((income_this_year - income_last_year) / income_last_year * 100), 1)
    curr_change_percent = str(curr_change_percent)  + '%' + ' vs. prior month'
    st.metric("REVENUE", round(income_this_year, 3),delta=curr_change_percent)

    #calculate operating expenses
    operating_Expenses_last_year = Total_operating_Expenses[1:13].astype(float).sum()
    operating_Expenses_this_year = Total_operating_Expenses[13:].astype(float).sum()
    curr_expenses_change_percent = round(((operating_Expenses_this_year - operating_Expenses_last_year) / operating_Expenses_last_year * 100), 2)
    curr_expenses_change_percent = str(curr_expenses_change_percent) + '%' + ' vs. prior month'
    st.metric("OPERATING EXPENSES", round(operating_Expenses_this_year, 3), delta=curr_expenses_change_percent)

    #calculate quick ratio
    quick_ratio_last_year = quick_ratio[1:13].astype(float).mean()
    quick_ratio_this_year = quick_ratio[13:].astype(float).mean()
    curr_quickratio_change_percent = round(((quick_ratio_this_year - quick_ratio_last_year)/ quick_ratio_last_year * 100), 2)
    curr_quickratio_change_percent = str(curr_quickratio_change_percent) + ' vs. prior month'
    st.metric("QUICK RATIO", round(quick_ratio_this_year, 2), delta=curr_quickratio_change_percent)