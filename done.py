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


c1,c2,c3, c4 = st.columns(4)
with c1:
    #calculate income revenue
    income_last_year = income[1:13].astype(float).mean()
    income_this_year = income[13:].astype(float).mean()
    curr_change_percent = round(((income_this_year - income_last_year) / income_last_year * 100), 1)
    curr_change_percent = str(curr_change_percent)  + '%' + ' vs. prior year'
    st.metric("REVENUE", round(income_this_year, 2),delta=curr_change_percent)

    #calculate operating expenses
    operating_Expenses_last_year = Total_operating_Expenses[1:13].astype(float).mean()
    operating_Expenses_this_year = Total_operating_Expenses[13:].astype(float).mean()
    curr_expenses_change_percent = round(((operating_Expenses_this_year - operating_Expenses_last_year) / operating_Expenses_last_year * 100), 2)
    curr_expenses_change_percent = str(curr_expenses_change_percent) + '%' + ' vs. prior year'
    st.metric("OPERATING EXPENSES", round(operating_Expenses_this_year, 2), delta=curr_expenses_change_percent)

    #calculate quick ratio
    quick_ratio_last_year = quick_ratio[1:13].astype(float).mean()
    quick_ratio_this_year = quick_ratio[13:].astype(float).mean()
    curr_quickratio_change_percent = round(((quick_ratio_this_year - quick_ratio_last_year)/ quick_ratio_last_year * 100), 2)
    curr_quickratio_change_percent = str(curr_quickratio_change_percent) + ' vs. prior year'
    st.metric("QUICK RATIO", round(quick_ratio_this_year, 2), delta=curr_quickratio_change_percent)
   
with c2:
    #COST of good sold
    cost_of_good_sold_last_year = cost_of_good_sold[1:13].astype(float).mean()
    cost_of_good_sold_this_year = cost_of_good_sold[13:].astype(float).mean()
    cost_of_good_sold_change_percent = round(((cost_of_good_sold_this_year - cost_of_good_sold_last_year) / cost_of_good_sold_last_year * 100), 1)
    cost_of_good_sold_change_percent = str(cost_of_good_sold_change_percent) + '%' + ' vs. prior year'
    st.metric("COST OF GOODS SOLD", round(cost_of_good_sold_this_year, 2),delta=cost_of_good_sold_change_percent)

    #operating_profit
    ebit_sold_last_year = operating_profit[1:13].astype(float).mean()
    ebit_sold_this_year = operating_profit[13:].astype(float).mean()
    ebit_sold_change_percent = round(((ebit_sold_this_year - ebit_sold_last_year) / ebit_sold_last_year * 100), 1)
    ebit_sold_change_percent = str(ebit_sold_change_percent) + '%' + ' vs. prior year'
    st.metric("EBIT", round(ebit_sold_this_year, 2),delta=ebit_sold_change_percent)

    #current ratio
    curr_ratio_last_year = curr_ratio[1:13].astype(float).mean()
    curr_ratio_this_year = curr_ratio[13:].astype(float).mean()
    curr_ratiochange_percent = round(((curr_ratio_this_year - curr_ratio_last_year) / curr_ratio_last_year * 100), 1)
    curr_ratiochange_percent = str(curr_ratiochange_percent) + '%' + ' vs. prior year'
    st.metric("CURRENT RATIO", round(curr_ratio_this_year, 2),delta=curr_ratiochange_percent)


with c3:
    #gross profit  
    gross_profit_last_year = gross_profit[1:13].astype(float).mean()
    gross_profit_this_year = gross_profit[13:].astype(float).mean()
    gross_profit_change_percent = round(((gross_profit_this_year - gross_profit_last_year) / gross_profit_last_year * 100), 1)
    gross_profit_change_percent = str(gross_profit_change_percent) + '%' + ' vs. prior year'
    st.metric("GROSS PROFIT", round(gross_profit_this_year, 2),delta=gross_profit_change_percent)

    #net profit
    net_profit_last_year = net_profit[1:13].astype(float).mean()
    net_profit_this_year = net_profit[13:].astype(float).mean()
    net_profit_change_percent = round(((ebit_sold_this_year - ebit_sold_last_year) / ebit_sold_last_year * 100), 1)
    net_profit_change_percent = str(ebit_sold_change_percent) + '%' + ' vs. prior year'
    st.metric("NET PROFIT", round(ebit_sold_this_year, 2),delta=ebit_sold_change_percent)

    #acount receivable
    account_receivable_last_year = account_receivable[1:13].astype(float).mean()
    account_receivable_this_year = account_receivable[13:].astype(float).mean()
    account_receivable_change_percent = round(((account_receivable_this_year - account_receivable_last_year) / account_receivable_last_year * 100), 1)
    account_receivable_change_percent = str(account_receivable_change_percent) + '%' + ' vs. prior year'
    st.metric("ACCOUNT RECEIVABLE", round(account_receivable_this_year, 2),delta=account_receivable_change_percent)

with c4:
    #gross profit margin
    total_revenue = income[1:13].astype(float).sum()
    Cost_of_good_solds = cost_of_good_sold[1:13].astype(float).sum()
    gross_profit_margin_last_year = round((((total_revenue - Cost_of_good_solds) / total_revenue)* 100), 1)

    total_revenue_this_year = income[13:].astype(float).sum()
    Cost_of_good_solds_this_year = cost_of_good_sold[13:].astype(float).sum()
    gross_profit_margin_this_year = round((((total_revenue_this_year - Cost_of_good_solds_this_year) / total_revenue_this_year)* 100), 1)
    gross_profit_margin_change_percent = round(((gross_profit_margin_this_year - gross_profit_margin_last_year) / gross_profit_margin_last_year * 100), 1)
    gross_profit_margin_change_percent = str(gross_profit_margin_change_percent) + '%' + ' vs. prior year'
    st.metric("GROSS PROFIT MARGIN", round(gross_profit_margin_this_year, 2),delta=gross_profit_margin_change_percent)

    #net profit margin
    Net_income = net_profit[1:13].astype(float).sum()
    Net_income_this_year = net_profit[13:].astype(float).sum()
    
    net_profit_margin_last_year = round(((Net_income / total_revenue) * 100), 1)
    net_profit_margin_this_year = round(((Net_income_this_year / total_revenue_this_year) * 100), 1)
    net_profit_margin_change_percent = round(((net_profit_margin_this_year - net_profit_margin_last_year) / net_profit_margin_last_year * 100), 1)
    net_profit_margin_change_percent = str(net_profit_margin_change_percent) + '%' + ' vs. prior year'
    st.metric("NET PROFIT MARGIN", round(net_profit_margin_this_year, 2),delta=net_profit_margin_change_percent)

    #account payable
    account_payable_last_year = account_payable[1:13].astype(float).mean()
    account_payable_this_year = account_payable[13:].astype(float).mean()
    account_payable_change_percent = round(((account_payable_this_year - account_payable_last_year) / account_payable_last_year * 100), 1)
    account_payable_change_percent = str(account_payable_change_percent) + '%' + ' vs. prior year'
    st.metric("ACCOUNT PAYABLE", round(account_payable_this_year, 2),delta=account_payable_change_percent)
