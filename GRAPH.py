import pandas as pd
import streamlit as st
import numpy as np
import matplotlib.pyplot as plt


excel_file = './Finance Dashboard Template.xlsx' #path
file = pd.read_excel(excel_file, engine='openpyxl')
# st.write(file)

print(file)

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

print(income)

c1,c2,c3, c4 = st.columns(4)
with c1:
    #calculate income revenue average
    income_last_year = income[13].astype(float)
    print(income_last_year)
    income_this_year = income[24].astype(float)
    print(income_this_year)

    #Pisahkan data tahun lalu dan tahun ini ke beda variabel
    last_year = list(income[1:13])
    this_year = list(income[13:])

    #gunakan rumus untuk mendapatkan persentase pergantian
    curr_change_percent = round(((income_this_year - income_last_year) / income_last_year * 100), 1)
    curr_change_percent = str(curr_change_percent)  + '%' + ' vs. prior year'

    #gabung menjadi satu dataframe data last dan current year
    year = pd.DataFrame({
        "last":last_year,
        "current":this_year
    })

    #data hasil kalkulasi beserta persentase perubahan akan dimasukkan ke dalam elemen metric
    st.metric("REVENUE", round(income_this_year, 2),delta=curr_change_percent)
    
    #data akan dilabel dan dichart di dalam line char
    st.line_chart(data=year)

    #calculate operating expenses
    operating_Expenses_last_year = Total_operating_Expenses[13].astype(float)
    operating_Expenses_this_year = Total_operating_Expenses[24].astype(float)
    curr_expenses_change_percent = round(((operating_Expenses_this_year - operating_Expenses_last_year) / operating_Expenses_last_year * 100), 2)
    curr_expenses_change_percent = str(curr_expenses_change_percent) + '%' + ' vs. prior year'
    st.metric("OPERATING EXPENSES", round(operating_Expenses_this_year, 2), delta=curr_expenses_change_percent)

    #calculate quick ratio
    quick_ratio_last_year = quick_ratio[13].astype(float)
    quick_ratio_this_year = quick_ratio[24].astype(float)
    curr_quickratio_change_percent = round(((quick_ratio_this_year - quick_ratio_last_year)/ quick_ratio_last_year * 100), 2)
    curr_quickratio_change_percent = str(curr_quickratio_change_percent) + ' vs. prior year'
    st.metric("QUICK RATIO", round(quick_ratio_this_year, 2), delta=curr_quickratio_change_percent)
   
with c2:
    #calculate income revenue average
    income_last_year = cost_of_good_sold[13].astype(float)
    income_this_year = cost_of_good_sold[24].astype(float)

    # Generate some data
    last_year = list(cost_of_good_sold[1:13])
    current_year = list(cost_of_good_sold[13:])

    #gunakan rumus untuk mendapatkan persentase pergantian
    curr_change_percent = round(((income_this_year - income_last_year) / income_last_year * 100), 1)
    curr_change_percent = str(curr_change_percent)  + '%' + ' vs. prior year'

    year = pd.DataFrame({
        "last":last_year,
        "current":current_year
    })

    #data hasil kalkulasi beserta persentase perubahan akan dimasukkan ke dalam elemen metric
    st.metric("Cost of Good Sold", round(income_this_year, 2),delta=curr_change_percent)

    st.line_chart(data=year)


    data = pd.DataFrame({'last_year': last_year, 'current_year': current_year})
    print(data)
    #Create a line chart
    c2.line_chart(data)
    #operating_profit
    ebit_sold_last_year = operating_profit[13].astype(float)
    ebit_sold_this_year = operating_profit[24].astype(float)
    ebit_sold_change_percent = round(((ebit_sold_this_year - ebit_sold_last_year) / ebit_sold_last_year * 100), 1)
    ebit_sold_change_percent = str(ebit_sold_change_percent) + '%' + ' vs. prior year'
    st.metric("EBIT", round(ebit_sold_this_year, 2),delta=ebit_sold_change_percent)

    #current ratio
    curr_ratio_last_year = curr_ratio[13].astype(float)
    curr_ratio_this_year = curr_ratio[24].astype(float)
    curr_ratiochange_percent = round(((curr_ratio_this_year - curr_ratio_last_year) / curr_ratio_last_year * 100), 1)
    curr_ratiochange_percent = str(curr_ratiochange_percent) + '%' + ' vs. prior year'
    st.metric("CURRENT RATIO", round(curr_ratio_this_year, 2),delta=curr_ratiochange_percent)

with c3:
    #calculate income revenue average
    income_last_year = gross_profit[13].astype(float)
    income_this_year = gross_profit[24].astype(float)

    # Generate some data
    last_year = list(gross_profit[1:13])
    current_year = list(gross_profit[13:])

    #gunakan rumus untuk mendapatkan persentase pergantian
    curr_change_percent = round(((income_this_year - income_last_year) / income_last_year * 100), 1)
    curr_change_percent = str(curr_change_percent)  + '%' + ' vs. prior year'

    year = pd.DataFrame({
        "last":last_year,
        "current":current_year
    })

    #data hasil kalkulasi beserta persentase perubahan akan dimasukkan ke dalam elemen metric
    st.metric("Gross Profit", round(income_this_year, 2),delta=curr_change_percent)

    st.line_chart(data=year)

####################################################################################################
with c3:
    #gross profit  
    gross_profit_last_year = gross_profit[13].astype(float)
    gross_profit_this_year = gross_profit[24].astype(float)
    gross_profit_change_percent = round(((gross_profit_this_year - gross_profit_last_year) / gross_profit_last_year * 100), 1)
    gross_profit_change_percent = str(gross_profit_change_percent) + '%' + ' vs. prior year'
    st.metric("GROSS PROFIT", round(gross_profit_this_year, 2),delta=gross_profit_change_percent)

    #net profit
    net_profit_last_year = net_profit[13].astype(float)
    net_profit_this_year = net_profit[24].astype(float)
    net_profit_change_percent = round(((ebit_sold_this_year - ebit_sold_last_year) / ebit_sold_last_year * 100), 1)
    net_profit_change_percent = str(ebit_sold_change_percent) + '%' + ' vs. prior year'
    st.metric("NET PROFIT", round(ebit_sold_this_year, 2),delta=ebit_sold_change_percent)

    #acount receivable
    account_receivable_last_year = account_receivable[13].astype(float)
    account_receivable_this_year = account_receivable[24].astype(float)
    account_receivable_change_percent = round(((account_receivable_this_year - account_receivable_last_year) / account_receivable_last_year * 100), 1)
    account_receivable_change_percent = str(account_receivable_change_percent) + '%' + ' vs. prior year'
    st.metric("ACCOUNT RECEIVABLE", round(account_receivable_this_year, 2),delta=account_receivable_change_percent)
####################################################################################################
with c4:
    #calculate income revenue average
    income_last_year = Total_operating_Expenses[13].astype(float)
    income_this_year = Total_operating_Expenses[24].astype(float)

    # Generate some data
    last_year = list(Total_operating_Expenses[1:13])
    current_year = list(Total_operating_Expenses[13:])

    #gunakan rumus untuk mendapatkan persentase pergantian
    curr_change_percent = round(((income_this_year - income_last_year) / income_last_year * 100), 1)
    curr_change_percent = str(curr_change_percent)  + '%' + ' vs. prior year'

    year = pd.DataFrame({
        "last":last_year,
        "current":current_year
    })

    #data hasil kalkulasi beserta persentase perubahan akan dimasukkan ke dalam elemen metric
    st.metric("Total_operating_Expenses", round(income_this_year, 2),delta=curr_change_percent)

    st.line_chart(data=year)
########################################################################################################################################################################
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
    account_payable_last_year = account_payable[13].astype(float)
    account_payable_this_year = account_payable[24].astype(float)
    account_payable_change_percent = round(((account_payable_this_year - account_payable_last_year) / account_payable_last_year * 100), 1)
    account_payable_change_percent = str(account_payable_change_percent) + '%' + ' vs. prior year'
    st.metric("ACCOUNT PAYABLE", round(account_payable_this_year, 2),delta=account_payable_change_percent)
########################################################################################################################################################################
c5, c6, c7, c8 = st.columns(4)
with c5:
    #calculate income revenue average
    income_last_year = operating_profit[13].astype(float)
    income_this_year = operating_profit[24].astype(float)

    # Generate some data
    last_year = list(operating_profit[1:13])
    current_year = list(operating_profit[13:])

    #gunakan rumus untuk mendapatkan persentase pergantian
    curr_change_percent = round(((income_this_year - income_last_year) / income_last_year * 100), 1)
    curr_change_percent = str(curr_change_percent)  + '%' + ' vs. prior year'

    year = pd.DataFrame({
        "last":last_year,
        "current":current_year
    })

    #data hasil kalkulasi beserta persentase perubahan akan dimasukkan ke dalam elemen metric
    st.metric("Operating Profit", round(income_this_year, 2),delta=curr_change_percent)

    st.line_chart(data=year)

with c6:
    #calculate income revenue average
    income_last_year = taxes[13].astype(float)
    income_this_year = taxes[24].astype(float)

    # Generate some data
    last_year = list(taxes[1:13])
    current_year = list(taxes[13:])

    #gunakan rumus untuk mendapatkan persentase pergantian
    curr_change_percent = round(((income_this_year - income_last_year) / income_last_year * 100), 1)
    curr_change_percent = str(curr_change_percent)  + '%' + ' vs. prior year'

    year = pd.DataFrame({
        "last":last_year,
        "current":current_year
    })

    #data hasil kalkulasi beserta persentase perubahan akan dimasukkan ke dalam elemen metric
    st.metric("Taxes", round(income_this_year, 2),delta=curr_change_percent)

    st.line_chart(data=year)

with c7:
    #calculate income revenue average
    income_last_year = net_profit[13].astype(float)
    income_this_year = net_profit[24].astype(float)

    # Generate some data
    last_year = list(net_profit[1:13])
    current_year = list(net_profit[13:])

    #gunakan rumus untuk mendapatkan persentase pergantian
    curr_change_percent = round(((income_this_year - income_last_year) / income_last_year * 100), 1)
    curr_change_percent = str(curr_change_percent)  + '%' + ' vs. prior year'

    year = pd.DataFrame({
        "last":last_year,
        "current":current_year
    })

    #data hasil kalkulasi beserta persentase perubahan akan dimasukkan ke dalam elemen metric
    st.metric("Net Profit", round(income_this_year, 2),delta=curr_change_percent)

    st.line_chart(data=year)

with c8:
    #calculate income revenue average
    income_last_year = expenses[13].astype(float)
    income_this_year = expenses[24].astype(float)

    # Generate some data
    last_year = list(expenses[1:13])
    current_year = list(expenses[13:])

    #gunakan rumus untuk mendapatkan persentase pergantian
    curr_change_percent = round(((income_this_year - income_last_year) / income_last_year * 100), 1)
    curr_change_percent = str(curr_change_percent)  + '%' + ' vs. prior year'

    year = pd.DataFrame({
        "last":last_year,
        "current":current_year
    })

    #data hasil kalkulasi beserta persentase perubahan akan dimasukkan ke dalam elemen metric
    st.metric("Expenses", round(income_this_year, 2),delta=curr_change_percent)

    st.line_chart(data=year)

c9, c10, c11, c12 = st.columns(4)
with c9:
    #calculate income revenue average
    income_last_year = cash_at_eom[13].astype(float)
    income_this_year = cash_at_eom[24].astype(float)

    # Generate some data
    last_year = list(cash_at_eom[1:13])
    current_year = list(cash_at_eom[13:])

    #gunakan rumus untuk mendapatkan persentase pergantian
    curr_change_percent = round(((income_this_year - income_last_year) / income_last_year * 100), 1)
    curr_change_percent = str(curr_change_percent)  + '%' + ' vs. prior year'

    year = pd.DataFrame({
        "last":last_year,
        "current":current_year
    })

    #data hasil kalkulasi beserta persentase perubahan akan dimasukkan ke dalam elemen metric
    st.metric("Cash at EOM", round(income_this_year, 2),delta=curr_change_percent)

    st.line_chart(data=year)

with c10:
    #calculate income revenue average
    income_last_year = quick_ratio[13].astype(float)
    income_this_year = quick_ratio[24].astype(float)

    # Generate some data
    last_year = list(quick_ratio[1:13])
    current_year = list(quick_ratio[13:])

    #gunakan rumus untuk mendapatkan persentase pergantian
    curr_change_percent = round(((income_this_year - income_last_year) / income_last_year * 100), 1)
    curr_change_percent = str(curr_change_percent)  + '%' + ' vs. prior year'

    year = pd.DataFrame({
        "last":last_year,
        "current":current_year
    })

    #data hasil kalkulasi beserta persentase perubahan akan dimasukkan ke dalam elemen metric
    st.metric("Quick Ratio", round(income_this_year, 2),delta=curr_change_percent)

    st.line_chart(data=year)

with c11:
    #calculate income revenue average
    income_last_year = curr_ratio[13].astype(float)
    income_this_year = curr_ratio[24].astype(float)

    # Generate some data
    last_year = list(curr_ratio[1:13])
    current_year = list(curr_ratio[13:])

    #gunakan rumus untuk mendapatkan persentase pergantian
    curr_change_percent = round(((income_this_year - income_last_year) / income_last_year * 100), 1)
    curr_change_percent = str(curr_change_percent)  + '%' + ' vs. prior year'

    year = pd.DataFrame({
        "last":last_year,
        "current":current_year
    })

    #data hasil kalkulasi beserta persentase perubahan akan dimasukkan ke dalam elemen metric
    st.metric("Curr Ratio", round(income_this_year, 2),delta=curr_change_percent)

    st.line_chart(data=year)

with c12:
    #calculate income revenue average
    income_last_year = account_receivable[13].astype(float)
    income_this_year = account_receivable[24].astype(float)

    # Generate some data
    last_year = list(account_receivable[1:13])
    current_year = list(account_receivable[13:])

    #gunakan rumus untuk mendapatkan persentase pergantian
    curr_change_percent = round(((income_this_year - income_last_year) / income_last_year * 100), 1)
    curr_change_percent = str(curr_change_percent)  + '%' + ' vs. prior year'

    year = pd.DataFrame({
        "last":last_year,
        "current":current_year
    })

    #data hasil kalkulasi beserta persentase perubahan akan dimasukkan ke dalam elemen metric
    st.metric("account receivable", round(income_this_year, 2),delta=curr_change_percent)

    st.line_chart(data=year)

c13, c14, c15, c16 = st.columns(4)
with c13:
    #calculate income revenue average
    income_last_year = account_payable[13].astype(float)
    income_this_year = account_payable[24].astype(float)

    # Generate some data
    last_year = list(account_payable[1:13])
    current_year = list(account_payable[13:])

    #gunakan rumus untuk mendapatkan persentase pergantian
    curr_change_percent = round(((income_this_year - income_last_year) / income_last_year * 100), 1)
    curr_change_percent = str(curr_change_percent)  + '%' + ' vs. prior year'

    year = pd.DataFrame({
        "last":last_year,
        "current":current_year
    })

    #data hasil kalkulasi beserta persentase perubahan akan dimasukkan ke dalam elemen metric
    st.metric("account payable", round(income_this_year, 2),delta=curr_change_percent)

    st.line_chart(data=year)
