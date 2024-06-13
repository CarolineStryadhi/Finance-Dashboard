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


   
with c2:
    #calculate income revenue average
    cogs_last_year = cost_of_good_sold[13].astype(float)
    cogs_this_year = cost_of_good_sold[24].astype(float)

    # Generate some data
    last_year = list(cost_of_good_sold[1:13])
    current_year = list(cost_of_good_sold[13:])

    #gunakan rumus untuk mendapatkan persentase pergantian
    curr_change_percent = round(((cogs_this_year - cogs_last_year) / cogs_last_year * 100), 1)
    curr_change_percent = str(curr_change_percent)  + '%' + ' vs. prior year'

    year = pd.DataFrame({
        "last":last_year,
        "current":current_year
    })

    #data hasil kalkulasi beserta persentase perubahan akan dimasukkan ke dalam elemen metric
    st.metric("Cost of Good Sold", round(cogs_this_year, 2),delta=curr_change_percent)

    st.line_chart(data=year)




with c3:
    #calculate income revenue average
    gross_last_year = gross_profit[13].astype(float)
    gross_this_year = gross_profit[24].astype(float)

    # Generate some data
    last_year = list(gross_profit[1:13])
    current_year = list(gross_profit[13:])

    #gunakan rumus untuk mendapatkan persentase pergantian
    curr_change_percent = round(((gross_this_year - gross_last_year) / gross_last_year * 100), 1)
    curr_change_percent = str(curr_change_percent)  + '%' + ' vs. prior year'

    year = pd.DataFrame({
        "last":last_year,
        "current":current_year
    })

    #data hasil kalkulasi beserta persentase perubahan akan dimasukkan ke dalam elemen metric
    st.metric("Gross Profit", round(gross_this_year, 2),delta=curr_change_percent)

    st.line_chart(data=year)


with c4:
    #calculate income revenue average
    toe_last_year = Total_operating_Expenses[13].astype(float)
    toe_this_year = Total_operating_Expenses[24].astype(float)

    # Generate some data
    last_year = list(Total_operating_Expenses[1:13])
    current_year = list(Total_operating_Expenses[13:])

    #gunakan rumus untuk mendapatkan persentase pergantian
    curr_change_percent = round(((toe_this_year - toe_last_year) / toe_last_year * 100), 1)
    curr_change_percent = str(curr_change_percent)  + '%' + ' vs. prior year'

    year = pd.DataFrame({
        "last":last_year,
        "current":current_year
    })

    #data hasil kalkulasi beserta persentase perubahan akan dimasukkan ke dalam elemen metric
    st.metric("Total_operating_Expenses", round(toe_this_year, 2),delta=curr_change_percent)

    st.line_chart(data=year)

c5, c6, c7, c8 = st.columns(4)
with c5:
    #calculate income revenue average
    op_last_year = operating_profit[13].astype(float)
    op_this_year = operating_profit[24].astype(float)

    # Generate some data
    last_year = list(operating_profit[1:13])
    current_year = list(operating_profit[13:])

    #gunakan rumus untuk mendapatkan persentase pergantian
    curr_change_percent = round(((op_this_year - op_last_year) / op_last_year * 100), 1)
    curr_change_percent = str(curr_change_percent)  + '%' + ' vs. prior year'

    year = pd.DataFrame({
        "last":last_year,
        "current":current_year
    })

    #data hasil kalkulasi beserta persentase perubahan akan dimasukkan ke dalam elemen metric
    st.metric("Operating Profit", round(op_this_year, 2),delta=curr_change_percent)

    st.line_chart(data=year)

with c6:
    #calculate income revenue average
    taxes_last_year = taxes[13].astype(float)
    taxes_this_year = taxes[24].astype(float)

    # Generate some data
    last_year = list(taxes[1:13])
    current_year = list(taxes[13:])

    #gunakan rumus untuk mendapatkan persentase pergantian
    curr_change_percent = round(((taxes_this_year - taxes_last_year) / taxes_last_year * 100), 1)
    curr_change_percent = str(curr_change_percent)  + '%' + ' vs. prior year'

    year = pd.DataFrame({
        "last":last_year,
        "current":current_year
    })

    #data hasil kalkulasi beserta persentase perubahan akan dimasukkan ke dalam elemen metric
    st.metric("Taxes", round(taxes_this_year, 2),delta=curr_change_percent)

    st.line_chart(data=year)

with c7:
    #calculate income revenue average
    np_last_year = net_profit[13].astype(float)
    np_this_year = net_profit[24].astype(float)

    # Generate some data
    last_year = list(net_profit[1:13])
    current_year = list(net_profit[13:])

    #gunakan rumus untuk mendapatkan persentase pergantian
    curr_change_percent = round(((np_this_year - np_last_year) / np_last_year * 100), 1)
    curr_change_percent = str(curr_change_percent)  + '%' + ' vs. prior year'

    year = pd.DataFrame({
        "last":last_year,
        "current":current_year
    })

    #data hasil kalkulasi beserta persentase perubahan akan dimasukkan ke dalam elemen metric
    st.metric("Net Profit", round(np_this_year, 2),delta=curr_change_percent)

    st.line_chart(data=year)

with c8:
    #calculate income revenue average
    expenses_last_year = expenses[13].astype(float)
    expenses_this_year = expenses[24].astype(float)

    # Generate some data
    last_year = list(expenses[1:13])
    current_year = list(expenses[13:])

    #gunakan rumus untuk mendapatkan persentase pergantian
    curr_change_percent = round(((expenses_this_year - expenses_last_year) / expenses_last_year * 100), 1)
    curr_change_percent = str(curr_change_percent)  + '%' + ' vs. prior year'

    year = pd.DataFrame({
        "last":last_year,
        "current":current_year
    })

    #data hasil kalkulasi beserta persentase perubahan akan dimasukkan ke dalam elemen metric
    st.metric("Expenses", round(expenses_this_year, 2),delta=curr_change_percent)

    st.line_chart(data=year)

c9, c10, c11, c12 = st.columns(4)
with c9:
    #calculate income revenue average
    EOM_last_year = cash_at_eom[13].astype(float)
    EOM_this_year = cash_at_eom[24].astype(float)

    # Generate some data
    last_year = list(cash_at_eom[1:13])
    current_year = list(cash_at_eom[13:])

    #gunakan rumus untuk mendapatkan persentase pergantian
    curr_change_percent = round(((EOM_this_year - EOM_last_year) / EOM_last_year * 100), 1)
    curr_change_percent = str(curr_change_percent)  + '%' + ' vs. prior year'

    year = pd.DataFrame({
        "last":last_year,
        "current":current_year
    })

    #data hasil kalkulasi beserta persentase perubahan akan dimasukkan ke dalam elemen metric
    st.metric("Cash at EOM", round(EOM_this_year, 2),delta=curr_change_percent)

    st.line_chart(data=year)

with c10:
    #calculate income revenue average
    QR_last_year = quick_ratio[13].astype(float)
    QR_this_year = quick_ratio[24].astype(float)

    # Generate some data
    last_year = list(quick_ratio[1:13])
    current_year = list(quick_ratio[13:])

    #gunakan rumus untuk mendapatkan persentase pergantian
    curr_change_percent = round(((QR_this_year - QR_last_year) / QR_last_year * 100), 1)
    curr_change_percent = str(curr_change_percent)  + '%' + ' vs. prior year'

    year = pd.DataFrame({
        "last":last_year,
        "current":current_year
    })

    #data hasil kalkulasi beserta persentase perubahan akan dimasukkan ke dalam elemen metric
    st.metric("Quick Ratio", round(QR_this_year, 2),delta=curr_change_percent)

    st.line_chart(data=year)

with c11:
    #calculate income revenue average
    CR_last_year = curr_ratio[13].astype(float)
    CR_this_year = curr_ratio[24].astype(float)

    # Generate some data
    last_year = list(curr_ratio[1:13])
    current_year = list(curr_ratio[13:])

    #gunakan rumus untuk mendapatkan persentase pergantian
    curr_change_percent = round(((CR_this_year - CR_last_year) / CR_last_year * 100), 1)
    curr_change_percent = str(curr_change_percent)  + '%' + ' vs. prior year'

    year = pd.DataFrame({
        "last":last_year,
        "current":current_year
    })

    #data hasil kalkulasi beserta persentase perubahan akan dimasukkan ke dalam elemen metric
    st.metric("Curr Ratio", round(CR_this_year, 2),delta=curr_change_percent)

    st.line_chart(data=year)

with c12:
    #calculate income revenue average
    AR_last_year = account_receivable[13].astype(float)
    AR_this_year = account_receivable[24].astype(float)

    # Generate some data
    last_year = list(account_receivable[1:13])
    current_year = list(account_receivable[13:])

    #gunakan rumus untuk mendapatkan persentase pergantian
    curr_change_percent = round(((AR_this_year - AR_last_year) / AR_last_year * 100), 1)
    curr_change_percent = str(curr_change_percent)  + '%' + ' vs. prior year'

    year = pd.DataFrame({
        "last":last_year,
        "current":current_year
    })

    #data hasil kalkulasi beserta persentase perubahan akan dimasukkan ke dalam elemen metric
    st.metric("account receivable", round(AR_this_year, 2),delta=curr_change_percent)

    st.line_chart(data=year)

c13, c14, c15, c16 = st.columns(4)
with c13:
    #calculate income revenue average
    AP_last_year = account_payable[13].astype(float)
    AP_this_year = account_payable[24].astype(float)

    # Generate some data
    last_year = list(account_payable[1:13])
    current_year = list(account_payable[13:])

    #gunakan rumus untuk mendapatkan persentase pergantian
    curr_change_percent = round(((AP_this_year - AP_last_year) / AP_last_year * 100), 1)
    curr_change_percent = str(curr_change_percent)  + '%' + ' vs. prior year'

    year = pd.DataFrame({
        "last":last_year,
        "current":current_year
    })

    #data hasil kalkulasi beserta persentase perubahan akan dimasukkan ke dalam elemen metric
    st.metric("account payable", round(AP_this_year, 2),delta=curr_change_percent)

    st.line_chart(data=year)
