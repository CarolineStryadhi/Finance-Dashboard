import streamlit as st
import pandas as pd
import matplotlib.pyplot as plt

# Set layout menjadi wide
st.set_page_config(layout="wide")

# Baca data Excel
excel_file = 'C:/Users/norma/OneDrive/Pictures/folder tugas 2/Finance Dashboard Template.xlsx'  # Ganti 'nama_file.xlsx' dengan nama file Excel yang sesuai
df = pd.read_excel(excel_file)

# Ambil semua kategori
kategori_list = df['Indicator Name'].unique()

# Buat figure dengan subplots
fig, axes = plt.subplots(4, 4, figsize=(25, 20))

# Loop untuk membuat chart untuk setiap kategori
for i, kategori in enumerate(kategori_list):
    data_kategori = df[df['Indicator Name'] == kategori]
    
    # Tentukan posisi subplot
    row = i // 4
    col = i % 4
    
    # Plot data
    axes[row, col].plot(data_kategori.columns[1:], data_kategori.iloc[0, 1:])
    axes[row, col].set_title(f'Line Chart untuk {kategori}')
    axes[row, col].set_xlabel('Bulan')
    axes[row, col].set_ylabel('Nilai')
    axes[row, col].tick_params(axis='x', rotation=45)

# Atur layout
plt.tight_layout()

# Tampilkan plot menggunakan Streamlit
st.pyplot(fig)
