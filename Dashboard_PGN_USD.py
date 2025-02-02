# -*- coding: utf-8 -*-
"""
Dashboard Usage dan Monitoring PGN
"""

import base64
import io
import datetime
import calendar

import pandas as pd
import numpy as np

import plotly.express as px
import plotly.graph_objs as go

# Streamlit Components
import streamlit as st

# -----------------------------------------------------
# 0. Konfigurasi Halaman & CSS (Tampilan Profesional)
# -----------------------------------------------------
st.set_page_config(
    page_title="Dashboard PGN",
    page_icon=":bar_chart:",
    layout="wide"
)

# Gaya CSS untuk tampilan yang lebih profesional
custom_css = """
<style>
/* Global Font Family */
body, button, input, textarea, select, h1, h2, h3, h4, h5, h6, p, div, span {
    font-family: 'Segoe UI', Tahoma, Geneva, Verdana, sans-serif;
}

/* Background Color of the Main App */
[data-testid="stAppViewContainer"] {
    background-color: #ffffff; /* Putih untuk tampilan bersih dan profesional */
}

/* Header Styling */
[data-testid="stHeader"] {
    background-color: #003366; /* Biru korporat gelap */
    padding: 15px 30px;
    box-shadow: 0 4px 6px rgba(0, 0, 0, 0.1);
}

/* Header Text Styling */
[data-testid="stHeader"] * {
    color: #ffffff;
    font-weight: 700;
    text-transform: uppercase;
    letter-spacing: 1px;
}

/* Sidebar Styling */
[data-testid="stSidebar"] {
    background-color: #f0f4f8; /* Biru-abu muda untuk sidebar */
    padding: 20px;
    border-right: 1px solid #e0e0e0;
}

/* Upload Box in Sidebar */
.sidebar .upload-box {
    border: 2px solid #003366; /* Border biru korporat */
    border-radius: 10px;
    padding: 25px;
    text-align: center;
    color: #003366;
    font-weight: 600;
    background-color: #ffffff;
    box-shadow: 0 2px 4px rgba(0,0,0,0.05);
    transition: background-color 0.3s, box-shadow 0.3s;
}

.sidebar .upload-box:hover {
    background-color: #e6f0ff;
    box-shadow: 0 4px 8px rgba(0,0,0,0.15);
}

/* Page Title Styling */
.big-title {
    color: #003366; 
    text-align: center; 
    font-weight: 700; 
    margin-top: 30px;
    margin-bottom: 25px;
    font-size: 2.5rem;
    letter-spacing: 1px;
}

/* Streamlit Button Styling */
button.css-1emrehy.edgvbvh3 {
    background-color: #003366;
    color: #ffffff;
    border: none;
    border-radius: 5px;
    padding: 10px 20px;
    font-weight: 600;
    transition: background-color 0.3s;
}

button.css-1emrehy.edgvbvh3:hover {
    background-color: #00509e;
    color: #ffffff;
}

/* Tables Styling */
.stDataFrame table {
    border-collapse: collapse;
    width: 100%;
    font-size: 0.95rem;
    color: #333333;
}

.stDataFrame th, .stDataFrame td {
    border: 1px solid #dddddd;
    text-align: center;
    padding: 10px;
}

.stDataFrame th {
    background-color: #f2f2f2;
    font-weight: 700;
}

.stDataFrame tr:nth-child(even) {
    background-color: #fafafa;
}

/* Plotly Chart Titles */
.plotly .main-svg .title {
    font-family: 'Segoe UI', Tahoma, Geneva, Verdana, sans-serif;
    font-size: 1.5rem;
    color: #003366;
}

/* Select Box Styling */
.css-1siy2j7 {
    background-color: #ffffff;
    border: 1px solid #003366;
    border-radius: 5px;
    padding: 6px;
    font-size: 1rem;
    color: #003366;
}

/* Enhance Select Box on Hover */
.css-1siy2j7:hover {
    border-color: #00509e;
}

/* Page Content Padding */
.main-content {
    padding: 20px;
}

/* Add Smooth Transitions to Interactive Elements */
.sidebar .upload-box, button.css-1emrehy.edgvbvh3, .css-1siy2j7 {
    transition: all 0.3s ease;
}

/* Tables Text Alignment */
.stDataFrame th, .stDataFrame td {
    text-align: center;
}

/* Select Box Alignment */
.css-1siy2j7 {
    text-align: center;
}

</style>
"""
st.markdown(custom_css, unsafe_allow_html=True)

# -----------------------------------------------------
# 1. Judul Aplikasi
# -----------------------------------------------------
st.markdown("<h1 class='big-title'>Dashboard PGN USD</h1>", unsafe_allow_html=True)
st.write(" ")

# -----------------------------------------------------
# 2. Sidebar dengan Pilihan Dashboard
# -----------------------------------------------------
st.sidebar.markdown("### Pilih Dashboard")
dashboard_choice = st.sidebar.radio(
    "Jenis Dashboard",
    ["Dashboard Usage USD", "Dashboard Monitoring USD"]
)

# -----------------------------------------------------
# 3. Session State Inisialisasi untuk Kedua Dashboard
# -----------------------------------------------------
# Dashboard Usage
if "df_daily_usage" not in st.session_state:
    st.session_state["df_daily_usage"] = None
if "df_long_daily_usage" not in st.session_state:
    st.session_state["df_long_daily_usage"] = None
if "df_annual_usage" not in st.session_state:
    st.session_state["df_annual_usage"] = None
if "df_long_annual_usage" not in st.session_state:
    st.session_state["df_long_annual_usage"] = None

# Dashboard Monitoring
if "df_monitoring_bulanan" not in st.session_state:
    st.session_state["df_monitoring_bulanan"] = None
if "df_monitoring_harian" not in st.session_state:
    st.session_state["df_monitoring_harian"] = None
if "list_pelanggan_monitoring" not in st.session_state:
    st.session_state["list_pelanggan_monitoring"] = []
if "list_tahun_monitoring" not in st.session_state:
    st.session_state["list_tahun_monitoring"] = []

# -----------------------------------------------------
# 4. FUNGSI PARSING FILE EXCEL untuk Dashboard Usage
# -----------------------------------------------------
def parse_excel_usage(contents, filename):
    """
    Fungsi ini menerima 'contents' dari file uploader (binary),
    lalu parse menjadi DataFrame (Daily, Monthly, Annual) yang sudah dimodifikasi.
    Kembalikan dictionary berisi data frame yang siap dipakai di layout.
    """
    decoded = contents.read()  # Membaca semua isi file

    try:
        if 'xlsx' in filename:
            # Baca excel
            xls_file = io.BytesIO(decoded)

            # --------------------------------------
            # DAILY
            # --------------------------------------
            df_daily = pd.read_excel(xls_file, sheet_name="Daily Real", skiprows=2)

            # Format nama kolom: hilangkan spasi, ubah dateTime jadi YYYY-MM-DD
            formatted_columns = []
            for col in df_daily.columns:
                if isinstance(col, str) and '00:00:00' in col:
                    try:
                        formatted_columns.append(pd.to_datetime(col).strftime('%Y-%m-%d'))
                    except:
                        formatted_columns.append(col)
                elif isinstance(col, str):
                    formatted_columns.append(col.strip())
                elif isinstance(col, pd.Timestamp) or isinstance(col, datetime.datetime):
                    formatted_columns.append(col.strftime('%Y-%m-%d'))
                else:
                    formatted_columns.append(col)
            df_daily.columns = formatted_columns

            # Identifikasi kolom tanggal (yang mengandung '-')
            tanggal_cols = [col for col in df_daily.columns if isinstance(col, str) and '-' in col]

            # Transformasi melt (long format)
            df_long_daily = df_daily.melt(
                id_vars=['ID', 'Nama Pelanggan', 'CM', 'Segment', 'Ketentuan', 'Kepmen'],
                value_vars=tanggal_cols,
                var_name="Tanggal",
                value_name="Penggunaan"
            )

            df_long_daily["Tanggal"] = pd.to_datetime(df_long_daily["Tanggal"], errors='coerce')
            df_long_daily['Bulan'] = df_long_daily['Tanggal'].dt.month_name()
            df_long_daily['Kategori'] = df_long_daily['Tanggal'].dt.day_name().apply(
                lambda x: 'Weekend' if x in ['Saturday', 'Sunday'] else 'Weekday'
            )

            # --------------------------------------
            # MONTHLY
            # --------------------------------------
            xls_file.seek(0)
            df_monthly = pd.read_excel(xls_file, sheet_name="Monthly", skiprows=2)

            # Format nama kolom
            formatted_columns = []
            for col in df_monthly.columns:
                if isinstance(col, str) and '00:00:00' in col:
                    try:
                        formatted_columns.append(pd.to_datetime(col).strftime('%Y-%m'))
                    except:
                        formatted_columns.append(col)
                elif isinstance(col, str):
                    formatted_columns.append(col.strip())
                elif isinstance(col, pd.Timestamp) or isinstance(col, datetime.datetime):
                    formatted_columns.append(col.strftime('%Y-%m'))
                else:
                    formatted_columns.append(col)
            df_monthly.columns = formatted_columns

            # Normalisasi kolom 'Sektor'
            df_monthly["Sektor"] = df_monthly["Sektor"].str.strip().str.replace('-', ' ').str.split().str.join(' ')

            # Identifikasi kolom yang mengandung '-' (YYYY-MM)
            tanggal_cols = [col for col in df_monthly.columns if isinstance(col, str) and '-' in col]

            df_long_monthly = df_monthly.melt(
                id_vars=["ID", "Nama Pelanggan", "CM", "Sektor", "Produk Utama", "Kontrak", "Segment",
                         "Ketentuan", "Kepmen", "Harga Gas", "Kmin", "Kmax"],
                value_vars=tanggal_cols,
                var_name="Tanggal",
                value_name="Nilai"
            )
            df_long_monthly["Tanggal"] = pd.to_datetime(df_long_monthly["Tanggal"], format='%Y-%m', errors='coerce')
            df_long_monthly["Tahun"] = df_long_monthly["Tanggal"].dt.year
            df_long_monthly["Bulan"] = df_long_monthly["Tanggal"].dt.month_name()

            # --------------------------------------
            # ANNUAL
            # --------------------------------------
            xls_file.seek(0)
            df_annual = pd.read_excel(xls_file, sheet_name="Annual", skiprows=2)

            # Format kolom (ambil tahunnya saja)
            formatted_columns = []
            for col in df_annual.columns:
                if isinstance(col, str) and '-' in col:
                    formatted_columns.append(pd.to_datetime(col).strftime('%Y'))
                elif isinstance(col, pd.Timestamp) or isinstance(col, datetime.datetime):
                    formatted_columns.append(col.strftime('%Y'))
                else:
                    formatted_columns.append(col.strip())
            df_annual.columns = formatted_columns

            # Identifikasi kolom tahun
            year_columns = [c for c in df_annual.columns if c.isdigit()]

            # Melt ke format long
            df_long_annual = df_annual.melt(
                id_vars=['ID', 'Nama Pelanggan', 'CM', 'Sektor', 'Segment', 'Ketentuan', 'Kepmen',
                         'Harga Gas', 'Kmin', 'Kmax'],
                value_vars=year_columns,
                var_name='Tahun',
                value_name='Nilai'
            )

            # Kembalikan dictionary
            return {
                'df_daily_usage': df_daily,
                'df_long_daily_usage': df_long_daily,
                'df_monthly_usage': df_monthly,
                'df_long_monthly_usage': df_long_monthly,
                'df_annual_usage': df_annual,
                'df_long_annual_usage': df_long_annual
            }
        else:
            return None
    except Exception as e:
        print("Error parse_excel_usage:", e)
        return None
# -----------------------------------------------------
# 5. FUNGSI PARSING FILE EXCEL untuk Dashboard Monitoring
# -----------------------------------------------------
def parse_excel_monitoring(file):
    """
    Fungsi ini menerima file uploader untuk Dashboard Monitoring,
    lalu parse menjadi DataFrame yang siap digunakan (daily & monthly).
    """
    try:
        xls = pd.ExcelFile(file)

        # ----------------------
        # 1) Daily Real
        # ----------------------
        df_daily_real = pd.read_excel(xls, 'Daily Real', skiprows=2)
        df_daily_real.columns = [
            pd.to_datetime(col).strftime('%Y-%m-%d') if isinstance(col, (str, datetime.datetime, pd.Timestamp)) and '00:00:00' in str(col)
            else col.strip()
            for col in df_daily_real.columns
        ]
        tanggal_cols_real = [col for col in df_daily_real.columns if isinstance(col, str) and '-' in col]
        df_long_real = df_daily_real.melt(
            id_vars=['ID', 'Nama Pelanggan', 'CM', 'Segment', 'Ketentuan', 'Kepmen'],
            value_vars=tanggal_cols_real,
            var_name="Tanggal",
            value_name="Real_Daily"
        )
        df_long_real['Tanggal'] = pd.to_datetime(df_long_real['Tanggal'], errors='coerce')
        df_long_real['Tahun'] = df_long_real['Tanggal'].dt.year
        df_long_real['Bulan'] = df_long_real['Tanggal'].dt.strftime('%b-%Y')

        # ----------------------
        # 2) Daily Nominee
        # ----------------------
        df_daily_nominee = pd.read_excel(xls, 'Daily Nominee', skiprows=2)
        df_daily_nominee.columns = [
            pd.to_datetime(col).strftime('%Y-%m-%d') if isinstance(col, (str, datetime.datetime, pd.Timestamp)) and '00:00:00' in str(col)
            else col.strip()
            for col in df_daily_nominee.columns
        ]
        tanggal_cols_nominee = [col for col in df_daily_nominee.columns if isinstance(col, str) and '-' in col]
        df_long_nominee = df_daily_nominee.melt(
            id_vars=['ID', 'Nama Pelanggan', 'CM', 'Segment', 'Ketentuan', 'Kepmen'],
            value_vars=tanggal_cols_nominee,
            var_name="Tanggal",
            value_name="Nominasi"
        )
        df_long_nominee['Tanggal'] = pd.to_datetime(df_long_nominee['Tanggal'], errors='coerce')
        df_long_nominee['Tahun'] = df_long_nominee['Tanggal'].dt.year
        df_long_nominee['Bulan'] = df_long_nominee['Tanggal'].dt.strftime('%b-%Y')

        # ----------------------
        # 3) Monthly
        # ----------------------
        df_monthly = pd.read_excel(xls, 'Monthly', skiprows=2)
        df_monthly.columns = [
            pd.to_datetime(col).strftime('%Y-%m') if isinstance(col, (str, datetime.datetime, pd.Timestamp)) and '00:00:00' in str(col)
            else col.strip()
            for col in df_monthly.columns
        ]
        tanggal_cols_monthly = [col for col in df_monthly.columns if isinstance(col, str) and '-' in col]
        df_long_monthly = df_monthly.melt(
            id_vars=["ID", "Nama Pelanggan", "Sektor", "Produk Utama", "Kontrak",
                     "CM", "Segment", "Ketentuan", "Kepmen", "Harga Gas", "Kmin", "Kmax"],
            value_vars=tanggal_cols_monthly,
            var_name="Tanggal",
            value_name="Real_Monthly"
        )
        df_long_monthly['Tanggal'] = pd.to_datetime(df_long_monthly['Tanggal'], format='%Y-%m', errors='coerce')
        df_long_monthly['Bulan'] = df_long_monthly['Tanggal'].dt.strftime('%b-%Y')
        df_long_monthly['Tahun'] = df_long_monthly['Tanggal'].dt.year

        # ----------------------
        # 4) RKAP
        # ----------------------
        df_rkap = pd.read_excel(xls, 'RKAP', skiprows=2)
        df_rkap.columns = [
            pd.to_datetime(col).strftime('%Y-%m') if isinstance(col, (str, datetime.datetime, pd.Timestamp)) and '00:00:00' in str(col)
            else col.strip()
            for col in df_rkap.columns
        ]
        tanggal_cols_rkap = [col for col in df_rkap.columns if isinstance(col, str) and '-' in col]
        df_long_rkap = df_rkap.melt(
            id_vars=['ID', 'Nama Pelanggan', 'CM', 'Segment', 'Ketentuan', 'Kepmen', 'Harga Gas', 'Kmin', 'Kmax'],
            value_vars=tanggal_cols_rkap,
            var_name="Tanggal",
            value_name="RKAP_Actual"
        )
        df_long_rkap['Tanggal'] = pd.to_datetime(df_long_rkap['Tanggal'], format='%Y-%m', errors='coerce')
        df_long_rkap['Bulan'] = df_long_rkap['Tanggal'].dt.strftime('%b-%Y')
        df_long_rkap['Tahun'] = df_long_rkap['Tanggal'].dt.year

        # ----------------------
        # 5) Gabungkan data daily real & nominee
        # ----------------------
        df_combined_daily = pd.merge(
            df_long_real,
            df_long_nominee,
            on=['ID', 'Nama Pelanggan', 'Tanggal', 'Tahun', 'Bulan', 'CM', 'Segment', 'Ketentuan', 'Kepmen'],
            how='left'
        )

        # ----------------------
        # 6) Gabungkan dengan data monthly
        # ----------------------
        df_long_monthly_for_merge = df_long_monthly[[
            'ID', 'Nama Pelanggan', 'Bulan', 'Tahun',
            'Sektor', 'CM', 'Segment', 'Ketentuan', 'Kepmen', 'Real_Monthly', 'Kmin', 'Kmax'
        ]]

        df_combined = pd.merge(
            df_combined_daily,
            df_long_monthly_for_merge,
            on=['ID', 'Nama Pelanggan', 'Bulan', 'Tahun'],
            how='left',
            suffixes=('_daily', '_monthly')
        )

        # Overwrite kolom CM & Segment pakai data monthly jika data daily kosong
        df_combined['CM'] = np.where(
            (df_combined['CM_daily'].eq(0)) | (df_combined['CM_daily'].eq('')),
            df_combined['CM_monthly'],
            df_combined['CM_daily']
        )
        df_combined['Segment'] = np.where(
            (df_combined['Segment_daily'].eq(0)) | (df_combined['Segment_daily'].eq('')),
            df_combined['Segment_monthly'],
            df_combined['Segment_daily']
        )
        df_combined['Ketentuan'] = np.where(
            (df_combined['Ketentuan_daily'].eq(0)) | (df_combined['Ketentuan_daily'].eq('')),
            df_combined['Ketentuan_monthly'],
            df_combined['Ketentuan_daily']
        )
        df_combined['Kepmen'] = np.where(
            (df_combined['Kepmen_daily'].eq(0)) | (df_combined['Kepmen_daily'].eq('')),
            df_combined['Kepmen_monthly'],
            df_combined['Kepmen_daily']
        )

        df_combined.drop(columns=['CM_daily', 'CM_monthly', 'Segment_daily', 'Segment_monthly', 'Ketentuan_daily', 'Ketentuan_monthly', 'Kepmen_daily', 'Kepmen_monthly'], inplace=True, errors='ignore')

        # ----------------------
        # 7) Gabungkan dengan data RKAP
        # ----------------------
        df_combined = pd.merge(
            df_combined,
            df_long_rkap[['ID', 'Nama Pelanggan', 'Bulan', 'Tahun', 'RKAP_Actual']],
            on=['ID', 'Nama Pelanggan', 'Bulan', 'Tahun'],
            how='left'
        )

        # ----------------------
        # 8) Isi nilai NaN
        # ----------------------
        df_combined.fillna(0, inplace=True)

        # ----------------------
        # 9) Hitung RR%
        # ----------------------
        df_combined['RR%'] = np.where(df_combined['RKAP_Actual'] == 0, 0,
                                    (df_combined['Real_Monthly'] / df_combined['RKAP_Actual']) * 100)
        df_combined['RR%'] = df_combined['RR%'].round(2)

        # ----------------------
        # 10) Persiapkan tabel bulanan & harian
        # ----------------------
        df_table_bulanan = df_combined.groupby(['Nama Pelanggan', 'Bulan', 'Tahun', 'CM', 'Segment', 'Sektor', 'Ketentuan', 'Kepmen']).agg({
            'Kmin': 'mean',
            'Kmax': 'mean',
            'Nominasi': 'sum',
            'Real_Monthly': 'first',
            'RKAP_Actual': 'first',
            'RR%': 'mean'
        }).reset_index()
        
        df_table_bulanan[['Kmin', 'Kmax', 'Nominasi', 'Real_Monthly', 'RKAP_Actual', 'RR%']] = \
            df_table_bulanan[['Kmin', 'Kmax', 'Nominasi', 'Real_Monthly', 'RKAP_Actual', 'RR%']].round(2)

        df_table_harian = df_combined[[
            'Nama Pelanggan', 'Tanggal', 'Bulan', 'Tahun', 'CM', 'Segment', 'Sektor', 'Ketentuan', 'Kepmen',
            'Kmin', 'Kmax', 'Nominasi', 'Real_Daily', 'RKAP_Actual', 'RR%'
        ]]

        list_pelanggan = sorted(df_combined['Nama Pelanggan'].unique())
        list_tahun = sorted(df_combined['Tahun'].unique())

        return df_table_bulanan, df_table_harian, list_pelanggan, list_tahun

    except Exception as e:
        print("Error parse_excel_monitoring:", e)
        return None, None, [], []

# -----------------------------------------------------
# 6. FUNGSI PLOTTING untuk Dashboard Usage
# -----------------------------------------------------
month_order = [
    "January", "February", "March", "April", "May", "June",
    "July", "August", "September", "October", "November", "December"
]

def plot_customer_usage(filtered_df, selected_customer):
    """Plot Weekend vs Weekday usage."""
    fig = go.Figure()
    for category in ['Weekday', 'Weekend']:
        cat_df = filtered_df[filtered_df['Kategori'] == category]
        color = 'blue' if category == 'Weekend' else 'green'
        if selected_customer == 'All':
            # Jika 'All' dipilih, agregasikan penggunaan
            cat_usage = cat_df.groupby('Tanggal')['Penggunaan'].sum().reset_index()
            fig.add_trace(go.Scatter(
                x=cat_usage['Tanggal'],
                y=cat_usage['Penggunaan'],
                mode='lines',
                name=category,
                line=dict(color=color)
            ))
        else:
            fig.add_trace(go.Scatter(
                x=cat_df['Tanggal'],
                y=cat_df['Penggunaan'],
                mode='lines',
                name=category,
                line=dict(color=color)
            ))
    fig.update_layout(
        title=f'Penggunaan Harian - {selected_customer}',
        xaxis_title='Tanggal',
        yaxis_title='Penggunaan (MMBtu)',
        template='plotly_white',
        height=600
    )
    return fig

# -----------------------------------------------------
# 7. FUNGSI PLOTTING Preview untuk Overview
# -----------------------------------------------------
def plot_daily_usage_preview():
    """Plot pratinjau penggunaan harian."""
    if st.session_state["df_long_daily_usage"] is not None:
        df_melted = st.session_state["df_long_daily_usage"].copy()
        df_melted['Tanggal'] = pd.to_datetime(df_melted['Tanggal'], errors='coerce')
        filtered_df = df_melted.copy()

        # Agregasikan data jika 'All' pelanggan
        cat_usage = filtered_df.groupby(['Tanggal', 'Kategori'])['Penggunaan'].sum().reset_index()

        fig = go.Figure()
        for category, color in zip(['Weekday', 'Weekend'], ['green', 'blue']):
            cat_df = cat_usage[cat_usage['Kategori'] == category]
            fig.add_trace(go.Scatter(
                x=cat_df['Tanggal'],
                y=cat_df['Penggunaan'],
                mode='lines+markers',
                name=category,
                line=dict(width=3, color=color),
                marker=dict(size=6, color=color),
                hovertemplate='<b>Tanggal:</b> %{x}<br><b>Penggunaan:</b> %{y} MMBtu<extra></extra>'
            ))
        fig.update_layout(
            title="Penggunaan Harian",
            xaxis_title="Tanggal",
            yaxis_title="Penggunaan (MMBtu)",
            template="plotly_white",
            height=300
        )
        st.plotly_chart(fig, use_container_width=True)

def plot_weekend_weekday_preview():
    """Plot pratinjau penggunaan Weekend vs Weekday."""
    if st.session_state["df_long_daily_usage"] is not None:
        df_melted = st.session_state["df_long_daily_usage"].copy()
        df_melted['Tanggal'] = pd.to_datetime(df_melted['Tanggal'], errors='coerce')
        filtered_df = df_melted.copy()

        fig = go.Figure()
        for category, color in zip(['Weekday', 'Weekend'], ['green', 'blue']):
            cat_df = filtered_df[filtered_df['Kategori'] == category]
            cat_usage = cat_df.groupby('Tanggal')['Penggunaan'].sum().reset_index()
            fig.add_trace(go.Scatter(
                x=cat_usage['Tanggal'],
                y=cat_usage['Penggunaan'],
                mode='lines+markers',
                name=category,
                line=dict(color=color),
                marker=dict(size=6, color=color),
                hovertemplate='<b>Tanggal:</b> %{x}<br><b>Penggunaan:</b> %{y} MMBtu<extra></extra>'
            ))
        fig.update_layout(
            title="Weekend vs Weekday",
            xaxis_title="Tanggal",
            yaxis_title="Penggunaan (MMBtu)",
            template="plotly_white",
            height=300
        )
        st.plotly_chart(fig, use_container_width=True)

def plot_annual_preview():
    """Plot pratinjau Annual."""
    if st.session_state["df_long_annual_usage"] is not None:
        df_long_annual = st.session_state["df_long_annual_usage"].copy()
        yearly_stats = df_long_annual.groupby('Tahun').agg(
            Total=('Nilai', 'sum'),
            Average=('Nilai', 'mean'),
            Minimum=('Nilai', 'min'),
            Maximum=('Nilai', 'max')
        ).reset_index()

        fig = px.line(
            yearly_stats,
            x='Tahun',
            y='Total',
            title="Annual",
            labels={'Total': 'Total Konsumsi (MMBtu)', 'Tahun': 'Tahun'},
            template='plotly_white',
            height=300
        )
        fig.update_traces(mode='lines+markers')
        st.plotly_chart(fig, use_container_width=True)

def plot_segment_preview():
    """Plot pratinjau Segment."""
    if st.session_state["df_long_monthly_usage"] is not None:
        df_long_pie = st.session_state["df_long_monthly_usage"].copy()
        segment_stats = df_long_pie.groupby('Segment')['Nilai'].sum().reset_index()

        fig = px.pie(
            segment_stats,
            names='Segment',
            values='Nilai',
            title="Segment",
            hole=0.4,
            template='plotly_white',
            height=300
        )
        fig.update_traces(
            textposition='inside',
            textinfo='percent+label',
            hovertemplate='<b>Segment:</b> %{label}<br>'
                          '<b>Total Nilai:</b> %{value:,.2f}<extra></extra>'
        )
        st.plotly_chart(fig, use_container_width=True)

def plot_sector_preview():
    """Plot pratinjau Sektor."""
    if st.session_state["df_long_monthly_usage"] is not None:
        df_sector = st.session_state["df_long_monthly_usage"].copy()
        sector_stats = df_sector.groupby('Sektor')['Nilai'].sum().reset_index()

        fig = px.pie(
            sector_stats,
            names='Sektor',
            values='Nilai',
            title="Sektor",
            hole=0.4,
            template='plotly_white',
            height=300
        )
        fig.update_traces(
            textposition='inside',
            textinfo='percent+label',
            hovertemplate='<b>Sektor:</b> %{label}<br>'
                          '<b>Total Nilai:</b> %{value:,.2f}<extra></extra>'
        )
        st.plotly_chart(fig, use_container_width=True)

def plot_sector_compare_preview():
    """
    Plot pratinjau (preview) 'Sektor Compare' yang sederhana.
    Misal: Perbandingan total konsumsi antar dua tahun (min dan max) 
    untuk masing-masing Sektor, lalu tampilkan dalam barchart side-by-side.
    """
    if st.session_state["df_long_monthly_usage"] is not None:
        df_sector_compare = st.session_state["df_long_monthly_usage"].copy()

        # Pastikan kolom 'Tahun' numeric
        df_sector_compare['Tahun'] = pd.to_numeric(df_sector_compare['Tahun'], errors='coerce')

        # Jika data kosong (misal tidak ada kolom Tahun)
        if df_sector_compare.empty or df_sector_compare['Tahun'].dropna().empty:
            st.info("Data kosong atau kolom Tahun tidak tersedia.")
            return

        # Contoh: ambil tahun terendah (min) dan tertinggi (max) untuk dibandingkan
        min_year = int(df_sector_compare['Tahun'].min())
        max_year = int(df_sector_compare['Tahun'].max())

        # Filter hanya min_year & max_year
        df_filtered = df_sector_compare[df_sector_compare['Tahun'].isin([min_year, max_year])].copy()

        # Agregasi total per Sektor per Tahun
        sector_compare_data = df_filtered.groupby(['Tahun', 'Sektor'], as_index=False)['Nilai'].sum()

        # Jika ingin urutkan Sektor secara alfabetis (opsional)
        sector_compare_data['Sektor'] = sector_compare_data['Sektor'].astype(str)
        sector_compare_data.sort_values('Sektor', inplace=True)

        # Plot barchart side-by-side (barmode='group')
        fig = px.bar(
            sector_compare_data,
            x='Sektor',
            y='Nilai',
            color='Tahun',
            barmode='group',
            title=f"Sektor Compare (Tahun {min_year} vs {max_year})",
            labels={'Nilai': 'Total Nilai (MMBtu)', 'Sektor': 'Sektor'},
            template='plotly_white',
            height=300
        )
        fig.update_traces(
            texttemplate='%{y:,.2f}',
            textposition='outside',
            hovertemplate=(
                "<b>Sektor:</b> %{x}<br>"
                "<b>Tahun:</b> %{marker.color}<br>"
                "<b>Total Nilai:</b> %{y:,.2f} MMBtu<extra></extra>"
            )
        )
        fig.update_layout(
            xaxis_title="Sektor",
            yaxis_title="Total Nilai (MMBtu)",
            xaxis_tickangle=-45,   # Jika nama sektornya panjang, miringkan sumbu-X
            margin=dict(l=40, r=40, t=80, b=80)
        )

        st.plotly_chart(fig, use_container_width=True)
    else:
        st.info("Belum ada data untuk 'SMonthly'.")

# -----------------------------------------------------
# 7. Dashboard Usage
# -----------------------------------------------------
if dashboard_choice == "Dashboard Usage USD":
    # Upload File Excel di Sidebar
    st.sidebar.markdown("### Upload File Excel untuk Dashboard Usage")
    uploaded_file_usage = st.sidebar.file_uploader(
        label="",
        type=["xlsx"],
        help="Upload file Excel (.xlsx) yang sudah sesuai format template.",
        key="usage_uploader"
    )

    # Proses Upload File Usage
    if uploaded_file_usage is not None:
        parsed_result_usage = parse_excel_usage(uploaded_file_usage, uploaded_file_usage.name)
        if parsed_result_usage is not None:
            st.session_state["df_daily_usage"] = parsed_result_usage["df_daily_usage"]
            st.session_state["df_long_daily_usage"] = parsed_result_usage["df_long_daily_usage"]
            st.session_state["df_monthly_usage"] = parsed_result_usage["df_monthly_usage"]
            st.session_state["df_long_monthly_usage"] = parsed_result_usage["df_long_monthly_usage"]
            st.session_state["df_annual_usage"] = parsed_result_usage["df_annual_usage"]
            st.session_state["df_long_annual_usage"] = parsed_result_usage["df_long_annual_usage"]

            st.sidebar.success(f"File {uploaded_file_usage.name} berhasil di-upload dan diproses!")
        else:
            st.sidebar.error("File tidak valid atau gagal diproses.")
    else:
        st.sidebar.warning("Belum ada file yang diupload untuk Dashboard Usage.")

    # Jika data sudah diupload, tampilkan tabs
    if (
        st.session_state["df_long_daily_usage"] is not None and
        st.session_state["df_long_annual_usage"] is not None
    ):
        tabs_usage = st.tabs([
            "Overview",
            "Daily Usage",
            "Weekend vs Weekday",
            "Annual",
            "Segment",
            "Sektor",
            "Monthly"
        ])


        # -------------------------------------------------
        # 1. TAB: OVERVIEW
        # -------------------------------------------------
        with tabs_usage[0]:
            st.subheader("Overview Dashboard")
            
            # Membuat dua baris dengan beberapa kolom untuk visualisasi
            # Baris 1: Penggunaan Harian, Weekend vs Weekday, Penggunaan Bulanan, Compare
            col1, col2, col3 = st.columns(3)
            with col1:
                st.markdown("### Penggunaan Harian")
                plot_daily_usage_preview()
            with col2:
                st.markdown("### Weekend vs Weekday")
                plot_weekend_weekday_preview()
            with col3:
                st.markdown("### Annual")
                plot_annual_preview()
            st.markdown("---")
            
                # Baris 2 diperbesar dari 3 kolom menjadi 4 kolom
            col4, col5, col6 = st.columns(3)
            with col4:
                st.markdown("### Segment")
                plot_segment_preview()
            with col5:
                st.markdown("### Sektor")
                plot_sector_preview()
            with col6:
                st.markdown("### Monthly")
                plot_sector_compare_preview()  

            # Tidak perlu scroll banyak karena visualisasi diatur dalam baris dan kolom

# -------------------------------------------------
        # 2. TAB: DAILY USAGE
        # -------------------------------------------------
        with tabs_usage[1]:
            st.subheader("Daily Usage Comparison")

            # Add a toggle switch for comparison mode
            comparison_mode = st.toggle("Enable Comparison Mode", key="daily_usage_comparison")

            # Function to create filter and analysis for a panel
            def create_usage_panel(panel_key, df_original):
                # -- Copy df for manipulation --
                df_melted = df_original.copy()
                df_melted['Tanggal'] = pd.to_datetime(df_melted['Tanggal'], errors='coerce')

                # Filter untuk subset data sementara berdasarkan filter sebelumnya
                filtered_df = df_melted.copy()

                # Set USD as the default unit
                selected_unit = "USD"

                # ======================================
                # BARIS 1: 5 DROPDOWN (Nama Pelanggan, CM, Segment, Ketentuan, Kepmen)
                # ======================================
                col1, col2, col3, col4, col5 = st.columns(5)

                with col2:
                    # CM
                    all_cm = sorted(filtered_df['CM'].dropna().unique())
                    cm_options = ["All", "COR", "NON COR"] + [x for x in all_cm if x not in ["COR", "NON COR"]]
                    selected_cm = st.selectbox(f"Pilih CM {panel_key}", cm_options, index=0, key=f"cm_{panel_key}")

                    # Filter berdasarkan CM
                    if selected_cm == "NON COR":
                        filtered_df = filtered_df[filtered_df['CM'] != "COR"]
                    elif selected_cm != "All":
                        filtered_df = filtered_df[filtered_df['CM'] == selected_cm]

                with col1:
                    # Nama Pelanggan
                    all_pelanggan = ['All'] + sorted(filtered_df['Nama Pelanggan'].dropna().unique())
                    selected_customer = st.selectbox(f"Pilih Nama Pelanggan {panel_key}", all_pelanggan, index=0, key=f"customer_{panel_key}")

                    # Filter berdasarkan Nama Pelanggan
                    if selected_customer != 'All':
                        filtered_df = filtered_df[filtered_df['Nama Pelanggan'] == selected_customer]

                with col3:
                    # Segment
                    segment_options = ["All"] + sorted(filtered_df['Segment'].dropna().unique())
                    selected_segment = st.selectbox(f"Pilih Segment {panel_key}", segment_options, index=0, key=f"segment_{panel_key}")

                    # Filter berdasarkan Segment
                    if selected_segment != "All":
                        filtered_df = filtered_df[filtered_df['Segment'] == selected_segment]

                with col4:
                    # Ketentuan
                    ketentuan_options = ["All"] + sorted(filtered_df['Ketentuan'].dropna().unique())
                    selected_ketentuan = st.selectbox(f"Pilih Ketentuan {panel_key}", ketentuan_options, index=0, key=f"ketentuan_{panel_key}")

                    # Filter berdasarkan Ketentuan
                    if selected_ketentuan != "All":
                        filtered_df = filtered_df[filtered_df['Ketentuan'] == selected_ketentuan]

                with col5:
                    # Kepmen
                    kepmen_options = ["All"] + sorted(filtered_df['Kepmen'].dropna().unique())
                    selected_kepmen = st.selectbox(f"Pilih Kepmen {panel_key}", kepmen_options, index=0, key=f"kepmen_{panel_key}")

                    # Filter berdasarkan Kepmen
                    if selected_kepmen != "All":
                        filtered_df = filtered_df[filtered_df['Kepmen'] == selected_kepmen]

                # ======================================
                # BARIS 2: TANGGAL AWAL, TANGGAL AKHIR
                # ======================================
                col6, col7 = st.columns(2)
                with col6:
                    min_date = filtered_df['Tanggal'].min()
                    max_date = filtered_df['Tanggal'].max()
                    start_date = st.date_input(f"Tanggal Awal {panel_key}", value=min_date, 
                                            min_value=min_date, max_value=max_date, key=f"start_date_{panel_key}")
                with col7:
                    end_date = st.date_input(f"Tanggal Akhir {panel_key}", value=max_date, 
                                            min_value=min_date, max_value=max_date, key=f"end_date_{panel_key}")

                # -----------------------------------------------
                # TERAPKAN FILTER
                # -----------------------------------------------
                # Filter Tanggal
                filtered_df = filtered_df[
                    (filtered_df['Tanggal'] >= pd.to_datetime(start_date)) &
                    (filtered_df['Tanggal'] <= pd.to_datetime(end_date))
                ]

                # -----------------------------------------------
                # CEK DATA HASIL FILTER
                # -----------------------------------------------
                if filtered_df.empty:
                    st.info(f"Data tidak tersedia untuk filter tersebut di Panel {panel_key}.")
                    return None

                # Plot
                fig = go.Figure()
                # Jika 'All' -> break down Weekday vs Weekend
                if selected_customer == 'All':
                    weekday_df = filtered_df[filtered_df['Kategori'] == 'Weekday']
                    weekday_usage = weekday_df.groupby('Tanggal')['Penggunaan'].sum().reset_index()
                    fig.add_trace(go.Scatter(
                        x=weekday_usage['Tanggal'], y=weekday_usage['Penggunaan'],
                        mode='lines+markers', name='Weekday'
                    ))

                    weekend_df = filtered_df[filtered_df['Kategori'] == 'Weekend']
                    weekend_usage = weekend_df.groupby('Tanggal')['Penggunaan'].sum().reset_index()
                    fig.add_trace(go.Scatter(
                        x=weekend_usage['Tanggal'], y=weekend_usage['Penggunaan'],
                        mode='lines+markers', name='Weekend'
                    ))
                    fig.update_layout(title=f"Penggunaan Harian - All Pelanggan (Panel {panel_key})")
                else:
                    # Pelanggan tertentu
                    fig.add_trace(go.Scatter(
                        x=filtered_df['Tanggal'], y=filtered_df['Penggunaan'],
                        mode='lines+markers', name=selected_customer
                    ))
                    fig.update_layout(title=f"Penggunaan Harian - {selected_customer} (Panel {panel_key})")

                fig.update_layout(
                    xaxis_title="Tanggal",
                    yaxis_title=f"Penggunaan (USD)",
                    template="plotly_white",
                    height=600
                )
                st.plotly_chart(fig, use_container_width=True)

                # Statistik
                total_val = filtered_df['Penggunaan'].sum()
                avg_val   = filtered_df['Penggunaan'].mean()
                min_val   = filtered_df['Penggunaan'].min()
                max_val   = filtered_df['Penggunaan'].max()

                stats_data = [{
                    'Total': f"{total_val:,.2f} USD",
                    'Average': f"{avg_val:,.2f} USD",
                    'Minimum': f"{min_val:,.2f} USD",
                    'Maximum': f"{max_val:,.2f} USD",
                }]
                st.write(f"#### Statistik Penggunaan (Panel {panel_key})")
                st.table(stats_data)

                # Tabel Data
                df_display = filtered_df.copy()
                df_display['Tanggal'] = df_display['Tanggal'].dt.strftime('%d-%m-%Y')
                df_display['Penggunaan'] = df_display['Penggunaan'].apply(lambda x: f"{x:,.2f}")

                # Replace the table section with this code:
                st.write(f"#### Data Penggunaan Harian - Filtered (Panel {panel_key})")

                # Add table filter dropdown
                table_filter = st.selectbox(
                    f"Filter Table View {panel_key}",
                    options=[
                        'All',
                        'All Highest', 'All Lowest',
                        'Top 5', 'Top 10', 'Top 25', 'Top 50', 'Top 100',
                        'Low 5', 'Low 10', 'Low 25', 'Low 50', 'Low 100'
                    ],
                    key=f"table_filter_{panel_key}"
                )

                # Prepare display dataframe
                df_display = filtered_df.copy()
                df_display['Tanggal'] = df_display['Tanggal'].dt.strftime('%d-%m-%Y')

                # Apply table filters
                if table_filter != 'All':
                    df_sorted = df_display.copy()
                    df_sorted['Penggunaan'] = pd.to_numeric(df_sorted['Penggunaan'], errors='coerce')
                    
                    if 'Highest' in table_filter:
                        df_display = df_sorted.sort_values('Penggunaan', ascending=False)
                    elif 'Lowest' in table_filter:
                        df_display = df_sorted.sort_values('Penggunaan', ascending=True)
                    elif 'Top' in table_filter:
                        n = int(table_filter.split()[1])
                        df_display = df_sorted.sort_values('Penggunaan', ascending=False).head(n)
                    elif 'Low' in table_filter:
                        n = int(table_filter.split()[1])
                        df_display = df_sorted.sort_values('Penggunaan', ascending=True).head(n)

                # Format penggunaan as string with thousand separator and 2 decimal places
                df_display['Penggunaan'] = df_display['Penggunaan'].apply(lambda x: f"{x:,.2f}")

                # Display the filtered table
                st.dataframe(df_display, use_container_width=True)

                # Add Excel export functionality
                if not df_display.empty:
                    excel_buffer = io.BytesIO()
                    with pd.ExcelWriter(excel_buffer, engine='openpyxl') as writer:
                        df_display.to_excel(writer, index=False, sheet_name='Data Penggunaan Harian')
                    excel_buffer.seek(0)
                    
                    filename = f"Data_Penggunaan_Harian_{selected_unit}_{table_filter}_Panel_{panel_key}.xlsx"
                    
                    st.download_button(
                        label=f"Download Data Penggunaan Harian (Excel) - {selected_unit}",
                        data=excel_buffer,
                        file_name=filename,
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                        key=f"download_daily_usage_{panel_key}"
                    )

                return filtered_df

            # Jika mode perbandingan aktif, buat dua panel
            if comparison_mode:
                col_left, col_right = st.columns(2)
                
                with col_left:
                    st.write("### Panel Kiri")
                    left_df = create_usage_panel('Left', st.session_state["df_long_daily_usage"])
                
                with col_right:
                    st.write("### Panel Kanan")
                    right_df = create_usage_panel('Right', st.session_state["df_long_daily_usage"])
            else:
                # Mode normal, satu panel
                create_usage_panel('Single', st.session_state["df_long_daily_usage"])

# -------------------------------------------------
        # 3. TAB: WEEKEND vs WEEKDAY
        # -------------------------------------------------
        with tabs_usage[2]:
            st.subheader("Weekend vs Weekday Comparison")

            # Add a toggle switch for comparison mode
            comparison_mode = st.toggle("Enable Comparison Mode", key="weekend_comparison")

            # Function to create weekend vs weekday panel
            def create_weekend_weekday_panel(panel_key, df_original):
                df_melted = df_original.copy()
                df_melted['Tanggal'] = pd.to_datetime(df_melted['Tanggal'], errors='coerce')

                # Filter untuk subset data sementara berdasarkan filter sebelumnya
                filtered_df = df_melted.copy()

                # Set USD as the default unit
                selected_unit = "USD"

                # ======================================
                # BARIS 1: 5 DROPDOWN (Nama Pelanggan, CM, Segment, Ketentuan, Kepmen)
                # ======================================
                col1, col2, col3, col4, col5 = st.columns(5)

                with col2:
                    # CM
                    all_cm = sorted(filtered_df['CM'].dropna().unique())
                    cm_options = ["All", "COR", "NON COR"] + [x for x in all_cm if x not in ["COR", "NON COR"]]
                    selected_cm = st.selectbox(f"Pilih CM {panel_key}", cm_options, index=0, key=f"weekend_cm_{panel_key}")

                    # Filter berdasarkan CM
                    if selected_cm == "NON COR":
                        filtered_df = filtered_df[filtered_df['CM'] != "COR"]
                    elif selected_cm != "All":
                        filtered_df = filtered_df[filtered_df['CM'] == selected_cm]

                with col1:
                    # Nama Pelanggan
                    all_pelanggan = ['All'] + sorted(filtered_df['Nama Pelanggan'].dropna().unique())
                    selected_customer = st.selectbox(f"Pilih Nama Pelanggan {panel_key}", all_pelanggan, index=0, key=f"weekend_customer_{panel_key}")

                    # Filter berdasarkan Nama Pelanggan
                    if selected_customer != 'All':
                        filtered_df = filtered_df[filtered_df['Nama Pelanggan'] == selected_customer]

                with col3:
                    # Segment
                    segment_options = ["All"] + sorted(filtered_df['Segment'].dropna().unique())
                    selected_segment = st.selectbox(f"Pilih Segment {panel_key}", segment_options, index=0, key=f"weekend_segment_{panel_key}")

                    # Filter berdasarkan Segment
                    if selected_segment != "All":
                        filtered_df = filtered_df[filtered_df['Segment'] == selected_segment]

                with col4:
                    # Ketentuan
                    ketentuan_options = ["All"] + sorted(filtered_df['Ketentuan'].dropna().unique())
                    selected_ketentuan = st.selectbox(f"Pilih Ketentuan {panel_key}", ketentuan_options, index=0, key=f"weekend_ketentuan_{panel_key}")

                    # Filter berdasarkan Ketentuan
                    if selected_ketentuan != "All":
                        filtered_df = filtered_df[filtered_df['Ketentuan'] == selected_ketentuan]

                with col5:
                    # Kepmen
                    kepmen_options = ["All"] + sorted(filtered_df['Kepmen'].dropna().unique())
                    selected_kepmen = st.selectbox(f"Pilih Kepmen {panel_key}", kepmen_options, index=0, key=f"weekend_kepmen_{panel_key}")

                    # Filter berdasarkan Kepmen
                    if selected_kepmen != "All":
                        filtered_df = filtered_df[filtered_df['Kepmen'] == selected_kepmen]

                # ======================================
                # BARIS 2: TANGGAL AWAL, TANGGAL AKHIR
                # ======================================
                col6, col7 = st.columns(2)
                with col6:
                    min_date = filtered_df['Tanggal'].min()
                    max_date = filtered_df['Tanggal'].max()
                    start_date = st.date_input(f"Tanggal Awal {panel_key}", value=min_date, 
                                            min_value=min_date, max_value=max_date, 
                                            key=f"weekend_start_{panel_key}")
                with col7:
                    end_date = st.date_input(f"Tanggal Akhir {panel_key}", value=max_date, 
                                            min_value=min_date, max_value=max_date, 
                                            key=f"weekend_end_{panel_key}")

                # -----------------------------------------------
                # TERAPKAN FILTER
                # -----------------------------------------------
                # Filter Tanggal
                filtered_df = filtered_df[
                    (filtered_df['Tanggal'] >= pd.to_datetime(start_date)) &
                    (filtered_df['Tanggal'] <= pd.to_datetime(end_date))
                ]

                # ======================================
                # Tambahan Filter Kategori (Opsional)
                # ======================================
                category_filter = st.selectbox(f"Filter Kategori {panel_key}", ["All", "Weekday", "Weekend"], 
                                            key=f"weekend_category_filter_{panel_key}")
                if category_filter != 'All':
                    filtered_df = filtered_df[filtered_df['Kategori'] == category_filter]

                # -----------------------------------------------
                # CEK DATA
                # -----------------------------------------------
                if filtered_df.empty:
                    st.info(f"Data tidak tersedia untuk filter tersebut di Panel {panel_key}.")
                    return None

                # Contoh Plot
                fig = go.Figure()
                if selected_customer == 'All':
                    weekday_df = filtered_df[filtered_df['Kategori'] == 'Weekday']
                    weekday_usage = weekday_df.groupby('Tanggal')['Penggunaan'].sum().reset_index()
                    fig.add_trace(go.Scatter(
                        x=weekday_usage['Tanggal'], y=weekday_usage['Penggunaan'],
                        mode='lines+markers', name='Weekday'
                    ))

                    weekend_df = filtered_df[filtered_df['Kategori'] == 'Weekend']
                    weekend_usage = weekend_df.groupby('Tanggal')['Penggunaan'].sum().reset_index()
                    fig.add_trace(go.Scatter(
                        x=weekend_usage['Tanggal'], y=weekend_usage['Penggunaan'],
                        mode='lines+markers', name='Weekend'
                    ))
                    fig.update_layout(title=f"Weekend vs Weekday - All Pelanggan (Panel {panel_key})")
                else:
                    # Pelanggan tertentu
                    fig.add_trace(go.Scatter(
                        x=filtered_df['Tanggal'], y=filtered_df['Penggunaan'],
                        mode='lines+markers',
                        name=selected_customer
                    ))
                    fig.update_layout(title=f"Weekend vs Weekday - {selected_customer} (Panel {panel_key})")

                fig.update_layout(
                    xaxis_title="Tanggal",
                    yaxis_title=f"Penggunaan (USD)",
                    template="plotly_white",
                    height=600
                )
                st.plotly_chart(fig, use_container_width=True)

                # Contoh Statistik Weekend vs Weekday
                categories = ['Weekday', 'Weekend']
                stats_data = []
                for cat in categories:
                    cat_df = filtered_df[filtered_df['Kategori'] == cat]
                    if cat_df.empty:
                        stats_data.append({
                            'Kategori': cat,
                            'Total': f'0 USD',
                            'Average': f'0 USD',
                            'Minimum': f'0 USD',
                            'Maximum': f'0 USD'
                        })
                    else:
                        total_cat = cat_df['Penggunaan'].sum()
                        min_cat   = cat_df['Penggunaan'].min()
                        max_cat   = cat_df['Penggunaan'].max()
                        avg_cat   = cat_df['Penggunaan'].mean()

                        stats_data.append({
                            'Kategori': cat,
                            'Total': f"{total_cat:,.2f} USD",
                            'Average': f"{avg_cat:,.2f} USD",
                            'Minimum': f"{min_cat:,.2f} USD",
                            'Maximum': f"{max_cat:,.2f} USD"
                        })

                st.write(f"#### Statistik Weekend vs Weekday (Panel {panel_key})")
                st.table(stats_data)

                # Contoh Tabel
                df_display = filtered_df.copy()
                df_display['Tanggal'] = df_display['Tanggal'].dt.strftime('%d-%m-%Y')
                df_display['Penggunaan'] = df_display['Penggunaan'].apply(lambda x: f"{x:,.2f}")

                # Replace the Weekend vs Weekday table section with this code:
                st.write(f"#### Data Weekend vs Weekday - Filtered (Panel {panel_key})")

                # Add table filter dropdown
                weekday_table_filter = st.selectbox(
                    f"Filter Weekend vs Weekday Table View {panel_key}",
                    options=[
                        'All',
                        'All Highest', 'All Lowest',
                        'Top 5', 'Top 10', 'Top 25', 'Top 50', 'Top 100',
                        'Low 5', 'Low 10', 'Low 25', 'Low 50', 'Low 100'
                    ],
                    key=f"weekday_table_filter_{panel_key}"
                )

                # Prepare display dataframe
                df_display = filtered_df.copy()
                df_display['Tanggal'] = df_display['Tanggal'].dt.strftime('%d-%m-%Y')

                # Apply table filters
                if weekday_table_filter != 'All':
                    df_sorted = df_display.copy()
                    df_sorted['Penggunaan'] = pd.to_numeric(df_sorted['Penggunaan'], errors='coerce')
                    
                    if 'Highest' in weekday_table_filter:
                        df_display = df_sorted.sort_values('Penggunaan', ascending=False)
                    elif 'Lowest' in weekday_table_filter:
                        df_display = df_sorted.sort_values('Penggunaan', ascending=True)
                    elif 'Top' in weekday_table_filter:
                        n = int(weekday_table_filter.split()[1])
                        df_display = df_sorted.sort_values('Penggunaan', ascending=False).head(n)
                    elif 'Low' in weekday_table_filter:
                        n = int(weekday_table_filter.split()[1])
                        df_display = df_sorted.sort_values('Penggunaan', ascending=True).head(n)

                # Format penggunaan as string with thousand separator and 2 decimal places
                df_display['Penggunaan'] = df_display['Penggunaan'].apply(lambda x: f"{x:,.2f}")

                # Display the filtered table
                st.dataframe(df_display, use_container_width=True)

                # Add Excel export functionality
                if not df_display.empty:
                    excel_buffer = io.BytesIO()
                    with pd.ExcelWriter(excel_buffer, engine='openpyxl') as writer:
                        df_display.to_excel(writer, index=False, sheet_name='Data Weekend vs Weekday')
                    excel_buffer.seek(0)
                    
                    filename = f"Data_Weekend_Weekday_{selected_unit}_{weekday_table_filter}_Panel_{panel_key}.xlsx"
                    
                    st.download_button(
                        label=f"Download Data Weekend vs Weekday (Excel) - {selected_unit}",
                        data=excel_buffer,
                        file_name=filename,
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                        key=f"download_weekday_weekend_{panel_key}"
                    )

                return filtered_df
            # Jika mode perbandingan aktif, buat dua panel
            if comparison_mode:
                col_left, col_right = st.columns(2)
                
                with col_left:
                    st.write("### Panel Kiri")
                    left_df = create_weekend_weekday_panel('Left', st.session_state["df_long_daily_usage"])
                
                with col_right:
                    st.write("### Panel Kanan")
                    right_df = create_weekend_weekday_panel('Right', st.session_state["df_long_daily_usage"])
            else:
                # Mode normal, satu panel
                create_weekend_weekday_panel('Single', st.session_state["df_long_daily_usage"])        

# -------------------------------------------------
        # 5. TAB: ANNUAL
        # -------------------------------------------------
        with tabs_usage[3]:
            st.subheader("Annual")
            
            # Initialize session state for comparison mode if not exists
            if 'comparison_mode' not in st.session_state:
                st.session_state.comparison_mode = False
            
            # Button to toggle comparison mode using st.toggle
            st.session_state.comparison_mode = st.toggle(
                "Enable Comparison View", 
                value=st.session_state.get('comparison_mode', False),
                help="Switch between single view and comparison view"
            )
            df_long_annual = st.session_state["df_long_annual_usage"].copy()
            df_long_annual['Tahun'] = pd.to_numeric(df_long_annual['Tahun'], errors='coerce')
            
            def get_days_in_month(year, month):
                """Helper function to get number of days in a month"""
                import calendar
                return calendar.monthrange(int(year), int(month))[1]
            
            # Fungsi untuk menangani setiap panel (Kiri dan Kanan)
            def annual_panel(panel, panel_side):
                """
                Fungsi untuk membuat panel Annual (kiri atau kanan).
                Args:
                    panel (Streamlit column): Objek kolom Streamlit (col_left atau col_right).
                    panel_side (str): 'left' atau 'right' untuk identifikasi panel.
                """
                with panel:
                    st.markdown(f"### Panel {panel_side.capitalize()}")

                    # Set USD as default unit
                    selected_unit = "USD"

                    # ======================================
                    # FILTER: Pilih Pelanggan, CM, Segment, Sektor, Ketentuan, Kepmen, Tahun Range
                    # ======================================
                    # Membuat enam kolom untuk filter
                    filter_col1, filter_col2, filter_col3 = st.columns(3)
                    filter_col4, filter_col5, filter_col6 = st.columns(3)
                    filter_col7, filter_col8 = st.columns(2)

                    # Pilih CM
                    with filter_col2:
                        all_cm = sorted(df_long_annual['CM'].dropna().unique())
                        cm_options = ["All", "COR", "NON COR"] + [x for x in all_cm if x not in ["COR", "NON COR"]]
                        selected_cm = st.selectbox(
                            f"Pilih CM ({panel_side.capitalize()})",
                            cm_options,
                            index=0,
                            key=f"annual_{panel_side}_cm"
                        )

                    # Filter berdasarkan CM untuk filter lainnya
                    if selected_cm == "NON COR":
                        filtered_df_for_other = df_long_annual[df_long_annual['CM'] != "COR"]
                    elif selected_cm != "All":
                        filtered_df_for_other = df_long_annual[df_long_annual['CM'] == selected_cm]
                    else:
                        filtered_df_for_other = df_long_annual

                    # Pilih Pelanggan (Update berdasarkan CM yang dipilih)
                    with filter_col1:
                        if selected_cm == 'All':
                            all_pelanggan = ['All'] + sorted(filtered_df_for_other['Nama Pelanggan'].dropna().unique())
                        else:
                            filtered_pelanggan = filtered_df_for_other['Nama Pelanggan'].dropna().unique()
                            all_pelanggan = ['All'] + sorted(filtered_pelanggan)
                        selected_customer = st.selectbox(
                            f"Pilih Pelanggan ({panel_side.capitalize()})",
                            all_pelanggan,
                            index=0,
                            key=f"annual_{panel_side}_pelanggan"
                        )

                    # Pilih Segment (Update berdasarkan CM)
                    with filter_col3:
                        segment_options = ["All"] + sorted(filtered_df_for_other['Segment'].dropna().unique())
                        selected_segment = st.selectbox(
                            f"Pilih Segment ({panel_side.capitalize()})",
                            segment_options,
                            index=0,
                            key=f"annual_{panel_side}_segment"
                        )

                    # Pilih Sektor (Update berdasarkan CM)
                    with filter_col4:
                        sektor_options = ["All"] + sorted(filtered_df_for_other['Sektor'].dropna().unique())
                        selected_sektor = st.selectbox(
                            f"Pilih Sektor ({panel_side.capitalize()})",
                            sektor_options,
                            index=0,
                            key=f"annual_{panel_side}_sektor"
                        )

                    # Pilih Ketentuan (Update berdasarkan CM)
                    with filter_col5:
                        ketentuan_options = ["All"] + sorted(filtered_df_for_other['Ketentuan'].dropna().unique())
                        selected_ketentuan = st.selectbox(
                            f"Pilih Ketentuan ({panel_side.capitalize()})",
                            ketentuan_options,
                            index=0,
                            key=f"annual_{panel_side}_ketentuan"
                        )

                    # Pilih Kepmen (Update berdasarkan CM)
                    with filter_col6:
                        kepmen_options = ["All"] + sorted(filtered_df_for_other['Kepmen'].dropna().unique())
                        selected_kepmen = st.selectbox(
                            f"Pilih Kepmen ({panel_side.capitalize()})",
                            kepmen_options,
                            index=0,
                            key=f"annual_{panel_side}_kepmen"
                        )

                    # Pilih Tahun Range
                    with filter_col7:
                        min_year = int(df_long_annual['Tahun'].min())
                        max_year = int(df_long_annual['Tahun'].max())
                        start_year = st.number_input(
                            f"Tahun Awal ({panel_side.capitalize()})",
                            min_value=min_year,
                            max_value=max_year,
                            value=min_year,
                            step=1,
                            key=f"annual_{panel_side}_start_year"
                        )

                    with filter_col8:
                        end_year = st.number_input(
                            f"Tahun Akhir ({panel_side.capitalize()})",
                            min_value=start_year,
                            max_value=max_year,
                            value=max_year,
                            step=1,
                            key=f"annual_{panel_side}_end_year"
                        )

                    # ======================================
                    # TERAPKAN FILTER
                    # ======================================
                    filtered_df = df_long_annual[
                        (df_long_annual['Tahun'] >= start_year) &
                        (df_long_annual['Tahun'] <= end_year)
                    ]

                    # Filter Pelanggan
                    if selected_customer != 'All':
                        filtered_df = filtered_df[filtered_df['Nama Pelanggan'] == selected_customer]

                    # Filter CM
                    if selected_cm == "NON COR":
                        filtered_df = filtered_df[filtered_df["CM"] != "COR"]
                    elif selected_cm != "All":
                        filtered_df = filtered_df[filtered_df["CM"] == selected_cm]

                    # Filter Segment
                    if selected_segment != "All":
                        filtered_df = filtered_df[filtered_df["Segment"] == selected_segment]

                    # Filter Sektor
                    if selected_sektor != "All":
                        filtered_df = filtered_df[filtered_df["Sektor"] == selected_sektor]

                    # Filter Ketentuan
                    if selected_ketentuan != "All":
                        filtered_df = filtered_df[filtered_df["Ketentuan"] == selected_ketentuan]

                    # Filter Kepmen
                    if selected_kepmen != "All":
                        filtered_df = filtered_df[filtered_df["Kepmen"] == selected_kepmen]

                    # -----------------------------------------------
                    # CEK DATA HASIL FILTER
                    # -----------------------------------------------
                    if filtered_df.empty:
                        st.info(f"Tidak ada data yang sesuai dengan filter di Panel {panel_side.capitalize()}.")
                    else:
                        # Menghitung Statistik
                        yearly_stats = filtered_df.groupby('Tahun').agg(
                            Total=('Nilai', 'sum'),
                            Average=('Nilai', 'mean'),
                            Minimum=('Nilai', 'min'),
                            Maximum=('Nilai', 'max')
                        ).reset_index()

                        yearly_stats.sort_values('Tahun', inplace=True)
                        yearly_stats['Perubahan (%)'] = yearly_stats['Total'].pct_change().fillna(0) * 100

                        # Membuat Grafik Bar
                        fig = px.bar(
                            yearly_stats,
                            x='Tahun',
                            y='Total',
                            title=f"Total Gas Consumption {'for ' + selected_customer if selected_customer != 'All' else 'Per Year'} ({panel_side.capitalize()})",
                            labels={'Total': f'Total Gas Consumption (USD)', 'Tahun': 'Year'},
                            text='Total',
                            template='plotly_white',
                            height=400
                        )
                        fig.update_traces(texttemplate='%{text:,.2f}', textposition='outside')
                        fig.update_layout(
                            xaxis_title="Year",
                            yaxis_title=f"Total Consumption (USD)",
                            showlegend=False
                        )
                        st.plotly_chart(fig, use_container_width=True)

                        # Menampilkan Insight
                        max_year_val = yearly_stats.iloc[yearly_stats['Total'].idxmax()]['Tahun']
                        min_year_val = yearly_stats.iloc[yearly_stats['Total'].idxmin()]['Tahun']
                        max_total = yearly_stats['Total'].max()
                        min_total = yearly_stats['Total'].min()

                        insight_text = (f"Tahun dengan konsumsi gas tertinggi adalah *{int(max_year_val)}* "
                                    f"dengan total *{max_total:,.2f} USD*. "
                                    f"Tahun dengan konsumsi gas terendah adalah *{int(min_year_val)}* "
                                    f"dengan total *{min_total:,.2f} USD*.")
                        st.write("#### Insight")
                        st.info(insight_text)

                        # Format data table
                        data_table = yearly_stats.copy()
                        data_table['Total'] = data_table['Total'].apply(lambda x: f"{x:,.2f}")
                        data_table['Average'] = data_table['Average'].apply(lambda x: f"{x:,.2f}")
                        data_table['Minimum'] = data_table['Minimum'].apply(lambda x: f"{x:,.2f}")
                        data_table['Maximum'] = data_table['Maximum'].apply(lambda x: f"{x:,.2f}")
                        data_table['Perubahan (%)'] = data_table['Perubahan (%)'].apply(lambda x: f"{x:.2f}%")

                        st.write(f"#### Yearly Gas Consumption Statistics (USD)")
                        st.dataframe(data_table, use_container_width=True)

                        # -----------------------------------------------
                        # TABEL KONTRIBUTOR
                        # -----------------------------------------------
                        st.write("#### Top Contributors")
                        
                        # Prepare contributor data
                        contributor_data = filtered_df.groupby(
                            ['Nama Pelanggan', 'CM', 'Sektor', 'Segment', 'Ketentuan', 'Kepmen']
                        ).agg(
                            Total_Penggunaan=('Nilai', 'sum')
                        ).reset_index()
                        
                        # Add selection for ranking options
                        ranking_options = [
                            "All",
                            "5 Teratas",
                            "5 Terbawah",
                            "10 Teratas",
                            "10 Terbawah",
                            "15 Teratas",
                            "15 Terbawah",
                            "25 Teratas",
                            "25 Terbawah",
                            "50 Teratas",
                            "50 Terbawah",
                            "100 Teratas",
                            "100 Terbawah",
                            "Urut dari Teratas",
                            "Urut dari Terbawah"
                        ]
                        
                        selected_ranking = st.selectbox(
                            f"Tampilkan Kontributor ({panel_side.capitalize()})",
                            options=ranking_options,
                            key=f"contributor_ranking_{panel_side}"
                        )
                        
                        # Sort data based on selection
                        if "Terbawah" in selected_ranking:
                            contributor_data = contributor_data.sort_values('Total_Penggunaan', ascending=True)
                        else:
                            contributor_data = contributor_data.sort_values('Total_Penggunaan', ascending=False)
                        
                        # Filter data based on selection
                        if selected_ranking != "All":
                            if "Urut dari" not in selected_ranking:
                                # Extract number from the selection (e.g., "5" from "5 Teratas")
                                num_display = int(''.join(filter(str.isdigit, selected_ranking)))
                                contributor_data = contributor_data.head(num_display)
                        
                        # Format the Total_Penggunaan column
                        contributor_data['Total_Penggunaan'] = contributor_data['Total_Penggunaan'].apply(
                            lambda x: f"{x:,.2f} USD"
                        )
                        
                        # Add rank column
                        contributor_data.insert(0, 'Rank', range(1, len(contributor_data) + 1))
                        
                        # Display the table with built-in sorting capabilities
                        st.dataframe(
                            contributor_data,
                            use_container_width=True,
                            column_config={
                                "Rank": "Ranking",
                                "Nama Pelanggan": "Customer Name",
                                "Total_Penggunaan": f"Total Usage (USD)"
                            },
                            hide_index=True
                        )
                        # Add Excel download functionality
                        excel_buffer = io.BytesIO()
                        with pd.ExcelWriter(excel_buffer, engine='openpyxl') as writer:
                            contributor_data.to_excel(writer, index=False, sheet_name='Contributors')
                        excel_buffer.seek(0)

                        st.download_button(
                            label=f"Download Contributors Table (Excel) - {selected_unit}",
                            data=excel_buffer,
                            file_name=f"Contributors_Table_{end_year}_{selected_unit}.xlsx",
                            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                            key=f"download_contributors_{panel_side}"
                        )

            # Layout handling based on comparison mode
            if st.session_state.comparison_mode:
                col_left, col_right = st.columns(2)
                annual_panel(col_left, 'left')
                annual_panel(col_right, 'right')
            else:
                # Single panel mode - just show the left panel
                annual_panel(st.container(), 'left')

        # -------------------------------------------------
        # 6. TAB: SEGMENT
        # -------------------------------------------------
        with tabs_usage[4]:
            st.subheader("Segment")
            
            # Add toggle switch for comparison mode
            enable_comparison = st.toggle("Enable Comparison Mode", key="enable_comparison_mode")
            
            if enable_comparison:
                # Create two columns for side-by-side comparison
                left_col, right_col = st.columns(2)
                
                # Function to create segment analysis panel
                def create_segment_panel(panel_key):
                    df_long_pie = st.session_state["df_long_monthly_usage"].copy()
                    df_long_pie['Tahun'] = pd.to_numeric(df_long_pie['Tahun'], errors='coerce')
                    
                    # Select unit (now only USD)
                    selected_unit = "USD"
                    
                    # Pilih Tahun
                    tahun_options = sorted(df_long_pie['Tahun'].dropna().unique())
                    if not tahun_options:
                        st.info("Data kosong.")
                        return
                    
                    selected_year = st.selectbox(
                        "Pilih Tahun", 
                        tahun_options, 
                        index=0, 
                        key=f"pie_year_usage_{panel_key}"
                    )
                    
                    # Pilih CM
                    cm_options_list = sorted(df_long_pie['CM'].dropna().unique())
                    selected_cm = st.selectbox(
                        "Pilih CM", 
                        cm_options_list, 
                        index=0 if len(cm_options_list) > 0 else -1, 
                        key=f"pie_cm_usage_{panel_key}"
                    )
                    
                    # Pilih Bulan
                    pie_month_options = ['All'] + month_order
                    selected_month = st.selectbox(
                        "Pilih Bulan", 
                        pie_month_options, 
                        index=0, 
                        key=f"pie_month_usage_{panel_key}"
                    )
                    
                    # Filter
                    df_seg_filtered = df_long_pie[df_long_pie['Tahun'] == selected_year]
                    df_seg_filtered = df_seg_filtered[df_seg_filtered['CM'] == selected_cm]
                    if selected_month != 'All':
                        df_seg_filtered = df_seg_filtered[df_seg_filtered['Bulan'] == selected_month]
                    
                    if df_seg_filtered.empty:
                        st.info("Tidak ada data.")
                        return
                    
                    # Calculate segment statistics
                    segment_stats = df_seg_filtered.groupby('Segment').agg(
                        Total_Nilai=('Nilai', 'sum'),
                        Jumlah_Pelanggan=('Nama Pelanggan', 'nunique')
                    ).reset_index()
                    
                    # Create pie chart
                    fig = px.pie(
                        segment_stats,
                        names='Segment',
                        values='Total_Nilai',
                        title=f"Distribution of Values by Segment - {selected_cm}, {selected_month}, {selected_year} (USD)",
                        hole=0.4,
                        template='plotly_white',
                        height=600,
                        custom_data=['Jumlah_Pelanggan']
                    )
                    
                    fig.update_traces(
                        textposition='inside',
                        textinfo='percent+label',
                        hovertemplate=(
                            '<b>Segment:</b> %{label}<br>'
                            '<b>Total Nilai (USD):</b> %{value:,.2f}<br>'
                            '<b>Jumlah Pelanggan:</b> %{customdata[0]}<extra></extra>'
                        )
                    )
                    
                    st.plotly_chart(fig, use_container_width=True, key=f"pie_chart_{panel_key}")
                    
                    # Prepare segment statistics table
                    total_konsumsi = segment_stats['Total_Nilai'].sum()
                    segment_stats['Persentase Konsumsi (%)'] = (segment_stats['Total_Nilai']/total_konsumsi)*100
                    segment_stats['Rata-rata Konsumsi'] = (segment_stats['Total_Nilai']/segment_stats['Jumlah_Pelanggan'])
                    
                    segment_stats_display = segment_stats.copy()
                    segment_stats_display.rename(columns={
                        'Total_Nilai': 'Total_Konsumsi',
                    }, inplace=True)
                    
                    # Format the display values
                    segment_stats_display['Total_Konsumsi'] = segment_stats_display['Total_Konsumsi'].apply(lambda x: f"{x:,.2f}")
                    segment_stats_display['Persentase Konsumsi (%)'] = segment_stats_display['Persentase Konsumsi (%)'].apply(lambda x: f"{x:.2f}%")
                    segment_stats_display['Rata-rata Konsumsi'] = segment_stats_display['Rata-rata Konsumsi'].apply(lambda x: f"{x:,.2f}")
                    
                    st.write("#### Statistik Segment")
                    st.dataframe(segment_stats_display, use_container_width=True)
                    
                    # Show highest and lowest segment indicators
                    numeric_konsumsi = segment_stats_display.copy()
                    numeric_konsumsi['Total_Konsumsi'] = numeric_konsumsi['Total_Konsumsi'].replace(',', '', regex=True).astype(float)
                    
                    max_idx = numeric_konsumsi['Total_Konsumsi'].idxmax()
                    min_idx = numeric_konsumsi['Total_Konsumsi'].idxmin()
                    top_segment_data = [
                        {
                            'Indikator': 'Segmen Tertinggi',
                            'Segment': numeric_konsumsi.loc[max_idx, 'Segment'],
                            'Total Konsumsi': f"{numeric_konsumsi.loc[max_idx, 'Total_Konsumsi']:,.2f} USD"
                        },
                        {
                            'Indikator': 'Segmen Terendah',
                            'Segment': numeric_konsumsi.loc[min_idx, 'Segment'],
                            'Total Konsumsi': f"{numeric_konsumsi.loc[min_idx, 'Total_Konsumsi']:,.2f} USD"
                        }
                    ]
                    
                    st.write("#### Indikator Segmen Tertinggi dan Terendah")
                    st.table(top_segment_data)
                    
                    # Customer list by segment
                    segment_filter_list = ['All'] + sorted(df_seg_filtered['Segment'].dropna().unique())
                    selected_segment = st.selectbox(
                        "Pilih Segment", 
                        segment_filter_list, 
                        index=0, 
                        key=f"segment_filter_usage_{panel_key}"
                    )
                    
                    ranking_options = [
                        'All',
                        'All Highest', 'All Lowest',
                        'Top 5', 'Top 10', 'Top 25', 'Top 50', 'Top 100',
                        'Low 5', 'Low 10', 'Low 25', 'Low 50', 'Low 100'
                    ]
                    
                    ranking_filter = st.selectbox(
                        "Pilih Ranking Filter", 
                        ranking_options, 
                        index=0, 
                        key=f"segment_ranking_filter_usage_{panel_key}"
                    )
                    
                    df_segment_filtered = df_seg_filtered.copy()
                    if selected_segment != 'All':
                        df_segment_filtered = df_segment_filtered[df_segment_filtered['Segment'] == selected_segment]
                    
                    if not df_segment_filtered.empty:
                        customer_stats = df_segment_filtered.groupby('Nama Pelanggan')['Nilai'].sum().reset_index()
                        customer_stats.columns = ['Nama Pelanggan', 'Total_Konsumsi']
                        
                        # Apply ranking filters
                        if ranking_filter == 'All Highest':
                            customer_stats = customer_stats.sort_values(by='Total_Konsumsi', ascending=False)
                        elif ranking_filter == 'All Lowest':
                            customer_stats = customer_stats.sort_values(by='Total_Konsumsi', ascending=True)
                        elif ranking_filter.startswith('Top'):
                            n = int(ranking_filter.split(' ')[-1])
                            customer_stats = customer_stats.nlargest(n, 'Total_Konsumsi')
                        elif ranking_filter.startswith('Low'):
                            n = int(ranking_filter.split(' ')[-1])
                            customer_stats = customer_stats.nsmallest(n, 'Total_Konsumsi')
                        
                        customer_stats['Total_Konsumsi'] = customer_stats['Total_Konsumsi'].apply(lambda x: f"{x:,.2f} USD")
                        
                        st.write("#### Daftar Pelanggan Berdasarkan Segmen")
                        st.dataframe(customer_stats, use_container_width=True)
                    else:
                        st.info("Tidak ada data untuk segmen atau filter terpilih.")
                
                # Create left panel
                with left_col:
                    st.markdown("### Panel Kiri")
                    create_segment_panel("left")
                
                # Create right panel
                with right_col:
                    st.markdown("### Panel Kanan")
                    create_segment_panel("right")
            
            else:
                # Original single panel code
                df_long_pie = st.session_state["df_long_monthly_usage"].copy()
                df_long_pie['Tahun'] = pd.to_numeric(df_long_pie['Tahun'], errors='coerce')
                
                # Set unit to USD
                selected_unit = "USD"
                
                # Pilih Tahun
                tahun_options = sorted(df_long_pie['Tahun'].dropna().unique())
                if not tahun_options:
                    st.info("Data kosong.")
                else:
                    selected_year = st.selectbox("Pilih Tahun", tahun_options, index=0, key="pie_year_usage")

                    # Pilih CM
                    cm_options_list = sorted(df_long_pie['CM'].dropna().unique())
                    selected_cm = st.selectbox(
                        "Pilih CM", 
                        cm_options_list, 
                        index=0 if len(cm_options_list) > 0 else -1, 
                        key="pie_cm_usage"
                    )

                    # Pilih Bulan
                    pie_month_options = ['All'] + month_order
                    selected_month = st.selectbox("Pilih Bulan", pie_month_options, index=0, key="pie_month_usage")

                    # Filter
                    df_seg_filtered = df_long_pie[df_long_pie['Tahun'] == selected_year]
                    df_seg_filtered = df_seg_filtered[df_seg_filtered['CM'] == selected_cm]
                    if selected_month != 'All':
                        df_seg_filtered = df_seg_filtered[df_seg_filtered['Bulan'] == selected_month]

                    if df_seg_filtered.empty:
                        st.info("Tidak ada data.")
                    else:
                        segment_stats = df_seg_filtered.groupby('Segment').agg(
                            Total_Nilai=('Nilai', 'sum'),
                            Jumlah_Pelanggan=('Nama Pelanggan', 'nunique')
                        ).reset_index()

                        fig = px.pie(
                            segment_stats,
                            names='Segment',
                            values='Total_Nilai',
                            title=f"Distribution of Values by Segment - {selected_cm}, {selected_month}, {selected_year} (USD)",
                            hole=0.4,
                            template='plotly_white',
                            height=600,
                            custom_data=['Jumlah_Pelanggan']
                        )
                        
                        fig.update_traces(
                            textposition='inside',
                            textinfo='percent+label',
                            hovertemplate=(
                                '<b>Segment:</b> %{label}<br>'
                                '<b>Total Nilai (USD):</b> %{value:,.2f}<br>'
                                '<b>Jumlah Pelanggan:</b> %{customdata[0]}<extra></extra>'
                            )
                        )

                        st.plotly_chart(fig, use_container_width=True)

                        # Tabel segment-stats
                        total_konsumsi = segment_stats['Total_Nilai'].sum()
                        segment_stats['Persentase Konsumsi (%)'] = (segment_stats['Total_Nilai']/total_konsumsi)*100
                        segment_stats['Rata-rata Konsumsi'] = (segment_stats['Total_Nilai']/segment_stats['Jumlah_Pelanggan'])

                        segment_stats_display = segment_stats.copy()
                        segment_stats_display.rename(columns={
                            'Total_Nilai': 'Total_Konsumsi',
                        }, inplace=True)

                        # Format
                        segment_stats_display['Total_Konsumsi'] = segment_stats_display['Total_Konsumsi'].apply(lambda x: f"{x:,.2f}")
                        segment_stats_display['Persentase Konsumsi (%)'] = segment_stats_display['Persentase Konsumsi (%)'].apply(lambda x: f"{x:.2f}%")
                        segment_stats_display['Rata-rata Konsumsi'] = segment_stats_display['Rata-rata Konsumsi'].apply(lambda x: f"{x:,.2f}")

                        st.write("#### Statistik Segment")
                        st.dataframe(segment_stats_display, use_container_width=True)

                        # Indikator Segmen Tertinggi & Terendah
                        numeric_konsumsi = segment_stats_display.copy()
                        numeric_konsumsi['Total_Konsumsi'] = numeric_konsumsi['Total_Konsumsi'].replace(',', '', regex=True).astype(float)

                        max_idx = numeric_konsumsi['Total_Konsumsi'].idxmax()
                        min_idx = numeric_konsumsi['Total_Konsumsi'].idxmin()
                        top_segment_data = [
                            {
                                'Indikator': 'Segmen Tertinggi',
                                'Segment': numeric_konsumsi.loc[max_idx, 'Segment'],
                                'Total Konsumsi': f"{numeric_konsumsi.loc[max_idx, 'Total_Konsumsi']:,.2f} USD"
                            },
                            {
                                'Indikator': 'Segmen Terendah',
                                'Segment': numeric_konsumsi.loc[min_idx, 'Segment'],
                                'Total Konsumsi': f"{numeric_konsumsi.loc[min_idx, 'Total_Konsumsi']:,.2f} USD"
                            }
                        ]
                        
                        st.write("#### Indikator Segmen Tertinggi dan Terendah")
                        st.table(top_segment_data)

                        # Daftar Pelanggan Berdasarkan Segmen
                        segment_filter_list = ['All'] + sorted(df_seg_filtered['Segment'].dropna().unique())
                        selected_segment = st.selectbox(
                            "Pilih Segment", 
                            segment_filter_list, 
                            index=0, 
                            key="segment_filter_usage"
                        )

                        ranking_options = [
                            'All',
                            'All Highest', 'All Lowest',
                            'Top 5', 'Top 10', 'Top 25', 'Top 50', 'Top 100',
                            'Low 5', 'Low 10', 'Low 25', 'Low 50', 'Low 100'
                        ]
                        
                        ranking_filter = st.selectbox(
                            "Pilih Ranking Filter", 
                            ranking_options, 
                            index=0, 
                            key="segment_ranking_filter_usage"
                        )

                        df_segment_filtered = df_seg_filtered.copy()
                        if selected_segment != 'All':
                            df_segment_filtered = df_segment_filtered[df_segment_filtered['Segment'] == selected_segment]

                        if not df_segment_filtered.empty:
                            customer_stats = df_segment_filtered.groupby('Nama Pelanggan')['Nilai'].sum().reset_index()
                            customer_stats.columns = ['Nama Pelanggan', 'Total_Konsumsi']
                            
                            # Apply ranking filters
                            if ranking_filter == 'All Highest':
                                customer_stats = customer_stats.sort_values(by='Total_Konsumsi', ascending=False)
                            elif ranking_filter == 'All Lowest':
                                customer_stats = customer_stats.sort_values(by='Total_Konsumsi', ascending=True)
                            elif ranking_filter.startswith('Top'):
                                n = int(ranking_filter.split(' ')[-1])
                                customer_stats = customer_stats.nlargest(n, 'Total_Konsumsi')
                            elif ranking_filter.startswith('Low'):
                                n = int(ranking_filter.split(' ')[-1])
                                customer_stats = customer_stats.nsmallest(n, 'Total_Konsumsi')
                            
                            customer_stats['Total_Konsumsi'] = customer_stats['Total_Konsumsi'].apply(lambda x: f"{x:,.2f} USD")
                            
                            st.write("#### Daftar Pelanggan Berdasarkan Segmen")
                            st.dataframe(customer_stats, use_container_width=True)
                        else:
                            st.info("Tidak ada data untuk segmen atau filter terpilih.") 

        # ------------------------------------------------- 
        # 7. TAB: SEKTOR
        # -------------------------------------------------                                          
        with tabs_usage[5]:
            st.subheader("Sektor")
            
            # Add toggle switch for comparison mode
            enable_comparison = st.toggle("Enable Comparison Mode", key="enable_comparison_mode_sektor")
            
            if enable_comparison:
                # Create two columns for side-by-side comparison
                left_col, right_col = st.columns(2)
                
                # Function to create sektor analysis panel
                def create_sektor_panel(panel_key):
                    df_sector = st.session_state["df_long_monthly_usage"].copy()
                    df_sector['Tahun'] = pd.to_numeric(df_sector['Tahun'], errors='coerce')

                    if df_sector.empty:
                        st.info("Data kosong.")
                        return
                    
                    # Pilih Tahun
                    tahun_options = sorted(df_sector['Tahun'].dropna().unique())
                    selected_year = st.selectbox(
                        "Pilih Tahun", 
                        tahun_options, 
                        index=0, 
                        key=f"sector_year_usage_{panel_key}"
                    )

                    # Pilih CM
                    cm_options_list = sorted(df_sector['CM'].dropna().unique())
                    selected_cm = st.selectbox(
                        "Pilih CM", 
                        cm_options_list, 
                        index=0 if len(cm_options_list) > 0 else -1, 
                        key=f"sector_cm_usage_{panel_key}"
                    )

                    # Pilih Bulan
                    sector_month_options = ['All'] + month_order
                    selected_month = st.selectbox(
                        "Pilih Bulan", 
                        sector_month_options, 
                        index=0, 
                        key=f"sector_month_usage_{panel_key}"
                    )

                    # Filter
                    filtered_df = df_sector[df_sector['Tahun'] == selected_year]
                    if selected_cm:
                        filtered_df = filtered_df[df_sector['CM'] == selected_cm]
                    if selected_month != 'All':
                        filtered_df = filtered_df[df_sector['Bulan'] == selected_month]

                    if filtered_df.empty:
                        st.info("Tidak ada data.")
                        return

                    # Pie Chart Sektor
                    sector_stats = filtered_df.groupby('Sektor').agg(
                        Total_Nilai=('Nilai', 'sum'),
                        Jumlah_Pelanggan=('Nama Pelanggan', 'nunique')
                    ).reset_index()
                    
                    fig = px.pie(
                        sector_stats,
                        names='Sektor',
                        values='Total_Nilai',
                        title=f"Distribution of Values by Sektor - {selected_cm}, {selected_month}, {selected_year} (USD)",
                        hole=0.4,
                        template='plotly_white',
                        height=600,
                        custom_data=['Jumlah_Pelanggan']
                    )
                    fig.update_traces(
                        textposition='inside',
                        textinfo='percent+label',
                        hovertemplate=(
                            '<b>Sektor:</b> %{label}<br>'
                            '<b>Total Nilai (USD):</b> %{value:,.2f}<br>'
                            '<b>Jumlah Pelanggan:</b> %{customdata[0]}<extra></extra>'
                        )
                    )

                    # Add unique key for plotly chart
                    st.plotly_chart(fig, use_container_width=True, key=f"pie_chart_sector_{panel_key}")

                    # Tabel Statistik Sektor
                    total_konsumsi = sector_stats['Total_Nilai'].sum()
                    sector_stats['Persentase Konsumsi (%)'] = (sector_stats['Total_Nilai']/total_konsumsi)*100
                    sector_stats['Rata-rata Konsumsi'] = (sector_stats['Total_Nilai']/sector_stats['Jumlah_Pelanggan'])

                    sector_stats_display = sector_stats.copy()
                    sector_stats_display.rename(columns={
                        'Total_Nilai': 'Total_Konsumsi_USD',
                    }, inplace=True)

                    # Pastikan 'Rata-rata Konsumsi' selalu ada
                    if 'Rata-rata Konsumsi' not in sector_stats_display.columns:
                        sector_stats_display['Rata-rata Konsumsi'] = 0

                    # Format
                    sector_stats_display['Total_Konsumsi_USD'] = sector_stats_display['Total_Konsumsi_USD'].apply(lambda x: f"{x:,.2f}")
                    sector_stats_display['Persentase Konsumsi (%)'] = sector_stats_display['Persentase Konsumsi (%)'].apply(lambda x: f"{x:.2f}%")
                    sector_stats_display['Rata-rata Konsumsi'] = sector_stats_display['Rata-rata Konsumsi'].apply(lambda x: f"{x:,.2f}")

                    st.write("#### Statistik Sektor")
                    st.dataframe(sector_stats_display, use_container_width=True)

                    # Daftar Pelanggan Berdasarkan Sektor
                    sektor_list = ['All'] + sorted(df_sector['Sektor'].dropna().unique())
                    selected_sektor = st.selectbox(
                        "Pilih Sektor", 
                        sektor_list, 
                        index=0, 
                        key=f"sector_filter_usage_{panel_key}"
                    )

                    # Ranking filter dengan opsi 'All'
                    ranking_options = [
                        'All',
                        'All Highest', 'All Lowest',
                        'Top 5', 'Top 10', 'Top 25', 'Top 50', 'Top 100',
                        'Low 5', 'Low 10', 'Low 25', 'Low 50', 'Low 100'
                    ]
                    sector_ranking_filter = st.selectbox(
                        "Pilih Ranking Filter", 
                        ranking_options, 
                        index=0, 
                        key=f"sector_customer_ranking_usage_{panel_key}"
                    )

                    df_sector_filtered = filtered_df.copy()
                    if selected_sektor != 'All':
                        df_sector_filtered = df_sector_filtered[df_sector_filtered['Sektor'] == selected_sektor]

                    if not df_sector_filtered.empty:
                        df_customers = df_sector_filtered.groupby(['Nama Pelanggan', 'Produk Utama'])['Nilai'].sum().reset_index()
                        df_customers.columns = ['Nama_Pelanggan', 'Produk_Utama', 'Total_Konsumsi']

                        # Ranking filter
                        if sector_ranking_filter == 'All Highest':
                            df_customers = df_customers.sort_values(by='Total_Konsumsi', ascending=False)
                        elif sector_ranking_filter == 'All Lowest':
                            df_customers = df_customers.sort_values(by='Total_Konsumsi', ascending=True)
                        elif sector_ranking_filter.startswith('Top'):
                            n = int(sector_ranking_filter.split(' ')[-1])
                            df_customers = df_customers.nlargest(n, 'Total_Konsumsi')
                        elif sector_ranking_filter.startswith('Low'):
                            n = int(sector_ranking_filter.split(' ')[-1])
                            df_customers = df_customers.nsmallest(n, 'Total_Konsumsi')
                        elif sector_ranking_filter == 'All':
                            pass  # Tampilkan semua tanpa filter

                        df_customers['Total_Konsumsi'] = df_customers['Total_Konsumsi'].apply(lambda x: f"{x:,.2f} USD")

                        st.write("#### Daftar Pelanggan Berdasarkan Sektor")
                        st.dataframe(df_customers, use_container_width=True)
                    else:
                        st.info("Tidak ada data pelanggan untuk sektor/filter terpilih.")
                
                # Create left panel
                with left_col:
                    st.markdown("### Panel Kiri")
                    create_sektor_panel("left")
                
                # Create right panel
                with right_col:
                    st.markdown("### Panel Kanan")
                    create_sektor_panel("right")
                    
            else:
                # Original single panel code
                df_sector = st.session_state["df_long_monthly_usage"].copy()
                df_sector['Tahun'] = pd.to_numeric(df_sector['Tahun'], errors='coerce')

                if df_sector.empty:
                    st.info("Data kosong.")
                else:
                    # Pilih Tahun
                    tahun_options = sorted(df_sector['Tahun'].dropna().unique())
                    selected_year = st.selectbox("Pilih Tahun", tahun_options, index=0, key="sector_year_usage_single")

                    # Pilih CM
                    cm_options_list = sorted(df_sector['CM'].dropna().unique())
                    selected_cm = st.selectbox(
                        "Pilih CM", 
                        cm_options_list, 
                        index=0 if len(cm_options_list) > 0 else -1, 
                        key="sector_cm_usage_single"
                    )

                    # Pilih Bulan
                    sector_month_options = ['All'] + month_order
                    selected_month = st.selectbox(
                        "Pilih Bulan", 
                        sector_month_options, 
                        index=0, 
                        key="sector_month_usage_single"
                    )

                    # Filter
                    filtered_df = df_sector[df_sector['Tahun'] == selected_year]
                    if selected_cm:
                        filtered_df = filtered_df[df_sector['CM'] == selected_cm]
                    if selected_month != 'All':
                        filtered_df = filtered_df[df_sector['Bulan'] == selected_month]

                    if filtered_df.empty:
                        st.info("Tidak ada data.")
                    else:
                        # Pie Chart Sektor
                        sector_stats = filtered_df.groupby('Sektor').agg(
                            Total_Nilai=('Nilai', 'sum'),
                            Jumlah_Pelanggan=('Nama Pelanggan', 'nunique')
                        ).reset_index()
                        
                        fig = px.pie(
                            sector_stats,
                            names='Sektor',
                            values='Total_Nilai',
                            title=f"Distribution of Values by Sektor - {selected_cm}, {selected_month}, {selected_year} (USD)",
                            hole=0.4,
                            template='plotly_white',
                            height=600,
                            custom_data=['Jumlah_Pelanggan']
                        )
                        fig.update_traces(
                            textposition='inside',
                            textinfo='percent+label',
                            hovertemplate=(
                                '<b>Sektor:</b> %{label}<br>'
                                '<b>Total Nilai (USD):</b> %{value:,.2f}<br>'
                                '<b>Jumlah Pelanggan:</b> %{customdata[0]}<extra></extra>'
                            )
                        )

                        st.plotly_chart(fig, use_container_width=True, key="pie_chart_sector_single_panel")

                        # Tabel Statistik Sektor
                        total_konsumsi = sector_stats['Total_Nilai'].sum()
                        sector_stats['Persentase Konsumsi (%)'] = (sector_stats['Total_Nilai']/total_konsumsi)*100
                        sector_stats['Rata-rata Konsumsi'] = (sector_stats['Total_Nilai']/sector_stats['Jumlah_Pelanggan'])

                        sector_stats_display = sector_stats.copy()
                        sector_stats_display.rename(columns={
                            'Total_Nilai': 'Total_Konsumsi_USD',
                        }, inplace=True)

                        # Pastikan 'Rata-rata Konsumsi' selalu ada
                        if 'Rata-rata Konsumsi' not in sector_stats_display.columns:
                            sector_stats_display['Rata-rata Konsumsi'] = 0

                        # Format
                        sector_stats_display['Total_Konsumsi_USD'] = sector_stats_display['Total_Konsumsi_USD'].apply(lambda x: f"{x:,.2f}")
                        sector_stats_display['Persentase Konsumsi (%)'] = sector_stats_display['Persentase Konsumsi (%)'].apply(lambda x: f"{x:.2f}%")
                        sector_stats_display['Rata-rata Konsumsi'] = sector_stats_display['Rata-rata Konsumsi'].apply(lambda x: f"{x:,.2f}")

                        st.write("#### Statistik Sektor")
                        st.dataframe(sector_stats_display, use_container_width=True)

                        # Daftar Pelanggan Berdasarkan Sektor
                        sektor_list = ['All'] + sorted(df_sector['Sektor'].dropna().unique())
                        selected_sektor = st.selectbox(
                            "Pilih Sektor", 
                            sektor_list, 
                            index=0, 
                            key="sector_filter_usage_single"
                        )

                        # Ranking filter dengan opsi 'All'
                        ranking_options = [
                            'All',
                            'All Highest', 'All Lowest',
                            'Top 5', 'Top 10', 'Top 25', 'Top 50', 'Top 100',
                            'Low 5', 'Low 10', 'Low 25', 'Low 50', 'Low 100'
                        ]
                        sector_ranking_filter = st.selectbox(
                            "Pilih Ranking Filter", 
                            ranking_options, 
                            index=0, 
                            key="sector_customer_ranking_usage_single"
                        )

                        df_sector_filtered = filtered_df.copy()
                        if selected_sektor != 'All':
                            df_sector_filtered = df_sector_filtered[df_sector_filtered['Sektor'] == selected_sektor]

                        if not df_sector_filtered.empty:
                            df_customers = df_sector_filtered.groupby(['Nama Pelanggan', 'Produk Utama'])['Nilai'].sum().reset_index()
                            df_customers.columns = ['Nama_Pelanggan', 'Produk_Utama', 'Total_Konsumsi']

                            # Ranking filter
                            if sector_ranking_filter == 'All Highest':
                                df_customers = df_customers.sort_values(by='Total_Konsumsi', ascending=False)
                            elif sector_ranking_filter == 'All Lowest':
                                df_customers = df_customers.sort_values(by='Total_Konsumsi', ascending=True)
                            elif sector_ranking_filter.startswith('Top'):
                                n = int(sector_ranking_filter.split(' ')[-1])
                                df_customers = df_customers.nlargest(n, 'Total_Konsumsi')
                            elif sector_ranking_filter.startswith('Low'):
                                n = int(sector_ranking_filter.split(' ')[-1])
                                df_customers = df_customers.nsmallest(n, 'Total_Konsumsi')
                            elif sector_ranking_filter == 'All':
                                pass  # Tampilkan semua tanpa filter

                            df_customers['Total_Konsumsi'] = df_customers['Total_Konsumsi'].apply(lambda x: f"{x:,.2f} USD")

                            st.write("#### Daftar Pelanggan Berdasarkan Sektor")
                            st.dataframe(df_customers, use_container_width=True)
                        else:
                            st.info("Tidak ada data pelanggan untuk sektor/filter terpilih.")                                         

        # -------------------------------------------------
        # 9. TAB: SEKTOR COMPARE
        # -------------------------------------------------
        with tabs_usage[6]:
            st.subheader("Monthly")

            # Add a switch for dual panel view
            dual_panel_mode = st.toggle("Dual Panel Comparison", key="dual_panel_toggle_unique")

            # Function to create the comparison section
            def create_comparison_section(key_prefix):
                df_sector_compare = st.session_state["df_long_monthly_usage"].copy()
                if df_sector_compare.empty:
                    st.info("Data kosong atau belum diunggah.")
                    return None

                # Set unit label
                selected_unit = "United States Dollar (USD)"

                # Pastikan kolom Tahun numeric
                df_sector_compare['Tahun'] = pd.to_numeric(df_sector_compare['Tahun'], errors='coerce')

                # Filter untuk subset data sementara berdasarkan filter sebelumnya
                filtered_df = df_sector_compare.copy()

                # Create first row of filters using columns
                col1, col2, col3 = st.columns(3)
                
                with col1:
                    list_cm = ["All", "COR", "NON COR"] + sorted(
                        [x for x in df_sector_compare['CM'].dropna().unique() if x not in ["COR", "NON COR"]]
                    )
                    selected_cm = st.selectbox("Pilih CM", list_cm, index=0, key=f"compare_sektor_cm_{key_prefix}")

                with col2:
                    if selected_cm == "NON COR":
                        filtered_df = filtered_df[filtered_df['CM'] != "COR"]
                    elif selected_cm != "All":
                        filtered_df = filtered_df[filtered_df['CM'] == selected_cm]
                    list_pelanggan = ['All'] + sorted(filtered_df['Nama Pelanggan'].dropna().unique())
                    selected_customer = st.selectbox("Pilih Nama Pelanggan", list_pelanggan, index=0, key=f"compare_sektor_pelanggan_{key_prefix}")

                with col3:
                    if selected_customer != 'All':
                        filtered_df = filtered_df[filtered_df['Nama Pelanggan'] == selected_customer]
                    list_sektor = ['All'] + sorted(filtered_df['Sektor'].dropna().unique())
                    selected_sektor = st.selectbox("Pilih Sektor", list_sektor, index=0, key=f"compare_sektor_sektor_{key_prefix}")

                # Create second row of filters using columns
                col4, col5, col6 = st.columns(3)
                
                with col4:
                    if selected_sektor != 'All':
                        filtered_df = filtered_df[filtered_df['Sektor'] == selected_sektor]
                    list_segment = ['All'] + sorted(filtered_df['Segment'].dropna().unique())
                    selected_segment = st.selectbox("Pilih Segment", list_segment, index=0, key=f"compare_sektor_segment_{key_prefix}")

                with col5:
                    if selected_segment != 'All':
                        filtered_df = filtered_df[filtered_df['Segment'] == selected_segment]
                    list_ketentuan = ['All'] + sorted(filtered_df['Ketentuan'].dropna().unique())
                    selected_ketentuan = st.selectbox("Pilih Ketentuan", list_ketentuan, index=0, key=f"compare_sektor_ketentuan_{key_prefix}")

                with col6:
                    if selected_ketentuan != 'All':
                        filtered_df = filtered_df[filtered_df['Ketentuan'] == selected_ketentuan]
                    list_kepmen = ['All'] + sorted(filtered_df['Kepmen'].dropna().unique())
                    selected_kepmen = st.selectbox("Pilih Kepmen", list_kepmen, index=0, key=f"compare_sektor_kepmen_{key_prefix}")

                # Filter based on Kepmen
                if selected_kepmen != 'All':
                    filtered_df = filtered_df[filtered_df['Kepmen'] == selected_kepmen]

                # -- Pilih rentang Tahun dengan slider
                min_year = int(df_sector_compare['Tahun'].min())
                max_year = int(df_sector_compare['Tahun'].max())
                year_range = st.slider(
                    "Filter Tahun",
                    min_year,
                    max_year,
                    (min_year, max_year),
                    step=1,
                    key=f"compare_sektor_year_range_{key_prefix}"
                )

                # Filter data sesuai pilihan tahun
                filtered_df = filtered_df[
                    (filtered_df['Tahun'] >= year_range[0]) & (filtered_df['Tahun'] <= year_range[1])
                ]

                # Urutan bulan
                month_order = [
                    "January", "February", "March", "April", "May", "June",
                    "July", "August", "September", "October", "November", "December"
                ]

                # ----------------------------------------------
                # 1) Tampilkan Bar Chart Perbandingan
                # ----------------------------------------------
                if filtered_df.empty:
                    st.info("Tidak ada data sesuai filter.")
                    return None

                filtered_df['Bulan'] = pd.Categorical(filtered_df['Bulan'], categories=month_order, ordered=True)
                filtered_df.sort_values('Bulan', inplace=True)

                import plotly.graph_objects as go
                fig_sektor_compare = go.Figure()

                # Loop tiap tahun dalam rentang slider
                for yr in range(year_range[0], year_range[1] + 1):
                    df_year = filtered_df[filtered_df['Tahun'] == yr]
                    if not df_year.empty:
                        # Agregasi Nilai per bulan
                        monthly_grouped = df_year.groupby('Bulan')['Nilai'].sum().reset_index()

                        # Inside create_comparison_section function, update the fig_sektor_compare.add_trace() call:
                        fig_sektor_compare.add_trace(go.Bar(
                            x=monthly_grouped['Bulan'],
                            y=monthly_grouped['Nilai'],
                            name=str(yr),
                            hovertemplate=(
                                f"Pelanggan: {'All Pelanggan' if selected_customer == 'All' else selected_customer}<br>" +
                                f"CM: {'All' if selected_cm == 'All' else selected_cm}<br>" +
                                f"Sektor: {'All Sektor' if selected_sektor == 'All' else selected_sektor}<br>" +
                                f"Tahun: {yr}<br>" +
                                "Bulan: %{x}<br>" +  # Fixed syntax
                                "Penggunaan: %{y:.2f} United States Dollar (USD)<extra></extra>"
                            )
                        ))

                fig_sektor_compare.update_layout(
                    title=(
                        f"Perbandingan Penggunaan (Monthly Compare)<br>"
                        f"({'All Pelanggan' if selected_customer == 'All' else selected_customer}) | "
                        f"({'All' if selected_cm == 'All' else selected_cm}) | "
                        f"({'All Sektor' if selected_sektor == 'All' else selected_sektor}) | "
                        f"({'All Segment' if selected_segment == 'All' else selected_segment}) | "
                        f"({'All Ketentuan' if selected_ketentuan == 'All' else selected_ketentuan}) | "
                        f"({'All Kepmen' if selected_kepmen == 'All' else selected_kepmen})"
                    ),
                    xaxis_title="Bulan",
                    yaxis_title=f"Penggunaan ({selected_unit})",
                    barmode="group",
                    template="plotly_white",
                    height=600,
                    legend_title="Tahun",
                    xaxis=dict(categoryorder='array', categoryarray=month_order)
                )
                st.plotly_chart(fig_sektor_compare, use_container_width=True, key=f"plotly_chart_{key_prefix}")

                # ----------------------------------------------
                # 2) Statistik Perbandingan
                # ----------------------------------------------
                df_awal = filtered_df[filtered_df['Tahun'] == year_range[0]]
                df_akhir = filtered_df[filtered_df['Tahun'] == year_range[1]]

                total_awal = df_awal['Nilai'].sum() if not df_awal.empty else 0
                total_akhir = df_akhir['Nilai'].sum() if not df_akhir.empty else 0
                selisih = total_akhir - total_awal
                persen_perubahan = (selisih / total_awal * 100) if total_awal != 0 else 0

                # Hitung CAGR
                if total_awal != 0 and (year_range[1] > year_range[0]):
                    cagr = ((total_akhir / total_awal)**(1 / (year_range[1] - year_range[0])) - 1) * 100
                else:
                    cagr = 0

                # Average Usage
                avg_awal = df_awal['Nilai'].mean() if not df_awal.empty else 0
                avg_akhir = df_akhir['Nilai'].mean() if not df_akhir.empty else 0

                stats_table_compare = [
                    {
                        'Metric': f'Total Usage ({selected_unit})',
                        'Tahun Awal': f"{total_awal:,.2f}",
                        'Tahun Akhir': f"{total_akhir:,.2f}",
                        'Perbandingan (Selisih)': f"{selisih:,.2f}",
                        'Perbandingan (%)': f"{persen_perubahan:.2f}%",
                        'CAGR (%)': f"{cagr:.2f}%"
                    },
                    {
                        'Metric': f'Average Usage ({selected_unit})',
                        'Tahun Awal': f"{avg_awal:,.2f}",
                        'Tahun Akhir': f"{avg_akhir:,.2f}",
                        'Perbandingan (Selisih)': f"{(avg_akhir - avg_awal):,.2f}",
                        'Perbandingan (%)': (
                            f"{((avg_akhir - avg_awal)/avg_awal*100) if avg_awal != 0 else 0:.2f}%"
                        ),
                        'CAGR (%)': 'N/A'
                    }
                ]

                st.write("### Statistik Perbandingan")
                st.table(stats_table_compare)

                # ----------------------------------------------
                # 3) Distribusi Bulanan
                # ----------------------------------------------
                st.write("### Distribusi Bulanan")
                monthly_data = []
                for m in month_order:
                    awal_val = df_awal[df_awal['Bulan'] == m]['Nilai'].sum() if not df_awal.empty else 0
                    akhir_val = df_akhir[df_akhir['Bulan'] == m]['Nilai'].sum() if not df_akhir.empty else 0
                    perubahan_abs = akhir_val - awal_val
                    perubahan_pct = (perubahan_abs / awal_val * 100) if awal_val != 0 else 0
                    monthly_data.append({
                        'Bulan': m,
                        f'Konsumsi Tahun Awal ({selected_unit})': f"{awal_val:,.2f}",
                        f'Konsumsi Tahun Akhir ({selected_unit})': f"{akhir_val:,.2f}",
                        f'Perubahan Absolut ({selected_unit})': f"{perubahan_abs:,.2f}",
                        'Perubahan (%)': f"{perubahan_pct:.2f}%"
                    })
                st.dataframe(monthly_data, use_container_width=True)

                # Add after monthly distribution table (st.dataframe(monthly_data))
                excel_buffer_monthly = io.BytesIO()
                with pd.ExcelWriter(excel_buffer_monthly, engine='openpyxl') as writer:
                    # Convert list of dictionaries to DataFrame first
                    monthly_df = pd.DataFrame(monthly_data)
                    monthly_df.to_excel(writer, index=False, sheet_name='Monthly Distribution')
                excel_buffer_monthly.seek(0)

                st.download_button(
                    label=f"Download Distribusi Bulanan (Excel) - {selected_unit}",
                    data=excel_buffer_monthly,
                    file_name=f"Monthly_Distribution_{year_range[0]}_{year_range[1]}_{selected_unit}.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    key=f"download_monthly_{key_prefix}"
                )

                # ----------------------------------------------
                # 4) Opsional: Kontributor Utama
                # ----------------------------------------------
                st.write("### Kontributor Utama")
                contrib_filter_options = [
                    'All',
                    'Top 5', 'Top 10', 'Top 25', 'Top 50', 'Top 100',
                    'Bottom 5', 'Bottom 10', 'Bottom 25', 'Bottom 50', 'Bottom 100',
                    'Highest Change', 'Lowest Change'
                ]
                selected_contrib = st.selectbox("Pilih Filter Kontributor", contrib_filter_options, index=0, key=f"compare_sektor_contrib_filter_{key_prefix}")

                df_contributor = df_sector_compare[
                    (df_sector_compare['Tahun'] >= year_range[0]) &
                    (df_sector_compare['Tahun'] <= year_range[1])
                ].copy()

                if selected_customer != 'All':
                    df_contributor = df_contributor[df_contributor['Nama Pelanggan'] == selected_customer]
                if selected_cm != 'All':
                    df_contributor = df_contributor[df_contributor['CM'] == selected_cm]
                if selected_sektor != 'All':
                    df_contributor = df_contributor[df_contributor['Sektor'] == selected_sektor]

                if df_contributor.empty:
                    st.info("Tidak ada data kontributor.")
                else:
                    # Group by all relevant columns including the new ones
                    contributor_grouped = df_contributor.groupby([
                        'Nama Pelanggan', 'CM', 'Sektor', 'Segment', 'Ketentuan', 'Kepmen'
                    ])['Nilai'].sum().reset_index()
                    contributor_grouped.columns = ['Nama', 'CM', 'Sektor', 'Segment', 'Ketentuan', 'Kepmen', 'Total Konsumsi']

                    # Calculate yearly values
                    awal_df_ = df_contributor[df_contributor['Tahun'] == year_range[0]].groupby(['Nama Pelanggan', 'Bulan'])['Nilai'].sum()
                    akhir_df_ = df_contributor[df_contributor['Tahun'] == year_range[1]].groupby(['Nama Pelanggan', 'Bulan'])['Nilai'].sum()

                    # Calculate yearly totals
                    yearly_totals_awal = awal_df_.groupby('Nama Pelanggan').sum()
                    yearly_totals_akhir = akhir_df_.groupby('Nama Pelanggan').sum()

                    contributor_grouped['Tahun Awal'] = contributor_grouped['Nama'].map(yearly_totals_awal).fillna(0)
                    contributor_grouped['Tahun Akhir'] = contributor_grouped['Nama'].map(yearly_totals_akhir).fillna(0)
                    contributor_grouped['Perubahan Absolut'] = contributor_grouped['Tahun Akhir'] - contributor_grouped['Tahun Awal']
                    contributor_grouped['Perubahan (%)'] = contributor_grouped.apply(
                        lambda row: ((row['Perubahan Absolut'] / row['Tahun Awal']) * 100) if row['Tahun Awal'] != 0 else 0,
                        axis=1
                    )

                    # Terapkan filter top/bottom
                    if selected_contrib == 'Top 5':
                        contributor_grouped = contributor_grouped.nlargest(5, 'Total Konsumsi')
                    elif selected_contrib == 'Top 10':
                        contributor_grouped = contributor_grouped.nlargest(10, 'Total Konsumsi')
                    elif selected_contrib == 'Top 25':
                        contributor_grouped = contributor_grouped.nlargest(25, 'Total Konsumsi')
                    elif selected_contrib == 'Top 50':
                        contributor_grouped = contributor_grouped.nlargest(50, 'Total Konsumsi')
                    elif selected_contrib == 'Top 100':
                        contributor_grouped = contributor_grouped.nlargest(100, 'Total Konsumsi')
                    elif selected_contrib == 'Bottom 5':
                        contributor_grouped = contributor_grouped.nsmallest(5, 'Total Konsumsi')
                    elif selected_contrib == 'Bottom 10':
                        contributor_grouped = contributor_grouped.nsmallest(10, 'Total Konsumsi')
                    elif selected_contrib == 'Bottom 25':
                        contributor_grouped = contributor_grouped.nsmallest(25, 'Total Konsumsi')
                    elif selected_contrib == 'Bottom 50':
                        contributor_grouped = contributor_grouped.nsmallest(50, 'Total Konsumsi')
                    elif selected_contrib == 'Bottom 100':
                        contributor_grouped = contributor_grouped.nsmallest(100, 'Total Konsumsi')
                    elif selected_contrib == 'Highest Change':
                        contributor_grouped = contributor_grouped.nlargest(5, 'Perubahan Absolut')
                    elif selected_contrib == 'Lowest Change':
                        contributor_grouped = contributor_grouped.nsmallest(5, 'Perubahan Absolut')
                    elif selected_contrib == 'All':
                        pass  # Tampilkan semua

                    # Format angka
                    contributor_grouped['Total Konsumsi'] = contributor_grouped['Total Konsumsi'].apply(lambda x: f"{x:,.2f} United States Dollar (USD)")
                    contributor_grouped['Tahun Awal'] = contributor_grouped['Tahun Awal'].apply(lambda x: f"{x:,.2f} United States Dollar (USD)")
                    contributor_grouped['Tahun Akhir'] = contributor_grouped['Tahun Akhir'].apply(lambda x: f"{x:,.2f} United States Dollar (USD)")
                    contributor_grouped['Perubahan Absolut'] = contributor_grouped['Perubahan Absolut'].apply(lambda x: f"{x:,.2f} United States Dollar (USD)")
                    contributor_grouped['Perubahan (%)'] = contributor_grouped['Perubahan (%)'].apply(lambda x: f"{x:,.2f}%")

                    # Reorder columns to a more logical sequence
                    column_order = [
                        'Nama', 'CM', 'Sektor', 'Segment', 'Ketentuan', 'Kepmen',
                        'Total Konsumsi', 'Tahun Awal', 'Tahun Akhir',
                        'Perubahan Absolut', 'Perubahan (%)'
                    ]
                    contributor_grouped = contributor_grouped[column_order]

                    st.dataframe(contributor_grouped, use_container_width=True)

                    # Add after contributors table (st.dataframe(contributor_grouped))
                    if not df_contributor.empty:
                        excel_buffer_contrib = io.BytesIO()
                        with pd.ExcelWriter(excel_buffer_contrib, engine='openpyxl') as writer:
                            contributor_grouped.to_excel(writer, index=False, sheet_name='Contributors')
                        excel_buffer_contrib.seek(0)

                        st.download_button(
                            label=f"Download Kontributor Utama (Excel) - {selected_unit}",
                            data=excel_buffer_contrib,
                            file_name=f"Contributors_{year_range[0]}_{year_range[1]}_{selected_contrib}_{selected_unit}.xlsx",
                            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                            key=f"download_contrib_{key_prefix}"
                        )
                        
            # If dual panel mode is on, create two columns
            if dual_panel_mode:
                col1, col2 = st.columns(2)
                
                with col1:
                    st.write("#### Panel 1")
                    create_comparison_section("panel1")
                
                with col2:
                    st.write("#### Panel 2")
                    create_comparison_section("panel2")
            else:
                # Single panel mode
                create_comparison_section("single")
# -----------------------------------------------------
# 7. Dashboard Monitoring
# -----------------------------------------------------
elif dashboard_choice == "Dashboard Monitoring USD":
    st.header("Dashboard Monitoring USD")

    # Upload File Excel di Sidebar
    st.sidebar.markdown("### Upload File Excel untuk Dashboard Monitoring")
    uploaded_file_monitoring = st.sidebar.file_uploader(
        label="Upload File Monitoring",
        type=["xlsx"],
        help="Upload file Excel (.xlsx) untuk Dashboard Monitoring.",
        key="monitoring_uploader"
    )

    if uploaded_file_monitoring is not None:
        df_bulanan_monitoring, df_harian_monitoring, list_pelanggan_monitoring, list_tahun_monitoring = parse_excel_monitoring(uploaded_file_monitoring)
        if df_bulanan_monitoring is not None and df_harian_monitoring is not None:
            st.sidebar.success(f"File {uploaded_file_monitoring.name} berhasil di-upload dan diproses!")
            st.session_state["df_monitoring_bulanan"] = df_bulanan_monitoring
            st.session_state["df_monitoring_harian"] = df_harian_monitoring
            st.session_state["list_pelanggan_monitoring"] = list_pelanggan_monitoring
            st.session_state["list_tahun_monitoring"] = list_tahun_monitoring

            # Penting: Pastikan kolom2 yang akan di-sort adalah string
            st.session_state["df_monitoring_bulanan"]["CM"] = st.session_state["df_monitoring_bulanan"]["CM"].astype(str)
            st.session_state["df_monitoring_bulanan"]["Segment"] = st.session_state["df_monitoring_bulanan"]["Segment"].astype(str)
            st.session_state["df_monitoring_bulanan"]["Sektor"] = st.session_state["df_monitoring_bulanan"]["Sektor"].astype(str)

        else:
            st.sidebar.error("File tidak valid atau gagal diproses.")
    else:
        st.sidebar.warning("Belum ada file yang diupload untuk Dashboard Monitoring.")

    # Jika data sudah diupload, tampilkan tabs monitoring
    if (
        st.session_state["df_monitoring_bulanan"] is not None and
        st.session_state["df_monitoring_harian"] is not None
    ):
        tabs_monitoring = st.tabs([
            "Monitoring Monthly",
            "Monitoring Daily"
        ])

        # -------------------------------------------------
        # 1. TAB: MONITORING MONTHLY
        # -------------------------------------------------
        with tabs_monitoring[0]:
            st.subheader("Monitoring Monthly")

            # Function to create monitoring panel
            def create_monitoring_panel(df_bulanan_all, panel_key):
                # Add year range slider for each panel
                all_years = sorted(df_bulanan_all["Tahun"].unique())
                min_year = min(all_years)
                max_year = max(all_years)
                
                year_range = st.slider(
                    "Pilih Rentang Tahun",
                    min_value=min_year,
                    max_value=max_year,
                    value=(min_year, max_year),  # Default to full range
                    step=1,
                    key=f"year_range_slider_{panel_key}"
                )

                # Add unit selection inside panel
                selected_unit = st.radio(
                    "Pilih Satuan",
                    options=['United States Dollar (USD)'],
                    horizontal=True,
                    key=f"unit_selection_{panel_key}"
                )
                
                # Filter untuk subset data sementara berdasarkan filter sebelumnya
                filtered_df = df_bulanan_all.copy()
                
                # Filter based on year range first
                filtered_df = filtered_df[filtered_df['Tahun'].between(year_range[0], year_range[1])]

                # Dapatkan daftar unique
                list_cm_raw = sorted(filtered_df["CM"].dropna().unique())
                cm_options = ["All CM", "COR", "NON COR"] + [x for x in list_cm_raw if x not in ["COR", "NON COR"]]

                # Row 1: CM, Pelanggan, Sektor
                filter_row1 = st.columns(3)
                
                with filter_row1[0]:
                    selected_cm = st.selectbox(
                        "Pilih CM",
                        options=cm_options,
                        index=0,
                        key=f"cm_select_{panel_key}"
                    )

                # Filter berdasarkan CM
                if selected_cm == "NON COR":
                    filtered_df = filtered_df[filtered_df['CM'] != "COR"]
                elif selected_cm != "All CM":
                    filtered_df = filtered_df[filtered_df['CM'] == selected_cm]

                with filter_row1[1]:
                    selected_pelanggan = st.selectbox(
                        "Pilih Pelanggan",
                        options=["All Pelanggan"] + sorted(filtered_df["Nama Pelanggan"].dropna().unique()),
                        index=0,
                        key=f"pelanggan_select_{panel_key}"
                    )

                with filter_row1[2]:
                    selected_sektor = st.selectbox(
                        "Pilih Sektor",
                        options=["All Sektor"] + sorted(filtered_df["Sektor"].dropna().unique()),
                        index=0,
                        key=f"sektor_select_{panel_key}"
                    )

                # Row 2: Segment, Ketentuan, Kepmen
                filter_row2 = st.columns(3)
                
                with filter_row2[0]:
                    selected_segment = st.selectbox(
                        "Pilih Segment",
                        options=["All Segment"] + sorted(filtered_df["Segment"].dropna().unique()),
                        index=0,
                        key=f"segment_select_{panel_key}"
                    )

                with filter_row2[1]:
                    selected_ketentuan = st.selectbox(
                        "Pilih Ketentuan",
                        options=["All Ketentuan"] + sorted(filtered_df["Ketentuan"].dropna().unique()),
                        index=0,
                        key=f"ketentuan_select_{panel_key}"
                    )

                with filter_row2[2]:
                    selected_kepmen = st.selectbox(
                        "Pilih Kepmen",
                        options=["All Kepmen"] + sorted(filtered_df["Kepmen"].dropna().unique()),
                        index=0,
                        key=f"kepmen_select_{panel_key}"
                    )

                # Apply all filters
                if selected_pelanggan != "All Pelanggan":
                    filtered_df = filtered_df[filtered_df['Nama Pelanggan'] == selected_pelanggan]
                if selected_sektor != "All Sektor":
                    filtered_df = filtered_df[filtered_df['Sektor'] == selected_sektor]
                if selected_segment != "All Segment":
                    filtered_df = filtered_df[filtered_df['Segment'] == selected_segment]
                if selected_ketentuan != "All Ketentuan":
                    filtered_df = filtered_df[filtered_df['Ketentuan'] == selected_ketentuan]
                if selected_kepmen != "All Kepmen":
                    filtered_df = filtered_df[filtered_df['Kepmen'] == selected_kepmen]

                # Process monthly data with accumulation
                if selected_pelanggan == "All Pelanggan":
                    # First group by customer and date to get customer-level monthly totals
                    customer_monthly = filtered_df.groupby(['Nama Pelanggan', 'Bulan', 'Tahun']).agg({
                        'Real_Monthly': 'sum',
                        'Kmin': 'sum',
                        'Kmax': 'sum',
                        'Nominasi': 'sum',
                        'RKAP_Actual': 'sum',
                        'RR%': 'mean'
                    }).reset_index()

                    # Create continuous date for sorting
                    bulan_order = {
                        'Jan': 1, 'Feb': 2, 'Mar': 3, 'Apr': 4, 'May': 5, 'Jun': 6,
                        'Jul': 7, 'Aug': 8, 'Sep': 9, 'Oct': 10, 'Nov': 11, 'Dec': 12
                    }
                    
                    customer_monthly['Date'] = customer_monthly.apply(
                        lambda x: pd.to_datetime(f"{x['Tahun']}-{bulan_order[x['Bulan'].split('-')[0]]:02d}-01"),
                        axis=1
                    )
                    
                    # Sort and calculate accumulation per customer
                    customer_monthly = customer_monthly.sort_values(['Nama Pelanggan', 'Date'])
                    customer_monthly['RKAP_Ac'] = customer_monthly.groupby(['Nama Pelanggan', 'Tahun'])['RKAP_Actual'].cumsum()
                    customer_monthly['Real_Ac_Monthly'] = customer_monthly.groupby(['Nama Pelanggan', 'Tahun'])['Real_Monthly'].cumsum()
                    
                    # Now aggregate to total level
                    filtered_bulanan = customer_monthly.groupby(['Bulan', 'Tahun', 'Date']).agg({
                        'Real_Monthly': 'sum',
                        'Kmin': 'sum',
                        'Kmax': 'sum',
                        'Nominasi': 'sum',
                        'RKAP_Actual': 'sum',
                        'RKAP_Ac': 'sum',
                        'Real_Ac_Monthly': 'sum',
                        'RR%': 'mean'
                    }).reset_index()
                    
                    # Calculate Ac% after aggregation
                    filtered_bulanan['Ac%'] = (filtered_bulanan['Real_Ac_Monthly'] / filtered_bulanan['RKAP_Ac'] * 100).round(2)
                    filtered_bulanan['Nama Pelanggan'] = "All Pelanggan"

                else:
                    # For individual customer, just calculate directly
                    filtered_bulanan = filtered_df.copy()
                    
                    # Create continuous date for sorting
                    bulan_order = {
                        'Jan': 1, 'Feb': 2, 'Mar': 3, 'Apr': 4, 'May': 5, 'Jun': 6,
                        'Jul': 7, 'Aug': 8, 'Sep': 9, 'Oct': 10, 'Nov': 11, 'Dec': 12
                    }
                    
                    filtered_bulanan['Date'] = filtered_bulanan.apply(
                        lambda x: pd.to_datetime(f"{x['Tahun']}-{bulan_order[x['Bulan'].split('-')[0]]:02d}-01"),
                        axis=1
                    )
                    
                    # Sort and calculate accumulation
                    filtered_bulanan = filtered_bulanan.sort_values(['Date'])
                    filtered_bulanan['RKAP_Ac'] = filtered_bulanan.groupby('Tahun')['RKAP_Actual'].cumsum()
                    filtered_bulanan['Real_Ac_Monthly'] = filtered_bulanan.groupby('Tahun')['Real_Monthly'].cumsum()
                    filtered_bulanan['Ac%'] = (filtered_bulanan['Real_Ac_Monthly'] / filtered_bulanan['RKAP_Ac'] * 100).round(2)

                if not filtered_bulanan.empty and 'Tahun' in filtered_bulanan.columns:
                    filtered_bulanan['Tahun'] = filtered_bulanan['Tahun'].astype(str)

                # Add display options
                display_option = st.radio(
                    "Opsi Tampilan Data",
                    options=['Tampilkan Data Real', 'Tampilkan Akumulasi', 'Tampilkan Semua'],
                    horizontal=True,
                    key=f"display_option_{panel_key}"
                )

                # Create plot
                if not filtered_bulanan.empty:
                    # Sort by date for continuous plotting
                    filtered_bulanan_sorted = filtered_bulanan.sort_values('Date')

                    fig_bulanan = go.Figure()

                    # Define metrics based on display option
                    if display_option == 'Tampilkan Data Real':
                        metrics = [
                            ('Real_Monthly', 'blue'),
                            ('Kmin', 'green'),
                            ('Kmax', 'orange'),
                            ('Nominasi', 'purple'),
                            ('RKAP_Actual', 'red')
                        ]
                    elif display_option == 'Tampilkan Akumulasi':
                        metrics = [
                            ('RKAP_Ac', 'red'),
                            ('Real_Ac_Monthly', 'blue')
                        ]
                        # Add Ac% on secondary y-axis
                        fig_bulanan.add_trace(go.Scatter(
                            x=filtered_bulanan_sorted['Date'],
                            y=filtered_bulanan_sorted['Ac%'],
                            mode='lines+markers',
                            name='Ac%',
                            line=dict(color='purple', dash='dot'),
                            marker=dict(size=8),
                            yaxis='y2'
                        ))
                    else:  # Tampilkan Semua
                        metrics = [
                            ('Real_Monthly', 'blue'),
                            ('Kmin', 'green'),
                            ('Kmax', 'orange'),
                            ('Nominasi', 'purple'),
                            ('RKAP_Actual', 'red'),
                            ('RKAP_Ac', 'brown'),
                            ('Real_Ac_Monthly', 'cyan')
                        ]
                        # Add Ac% on secondary y-axis
                        fig_bulanan.add_trace(go.Scatter(
                            x=filtered_bulanan_sorted['Date'],
                            y=filtered_bulanan_sorted['Ac%'],
                            mode='lines+markers',
                            name='Ac%',
                            line=dict(color='purple', dash='dot'),
                            marker=dict(size=8),
                            yaxis='y2'
                        ))

                    # Add main metrics
                    for metric, color in metrics:
                        fig_bulanan.add_trace(go.Scatter(
                            x=filtered_bulanan_sorted['Date'],
                            y=filtered_bulanan_sorted[metric],
                            mode='lines+markers',
                            name=metric,
                            line=dict(color=color),
                            marker=dict(size=8)
                        ))

                    # Add unit to title
                    year_range_text = f"{year_range[0]}-{year_range[1]}" if year_range[0] != year_range[1] else str(year_range[0])
                    title_text = (f'Grafik Bulanan: {selected_pelanggan} - {year_range_text} (United States Dollar - USD)'
                                if selected_pelanggan != "All Pelanggan"
                                else f'Grafik Bulanan: All Pelanggan - {year_range_text} (United States Dollar - USD)')
                    
                    # Update layout with secondary y-axis if showing accumulation or all data
                    if display_option in ['Tampilkan Akumulasi', 'Tampilkan Semua']:
                        fig_bulanan.update_layout(
                            title={
                                'text': title_text,
                                'y': 0.9,
                                'x': 0.5,
                                'xanchor': 'center',
                                'yanchor': 'top'
                            },
                            xaxis_title="Date",
                            yaxis_title="United States Dollar (USD)",
                            yaxis2=dict(
                                title="Ac%",
                                overlaying="y",
                                side="right",
                                ticksuffix="%"
                            ),
                            hovermode='x unified'
                        )
                    else:
                        fig_bulanan.update_layout(
                            title={
                                'text': title_text,
                                'y': 0.9,
                                'x': 0.5,
                                'xanchor': 'center',
                                'yanchor': 'top'
                            },
                            xaxis_title="Date",
                            yaxis_title="United States Dollar (USD)",
                            hovermode='x unified'
                        )

                    # Format x-axis to show month and year
                    fig_bulanan.update_xaxes(
                        tickformat="%b %Y",
                        tickangle=45,
                        tickmode='auto',
                        nticks=20
                    )

                    st.plotly_chart(fig_bulanan, use_container_width=True, key=f"chart_{panel_key}")

                    # Display tables and additional functionality
                    columns_to_display = [
                        'Nama Pelanggan', 'Bulan', 'Tahun',
                        'Kmin', 'Kmax', 'Nominasi', 'Real_Monthly', 'RKAP_Actual', 'RR%',
                        'RKAP_Ac', 'Real_Ac_Monthly', 'Ac%'
                    ]
                    cols_available = [col for col in columns_to_display if col in filtered_bulanan.columns]

                    # Sort table by Date
                    filtered_bulanan_sorted = filtered_bulanan.sort_values('Date')

                    # Display the sorted table with formatted numbers
                    st.dataframe(
                        filtered_bulanan_sorted[cols_available].style.format({
                            'Kmin': '{:.2f}',
                            'Kmax': '{:.2f}',
                            'Nominasi': '{:.2f}',
                            'Real_Monthly': '{:.2f}',
                            'RKAP_Actual': '{:.2f}',
                            'RKAP_Ac': '{:.2f}',
                            'Real_Ac_Monthly': '{:.2f}',
                            'Ac%': '{:.2f}',
                            'RR%': '{:.2f}'
                        }),
                        use_container_width=True,
                        key=f"table_{panel_key}"
                    )

                    # Add Excel download functionality
                    excel_buffer = io.BytesIO()
                    with pd.ExcelWriter(excel_buffer, engine='openpyxl') as writer:
                        filtered_bulanan_sorted[cols_available].to_excel(writer, index=False, sheet_name='Monthly Data')
                    excel_buffer.seek(0)
                    
                    st.download_button(
                        label=f"Download Data Bulanan (Excel) - United States Dollar (USD)",
                        data=excel_buffer,
                        file_name=f"Monthly_Data_{selected_pelanggan}_{year_range_text}_USD.xlsx",
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                        key=f"download_{panel_key}"
                    )

                    # Display summary
                    st.write("#### Ringkasan Data Bulanan")
                    if not filtered_bulanan.empty:
                        summary_stats = filtered_bulanan[['Real_Monthly', 'RKAP_Actual', 'RKAP_Ac', 
                                                        'Real_Ac_Monthly', 'Ac%', 'RR%']].describe()
                        st.dataframe(summary_stats.style.format("{:.2f}"), 
                                use_container_width=True, 
                                key=f"summary_{panel_key}")
                    else:
                        st.info("Data kosong, tidak bisa menampilkan ringkasan.")

                    # Display contributor table
                    st.write("### Tabel Kontributor")
                    st.caption("Lihat kontributor (mis. top-botom) berdasarkan Real_Monthly.")

                    if filtered_df.empty:
                        st.info("Tidak ada data. Tabel kontributor kosong.")
                    else:
                        if 'Real_Monthly' not in filtered_df.columns:
                            st.info("Kolom 'Real_Monthly' tidak ditemukan. Tidak bisa menampilkan kontributor.")
                        else:
                            df_contributor = filtered_df.groupby(
                                ['Nama Pelanggan', 'CM', 'Sektor', 'Segment', 'Ketentuan', 'Kepmen'],
                                as_index=False
                            )['Real_Monthly'].sum()

                            df_contributor.rename(columns={'Real_Monthly': 'Total_Real_Monthly'}, inplace=True)

                            contributor_filter = st.selectbox(
                                "Filter Kontributor",
                                [
                                    "All",
                                    "All Highest",
                                    "All Lowest",
                                    "Top 5",
                                    "Top 10",
                                    "Bottom 5",
                                    "Bottom 10"
                                ],
                                key=f"contributor_filter_{panel_key}"
                            )

                            df_display = df_contributor.copy()

                            if contributor_filter == "All":
                                pass
                            elif contributor_filter == "All Highest":
                                df_display = df_display.sort_values('Total_Real_Monthly', ascending=False)
                            elif contributor_filter == "All Lowest":
                                df_display = df_display.sort_values('Total_Real_Monthly', ascending=True)
                            elif contributor_filter == "Top 5":
                                df_display = df_display.nlargest(5, 'Total_Real_Monthly')
                            elif contributor_filter == "Top 10":
                                df_display = df_display.nlargest(10, 'Total_Real_Monthly')
                            elif contributor_filter == "Bottom 5":
                                df_display = df_display.nsmallest(5, 'Total_Real_Monthly')
                            elif contributor_filter == "Bottom 10":
                                df_display = df_display.nsmallest(10, 'Total_Real_Monthly')

                            df_display['Total_Real_Monthly'] = df_display['Total_Real_Monthly'].round(2)
                            st.dataframe(df_display, use_container_width=True, key=f"contributor_table_{panel_key}")
                            
                            # Add this code after st.dataframe(df_display)
                            if not df_display.empty:
                                excel_buffer = io.BytesIO()
                                with pd.ExcelWriter(excel_buffer, engine='openpyxl') as writer:
                                    df_display.to_excel(writer, index=False, sheet_name='Contributors')
                                excel_buffer.seek(0)

                                filename = f"Monthly_Contributors_{selected_unit}_{contributor_filter}.xlsx"

                                st.download_button(
                                    label=f"Download Tabel Kontributor Bulanan (Excel) - {selected_unit}",
                                    data=excel_buffer,
                                    file_name=filename,
                                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                                    key=f"download_contributor_{panel_key}"
                                )
            # Add toggle button for comparison view
            enable_comparison = st.toggle("Enable Comparison View", key="enable_comparison_view")

            if enable_comparison:
                # Create two columns for side-by-side comparison
                left_col, right_col = st.columns(2)

                # Create left panel
                with left_col:
                    st.markdown("### Left Panel")
                    create_monitoring_panel(st.session_state["df_monitoring_bulanan"], "panel_left")

                # Create right panel
                with right_col:
                    st.markdown("### Right Panel")
                    create_monitoring_panel(st.session_state["df_monitoring_bulanan"], "panel_right")

            else:
                # Original single panel view
                df_bulanan_all = st.session_state["df_monitoring_bulanan"]
                create_monitoring_panel(df_bulanan_all, "panel_single")

        # -------------------------------------------------
        # 2. TAB: MONITORING DAILY
        # -------------------------------------------------
        with tabs_monitoring[1]:
            st.subheader("Monitoring Daily")

            # Add toggle button for comparison view
            show_comparison = st.toggle("Show Comparison View", key="show_comparison_view")

            # Ambil data harian all
            df_harian_all = st.session_state["df_monitoring_harian"].copy()

            def create_monitoring_panel(df_harian_all, panel_key=""):
                # Add unit selection radio button with unique key for each panel
                selected_unit = st.radio(
                    "Pilih Satuan",
                    options=['United States Dollar (USD)'],
                    horizontal=True,
                    key=f"unit_selection{panel_key}"
                )

                # Function to convert values based on selected unit
                def convert_values(df, unit):
                    return df.copy()  # No conversion needed since we only have USD

                # Apply unit conversion to the dataframe
                df_harian_all_converted = convert_values(df_harian_all, selected_unit)

                # --- Pastikan kolomnya ada dan formatkan tipe data ---
                columns_to_check = ["CM", "Sektor", "Segment", "Ketentuan", "Kepmen"]
                for col in columns_to_check:
                    if col not in df_harian_all_converted.columns:
                        df_harian_all_converted[col] = ""
                    # Convert to string and handle NaN/None values
                    df_harian_all_converted[col] = df_harian_all_converted[col].fillna("").astype(str)

                # --- Dapatkan list unik CM, Sektor, Segment ---
                list_cm_raw = sorted([str(x) for x in df_harian_all_converted["CM"].unique() if x and x != "nan"])
                cm_options = ["All CM", "COR", "NON COR"] + [x for x in list_cm_raw if x not in ["COR"]]
                
                # --------------------------------------------
                # BARIS 1: 6 FILTER
                # --------------------------------------------
                col1, col2, col3, col4, col5, col6 = st.columns(6)

                # 1. Filter CM first
                with col2:
                    selected_cm_daily = st.selectbox(
                        "Pilih CM",
                        options=cm_options,
                        index=0,
                        key=f"monitoring_daily_cm_select{panel_key}"
                    )

                # Filter berdasarkan CM
                df_filtered = df_harian_all_converted.copy()
                if selected_cm_daily == "NON COR":
                    df_filtered = df_filtered[df_filtered['CM'] != "COR"]
                elif selected_cm_daily != "All CM":
                    df_filtered = df_filtered[df_filtered['CM'] == selected_cm_daily]

                # 2. Filter Sektor
                list_sektor = ["All Sektor"] + sorted([str(x) for x in df_filtered["Sektor"].unique() if x and x != "nan"])
                with col3:
                    selected_sektor_daily = st.selectbox(
                        "Pilih Sektor",
                        options=list_sektor,
                        index=0,
                        key=f"monitoring_daily_sektor_select{panel_key}"
                    )

                if selected_sektor_daily != "All Sektor":
                    df_filtered = df_filtered[df_filtered['Sektor'] == selected_sektor_daily]

                # 3. Filter Segment
                list_segment = ["All Segment"] + sorted([str(x) for x in df_filtered["Segment"].unique() if x and x != "nan"])
                with col4:
                    selected_segment_daily = st.selectbox(
                        "Pilih Segment",
                        options=list_segment,
                        index=0,
                        key=f"monitoring_daily_segment_select{panel_key}"
                    )

                if selected_segment_daily != "All Segment":
                    df_filtered = df_filtered[df_filtered['Segment'] == selected_segment_daily]

                # 4. Filter Ketentuan
                list_ketentuan = ["All Ketentuan"] + sorted([str(x) for x in df_filtered["Ketentuan"].unique() if x and x != "nan"])
                with col5:
                    selected_ketentuan_daily = st.selectbox(
                        "Pilih Ketentuan",
                        options=list_ketentuan,
                        index=0,
                        key=f"monitoring_daily_ketentuan_select{panel_key}"
                    )

                if selected_ketentuan_daily != "All Ketentuan":
                    df_filtered = df_filtered[df_filtered['Ketentuan'] == selected_ketentuan_daily]

                # 5. Filter Kepmen
                list_kepmen = ["All Kepmen"] + sorted([str(x) for x in df_filtered["Kepmen"].unique() if x and x != "nan"])
                with col6:
                    selected_kepmen_daily = st.selectbox(
                        "Pilih Kepmen",
                        options=list_kepmen,
                        index=0,
                        key=f"monitoring_daily_kepmen_select{panel_key}"
                    )

                if selected_kepmen_daily != "All Kepmen":
                    df_filtered = df_filtered[df_filtered['Kepmen'] == selected_kepmen_daily]

                # 6. Filter Pelanggan
                list_pelanggan = ["All Pelanggan"] + sorted([str(x) for x in df_filtered["Nama Pelanggan"].unique() if x and x != "nan"])
                with col1:
                    selected_pelanggan_daily = st.selectbox(
                        "Pilih Pelanggan",
                        options=list_pelanggan,
                        index=0,
                        key=f"monitoring_daily_pelanggan_select{panel_key}"
                    )

                if selected_pelanggan_daily != "All Pelanggan":
                    df_filtered = df_filtered[df_filtered['Nama Pelanggan'] == selected_pelanggan_daily]

                # --------------------------------------------
                # BARIS 2: PILIH RENTANG TANGGAL
                # --------------------------------------------
                min_date = df_filtered['Tanggal'].min()
                max_date = df_filtered['Tanggal'].max()
                default_range = (min_date.date(), max_date.date()) if (min_date and max_date) else None

                date_range = st.date_input(
                    "Pilih Rentang Tanggal",
                    value=default_range,
                    min_value=min_date.date() if min_date else None,
                    max_value=max_date.date() if max_date else None,
                    key=f"monitoring_daily_date_range{panel_key}"
                )

                if isinstance(date_range, tuple) and len(date_range) == 2:
                    start_date, end_date = date_range
                else:
                    start_date, end_date = default_range if default_range else (None, None)

                if start_date and end_date:
                    df_filtered = df_filtered[
                        (df_filtered['Tanggal'] >= pd.to_datetime(start_date)) &
                        (df_filtered['Tanggal'] <= pd.to_datetime(end_date))
                    ]

                # Modification for All Pelanggan view
                if selected_pelanggan_daily == "All Pelanggan":
                    # Aggregate data for All Pelanggan
                    filtered_harian = df_filtered.groupby('Tanggal').agg({
                        'CM': 'first',
                        'Sektor': 'first',
                        'Segment': 'first',
                        'Ketentuan': 'first',
                        'Kepmen': 'first',
                        'Real_Daily': 'sum',
                        'Kmin': 'mean',
                        'Kmax': 'mean',
                        'Nominasi': 'sum',
                        'RKAP_Actual': 'sum',
                        'Tahun': 'first'
                    }).reset_index()
                    
                    # Set Nama Pelanggan to "All Pelanggan"
                    filtered_harian['Nama Pelanggan'] = "All Pelanggan"
                else:
                    filtered_harian = df_filtered

                # Add calculation for accumulated columns
                filtered_harian['Month'] = filtered_harian['Tanggal'].dt.to_period('M')
                filtered_harian['Year'] = filtered_harian['Tanggal'].dt.year

                if selected_pelanggan_daily == "All Pelanggan":
                    # For All Pelanggan view, sum by date first
                    accumulated = filtered_harian.groupby(['Year', 'Month', 'Tanggal']).agg({
                        'RKAP_Actual': 'sum',
                        'Real_Daily': 'sum'
                    }).reset_index()
                else:
                    accumulated = filtered_harian.copy()

                # Sort by date to ensure proper accumulation
                accumulated = accumulated.sort_values(['Year', 'Month', 'Tanggal'])

                # Calculate accumulations within each month
                accumulated['RKAP_Ac'] = accumulated.groupby(['Year', 'Month'])['RKAP_Actual'].cumsum()
                accumulated['Real_Ac_Daily'] = accumulated.groupby(['Year', 'Month'])['Real_Daily'].cumsum()

                # Calculate Ac%
                accumulated['Ac%'] = np.where(
                    accumulated['RKAP_Ac'] == 0,
                    0,
                    (accumulated['Real_Ac_Daily'] / accumulated['RKAP_Ac']) * 100
                ).round(2)

                # Merge accumulated columns back to filtered_harian
                filtered_harian = filtered_harian.merge(
                    accumulated[['Tanggal', 'RKAP_Ac', 'Real_Ac_Daily', 'Ac%']],
                    on='Tanggal',
                    how='left'
                )

                # Normalize Kmin and Kmax
                filtered_harian['Days_in_Month'] = filtered_harian['Tanggal'].dt.days_in_month
                filtered_harian['Kmin'] = filtered_harian['Kmin'] / filtered_harian['Days_in_Month']
                filtered_harian['Kmax'] = filtered_harian['Kmax'] / filtered_harian['Days_in_Month']

                # Calculate RR%
                filtered_harian['RR%'] = np.where(
                    filtered_harian['RKAP_Actual'] == 0,
                    0,
                    (filtered_harian['Real_Daily'] / filtered_harian['RKAP_Actual']) * 100
                ).round(2)

                # Ensure Tahun is string
                filtered_harian['Tahun'] = filtered_harian['Tahun'].astype(str)

                # Add display options
                display_option = st.radio(
                    "Opsi Tampilan Data",
                    options=['Tampilkan Data Real', 'Tampilkan Akumulasi', 'Tampilkan Semua'],
                    horizontal=True,
                    key=f"display_option{panel_key}"
                )

                # --------------------------------------------
                # PLOT
                # --------------------------------------------
                if not filtered_harian.empty:
                    fig_harian = go.Figure()
                    
                    # Function to create trace with hover information
                    def create_trace(y_column, name, color, show_on_second_y=False):
                        hover_text = filtered_harian.apply(
                            lambda row: f"Nama Pelanggan: {row['Nama Pelanggan']}<br>" + 
                                    f"Tanggal: {row['Tanggal'].strftime('%Y-%m-%d')}<br>" + 
                                    f"{name}: {row[y_column]:.2f}" + 
                                    (" %" if name in ['RR%', 'Ac%'] else f" USD"),
                            axis=1
                        )
                        
                        return go.Scatter(
                            x=filtered_harian['Tanggal'],
                            y=filtered_harian[y_column],
                            mode='lines+markers',
                            name=name,
                            line=dict(color=color),
                            marker=dict(size=8),
                            hovertemplate='%{text}',
                            text=hover_text,
                            yaxis='y2' if show_on_second_y else 'y'
                        )

                    # Define metrics based on display option
                    if display_option == 'Tampilkan Data Real':
                        metrics = [
                            ('Real_Daily', 'Real_Daily', 'blue', False),
                            ('Kmin', 'Kmin', 'green', False),
                            ('Kmax', 'Kmax', 'orange', False),
                            ('Nominasi', 'Nominasi', 'purple', False),
                            ('RKAP_Actual', 'RKAP_Actual', 'red', False),
                            ('RR%', 'RR%', 'brown', True)
                        ]
                    elif display_option == 'Tampilkan Akumulasi':
                        metrics = [
                            ('RKAP_Ac', 'RKAP_Ac', 'red', False),
                            ('Real_Ac_Daily', 'Real_Ac_Daily', 'blue', False),
                            ('Ac%', 'Ac%', 'purple', True)
                        ]
                    else:  # Tampilkan Semua
                        metrics = [
                            ('Real_Daily', 'Real_Daily', 'blue', False),
                            ('Kmin', 'Kmin', 'green', False),
                            ('Kmax', 'Kmax', 'orange', False),
                            ('Nominasi', 'Nominasi', 'purple', False),
                            ('RKAP_Actual', 'RKAP_Actual', 'red', False),
                            ('RKAP_Ac', 'RKAP_Ac', 'brown', False),
                            ('Real_Ac_Daily', 'Real_Ac_Daily', 'cyan', False),
                            ('RR%', 'RR%', 'darkred', True),
                            ('Ac%', 'Ac%', 'darkgreen', True)
                        ]

                    # Add all traces
                    for col, name, color, second_y in metrics:
                        fig_harian.add_trace(create_trace(col, name, color, second_y))

                    title_str = f"Grafik Harian: {selected_pelanggan_daily} - {start_date} s/d {end_date} (USD)"
                    
                    # Update layout with secondary y-axis for percentage metrics
                    fig_harian.update_layout(
                        title=title_str,
                        hovermode='closest',
                        hoverdistance=100,
                        spikedistance=1000,
                        yaxis=dict(
                            title='United States Dollar (USD)',
                            side='left'
                        ),
                        yaxis2=dict(
                            title='Percentage (%)',
                            side='right',
                            overlaying='y',
                            ticksuffix='%'
                        )
                    )
                    
                    # Tambahkan unique key untuk plotly chart
                    st.plotly_chart(fig_harian, use_container_width=True, key=f"plotly_chart{panel_key}")
                else:
                    st.info("Tidak ada data harian untuk filter ini.")

                # --------------------------------------------
                # TAMPILKAN TABEL HARIAN
                # --------------------------------------------
                st.write(f"#### Data Harian (United States Dollar)")
                columns_to_display_daily = [
                    'Nama Pelanggan', 'Tanggal', 'Tahun',
                    'CM', 'Segment', 'Sektor',
                    'Kmin', 'Kmax', 'Nominasi',
                    'Real_Daily', 'RKAP_Actual', 'RR%',
                    'Real_Ac_Daily', 'RKAP_Ac', 'Ac%'  # Added new columns
                ]
                cols_daily_available = [c for c in columns_to_display_daily if c in filtered_harian.columns]

                # Add pagination
                page_size = 100
                total_rows = len(filtered_harian)

                # Checkbox to toggle full data view
                show_full_data = st.checkbox("Show Full Data", key=f"show_full_data_{panel_key}")

                if show_full_data:
                    # Display full data
                    st.dataframe(filtered_harian[cols_daily_available], use_container_width=True)
                else:
                    # Display first 100 rows
                    st.dataframe(filtered_harian[cols_daily_available].head(page_size), use_container_width=True)
                    
                
                # Always provide option to download full data
                def to_excel(df):
                    output = io.BytesIO()
                    with pd.ExcelWriter(output, engine="openpyxl") as writer:
                        df.to_excel(writer, index=False, sheet_name="Data Harian")
                    output.seek(0)
                    return output

                # Button to save Excel file with full data
                excel_data = to_excel(filtered_harian[cols_daily_available])
                st.download_button(
                    label="Download Full Data as Excel (United States Dollar)",
                    data=excel_data,
                    file_name="data_harian_full_USD.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                )

                # Display row count information
                st.write(f"Showing {min(page_size, total_rows)} of {total_rows} rows")

                # --------------------------------------------
                # RINGKASAN DATA HARIAN
                # --------------------------------------------
                st.write("#### Ringkasan Data Harian (United States Dollar)")
                if not filtered_harian.empty:
                    numeric_cols = [
                        'Real_Daily', 'Kmin', 'Kmax', 'Nominasi',
                        'RKAP_Actual', 'RR%'
                    ]
                    numeric_cols_available = [col for col in numeric_cols if col in filtered_harian.columns]

                    if numeric_cols_available:
                        df_numeric_summary = filtered_harian[numeric_cols_available].describe().transpose()
                        st.dataframe(df_numeric_summary, use_container_width=True)

                    if 'Tanggal' in filtered_harian.columns:
                        min_tanggal = filtered_harian['Tanggal'].min()
                        max_tanggal = filtered_harian['Tanggal'].max()
                        st.write(f"Rentang Tanggal: {min_tanggal.date()} s/d {max_tanggal.date()}")
                else:
                    st.info("Tidak ada data untuk ringkasan.")

                # --------------------------------------------
                # TABEL KONTRIBUTOR
                # --------------------------------------------
                st.write("### Tabel Kontributor Harian (United States Dollar)")
                st.caption("Lihat kontributor (mis. top-bottom) berdasarkan Real_Daily, menampilkan CM, Sektor, Segment, 'Ketentuan', 'Kepmen'.")

                if df_filtered.empty:
                    st.info("Tidak ada data. Tabel kontributor kosong.")
                else:
                    if 'Real_Daily' not in df_filtered.columns:
                        st.info("Kolom 'Real_Daily' tidak ditemukan. Tidak bisa menampilkan kontributor.")
                    else:
                        df_contributor = df_filtered.groupby(
                            ['Nama Pelanggan', 'CM', 'Sektor', 'Segment', 'Ketentuan', 'Kepmen'],
                            as_index=False
                        )['Real_Daily'].sum()

                        df_contributor.rename(columns={'Real_Daily': 'Total_Real_Daily'}, inplace=True)

                        contributor_filter = st.selectbox(
                            "Filter Kontributor",
                            [
                                "All",          # Tampilkan apa adanya (tanpa sort)
                                "All Highest",  # sort desc, tampil semua
                                "All Lowest",   # sort asc, tampil semua
                                "Top 5",
                                "Top 10",
                                "Bottom 5",
                                "Bottom 10"
                            ],
                            key=f"kontributor_filter_daily{panel_key}"
                        )

                        df_display = df_contributor.copy()

                        if contributor_filter == "All":
                            pass  # no sort
                        elif contributor_filter == "All Highest":
                            df_display = df_display.sort_values('Total_Real_Daily', ascending=False)
                        elif contributor_filter == "All Lowest":
                            df_display = df_display.sort_values('Total_Real_Daily', ascending=True)
                        elif contributor_filter == "Top 5":
                            df_display = df_display.nlargest(5, 'Total_Real_Daily')
                        elif contributor_filter == "Top 10":
                            df_display = df_display.nlargest(10, 'Total_Real_Daily')
                        elif contributor_filter == "Bottom 5":
                            df_display = df_display.nsmallest(5, 'Total_Real_Daily')
                        elif contributor_filter == "Bottom 10":
                            df_display = df_display.nsmallest(10, 'Total_Real_Daily')

                        # Format tampilan
                        df_display['Total_Real_Daily'] = df_display['Total_Real_Daily'].round(2)
                        st.dataframe(df_display, use_container_width=True)

                        # Add this code after st.dataframe(df_display)
                        if not df_display.empty:
                            excel_buffer = io.BytesIO()
                            with pd.ExcelWriter(excel_buffer, engine='openpyxl') as writer:
                                df_display.to_excel(writer, index=False, sheet_name='Daily Contributors')
                            excel_buffer.seek(0)

                            filename = f"Daily_Contributors_{selected_unit}_{contributor_filter}.xlsx"

                            st.download_button(
                                label=f"Download Tabel Kontributor Harian (Excel) - {selected_unit}",
                                data=excel_buffer,
                                file_name=filename,
                                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                                key=f"download_daily_contributor_{panel_key}"
                            )

            # Render panels based on comparison toggle
            if show_comparison:
                st.write("### Panel Comparison View")
                
                # Create two columns for side-by-side comparison
                left_col, right_col = st.columns(2)
                
                with left_col:
                    st.write("#### Left Panel")
                    create_monitoring_panel(df_harian_all, panel_key="_left")
                    
                with right_col:
                    st.write("#### Right Panel")
                    create_monitoring_panel(df_harian_all, panel_key="_right")
            else:
                # Show single panel when comparison is not enabled
                create_monitoring_panel(df_harian_all, panel_key="_single")