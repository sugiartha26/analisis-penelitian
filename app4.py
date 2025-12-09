import streamlit as st
import pandas as pd
import altair as alt
import re
from io import BytesIO

st.set_page_config(page_title="Research Data Explorer", layout="wide")

st.title("📊 Research Data Explorer – Universitas Negeri Padang")
st.write("Upload beberapa file Excel. Header otomatis dibaca dari baris ke-5.")

# ======================================================
# Fungsi konversi "Rp. 77.107.000" → 77107000
# ======================================================
def convert_rupiah_to_int(value):
    if pd.isna(value):
        return 0
    value = str(value)
    value = re.sub(r'[^0-9]', '', value)
    return int(value) if value.isdigit() else 0

# ======================================================
# Fungsi export excel
# ======================================================
def export_excel(df):
    output = BytesIO()
    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        df.to_excel(writer, index=False, sheet_name="Data")
    return output.getvalue()

# ======================================================
# Fungsi baca file Excel dengan header baris ke-5
# ======================================================
def load_excel(file):
    try:
        df = pd.read_excel(file, header=4)
        df = df.loc[:, ~df.columns.str.contains("^Unnamed")]

        if "DANA DISETUJUI" in df.columns:
            df["DANA_DISETUJUI_NUM"] = df["DANA DISETUJUI"].apply(convert_rupiah_to_int)

        return df
    except:
        return None

# ======================================================
# Upload multiple files
# ======================================================
uploaded_files = st.file_uploader(
    "📁 Upload beberapa file Excel",
    type=["xlsx", "xls"],
    accept_multiple_files=True
)

if uploaded_files:

    all_data = []  # list penyimpanan sementara

    for file in uploaded_files:
        df = load_excel(file)

        if df is None:
            st.error(f"❌ File '{file.name}' tidak sesuai format! Melewati file ini.")
        else:
            df["FILE_SUMBER"] = file.name  # opsional: jejak asal file
            all_data.append(df)

    if len(all_data) == 0:
        st.error("Tidak ada file yang valid.")
    else:
        # Gabungkan semua file
        df_all = pd.concat(all_data, ignore_index=True)

        st.success(f"✔ {len(all_data)} file berhasil diproses & digabungkan!")

        st.subheader("📄 Preview Data Gabungan")
        st.dataframe(df_all.head())

        # Tombol Export Tabel Gabungan
        st.download_button(
            label="📥 Download Data Gabungan (Excel)",
            data=export_excel(df_all),
            file_name="data_gabungan.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )

        # ======================================================
        # SIDEBAR FILTER
        # ======================================================
        st.sidebar.header("🔍 Filter Data")

        tahun_usulan = st.sidebar.multiselect(
            "Tahun Usulan",
            sorted(df_all["TAHUN USULAN KEGIATAN"].dropna().unique())
        )

        tahun_pelaksanaan = st.sidebar.multiselect(
            "Tahun Pelaksanaan",
            sorted(df_all["TAHUN PELAKSANAAN KEGIATAN"].dropna().unique())
        )

        bidang_fokus = st.sidebar.multiselect(
            "Bidang Fokus",
            sorted(df_all["BIDANG FOKUS"].dropna().unique())
        )

        program_hibah = st.sidebar.multiselect(
            "Program Hibah",
            sorted(df_all["PROGRAM HIBAH"].dropna().unique())
        )

        dana_min = int(df_all["DANA_DISETUJUI_NUM"].min())
        dana_max = int(df_all["DANA_DISETUJUI_NUM"].max())

        dana_range = st.sidebar.slider(
            "Rentang Dana Disetujui",
            min_value=dana_min,
            max_value=dana_max,
            value=(dana_min, dana_max),
            step=1000000
        )

        # ======================================================
        # APPLY FILTER
        # ======================================================
        df_filtered = df_all.copy()

        if tahun_usulan:
            df_filtered = df_filtered[df_filtered["TAHUN USULAN KEGIATAN"].isin(tahun_usulan)]

        if tahun_pelaksanaan:
            df_filtered = df_filtered[df_filtered["TAHUN PELAKSANAAN KEGIATAN"].isin(tahun_pelaksanaan)]

        if bidang_fokus:
            df_filtered = df_filtered[df_filtered["BIDANG FOKUS"].isin(bidang_fokus)]

        if program_hibah:
            df_filtered = df_filtered[df_filtered["PROGRAM HIBAH"].isin(program_hibah)]

        df_filtered = df_filtered[
            (df_filtered["DANA_DISETUJUI_NUM"] >= dana_range[0]) &
            (df_filtered["DANA_DISETUJUI_NUM"] <= dana_range[1])
        ]

        st.subheader("📊 Filtered Data")
        st.dataframe(df_filtered)

        # Tombol Export Data Filtered
        st.download_button(
            label="📥 Download Data Filtered (Excel)",
            data=export_excel(df_filtered),
            file_name="data_filtered.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )

        # ======================================================
        # VISUALISASI
        # ======================================================
        st.subheader("📈 Visualisasi Dana")

        col1, col2 = st.columns(2)

        # -------------------
        # Dana per tahun (bar)
        # -------------------
        with col1:
            st.markdown("### Total Dana per Tahun Usulan (Bar Chart)")
            chart_dana_bar = (
                alt.Chart(df_filtered)
                .mark_bar()
                .encode(
                    x="TAHUN USULAN KEGIATAN:O",
                    y="sum(DANA_DISETUJUI_NUM):Q",
                    tooltip=["TAHUN USULAN KEGIATAN", "sum(DANA_DISETUJUI_NUM)"]
                )
            )
            st.altair_chart(chart_dana_bar, use_container_width=True)

        with col2:
            st.markdown("### Tren Dana per Tahun Usulan (Line Chart)")
            chart_dana_line = (
                alt.Chart(df_filtered)
                .mark_line(point=True)
                .encode(
                    x="TAHUN USULAN KEGIATAN:O",
                    y="sum(DANA_DISETUJUI_NUM):Q",
                    tooltip=["TAHUN USULAN KEGIATAN", "sum(DANA_DISETUJUI_NUM)"]
                )
            )
            st.altair_chart(chart_dana_line, use_container_width=True)

        # -------------------
        # Dana per bidang fokus
        # -------------------
        st.markdown("### Dana per Bidang Fokus")
        bidang_chart = (
            alt.Chart(df_filtered)
            .mark_bar()
            .encode(
                x="sum(DANA_DISETUJUI_NUM):Q",
                y=alt.Y("BIDANG FOKUS:N", sort='-x'),
                tooltip=["BIDANG FOKUS", "sum(DANA_DISETUJUI_NUM)"]
            )
        )
        st.altair_chart(bidang_chart, use_container_width=True)

        # -------------------
        # Dana per program hibah
        # -------------------
        st.markdown("### Dana per Program Hibah")
        hibah_chart = (
            alt.Chart(df_filtered)
            .mark_bar()
            .encode(
                x="sum(DANA_DISETUJUI_NUM):Q",
                y=alt.Y("PROGRAM HIBAH:N", sort='-x'),
                tooltip=["PROGRAM HIBAH", "sum(DANA_DISETUJUI_NUM)"]
            )
        )
        st.altair_chart(hibah_chart, use_container_width=True)

else:
    st.info("Silakan upload satu atau beberapa file Excel.")
