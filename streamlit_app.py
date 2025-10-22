import streamlit as st
import pandas as pd
import io

# ----------------------------------------
# Page Configuration
# ----------------------------------------
st.set_page_config(
    page_title="Complaint Report by Branch",
    page_icon="📊",
    layout="wide"
)

# ----------------------------------------
# Custom Light-Themed Styling
# ----------------------------------------
st.markdown("""
    <style>
    html, body, [class*="css"] {
        font-family: 'Poppins', sans-serif;
        background-color: #f7f9fc;
        color: #2c3e50;
    }

    /* Title */
    .title {
        font-size: 2rem;
        font-weight: 700;
        color: #0078d7;
        text-align: center;
        margin-bottom: 1.5rem;
    }

    /* File Upload */
    div[data-testid="stFileUploader"] {
        background-color: #ffffff;
        border: 2px dashed #0078d7;
        border-radius: 10px;
        padding: 20px;
    }

    /* Selectbox */
    div[data-baseweb="select"] {
        background-color: #ffffff !important;
        color: #2c3e50 !important;
        border-radius: 10px !important;
    }

    /* Table Styling */
    table {
        border-collapse: collapse;
        width: 100%;
        background-color: white;
        border-radius: 10px;
        overflow: hidden;
        border: 1px solid #e0e6ed;
    }
    thead tr {
        background-color: #0078d7;
        color: white;
        text-align: left;
    }
    tbody tr:nth-child(even) {
        background-color: #f2f6fb;
    }
    tbody tr:hover {
        background-color: #e8f0fd;
    }

    /* Buttons */
    .stDownloadButton button {
        background: linear-gradient(90deg, #0078d7, #0094ff);
        color: white;
        border: none;
        border-radius: 8px;
        font-weight: 600;
        padding: 0.6rem 1.2rem;
    }
    .stDownloadButton button:hover {
        background: linear-gradient(90deg, #0094ff, #0078d7);
        transform: scale(1.02);
        transition: 0.2s ease-in-out;
    }

    /* Info box */
    .stAlert {
        border-radius: 10px;
        padding: 15px;
        background-color: #eef5ff;
        border-left: 4px solid #0078d7;
    }
    </style>
""", unsafe_allow_html=True)

# ----------------------------------------
# Title
# ----------------------------------------
st.markdown("<div class='title'>📊 Complaint Report by Branch</div>", unsafe_allow_html=True)

# ----------------------------------------
# Upload Section
# ----------------------------------------
uploaded_file = st.file_uploader("📂 Upload 'Data for Working.xlsx'", type=["xlsx"])

# MOP List
try:
    df_mop = pd.read_excel("MOP LIST.xlsx")
    df_mop = df_mop.rename(columns={'Item code': 'Item Code'})
except Exception as e:
    st.error("⚠️ Error loading 'MOP LIST.xlsx'. Please make sure it exists in the same directory.")
    st.stop()

# ----------------------------------------
# Data Handling
# ----------------------------------------
if uploaded_file is not None:
    df_complaints = pd.read_excel(uploaded_file)

    # Merge Data
    df = pd.merge(df_complaints, df_mop, on='Item Code', how='left')

    # Ensure numeric types
    df['MOP'] = pd.to_numeric(df['MOP'], errors='coerce')
    df['Days'] = pd.to_numeric(df['Days'], errors='coerce')

    # Drop invalid rows
    df = df.dropna(subset=['MOP', 'Days', 'Branch'])

    # Brand Filter
    brands = sorted(df['Brand'].dropna().unique())
    brands.insert(0, 'All')

    with st.sidebar:
        st.markdown("### 🔍 Filter Options")
        selected_brand = st.selectbox("Select Brand", options=brands)
        st.markdown("---")
        st.markdown("💡 *Use this filter to view reports for a specific brand.*")

    # Apply Filter
    filtered_df = df if selected_brand == 'All' else df[df['Brand'] == selected_brand]

    # ----------------------------------------
    # Report Generation
    # ----------------------------------------
    if not filtered_df.empty:
        report_df = filtered_df.groupby('Branch').agg(
            SumofMOP=('MOP', 'sum'),
            AverageofDays=('Days', 'mean'),
            CountofComplaintMode=('Complaint Mode', 'count')
        ).reset_index()

        report_df = report_df.rename(columns={
            'SumofMOP': 'Sum of MOP',
            'AverageofDays': 'Average of Days',
            'CountofComplaintMode': 'Count of Complaint Mode'
        }).sort_values('Sum of MOP', ascending=False)

        # ----------------------------------------
        # Display Data
        # ----------------------------------------
        st.markdown("### 📈 Branch Report")
        st.dataframe(
            report_df.style.format({'Sum of MOP': '{:,.2f}', 'Average of Days': '{:.1f}'})
        )

        # ----------------------------------------
        # Download Report
        # ----------------------------------------
        buffer = io.BytesIO()
        with pd.ExcelWriter(buffer, engine='xlsxwriter') as writer:
            report_df.to_excel(writer, index=False)
        buffer.seek(0)

        st.download_button(
            label="💾 Download Report as Excel",
            data=buffer,
            file_name=f"Complaint_Report_{selected_brand}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )

    else:
        st.info("No data available for the selected brand.")

else:
    st.info("Please upload the data file to proceed.")
