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
# Custom Light Theme
# ----------------------------------------
st.markdown("""
    <style>
    html, body, [class*="css"] {
        font-family: 'Poppins', sans-serif;
        background-color: #f7f9fc;
        color: #2c3e50;
    }

    .title {
        font-size: 2rem;
        font-weight: 700;
        color: #0078d7;
        text-align: center;
        margin-bottom: 1.5rem;
    }

    div[data-testid="stFileUploader"] {
        background-color: #ffffff;
        border: 2px dashed #0078d7;
        border-radius: 10px;
        padding: 20px;
    }

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
    }
    tbody tr:nth-child(even) {
        background-color: #f2f6fb;
    }
    tbody tr:hover {
        background-color: #e8f0fd;
    }

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
    </style>
""", unsafe_allow_html=True)

# ----------------------------------------
# Title
# ----------------------------------------
st.markdown("<div class='title'>📊 Complaint Report by Branch</div>", unsafe_allow_html=True)

# ----------------------------------------
# File Uploads
# ----------------------------------------
uploaded_complaints = st.file_uploader("📂 Upload 'Data for Working.xlsx'", type=["xlsx"])
uploaded_mop = st.file_uploader("📂 Upload 'MOP LIST.xlsx'", type=["xlsx"])

if uploaded_complaints is None or uploaded_mop is None:
    st.info("Please upload both 'Data for Working.xlsx' and 'MOP LIST.xlsx' to proceed.")
    st.stop()

# ----------------------------------------
# Load Files
# ----------------------------------------
try:
    df_complaints = pd.read_excel(uploaded_complaints)
    df_mop = pd.read_excel(uploaded_mop)
    df_mop = df_mop.rename(columns={'Item code': 'Item Code'})
except Exception as e:
    st.error(f"⚠️ Error reading files: {e}")
    st.stop()

# ----------------------------------------
# Data Processing
# ----------------------------------------
# Merge DataFrames
df = pd.merge(df_complaints, df_mop, on='Item Code', how='left')

# Ensure numeric types
df['MOP'] = pd.to_numeric(df['MOP'], errors='coerce')
df['Days'] = pd.to_numeric(df['Days'], errors='coerce')

# Drop invalid rows
df = df.dropna(subset=['MOP', 'Days', 'Branch'])

# Brand filter
brands = sorted(df['Brand'].dropna().unique())
brands.insert(0, 'All')

with st.sidebar:
    st.markdown("### 🔍 Filter Options")
    selected_brand = st.selectbox("Select Brand", options=brands)
    st.markdown("---")
    st.markdown("💡 *Use this filter to view reports for a specific brand.*")

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
    # Dynamic Number Formatting Function
    # ----------------------------------------
    def format_dynamic(x):
        if pd.isna(x):
            return ""
        elif float(x).is_integer():
            return int(x)
        else:
            return round(x, 1)

    # Display in Streamlit
    display_df = report_df.copy()
    display_df['Sum of MOP'] = display_df['Sum of MOP'].apply(format_dynamic)
    display_df['Average of Days'] = display_df['Average of Days'].apply(format_dynamic)
    st.markdown("### 📈 Branch Performance Summary")
    st.dataframe(display_df)

    # ----------------------------------------
    # Styled Excel Export
    # ----------------------------------------
    buffer = io.BytesIO()
    with pd.ExcelWriter(buffer, engine='xlsxwriter') as writer:
        report_df.to_excel(writer, index=False, sheet_name='Report')
        workbook = writer.book
        worksheet = writer.sheets['Report']

        # Formats
        header_format = workbook.add_format({
            'bold': True, 'bg_color': '#0078D7', 'font_color': 'white',
            'border': 1, 'align': 'center', 'valign': 'vcenter'
        })
        even_row = workbook.add_format({'bg_color': '#F2F6FB', 'border': 1})
        odd_row = workbook.add_format({'bg_color': '#FFFFFF', 'border': 1})
        number_fmt_int = workbook.add_format({'num_format': '0', 'border': 1})
        number_fmt_float = workbook.add_format({'num_format': '0.0', 'border': 1})
        text_fmt = workbook.add_format({'border': 1})

        # Header
        for col_num, value in enumerate(report_df.columns.values):
            worksheet.write(0, col_num, value, header_format)

        # Body with dynamic formatting
        for row_num, row_data in enumerate(report_df.values, start=1):
            fmt = even_row if row_num % 2 == 0 else odd_row
            for col_num, cell_value in enumerate(row_data):
                if isinstance(cell_value, (int, float)):
                    if float(cell_value).is_integer():
                        worksheet.write(row_num, col_num, cell_value, number_fmt_int)
                    else:
                        worksheet.write(row_num, col_num, round(cell_value,1), number_fmt_float)
                else:
                    worksheet.write(row_num, col_num, cell_value, text_fmt)

        # Auto column width
        for i, col in enumerate(report_df.columns):
            max_len = max(report_df[col].astype(str).map(len).max(), len(col)) + 2
            worksheet.set_column(i, i, max_len)

    buffer.seek(0)

    # Download button
    st.download_button(
        label="💾 Download Beautiful Excel Report",
        data=buffer,
        file_name=f"Complaint_Report_{selected_brand}.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )

else:
    st.info("No data available for the selected brand.")
