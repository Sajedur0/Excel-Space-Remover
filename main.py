import streamlit as st
import pandas as pd
import io
import os

# Page Configuration
st.set_page_config(
    page_title="Excel Space Cleaner",
    page_icon="🧹",
    layout="wide",
    initial_sidebar_state="expanded"
)

# -------------------- Custom CSS --------------------
st.markdown("""
    <style>
        .main-header {
            font-size: 2.5rem;
            font-weight: bold;
            text-align: center;
            color: #1E3A5F;
            margin-bottom: 0.5rem;
        }
        .sub-header {
            text-align: center;
            color: #4A4A4A;
            font-size: 1.1rem;
            margin-bottom: 2rem;
        }
        .metric-card {
            background-color: #f0f2f6;
            border-radius: 12px;
            padding: 20px;
            text-align: center;
            box-shadow: 0 2px 8px rgba(0,0,0,0.1);
            margin-bottom: 1rem;
        }
        .file-uploader {
            border: 2px dashed #1E3A5F;
            border-radius: 12px;
            padding: 30px;
            text-align: center;
            background-color: #f8f9fa;
            margin-bottom: 2rem;
        }
        .stButton>button {
            width: 100%;
            border-radius: 8px;
            font-weight: bold;
        }
        .info-box {
            background-color: #e8f4e8;
            border-left: 4px solid #28a745;
            padding: 12px 16px;
            border-radius: 0 8px 8px 0;
            margin-bottom: 1rem;
        }
        .warning-box {
            background-color: #fff3cd;
            border-left: 4px solid #ffc107;
            padding: 12px 16px;
            border-radius: 0 8px 8px 0;
            margin-bottom: 1rem;
        }
        .footer {
            text-align: center;
            color: #999;
            font-size: 0.85rem;
            margin-top: 3rem;
            padding-top: 1rem;
            border-top: 1px solid #e0e0e0;
        }
    </style>
""", unsafe_allow_html=True)


# -------------------- Helper Functions --------------------
def detect_spaces(df):
    """Detect cells that have leading or trailing spaces."""
    space_mask = df.map(lambda x: str(x).startswith(' ') or str(x).endswith(' ') if pd.notna(x) else False)
    return space_mask


def count_spaces(df):
    """Count total cells with leading or trailing spaces."""
    space_mask = detect_spaces(df)
    total_spaces = space_mask.sum().sum()
    rows_affected = space_mask.any(axis=1).sum()
    cols_affected = space_mask.any(axis=0).sum()
    return int(total_spaces), int(rows_affected), int(cols_affected)


def highlight_spaces(val):
    """Highlight cells with leading or trailing spaces."""
    try:
        if pd.isna(val):
            return ''
        val_str = str(val)
        if val_str.startswith(' ') or val_str.endswith(' '):
            return 'background-color: #FFCCCC; color: black; font-weight: bold;'
        return ''
    except:
        return ''


def clean_spaces(df):
    """Remove leading and trailing spaces from all cells."""
    return df.map(lambda x: str(x).strip() if pd.notna(x) else x)


def convert_df_to_excel(df):
    """Convert DataFrame to bytes for Excel download."""
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        df.to_excel(writer, index=False, sheet_name='Cleaned_Data')
    return output.getvalue()


def validate_file(uploaded_file):
    """Validate the uploaded file."""
    if uploaded_file is None:
        return False, "No file uploaded."
    
    # Check file size (max 50 MB)
    max_size = 50 * 1024 * 1024
    if uploaded_file.size > max_size:
        return False, f"File size exceeds 50 MB limit. Please upload a smaller file."
    
    # Check file extension
    name = uploaded_file.name.lower()
    if not (name.endswith('.xlsx') or name.endswith('.xls') or name.endswith('.csv')):
        return False, "Invalid file type. Please upload an .xlsx, .xls, or .csv file."
    
    return True, "File is valid."


# -------------------- Main App --------------------
def main():
    # Header
    st.markdown('<div class="main-header">🧹 Excel Space Cleaner</div>', unsafe_allow_html=True)
    st.markdown('<div class="sub-header">Upload your Excel file to detect and remove unwanted leading & trailing spaces.</div>', unsafe_allow_html=True)

    # -------------------- STEP 1: File Upload --------------------
    st.markdown("### 📂 Step 1: Load Your File")
    
    upload_col1, upload_col2 = st.columns([2, 1])
    
    with upload_col1:
        file_type = st.selectbox(
            "Select file source:",
            options=["Upload from Computer", "Load from URL", "Use Sample Data"],
            help="Choose how you want to load your Excel file."
        )
    
    with upload_col2:
        sheet_option = st.checkbox("Select specific sheet", value=False, help="If checked, you can pick which sheet to load.")

    df = None
    file_name = ""

    if file_type == "Upload from Computer":
        st.markdown('<div class="file-uploader">', unsafe_allow_html=True)
        uploaded_file = st.file_uploader(
            "Drag & drop or click to browse your Excel/CSV file",
            type=["xlsx", "xls", "csv"],
            help="Supported formats: .xlsx, .xls, .csv"
        )
        st.markdown('</div>', unsafe_allow_html=True)

        if uploaded_file:
            valid, msg = validate_file(uploaded_file)
            if not valid:
                st.warning(f"⚠️ {msg}")
                return
            file_name = uploaded_file.name

            try:
                if uploaded_file.name.endswith('.csv'):
                    df = pd.read_csv(uploaded_file)
                else:
                    if sheet_option:
                        sheet_names = pd.ExcelFile(uploaded_file).sheet_names
                        selected_sheet = st.selectbox("Select Sheet", sheet_names)
                        df = pd.read_excel(uploaded_file, sheet_name=selected_sheet)
                    else:
                        df = pd.read_excel(uploaded_file)
                
                df = df.astype(str)
                st.success(f"✅ File **'{file_name}'** loaded successfully!")
            except Exception as e:
                st.error(f"❌ Error loading file: {e}")
                return

    elif file_type == "Load from URL":
        url_input = st.text_input(
            "Enter the direct URL to your Excel/CSV file:",
            placeholder="https://example.com/your-file.xlsx",
            help="Provide a direct download link to an Excel or CSV file."
        )
        
        if st.button("🔗 Load from URL"):
            if url_input.strip() == "":
                st.warning("⚠️ Please enter a valid URL.")
                return
            try:
                file_name = url_input.split("/")[-1].split("?")[0]
                if file_name.endswith('.csv'):
                    df = pd.read_csv(url_input)
                else:
                    if sheet_option:
                        sheet_names = pd.ExcelFile(url_input).sheet_names
                        selected_sheet = st.selectbox("Select Sheet", sheet_names)
                        df = pd.read_excel(url_input, sheet_name=selected_sheet)
                    else:
                        df = pd.read_excel(url_input)
                
                df = df.astype(str)
                st.success(f"✅ File loaded successfully from URL!")
            except Exception as e:
                st.error(f"❌ Error loading file from URL: {e}")
                return

    elif file_type == "Use Sample Data":
        sample_data = {
            "Name": ["  John", "Jane  ", "  Bob  ", "Alice"],
            "City": ["New York  ", "  London", "Paris  ", "  Tokyo"],
            "Email": ["  john@example.com", "jane@example.com  ", " bob@example.com ", "alice@example.com"],
            "Score": ["  95", "88  ", "  92  ", "85"]
        }
        df = pd.DataFrame(sample_data).astype(str)
        file_name = "sample_data.xlsx"
        st.info("ℹ️ Loaded sample data for demonstration purposes.")

    # -------------------- STEP 2: Analysis --------------------
    if df is not None:
        st.markdown("---")
        st.markdown("### 🔍 Step 2: Space Analysis")

        # Metrics
        total_spaces, rows_affected, cols_affected = count_spaces(df)

        col1, col2, col3, col4 = st.columns(4)
        col1.metric("Total Cells Scanned", f"{df.shape[0] * df.shape[1]:,}")
        col2.metric("Cells with Spaces", f"{total_spaces:,}")
        col3.metric("Rows Affected", f"{rows_affected:,}")
        col4.metric("Columns Affected", f"{cols_affected:,}")

        # Space details
        if total_spaces > 0:
            with st.expander("📋 Click to see which cells have spaces"):
                space_mask = detect_spaces(df)
                detail_rows = []
                for i in range(df.shape[0]):
                    for j in range(df.shape[1]):
                        if space_mask.iloc[i, j]:
                            detail_rows.append({
                                "Row": i + 1,
                                "Column": df.columns[j],
                                "Value": repr(df.iloc[i, j]),
                                "Has Leading Space": str(df.iloc[i, j]).startswith(' '),
                                "Has Trailing Space": str(df.iloc[i, j]).endswith(' ')
                            })
                
                if detail_rows:
                    detail_df = pd.DataFrame(detail_rows)
                    st.dataframe(detail_df, use_container_width=True, height=min(len(detail_rows) * 35 + 40, 400))
                else:
                    st.info("No spaces detected.")

        # -------------------- STEP 3: Preview --------------------
        st.markdown("---")
        st.markdown("### 👁️ Step 3: Preview (Cells with spaces are highlighted in red)")

        styled_df = df.style.map(highlight_spaces)
        st.dataframe(styled_df, use_container_width=True, height=min(len(df) * 35 + 80, 500))

        # -------------------- STEP 4: Clean & Download --------------------
        st.markdown("---")
        st.markdown("### 🧹 Step 4: Clean Spaces")

        clean_col1, clean_col2 = st.columns(2)

        with clean_col1:
            trim_option = st.radio(
                "Select cleaning option:",
                options=[
                    "Remove leading & trailing spaces (Trim)",
                    "Remove leading spaces only (LTrim)",
                    "Remove trailing spaces only (RTrim)",
                    "Remove all spaces (including between words)"
                ],
                index=0
            )

        with clean_col2:
            output_format = st.selectbox(
                "Select output format:",
                options=["Excel (.xlsx)", "CSV (.csv)"],
                index=0
            )

        if st.button("🚀 Clean & Download", type="primary"):
            with st.spinner("Cleaning spaces..."):
                if "Trim" in trim_option:
                    if "Leading" in trim_option and "trailing" not in trim_option.lower():
                        cleaned_df = df.map(lambda x: str(x).lstrip() if pd.notna(x) else x)
                    elif "Trailing" in trim_option and "leading" not in trim_option.lower():
                        cleaned_df = df.map(lambda x: str(x).rstrip() if pd.notna(x) else x)
                    else:
                        cleaned_df = df.map(lambda x: str(x).strip() if pd.notna(x) else x)
                elif "all spaces" in trim_option.lower():
                    cleaned_df = df.map(lambda x: str(x).replace(' ', '') if pd.notna(x) else x)
                else:
                    cleaned_df = df

                st.success("✅ All spaces have been successfully removed!")

                # Show cleaned preview
                st.markdown("#### ✨ Cleaned Data Preview")
                st.dataframe(cleaned_df, use_container_width=True, height=min(len(cleaned_df) * 35 + 80, 500))

                # Provide download
                st.markdown("#### 📥 Download Your Cleaned File")
                
                if output_format == "Excel (.xlsx)":
                    excel_data = convert_df_to_excel(cleaned_df)
                    base_name = os.path.splitext(file_name)[0] if file_name else "cleaned_data"
                    st.download_button(
                        label="📥 Download Cleaned Excel File",
                        data=excel_data,
                        file_name=f"{base_name}_cleaned.xlsx",
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                        key="download_excel"
                    )
                else:
                    csv_data = cleaned_df.to_csv(index=False).encode('utf-8')
                    base_name = os.path.splitext(file_name)[0] if file_name else "cleaned_data"
                    st.download_button(
                        label="📥 Download Cleaned CSV File",
                        data=csv_data,
                        file_name=f"{base_name}_cleaned.csv",
                        mime="text/csv",
                        key="download_csv"
                    )

                # Summary
                new_spaces, _, _ = count_spaces(cleaned_df)
                st.info(f"📊 Summary: {total_spaces} spaces were found and {total_spaces - new_spaces} were removed. "
                        f"{new_spaces} spaces remaining (if any).")


    # -------------------- Footer --------------------
    st.markdown("---")
    st.markdown(
        '<div class="footer">Made with ❤️ by Excel Space Cleaner | Supports .xlsx, .xls, .csv formats</div>',
        unsafe_allow_html=True
    )


# -------------------- Run App --------------------
if __name__ == "__main__":
    main()