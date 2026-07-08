import streamlit as st
from merger import merge_files

st.set_page_config(
    page_title="CSV & Excel Merger",
    page_icon="📁",
    layout="wide"
)

st.title("📁 CSV & Excel Merger")

st.write(
    "Merge multiple CSV and Excel files into a single Excel workbook."
)

uploaded_files = st.file_uploader(
    "Choose CSV or Excel files",
    type=["csv", "xlsx"],
    accept_multiple_files=True
)

if uploaded_files:

    try:

        output, merged_df, total_rows = merge_files(uploaded_files)

        st.success("✅ Files merged successfully!")

        col1, col2, col3 = st.columns(3)

        col1.metric(
            "Files Uploaded",
            len(uploaded_files)
        )

        col2.metric(
            "Rows Imported",
            total_rows
        )

        col3.metric(
            "Rows Output",
            len(merged_df)
        )

        st.subheader("Preview")

        st.dataframe(
            merged_df.head(50),
            use_container_width=True
        )

        st.download_button(
            "⬇ Download Merged Excel",
            output,
            file_name="Merged_File.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )

    except Exception as e:

        st.error(str(e))

        print("hello")