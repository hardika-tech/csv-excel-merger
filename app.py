import time
import traceback

import streamlit as st

from merger import merge_files


# --------------------------------------------------
# Page Configuration
# --------------------------------------------------

st.set_page_config(
    page_title="CSV & Excel Merger",
    page_icon="📁",
    layout="wide"
)

# --------------------------------------------------
# Header
# --------------------------------------------------

st.title("📁 CSV & Excel Merger")

st.write(
    "Merge multiple CSV and Excel files into a single Excel workbook."
)

st.info(
    "💡 Tip: For best performance on Render Free, "
    "try to keep the total upload size below approximately 200 MB."
)

st.divider()

# --------------------------------------------------
# File Upload
# --------------------------------------------------

uploaded_files = st.file_uploader(
    "Choose CSV or Excel files",
    type=["csv", "xlsx"],
    accept_multiple_files=True
)

# --------------------------------------------------
# Main Processing
# --------------------------------------------------

if uploaded_files:

    st.subheader("Selected Files")

    total_size = 0

    for file in uploaded_files:

        size_mb = file.size / (1024 * 1024)

        total_size += file.size

        st.write(
            f"📄 **{file.name}**   `{size_mb:.2f} MB`"
        )

    st.write("")

    st.info(
        f"Total Files: **{len(uploaded_files)}** | "
        f"Total Size: **{total_size / (1024 * 1024):.2f} MB**"
    )

    st.divider()

    progress = st.progress(0)

    status = st.empty()

    start_time = time.time()

    try:

        with st.spinner("Processing files..."):

            st.write("✅ Step 1: Starting merge...")

            output, summary = merge_files(
                uploaded_files,
                progress_callback=lambda value, text: (
                    progress.progress(value),
                    status.markdown(
                        f"**{value}% Complete**  \n{text}"
                    )
                )
            )

            st.write("✅ Step 2: merge_files() completed")

            st.write(f"Output Type: {type(output)}")

            st.write(
                f"Output Size: {output.getbuffer().nbytes:,} bytes"
            )

        elapsed = time.time() - start_time

        progress.empty()
        status.empty()

        st.write("✅ Step 3: Merge finished")

        st.success("✅ Files merged successfully!")

        st.divider()

        col1, col2, col3, col4 = st.columns(4)

        col1.metric(
            "Files",
            summary["files"]
        )

        col2.metric(
            "Rows Imported",
            f"{summary['rows_imported']:,}"
        )

        col3.metric(
            "Rows Output",
            f"{summary['rows_output']:,}"
        )

        col4.metric(
            "Time",
            f"{elapsed:.1f} sec"
        )

        st.write("")

        st.write("✅ Step 4: Creating download button")

        st.download_button(
            label="⬇ Download Merged Excel",
            data=output,
            file_name="Merged_File.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            use_container_width=True
        )

        st.write("✅ Step 5: Download button created")

    except Exception as e:

        progress.empty()
        status.empty()

        st.error("❌ Merge Failed")

        st.write("## Exception")
        st.code(str(e))

        st.write("## Full Traceback")
        st.code(traceback.format_exc())

        # Also print to Render logs
        print(traceback.format_exc())

        # Re-raise so Render captures the error
        raise

# --------------------------------------------------
# Footer
# --------------------------------------------------

st.divider()

st.caption(
    "📁 CSV & Excel Merger • Internal Tool • Version 2.0"
)