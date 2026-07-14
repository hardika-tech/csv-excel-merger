import gc
from io import BytesIO

import pandas as pd


# ----------------------------------------
# Date columns to format
# ----------------------------------------

DATE_COLUMNS = [
    "Attributed Touch Time",
    "Install Time",
    "Event Time"
]


# ----------------------------------------
# Read CSV safely
# ----------------------------------------

def read_csv_file(file):
    """
    Reads CSV using UTF-8 first.
    Falls back to latin1 if required.
    """

    try:
        file.seek(0)

        return pd.read_csv(
            file,
            low_memory=False
        )

    except UnicodeDecodeError:

        file.seek(0)

        return pd.read_csv(
            file,
            encoding="latin1",
            low_memory=False
        )


# ----------------------------------------
# Read Excel safely
# ----------------------------------------

def read_excel_file(file):
    """
    Reads XLSX files.
    """

    file.seek(0)

    return pd.read_excel(
        file,
        engine="openpyxl"
    )


# ----------------------------------------
# Main Merge Function
# ----------------------------------------

def merge_files(uploaded_files, progress_callback=None):
    """
    Merge uploaded CSV/XLSX files into
    a single Excel workbook.
    """

    if not uploaded_files:
        raise Exception("No files uploaded.")

    total_files = len(uploaded_files)
    total_rows = 0
    dataframes = []

    # ----------------------------------------
    # Read Every Uploaded File
    # ----------------------------------------

    for index, file in enumerate(uploaded_files):

        if progress_callback:
            progress = int(((index + 1) / total_files) * 100)

            progress_callback(
                progress,
                f"Processing {index + 1} of {total_files}: {file.name}"
            )

        try:

            if file.name.lower().endswith(".csv"):
                df = read_csv_file(file)

            elif file.name.lower().endswith(".xlsx"):
                df = read_excel_file(file)

            else:
                raise Exception(f"Unsupported file type: {file.name}")

        except Exception as e:

            raise Exception(
                f"Unable to read\n\n{file.name}\n\n{e}"
            )

        total_rows += len(df)
        dataframes.append(df)

    # ----------------------------------------
    # Merge All Files
    # ----------------------------------------

    print(">>> Before concat")

    if len(dataframes) == 1:
        merged_df = dataframes[0]
    else:
        merged_df = pd.concat(
            dataframes,
            ignore_index=True,
            sort=False,
            copy=False
        )

    print(">>> After concat")

    dataframes.clear()
    del dataframes
    gc.collect()

    # ----------------------------------------
    # Format Date Columns
    # ----------------------------------------

    for col in DATE_COLUMNS:

        if col in merged_df.columns:

            merged_df[col] = pd.to_datetime(
                merged_df[col],
                errors="coerce",
                utc=False
            )

    print(">>> After datetime formatting")

    # ----------------------------------------
    # Create Excel File (DEBUG VERSION)
    # ----------------------------------------

    print(">>> Before creating BytesIO")

    output = BytesIO()

    print(">>> Before ExcelWriter")

    with pd.ExcelWriter(
        output,
        engine="openpyxl"
    ) as writer:

        print(">>> Before to_excel")

        merged_df.to_excel(
            writer,
            sheet_name="Merged Data",
            index=False
        )

        print(">>> After to_excel")

    print(">>> After ExcelWriter")

    # ----------------------------------------
    # Prepare Summary
    # ----------------------------------------

    summary = {
        "files": total_files,
        "rows_imported": total_rows,
        "rows_output": len(merged_df)
    }

    del merged_df
    gc.collect()

    output.seek(0)

    return output, summary