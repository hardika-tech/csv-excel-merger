import pandas as pd
from io import BytesIO


def merge_files(uploaded_files):

    dfs = []
    total_rows = 0

    for file in uploaded_files:

        try:

            if file.name.endswith(".csv"):

                try:
                    df = pd.read_csv(file)

                except:
                    file.seek(0)
                    df = pd.read_csv(file, encoding="latin1")

            else:

                df = pd.read_excel(file)

            dfs.append(df)
            total_rows += len(df)

        except Exception as e:
            raise Exception(f"Error reading {file.name}\n{e}")

    merged_df = pd.concat(
        dfs,
        ignore_index=True,
        sort=False
    )

    date_columns = [
        "Attributed Touch Time",
        "Install Time",
        "Event Time"
    ]

    for col in date_columns:

        if col in merged_df.columns:

            merged_df[col] = pd.to_datetime(
                merged_df[col],
                errors="coerce"
            )

    output = BytesIO()

    with pd.ExcelWriter(
        output,
        engine="openpyxl"
    ) as writer:

        merged_df.to_excel(
            writer,
            index=False
        )

    output.seek(0)

    return output, merged_df, total_rows