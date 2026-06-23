```python
import streamlit as st
import pandas as pd
import zipfile
import io
import csv

st.set_page_config(page_title="TXT → Excel Splitter", layout="wide")

st.title("📄 TXT ➜ Excel (Tự nhận diện cột & Tách file)")

uploaded_file = st.file_uploader(
    "Chọn file TXT",
    type=["txt"]
)

rows_per_file = st.number_input(
    "Số dòng mỗi file Excel",
    min_value=1000,
    value=100000,
    step=1000
)

if uploaded_file and st.button("Xử lý"):

    # Đọc vài dòng đầu để đoán delimiter
    sample = uploaded_file.read(10240).decode(
        "utf-8",
        errors="ignore"
    )

    uploaded_file.seek(0)

    try:
        dialect = csv.Sniffer().sniff(
            sample,
            delimiters="|;,\t"
        )
        delimiter = dialect.delimiter
    except:
        delimiter = "|"

    st.info(f"Delimiter nhận diện: '{delimiter}'")

    text_stream = io.TextIOWrapper(
        uploaded_file,
        encoding="utf-8",
        errors="ignore"
    )

    zip_buffer = io.BytesIO()

    with zipfile.ZipFile(
        zip_buffer,
        mode="w",
        compression=zipfile.ZIP_DEFLATED
    ) as zipf:

        chunk = []
        file_index = 1

        for line_no, line in enumerate(text_stream, start=1):

            row = line.rstrip("\n").split(delimiter)
            chunk.append(row)

            if len(chunk) >= rows_per_file:

                max_cols = max(len(r) for r in chunk)

                cols = [
                    f"COL_{i+1}"
                    for i in range(max_cols)
                ]

                df = pd.DataFrame(
                    chunk,
                    columns=cols
                )

                excel_buffer = io.BytesIO()

                with pd.ExcelWriter(
                    excel_buffer,
                    engine="openpyxl"
                ) as writer:

                    df.to_excel(
                        writer,
                        index=False,
                        sheet_name="Data"
                    )

                zipf.writestr(
                    f"output_{file_index}.xlsx",
                    excel_buffer.getvalue()
                )

                file_index += 1
                chunk = []

        # phần còn lại
        if chunk:

            max_cols = max(len(r) for r in chunk)

            cols = [
                f"COL_{i+1}"
                for i in range(max_cols)
            ]

            df = pd.DataFrame(
                chunk,
                columns=cols
            )

            excel_buffer = io.BytesIO()

            with pd.ExcelWriter(
                excel_buffer,
                engine="openpyxl"
            ) as writer:

                df.to_excel(
                    writer,
                    index=False,
                    sheet_name="Data"
                )

            zipf.writestr(
                f"output_{file_index}.xlsx",
                excel_buffer.getvalue()
            )

    zip_buffer.seek(0)

    st.success(
        f"Hoàn thành. Tạo {file_index} file Excel."
    )

    st.download_button(
        label="📦 Tải ZIP",
        data=zip_buffer,
        file_name="excel_split.zip",
        mime="application/zip"
    )
```
