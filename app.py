import streamlit as st
import pandas as pd
import zipfile
import io
import csv

st.set_page_config(
page_title="TXT ➜ Excel Splitter",
layout="wide"
)

st.title("📄 TXT ➜ Excel (Tách file Excel và nén ZIP)")

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

if uploaded_file:

```
st.info(
    "Hỗ trợ tự nhận diện dấu phân cách: | ; , TAB"
)

if st.button("🚀 Bắt đầu xử lý"):

    try:

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
        except Exception:
            delimiter = "|"

        st.success(
            f"Đã nhận diện delimiter: '{delimiter}'"
        )

        text_stream = io.TextIOWrapper(
            uploaded_file,
            encoding="utf-8",
            errors="ignore"
        )

        zip_buffer = io.BytesIO()

        progress = st.progress(0)
        status = st.empty()

        with zipfile.ZipFile(
            zip_buffer,
            mode="w",
            compression=zipfile.ZIP_DEFLATED
        ) as zipf:

            chunk = []
            file_index = 1
            total_rows = 0

            for line in text_stream:

                total_rows += 1

                row = line.rstrip("\n").split(delimiter)
                chunk.append(row)

                if total_rows % 5000 == 0:
                    status.text(
                        f"Đang xử lý: {total_rows:,} dòng"
                    )

                if len(chunk) >= rows_per_file:

                    max_cols = max(
                        len(r)
                        for r in chunk
                    )

                    columns = [
                        f"COL_{i+1}"
                        for i in range(max_cols)
                    ]

                    df = pd.DataFrame(
                        chunk,
                        columns=columns
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

            if chunk:

                max_cols = max(
                    len(r)
                    for r in chunk
                )

                columns = [
                    f"COL_{i+1}"
                    for i in range(max_cols)
                ]

                df = pd.DataFrame(
                    chunk,
                    columns=columns
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

        progress.progress(100)

        zip_buffer.seek(0)

        st.success(
            f"Hoàn thành! Đã xử lý {total_rows:,} dòng."
        )

        st.download_button(
            label="📦 Tải file ZIP",
            data=zip_buffer,
            file_name="excel_split.zip",
            mime="application/zip"
        )

    except Exception as e:
        st.error(f"Lỗi: {str(e)}")
```
