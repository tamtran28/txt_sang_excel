import streamlit as st
import pandas as pd
import zipfile
import io
import csv
import re  # Thêm thư viện re để lọc ký tự lỗi

st.set_page_config(page_title="TXT ➜ Excel Splitter", layout="wide")

st.title("📄 TXT ➜ Excel (Tách file + ZIP)")

uploaded_file = st.file_uploader("Chọn file TXT", type=["txt"])

rows_per_file = st.number_input(
    "Số dòng mỗi file Excel",
    min_value=1000,
    value=100000,
    step=1000
)

# Hàm loại bỏ ký tự không hợp lệ với Excel
# Giữ lại các ký tự in được, dấu xuống dòng (\n, \r) và tab (\t)
ILLEGAL_CHARACTERS_RE = re.compile(r'[\x00-\x08\x0b\x0c\x0e-\x1f\x7f-\xff]')

def clean_string(val):
    if isinstance(val, str):
        return ILLEGAL_CHARACTERS_RE.sub("", val)
    return val

if uploaded_file is not None:

    st.info("Hỗ trợ tự nhận diện dấu phân cách: | ; , TAB")

    if st.button("🚀 Bắt đầu xử lý"):

        # đọc sample để detect delimiter
        sample = uploaded_file.read(10240).decode("utf-8", errors="ignore")
        uploaded_file.seek(0)

        try:
            dialect = csv.Sniffer().sniff(sample, delimiters="|;,\t")
            delimiter = dialect.delimiter
        except:
            delimiter = "|"

        st.success(f"Delimiter nhận diện: '{delimiter}'")

        text_stream = io.TextIOWrapper(uploaded_file, encoding="utf-8", errors="ignore")

        zip_buffer = io.BytesIO()

        progress = st.progress(0)
        status = st.empty()

        with zipfile.ZipFile(zip_buffer, "w", zipfile.ZIP_DEFLATED) as zipf:

            chunk = []
            file_index = 1
            total_rows = 0

            for line in text_stream:

                total_rows += 1

                # Tách dòng và làm sạch từng phần tử ngay lập tức
                row = [clean_string(item) for item in line.strip().split(delimiter)]
                chunk.append(row)

                if total_rows % 5000 == 0:
                    status.text(f"Đang xử lý: {total_rows:,} dòng")

                if len(chunk) >= rows_per_file:

                    max_cols = max(len(r) for r in chunk)
                    cols = [f"COL_{i+1}" for i in range(max_cols)]

                    # Đảm bảo các dòng có đủ số cột bằng cách padding chuỗi rỗng
                    padded_chunk = [r + [""] * (max_cols - len(r)) for r in chunk]

                    df = pd.DataFrame(padded_chunk, columns=cols)

                    buffer = io.BytesIO()
                    with pd.ExcelWriter(buffer, engine="openpyxl") as writer:
                        df.to_excel(writer, index=False, sheet_name="Data")

                    zipf.writestr(f"output_{file_index}.xlsx", buffer.getvalue())

                    file_index += 1
                    chunk = []

            # xử lý phần còn lại cuối cùng
            if chunk:
                max_cols = max(len(r) for r in chunk)
                cols = [f"COL_{i+1}" for i in range(max_cols)]
                padded_chunk = [r + [""] * (max_cols - len(r)) for r in chunk]

                df = pd.DataFrame(padded_chunk, columns=cols)

                buffer = io.BytesIO()
                with pd.ExcelWriter(buffer, engine="openpyxl") as writer:
                    df.to_excel(writer, index=False, sheet_name="Data")

                zipf.writestr(f"output_{file_index}.xlsx", buffer.getvalue())

        zip_buffer.seek(0)
        progress.progress(100)
        st.success(f"Xong! Đã xử lý {total_rows:,} dòng.")

        st.download_button(
            "📦 Tải file ZIP",
            data=zip_buffer,
            file_name="excel_split.zip",
            mime="application/zip"
        )
