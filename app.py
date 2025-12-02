import io
import math

import pandas as pd
import streamlit as st


st.set_page_config(page_title="Tách Excel thành 3 file", layout="centered")

st.title("Tách dữ liệu Excel thành 3 file (Streamlit)")
st.write(
    "Ứng dụng này sẽ:\n"
    "- Chỉ lấy **22 dòng đầu** trong file Excel (1 dòng tiêu đề + 21 dòng dữ liệu).\n"
    "- Chia **21 dòng dữ liệu** đó thành **3 phần**.\n"
    "- Tạo 3 file Excel tương ứng: `1.xlsx`, `2.xlsx`, `3.xlsx` (mỗi file có lặp lại dòng tiêu đề)."
)


uploaded_file = st.file_uploader(
    "Chọn file Excel (ví dụ: thunghiem.xlsx)", type=["xlsx", "xls"]
)


def split_dataframe_into_three(df: pd.DataFrame):
    """
    Nhận vào DataFrame đã cắt còn tối đa 21 dòng dữ liệu,
    chia thành tối đa 3 phần và trả về list các DataFrame con.
    """
    max_data_rows = 21
    total_rows = min(len(df), max_data_rows)
    df_limit = df.iloc[:total_rows]  # chỉ lấy tối đa 21 dòng dữ liệu

    if total_rows == 0:
        return []

    rows_per_part = math.ceil(total_rows / 3)

    parts = []
    for i in range(3):
        start_row = i * rows_per_part
        end_row = min((i + 1) * rows_per_part, total_rows)
        if start_row >= total_rows:
            break
        parts.append(df_limit.iloc[start_row:end_row])

    return parts


def dataframe_to_excel_bytes(df: pd.DataFrame) -> bytes:
    """
    Ghi DataFrame ra file Excel trong bộ nhớ và trả về bytes,
    để dùng cho st.download_button.
    """
    buffer = io.BytesIO()
    with pd.ExcelWriter(buffer, engine="openpyxl") as writer:
        df.to_excel(writer, index=False)
    buffer.seek(0)
    return buffer.read()


if uploaded_file is not None:
    try:
        # Đọc file, dòng đầu tiên là tiêu đề (header)
        df = pd.read_excel(uploaded_file)

        st.success(f"Đã đọc file, tổng số dòng dữ liệu (không tính tiêu đề): {len(df)}")

        parts = split_dataframe_into_three(df)

        if not parts:
            st.warning("Không có dữ liệu (hoặc file không có dòng nào sau tiêu đề).")
        else:
            st.subheader("Tải về các file đã tách")
            for idx, part_df in enumerate(parts, start=1):
                file_name = f"{idx}.xlsx"
                excel_bytes = dataframe_to_excel_bytes(part_df)

                st.write(f"**File {file_name}** – số dòng dữ liệu: {len(part_df)}")
                st.download_button(
                    label=f"Tải {file_name}",
                    data=excel_bytes,
                    file_name=file_name,
                    mime=(
                        "application/vnd.openxmlformats-officedocument."
                        "spreadsheetml.sheet"
                    ),
                    key=f"download_{idx}",
                )

    except Exception as e:
        st.error(f"Có lỗi khi đọc hoặc xử lý file Excel: {e}")


