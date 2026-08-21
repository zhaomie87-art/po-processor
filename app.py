import streamlit as st
import tempfile, os, shutil
from process_po import process_folder

st.set_page_config(page_title="Xử Lý PO Tự Động", layout="centered")
st.title("Tự Động Điền PO Siêu Thị Vào Form BÁN")

template_file = st.file_uploader("1. Tải lên Form_Hàng_MT10.xlsx mẫu", type=["xlsx"])
pdf_files = st.file_uploader("2. Tải lên các file PO (PDF)", type=["pdf"], accept_multiple_files=True)

if st.button("Bắt đầu xử lý") and template_file and pdf_files:
    with st.spinner("Đang đọc PDF và đối chiếu dữ liệu..."):
        with tempfile.TemporaryDirectory() as tmpdir:
            # Lưu tạm file mẫu và PDF vào thư mục xử lý
            tmpl_path = os.path.join(tmpdir, "Form_Hàng_MT10.xlsx")
            with open(tmpl_path, "wb") as f:
                f.write(template_file.getbuffer())
            for pdf in pdf_files:
                with open(os.path.join(tmpdir, pdf.name), "wb") as f:
                    f.write(pdf.getbuffer())

            # Gọi hàm xử lý cốt lõi
            out_path = process_folder(tmpdir)

            with open(out_path, "rb") as f:
                st.success("Xử lý hoàn tất!")
                st.download_button(
                    label="Tải Về File Form_Hàng_MT10_đã_điền.xlsx",
                    data=f,
                    file_name="Form_Hàng_MT10_đã_điền.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                )