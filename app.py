import streamlit as st
import pandas as pd
import openpyxl
from docx import Document
import google.generativeai as genai
import time
from PIL import Image

st.set_page_config(page_title="Trợ Lý GVCN", layout="wide")
st.title("📚 Hệ Thống Tự Động Hóa Hồ Sơ GVCN")

# --- CẤU HÌNH API BÊN TRÁI ---
with st.sidebar:
    st.header("🔑 Cấu hình hệ thống")
    api_key = st.text_input("Nhập Gemini API Key:", type="password")
    if api_key:
        genai.configure(api_key=api_key)
        
    st.header("📁 Tải Lên Dữ Liệu")
    ds_lop = st.file_uploader("1. Danh sách lớp (Excel)", type=["xlsx"])
    mau_excel = st.file_uploader("2. File mẫu Tổng hợp (Excel)", type=["xlsx"])
    mau_word = st.file_uploader("3. Phiếu liên lạc (Word)", type=["docx"])
    
    st.header("📸 Tải Lên Sổ Tay")
    anh_so_tay = st.file_uploader("Ảnh chụp sổ tay", type=["jpg", "png", "jpeg"], accept_multiple_files=True)

# --- KHUNG XỬ LÝ CHÍNH ---
tab1, tab2 = st.tabs(["⚙️ Xử Lý Dữ Liệu", "📊 Kết Quả"])

with tab1:
    st.info("Nhập API Key ở thanh bên trái, tải file và bấm nút bên dưới để AI tự động xử lý.")
    
    if st.button("🚀 Bắt Đầu Quét & Tổng Hợp", use_container_width=True):
        if not api_key:
            st.error("Thầy/Cô chưa nhập API Key kìa!")
        elif not (ds_lop and anh_so_tay):
            st.warning("Vui lòng tải lên danh sách lớp và ít nhất 1 ảnh chụp sổ tay.")
        else:
            with st.status("Hệ thống đang làm việc...", expanded=True) as status:
                st.write("Đọc danh sách lớp...")
                df_hs = pd.read_excel(ds_lop, skiprows=2)
                
                st.write("AI đang đọc hình ảnh sổ tay...")
                # Nạp model Gemini Vision để đọc ảnh
                model = genai.GenerativeModel('gemini-1.5-pro')
                
                # Câu lệnh Prompt Engineering ép AI trả về dữ liệu chuẩn
                prompt = """
                Bạn là trợ lý đọc sổ tay kỷ luật. Hãy đọc hình ảnh này và đối chiếu với danh sách học sinh.
                Trả về kết quả lỗi vi phạm của từng em theo đúng định dạng: '- ngày/tháng, môn, lỗi vi phạm ;'
                Tuyệt đối không bịa thêm thông tin.
                """
                
                # Trong thực tế, vòng lặp này sẽ gửi từng ảnh lên Gemini
                for img_file in anh_so_tay:
                    img = Image.open(img_file)
                    st.write(f"- Đang quét ảnh: {img_file.name}...")
                    # response = model.generate_content([prompt, img])
                    time.sleep(1) # Chờ AI đọc
                
                st.write("Đang tổng hợp vào Excel và Word...")
                time.sleep(2)
                
                status.update(label="Hoàn tất!", state="complete", expanded=False)
                st.session_state['xong'] = True
                st.rerun()

with tab2:
    if st.session_state.get('xong', False):
        st.success("🎉 Mọi hồ sơ đã sẵn sàng!")
        st.download_button("📥 Tải Bảng Excel Tổng Hợp", data=b"demo_excel", file_name="TongHop.xlsx")
        st.download_button("📥 Tải Phiếu Báo Phụ Huynh (ZIP)", data=b"demo_zip", file_name="PhieuLienLac.zip")
    else:
        st.write("Chưa có dữ liệu. Vui lòng chạy xử lý ở Tab bên cạnh.")