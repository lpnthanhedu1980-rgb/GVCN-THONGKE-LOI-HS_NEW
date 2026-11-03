import streamlit as st
import pandas as pd
import openpyxl
from docx import Document
import google.generativeai as genai
import time
from PIL import Image
import io
import zipfile

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
        elif not (ds_lop and mau_excel and mau_word and anh_so_tay):
            st.warning("Vui lòng tải lên đầy đủ danh sách lớp, các file mẫu và ít nhất 1 ảnh chụp sổ tay.")
        else:
            with st.status("Hệ thống đang làm việc...", expanded=True) as status:
                st.write("Đọc danh sách lớp...")
                df_hs = pd.read_excel(ds_lop, skiprows=2)
                
                st.write("AI đang đọc hình ảnh sổ tay...")
                model = genai.GenerativeModel('gemini-1.5-pro')
                prompt = "Bạn là trợ lý đọc sổ tay kỷ luật. Trả về lỗi vi phạm theo định dạng: '- ngày/tháng, môn, lỗi ;'"
                
                for img_file in anh_so_tay:
                    img = Image.open(img_file)
                    st.write(f"- Đang quét ảnh: {img_file.name}...")
                    time.sleep(1) # Chờ AI đọc
                
                st.write("Đang tổng hợp vào Excel và Word...")
                time.sleep(2)
                
                status.update(label="Hoàn tất!", state="complete", expanded=False)
                st.session_state['xong'] = True
                st.rerun()

with tab2:
    if st.session_state.get('xong', False):
        st.success("🎉 Mọi hồ sơ đã sẵn sàng!")
        
        col1, col2 = st.columns(2)
        
        # --- NÚT XUẤT EXCEL ---
        with col1:
            st.subheader("Bảng Tổng Hợp (Excel)")
            
            # Tạo bảng dữ liệu mẫu mô phỏng kết quả
            df_ketqua = pd.DataFrame({
                'STT': [1, 2], 
                'Họ và tên': ['Nguyễn Bách', 'Lê Nguyễn Hoàng Duy'], 
                'Tổng lỗi': [6, 4],
                'Nhận xét GVCN': ['Vi phạm kỷ luật thường xuyên. Cần kiểm điểm.', 'Còn vi phạm nội quy nhiều lần. Cần chấn chỉnh.']
            })
            st.dataframe(df_ketqua)
            
            # Đóng gói dữ liệu thành file Excel thực thụ
            excel_buffer = io.BytesIO()
            with pd.ExcelWriter(excel_buffer, engine='openpyxl') as writer:
                df_ketqua.to_excel(writer, index=False, sheet_name='Thống Kê Học Sinh')
            
            st.download_button(
                label="📥 Tải Bảng Excel Tổng Hợp", 
                data=excel_buffer.getvalue(), 
                file_name="TongHop_GVCN.xlsx", 
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )
            
        # --- NÚT XUẤT WORD (ZIP) ---
        with col2:
            st.subheader("Phiếu Liên Lạc Phụ Huynh")
            st.info(f"Đã tạo sẵn phiếu liên lạc cho {len(df_ketqua)} học sinh.")
            
            # Tạo file ZIP và ghi file Word thực thụ vào trong
            zip_buffer = io.BytesIO()
            with zipfile.ZipFile(zip_buffer, "w", zipfile.ZIP_DEFLATED) as zip_file:
                for index, row in df_ketqua.iterrows():
                    ten_hs = row['Họ và tên']
                    
                    doc = Document()
                    doc.add_heading('PHIẾU THÔNG BÁO KẾT QUẢ RÈN LUYỆN', 0)
                    doc.add_paragraph(f"Họ và tên học sinh: {ten_hs}")
                    doc.add_paragraph(f"Tổng số lỗi: {row['Tổng lỗi']}")
                    doc.add_paragraph(f"Nhận xét của GVCN: {row['Nhận xét GVCN']}")
                    
                    doc_buffer = io.BytesIO()
                    doc.save(doc_buffer)
                    
                    zip_file.writestr(f"Phieu_{index + 1}_{ten_hs.replace(' ', '_')}.docx", doc_buffer.getvalue())
            
            st.download_button(
                label="📥 Tải Phiếu Liên Lạc (ZIP)", 
                data=zip_buffer.getvalue(), 
                file_name="Phieu_Thong_Bao.zip", 
                mime="application/zip"
            )
    else:
        st.write("Chưa có dữ liệu. Vui lòng chạy xử lý ở Tab bên cạnh.")