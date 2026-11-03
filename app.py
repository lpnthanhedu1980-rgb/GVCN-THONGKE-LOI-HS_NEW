import streamlit as st
import pandas as pd
import openpyxl
from docx import Document
import google.generativeai as genai
import time
from PIL import Image
import io
import json
from copy import deepcopy

st.set_page_config(page_title="Trợ Lý GVCN", layout="wide")

# --- UI ĐẦU TRANG ---
st.markdown("### 👨‍🏫 TRƯỜNG THCS ĐOÀN THỊ ĐIỂM - PH1")
st.markdown("#### GVTH: LÊ PHAN NGỌC THANH")
st.title("📚 GVCN NHẬN XÉT HS TRONG THÁNG")

# --- CẤU HÌNH BẢO MẬT BÊN TRÁI ---
with st.sidebar:
    st.header("🔑 Mở khóa hệ thống")
    ma_truy_cap = st.text_input("Nhập mã bảo vệ:", type="password")
    
    api_key = None
    if ma_truy_cap == "THANHVUIVE":
        try:
            api_key = st.secrets["GEMINI_API_KEY"]
            genai.configure(api_key=api_key)
            st.success("✅ Đã mở khóa thành công!")
        except Exception:
            st.error("⚠️ Chưa cài đặt API Key trong máy chủ Streamlit.")
    elif ma_truy_cap:
        st.error("❌ Sai mã bảo vệ!")
        
    st.header("📁 Tải Lên Dữ Liệu")
    ds_lop = st.file_uploader("1. Danh sách lớp (Excel)", type=["xlsx"])
    mau_excel = st.file_uploader("2. File mẫu Tổng hợp (Excel)", type=["xlsx"])
    mau_word = st.file_uploader("3. Phiếu liên lạc (Word)", type=["docx"])
    
    st.header("📸 Tải Lên Sổ Tay")
    anh_so_tay = st.file_uploader("Ảnh chụp sổ tay", type=["jpg", "png", "jpeg"], accept_multiple_files=True)

# --- KHUNG XỬ LÝ CHÍNH ---
tab1, tab2 = st.tabs(["⚙️ Xử Lý Dữ Liệu", "📊 Kết Quả"])

with tab1:
    st.info("Nhập mã bảo vệ, tải đầy đủ 4 mục bên trái và bấm nút bên dưới để hệ thống xử lý.")
    
    if st.button("🚀 Bắt Đầu Quét & Tổng Hợp", use_container_width=True):
        if not api_key:
            st.error("Vui lòng nhập đúng mã bảo vệ để mở khóa AI!")
        elif not (ds_lop and mau_excel and mau_word and anh_so_tay):
            st.warning("Vui lòng tải đủ Danh sách lớp, 2 file mẫu và ảnh sổ tay.")
        else:
            with st.status("Hệ thống đang xử lý...", expanded=True) as status:
                st.write("1. Đọc danh sách lớp...")
                try:
                    df_hs = pd.read_excel(ds_lop, skiprows=2)
                    if 'Họ và tên' not in df_hs.columns:
                        df_hs = pd.read_excel(ds_lop)
                except:
                    df_hs = pd.read_excel(ds_lop)
                
                danh_sach_ten = df_hs['Họ và tên'].dropna().tolist()
                st.write(f"Đã tìm thấy {len(danh_sach_ten)} học sinh.")
                
                st.write("2. 🧠 AI đang phân tích hình ảnh và đối chiếu danh sách (khoảng 1-2 phút)...")
                
                # --- SỬ DỤNG MODEL CHUẨN MỚI NHẤT ---
                model = genai.GenerativeModel('gemini-2.5-flash')
                
                prompt = """
Bạn là trợ lý đọc sổ tay kỷ luật và phong trào của giáo viên.
Đây là danh sách học sinh trong lớp: """ + str(danh_sach_ten) + """

Nhiệm vụ: 
Hãy đọc thật kỹ các bức ảnh chụp sổ tay, tìm tên học sinh và tổng hợp chính xác lỗi vi phạm, phong trào của TỪNG em có tên trong danh sách trên.

YÊU CẦU:
1. Nếu học sinh KHÔNG có lỗi hoặc KHÔNG tham gia phong trào, bắt buộc ghi "Không".
2. Nếu có lỗi, phải ghi đúng cấu trúc: "- ngày/tháng, môn, lỗi vi phạm ;"
3. Nếu có nhiều lỗi trong cùng một tuần, hãy nối chúng lại và xuống dòng bằng kí tự \\n.
4. Cấu trúc phong trào: "- Tham gia [Tên phong trào] ;"
5. ĐỐI VỚI NHỮNG HỌC SINH CÓ LỖI HOẶC CÓ PHONG TRÀO, GÁN ĐÚNG VÀO TUẦN TRONG ẢNH. Nếu không rõ tuần, gán vào Tuan5.

BẮT BUỘC TRẢ VỀ DUY NHẤT CHUỖI JSON ĐÚNG CẤU TRÚC SAU (không giải thích gì thêm):
{
  "ket_qua": [
    {
      "Ten": "Nguyễn Bách",
      "Tuan5": "Không",
      "Tuan6": "- 13/10, Nhạc, không đem thanh phách ;",
      "Tuan7": "Không",
      "Tuan8": "Không",
      "PhongTrao": "- Tham gia Thi điền kinh ;"
    }
  ]
}
"""
                contents = [prompt]
                for img_file in anh_so_tay:
                    contents.append(Image.open(img_file))
                
                try:
                    response = model.generate_content(contents)
                    text_resp = response.text.strip()
                    if text_resp.startswith("```json"):
                        text_resp = text_resp[7:-3].strip()
                    elif text_resp.startswith("```"):
                        text_resp = text_resp[3:-3].strip()
                        
                    data_ai = json.loads(text_resp)
                    dict_ai = {item.get('Ten', '').strip(): item for item in data_ai.get('ket_qua', [])}
                    st.write("✅ AI đã bóc tách xong dữ liệu thực tế!")
                except Exception as e:
                    st.error(f"Lỗi khi AI xử lý: {e}")
                    st.stop()
                
                st.write("3. Đang tự động điền dữ liệu vào File Excel mẫu và đếm lỗi...")
                wb = openpyxl.load_workbook(mau_excel)
                ws = wb.active
                
                processed_data = []
                
                for index, row in df_hs.iterrows():
                    if pd.isna(row.get('Họ và tên')):
                        continue
                        
                    r = index + 2 
                    ten_hs = str(row.get('Họ và tên')).strip()
                    stt = str(row.get('STT', index+1))
                    
                    info = dict_ai.get(ten_hs, {})
                    t5 = info.get('Tuan5', 'Không')
                    t6 = info.get('Tuan6', 'Không')
                    t7 = info.get('Tuan7', 'Không')
                    t8 = info.get('Tuan8', 'Không')
                    phong_trao = info.get('PhongTrao', 'Không')
                    
                    tong_loi = 0
                    for t in [t5, t6, t7, t8]:
                        if t and str(t).strip() != 'Không':
                            tong_loi += len([line for line in str(t).split(';') if line.strip() != ''])
                    
                    nhan_xet = ""
                    if tong_loi == 0:
                        nhan_xet = "Thực hiện rất tốt nội quy, nề nếp trường lớp."
                    elif tong_loi <= 2:
                        nhan_xet = "Có cố gắng trong học tập và rèn luyện. Cần khắc phục một vài vi phạm nhỏ."
                    elif tong_loi <= 5:
                        nhan_xet = "Còn vi phạm nội quy nhiều lần. Cần nghiêm túc chấn chỉnh lại nề nếp học tập."
                    else:
                        nhan_xet = "Vi phạm kỷ luật thường xuyên. Cần kiểm điểm và thay đổi thái độ học tập ngay."
                        
                    if phong_trao and str(phong_trao).strip() != 'Không':
                        nhan_xet += " Tích cực tham gia các phong trào."
                    
                    try:
                        ws.cell(row=r, column=3).value = ten_hs
                        ws.cell(row=r, column=4).value = t5 if t5 != "Không" else None
                        ws.cell(row=r, column=5).value = len([l for l in str(t5).split(';') if l.strip()]) if t5 != "Không" else 0
                        ws.cell(row=r, column=6).value = t6 if t6 != "Không" else None
                        ws.cell(row=r, column=7).value = len([l for l in str(t6).split(';') if l.strip()]) if t6 != "Không" else 0
                        ws.cell(row=r, column=8).value = t7 if t7 != "Không" else None
                        ws.cell(row=r, column=9).value = len([l for l in str(t7).split(';') if l.strip()]) if t7 != "Không" else 0
                        ws.cell(row=r, column=10).value = t8 if t8 != "Không" else None
                        ws.cell(row=r, column=12).value = len([l for l in str(t8).split(';') if l.strip()]) if t8 != "Không" else 0
                        ws.cell(row=r, column=13).value = tong_loi
                        ws.cell(row=r, column=14).value = phong_trao if phong_trao != "Không" else None
                        ws.cell(row=r, column=15).value = nhan_xet
                    except:
                        pass
                        
                    processed_data.append({
                        'STT': stt, 'Ten': ten_hs, 'T5': t5, 'T6': t6, 'T7': t7, 'T8': t8,
                        'Tong': tong_loi, 'PhongTrao': phong_trao, 'NhanXet': nhan_xet
                    })

                excel_buffer = io.BytesIO()
                wb.save(excel_buffer)
                st.session_state['excel_data'] = excel_buffer.getvalue()
                
                st.write("4. Đang thực hiện Mail Merge và gộp thành 1 file Word duy nhất...")
                
                master_doc = None
                
                for item in processed_data:
                    mau_word.seek(0)
                    doc = Document(mau_word)
                    
                    if len(doc.tables) > 0:
                        try:
                            doc.tables[0].cell(0,0).paragraphs[0].text = f"Họ và tên học sinh: {item['Ten']}"
                            doc.tables[0].cell(1,1).paragraphs[0].text = f"STT theo danh sách lớp: {item['STT']}"
                        except: pass
                                
                    if len(doc.tables) > 1:
                        t1 = doc.tables[1]
                        try:
                            t1.cell(1, 1).text = item['T5'] if item['T5'] != 'Không' else ""
                            t1.cell(2, 1).text = item['T6'] if item['T6'] != 'Không' else ""
                            t1.cell(3, 1).text = item['T7'] if item['T7'] != 'Không' else ""
                            t1.cell(4, 1).text = item['T8'] if item['T8'] != 'Không' else ""
                            t1.cell(5, 2).text = str(item['Tong'])
                        except: pass
                            
                    if len(doc.tables) > 2:
                        t2 = doc.tables[2]
                        try:
                            if item['PhongTrao'] != "Không":
                                t2.cell(1, 2).text = item['PhongTrao']
                                t2.cell(1, 3).text = "Đã tham gia"
                            else:
                                t2.cell(1, 2).text = "Không tham gia phong trào nào"
                                t2.cell(1, 3).text = ""
                        except: pass
                    
                    for p in doc.paragraphs:
                        if "1. Nhận xét:" in p.text:
                            p.text = ""
                            r1 = p.add_run("1. Nhận xét: ")
                            r1.bold = True
                            p.add_run(item['NhanXet'])
                            
                    if master_doc is None:
                        master_doc = doc
                    else:
                        master_doc.add_page_break()
                        body = doc.element.body
                        sectPr = body.find('.//{http://schemas.openxmlformats.org/wordprocessingml/2006/main}sectPr')
                        if sectPr is not None:
                            body.remove(sectPr)
                        for element in body:
                            master_doc.element.body.append(deepcopy(element))
                        
                master_buffer = io.BytesIO()
                master_doc.save(master_buffer)
                        
                st.session_state['word_data'] = master_buffer.getvalue()
                st.session_state['count'] = len(processed_data)
                
                status.update(label="Hoàn tất thành công!", state="complete", expanded=False)
                st.session_state['xong'] = True
                st.rerun()

with tab2:
    if st.session_state.get('xong', False):
        st.success(f"🎉 Đã xử lý thành công toàn bộ {st.session_state['count']} học sinh bằng AI!")
        col1, col2 = st.columns(2)
        with col1:
            st.subheader("Bảng Tổng Hợp (Excel)")
            st.download_button(
                label="📥 Tải Excel Tổng Hợp", 
                data=st.session_state['excel_data'], 
                file_name="TongHop_AI_Quet.xlsx", 
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )
        with col2:
            st.subheader("Phiếu Liên Lạc (Word Gộp)")
            st.download_button(
                label="📥 Tải Phiếu Liên Lạc (1 File Gộp)", 
                data=st.session_state['word_data'], 
                file_name="Phieu_Thong_Bao_Gop.docx", 
                mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
            )
    else:
        st.write("Vui lòng tải file và chạy xử lý ở Tab bên cạnh.")
