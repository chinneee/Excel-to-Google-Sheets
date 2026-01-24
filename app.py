import streamlit as st
import pandas as pd
import numpy as np
import gspread
from google.oauth2.service_account import Credentials
import json
from io import BytesIO
import time

# Page config
st.set_page_config(
    page_title="Excel → Google Sheets",
    page_icon="📊",
    layout="centered",
    initial_sidebar_state="collapsed"
)

# Enhanced Custom CSS
st.markdown("""
    <style>
    /* Main background */
    .stApp {
        background: linear-gradient(135deg, #667eea 0%, #764ba2 100%);
    }
    
    /* Main container */
    .main .block-container {
        padding-top: 2rem;
        padding-bottom: 2rem;
        max-width: 900px;
    }
    
    /* Custom card styling */
    .custom-card {
        background: white;
        border-radius: 20px;
        padding: 2.5rem;
        box-shadow: 0 20px 60px rgba(0,0,0,0.15);
        margin-bottom: 1.5rem;
        animation: slideUp 0.5s ease-out;
    }
    
    @keyframes slideUp {
        from {
            opacity: 0;
            transform: translateY(20px);
        }
        to {
            opacity: 1;
            transform: translateY(0);
        }
    }
    
    /* Header styling */
    .main-header {
        text-align: center;
        color: white;
        margin-bottom: 2rem;
        animation: fadeIn 0.8s ease-in;
    }
    
    .main-header h1 {
        font-size: 3rem;
        font-weight: 800;
        margin-bottom: 0.5rem;
        text-shadow: 2px 2px 4px rgba(0,0,0,0.2);
    }
    
    .main-header p {
        font-size: 1.1rem;
        opacity: 0.95;
    }
    
    @keyframes fadeIn {
        from { opacity: 0; }
        to { opacity: 1; }
    }
    

    /* Upload box */
    .upload-box {
        border: 3px dashed #667eea;
        border-radius: 15px;
        padding: 3rem 2rem;
        text-align: center;
        background: linear-gradient(135deg, #EEF2FF 0%, #E0E7FF 100%);
        transition: all 0.3s ease;
        cursor: pointer;
    }
    
    .upload-box:hover {
        border-color: #764ba2;
        background: linear-gradient(135deg, #E0E7FF 0%, #DDD6FE 100%);
        transform: translateY(-2px);
        box-shadow: 0 8px 20px rgba(102, 126, 234, 0.2);
    }
    
    /* Buttons */
    .stButton > button {
        border-radius: 10px;
        font-weight: 600;
        padding: 0.6rem 2rem;
        transition: all 0.3s ease;
        border: none;
    }
    
    .stButton > button:hover {
        transform: translateY(-2px);
        box-shadow: 0 5px 15px rgba(0,0,0,0.2);
    }
    
    /* File uploader */
    .stFileUploader {
        background: transparent;
    }
    
    /* Input fields */
    .stTextInput > div > div > input,
    .stNumberInput > div > div > input,
    .stTextArea > div > div > textarea {
        border-radius: 10px;
        border: 2px solid #E5E7EB;
        transition: all 0.3s ease;
    }
    
    .stTextInput > div > div > input:focus,
    .stNumberInput > div > div > input:focus,
    .stTextArea > div > div > textarea:focus {
        border-color: #667eea;
        box-shadow: 0 0 0 3px rgba(102, 126, 234, 0.1);
    }
    
    /* Success/Error messages */
    .stSuccess, .stError, .stInfo, .stWarning {
        border-radius: 10px;
        animation: slideIn 0.3s ease-out;
    }
    
    @keyframes slideIn {
        from {
            opacity: 0;
            transform: translateX(-20px);
        }
        to {
            opacity: 1;
            transform: translateX(0);
        }
    }
    
    /* Expander */
    .streamlit-expanderHeader {
        border-radius: 10px;
        background: linear-gradient(135deg, #F3F4F6 0%, #E5E7EB 100%);
        font-weight: 600;
    }
    
    /* Dataframe */
    .dataframe {
        border-radius: 10px;
        overflow: hidden;
    }
    
    /* Info boxes */
    .info-box {
        background: linear-gradient(135deg, #DBEAFE 0%, #BFDBFE 100%);
        border-left: 4px solid #3B82F6;
        border-radius: 10px;
        padding: 1rem 1.5rem;
        margin: 1rem 0;
    }
    
    .success-box {
        background: linear-gradient(135deg, #D1FAE5 0%, #A7F3D0 100%);
        border-left: 4px solid #10B981;
        border-radius: 10px;
        padding: 1rem 1.5rem;
        margin: 1rem 0;
    }
    
    .warning-box {
        background: linear-gradient(135deg, #FEF3C7 0%, #FDE68A 100%);
        border-left: 4px solid #F59E0B;
        border-radius: 10px;
        padding: 1rem 1.5rem;
        margin: 1rem 0;
    }
    
    /* Config display */
    .config-item {
        display: flex;
        justify-content: space-between;
        padding: 0.75rem 1rem;
        border-bottom: 1px solid #E5E7EB;
        transition: background 0.2s ease;
    }
    
    .config-item:hover {
        background: #F9FAFB;
    }
    
    .config-label {
        color: #6B7280;
        font-weight: 500;
    }
    
    .config-value {
        color: #111827;
        font-weight: 600;
    }
    
    /* Hide Streamlit branding */
    #MainMenu {visibility: hidden;}
    footer {visibility: hidden;}
    
    /* Spinner */
    .stSpinner > div {
        border-top-color: #667eea !important;
    }
    </style>
""", unsafe_allow_html=True)

# Header
st.markdown("""
    <div class="main-header">
        <h1>📊 Excel → Google Sheets</h1>
        <p>Upload file Excel và đẩy dữ liệu lên Google Sheets một cách dễ dàng</p>
    </div>
""", unsafe_allow_html=True)

# Initialize session state
if 'step' not in st.session_state:
    st.session_state.step = 1

# Progress Steps - Using Streamlit columns instead of HTML
st.markdown('<div class="custom-card">', unsafe_allow_html=True)

steps_data = [
    ("📤", "Upload"),
    ("⚙️", "Cấu hình"),
    ("✅", "Thực thi")
]

st.markdown("<br>", unsafe_allow_html=True)
cols = st.columns(3)
for i, (icon, label) in enumerate(steps_data):
    step_num = i + 1
    with cols[i]:
        if step_num < st.session_state.step:
            st.markdown(f"""
                <div style="text-align: center;">
                    <div style="width: 60px; height: 60px; border-radius: 50%; margin: 0 auto 0.5rem; display: flex; align-items: center; justify-content: center; font-size: 1.5rem; background: #10B981; color: white; border: 3px solid #10B981;">
                        {icon}
                    </div>
                    <div style="font-size: 0.9rem; font-weight: 600; color: #10B981;">{label}</div>
                </div>
            """, unsafe_allow_html=True)
        elif step_num == st.session_state.step:
            st.markdown(f"""
                <div style="text-align: center;">
                    <div style="width: 60px; height: 60px; border-radius: 50%; margin: 0 auto 0.5rem; display: flex; align-items: center; justify-content: center; font-size: 1.5rem; background: linear-gradient(135deg, #667eea 0%, #764ba2 100%); color: white; border: 3px solid #667eea; box-shadow: 0 4px 15px rgba(102, 126, 234, 0.4); transform: scale(1.1);">
                        {icon}
                    </div>
                    <div style="font-size: 0.9rem; font-weight: 600; color: #667eea;">{label}</div>
                </div>
            """, unsafe_allow_html=True)
        else:
            st.markdown(f"""
                <div style="text-align: center;">
                    <div style="width: 60px; height: 60px; border-radius: 50%; margin: 0 auto 0.5rem; display: flex; align-items: center; justify-content: center; font-size: 1.5rem; background: white; color: #9CA3AF; border: 3px solid #E5E7EB;">
                        {icon}
                    </div>
                    <div style="font-size: 0.9rem; font-weight: 600; color: #6B7280;">{label}</div>
                </div>
            """, unsafe_allow_html=True)

# Step 1: Upload Excel
if st.session_state.step == 1:
    st.markdown("### 📤 Bước 1: Upload File Excel")
    st.markdown("<br>", unsafe_allow_html=True)
    
    uploaded_file = st.file_uploader(
        "Chọn file Excel",
        type=['xlsx', 'xlsm', 'xls'],
        help="Hỗ trợ định dạng: .xlsx, .xlsm, .xls",
        label_visibility="collapsed"
    )
    
    if uploaded_file:
        st.markdown(f"""
            <div class="success-box">
                <strong>✅ Đã chọn file:</strong> {uploaded_file.name}<br>
                <small>Kích thước: {uploaded_file.size / 1024:.2f} KB</small>
            </div>
        """, unsafe_allow_html=True)
        
        # Preview
        with st.expander("👁️ Xem trước dữ liệu", expanded=False):
            try:
                df_preview = pd.read_excel(uploaded_file, sheet_name=0, nrows=10)
                st.dataframe(df_preview, use_container_width=True, height=300)
                st.caption(f"Hiển thị 10/{len(df_preview)} dòng đầu tiên")
            except Exception as e:
                st.error(f"⚠️ Không thể xem trước: {str(e)}")
        
        st.session_state.uploaded_file = uploaded_file
        
        st.markdown("<br>", unsafe_allow_html=True)
        col1, col2, col3 = st.columns([1, 2, 1])
        with col2:
            if st.button("Tiếp tục →", type="primary", use_container_width=True):
                st.session_state.step = 2
                st.rerun()

# Step 2: Configuration
elif st.session_state.step == 2:
    st.markdown("### ⚙️ Bước 2: Cấu hình")
    st.markdown("<br>", unsafe_allow_html=True)
    
    col1, col2 = st.columns(2)
    
    with col1:
        excel_sheet_name = st.text_input(
            "📄 Tên Sheet trong Excel",
            value="Template",
            help="Tên sheet cần đọc trong file Excel"
        )
        
        start_row = st.number_input(
            "📍 Dòng bắt đầu",
            min_value=1,
            value=7,
            help="Bắt đầu đọc từ dòng số mấy (bỏ qua header)"
        )
    
    with col2:
        worksheet_name = st.text_input(
            "📝 Tên Worksheet (Google Sheets)",
            value="Template",
            help="Tên worksheet đích trong Google Sheets"
        )
    
    gsheet_id = st.text_input(
        "🔗 Google Sheet ID",
        value="",
        help="Lấy từ URL: docs.google.com/spreadsheets/d/[SHEET_ID]/edit",
        placeholder="Ví dụ: 14g1NFdrmOFB_nyy74f5dNluV5bHG6_T9vS0_ez2Ao1s"
    )
    
    st.markdown("<br>", unsafe_allow_html=True)
    st.markdown("#### 🔑 Service Account Credentials")
    
    # Option to paste JSON or upload file
    auth_method = st.radio(
        "Chọn cách nhập credentials:",
        ["📋 Dán JSON", "📁 Upload file JSON"],
        horizontal=True
    )
    
    service_account_info = None
    
    if auth_method == "📋 Dán JSON":
        json_text = st.text_area(
            "Dán toàn bộ nội dung file JSON vào đây",
            height=150,
            help="Copy từ Google Cloud Console",
            placeholder='{\n  "type": "service_account",\n  "project_id": "...",\n  ...\n}'
        )
        if json_text:
            try:
                service_account_info = json.loads(json_text)
                st.success("✅ JSON hợp lệ!")
            except:
                st.error("❌ JSON không hợp lệ. Vui lòng kiểm tra lại format.")
    else:
        json_file = st.file_uploader(
            "Upload file credentials",
            type=['json'],
            help="File JSON từ Google Cloud Console"
        )
        if json_file:
            try:
                service_account_info = json.load(json_file)
                st.success(f"✅ Đã load: {json_file.name}")
            except:
                st.error("❌ File JSON không hợp lệ.")
    
    st.markdown("<br>", unsafe_allow_html=True)
    clear_before_append = st.checkbox(
        "🧹 Xóa dữ liệu cũ trước khi thêm dữ liệu mới",
        value=True,
        help="Xóa toàn bộ dữ liệu từ dòng bắt đầu trở đi"
    )
    
    # Info box
    st.markdown("""
        <div class="info-box">
            💡 <strong>Lưu ý:</strong> Service Account phải được share quyền <strong>Editor</strong> trên Google Sheet
        </div>
    """, unsafe_allow_html=True)
    
    # Save to session state
    st.session_state.config = {
        'excel_sheet_name': excel_sheet_name,
        'start_row': start_row,
        'gsheet_id': gsheet_id,
        'worksheet_name': worksheet_name,
        'service_account_info': service_account_info,
        'clear_before_append': clear_before_append
    }
    
    st.markdown("<br>", unsafe_allow_html=True)
    col1, col2, col3 = st.columns([1, 1, 1])
    with col1:
        if st.button("← Quay lại", use_container_width=True):
            st.session_state.step = 1
            st.rerun()
    with col3:
        can_proceed = (
            gsheet_id and 
            service_account_info is not None and
            hasattr(st.session_state, 'uploaded_file')
        )
        if st.button("Tiếp tục →", type="primary", disabled=not can_proceed, use_container_width=True):
            st.session_state.step = 3
            st.rerun()
        
        if not can_proceed:
            st.caption("⚠️ Vui lòng điền đầy đủ thông tin")

# Step 3: Execute
elif st.session_state.step == 3:
    st.markdown("### ✅ Bước 3: Xác nhận và Thực thi")
    st.markdown("<br>", unsafe_allow_html=True)
    
    # Display configuration
    config = st.session_state.config
    
    st.markdown("#### 📋 Thông tin cấu hình")
    
    config_html = f"""
    <div style="background: linear-gradient(135deg, #F9FAFB 0%, #F3F4F6 100%); border-radius: 12px; padding: 1rem; margin: 1rem 0;">
        <div class="config-item">
            <span class="config-label">📄 File Excel</span>
            <span class="config-value">{st.session_state.uploaded_file.name}</span>
        </div>
        <div class="config-item">
            <span class="config-label">📊 Sheet</span>
            <span class="config-value">{config['excel_sheet_name']}</span>
        </div>
        <div class="config-item">
            <span class="config-label">📍 Dòng bắt đầu</span>
            <span class="config-value">{config['start_row']}</span>
        </div>
        <div class="config-item">
            <span class="config-label">🔗 Google Sheet ID</span>
            <span class="config-value" style="font-size: 0.85em; font-family: monospace;">{config['gsheet_id'][:20]}...</span>
        </div>
        <div class="config-item">
            <span class="config-label">📝 Worksheet</span>
            <span class="config-value">{config['worksheet_name']}</span>
        </div>
        <div class="config-item" style="border-bottom: none;">
            <span class="config-label">🧹 Xóa dữ liệu cũ</span>
            <span class="config-value" style="color: {'#10B981' if config['clear_before_append'] else '#EF4444'};">
                {'✅ Có' if config['clear_before_append'] else '❌ Không'}
            </span>
        </div>
    </div>
    """
    st.markdown(config_html, unsafe_allow_html=True)
    
    st.markdown("<br>", unsafe_allow_html=True)
    col1, col2 = st.columns(2)
    
    with col1:
        if st.button("← Sửa cấu hình", use_container_width=True):
            st.session_state.step = 2
            st.rerun()
    
    with col2:
        if st.button("🚀 Bắt đầu đẩy dữ liệu", type="primary", use_container_width=True):
            
            progress_bar = st.progress(0)
            status_text = st.empty()
            
            try:
                # Read Excel
                status_text.markdown("### 📥 Đang đọc file Excel...")
                progress_bar.progress(10)
                time.sleep(0.5)
                
                df = pd.read_excel(
                    st.session_state.uploaded_file,
                    sheet_name=config['excel_sheet_name'],
                    engine="openpyxl",
                    skiprows=config['start_row'] - 1,
                    header=None
                )
                
                # Clean data
                df = df.dropna(how="all")
                df = df.replace([np.nan, np.inf, -np.inf], "")
                
                if df.empty:
                    st.warning("⚠️ Không tìm thấy dữ liệu trong file.")
                    st.stop()
                
                progress_bar.progress(30)
                st.success(f"✅ Đã đọc {df.shape[0]} dòng × {df.shape[1]} cột")
                time.sleep(0.3)
                
                # Authenticate
                status_text.markdown("### 🔐 Đang xác thực với Google...")
                progress_bar.progress(50)
                time.sleep(0.5)
                
                scopes = [
                    "https://www.googleapis.com/auth/spreadsheets",
                    "https://www.googleapis.com/auth/drive"
                ]
                
                creds = Credentials.from_service_account_info(
                    config['service_account_info'],
                    scopes=scopes
                )
                gc = gspread.authorize(creds)
                
                progress_bar.progress(60)
                st.success("✅ Xác thực thành công!")
                time.sleep(0.3)
                
                # Open worksheet
                status_text.markdown("### 📂 Đang mở Google Sheet...")
                progress_bar.progress(70)
                time.sleep(0.5)
                
                ws = gc.open_by_key(config['gsheet_id']).worksheet(config['worksheet_name'])
                
                progress_bar.progress(75)
                st.success("✅ Đã kết nối với Google Sheet!")
                time.sleep(0.3)
                
                # Clear if needed
                if config['clear_before_append']:
                    status_text.markdown(f"### 🧹 Đang xóa dữ liệu cũ từ dòng {config['start_row']}...")
                    progress_bar.progress(80)
                    time.sleep(0.5)
                    ws.batch_clear([f"A{config['start_row']}:ZZ"])
                    st.success("✅ Đã xóa dữ liệu cũ!")
                    time.sleep(0.3)
                
                # Append data
                status_text.markdown("### 🚀 Đang đẩy dữ liệu lên Google Sheets...")
                progress_bar.progress(90)
                time.sleep(0.5)
                
                ws.append_rows(
                    df.values.tolist(),
                    value_input_option="RAW",
                    table_range=f"A{config['start_row']}"
                )
                
                progress_bar.progress(100)
                status_text.empty()
                
                # Success message
                st.markdown("""
                    <div class="success-box" style="text-align: center; padding: 2rem;">
                        <h2 style="color: #10B981; margin: 0;">🎉 Hoàn thành!</h2>
                        <p style="font-size: 1.1rem; margin: 1rem 0;">
                            Đã đẩy <strong>{}</strong> dòng lên Google Sheets thành công!
                        </p>
                    </div>
                """.format(len(df)), unsafe_allow_html=True)
                
                st.balloons()
                
                st.markdown("<br>", unsafe_allow_html=True)
                col1, col2, col3 = st.columns([1, 2, 1])
                with col2:
                    if st.button("🔄 Upload file mới", use_container_width=True):
                        st.session_state.step = 1
                        if 'uploaded_file' in st.session_state:
                            del st.session_state.uploaded_file
                        st.rerun()
                
            except Exception as e:
                progress_bar.empty()
                status_text.empty()
                st.error(f"❌ **Đã xảy ra lỗi:**\n\n{str(e)}")
                
                with st.expander("🔍 Chi tiết lỗi"):
                    st.exception(e)

st.markdown('</div>', unsafe_allow_html=True)

# Footer
st.markdown("<br><br>", unsafe_allow_html=True)
st.markdown("""
    <div style='text-align: center; color: white; font-size: 0.9em; padding: 1rem; background: rgba(255,255,255,0.1); border-radius: 10px;'>
        💡 <strong>Lưu ý:</strong> Service Account cần có quyền <strong>Editor</strong> trên Google Sheet<br>
        <small>Tạo với ❤️ bằng Streamlit</small>
    </div>
""", unsafe_allow_html=True)
