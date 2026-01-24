# 📊 Excel to Google Sheets - Streamlit App

Ứng dụng web để upload file Excel và tự động đẩy dữ liệu lên Google Sheets.

## 🚀 Deploy lên Streamlit Cloud

### Bước 1: Chuẩn bị GitHub Repository

1. Tạo repository mới trên GitHub
2. Upload 2 file:
   - `app.py` (code chính)
   - `requirements.txt` (dependencies)

### Bước 2: Deploy trên Streamlit Cloud

1. Truy cập: https://share.streamlit.io/
2. Đăng nhập bằng GitHub
3. Click **"New app"**
4. Chọn repository vừa tạo
5. Main file path: `app.py`
6. Click **Deploy**

### Bước 3: Cấu hình Google Service Account

1. Truy cập [Google Cloud Console](https://console.cloud.google.com/)
2. Tạo project mới hoặc chọn project có sẵn
3. Bật **Google Sheets API** và **Google Drive API**
4. Tạo **Service Account**:
   - IAM & Admin → Service Accounts → Create Service Account
   - Nhập tên, click Create
   - Bỏ qua Grant access
   - Click Done
5. Tạo JSON key:
   - Click vào Service Account vừa tạo
   - Keys → Add Key → Create New Key → JSON
   - Download file JSON

6. **Chia sẻ Google Sheet** với Service Account:
   - Mở file JSON, copy email (dạng: `xxx@xxx.iam.gserviceaccount.com`)
   - Mở Google Sheet cần đẩy dữ liệu
   - Click **Share** → Paste email → Cho quyền **Editor**

## 📝 Cách sử dụng

### Bước 1: Upload Excel
- Chọn file Excel (.xlsx, .xlsm, .xls)
- Xem trước dữ liệu

### Bước 2: Cấu hình
- **Tên Sheet**: Tên sheet trong file Excel (mặc định: "Template")
- **Dòng bắt đầu**: Đọc từ dòng số mấy (mặc định: 7)
- **Google Sheet ID**: Lấy từ URL
  ```
  https://docs.google.com/spreadsheets/d/[SHEET_ID]/edit
  ```
- **Worksheet name**: Tên worksheet trong Google Sheets
- **Service Account JSON**: 
  - Cách 1: Dán nội dung JSON
  - Cách 2: Upload file JSON
- **Xóa dữ liệu cũ**: Tích nếu muốn xóa dữ liệu từ dòng bắt đầu trở đi

### Bước 3: Thực thi
- Xem lại thông tin cấu hình
- Click "Bắt đầu đẩy dữ liệu"
- Chờ hoàn thành

## ⚙️ Tính năng

✅ Upload file Excel (.xlsx, .xlsm, .xls)  
✅ Tự động xóa NaN, Inf trong dữ liệu  
✅ Xóa dữ liệu cũ trước khi thêm (tuỳ chọn)  
✅ Giao diện thân thiện, dễ sử dụng  
✅ Progress indicator cho từng bước  
✅ Xử lý lỗi chi tiết  

## 🔒 Bảo mật

- **KHÔNG** commit file JSON credentials lên GitHub
- Sử dụng Streamlit Secrets để lưu credentials an toàn:
  - Settings → Secrets → Add credentials

## 🛠️ Chạy local (Development)

```bash
# Clone repository
git clone https://github.com/your-username/your-repo.git
cd your-repo

# Cài đặt dependencies
pip install -r requirements.txt

# Chạy app
streamlit run app.py
```

## 📦 Dependencies

- `streamlit` - Web framework
- `pandas` - Xử lý Excel
- `numpy` - Xử lý số liệu
- `gspread` - Google Sheets API
- `google-auth` - Xác thực Google
- `openpyxl` - Đọc Excel

## 💡 Tips

1. **Google Sheet ID**: Lấy từ URL giữa `/d/` và `/edit`
2. **Service Account**: Phải có quyền Editor trên Sheet
3. **Dòng bắt đầu**: Thường là dòng đầu tiên có dữ liệu (bỏ qua header)
4. **Clear data**: Nên bật để tránh duplicate dữ liệu

## 🐛 Xử lý lỗi thường gặp

### Lỗi: "Permission denied"
→ Chưa share Google Sheet với Service Account email

### Lỗi: "Sheet not found"
→ Kiểm tra lại tên worksheet

### Lỗi: "Invalid credentials"
→ File JSON không đúng hoặc đã bị revoke

### Lỗi: "API not enabled"
→ Chưa bật Google Sheets API trong Google Cloud Console

## 📞 Support

Nếu gặp vấn đề, hãy:
1. Check lại tất cả cấu hình
2. Xem error message chi tiết
3. Kiểm tra quyền truy cập Google Sheet

