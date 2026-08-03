# Monthly Domain Timesheet Generator App 🚀

Ứng dụng và công cụ dòng lệnh (CLI) tự động khởi tạo file Excel Timesheet hàng tháng cho 5 Domain:
1. **CBG CRM-OM Domain** (`FPT_CBG CRM-OM Domain_Timesheet_May2026.xlsx`)
2. **EBG CRM-OM Domain** (`FPT_EBG_CRM-OM_Timesheet_May2026.xlsx`)
3. **Identity Domain** (`FPT_Identity Domain_Timesheet_May2026.xlsx`)
4. **Provisioning Domain** (`FPT_Provisioning Domain_Timesheet_May2026.xlsx`)
5. **RWFM Domain** (`FPT_RWFM Domain_Timesheet_May2026.xlsx`)

---

## Tính Năng Nổi Bật

- ✅ **Bảo toàn 100% công thức Excel**: Giữ nguyên tất cả các hàm Excel (`SUMIFS`, `COUNTIFS`, `SUM`, `=ROW()-1`, `=ROW()-2`, `=Summary!K...`, `=G3-MAX(I3-H3,0)`).
- 🚫 **Loại bỏ ngày cuối tuần (Thứ 7 & Chủ nhật)**: Chỉ sinh công nhật cho các ngày làm việc tiêu chuẩn (Monday to Friday).
- 🇸🇬 **Tự động nhận diện Ngày lễ Singapore (Public Holidays)**: Tự động điền ngày lễ chuẩn của Singapore (Labour Day, Vesak Day, National Day, Deepavali, Christmas, ...) với `0h` và `Public Holiday`.
- 📅 **Cập nhật ngày tháng động**: Tự động cập nhật cột `Month`, `Date` và các tiêu đề ở sheet `Balance Leave` / `Leave Balance`.
- 💻 **Giao diện Web App Glassmorphism**: Chọn tháng, chọn domain, xem trước bảng công và tải xuống file Excel/ZIP.
- ⚡ **Công cụ CLI**: Chạy nhanh trực tiếp từ dòng lệnh.

---

## Cấu Trúc Thư Mục

```text
timesheet-generator-app/
├── app.py                     # Web App Server (Flask)
├── timesheet_generator.py     # Core engine xử lý Excel, công thức & ngày lễ
├── generate_timesheets_cli.py # Tool dòng lệnh (CLI)
├── requirements.txt           # Danh sách thư viện Python
├── README.md                  # Hướng dẫn sử dụng
├── timesheets_extracted/      # Thư mục chứa các file template Excel mẫu
│   ├── FPT_CBG CRM-OM Domain_Timesheet_May2026.xlsx
│   ├── FPT_EBG_CRM-OM_Timesheet_May2026.xlsx
│   ├── FPT_Identity Domain_Timesheet_May2026.xlsx
│   ├── FPT_Provisioning Domain_Timesheet_May2026.xlsx
│   └── FPT_RWFM Domain_Timesheet_May2026.xlsx
├── templates/                 # Giao diện HTML Web App
│   └── index.html
├── static/                    # Style CSS & JS frontend
│   ├── style.css
│   └── app.js
└── generated_timesheets/      # Thư mục chứa các tệp Excel đầu ra
```

---

## Hướng Dẫn Cài Đặt & Sử Dụng

### 1. Cài đặt thư viện
```bash
pip install -r requirements.txt
```

### 2. Chạy Giao diện Web App
```bash
python app.py
```
Truy cập trình duyệt tại địa chỉ: **[http://127.0.0.1:5050](http://127.0.0.1:5050)**

### 3. Chạy bằng Dòng lệnh CLI
Tạo Timesheet cho tháng 06/2026 cho tất cả các domain:
```bash
python generate_timesheets_cli.py --month 2026-06 --all
```

Tạo Timesheet cho domain cụ thể:
```bash
python generate_timesheets_cli.py --month 2026-07 --domains provisioning cbg
```

Nếu không muốn tự động nạp ngày lễ Singapore:
```bash
python generate_timesheets_cli.py --month 2026-06 --all --no-sg-holidays
```
