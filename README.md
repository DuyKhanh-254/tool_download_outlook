# SharePoint & OneDrive PDF Downloader (Restricted View)

Công cụ hỗ trợ tải xuống file tài liệu PDF từ **SharePoint Online** và **OneDrive** (đặc biệt là các tenant giáo dục và doanh nghiệp như `ueteduvn-my.sharepoint.com`) khi tính năng tải xuống đã bị chặn (Download Disabled / View-Only mode).

Dự án cung cấp **2 giải pháp hoàn chỉnh**:
1. **Chrome Extension (Khuyên dùng)**: Tiện ích mở rộng chạy trực tiếp trên trình duyệt Chrome/Edge, tận dụng 100% phiên đăng nhập SSO Microsoft sẵn có, trích xuất nguyên bản từ thẻ `<canvas>` và xuất file PDF chỉ với 1 cú click.
2. **Python CLI Tool**: Script dòng lệnh Python tự động hóa bằng Selenium và Pillow dành cho các tác vụ tải hàng loạt hoặc chạy tự động hóa trên server.

---

## PHẦN 1: HƯỚNG DẪN CÀI ĐẶT & SỬ DỤNG CHROME EXTENSION (KHUYÊN DÙNG)

### 🌟 Tại sao nên dùng Chrome Extension?
- **Không cần cấu hình phức tạp:** Không cần cài WebDriver, không xung đột profile (`user data directory is in use`).
- **Không cần nhập mật khẩu hay OTP/MFA:** Tận dụng trực tiếp tài khoản trường/công ty mà bạn đã đăng nhập trên trình duyệt.
- **Chất lượng ảnh nguyên bản (Lossless):** Lấy trực tiếp dữ liệu đồ họa từ thẻ `<canvas>` của SharePoint viewer, không bị dính thanh công cụ hay thanh cuộn.
- **Tự động đóng gói PDF:** Tích hợp sẵn thư viện `jsPDF`, tự động cuộn qua từng trang, đóng gói và tải về máy.

---

### 🚀 Cách cài đặt Extension vào Chrome / Edge trong 3 bước:

1. **Mở trang quản lý tiện ích:**
   - Trên **Google Chrome**, nhập vào thanh địa chỉ: `chrome://extensions/`
   - Trên **Microsoft Edge**, nhập vào thanh địa chỉ: `edge://extensions/`

2. **Bật chế độ dành cho nhà phát triển (Developer mode):**
   - Bật công tắc **Developer mode** ở góc trên cùng bên phải màn hình.

3. **Nạp thư mục Extension:**
   - Nhấn nút **Load unpacked** (Tải tiện ích đã giải nén) ở góc trên bên trái.
   - Chọn thư mục:
     ```
     d:\tool-download-onedrive\extension
     ```
   - Extension **OneDrive & SharePoint PDF Downloader** sẽ xuất hiện ngay lập tức trên thanh công cụ!

---

### 📖 Cách sử dụng Extension:

1. **Cập nhật lại extension (nếu đã nạp trước đó):**
   - Vào `chrome://extensions/`, tìm **OneDrive & SharePoint PDF Downloader** và nhấn vào biểu tượng **Tải lại (biểu tượng mũi tên xoay tròn 🔄)**.
   - Quay lại tab SharePoint và nhấn **F5 (Reload lại trang)**.

2. **Tiến hành tải file:**
   - Mở file tài liệu PDF hoặc Slide bài giảng trên SharePoint / OneDrive.
   - Bạn có thể tải file bằng 1 trong 2 cách:
     - **Cách 1 (Nút nổi trực tiếp trên trang):** Nhìn góc dưới bên phải, bấm vào nút tròn màu xanh **📥 Tải PDF**. Bạn có thể chỉnh ô **"Từ [ 1 ] đến [ 88 ]"** rồi bấm **Bắt đầu tải PDF**.
     - **Cách 2 (Popup Extension):** Bấm vào icon extension trên thanh công cụ:
       - Bạn có thể nhập dải trang cần tải (ví dụ: Từ `1` đến `88`, hoặc để mặc định tích chọn **"Quét cuộn liên tục đến hết file"**).
       - Bấm **"Bắt đầu tải PDF ngay"**.
   - Extension sẽ tự động đưa tài liệu về trang 1, cuộn quét từng trang (kể cả khi SharePoint dùng cơ chế lazy-load trang), chờ canvas vẽ xong, ghép toàn bộ thành file PDF chuẩn nét và lưu vào máy tính của bạn!

---

## PHẦN 2: HƯỚNG DẪN SỬ DỤNG BẢN PYTHON CLI

Nếu bạn muốn chạy script tự động bằng Python:

### 1. Cài đặt các thư viện phụ thuộc:
```bash
pip install -r requirements.txt
```

### 2. Cú pháp dòng lệnh:
```bash
python onedrive_downloader.py --url "<SHAREPOINT_URL>" [TÙY_CHỌN]
```

### 3. Các tham số hỗ trợ:
| Tham số | Ý nghĩa | Mặc định |
|---|---|---|
| `--url`, `-u` | **(Bắt buộc)** URL của file SharePoint / OneDrive | Không |
| `--output`, `-o` | Đường dẫn file PDF đầu ra | Tên gốc của file |
| `--browser`, `-b` | Trình duyệt (`chrome` hoặc `firefox`) | `chrome` |
| `--profile-dir`, `-p` | Đường dẫn profile để giữ phiên đăng nhập | `None` |
| `--headless` | Chạy ẩn trình duyệt không mở cửa sổ | `False` |
| `--page-delay`, `-d` | Thời gian chờ render giữa các trang (giây) | `2.0` |
| `--keep-temp` | Giữ lại thư mục ảnh tạm | `False` |
| `--verbose`, `-v` | In log chi tiết phục vụ gỡ lỗi | `False` |

### 4. Các ví dụ mẫu:
- **Tải cơ bản:**
  ```bash
  python onedrive_downloader.py --url "https://ueteduvn-my.sharepoint.com/personal/user_uet_edu_vn/_layouts/15/onedrive.aspx?id=..."
  ```
- **Chỉ định tên file và tăng thời gian chờ:**
  ```bash
  python onedrive_downloader.py --url "https://..." --output "bai_giang.pdf" --page-delay 2.5
  ```
- **Chạy với Chrome Profile đã đăng nhập sẵn:**
  *(Lưu ý: Cần tắt các cửa sổ Chrome đang mở trước khi chạy lệnh có profile để tránh xung đột khóa file)*
  ```bash
  python onedrive_downloader.py --url "https://..." --profile-dir "C:\Users\<Tên_User>\AppData\Local\Google\Chrome\User Data"
  ```
- **Chạy ở chế độ debug và giữ ảnh tạm:**
  ```bash
  python onedrive_downloader.py --url "https://..." --verbose --keep-temp
  ```

---

## ⚙️ CẤU TRÚC DỰ ÁN
```
d:\tool-download-onedrive\
├── extension/                     # Toàn bộ mã nguồn Chrome Extension (Manifest V3)
│   ├── manifest.json              # Cấu hình Extension
│   ├── icons/                     # Icon các kích thước 16, 48, 128
│   ├── libs/
│   │   └── jspdf.umd.min.js       # Thư viện đóng gói PDF client-side
│   ├── popup/                     # Giao diện điều khiển Popup
│   │   ├── popup.html
│   │   ├── popup.css
│   │   └── popup.js
│   ├── content_scripts/           # Script trích xuất trực tiếp trên tab SharePoint
│   │   ├── content.js
│   │   └── content.css
│   └── background/
│       └── service-worker.js      # Background Worker quản lý download
├── onedrive_downloader.py         # Script Python CLI dự phòng
├── requirements.txt               # Danh sách thư viện Python
└── README.md                      # Hướng dẫn sử dụng
```

---

## 🔒 LƯU Ý VÀ BẢN QUYỀN
- Công cụ này chỉ được sử dụng cho mục đích cá nhân, học tập và nghiên cứu trên những tài liệu mà người dùng đã được cấp quyền truy cập xem hợp pháp.
- Người dùng tự chịu trách nhiệm về việc bảo mật và quyền sở hữu tài liệu tải về.
