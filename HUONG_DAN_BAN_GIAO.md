# Hướng Dẫn Sử Dụng và Bàn Giao Tool Crawl Sản Phẩm Samsung

## 1. Môi Trường Yêu Cầu
- **Python**: Phiên bản 3.8 trở lên.
- **Trình duyệt**: Google Chrome (Công cụ sẽ tự động tải WebDriver tương thích thông qua `webdriver_manager`).

## 2. Cài Đặt
Bật Terminal (hoặc Command Prompt) tại thư mục chứa source code và làm theo các bước sau:

**Bước 1:** Khởi tạo môi trường ảo (Khuyên dùng):
```bash
python -m venv venv
source venv/bin/activate  # (Với macOS/Linux)
venv\Scripts\activate     # (Với Windows)
```

**Bước 2:** Cài đặt các thư viện cần thiết:
```bash
pip install -r requirements.txt
```

## 3. Cách Sử Dụng (File chính: `crawl_samsung_auto.py`)
Tool sử dụng `Selenium` kết hợp với việc bắt các API ẩn (Network Interception) trong quá trình cuộn trang để lấy toàn bộ dữ liệu (cả các dữ liệu ẩn) một cách tự động và không bị block bởi rate-limit.

Bạn có thể chạy tool với các tham số (arguments) khác nhau để lấy danh mục sản phẩm mong muốn. Nếu không truyền tham số, tool sẽ mặc định lấy **Smartphones**.

**Lệnh chạy theo từng danh mục:**

- **Smartphones (Mặc định)**:
  ```bash
  python crawl_samsung_auto.py
  ```
  *File lưu trữ: `samsung_product_bulk.xlsx`*

- **Tablets (Máy tính bảng)**:
  ```bash
  python crawl_samsung_auto.py tablets
  ```
  *File lưu trữ: `samsung_tab_product_bulk.xlsx`*

- **Watches (Đồng hồ thông minh)**:
  ```bash
  python crawl_samsung_auto.py watches
  ```
  *File lưu trữ: `samsung_watch_product_bulk.xlsx`*

- **Tai nghe (Audio/Buds)**:
  ```bash
  python crawl_samsung_auto.py audio
  # Hoặc: python crawl_samsung_auto.py buds
  ```
  *File lưu trữ: `samsung_buds_product_bulk.xlsx`*

- **Nhẫn thông minh (Rings)**:
  ```bash
  python crawl_samsung_auto.py rings
  ```
  *File lưu trữ: `samsung_ring_product_bulk.xlsx`*

- **Phụ kiện (Accessories)**:
  ```bash
  python crawl_samsung_auto.py accessories
  ```
  *File lưu trữ: `samsung_accessories_product_bulk.xlsx`*

## 4. Cấu Trúc Dữ Liệu
Các file Excel xuất ra sẽ bao gồm các cột sau để import trực tiếp vào hệ thống:
- `name`: Tên sản phẩm
- `description`: Mô tả ngắn gọn của sản phẩm
- `category_id`: 55 (mặc định)
- `multi_categories`: "9, 55"
- `brand_id`: 24
- `video_provider`, `video_link`: Thông tin video (youtube)
- `tags`: Dùng cho Alt ảnh hoặc tag SEO
- `unit_price`: Giá sản phẩm (Đã lọc tiền tệ VND và dấu phẩy, ưu tiên giá khuyến mãi `promotionPrice`, không có sẽ để `1`)
- `slug`: URL rút gọn lấy từ `originPdpUrl`
- `sku`: Mã model sản phẩm
- `meta_title`, `meta_description`: Phục vụ SEO
- `thumbnail_img`: Ảnh đại diện sản phẩm (đã bọc thành link https chuẩn)
- `photos`: Danh sách ảnh gallery (cách nhau bởi dấu phẩy)

## 5. Lưu Ý Kỹ Thuật
- Quá trình chạy sẽ mở một cửa sổ Chrome ẩn danh. Trình duyệt sẽ tự động cuộn trang dần dần để gọi API như một người dùng thật.
- Vui lòng **không đóng cửa sổ Chrome** trong lúc tool đang chạy. Tool sẽ tự động đóng trình duyệt khi crawl xong và ghi file Excel hoàn tất.
- Tool có cơ chế tự lọc các sản phẩm rác như *Pre-Registration eVoucher* hay *sản phẩm chỉ định.* 
- Quá trình log ra console sẽ cho biết tool đang scroll ở đâu, lấy được bao nhiêu sản phẩm hiện tại.
