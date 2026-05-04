# Samsung Product Crawler

Project crawl dữ liệu sản phẩm từ API Samsung và lưu vào file Excel.

## Cài đặt

1. Cài đặt các dependencies:
```bash
pip install -r requirements.txt
```

## Sử dụng

Chạy script để crawl dữ liệu:
```bash
python crawl_samsung.py
```

Script sẽ:
1. Gọi API Samsung với pagination để lấy tất cả sản phẩm
2. Extract và map dữ liệu theo các cột đã định nghĩa
3. Lưu vào file `samsung_product_bulk.xlsx`

## Cấu trúc dữ liệu

Dữ liệu được lưu vào Excel với các cột sau:

- `name`: fmyMarketingName
- `description`: "" (để trống)
- `category_id`: 55
- `multi_categories`: "9, 55"
- `brand_id`: 24
- `video_provider`: "youtube"
- `video_link`: "youtube"
- `tags`: galleryImageAlt
- `unit_price`: priceDisplay
- `unit`: "Chiếc"
- `slug`: originPdpUrl (đã xử lý)
- `current_stock`: 100
- `est_shipping_days`: 5
- `sku`: modelList[0].modelCode
- `meta_title`: displayName
- `meta_description`: keySummary[0].description (nhiều mô tả cách nhau bởi dấu phẩy)
- `thumbnail_img`: modelList[0].thumbUrl (nhiều thumbnail cách nhau bởi dấu phẩy, đã thêm https:)
- `photos`: modelList[0].galleryImage (nhiều ảnh cách nhau bởi dấu phẩy, đã thêm https:)

## Lưu ý

- Script tự động xử lý pagination để lấy tất cả sản phẩm
- Có delay 1 giây giữa các request để tránh rate limiting
- Nếu file Excel đã tồn tại, dữ liệu mới sẽ được append vào cuối file
- Các URL ảnh sẽ tự động được thêm `https:` nếu thiếu

