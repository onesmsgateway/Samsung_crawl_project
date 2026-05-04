import json
from crawl_samsung_auto import SamsungCrawlerAuto

# Tạo crawler và lấy một response mẫu
crawler = SamsungCrawlerAuto()
crawler.setup_driver()
crawler.driver.get("https://www.samsung.com/vn/smartphones/all-smartphones/")
import time
time.sleep(5)

# Scroll một chút
crawler.driver.execute_script("window.scrollTo(0, 1000);")
time.sleep(3)

# Lấy responses
responses = crawler.extract_network_responses()

if responses:
    print("Tìm thấy response, đang phân tích...")
    response = responses[0]
    
    # Tìm product list
    product_list = None
    if 'response' in response and 'resultData' in response['response']:
        product_list = response['response']['resultData'].get('productList', [])
    elif 'resultData' in response:
        product_list = response['resultData'].get('productList', [])
    elif 'productList' in response:
        product_list = response['productList']
    
    if product_list and len(product_list) > 0:
        product = product_list[0]
        print("\n" + "="*60)
        print("CẤU TRÚC SẢN PHẨM ĐẦU TIÊN:")
        print("="*60)
        print(f"Keys ở root level: {list(product.keys())}")
        print(f"\npriceDisplay: {product.get('priceDisplay', 'KHÔNG CÓ')}")
        print(f"price: {product.get('price', 'KHÔNG CÓ')}")
        print(f"salePrice: {product.get('salePrice', 'KHÔNG CÓ')}")
        
        if 'modelList' in product and product['modelList']:
            model = product['modelList'][0]
            print(f"\nModel keys: {list(model.keys())}")
            print(f"Model priceDisplay: {model.get('priceDisplay', 'KHÔNG CÓ')}")
            print(f"Model price: {model.get('price', 'KHÔNG CÓ')}")
            print(f"Model salePrice: {model.get('salePrice', 'KHÔNG CÓ')}")
        
        print("\n" + "="*60)
        print("JSON đầy đủ (một phần):")
        print("="*60)
        print(json.dumps(product, indent=2, ensure_ascii=False)[:2000])
    
crawler.driver.quit()

