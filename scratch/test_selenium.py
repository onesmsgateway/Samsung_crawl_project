
from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.chrome.options import Options
from webdriver_manager.chrome import ChromeDriverManager
import os

try:
    print("Attempting to install ChromeDriver...")
    driver_path = ChromeDriverManager().install()
    print(f"Driver installed at: {driver_path}")
    
    chrome_options = Options()
    chrome_options.add_argument('--headless') # Run headless to avoid GUI issues
    chrome_options.add_argument('--no-sandbox')
    chrome_options.add_argument('--disable-dev-shm-usage')
    
    print("Launching Chrome...")
    service = Service(driver_path)
    driver = webdriver.Chrome(service=service, options=chrome_options)
    print("Chrome launched successfully!")
    driver.get("https://www.google.com")
    print(f"Page title: {driver.title}")
    driver.quit()
except Exception as e:
    print(f"An error occurred: {e}")
    import traceback
    traceback.print_exc()
