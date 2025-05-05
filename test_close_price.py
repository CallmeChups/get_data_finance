import yfinance as yf
import pandas as pd
from datetime import datetime, timedelta
from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from bs4 import BeautifulSoup
import time

def setup_driver(headless=True):
    chrome_options = Options()
    chrome_options.add_argument("--user-data-dir=C:/Users/USER/AppData/Local/Google/Chrome/User Data")  # Thay YourUsername
    chrome_options.add_argument("--profile-directory=Profile 7")
    if headless:
        chrome_options.add_argument("--headless")
    chrome_options.add_argument("--user-agent=Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/135.0.7049.42 Safari/537.36")
    chrome_options.add_argument("--ignore-certificate-errors")
    chrome_options.add_argument("--disable-gpu")
    chrome_options.add_argument("--no-sandbox")
    chrome_options.add_argument("--disable-dev-shm-usage")
    
    chrome_driver_path = "C:/WebDrivers/chromedriver.exe"
    service = Service(chrome_driver_path)
    driver = webdriver.Chrome(service=service, options=chrome_options)
    return driver

def get_vietstock_close_price(symbol, max_retries=3):
    url = f"https://finance.vietstock.vn/{symbol}/phan-tich-ky-thuat.htm"
    
    for attempt in range(max_retries):
        driver = setup_driver()
        try:
            driver.get(url)
            # Đợi cho đến khi element có id="stockprice" xuất hiện
            WebDriverWait(driver, 10).until(
                EC.presence_of_element_located((By.ID, "stockprice"))
            )
            
            # Lấy giá đóng cửa và ngày
            soup = BeautifulSoup(driver.page_source, "html.parser")
            price_element = soup.find("h2", {"id": "stockprice"})
            if price_element and price_element.find("span", {"class": "price"}):
                price = price_element.find("span", {"class": "price"}).text.strip()
                # Lấy ngày hiện tại
                current_date = datetime.now().strftime("%Y-%m-%d")
                driver.quit()
                return price, current_date
            
        except Exception as e:
            print(f"Lỗi khi lấy giá đóng cửa cho {symbol} ở lần thử {attempt + 1}: {e}")
            
        finally:
            driver.quit()
            
    print(f"Không lấy được giá đóng cửa cho {symbol}")
    return None, None

def get_weekly_close_data(symbols, num_candles=60):
    # Tính khoảng thời gian dựa trên số cây nến (mỗi cây là 1 tuần)
    end_date = datetime.now()
    start_date = end_date - timedelta(weeks=num_candles)
    
    # Tạo dictionary để lưu dữ liệu
    data = {}
    
    for symbol in symbols:
        # Thử cả hai định dạng: symbol gốc và symbol + .VN
        possible_symbols = [symbol, f"{symbol}.VN"]
        found_data = False
        
        for yf_symbol in possible_symbols:
            try:
                # Tải dữ liệu từ yfinance
                stock = yf.Ticker(yf_symbol)
                df = stock.history(start=start_date, end=end_date, interval="1wk")
                
                if not df.empty:
                    # Lấy cột Close và lưu vào dictionary
                    data[symbol] = df['Close']
                    found_data = True
                    print(f"Tìm thấy dữ liệu cho {yf_symbol}")
                    break  # Thoát vòng lặp nếu tìm thấy dữ liệu
                else:
                    print(f"Không có dữ liệu cho {yf_symbol}")
                    
            except Exception as e:
                print(f"Lỗi khi tải dữ liệu cho {yf_symbol}: {e}")
        
        if not found_data:
            print(f"Không tìm thấy dữ liệu cho {symbol} dưới bất kỳ định dạng nào")
    
    # Chuyển dictionary thành DataFrame
    result = pd.DataFrame(data)
    return result

# Danh sách symbols của bạn
symbols = ["ACB", "BCM", "BID", "BVH", "CTG", "FPT", "GAS", "GVR", "HDB", "HPG", 
           "LPB", "MBB", "MSN", "MWG", "PLX", "SAB", "SHB", "SSB", "SSI", "STB", 
           "TCB", "TPB", "VCB", "VHM", "VIB", "VIC", "VJC", "VNM", "VPB", "VRE"]

# Gọi hàm để lấy dữ liệu (n cây nến)
df = get_weekly_close_data(symbols, num_candles=100)

# In kết quả
print(df)

# Lưu vào file CSV
df.to_csv("weekly_close_data.csv")

# Tạo file .txt với dữ liệu từ Vietstock
with open("last_close_prices.txt", "w", encoding="utf-8") as f:
    for symbol in symbols:
        price, date = get_vietstock_close_price(symbol)
        if price and date:
            f.write(f"{symbol}_{date}: {price}\n")