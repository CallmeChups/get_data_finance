from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.common.by import By
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium_stealth import stealth
from bs4 import BeautifulSoup
from unidecode import unidecode
import re
import time
import pandas as pd
from icecream import ic
from datetime import datetime

# Dictionary để lưu trữ URL của các symbol đã tìm thấy
symbol_urls = {}

# Hàm khởi tạo Chrome driver chung
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
    chrome_options.add_argument("--disable-blink-features=AutomationControlled")
    chrome_options.add_experimental_option("excludeSwitches", ["enable-automation"])
    chrome_options.add_experimental_option('useAutomationExtension', False)

    chrome_driver_path = "C:/WebDrivers/chromedriver.exe"
    service = Service(chrome_driver_path)
    driver = webdriver.Chrome(service=service, options=chrome_options)
    driver.set_page_load_timeout(60)
    driver.set_script_timeout(60)

    # Áp dụng stealth để tránh phát hiện bot
    stealth(driver,
            languages=["en-US", "en"],
            vendor="Google Inc.",
            platform="Win32",
            webgl_vendor="Intel Inc.",
            renderer="Intel Iris OpenGL Engine",
            fix_hairline=True)
    return driver

# Hàm chuyển tên công ty thành slug
def to_slug(text):
    text = unidecode(text).lower()
    return re.sub(r'[^a-z0-9]+', '-', text).strip('-')

# Hàm lấy URL chính xác của công ty dựa trên symbol
def get_company_url(symbol, max_retries=3):
    if symbol in symbol_urls:
        ic(f"Đã tìm thấy URL cho {symbol} trong bộ nhớ: {symbol_urls[symbol]}")
        return symbol_urls[symbol]

    for attempt in range(max_retries):
        driver = setup_driver()
        try:
            driver.get("https://finance.vietstock.vn/doanh-nghiep-a-z")
            WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.ID, "az-container")))

            first_letter = symbol[0].upper()
            letter_button = WebDriverWait(driver, 10).until(
                EC.element_to_be_clickable((By.XPATH, f"//a[@class='title-link filter-char' and text()='{first_letter}']"))
            )
            driver.execute_script("arguments[0].click();", letter_button)
            ic(f"Đã nhấn nút '{first_letter}'")
            time.sleep(3)

            overview_url = None
            while True:
                soup = BeautifulSoup(driver.page_source, "html.parser")
                tbody = soup.select_one("#az-container table tbody")
                if not tbody:
                    break

                for row in tbody.find_all("tr"):
                    cells = row.find_all("td")
                    if len(cells) >= 3 and cells[1].text.strip() == symbol:
                        company_link = cells[1].find("a")
                        if company_link and "href" in company_link.attrs:
                            overview_url = company_link["href"]
                            break
                if overview_url:
                    break

                next_button = driver.find_elements(By.CSS_SELECTOR, "i.fa.fa-chevron-right")
                if not next_button or "disabled" in next_button[0].find_element(By.XPATH, "..").get_attribute("outerHTML"):
                    break
                driver.execute_script("arguments[0].click();", next_button[0])
                time.sleep(3)

            driver.quit()
            if overview_url:
                symbol_part = overview_url.split("/")[-1].split("-")[0]
                finance_url = f"https://finance.vietstock.vn/{symbol_part}/tai-chinh.htm?tab=BCTT"
                symbol_urls[symbol] = (overview_url, finance_url)
                ic(f"Đã lưu URL cho {symbol}: {symbol_urls[symbol]}")
                return overview_url, finance_url
            print(f"Không tìm thấy URL cho {symbol} ở lần thử {attempt + 1}")

        except Exception as e:
            print(f"Lỗi khi lấy URL cho {symbol} ở lần thử {attempt + 1}: {e}")
            driver.quit()
            continue

    print(f"Không lấy được URL cho {symbol} sau {max_retries} lần thử")
    return None, None

# Hàm lấy dữ liệu tài chính (EPS, ROE) từ trang Tổng quan
def get_financial_data(data_type, symbol, max_retries=3):
    if symbol not in symbol_urls:
        get_company_url(symbol)  # Đảm bảo URL đã được lấy trước
    if symbol not in symbol_urls:
        print(f"Không có URL cho {symbol}")
        return None
    url = symbol_urls[symbol][0]  # Sử dụng overview_url cho EPS và ROE

    driver = setup_driver()
    try:
        driver.get(url)
        WebDriverWait(driver, 15).until(
            EC.presence_of_element_located((By.ID, "table-2"))
        )
        driver.execute_script("window.scrollTo(0, document.body.scrollHeight);")
        time.sleep(2)

        # Lấy dữ liệu ban đầu (2021-2024)
        soup = BeautifulSoup(driver.page_source, "html.parser")
        table = soup.find("table", id="table-2")
        if not table:
            print(f"Không tìm thấy bảng 'table-2' cho {symbol}")
            driver.quit()
            return None

        rows = table.find("tbody").find_all("tr")
        data_2021_2024 = None
        for row in rows:
            cells = row.find_all("td")
            if cells and data_type in cells[0].text.strip():
                data_2021_2024 = [cell.text.strip() for cell in cells[1:5]]  # Lấy 4 năm 2021-2024
                ic(f"{data_type} của {symbol} (2021-2024): {data_2021_2024}")
                break

        if not data_2021_2024:
            print(f"Không tìm thấy {data_type} trong bảng 'table-2' cho {symbol}")
            driver.quit()
            return None

        # Nhấn nút "qua trái" để lấy dữ liệu 2017-2020, retry cho đến khi thành công
        left_button_clicked = False
        for attempt in range(max_retries):
            try:
                left_button = WebDriverWait(driver, 5).until(
                    EC.element_to_be_clickable((By.CSS_SELECTOR, "span.fa.fa-chevron-left.pull-left"))
                )
                driver.execute_script("arguments[0].click();", left_button)
                ic(f"Đã nhấn nút 'qua trái' cho {symbol} ở lần thử {attempt + 1}")
                time.sleep(3)  # Chờ dữ liệu load
                left_button_clicked = True
                break
            except Exception as e:
                print(f"Lỗi khi nhấn nút 'qua trái' cho {symbol} ở lần thử {attempt + 1}: {e}")
                time.sleep(2)

        if not left_button_clicked:
            ic(f"Không thể nhấn nút 'qua trái' sau {max_retries} lần thử, dùng dữ liệu hiện tại với 2020 = N/A")
            final_data = ["N/A"] + data_2021_2024
            driver.quit()
            ic(f"{data_type} của {symbol}: {final_data}")
            return final_data

        # Lấy dữ liệu sau khi nhấn "qua trái" (2017-2020)
        soup = BeautifulSoup(driver.page_source, "html.parser")
        table = soup.find("table", id="table-2")
        if not table:
            print(f"Không tìm thấy bảng 'table-2' sau khi nhấn 'qua trái' cho {symbol}")
            final_data = ["N/A"] + data_2021_2024
            driver.quit()
            ic(f"{data_type} của {symbol}: {final_data}")
            return final_data

        rows = table.find("tbody").find_all("tr")
        data_2017_2020 = None
        for row in rows:
            cells = row.find_all("td")
            if cells and data_type in cells[0].text.strip():
                data_2017_2020 = [cell.text.strip() for cell in cells[1:5]]  # Lấy 2017-2020
                ic(f"{data_type} của {symbol} (2017-2020): {data_2017_2020}")
                break

        if not data_2017_2020:
            print(f"Không tìm thấy {data_type} trong bảng 'table-2' sau khi nhấn 'qua trái' cho {symbol}")
            final_data = ["N/A"] + data_2021_2024
        else:
            final_data = [data_2017_2020[-1]] + data_2021_2024  # Lấy 2020 ghép với 2021-2024

        driver.quit()
        ic(f"{data_type} của {symbol}: {final_data}")
        return final_data

    except Exception as e:
        print(f"Lỗi khi lấy {data_type} cho {symbol}: {e}")
        driver.quit()
        return None

# Hàm lấy "KLGD khớp lệnh trung bình 10 phiên" từ cafef.vn
def get_avg_trading_volume(symbol, max_retries=3):
    for attempt in range(max_retries):
        driver = setup_driver(headless=True)
        try:
            ic(f"Bắt đầu lấy KLGD trung bình cho {symbol}, lần thử {attempt + 1}")
            driver.get("https://cafef.vn/du-lieu/cong-bo-thong-tin.chn")
            ic(f"Đã truy cập trang cong-bo-thong-tin.chn")

            input_field = WebDriverWait(driver, 10).until(
                EC.presence_of_element_located((By.ID, "acp-inp-disclosure"))
            )
            input_field.clear()
            input_field.send_keys(symbol)
            time.sleep(1)
            ic(f"Đã nhập mã {symbol} vào ô tìm kiếm")

            # Thử lại nếu dropdown không xuất hiện
            for _ in range(2):
                try:
                    dropdown_item = WebDriverWait(driver, 10).until(
                        EC.element_to_be_clickable((By.ID, "0yourself-transaction-search-stock-item"))
                    )
                    driver.execute_script("arguments[0].click();", dropdown_item)
                    ic(f"Đã chọn mục đầu tiên trong dropdown")
                    break
                except:
                    ic(f"Dropdown chưa xuất hiện, thử lại lần {_ + 1}")
                    time.sleep(2)

            search_button = WebDriverWait(driver, 10).until(
                EC.element_to_be_clickable((By.CLASS_NAME, "btn-disclosure"))
            )
            driver.execute_script("arguments[0].click();", search_button)
            time.sleep(3)
            ic(f"Đã nhấn nút tìm kiếm")

            soup = BeautifulSoup(driver.page_source, "html.parser")
            company_name = soup.select_one("tbody#render-table-information-disclosure p.ellipsis-two-line").text.strip()
            target_url = f"https://cafef.vn/du-lieu/hose/{symbol.lower()}-{to_slug(company_name)}.chn"
            ic(f"URL mục tiêu: {target_url}")

            driver.get(target_url)
            WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.CLASS_NAME, "dlt-right-half")))
            ic(f"Đã truy cập trang {target_url}")

            soup = BeautifulSoup(driver.page_source, "html.parser")
            avg_volume = soup.select_one(".dlt-right-half .dltl-other li.clearfix div.r").text.strip()
            ic(f"KLGD trung bình của {symbol}: {avg_volume}")
            driver.quit()
            return avg_volume

        except Exception as e:
            print(f"Lỗi khi lấy KLGD trung bình cho {symbol} ở lần thử {attempt + 1}: {e}")
            driver.quit()
            continue

    print(f"Không lấy được KLGD trung bình cho {symbol}")
    return None

# Hàm lấy "Khối lượng lưu hành" từ cophieu68.vn
def get_outstanding_shares(symbol, max_retries=3):
    url = f"https://www.cophieu68.vn/quote/profile.php?id={symbol.lower()}&stockname_search=Submit"
    for attempt in range(max_retries):
        driver = setup_driver()
        try:
            driver.get(url)
            WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.TAG_NAME, "table")))
            soup = BeautifulSoup(driver.page_source, "html.parser")
            table = soup.find("h2", text="THÔNG TIN CÔNG TY").find_next("table")
            for row in table.find("tbody").find_all("tr"):
                cells = row.find_all("td")
                if len(cells) >= 2 and cells[0].text.strip() == "KL lưu hành":
                    driver.quit()
                    return cells[1].text.strip()
            print(f"Không tìm thấy 'KL lưu hành' cho {symbol}")
            driver.quit()
            continue

        except Exception as e:
            print(f"Lỗi khi lấy KL lưu hành cho {symbol} ở lần thử {attempt + 1}: {e}")
            driver.quit()
            continue

    print(f"Không lấy được KL lưu hành cho {symbol}")
    return None

# Hàm lấy "% NN sở hữu" từ finance_url
def get_ownership_ratio(symbol, max_retries=3):
    _, finance_url = get_company_url(symbol)
    if not finance_url:
        return None
    for attempt in range(max_retries):
        driver = setup_driver()
        try:
            driver.get(finance_url)
            WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.ID, "ownedratio")))
            soup = BeautifulSoup(driver.page_source, "html.parser")
            owned_ratio = soup.find("b", id="ownedratio").text.strip()
            driver.quit()
            return owned_ratio

        except Exception as e:
            print(f"Lỗi khi lấy % NN sở hữu cho {symbol} ở lần thử {attempt + 1}: {e}")
            driver.quit()
            continue

    print(f"Không lấy được % NN sở hữu cho {symbol}")
    return None

# Hàm lấy "Lợi nhuận sau thuế" từ tab "Tài chính"
def get_profit_data(symbol, max_retries=3):
    _, finance_url = get_company_url(symbol)
    if not finance_url:
        return None
    for attempt in range(max_retries):
        driver = setup_driver()
        try:
            driver.get(finance_url)
            WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.ID, "tbl-data-BCTT-KQ")))
            soup = BeautifulSoup(driver.page_source, "html.parser")
            table = soup.find("table", id="tbl-data-BCTT-KQ")
            headers = [h.text.strip() for h in table.find("thead").find_all("th", class_="text-center") if h.text.strip()]
            rows = table.find("tbody").find_all("tr")
            if len(rows) < 2:
                print(f"Không đủ hàng trong bảng cho {symbol}")
                driver.quit()
                continue
            values = [cell.text.strip() for cell in rows[-2].find_all("td", class_="text-right")]
            driver.quit()
            return headers, values

        except Exception as e:
            print(f"Lỗi khi lấy Lợi nhuận sau thuế cho {symbol} ở lần thử {attempt + 1}: {e}")
            driver.quit()
            continue

    print(f"Không lấy được Lợi nhuận sau thuế cho {symbol}")
    return None

# Hàm lấy giá đóng cửa gần nhất từ Vietstock
def get_latest_close_price(symbol, max_retries=3):
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

# Hàm xuất dữ liệu ra Excel
def export_to_excel(symbols=["ACB"]):
    excel_file = "financial_data.xlsx"
    with pd.ExcelWriter(excel_file, engine="openpyxl") as writer:
        for symbol in symbols:
            ic(f"\nXử lý dữ liệu cho {symbol}...")
            data = {
                "EPS": get_financial_data("EPS", symbol) or ["N/A"] * 5,
                "ROE": get_financial_data("ROE", symbol) or ["N/A"] * 5,
                "Owned Ratio": get_ownership_ratio(symbol) or "N/A",
                "Outstanding Shares": get_outstanding_shares(symbol) or "N/A",
                "Avg Trading Volume": get_avg_trading_volume(symbol) or "N/A",
                "Profit": get_profit_data(symbol) or (["Q4/2023", "Q1/2024", "Q2/2024", "Q3/2024", "Q4/2024"], ["N/A"] * 5),
                "Latest Close": get_latest_close_price(symbol) or ("N/A", "N/A")
            }
            price, date = data["Latest Close"]

            if all(v == "N/A" or not v for v in data.values() if not isinstance(v, tuple)):
                print(f"Không lấy được dữ liệu cho {symbol}, bỏ qua")
                continue

            # Zone 1
            zone1_data = {
                "Năm tài chính": ["EPS", "ROE", "% NN sở hữu", "Khối lượng lưu hành", "KLGD khớp lệnh trung bình 10 phiên", f"Giá đóng cửa gần nhất tại {date}"],
                "2020": [data["EPS"][0], data["ROE"][0], data["Owned Ratio"], data["Outstanding Shares"], data["Avg Trading Volume"], price],
                "2021": [data["EPS"][1], data["ROE"][1], "", "", "", ""],
                "2022": [data["EPS"][2], data["ROE"][2], "", "", "", ""],
                "2023": [data["EPS"][3], data["ROE"][3], "", "", "", ""],
                "2024": [data["EPS"][4], data["ROE"][4], "", "", "", ""]
            }
            zone1_df = pd.DataFrame(zone1_data)

            # Zone 2
            periods, profit_values = data["Profit"]
            zone2_data = {"Năm tài chính": ["Lợi nhuận sau thuế"]}
            for period, value in zip(periods, profit_values):
                zone2_data[period] = [value]
            zone2_df = pd.DataFrame(zone2_data)

            # Ghép các zone lại với nhau
            final_df = pd.concat([
                zone1_df, 
                pd.DataFrame([[""] * len(zone1_df.columns)], columns=zone1_df.columns),
                zone2_df,
                pd.DataFrame([[""] * len(zone1_df.columns)], columns=zone1_df.columns),
            ], ignore_index=True)
            
            final_df.to_excel(writer, sheet_name=symbol, index=False)
            ic(f"Đã ghi dữ liệu cho {symbol} vào file Excel")

            # Ghi giá đóng cửa vào file txt riêng
            with open("last_close_prices.txt", "a", encoding="utf-8") as f:
                if price != "N/A" and date != "N/A":
                    f.write(f"{symbol}_{date}: {price}\n")

    print(f"\nĐã xuất dữ liệu ra file: {excel_file}")

# Chạy chương trình
symbols = ["ACB","BCM","BID","BVH","CTG","FPT","GAS","GVR","HDB","HPG","LPB","MBB","MSN","MWG","PLX","SAB","SHB","SSB","SSI","STB","TCB","TPB","VCB","VHM","VIB","VIC","VJC","VNM","VPB","VRE"]
# symbols = ["BVH"]
export_to_excel(symbols)
