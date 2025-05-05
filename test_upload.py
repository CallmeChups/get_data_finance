import pandas as pd
from icecream import ic
import gspread
from oauth2client.service_account import ServiceAccountCredentials
import numpy as np
import time

def convert_value(value):
    """Chuyển đổi giá trị thành định dạng phù hợp cho Google Sheets"""
    if pd.isna(value) or value == 'N/A':
        return None
    if isinstance(value, (int, float)):
        if np.isinf(value) or np.isnan(value):
            return None
        return value
    # Thử chuyển đổi chuỗi thành số
    try:
        # Xử lý chuỗi có dấu phẩy ngăn cách hàng nghìn
        if isinstance(value, str):
            value = value.replace(',', '')
        return float(value)
    except:
        return None

def convert_to_float(value):
    """Chuyển đổi giá trị thành float nếu có thể"""
    if pd.isna(value) or value == 'N/A':
        return None
    if isinstance(value, (int, float)):
        if np.isinf(value) or np.isnan(value):
            return None
        return float(value)
    # Thử chuyển đổi chuỗi thành số
    try:
        # Xử lý chuỗi có dấu phẩy ngăn cách hàng nghìn
        if isinstance(value, str):
            value = value.replace(',', '')
        return float(value)
    except:
        return None

# Hàm cập nhật Google Sheets
def update_google_sheets(symbol, data):
    # Thiết lập credentials cho Google Sheets API
    scope = ['https://spreadsheets.google.com/feeds',
             'https://www.googleapis.com/auth/drive']
    creds = ServiceAccountCredentials.from_json_keyfile_name('credentials.json', scope)
    client = gspread.authorize(creds)
    
    # Mở spreadsheet
    spreadsheet = client.open_by_url('https://docs.google.com/spreadsheets/d/1BrzaI-Il-H2IDliPGQEzn9aXu5LbM0Mltg6OMQure4U/edit?gid=969948111')
    
    # Lấy worksheet cho symbol
    try:
        worksheet = spreadsheet.worksheet(symbol)
    except:
        print(f"Không tìm thấy sheet cho {symbol}")
        return
    
    # Chuẩn bị dữ liệu để cập nhật theo batch
    batch_data = {
        'requests': []
    }
    
    # Cập nhật EPS (5 giá trị)
    eps_cells = ['K13', 'L13', 'M13', 'N13', 'O13']
    for cell, value in zip(eps_cells, data['EPS']):
        converted_value = convert_to_float(value)
        if converted_value is not None:
            batch_data['requests'].append({
                'updateCells': {
                    'range': {
                        'sheetId': worksheet.id,
                        'startRowIndex': int(cell[1:]) - 1,
                        'endRowIndex': int(cell[1:]),
                        'startColumnIndex': ord(cell[0]) - ord('A'),
                        'endColumnIndex': ord(cell[0]) - ord('A') + 1
                    },
                    'rows': [{
                        'values': [{
                            'userEnteredValue': {
                                'numberValue': converted_value
                            }
                        }]
                    }],
                    'fields': 'userEnteredValue'
                }
            })
    
    # Cập nhật ROE (5 giá trị)
    roe_cells = ['K15', 'L15', 'M15', 'N15', 'O15']
    for cell, value in zip(roe_cells, data['ROE']):
        converted_value = convert_to_float(value)
        if converted_value is not None:
            batch_data['requests'].append({
                'updateCells': {
                    'range': {
                        'sheetId': worksheet.id,
                        'startRowIndex': int(cell[1:]) - 1,
                        'endRowIndex': int(cell[1:]),
                        'startColumnIndex': ord(cell[0]) - ord('A'),
                        'endColumnIndex': ord(cell[0]) - ord('A') + 1
                    },
                    'rows': [{
                        'values': [{
                            'userEnteredValue': {
                                'numberValue': converted_value
                            }
                        }]
                    }],
                    'fields': 'userEnteredValue'
                }
            })
    
    # Cập nhật % NN sở hữu
    cell = 'C29'
    converted_value = convert_to_float(data['Owned Ratio'])
    if converted_value is not None:
        batch_data['requests'].append({
            'updateCells': {
                'range': {
                    'sheetId': worksheet.id,
                    'startRowIndex': int(cell[1:]) - 1,
                    'endRowIndex': int(cell[1:]),
                    'startColumnIndex': ord(cell[0]) - ord('A'),
                    'endColumnIndex': ord(cell[0]) - ord('A') + 1
                },
                'rows': [{
                    'values': [{
                        'userEnteredValue': {
                            'numberValue': converted_value
                        }
                    }]
                }],
                'fields': 'userEnteredValue'
            }
        })
    
    # Cập nhật Khối lượng lưu hành
    cell = 'O36'
    converted_value = convert_to_float(data['Outstanding Shares'])
    if converted_value is not None:
        batch_data['requests'].append({
            'updateCells': {
                'range': {
                    'sheetId': worksheet.id,
                    'startRowIndex': int(cell[1:]) - 1,
                    'endRowIndex': int(cell[1:]),
                    'startColumnIndex': ord(cell[0]) - ord('A'),
                    'endColumnIndex': ord(cell[0]) - ord('A') + 1
                },
                'rows': [{
                    'values': [{
                        'userEnteredValue': {
                            'numberValue': converted_value
                        }
                    }]
                }],
                'fields': 'userEnteredValue'
            }
        })
    
    # Cập nhật KLGD khớp lệnh trung bình 10 phiên
    cell = 'C24'
    converted_value = convert_to_float(data['Avg Trading Volume'])
    if converted_value is not None:
        batch_data['requests'].append({
            'updateCells': {
                'range': {
                    'sheetId': worksheet.id,
                    'startRowIndex': int(cell[1:]) - 1,
                    'endRowIndex': int(cell[1:]),
                    'startColumnIndex': ord(cell[0]) - ord('A'),
                    'endColumnIndex': ord(cell[0]) - ord('A') + 1
                },
                'rows': [{
                    'values': [{
                        'userEnteredValue': {
                            'numberValue': converted_value
                        }
                    }]
                }],
                'fields': 'userEnteredValue'
            }
        })
    
    # Cập nhật Giá đóng cửa gần nhất
    price, date = data['Latest Close']
    cell = 'K45'
    converted_value = convert_to_float(price)
    if converted_value is not None:
        batch_data['requests'].append({
            'updateCells': {
                'range': {
                    'sheetId': worksheet.id,
                    'startRowIndex': int(cell[1:]) - 1,
                    'endRowIndex': int(cell[1:]),
                    'startColumnIndex': ord(cell[0]) - ord('A'),
                    'endColumnIndex': ord(cell[0]) - ord('A') + 1
                },
                'rows': [{
                    'values': [{
                        'userEnteredValue': {
                            'numberValue': converted_value
                        }
                    }]
                }],
                'fields': 'userEnteredValue'
            }
        })
    
    # Cập nhật Lợi nhuận sau thuế (5 giá trị)
    profit_cells = ['K3', 'L3', 'M3', 'N3', 'O3']
    for cell, value in zip(profit_cells, data['Profit'][1]):  # Lấy phần values từ tuple (periods, values)
        converted_value = convert_to_float(value)
        if converted_value is not None:
            batch_data['requests'].append({
                'updateCells': {
                    'range': {
                        'sheetId': worksheet.id,
                        'startRowIndex': int(cell[1:]) - 1,
                        'endRowIndex': int(cell[1:]),
                        'startColumnIndex': ord(cell[0]) - ord('A'),
                        'endColumnIndex': ord(cell[0]) - ord('A') + 1
                    },
                    'rows': [{
                        'values': [{
                            'userEnteredValue': {
                                'numberValue': converted_value
                            }
                        }]
                    }],
                    'fields': 'userEnteredValue'
                }
            })
    
    # Thực hiện cập nhật theo batch
    try:
        if batch_data['requests']:  # Chỉ thực hiện nếu có request
            spreadsheet.batch_update(batch_data)
            print(f"Đã cập nhật dữ liệu cho {symbol} lên Google Sheets")
            time.sleep(2)  # Thêm delay 2 giây sau mỗi lần cập nhật
    except Exception as e:
        print(f"Lỗi khi cập nhật Google Sheets cho {symbol}: {e}")
        time.sleep(5)  # Thêm delay dài hơn nếu có lỗi

# Hàm đọc dữ liệu từ Excel và upload lên Google Sheets
def upload_excel_to_sheets():
    excel_file = "financial_data.xlsx"
    xls = pd.ExcelFile(excel_file)
    
    for symbol in xls.sheet_names:
        print(f"\nXử lý dữ liệu cho {symbol}...")
        df = pd.read_excel(excel_file, sheet_name=symbol)
        
        # In ra thông tin về DataFrame để debug
        print(f"Số hàng trong DataFrame: {len(df)}")
        print(f"Các cột trong DataFrame: {df.columns.tolist()}")
        print("\nDữ liệu từ hàng 7-10:")
        print(df.iloc[7:10])
        
        # Chuẩn bị dữ liệu
        data = {
            "EPS": df.iloc[0, 1:6].tolist(),  # Lấy 5 giá trị EPS từ cột 2-6
            "ROE": df.iloc[1, 1:6].tolist(),  # Lấy 5 giá trị ROE từ cột 2-6
            "Owned Ratio": df.iloc[2, 1],     # Lấy % NN sở hữu
            "Outstanding Shares": df.iloc[3, 1],  # Lấy Khối lượng lưu hành
            "Avg Trading Volume": df.iloc[4, 1],  # Lấy KLGD trung bình
            "Latest Close": (df.iloc[5, 1], df.iloc[5, 0].split("tại ")[-1]),  # Lấy giá đóng cửa và ngày
            "Profit": (df.columns[6:11].tolist(), df.iloc[7, 6:11].tolist())  # Lấy lợi nhuận sau thuế từ các cột Q1/2024-Q1/2025
        }
        
        # In ra để kiểm tra dữ liệu lợi nhuận
        print(f"\nDữ liệu lợi nhuận sau thuế cho {symbol}:")
        print(f"Periods: {data['Profit'][0]}")
        print(f"Values: {data['Profit'][1]}")
        
        # Upload lên Google Sheets
        try:
            update_google_sheets(symbol, data)
            time.sleep(1)  # Thêm delay 1 giây giữa các symbol
        except Exception as e:
            print(f"Lỗi khi cập nhật Google Sheets cho {symbol}: {e}")
            time.sleep(5)  # Thêm delay dài hơn nếu có lỗi

# Chạy chương trình
if __name__ == "__main__":
    upload_excel_to_sheets() 