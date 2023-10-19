# Overtime_Machine
    
加班時數統計神器，無情加班機器人的好幫手。

## 如何安裝

* ### 透過 PyPI 安裝

    使用以下指令即可安裝此最新commit的套件( branch : main )。

    ```bash
    pip install git+https://github.com/lianyun0502/Overtime_Machine.git
    ```

* ### 透過源程式碼安裝
    或直接 clone 源程式碼執行。
        
    ```bash
    git clone https://github.com/lianyun0502/Overtime_Machine.git
    ```
    並用指令安裝所需套件。

    ```bash
    pip install -r requirements.txt
    ```


## 使用說明

參考範例程式碼 example.py ，將程式碼中的變數值修改為自己的資料，並執行程式。
* E_NUMBER: 員工編號
* E_NAME: 員工姓名
* YEAR: 年度
* START_DATE: 加班起始日期
* END_DATE: 加班結束日期
* EXCEL_FILE: 原始加班紀錄EXCEL檔案路徑
* SAVE_WORD_FILE: 輸出的加班補休認可表檔案路徑

```python
from overtime_machine import happ_overtime
from datetime import datetime



E_NUMBER = 'K0081'
E_NAME = '李璉昀 Eric Li'

YEAR = '111'
START_DATE = datetime(2023, 8, 21)
END_DATE = datetime(2023, 9, 20)

EXCEL_FILE = r'20230821~20230920 Attendance records.xlsx'

SAVE_WORD_FILE = r'加班補休認可表-0081-Eric.docx'

if __name__ == '__main__':
    
    happ_overtime(e_number=E_NUMBER, 
                  e_name=E_NAME, 
                  year=YEAR, 
                  start_date=START_DATE, 
                  end_date=END_DATE, 
                  excel_file=EXCEL_FILE, 
                  save_word_file=SAVE_WORD_FILE)
```
## 共同開發

歡迎各位加班超人一起開發，請先 fork 此專案，並在自己的專案中進行開發，開發完成後再發起 pull request，我會盡快處理。
