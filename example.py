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

