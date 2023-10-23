from overtime_machine import happy_overtime
from datetime import datetime



E_NUMBER = 'K0081'
E_NAME = '李璉昀 Eric Li'

YEAR = '111'
START_DATE = datetime(2023, 9, 21)
END_DATE = datetime(2023, 10, 20)

EXCEL_FILE = r'C:\Users\eric.li\Desktop\My Github\Overtime_Machine\20230921~20231020 Attendance records.xlsx'

SAVE_WORD_FILE = r'加班補休認可表-0081-Eric.docx'

EXCLUDE_DATE = [datetime(2023, 9, 23)]

if __name__ == '__main__':

    happy_overtime(e_number=E_NUMBER, 
                  e_name=E_NAME, 
                  year=YEAR, 
                  start_date=START_DATE, 
                  end_date=END_DATE, 
                  excel_file=EXCEL_FILE,
                  exclude_date=EXCLUDE_DATE, 
                  save_word_file=SAVE_WORD_FILE)

