from overtime_machine.excel import *
from overtime_machine.word import *
from datetime import datetime

def happ_overtime(e_number:str, 
                  e_name:str, 
                  year:str, 
                  start_date:datetime, 
                  end_date:datetime, 
                  excel_file:str, 
                  save_word_file:str, 
                  template_word_file:str=r'overtime_machine\template_word\加班補休認可表.docx'):
    
    my_data = get_exel_data(excel_file, e_number)

    df = get_DataFrame(my_data, start_date)

    date_region = f'{start_date.month}/{start_date.day}-{end_date.month}/{end_date.day}'


    text1 = f'\t（\t{year} ）年（{date_region}）\t  工號：\t{e_number}   姓名： {e_name}'

    doc = docx.Document(template_word_file)

    print('段落數量： ', len(doc.paragraphs))
    print('表格數量： ', len(doc.tables))

    write_paragraph(doc.paragraphs, 3, text1, 14)

    for table in doc.tables:
        rows = list(table.rows[2:])
        print(f'table rows number:{len(rows)}')


    for idx in range(len(df)):
        # print(df.iloc[idx])
        row = rows[idx]
        write_reason(row, 0, '')
        write_predict_hour(row, 1, df.iloc[idx]['Over Time Hours'])
        write_overtime_region(row, 3, 
                            date=df.iloc[idx]['Date'], 
                            region=(df.iloc[idx]['Start Over Time'], df.iloc[idx]['Last Record']))
        write_actually_hour(row, 5, df.iloc[idx]['Over Time Hours'])

        if df.iloc[idx]['Weekday'] == '星期六':
            write_overtime_type(row, 7, 'v')
        elif df.iloc[idx]['Weekday'] == '星期日':
            write_overtime_type(row, 8, 'v')
        else:
            write_overtime_type(row, 6, 'v')

    doc.save(save_word_file)
    print(save_word_file)
    print('done')