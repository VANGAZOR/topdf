from win32com import client
from pywintypes import com_error
from datetime import datetime

excel = client.DispatchEx("Excel.Application")
excel.Interactive = False
excel.Visible = False

start_time = datetime.now()

try:
    print('Попытка конвертации')
    sheets = excel.Workbooks.Open(r'D:\DEV\docxtopf\222.xlsx')  #полный путь до фаила
    sheets.ActiveSheet.ExportAsFixedFormat(0, r'D:\DEV\docxtopf\222.pdf') #аналогично

except com_error as e:
    print('Конвертация провалена')
else:
    print('Конвертация успешна')
    end_time = datetime.now()
    all_time = end_time - start_time
    print(f'Расчетное время конвертации ${all_time}')

finally:
        try:
            print('Закрываем фаил')
            sheets.Close()
            excel.Quit()
        except com_error as e: #обработка ошибки
            if e.hresult == -2147418111:
                print('Возможно фаил большой и возникла проблема с закрытием')
            else:
                 raise



