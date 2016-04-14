import xlrd
import xlwt
from easygui import fileopenbox


fname = fileopenbox("Выберите файл с днями рождения", "")
if not fname:
    exit()
if not fname.endswith(".xlsx"):
        msgbox("Выбран не файл xlsx", ok_button="Закрыть", title="Проверьте тип файла!")

rb = xlrd.open_workbook(fname)

sheets_list = rb.sheet_names() #нашли количество листов в книге
ws=rb.sheet_by_index(len(sheets_list)-1)
# print(dirname)
vals = [ws.row_values(rownum) for rownum in range(ws.nrows)]
pass
exit()