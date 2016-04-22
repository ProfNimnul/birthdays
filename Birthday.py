import xlrd
from xlwt import Workbook
import datetime
from easygui import fileopenbox,msgbox
# from tempfile import TemporaryFile

def MySort(info):
    return int(info.split("$")[0])

fname = fileopenbox("Выберите файл с днями рождения", "")
if not fname:
    exit()
if not fname.endswith(".xlsx"):
        msgbox("Выбран не файл xlsx", ok_button="Закрыть", title="Проверьте тип файла!")
        exit()
rb = xlrd.open_workbook(fname)
all_persons = list()


sheets_list = rb.sheet_names() #нашли количество листов в книге
# ws=rb.sheet_by_index(len(sheets_list)-1)
sheet=rb.sheet_by_index(len(sheets_list)-1)
old_sheet_name=sheet.name
# print(dirname)
for rownum in range(sheet.nrows):

    rows=sheet.row_values(rownum)
    if sheet.cell(rownum,1).ctype !=3:
        continue
    date_tuple = xlrd.xldate_as_tuple(rows[1],rb.datemode)
    full_date =  datetime.datetime(*date_tuple)

     #єто день рождения (именно ДЕНЬ)
    day_of_birth = date_tuple[2]
    info = str(day_of_birth)+ "$"+rows[0] + "$" + rows[2]  # сформировали строку из ФИО и адреса
    try:

        all_persons.append(info)


    except AttributeError:
        print("Ошибка добавления в словарь!")

    info=None
all_persons.sort(key=MySort)
new_wb=Workbook(encoding="utf-8")
del sheet
sheet=new_wb.add_sheet("Сортировано - "+old_sheet_name)
sheet.write(0,0,"День рождения")
sheet.write(0,1,"ФИО")
sheet.write(0,2,"Адрес")

row_count=1
for info in all_persons:
    bd,fio,adr=info.split("$")
    sheet.write(row_count,0,bd)
    sheet.write(row_count,1 ,fio)
    sheet.write(row_count,2 ,adr)
    row_count+=1
new_wb.save("Sorted.xlsx")
# new_wb.save(TemporaryFile())

pass
exit()