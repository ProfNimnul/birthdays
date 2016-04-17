import xlrd
import xlwt
import datetime
from easygui import fileopenbox,msgbox


def update_person(info, birthday=list()):
    birthday.add(info)
    return birthday


fname = fileopenbox("Выберите файл с днями рождения", "")
if not fname:
    exit()
if not fname.endswith(".xlsx"):
        msgbox("Выбран не файл xlsx", ok_button="Закрыть", title="Проверьте тип файла!")
        exit()
rb = xlrd.open_workbook(fname)
all_persons = list()
persons_with_one_birth_data=dict()

sheets_list = rb.sheet_names() #нашли количество листов в книге
# ws=rb.sheet_by_index(len(sheets_list)-1)
sheet=rb.sheet_by_index(len(sheets_list)-1)
# print(dirname)
for rownum in range(sheet.nrows):

    rows=sheet.row_values(rownum)
    if sheet.cell(rownum,1).ctype !=3:
        continue
    date_tuple = xlrd.xldate_as_tuple(rows[1],rb.datemode)
    full_date =  datetime.datetime(*date_tuple)

    day_of_birth=date_tuple[2] #єто день рождения (именно ДЕНЬ)
    info = sheet.cell(rownum, 0) + "$" + sheet.cell(rownum, 2)  # сформировали строку из ФИО и адреса

    if  not all_persons.get(day_of_birth,None): #такой даты нет вообще
        persons_with_one_birth_data.update(str(day_of_birth),update_person(info,all_persons))
    else
        persons_with_one_birth_data[str(day_of_birth)]=update_person()

pass
exit()