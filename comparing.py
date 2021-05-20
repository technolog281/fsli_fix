import pyodbc
from openpyxl import load_workbook
from search_last_list import last_page

wb = load_workbook('xlsx/temp.xlsx')
sheet = wb['Temp_worker']

cnxn = pyodbc.connect('DRIVER={ODBC Driver 17 for SQL Server};'
                      'SERVER=DROID1347\SQL_MIS;'
                      'DATABASE=script_base;'
                      'UID=sa;'
                      'PWD=Apacheserver1390')

cnxn.setencoding(encoding='utf-8')
cursor = cnxn.cursor()


def finder(name):
    cursor.execute("select ID from Units where Name like ?", name)
    to_find = cursor.fetchone()
    return to_find


res_str = []

for num in range(1, last_page('b1')):
    val = sheet[f'B{num}'].value
    val_split = val.split(';')
    temp_dict = []
    for n in range(0, len(val_split)):
        x = val_split[n]
        f_id = finder(val_split[n])[0]
        if val_split[n] is not None:
            temp_dict.append(f_id)
        else:
            continue
    res_str.append(temp_dict)

print(res_str)
