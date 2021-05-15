from __init__ import *
from openpyxl import load_workbook
from search_last_list import last_page

wb = load_workbook('./RTKinds.xlsx')
sheet = wb['Sheet1']


def insert_script_sql(params):
    sb_insert = f"INSERT INTO lbr_ResearchTypeKind (ResTypeKindID, Code, Name, Description, UGUID) " \
                f"VALUES(" \
                f"'{params['ResTypeKindID']}', " \
                f"'{params['Code']}', " \
                f"'{params['Name']}', " \
                f"'{params['Description']}', " \
                f"'{params['UGUID']}') "
    sb_cursor.execute(sb_insert)
    sb_cursor.commit()


for num in range(1, last_page('a1') - 1):
    from_excel = {'ResTypeKindID': sheet[f'A{num}'].value,
                  'Code': sheet[f'B{num}'].value,
                  'Name': sheet[f'C{num}'].value,
                  'Description': sheet[f'D{num}'].value,
                  'UGUID': sheet[f'E{num}'].value}
    sb_cursor.execute('SELECT UGUID FROM lbr_ResearchTypeKind WHERE UGUID = ?', from_excel['UGUID'])
    UGUID_from_base = str(sb_cursor.fetchall())
    if from_excel['UGUID'] in UGUID_from_base:
        Name = from_excel['Name']
        print(f'Я уже добавил {Name}')
    else:
        insert_script_sql(from_excel)
