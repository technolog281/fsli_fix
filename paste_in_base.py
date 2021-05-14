from __init__ import *
from openpyxl import load_workbook
from search_last_list import last_page

wb = load_workbook('./names_v6.xlsx')
sheet = wb['Sheet3']

for num in range(1, last_page('b1')):
    ID = sheet[f'A{num}'].value
    Name = sheet[f'B{num}'].value
    sb_cursor.execute('insert into BIOMS(ID, Name, UGUID) values (?, ?, (select newID()))', ID, Name)
    sb_cursor.commit()
