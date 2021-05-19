import pyodbc
from openpyxl import load_workbook
import loguru as log
from search_last_list import last_page
import pandas as pd

wb = load_workbook('./names_v5.xlsx')
sheet = wb['Sheet1']


def values(num):
    result = {'Num': sheet[f'A{num}'].value,
              'ResearchName': sheet[f'B{num}'].value,
              'ParamName': sheet[f'C{num}'].value,
              'Code_MIS': sheet[f'D{num}'].value,
              'Biom': sheet[f'E{num}'].value,
              'Unit': sheet[f'F{num}'].value,
              'Code_Lis': sheet[f'G{num}'].value,
              'Lab_Type': sheet[f'H{num}'].value,
              'Research_Type': sheet[f'I{num}'].value}
    return result


work_test_db = pyodbc.connect("DRIVER={SQL Server};"
                              "SERVER=KRG-OVIS-0048\SQLMIS;"
                              "DATABASE=script_base;"
                              "UID=sa;"
                              "PWD=@pacheServer1390")

wtdb_cursor = work_test_db.cursor()


# # amur_mis_belbol = amur_mis_belbol_conn.setencoding(encoding='utf-8') ## -- Декодирование
# belbol = pyodbc.connect("DRIVER={SQL Server};"
#                         "SERVER=DROID1347\SQL_MIS;"
#                         "DATABASE=amur_mis_belbol;"
#                         "UID=sa;"
#                         "PWD=Apacheserver1390")
# belbol_cursor = belbol.cursor()
# # amur_mis_belbol = amur_mis_belbol_conn.setencoding(encoding='utf-8') ## -- Декодирование
#
#
# script_base = pyodbc.connect('DRIVER={SQL Server};'
#                              'SERVER=DROID1347\SQL_MIS;'
#                              'DATABASE=script_base;'
#                              'UID=sa;'
#                              'PWD=Apacheserver1390')
# sb_cursor = script_base.cursor()
# # script_base = amur_mis_belbol_conn.setencoding(encoding='utf-8') ## -- Декодирование
