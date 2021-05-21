import pyodbc
from openpyxl import load_workbook
import loguru as log
from search_last_list import last_page
import pandas as pd
import logging

et_list_of_words = {
    'У': 'US',  # УЗИ
    'ОФ': 'OR',  # Офтальмологические исследования
    'Ф': 'FS',  # Функциональные исследования
    'Э': 'EE',  # Эндоскопические исследования
    'РГ': 'XR',  # Рентгенография
    'МРТ': 'MRI',  # МРТ
    'КТ': 'CT',  # КТ
    'АР': 'AG',  # Ангиография
    'РД': 'RD',  # Радионуклидная диагностика
    'Б': 'BR',  # Биохимические исследования
    'К': 'CR',  # Клинические исследования
    'Г': 'HE',  # Гематологические исследования
    'КГ': 'CS',  # Коагулогические исследования
    'И': 'IS',  # Иммунологические исследования
    'Ц': 'CytS',  # Цитологические исследования
    'ДГ-ИФА': 'DELIS',  # Диагностика гормонов методом иммуноферментного анализа
    'ИФА': 'ELISA',  # Исследования методом иммуноферментного анализа (ИФА)
    'ПЦР': 'PCR',  # Исследования методом полимеразной цепной реакции (ПЦР)
    'БАК': 'BacR',  # Бактериологические исследования
    'ДПИ-ИФА': 'Parasite_ELISA',  # Диагностика паразитарных инфекций  методом иммуноферментного анализа (ИФА)
    'КДИ': 'CDS',  # Клинико-диагностическое исследование (КДИ)
    'ДИС': 'DIS',  # Диагностика иммунного статуса
    'ИХА': 'IA'}  # Иммунохроматографический анализ (ИХА)

wb = load_workbook('temp_list.xlsx')
sheet = wb['Sheet1']


def values(num):
    result = {'Num': sheet[f'A{num}'].value,
              'ResearchName': sheet[f'B{num}'].value,
              'ParamName': sheet[f'C{num}'].value,
              'Code_MIS': str(sheet[f'D{num}'].value),
              'Biom': sheet[f'E{num}'].value,
              'Unit': sheet[f'F{num}'].value,
              'Code_Lis': str(sheet[f'G{num}'].value),
              'Lab_Type': sheet[f'H{num}'].value,
              'Research_Type': str(sheet[f'I{num}'].value)}
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
