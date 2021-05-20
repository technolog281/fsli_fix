import logging

from __init__ import *
from openpyxl import load_workbook
from search_last_list import last_page


def insert_script_sql():
    for num in range(2, sheet.max_row + 1):
        vals = values(num)
        if vals['Code_MIS'] is None:
            vals['Code_MIS'] = 'NULL'
        sb_insert = f"INSERT INTO lbr_AllTemp (ResearchTypeID, ResearchName, ParamName, MIS_Code, rf_BIOMID," \
                    f"Unit, Code_Lis, rf_LaboratoryTypeID, rf_ResearchTypeKindID)" \
                    f"VALUES(" \
                    f"'{vals['Num']}', " \
                    f"'{vals['ResearchName']}'," \
                    f"'{vals['ParamName']}'," \
                    f"'{vals['Code_MIS']}'," \
                    f"(SELECT TOP(1) ID FROM lbr_BIOM where BIOM_Name like '{vals['Biom']}')," \
                    f"'{vals['Unit']}'," \
                    f"'{vals['Code_Lis']}'," \
                    f"'{vals['Lab_Type']}'," \
                    f"'{vals['Research_Type']}')"
        wtdb_cursor.execute(sb_insert)
        wtdb_cursor.commit()


insert_script_sql()

