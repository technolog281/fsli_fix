from __init__ import *
from openpyxl import load_workbook
from search_last_list import last_page


def insert_script_sql():
    for num in range(2, sheet.max_row + 1):
        sb_insert = f"INSERT INTO lbr_AllTemp (ResearchTypeID, ResearchName, ParamName, MIS_Code, rf_BIOMID," \
                    f"Unit, Code_Lis, rf_LaboratoryTypeID, rf_ResearchTypeKindID)" \
                    f"VALUES(" \
                    f"'{values(num)['Num']}', " \
                    f"'{values(num)['ResearchName']}'," \
                    f"'{values(num)['ParamName']}'," \
                    f"'{values(num)['Code_MIS']}'," \
                    f"(SELECT TOP(1) ID FROM lbr_BIOM where BIOM_Name like '{values(num)['Biom']}')," \
                    f"'{values(num)['Unit']}'," \
                    f"'{values(num)['Code_Lis']}'," \
                    f"'{values(num)['Lab_Type']}'," \
                    f"'{values(num)['Research_Type']}')"
        wtdb_cursor.execute(sb_insert)
        wtdb_cursor.commit()


insert_script_sql()

