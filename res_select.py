from __init__ import *

to_search_book = load_workbook('to_search.xlsx')
ts_sheet = to_search_book['Sheet1']

df = pd.DataFrame()


for num in range(1, ts_sheet.max_row):
    search_name = ts_sheet[f'A{num}'].value
    search_code = ts_sheet[f'B{num}'].value

    wtdb_cursor.execute(f"SELECT ResearchTypeID FROM lbr_AllTemp WHERE Code_Lis like '{search_code}'")
    ResearchTypeID = wtdb_cursor.fetchone()[0]
    wtdb_cursor.execute(f"SELECT ResearchName FROM lbr_AllTemp WHERE Code_Lis like '{search_code}'")
    ResearchName = wtdb_cursor.fetchone()[0]
    wtdb_cursor.execute(f"SELECT ParamName FROM lbr_AllTemp WHERE Code_Lis like '{search_code}'")
    ParamName = wtdb_cursor.fetchone()[0]
    wtdb_cursor.execute(f"SELECT MIS_Code FROM lbr_AllTemp WHERE Code_Lis like '{search_code}'")
    MIS_Code = wtdb_cursor.fetchone()[0]
    wtdb_cursor.execute(f"SELECT Unit FROM lbr_AllTemp WHERE Code_Lis like '{search_code}'")
    Unit = wtdb_cursor.fetchone()[0]
    wtdb_cursor.execute(f"SELECT Code_Lis FROM lbr_AllTemp WHERE Code_Lis like '{search_code}'")
    Code_Lis = wtdb_cursor.fetchone()[0]
    wtdb_cursor.execute(f"SELECT rf_ResearchTypeKindID FROM lbr_AllTemp WHERE Code_Lis like '{search_code}'")
    rf_ResearchTypeKindID = wtdb_cursor.fetchone()[0]
    result_dict = {'Search_Code': search_code,
                   'Search_Name': search_name,
                   'ResearchTypeID': ResearchTypeID,
                   'ResearchName': ResearchName,
                   'ParamName': ParamName,
                   'MIS_Code': MIS_Code,
                   'Unit': Unit,
                   'Code_Lis': Code_Lis,
                   'rf_ResearchTypeKindID': rf_ResearchTypeKindID}
    df = df.append(result_dict, ignore_index=True)

df.to_excel('results.xlsx')
