import logging

from __init__ import *

to_search_book = load_workbook('xlsx/to_search.xlsx')
ts_sheet = to_search_book['Sheet1']

df = pd.DataFrame()


def select_research(code: int):
    wtdb_cursor.execute(f"SELECT ResearchTypeID, ResearchName, ParamName, MIS_Code,"
                        f"Unit, Code_Lis, rf_ResearchTypeKindID FROM lbr_AllTemp "
                        f"WHERE Code_Lis = '{code}'")
    record = wtdb_cursor.fetchone()
    if not record:
        logging.warning(f'Record not found, search_code: {search_code}')
    else:
        logging.info(f'Record founded, search_code: {search_code}')
    return record


for num in range(1, ts_sheet.max_row):
    search_name = ts_sheet[f'A{num}'].value
    search_code = ts_sheet[f'B{num}'].value

    research = select_research(search_code)
    if research is None:
        result_dict = {'Search_Code': search_code,
                       'Search_Name': 'Ничего не нашёл :('}
    result_dict = {
        'Search_Code': search_code,
        'Search_Name': search_name,
        'ResearchTypeID': research[0],
        'ResearchName': research[1],
        'ParamName': research[2],
        'MIS_Code': research[3],
        'Unit': research[4],
        'Code_Lis': research[5],
        'rf_ResearchTypeKindID': research[6]
    }
    df = df.append(result_dict, ignore_index=True)

df.to_excel('results.xlsx')
