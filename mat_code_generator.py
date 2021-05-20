from __init__ import *

wb = load_workbook('xlsx/names_v6.xlsx')
sheet = wb['Sheet1']
df = pd.DataFrame(
    columns=['NUM', 'RESEARCH_TYPE', 'RESEARCH_PARAM', 'NMU', 'BIOM', 'UNIT', 'LIS_CODE', 'MAT_CODE' 'LAB_TYPE'])

for num in range(2, 23485):
    ex = values(num)
    biom = ex['BIOM']
    biom_x = f"'%{biom}%'"
    sb_cursor.execute(f'select ID from BIOMS WHERE BIOM_Name like {biom_x}')
    append_num = sb_cursor.fetchone()
    if ex['LIS_CODE'] and append_num is not None:
        ex['MAT_CODE'] = str(ex['LIS_CODE']) + '.' + str(append_num[0])
    else:
        ex['MAT_CODE'] = 'АЦАЦАЦАЦАЦАЦА'
    df = df.append(ex, ignore_index=True, sort=False)

df.to_excel('names_v7.xlsx')
