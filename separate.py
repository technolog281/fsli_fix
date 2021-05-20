from openpyxl import load_workbook
from search_last_list import last_page
import pandas as pd

wb = load_workbook('xlsx/names_v5.xlsx')
sheet = wb['Sheet1']
df = pd.DataFrame(columns=['NUM', 'RESEARCH_TYPE', 'RESEARCH_PARAM', 'NMU', 'BIOM', 'UNIT', 'LIS_CODE', 'LAB_TYPE'])

for num in range(1, 23485):
    VALUES = {'NUM': sheet[f'A{num}'].value,
              'RESEARCH_TYPE': sheet[f'B{num}'].value,
              'RESEARCH_PARAM': sheet[f'C{num}'].value,
              'NMU': sheet[f'D{num}'].value,
              'BIOM': sheet[f'E{num}'].value,
              'UNIT': sheet[f'F{num}'].value,
              'LIS_CODE': sheet[f'G{num}'].value,
              'LAB_TYPE': sheet[f'H{num}'].value}

    if 'в сыворотке или плазме' in VALUES['RESEARCH_PARAM'] and 'Сыворотка' in VALUES['BIOM']:
        VALUES['RESEARCH_PARAM'] = VALUES['RESEARCH_PARAM'].replace('или плазме ', '')
        df = df.append(VALUES, ignore_index=True, sort=False)
    elif 'в сыворотке или плазме' in VALUES['RESEARCH_PARAM'] and 'Плазма' in VALUES['BIOM']:
        VALUES['RESEARCH_PARAM'] = VALUES['RESEARCH_PARAM'].replace('сыворотке или ', '')
        df = df.append(VALUES, ignore_index=True, sort=False)
    elif 'в крови или образце' in VALUES['RESEARCH_PARAM'] and 'Ткань' in VALUES['BIOM']:
        VALUES['RESEARCH_PARAM'] = VALUES['RESEARCH_PARAM'].replace('крови или ', '')
        df = df.append(VALUES, ignore_index=True, sort=False)
    elif 'в крови или образце' in VALUES['RESEARCH_PARAM'] and 'Кровь' in VALUES['BIOM']:
        VALUES['RESEARCH_PARAM'] = VALUES['RESEARCH_PARAM'].replace(' или образце тканей', '')
        df = df.append(VALUES, ignore_index=True, sort=False)
    elif 'в плазме или венозной' in VALUES['RESEARCH_PARAM'] and 'Плазма' in VALUES['BIOM']:
        VALUES['RESEARCH_PARAM'] = VALUES['RESEARCH_PARAM'].replace('или венозной ', '')
        df = df.append(VALUES, ignore_index=True, sort=False)
    elif 'в плазме или венозной' in VALUES['RESEARCH_PARAM'] and 'Кровь' in VALUES['BIOM']:
        VALUES['RESEARCH_PARAM'] = VALUES['RESEARCH_PARAM'].replace('плазме или ', '')
        df = df.append(VALUES, ignore_index=True, sort=False)
    else:
        df = df.append(VALUES, ignore_index=True, sort=False)

# print(df)
df.to_excel('names_v6.xlsx')
