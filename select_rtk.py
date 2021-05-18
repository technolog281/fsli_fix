import pandas as pd
# from __init__ import *
# from encodings import ascii as us-ascii
from openpyxl import load_workbook
from openpyxl import Workbook

wb = load_workbook('./names_v5.xlsx')
sheet_1 = wb['Sheet1']
num = 1
values_2 = {'G_NAME': str(sheet_1[f'A{num}'].value),
            'NAME': str(sheet_1[f'B{num}'].value),
            'CODE': str(sheet_1[f'C{num}'].value)}
# values_2 = [str(sheet_1[f'A{num}'].value), str(sheet_1[f'B{num}'].value), str(sheet_1[f'C{num}'].value)]
num_symbols = ('1', '2', '3', '4', '5', '6', '7', '(')
wb_ext = Workbook('names_upgrade.xlsx')
ws = wb_ext.create_sheet('Sheet1')
print(values_2['G_NAME'])
ws.append([values_2['G_NAME'], values_2['NAME'], values_2['CODE']])
#
#
# for num in range(2, 200):
#     start_g_name = values_2['G_NAME']
#     start_name = values_2['NAME']
#     if 'производн' in start_g_name:
#         values_2['CODE'] = '200'
#         ws.append([values_2['G_NAME'], values_2['NAME'], values_2['CODE']])
#         # ws.append(values_2['NAME'])
#         # ws.append(values_2['CODE'])
#         # ws[f'A{num}'] = values_2['G_NAME']
#         # ws[f'B{num}'] = values_2['NAME']
#         # ws[f'C{num}'] = values_2['CODE']
#     elif 'Днк' or 'Рнк' in start_name:
#         values_2['CODE'] = '208'
#         # rf.append(values_2, ignore_index=True, sort=False)
#     elif 'иммунофермент' in start_name:
#         values_2['CODE'] = '206'
#         # rf.append(values_2, ignore_index=True, sort=False)
#     elif 'иммунохромат' in start_name:
#         values_2['CODE'] = '204'
#         # rf.append(values_2, ignore_index=True, sort=False)
#     elif 'рН' in start_g_name:
#         values_2['CODE'] = '200'
#         # rf.append(values_2, ignore_index=True, sort=False)
#     elif 'Массовая концентрация' or 'Массовое содержание' in start_name:
#         values_2['CODE'] = '200'
#         # rf.append(values_2, ignore_index=True, sort=False)
#     elif 'ромбоцит' or 'ейкоцит' or 'ритроцит' or 'имфорцит' in start_g_name:
#         values_2['CODE'] = '202'
#         # rf.append(values_2, ignore_index=True, sort=False)
#     elif 'бнаружени' in start_name:
#         values_2['CODE'] = '208'
#         # rf.append(values_2, ignore_index=True, sort=False)
#     elif 'иммунофлуор' in start_name:
#         values_2['CODE'] = '207'
#         # rf.append(values_2, ignore_index=True, sort=False)
#     elif 'иммунохеми' in start_name:
#         values_2['CODE'] = '207'
#         # rf.append(values_2, ignore_index=True, sort=False)
#     elif 'иммунолог' in start_name:
#         values_2['CODE'] = '204'
#         # rf.append(values_2, ignore_index=True, sort=False)
#     elif 'ремя' in start_name:
#         values_2['CODE'] = '201'
#         # rf.append(values_2, ignore_index=True, sort=False)
#     elif 'микроорган' or 'ультур' in start_name:
#         values_2['CODE'] = '208'
#         # rf.append(values_2, ignore_index=True, sort=False)
#     elif 'язкост' or 'Запах' or 'Цвет' in start_name:
#         values_2['CODE'] = '201'
#         # rf.append(values_2, ignore_index=True, sort=False)
#     elif 'олярная' in start_name:
#         values_2['CODE'] = '200'
#         # rf.append(values_2, ignore_index=True, sort=False)
#     # else:
#     #     # rf.append(values_2, ignore_index=True, sort=False)
#
# # print(rf)
# # df.to_excel('names_v9.xlsx')
#
