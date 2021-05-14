from openpyxl import load_workbook
import pandas as pd

wb = load_workbook('./names_v3.xlsx')
sheet = wb['Worker']

df = pd.DataFrame(columns=['Unit'])

list_of_units = []


def sum(num):
    A = sheet[f'A{num}'].value
    if A is not None:
        A.split(' ; ')
        A.split(' , ')
    return A


for num in range(1, 16731):
    list_of_units.append(sum(num))

fl = []
for lst in list_of_units:

    if lst is None:
        continue
    # elif ' ' in lst:
    #     continue
    elif lst == '':
        continue
    else:
        fl.append(lst)

fl = set(fl)

df = pd.concat([pd.DataFrame([i], columns=['BIOM']) for i in fl], ignore_index=True)
print(df)
df.to_excel('bioms.xlsx')
