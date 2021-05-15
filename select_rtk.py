from __init__ import *
from openpyxl import load_workbook

wb = load_workbook('./probe.xlsx')
sheet = wb['Sheet1']

for num in range(2, 4):
    worker1 = values(num)['RESEARCH_PARAM']
    worker2 = values(num)['RESEARCH_TYPE']
    worker3 = values(num)['RTKID']
    if worker1 == 'подтверждающим методом':
        worker2 = worker2 + ' ' + 'подтверждающим методом'
        worker3 = '10'
        worker1 = ''
    elif worker1 == 'предвариательным методом':
        worker2 = worker2 + ' ' + 'предвариательным методом'
        worker3 = '10'
        worker1 = ''
    elif worker1 == 'предвариательным методом':
        worker2 = worker2 + ' ' + 'предвариательным методом'
        worker3 = '10'
        worker1 = ''
