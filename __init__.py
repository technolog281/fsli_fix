from typing import Dict, Any
from fastapi import FastAPI
from uvicorn import run
from jinja2 import Template
from fastapi.responses import HTMLResponse
import pyodbc
import sqlalchemy as sa
from sqlalchemy import create_engine
from openpyxl import load_workbook
from search_last_list import last_page
import pandas as pd

wb = load_workbook('./names_v7.xlsx')
sheet = wb['Sheet1']


def values(num):
    result = {'NUM': sheet[f'A{num}'].value,
              'RESEARCH_TYPE': sheet[f'B{num}'].value,
              'RESEARCH_PARAM': sheet[f'C{num}'].value,
              'NMU': sheet[f'D{num}'].value,
              'BIOM': sheet[f'E{num}'].value,
              'UNIT': sheet[f'F{num}'].value,
              'LIS_CODE': sheet[f'G{num}'].value,
              'MAT_CODE': sheet[f'H{num}'].value,
              'LAB_TYPE': sheet[f'I{num}'].value}
    return result


# belbol = pyodbc.connect("DRIVER={SQL Server};"
#                         "SERVER=DROID1347\SQL_MIS;"
#                         "DATABASE=amur_mis_belbol;"
#                         "UID=sa;"
#                         "PWD=Apacheserver1390")
# belbol_cursor = belbol.cursor()
# amur_mis_belbol = amur_mis_belbol_conn.setencoding(encoding='utf-8') ## -- Декодирование


# script_base = pyodbc.connect('DRIVER={SQL Server};'
#                              'SERVER=DROID1347\SQL_MIS;'
#                              'DATABASE=script_base;'
#                              'UID=sa;'
#                              'PWD=Apacheserver1390')
# sb_cursor = script_base.cursor()
# script_base = amur_mis_belbol_conn.setencoding(encoding='utf-8') ## -- Декодирование
