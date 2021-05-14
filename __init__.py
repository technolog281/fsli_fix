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

wb = load_workbook('./names_v6.xlsx')
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


belbol = pyodbc.connect("DRIVER={SQL Server};"
                        "SERVER=DROID1347\SQL_MIS;"
                        "DATABASE=amur_mis_belbol;"
                        "UID=sa;"
                        "PWD=Apacheserver1390")
belbol_cursor = belbol.cursor()
# amur_mis_belbol = amur_mis_belbol_conn.setencoding(encoding='utf-8')


script_base = pyodbc.connect('DRIVER={SQL Server};'
                             'SERVER=DROID1347\SQL_MIS;'
                             'DATABASE=script_base;'
                             'UID=sa;'
                             'PWD=Apacheserver1390')
sb_cursor = script_base.cursor()
# script_base = amur_mis_belbol_conn.setencoding(encoding='utf-8')


# app = FastAPI()
# eng = create_engine('DRIVER={ODBC Driver 17 for SQL Server};'
#                     'SERVER=DROID1347\SQL_MIS;'
#                     'DATABASE=amur_mis_belbol;'
#                     'UID=sa;'
#                     'PWD=Apacheserver1390')
# import urllib
# eng = 'DRIVER={SQL Server};DATABASE=amur_mis_belbol;SERVER=DROID1347\SQL_MIS;UID=sa;PWD=Apacheserver1390'
# eng = urllib.quote_plus(eng)
# eng = "mssql+pyodbc:///?odbc_connect=%s" % eng
#
# eng = create_engine('mssql+pyodbc://sa:Apacheserver1390@127.0.0.1:2044/SQL_MIS/amur_mis_belbol?driver=SQL+Server', pool_pre_ping=True)
# jdbc:sqlserver://127.0.0.1\SQL_MIS:2044;database=amur_mis_belbol
# with eng.connect() as con:
#     rs = con.execute('select * from lbr_Laboratory')
#
#     data = rs.fetchone()[0]
#
#     print
#     "Data: %s" % data

# qwe = "select * from lbr_ResearchTypeParam"
#
#
# def read_root(ask):
#     cursor.execute(ask)
#     row = cursor.fetchall()
#     # nums = list(range(1, len(row) + 1))
#     # new_dict = {nums: row for nums, row in zip(nums, row)}
#     return row
#
#
# print(type(read_root(qwe)))
#
# def answ():
#     html = open("C:\\Users\\Droid1347\\PycharmProjects\\pythonProject\\html\\_form.html").read()
#     template = Template(html)
#     return template.render(content=html, status_code=200)
#
#
# @app.get("/items/", response_class=HTMLResponse)
# async def read_items():
#     return answ()
#
#
# if __name__ == '__main__':
#     run(app, host='localhost', port=8080)
