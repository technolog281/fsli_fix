from __init__ import *


def ident_select(col):
    new_ident = []
    for i in range(col):
        wtdb_cursor.execute('select newID()')
        new_ident.append(wtdb_cursor.fetchone()[0])
    return new_ident


for x in range(len(ident_select(8))):
    print(ident_select(8)[x])
