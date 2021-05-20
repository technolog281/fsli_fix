from __init__ import *

wb = load_workbook('xlsx/names_v7.xlsx')
sheet = wb['Sheet1']
df = pd.DataFrame(
    columns=['NUM', 'RESEARCH_TYPE', 'RESEARCH_PARAM', 'NMU', 'BIOM', 'UNIT', 'LIS_CODE', 'MAT_CODE', 'LAB_TYPE'])


def count(rp_split):
    for i in rp_split:
        if 'метод' in i:
            return rp_split.index(i)


for num in range(2, 23458):
    val = values(num)
    rp_split = val['RESEARCH_PARAM'].split(' ')
    pos = count(rp_split)
    len_split = len(rp_split)
    if pos is not None and pos + 1 == len_split:
        rp = rp_split[pos - 1] + ' ' + rp_split[pos]
        val['RESEARCH_PARAM'] = rp
        df = df.append(val, ignore_index=True, sort=False)
    elif pos is not None and pos + 1 != len_split:
        rp = ''
        for i in range(pos, len_split):
            rp = rp + ' ' + rp_split[i]
            val['RESEARCH_PARAM'] = rp
        df = df.append(val, ignore_index=True, sort=False)
    else:
        df = df.append(val, ignore_index=True, sort=False)


df.to_excel('probe.xlsx')




