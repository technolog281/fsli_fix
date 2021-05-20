from __init__ import *

to_search_book = load_workbook('xlsx/to_search.xlsx')
ts_sheet = to_search_book['Sheet1']


def research_select():
    df = pd.DataFrame()
    for num in range(2, 6):
        research_name = values(num)['ResearchName']
        param_name = values(num)['ParamName']
        vals = values(num)

        if 'Массов' and 'подтверждающим' in param_name:
            vals['ResearchName'] = research_name + ', подтверждающим методом'
            df = df.append(vals, ignore_index=True)
        elif 'Массов' and 'предварительным' in param_name:
            vals['ResearchName'] = research_name + ', предварительным методом'
            df = df.append(vals, ignore_index=True)
        else:
            df = df.append(vals, ignore_index=True)
            logging.warning('Not Found')

    set_list = set(df['ResearchName'].values.tolist())
    return set_list

# for i in set_list:
#     print(i)

