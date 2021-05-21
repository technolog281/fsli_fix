import logging
import json
import re
from __init__ import et_list_of_words

list_of_mis = ['Б05', 'Б01', 'Б05', 'Б04', 'Б06', 'Б02', 'Б05', 'Б03', 'Б09', 'Б11', 'Б54', 'Б05', 'Б03', 'Б05',
               'Б03', 'Б06', 'Б5', 'Б95', 'Б06', 'Б07', 'Б19', 'Б58', 'ИФА03']
json_file = 'json/list_of_miscodes.json'
list_of_index_mis = []
list_of_index_data = []
list_of_index_nums = []
temp_num = []

for i in list_of_mis:
    index = list_of_mis.index(i)
    num = re.findall(r'\d+', i)
    list_of_index_mis.extend(num)
list_of_index_mis.sort()
list_of_index_mis_result = list(set(list_of_index_mis))


def select_word(list_of_word):
    result_word_list = []
    for pos in list_of_word:
        word = re.sub(r'\d', '', pos)
        result_word_list.append(word)
        result_word_list = list(set(result_word_list))
    for word_in in result_word_list:
        value_word = list(et_list_of_words.values())[list(et_list_of_words.keys()).index(word_in)]
        result_word_list
        if value_word is None:
            logging.warning('Код исследования не найден')
        else:
            logging.info('Код исследования найден')
    return result_word_list


select_word(list_of_mis)


def input_word(word_to_search: str, listf: list):
    for n in listf:
        if int(n) <= 9:
            n_res = f'{word_to_search}0{n}'
            ind = listf.index(n)
            listf[ind] = n_res
        else:
            n_res = f'{word_to_search}{n}'
            ind = listf.index(n)
            listf[ind] = n_res
    return listf


def input_to_json(word: str, list_of_nums):
    if word == 'B':
        word_to_search = 'Б'
        input_word(word_to_search, list_of_nums)
        with open(json_file, 'r') as read_file:
            data = json.load(read_file)['Biochemical research']
        if data is None:
            data = []
    list_of_nums.extend(data)
    list_of_nums = list(set(list_of_nums))

    for b in list_of_nums:
        index_num = list_of_nums.index(b)
        num_num = int(list_of_nums[index_num][1:])
        temp_num.append(num_num)
        temp_num.sort()

    result_list = input_word(word_to_search, temp_num)

    with open(json_file, 'w') as write_file:
        json.dump({'Biochemical research': result_list}, write_file)


input_to_json('B', list_of_index_mis_result)
