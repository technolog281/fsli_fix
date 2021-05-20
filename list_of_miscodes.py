import json

list_of_mis = ['Б05', 'Б03', 'Б05', 'Б03', 'Б06', 'Б02', 'Б05', 'Б03', 'Б09', 'Б11', 'Б54', 'Б02']
json_file = 'json/list_of_miscodes.json'
list_of_index_mis = []
list_of_index_data = []

for i in list_of_mis:
    index = list_of_mis.index(i)
    num = int(list_of_mis[index][1:])
    list_of_index_mis.append(num)
list_of_index_mis.sort()
list_of_index_mis_result = list(set(list_of_index_mis))


def input_to_json(word: str, list_of_nums):
    if word == 'B':
        word_to_search = 'Б'
        for n in list_of_nums:
            if n <= 9:
                n_res = f'{word_to_search}0{n}'
                ind = list_of_nums.index(n)
                list_of_nums[ind] = n_res

            else:
                n_res = f'{word_to_search}{n}'
                ind = list_of_nums.index(n)
                list_of_nums[ind] = n_res

    with open(json_file, 'r') as read_file:
        data = json.load(read_file)['Biochemical research']
        if data is None:
            data = []
    data = list(set(data))
    # for q in data:
    #     data_index = data.index[q]

    return data
    # with open(json_file, 'w') as write_file:
    #     json.dump({'Biochemical research': data}, write_file)


print(input_to_json('B', list_of_index_mis_result))
