import json
from openpyxl import load_workbook


tb = load_workbook('sample.xlsx')
tb_names = tb.sheetnames

alph = 'ABCDEFGHIJKLMNOPQRSTUVW'  # алфавит для перебора по столбцам таблицы
keys = ['C', 'M', 'Y', 'K', 'Pantone']  # ключи полей

pantone = []


def parser():
    for item in tb_names:
        page = tb[item]
        for char in alph:
            obj = []
            for i in range(1, 21):
                data_str = page[char + str(i)].value
                if data_str and not ("PANTONE" in data_str):
                    for a in data_str.split(" "):
                        obj.append(int(a.split(":")[1]))
                elif data_str:
                    obj.append(data_str)
                if len(obj) == 5:
                    obj.append(obj.pop(0))
                    pantone.append(dict(zip(keys, obj)))
                    obj = []

    with open('data.json', 'w') as data_f:
        json.dump(pantone, data_f)


if __name__ == "__main__":
    parser()
