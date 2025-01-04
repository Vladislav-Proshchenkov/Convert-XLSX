import pandas as pd
import glob

data_before = []
data = {}
title_list = []

list_files = glob.glob("*.xlsx")

remove_title = "Воронка новая"           # Проверяю, чтобы не было новых воронок

for file_name in list_files:             # Если есть новые воронки, удаляю из списка для конвертации
    if remove_title in file_name:
        list_files.remove(file_name)

if len(list_files) == 0:                 # Проверяем если нет воронок
    print("Нет ни одной воронки")
    exit()

for file_name in list_files:
    sales_data = pd.ExcelFile(file_name)
    print(file_name)

for i in sales_data.sheet_names:
    if i == 'Инфо' or i == 'Фильтры':
        continue
    else:
        data_list = pd.read_excel(file_name, i, skiprows=1)
        data_before.append(data_list)
        worksheets_data = pd.concat(data_before)
        title_excel = pd.read_excel(file_name, i)

title = title_excel.iloc[0].tolist()
list_columns = [0, 9, 11, 13, 15, 17, 19, 21, 23, 25, 27, 28, 30]

for n in list_columns:
    print(n)
    data_in_excel = worksheets_data.iloc[:, n].tolist()
    data[title[n]] = data_in_excel

data = pd.DataFrame.from_dict(data)

data.to_excel('Воронка новая.xlsx', sheet_name='Воронка', index=False)
print("Готово")