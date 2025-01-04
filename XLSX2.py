import pandas as pd

sales_data = pd.ExcelFile('Воронка.xlsx')
data_before = []
data = {}
title_list = []
for i in sales_data.sheet_names:
    if i == 'Инфо' or i == 'Фильтры':
        continue
    else:
        data_list = pd.read_excel('Воронка.xlsx', i, skiprows=1)
        data_before.append(data_list)
        worksheets_data = pd.concat(data_before)
        title_excel = pd.read_excel('Воронка.xlsx', i)

title = title_excel.iloc[0].tolist()
list_columns = [0, 9, 11, 13, 15, 17, 19, 21, 23, 25, 27, 28, 30]

for n in list_columns:
    data_in_excel = worksheets_data.iloc[:, n].tolist()
    data[title[n]] = data_in_excel

data = pd.DataFrame.from_dict(data)
data.to_excel('Воронка новая.xlsx', sheet_name='Воронка', index=False)
print("Готово")