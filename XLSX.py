import pandas as pd

sheet_name = 'Товары'

sales_data = pd.read_excel('Воронка.xlsx', engine='openpyxl', sheet_name=sheet_name)

list_columns = [0, 9, 11, 13, 15, 17, 19, 21, 23, 25, 27, 28, 30]
data = {}
for i in list_columns:
    data_in_excel = pd.DataFrame(sales_data.iloc[:,i].tolist())
    data[data_in_excel[0][0]] = data_in_excel[0][1:]

data = pd.DataFrame.from_dict(data)

data.to_excel('Воронка новая.xlsx')