import pandas as pd
import glob
import datetime

count = 0
list_remove = []
new_file_name = f"Воронка новая {datetime.date.today()}.xlsx"

list_files = glob.glob("*.xlsx")

remove_title = "Воронка новая"                               # Проверяю, чтобы не было новых воронок

for file_name in list_files:                                 # Если есть новые воронки, удаляю из списка для конвертации
    if remove_title in file_name:
        list_remove.append(file_name)                        # Добавляю в список для удаления file_name
for file_remove in list_remove:
    if file_remove in list_files:
        list_files.remove(file_remove)

if len(list_files) == 0:                                     # Проверяем если нет воронок
    print("Нет ни одной воронки")
    exit()

for file_name in list_files:                                 # открываю каждую воронку
    sales_data = pd.ExcelFile(file_name)
    print(file_name)
    data_before = []
    title_list = []

    for i in sales_data.sheet_names:
        if i == 'Метрики' or i == 'Фильтры' or i == 'Общая информация':                    # исключаю 'Инфо' и 'Фильтры'
            continue
        else:                                                # для остальных листов
            data_list = pd.read_excel(file_name, i, skiprows=1)         # читаю воронку без первой строки
            data_before.append(data_list)                               # добавляю данные листа в список
            worksheets_data = pd.concat(data_before)                    # объединяю данные листов
            title_excel = pd.read_excel(file_name, i)                   # читаю названия столбцов

    title = title_excel.iloc[0].tolist()                                # добавляю названия столбцов в список
    list_columns = [0, 10, 12, 16, 18, 20, 22, 24, 26, 28, 30, 31, 33]   # список нужных столбцов

    data = {}
    for n in list_columns:                                              # перебираю столбцы
        data_in_excel = worksheets_data.iloc[:, n].tolist()             # добавляю данные столбца в список
        data[title[n]] = data_in_excel                                  # добавляю названия столбцов в словарь

    data = pd.DataFrame.from_dict(data)

    if "ИРИНА" in file_name.upper():
        new_file_name = f"Воронка новая (Ирина) {datetime.date.today()}.xlsx"
    elif "НАТАЛЬЯ" in file_name.upper():
        new_file_name = f"Воронка новая (Наталья) {datetime.date.today()}.xlsx"
    elif "ПАВЕЛ" in file_name.upper():
        new_file_name = f"Воронка новая (Павел) {datetime.date.today()}.xlsx"
    elif "АЛЕКСЕЙ" in file_name.upper():
        new_file_name = f"Воронка новая (Алексей) {datetime.date.today()}.xlsx"
    elif count == 0:
        new_file_name = f"Воронка новая {datetime.date.today()}.xlsx"
        count += 1
    elif count > 0:
        new_file_name = f"Воронка новая ({count}) {datetime.date.today()}.xlsx"
        count += 1

    data.to_excel(new_file_name, sheet_name='Воронка', index=False)
    print("Готово", new_file_name)