import os
import pandas as pd
import numpy as np

folder = input('Введите название папки с файлами: ')
files = os.listdir(folder)

# создание списка файлов Excel в папке == folder
excel_files =[]
for file in files:
    if file.endswith('.xlsx'):
        excel_files.append(file)

# получение списка дат
data_new = []
for i in range(len(excel_files)):
    data = excel_files[i][excel_files[i].find('(') + 1: excel_files[i].rfind(')')] #срез даты из имени файла
    data_new.append(data)

#val = 'Разъем RJ45, розетка на панель'
val = input('Что ищем? ')

#sheet = 'Склад Офис'
sheet = input('На какой вкладке? ')

#price = 'Цена'
price = input('Название столбца с ценой? ')

#создание пустой итоговой таблицы для результата
df_res = pd.DataFrame()

for i in range(len(excel_files)):
    filename = os.path.join(folder, excel_files[i]) # объединяет путь и имя файла, где i - порядковый индекс файла
    df = pd.read_excel(filename, sheet_name=sheet) # читаем файл filename и вкладку из переменной sheet
    idx = np.where(df == val) #вычисляем номер строки в котором находится искомое значение = val
    value = df.loc[idx[0], price] # находим цену в столбце price
    value_new = value.tolist() #преобразование серии Pandas в Список
    print(data_new[i], val, value_new[0])
    new_row = [data_new[i], val, value_new[0]] # создаём строку для внесения в новую таблицу
    new_df = pd.DataFrame([new_row], columns=['Дата', 'Наименование', 'Цена'])  # помещаем new row в созданную новую таблицу
    df_res = pd.concat([df_res, new_df], axis=0, ignore_index=True) # соединяем итоговую таблицу с новой

file_result_name = r'\output.xlsx'
file_result = os.path.join(folder,file_result_name)

df_res.to_excel(file_result)
print('ИТОГОВАЯ ТАБЛИЦА СОХРАНЕНА!')
print(f'Файл здесь: {file_result}')

input()
