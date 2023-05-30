import pandas as pd

#Считывание файла
data = pd.read_excel('import_files/impp.xlsx', index_col= None, header= None)

# ЧАСЫ

#Выделение Часов
data_hours_transpose = data.transpose()
condition1 = data_hours_transpose[0].str.contains('Внесите количество контактов МК')
data_hours_transpose.drop(data_hours_transpose[condition1].index, inplace= True)
hours = data_hours_transpose.transpose()

#Первая строка как заголовок
firstColmn = list(hours.loc[0])
hours.columns = firstColmn
hours.drop(index= hours.index[0], axis= 0, inplace= True)

#Изменение Даты
hours['Дата сбора контактов'] = pd.to_datetime(hours['Дата сбора контактов'])
hours["Дата сбора контактов"] = hours["Дата сбора контактов"].dt.strftime("%d.%m.%Y")

hours['Дата'] = pd.to_datetime(hours['Дата'])
hours["Дата"] = hours["Дата"].dt.strftime("%d.%m.%Y")

#Обратное свертывание
hoursForOut = hours.melt(id_vars=['Дата', 'Источник', 'Дата сбора контактов', 'Имя супервайзера'])

#Удаление лишних строк
hoursForOut = hoursForOut.dropna(subset= ["value"])
hoursForOut = hoursForOut[hoursForOut['value'] != "Не работал"]

#Удаление не нужного текста
hoursForOut['variable'] = hoursForOut['variable'].str.replace('Выберите механику и количество отработанных часов x ', '')

#Фрейм для Свода
pivotForObschFile = hoursForOut.drop_duplicates(subset='variable')
pivotForObschFile = pivotForObschFile.groupby('Имя супервайзера')['value'].count()

#Фрейм для Часов NPI
hoursForOutNPI = hoursForOut[hoursForOut['value'] != "Опт"]
hoursForOutNPI = hoursForOutNPI.copy()
hoursForOutNPI.loc[hoursForOutNPI['value'] == "8 часов DSS", 'Механика'] = 2
hoursForOutNPI = hoursForOutNPI.copy()
hoursForOutNPI.loc[hoursForOutNPI['value'] == "4 часа DSS", 'Механика'] = 2
hoursForOutNPI = hoursForOutNPI.copy()
hoursForOutNPI.loc[hoursForOutNPI['value'] == "8 KA", 'Механика'] = 1
hoursForOutNPI = hoursForOutNPI.copy()
hoursForOutNPI.loc[hoursForOutNPI['value'] == "4 KA", 'Механика'] = 1
hoursForOutNPI = hoursForOutNPI.copy()
hoursForOutNPI.loc[hoursForOutNPI['value'] == "8 часов DSS", 'Часы'] = 8
hoursForOutNPI = hoursForOutNPI.copy()
hoursForOutNPI.loc[hoursForOutNPI['value'] == "4 часа DSS", 'Часы'] = 4
hoursForOutNPI = hoursForOutNPI.copy()
hoursForOutNPI.loc[hoursForOutNPI['value'] == "8 KA", 'Часы'] = 8
hoursForOutNPI = hoursForOutNPI.copy()
hoursForOutNPI.loc[hoursForOutNPI['value'] == "4 KA", 'Часы'] = 4

"""
#Фрейм для Часов Опт
hoursForOutOpt = hoursForOut[hoursForOut['value'] == "Опт"]
hoursForOutOpt = hoursForOutOpt.copy()
hoursForOutOpt.loc[hoursForOutOpt['value'] == "Опт", 'Механика'] = 1
hoursForOutOpt = hoursForOutOpt.copy()
hoursForOutOpt.loc[hoursForOutOpt['value'] == "Опт", 'Часы'] = 4
"""
#Запись в файл
with pd.ExcelWriter('out_file/out.xlsx') as writer:
    pivotForObschFile.to_excel(writer, sheet_name='Свод для Общ файла', index=True)

#Запись в файл
with pd.ExcelWriter('out_file/out.xlsx', mode= 'a') as writer:
    hoursForOutNPI.to_excel(writer, sheet_name='Часы для NPI', index=False)
"""
#Запись в файл
with pd.ExcelWriter('out_file/out.xlsx', mode= 'a') as writer:
    hoursForOutOpt.to_excel(writer, sheet_name='Часы для Опт', index=False)
"""

# КОНТАКТЫ

#Выделение Контактов
data_contacts_transpose = data.transpose()
condition2 = data_contacts_transpose[0].str.contains('Выберите механику и количество отработанных ч')
data_contacts_transpose.drop(data_contacts_transpose[condition2].index, inplace= True)
contacts = data_contacts_transpose.transpose()

#Первая строка как заголовок
firstColmn = list(contacts.loc[0])
contacts.columns = firstColmn
contacts.drop(index= contacts.index[0], axis= 0, inplace= True)

#Изменение Даты
contacts['Дата сбора контактов'] = pd.to_datetime(contacts['Дата сбора контактов'])
contacts["Дата сбора контактов"] = contacts["Дата сбора контактов"].dt.strftime("%d.%m.%Y")

contacts['Дата'] = pd.to_datetime(contacts['Дата'])
contacts["Дата"] = contacts["Дата"].dt.strftime("%d.%m.%Y")

#Обратное свертывание
contactsForOut = contacts.melt(id_vars=['Дата', 'Источник', 'Дата сбора контактов', 'Имя супервайзера'])

#Удаление лишних строк
contactsForOut = contactsForOut.dropna(subset= ["value"])

#Удаление не нужного текста
contactsForOut['variable'] = contactsForOut['variable'].str.replace('Внесите количество контактов МК ', '')
contactsForOut['variable'] = contactsForOut['variable'].str.replace(' x Количество контактов', '')

#Разделение текста на столбцы
contactsForOut[['МК', 'Механика']] = contactsForOut['variable'].str.split(' x ', expand = True)

#Удаление ненужного столбца
del contactsForOut['variable']

#Переставление столбцов в нужный порядок
contactsForOut = contactsForOut[['Дата', 'Источник', 'Дата сбора контактов', 'Имя супервайзера', 'МК', 'Механика', 'value']]


#Удаление нулей
contactsForOut = contactsForOut[contactsForOut['value'] != 0]
contactsForOut = contactsForOut.copy()

#Фрейм для NPI
contactsForOutNPI = contactsForOut[(contactsForOut['Механика'] != "10+2") & (contactsForOut['Механика'] != "Retailer")]
contactsForOutNPI = contactsForOutNPI.copy()
"""
#Фрейм для Опт
contactsForOutOpt = contactsForOut[(contactsForOut['Механика'] == "10+2") | (contactsForOut['Механика'] == "Retailer")]
contactsForOutOpt = contactsForOutOpt.copy()
"""


#Запись в файл
with pd.ExcelWriter('out_file/out.xlsx', mode= 'a') as writer:
    contactsForOutNPI.to_excel(writer, sheet_name='Контакты NPI', index=False)
"""
#Запись в файл
with pd.ExcelWriter('out_file/out.xlsx', mode= 'a') as writer:
    contactsForOutOpt.to_excel(writer, sheet_name='Контакты Опт', index=False)
"""