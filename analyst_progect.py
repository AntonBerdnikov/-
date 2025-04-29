# 1. Импорт библиотек и загрузка данных

import pandas as pd
import numpy as np
import matplotlib.pyplot as plt

df = pd.read_excel('C:/Users/Антон/Downloads/SalesDashboard.xlsx') # Извлечение данных из excel
pd.set_option('display.max_columns', None) # Получение всех столбцов

# 2. Обзор таблицы. Предварительная обработка данных

# Определяем, за что отвечают каждая колонка в таблице и тип данных
# Дата - Дата продажи. Имеет тип: datetime
# Товар - Категории товаров, которые были проданыю Имеет тип: object
# Менеджер - Фамилия менеджера, который продал товар. Имеет тип: object
# Регион - В каком регионе была продажаю Имеет тип: object
# Выручка - Не чистый заработок в продажи. Имеет тип: float
# Вал_прибыль - Разница между выручкой и себестоимостью товара. Имеет тип: float
# Прибыль - Чистый заработок с продажи товара. Имеет тип: int
# Выручка_план - План компании по выручке за продажу. Имеет тип: float

df[['Менеджер', 'Регион', 'Товар']] = df[['Менеджер', 'Регион', 'Товар']].astype('str') # Изменим тип object на тип str

cnt_str = df.count() # Опледеляем количество строк. Всего строк 5544
df_new = df.dropna() # Удалим все строки с любыми пропусками
df_dupl = df[df_new.duplicated()] # Проверяем данные на дубликаты
df_new = df_new.drop_duplicates() # Удаляем все дубликаты и переприсваеваем DataFrame
cnt_str_new = df_new.count()# Определяем количество строк после очистки данных. Всего строк 5544.Количество строк не изменилось


# Формирование когорт:
# 1) Выполнение годового плана по выручке
# 2) Прибыль по регионам
# 3) Топ 10 товаров по продажам за год
# 4) Выпонение плана среди менеджеров
# 5) Рентабельность по валовой прибыли
# 1
plan_revenue_year = df_new.groupby(df_new['Дата'].dt.year)[['Выручка', 'Выручка_план']].sum().round()
plan_revenue_year['План_процент'] = plan_revenue_year['Выручка'] / plan_revenue_year['Выручка_план'] * 100
print(plan_revenue_year)

# 2
profit_reg = df_new.groupby('Регион')['Прибыль'].sum().round()

# 3
top_product = df_new.groupby('Товар')['Выручка'].sum().round().sort_values(ascending=False)

# 4
plan_revenue_men = df_new.groupby('Менеджер')[['Выручка', 'Выручка_план']].sum().round()
plan_revenue_men['План_процент'] = round(plan_revenue_men['Выручка'] / plan_revenue_men['Выручка_план'] * 100, 2)
plan_revenue_men = plan_revenue_men.sort_values(by='План_процент', ascending=False)

# 5
profitability = df_new.groupby('Товар')[['Выручка', 'Вал_прибыль']].sum().round()
profitability['Рентабельность'] = round(profitability['Вал_прибыль'] / profitability['Выручка'] * 100, 2)
profitability = profitability.sort_values(by='Рентабельность', ascending=False)

# Добавление графиков
"""""
plt.figure(figsize=(10, 6))
plan_revenue_year[['Выручка', 'Выручка_план']].plot(kind='bar')
plt.title('Выполнение годового плана по выручке')
plt.show()

profit_reg.plot(kind='bar')
plt.title('Прибыль по регионам')
plt.show()

top_product.head(10).plot(kind='bar')
plt.title('Топ 10 товаров по продажам')
plt.show()
"""""
''''
# Добавление сводной информации
print("Сводка по выполнению годового плана:")
print(plan_revenue_year)

print("\nТоп-10 товаров по выручке:")
print(top_product.head(10))

print("\nЛучшие менеджеры по выполнению плана:")
print(plan_revenue_men.head(10))

print("\nСамые рентабельные товары:")
print(profitability.head(10))
'''''