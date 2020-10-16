#загружаем модули для работы
import pandas as pd
import openpyxl
from collections import Counter

#загружаем эксель файл с продажами в датафрейм панд
df = pd.read_excel('/content/logs.xlsx', sheet_name='log')

#переименовываем колонки в английский язык для удобства работы с датафреймом
df.columns = ['IP', 'Sex', 'Age', 'Browser', 'Version', 'Site_time', 'Visit_date', 'Goods']

#находим семь популярных браузером путем группировки датафрейма по столбцу браузер и сортировки значений по убыванию
top_seven_browsers = df.groupby(by='Browser')['Browser'].count().sort_values(ascending=False)[:7]

#создаем словарь с популярными браузерами для удобства работы
top_seven_browsers = top_seven_browsers.to_dict()

def sales_list(data):
    """Функция для создания списка купленных товаров
    путем перебора позиций в столбце Goods
    data: датафрейм с данными"""
    sales = []
    for row in data.Goods.str.split(','):
      for good in row:
        sales.append(good)
    return sales

#создаем словарь и проводим замену в столбце 'Пол' по словарю для удобства работы
MF = {'м':'M', 'ж':'F'}
df.Sex = df.Sex.map(MF)

#проверяем корректность и полноту замены
print(df.Sex.isna().sum())

#отделяем мужчин и женщин в отдельные датафреймы
males = df[df.Sex == 'M']
females = df[df.Sex == 'F']

#создаем общий список покупок и отдельно по мужчинам и женщинам
sales = sales_list(df)
male_sales = sales_list(males)
female_sales = sales_list(females)

#создаем общий словарь с покупками
top_sales_dict = Counter(sales)

#отделяем 7 самых популярных по продажам товаров
top_seven_goods = top_sales_dict.most_common(7)

#определяем самые популярные и непопулярные товары у мужчин
male_sales_dict = Counter(male_sales)
top_one_males = male_sales_dict.most_common(1)[0]
last_one_males = male_sales_dict.most_common(len(male_sales_dict))

#определяем самые популярные и непопулярные товары у женщин
female_sales_dict = Counter(female_sales)
top_one_females = female_sales_dict.most_common(1)[0]
last_one_females = female_sales_dict.most_common(len(male_sales_dict))

#открываем файл эксель для записи результатов в лист1
wb = openpyxl.load_workbook(filename='/content/report.xlsx')
sheet = wb['Лист1']

#записываем самые популярные браузеры и количество заходов для каждого в ячейки А5-А11, B5-B11
counter_1 = 5
sum_v = 0
for k, v in top_seven_browsers.items():
    sheet['A'+ str(counter_1)] = k
    sheet['B' + str(counter_1)] = v
    counter_1 += 1
    sum_v + int(v)

sheet['B12'] = sum(top_seven_browsers.values())

#записываем самые популярные товары и количество покупок для каждого в ячейки А19-А25, B19-B25
counter_2 = 19
sum_top = 0
for top in top_seven_goods:
    sheet['A'+ str(counter_2)] = top[0]
    sheet['B' + str(counter_2)] = top[1]
    counter_2 += 1
    
sheet['B26'] = sum(top_seven_goods)

#записываем самые популярные и непопулярные товары у мужчин и женщин
sheet['B31'] = top_one_males[0]
sheet['B32'] = top_one_females[0]
sheet['B33'] = last_one_males[-1][0]
sheet['B34'] = last_one_females[-1][0]

#сохраняем книгу эксель
wb.save('report.xlsx')
