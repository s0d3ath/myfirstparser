import requests
from bs4 import BeautifulSoup
import pandas as pd

# 1. Заходим на сайт
url = "https://steamcommunity.com/id/s0d3ath/"
response = requests.get(url)

# 2. Ищем нужные данные
soup = BeautifulSoup(response.text, "html.parser")
items = soup.find_all("div", class_="miniprofile.hover")  # ищем блоки товаров

# 3. Складываем в таблицу
data = []
for item in items:
    name = item.find("h3").text
    price = item.find("span", class_="price").text
    data.append({"Название": name, "Цена": price})

# 4. Сохраняем в Excel
df = pd.DataFrame(data)
df.to_excel("result.xlsx", index=False)

print("Готово! Файл result.xlsx создан.")
import os
print("Текущая папка:", os.getcwd())
print("Файл существует:", os.path.exists("result.xlsx"))
if os.path.exists("result.xlsx"):
    print("Полный путь:", os.path.abspath("result.xlsx"))