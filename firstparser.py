# ============================================================
# БЛОК 1. ИМПОРТЫ
# ============================================================
# requests — библиотека для HTTP-запросов.
import requests
# BeautifulSoup — для разбора HTML.
from bs4 import BeautifulSoup
# pandas — для таблиц и экспорта в Excel.
import pandas as pd
# os — для работы с путями к файлам.
import os


# ============================================================
# БЛОК 2. НАСТРОЙКИ И ПУТЬ ДЛЯ СОХРАНЕНИЯ
# ============================================================
# Определяем папку, где лежит сам скрипт (а не откуда его запускают).
# __file__ — это путь к текущему .py файлу.
# os.path.abspath — делает путь абсолютным.
# os.path.dirname — берёт только папку, отбрасывая имя файла.
script_dir = os.path.dirname(os.path.abspath(__file__))

# Склеиваем путь к папке скрипта с именем файла.
# os.path.join сам подставит правильный разделитель (/ или \) под вашу ОС.
file_path = os.path.join(script_dir, "result.xlsx")
print(f"Файл будет сохранён в: {file_path}")


# ============================================================
# БЛОК 3. ФОРМИРОВАНИЕ URL
# ============================================================
# Название предмета. Должно точно совпадать с названием в Steam.
item_name = "AK-47 | Neon Revolution (Field-Tested)"

# URL-кодирование: пробелы и | нельзя использовать в URL.
encoded_name = item_name.replace(" ", "%20").replace("|", "%7C")

# Базовый адрес страницы предмета.
base_url = f"https://steamcommunity.com/market/listings/730/{encoded_name}"

# Адрес эндпоинта /render/, который возвращает ТОЛЬКО HTML с лотами.
# Именно этот адрес вызывает JavaScript на странице, чтобы подгрузить лоты.
render_url = (
    f"{base_url}/render/"
    f"?query=&start=0&count=10"     # первые 10 лотов
    f"&country=RU"                   # страна — влияет на валюту
    f"&language=english"             # язык описаний
    f"&currency=5"                   # 5 — российский рубль
)


# ============================================================
# БЛОК 4. HTTP-ЗАПРОС С ЗАГОЛОВКАМИ
# ============================================================
# Заголовки (headers) — это «паспорт» клиента, который он показывает серверу.
# Без User-Agent Steam может вернуть 403 (запрещено) или пустой ответ.
headers = {
    # User-Agent говорит серверу, что мы «браузер Chrome на Windows».
    "User-Agent": (
        "Mozilla/5.0 (Windows NT 10.0; Win64; x64) "
        "AppleWebKit/537.36 (KHTML, like Gecko) "
        "Chrome/120.0.0.0 Safari/537.36"
    ),
    # Accept-Language — предпочитаемый язык ответа.
    "Accept-Language": "ru-RU,ru;q=0.9,en;q=0.8",
}

# Отправляем GET-запрос с заголовками.
response = requests.get(render_url, headers=headers)

# Проверяем статус-код ответа.
if response.status_code != 200:
    print(f"Ошибка при запросе: {response.status_code}")
    print("Ответ сервера:", response.text[:500])  # первые 500 символов для отладки
    exit()


# ============================================================
# БЛОК 5. ПАРСИНГ HTML
# ============================================================
# Превращаем HTML-текст в дерево BeautifulSoup.
soup = BeautifulSoup(response.text, "html.parser")

# Ищем все строки лотов. Теперь они точно есть — мы получили правильный HTML.
items = soup.find_all("div", class_="market_listing_row")

# Отладочный вывод: сколько лотов найдено.
print(f"Найдено строк с лотами: {len(items)}")

# Если лотов нет — Steam что-то поменял или заблокировал. Сохраним ответ для анализа.
if not items:
    debug_path = os.path.join(script_dir, "debug_response.html")
    with open(debug_path, "w", encoding="utf-8") as f:
        f.write(response.text)
    print(f"Лоты не найдены. HTML сохранён в {debug_path} для анализа.")
    exit()


# ============================================================
# БЛОК 6. СБОР ДАННЫХ
# ============================================================
data = []

for item in items:
    # Цена с комиссией (то, что заплатит покупатель).
    price_span = item.find("span", class_="market_listing_price_with_fee")

    # Название предмета.
    name_span = item.find("span", class_="market_listing_item_name")

    # Пропускаем строку, если нет цены ИЛИ названия.
    if not price_span or not name_span:
        continue

    price = price_span.text.strip()
    name = name_span.text.strip()

    # Дополнительно попробуем вытащить «цену без комиссии» (получит продавец).
    # Это полезно для анализа, но можно и убрать.
    price_no_fee_span = item.find("span", class_="market_listing_price_no_fee")
    price_no_fee = price_no_fee_span.text.strip() if price_no_fee_span else "—"

    data.append({
        "Название": name,
        "Цена с комиссией": price,
        "Цена без комиссии": price_no_fee,
    })


# ============================================================
# БЛОК 7. СОЗДАНИЕ ТАБЛИЦЫ
# ============================================================
df = pd.DataFrame(data)


# ============================================================
# БЛОК 8. СОХРАНЕНИЕ В EXCEL
# ============================================================
# file_path — абсолютный путь, который мы построили в блоке 2.
df.to_excel(file_path, index=False)

print(f"Готово! Файл создан. Записей: {len(data)}")
print(f"Полный путь к файлу: {file_path}")


# ============================================================
# БЛОК 9. ФИНАЛЬНАЯ ПРОВЕРКА
# ============================================================
if os.path.exists(file_path):
    size_kb = os.path.getsize(file_path) / 1024  # размер в килобайтах
    print(f"Файл существует. Размер: {size_kb:.2f} КБ")
else:
    print("Файл почему-то не создан. Проверьте права на запись в папку.")