import requests
import json
from bs4 import BeautifulSoup
import pandas as pd


data = pd.read_excel(r"data.xlsx")
rows = data['Код SKU'].dropna()
results = []

for row in rows:
    # URL и заголовки запроса
    url = f"https://www.detmir.ru/search/results/?qt={row}&searchType=auto&searchMode=common"
    headers = {
        "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/126.0.0.0 YaBrowser/24.7.0.0 Safari/537.36",
        "Cookie": "your_cookie_here"  # Замените на ваши куки, если это необходимо
    }

    try:
        # Отправка GET-запроса к странице
        response = requests.get(url, headers=headers)

        # Проверка успешного ответа
        if response.status_code == 200:
            # Используем BeautifulSoup для обработки HTML
            soup = BeautifulSoup(response.text, 'html.parser')

            # Поиск всех <script> тегов
            script_datas = soup.find_all("script")
            for script_data in script_datas:
                if script_data.string and 'appData' in script_data.string:
                    try:
                        json_text = script_data.string
                        # Извлечение данных, используя простое регулярное выражение
                        json_text = json_text.split('window.appData = JSON.parse(')[-1][:-2]  # Получаем часть после JSON.parse
                        json_text = json_text.replace(r'\\\"', '')
                        json_text = json_text.replace(r'\\"', '')
                        json_text = json_text.replace(r'\"', '"')  # Убираем экранирование
                        json_text = json_text[1:]

                        app_data = json.loads(json_text)  # Преобразуем в словарь Python
                        break  # Завершаем после первого совпадения
                    except Exception as e:
                        print("Ошибка при обработке JSON:", e)

            # Поиск нужного товара по SKU и добавление данных в результаты
            for suggestion in app_data['search']['data']['suggestions']:
                if suggestion['type'] == 'product':
                    product = suggestion['filter'].get('product')
                    if product and product.get('code') == row:
                        title_found = product['title']
                        print(title_found)
                        # Добавляем результат в список
                        results.append({'Код SKU': row, 'Название': title_found})
                        break
        else:
            print(f"Ошибка при загрузке страницы: {response.status_code}")

    except Exception as e:
        print(f"Ошибка при обработке SKU {row}: {e}")

if results:
    output_df = pd.DataFrame(results)
    output_df.to_excel('results.xlsx', index=False)
else:
    print("Нет данных для записи.")