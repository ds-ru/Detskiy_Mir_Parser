import requests
import re
import json
from bs4 import BeautifulSoup
import pandas as pd
import time
import os

data = pd.read_excel(r"C:\Users\Дмитрий Сперанский\Desktop\Личное\Detmir\data.xlsx")
rows = data['Код SKU'].dropna()
results = []

for row in rows:
    # URL и заголовки запроса
    url = f"https://www.detmir.ru/search/results/?qt={row}&searchType=auto&searchMode=common"
    # url = f"https://www.detmir.ru/product/index/id/{row}/"
    headers = {
        "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/126.0.0.0 YaBrowser/24.7.0.0 Safari/537.36",
        "Cookie": "ab2_90=ab2_90old90; ab2_33=ab2_33old34; ab2_50=33; ab3_75=ab3_75old75; ab3_33=ab3_33new33; ab3_20=ab3_20_20_3; cc=0; geoCityDMCode=; auid=81725085-21ab-4135-8bdf-c3e3f7c01cee; dmuid=614dae91-90f4-46ed-919e-f07921c65aad; _ym_uid=1710149749708734341; tmr_lvid=a93ef143abc443a79c89868f6ee731f8; tmr_lvidTS=1710149749199; uxs_uid=ad003220-e782-11ee-a32d-437752343407; geoCityDM=%D0%9C%D0%BE%D1%81%D0%BA%D0%B2%D0%B0%20%D0%B8%20%D0%9C%D0%BE%D1%81%D0%BA%D0%BE%D0%B2%D1%81%D0%BA%D0%B0%D1%8F%20%D0%BE%D0%B1%D0%BB%D0%B0%D1%81%D1%82%D1%8C; geoCityDMIso=RU-MOW; adrcid=Ad98trXaBtRvbWqcywem_sw; _ym_d=1729766966; acs_3=%7B%22hash%22%3A%225c916bd2c1ace501cfd5%22%2C%22nextSyncTime%22%3A1730469477049%2C%22syncLog%22%3A%7B%22224%22%3A1730383077049%2C%221228%22%3A1730383077049%2C%221230%22%3A1730383077049%7D%7D; adrdel=1730383077097; mindboxDeviceUUID=0643d98d-6280-4791-b98f-34ed1c47fa06; directCrm-session=%7B%22deviceGuid%22%3A%220643d98d-6280-4791-b98f-34ed1c47fa06%22%7D; _ga_XT948K4GL0=GS1.2.1730383083.11.0.1730383083.0.0.0; _ga=GA1.1.253794051.1710149749; web_proxy=old; uid=CtIBOWeYzy4zQCOFA1/XAg==; _gcl_au=1.1.599317614.1738067759; JSESSIONID=09cc21ce-09b4-4584-a7e7-5847639eef97; detmir-cart=867d402c-0103-4793-ae76-74b9ca8d7a1c; srv_id=cubic-front11-prod; _sp_ses.2b21=*; oneTimeSessionId=7a9430fa-db18-4ae3-bcd4-955d976c8736; detmir-buy_now-cart=cc49145c-9867-44b1-bf9d-f3de9c0a50d8; dm_s=L-09cc21ce-09b4-4584-a7e7-5847639eef97|kH867d402c-0103-4793-ae76-74b9ca8d7a1c|Vj81725085-21ab-4135-8bdf-c3e3f7c01cee|gqcubic-front11-prod|qa867d402c-0103-4793-ae76-74b9ca8d7a1c|-N1710149750107|RK1738067769288|tUcc49145c-9867-44b1-bf9d-f3de9c0a50d8#d_6peyFQCzdrpzKtKnhBcDa86I8yCVyJ83VS4iNjAF0; _ga_87D5G6Z6JP=GS1.1.1738067758.17.1.1738069291.44.0.0; _sp_id.2b21=0df78a0a-e914-4411-bdde-67b47b41fe4b.1710149749.17.1738069292.1734776601.56acf17b-0052-45bb-a5aa-ba5cfb82f815.5026cedc-afc0-4da4-8920-debe808c5121.ff06339a-7256-4877-85d8-668b7bcfd761.1738067760565.18; _ga_MW06XXV5JP=GS1.1.1738067758.17.1.1738069292.0.0.0"  # Замените на ваши куки, если это необходимо
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
                print(suggestion)
                break
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

print(results)

if not os.path.exists(r"C:\Users\Дмитрий Сперанский\Desktop\Личное\Detmir\results.xlsx"):
    with open(r"C:\Users\Дмитрий Сперанский\Desktop\Личное\Detmir\results.xlsx", "w") as file:
        if results:
            output_df = pd.DataFrame(results)
            output_df.to_excel(file, index=False)
        else:
            print("Нет данных для записи.")