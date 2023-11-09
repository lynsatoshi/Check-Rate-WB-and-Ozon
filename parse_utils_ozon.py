import asyncio
import os
import requests
import json

from bs4 import BeautifulSoup
from excel_format import format_excel_data

api_token = os.environ['API_KEY']


async def fetch_data(art, params, max_attempts):
    for attempts in range(1, max_attempts):
        try:
            response = requests.get('https://api.scrapingant.com/v2/general', params=params)
            response.raise_for_status()  # Генерирует исключение, если запрос не успешен
            soup = BeautifulSoup(response.text, 'html.parser')
            json_str = soup.get_text()
            final_json = json.loads(json_str)

            pre_final_json = final_json['widgetStates']['webListReviews-3201466-reviewshelfpaginator-3']

            feedback_elements = json.loads(pre_final_json)

            rating_art = round(feedback_elements['productScore'], 2)
            feedbacks_count = feedback_elements['paging']['total']
            rate_feedbacks = []

            for feedback in feedback_elements['reviews'][:5]:
                rate_feedbacks.append(feedback['content']['score'])

            print(f'Обработан артикул {art}.')

            return {
                'art': art,
                'rating_art': rating_art,
                'feedbacks_count': feedbacks_count,
                'rate_feedbacks': rate_feedbacks
            }
        except Exception as ex:
            print(f'Ошибка с артикулом {art}. Попытка {attempts}. Причина: {ex}')
            await asyncio.sleep(1)  # Даем небольшую паузу перед повторным запросом
    else:
        print(f'Ошибка с артикулом {art}. Достигнуто максимальное число попыток.')
        return {
            'art': art,
            'rating_art': '',
            'feedbacks_count': '',
            'rate_feedbacks': []
        }


async def feedback_from_ozon_api(art_list, max_attempts=4):
    result = []

    for art in art_list:
        params = {
            'url': f'https://www.ozon.ru/api/entrypoint-api.bx/page/json/v2?url=/product/{art}/?layout_container'
                   f'=reviewshelfpaginator&layout_page_index=3&oos_search=false&sh=Rm3ObU9OPg&start_page_id'
                   f'=20062bb852d48452fa35b1e04b788fc3',
            'x-api-key': f'{api_token}',
            'proxy_country': 'RU',
        }

        data = await fetch_data(art, params, max_attempts)
        result.append(data)

    return result


async def excel_result_ozon(feedback_and_rate_result):
    # Создаем папку data_ozon_excel, если ее нет
    data_ozon_excel_path = "data_ozon_excel"
    if not os.path.exists(data_ozon_excel_path):
        os.makedirs(data_ozon_excel_path)

    # Для названия файла
    file_prefix = "ozon"

    # Передаем данные для формирования excel файла
    data_ozon = format_excel_data(feedback_and_rate_result, data_ozon_excel_path, file_prefix)

    return data_ozon
