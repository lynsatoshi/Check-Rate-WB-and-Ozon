import os
import requests
import json

from bs4 import BeautifulSoup
from excel_format import format_excel_data


async def parse_api_wb(art_list):
    result = []
    for art in art_list:
        try:
            # Делаем api запрос на получение информации со страницы артикула
            url = f'https://card.wb.ru/cards/detail?appType=1&curr=rub&dest=123585944&regions=80,38,83,4,64,33,68,70,' \
                  f'30,40,86,69,22,1,31,66,110,48,71,114&spp=30&nm={art}'
            response = requests.get(url)

            # Достаем из запроса общую оценку товара
            soup = BeautifulSoup(response.text, 'html.parser')

            json_str = soup.get_text()

            final_json = json.loads(json_str)

            rating_art = final_json['data']['products'][0]['reviewRating']
            feedbacks_count = final_json['data']['products'][0]['feedbacks']
            feedback_id = final_json['data']['products'][0]['root']

            result.append({
                'art': art,
                'rating_art': rating_art,
                'feedbacks_count': feedbacks_count,
                'feedback_id': feedback_id
            })

        except Exception as e:
            print(f"Похоже запрашиваемого артикула нет {art}: {str(e)}")
            result.append({
                'art': art,
                'rating_art': None,
                'feedbacks_count': None,
                'feedback_id': None
            })

    return result


def get_feedbacks_list(feedback_id):
    try:
        url = f'https://feedbacks1.wb.ru/feedbacks/v1/{feedback_id}'
        response = requests.get(url)
        feedbacks_list = response.json()['feedbacks']
    except Exception as ex:
        try:
            url = f'https://feedbacks2.wb.ru/feedbacks/v1/{feedback_id}'
            response = requests.get(url)
            feedbacks_list = response.json()['feedbacks']
        except Exception as ex:
            feedbacks_list = []
            print(f'Ошибка с запросом для feedback_id {feedback_id}: {ex}')
    return feedbacks_list


async def feedback_from_wb_api(result_api):
    result = []
    for feedback in result_api:
        if feedback.get('feedback_id') is not None:
            art = feedback['art']
            feedback_id = feedback['feedback_id']

            rate_feedbacks = []

            try:
                url = f'https://feedbacks1.wb.ru/feedbacks/v1/{feedback_id}'
                response = requests.get(url)
                feedbacks_list = response.json()['feedbacks']

                # Сортируем feedbacks_list по параметру feedback_createdDate
                feedbacks_list.sort(key=lambda x: x['createdDate'], reverse=True)

                for grade in feedbacks_list[:5]:
                    feedback_rate = grade['productValuation']
                    rate_feedbacks.append(feedback_rate)

                result.append({
                    'art': art,
                    'rating_art': feedback.get('rating_art'),
                    'feedbacks_count': feedback.get('feedbacks_count'),
                    'rate_feedbacks': rate_feedbacks
                })
            except Exception as ex:
                try:
                    url = f'https://feedbacks2.wb.ru/feedbacks/v1/{feedback_id}'
                    response = requests.get(url)
                    feedbacks_list = response.json()['feedbacks']

                    # Сортируем feedbacks_list по параметру feedback_createdDate
                    feedbacks_list.sort(key=lambda x: x['createdDate'], reverse=True)

                    for grade in feedbacks_list[:5]:
                        feedback_rate = grade['productValuation']
                        rate_feedbacks.append(feedback_rate)

                    result.append({
                        'art': art,
                        'rating_art': feedback.get('rating_art'),
                        'feedbacks_count': feedback.get('feedbacks_count'),
                        'rate_feedbacks': rate_feedbacks
                    })
                except Exception as ex:
                    print(f'Ошибка с артикулом {art}: {ex}')
                    result.append({
                        'art': art,
                    })
        else:
            art = feedback['art']
            result.append({
                'art': art,
            })

    return result


async def excel_result_wb(feedback_and_rate_result):
    # Создаем папку data_wb_excel, если ее нет
    data_wb_excel_path = "data_wb_excel"
    if not os.path.exists(data_wb_excel_path):
        os.makedirs(data_wb_excel_path)

    # Для названия файла
    file_prefix = "wb"

    # Передаем данные для формирования excel файла
    data_wb = format_excel_data(feedback_and_rate_result, data_wb_excel_path, file_prefix)

    # Отправляем файл пользователю
    return data_wb
