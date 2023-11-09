import os
from aiogram import Bot, Dispatcher, types
from aiogram.contrib.fsm_storage.memory import MemoryStorage
from aiogram.dispatcher import FSMContext
from parse_utils_wb import parse_api_wb, feedback_from_wb_api, excel_result_wb
from parse_utils_ozon import feedback_from_ozon_api, excel_result_ozon

storage = MemoryStorage()
bot = Bot(token=os.environ['TOKEN'])
dp = Dispatcher(bot, storage=storage)


# Ответ на /start от пользователя и создание кнопок
@dp.message_handler(commands=['start'])
async def start(message: types.Message):
    keyboard = types.ReplyKeyboardMarkup(resize_keyboard=True)
    btn_wb = types.KeyboardButton("WB")
    btn_ozon = types.KeyboardButton("Ozon")
    keyboard.add(btn_wb, btn_ozon)

    await message.answer("Привет! Я парсер рейтинга и отзывов артикулов на WB и Ozon. Нажми на нужную кнопку.",
                         reply_markup=keyboard)


# Если пользователь нажал на кнопку WB
@dp.message_handler(lambda message: message.text == "WB")
async def wb_handler(message: types.Message, state: FSMContext):
    await message.answer("Пришлите список артикулов. Каждый артикул должен быть с новой строки без каких-либо знаков. "
                         "Присылать что-то кроме номеров артикулов не нужно.")
    await state.set_state("waiting_for_wb_articles")


# Если пользователь нажал на кнопку Ozon
@dp.message_handler(lambda message: message.text == "Ozon")
async def ozon_handler(message: types.Message, state: FSMContext):
    await message.answer("Пришлите список артикулов. Каждый артикул должен быть с новой строки без каких-либо знаков. "
                         "Присылать что-то кроме номеров артикулов не нужно.")
    await state.set_state("waiting_for_ozon_articles")


# Обработка WB-артикулов, которые прислал пользователь и ответ ему с готовым excel файлом
@dp.message_handler(state="waiting_for_wb_articles")
async def handle_wb_article(message: types.Message, state: FSMContext):
    # Список разделяем по переносам и записываем получившийся массив в переменную
    art_list = message.text.split('\n')

    await message.answer("Список артикулов для WB получен. Обработка данных может занять некоторое время. "
                         "Пожалуйста, подождите. Нормальное время ожидания: 1 минута при 30 артикулах.")

    # Вызываем функцию сбора количества отзывов
    result_api = await parse_api_wb(art_list)

    # Передаем полученные данные в функцию для сбора оценок к товарам
    feedback_and_rate_result = await feedback_from_wb_api(result_api)

    # Передаем полученные данные в функцию формирования Excel файла
    final_result = await excel_result_wb(feedback_and_rate_result)

    # Отправляем Excel файл пользователю
    await message.answer_document(types.InputFile(final_result), caption="Таблица с рейтингом для WB:")
    await state.finish()


# Обработка Ozon-артикулов, которые прислал пользователь и ответ ему с готовым excel файлом
@dp.message_handler(state="waiting_for_ozon_articles")
async def handle_ozon_article(message: types.Message, state: FSMContext):
    # Список разделяем по переносам и записываем получившийся массив в переменную
    art_list = message.text.split('\n')

    await message.answer("Список артикулов для Ozon получен. Обработка данных может занять некоторое время. "
                         "Пожалуйста, подождите. Нормальное время ожидания: 10 минут при 30 артикулах.")

    # Передаем полученные данные в функцию для сбора оценок к товарам
    feedback_and_rate_result = await feedback_from_ozon_api(art_list)

    # Передаем полученные данные в функцию формирования Excel файла
    final_result = await excel_result_ozon(feedback_and_rate_result)

    # Отправляем Excel файл пользователю
    await message.answer_document(types.InputFile(final_result), caption="Таблица с рейтингом для Ozon:")
    await state.finish()


@dp.message_handler()
async def handle_other_messages(message: types.Message):
    await message.answer("Ответ на данное сообщение не предусмотрен. Нажмите /start или одну из кнопок, пожалуйста.")
