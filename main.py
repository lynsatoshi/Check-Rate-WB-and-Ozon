from aiogram import executor
from telegram_bot import dp
from background import keep_alive


def main():
    keep_alive()
    executor.start_polling(dp)


if __name__ == "__main__":
    main()
