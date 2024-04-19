import telebot
import openpyxl
from dotenv import dotenv_values

env_vars = dotenv_values(".env")

token = env_vars["TOKEN"]

bot = telebot.TeleBot(token)
file_path = env_vars["file_path"]


@bot.message_handler(commands=["start"])
def start_message(message):
    bot.send_message(
        message.chat.id,
        "Привет! Введите свой личный номер для получения расписания.",
    )


@bot.message_handler(func=lambda message: True)
def get_schedule(message):
    try:
        wb = openpyxl.load_workbook(file_path)
        sheet = wb.active
        number = message.text

        schedule_dict = {}

        fio = ""

        for row in sheet.iter_rows(min_row=2, values_only=True):
            if row[1] == number:
                fio = row[0]
                for i, day_schedule in enumerate(row[3:], 1):
                    date_key = str(i).zfill(2)
                    schedule_dict[date_key] = day_schedule

        if schedule_dict:
            resp = f"Расписание работы для сотрудника {fio}:\n"

            for date, schedule in schedule_dict.items():
                resp += f"{date}: {schedule}\n"

            bot.send_message(message.chat.id, resp)
        else:
            bot.send_message(
                message.chat.id, f"Расписание для сотрудника {number} не найдено."
            )

    except Exception as e:
        bot.send_message(message.chat.id, f"Произошла ошибка: {e}")


bot.polling()
