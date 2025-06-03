import teradatasql
import time
import win32com.client as win32
import os
import schedule
import pandas as pd

from time import sleep
from datetime import datetime
from getpass import getpass


# Функция для отправки сообщения на почту
def send_message_to_me(sender, message):
    mail = outlook.CreateItem(0)
    mail.To = sender
    mail.Subject = "Выполнение запросов"
    mail.Body = "Test"
    mail.HTMLBody = message  # this field is optional

    mail.Send()


# Функция возвращающая корректное время в определённом формате
def get_current_time():
    return f"[{datetime.now().strftime("%d/%m/%Y, %H:%M")}] "


# Функция для обновления файла (Синтаксис из VBA, за счёт библиотеки)
def update_excel_file(path_file):
    xlapp = win32.DispatchEx("Excel.Application")
    wb = xlapp.Workbooks.Open(path_file)
    xlapp.Visible = False
    wb.RefreshAll()
    xlapp.CalculateUntilAsyncQueriesDone()

    wb.Save()
    xlapp.Quit()


# Функция полного обновления отчётности (Определение недели, внесение новой недели в БД и обновление отчётов)
def update_all():

    # обновляем файлы
    for excel_file in excel_files:
        print(f"{get_current_time()}Обновляем файл {excel_file.split('\\')[-1]}")
        update_excel_file(excel_file)

    # отправляем сообщение
    print(f"{get_current_time()}Сообщение отправлено")
    send_message_to_me(mail, "Успешно внесли неделю и обновили файлы")


# Строка, которая заменяется на актуальную неделю
replace_str = "REPORT_WEEK"
mail = "faryshev_da@magnit.ru"

# Ссылки на файлы для обновления
excel_files = [
    r"R:\Конкурентный анализ\Ценовые индексы\1. Фактический индекс Магнит к Конкурентам\1. Основной_ММ\Парсинг\Детали SKU\202521_ММ_Пятёрочка_Магазин_SKU.xlsx",
    r"R:\Конкурентный анализ\Ценовые индексы\1. Фактический индекс Магнит к Конкурентам\1. Основной_ММ\Парсинг\Детали SKU\202521_ММ_Остальные Конкуренты_Магазин_SKU.xlsx",
]
outlook = win32.Dispatch("outlook.application")


# Задаём отложенный запуск
schedule.every().wednesday.at("09:00").do(update_all)

print(
    f"{get_current_time()}Данный скрипт предназначен для автоматического обновления отчётов с индексами"
)


while True:
    update_now = input(f"{get_current_time()}Обновить сейчас файлы? (y/n): ")
    if update_now == "y":
        update_all()
        print(f"{get_current_time()}Ожидаем среду")
        break
    elif update_now == "n":
        print(f"{get_current_time()}Ожидаем среду")
        break
    else:
        print("Нужно вводить только y или n")

while True:
    schedule.run_pending()
    sleep(1)
