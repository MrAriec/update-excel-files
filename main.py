import win32com.client as win32
import os

from time import sleep
from datetime import datetime


# Функция для отправки сообщения на почту
def send_message_to_me(sender, message):
    mail = outlook.CreateItem(0)
    mail.To = sender
    mail.Subject = "Выполнение запросов"
    mail.Body = "Test"
    mail.HTMLBody = message  # this field is optional

    mail.Send()


# Функция возвращающая корректное время в определённом формате
def print_with_time(text):
    return print(f"[{datetime.now().strftime('%d/%m/%Y, %H:%M')}] {text}")


# Функция для обновления файла (Синтаксис из VBA, за счёт библиотеки)
def update_excel_file(path_file, excel_pass):
    xlapp = win32.DispatchEx("Excel.Application")
    wb = xlapp.Workbooks.Open(path_file, 2, False, None, None, excel_pass)
    xlapp.Visible = False
    wb.RefreshAll()
    xlapp.CalculateUntilAsyncQueriesDone()

    wb.Save()
    xlapp.Quit()


# Функция полного обновления отчётности
def update_all(my_excel_files, excel_pass):
    for excel_file in my_excel_files:
        print_with_time(f"Обновляем файл {excel_file.split('\\')[-1]}")
        update_excel_file(excel_file, excel_pass)

    # отправляем сообщение
    print_with_time("Сообщение отправлено")
    send_message_to_me(mail, "Успешно обновили файлы")


# outlook = win32.Dispatch("outlook.application")

print_with_time(
    "Данный скрипт предназначен для автоматического обновления отчётов Excel"
)
print_with_time(
    "В папке с этим скриптом должен находится один текстовый файл с путями к excel файлам"
)

mail = input(
    f"[{datetime.now().strftime('%d/%m/%Y, %H:%M')}] Введите почту на которую будет отправленно сообщение: "
)

pass_excel = input(
    f"[{datetime.now().strftime('%d/%m/%Y, %H:%M')}] Введите пароль для файлов (Если нет, то пропустите): "
)

current_dir = os.path.dirname(os.path.abspath(__file__))

while True:
    txt_files = [f for f in os.listdir(current_dir) if f.endswith(".txt")]

    if not txt_files:
        print_with_time("Нет текстовых файлов")
    elif len(txt_files) == 1:
        print_with_time("Найден файл {txt_files[0]}")
        file_path = os.path.join(current_dir, txt_files[0])
        with open(file_path, "r", encoding="utf-8") as file:
            excel_files = [
                e.strip()
                for e in file.readlines()
                if os.path.exists(e.strip()) and e.strip().endswith(".xlsx")
            ]

        if not excel_files:
            print_with_time("В файле нет корректный путей.")
        else:
            print_with_time("Начинаю обновлять файлы")
            break
    else:
        print_with_time("Файлов два и более, оставьте один")

    sleep(15)

update_all(excel_files, pass_excel)
