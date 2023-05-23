import smtplib
# библиотека Питона для работы с эл. почтой. SMTP - протокол передачи по email
from email.mime.text import MIMEText  # Нужно для того, чтобы принимало кириллицу
import gspread  # библиотека для работы с Google Sheets
import pandas as pd
from check_email import check  # функция для проверки правильности записи e-mail


def send_email(table):  # функция на вход принимает имя Гугл таблицы

    sa = gspread.service_account('service_account_key.json')  # загружаем ключ json из сервисного аккаунта Гугл
    sh = sa.open(table)  # открываем Гугл таблицу по имени
    wks = sh.worksheet('Лист1')  # Обращаемся к первому листу таблицы
    df = pd.DataFrame(wks.get_all_records())  # Читаем Гугл таблицу как пандас датафрейм

    with open('1.txt', 'r') as f:  # Читаем файл с логином и паролем от почты, с которой будем рассылать
        lines = [line.rstrip() for line in f]
    login = lines[0]
    password = lines[1]

    server = smtplib.SMTP_SSL('smtp.yandex.com', 465)  # Подключение к серверу

    error_count = 0  # Счетчик ошибок при рассылке
    try:
        server.login(login, password)
        server.auth_plain()

        for row in df.itertuples():  # проходим по каждой строке в Гугл таблице
            try:
                if not check(list(row)[1]):  # Проверяем, правильно ли записан e-mail
                    wks.update_cell(list(row)[0] + 2, 3, f'Ошибка. Неверный формат e-mail.')
                    error_count += 1
                    continue

                msg = MIMEText(list(row)[2])  # Создаем объект сообщение
                msg['Subject'] = 'Это рассылка!'  # Задаем тему письма
                msg['From'] = 'alex.devgru@yandex.ru'  # Задаем отправителя в заголовках
                msg['To'] = list(row)[1]  # Задаем получателя в заголовках
                server.sendmail(login, list(row)[1], msg.as_string())

            except Exception as ex:
                # в случае ошибки отправки пишем в Гугл таблицу сообщение об ошибке:
                wks.update_cell(list(row)[0] + 2, 3, f'Ошибка. Сообщение не отправлено.\n{ex}')
                error_count += 1
                continue

        return f'Сообщения разосланы!\n' \
               f'Из {df.shape[0]} адресов успешно отправлено {df.shape[0] - error_count} адресатам.\n' \
               f'{error_count} ошибок(ки) при отправке.'

    except Exception as ex:
        return f'{ex}.\n' \
               f'Check your login or password.'


'''
в итоге будет обрабатываться 3 ошибки:
1) если неверный логин или пароль от почты, то сообщение от ошибке выведется в консоль
2) если неверный формат e-mail, то сообщение об ошибке выведется в Гугл таблицу
3) если сообщение не отправлено (например, e-mail не существует), то сообщение об ошибке также
выведется в Гугл таблицу
'''


def main():
    table = 'e-mails'
    print(send_email(table=table))


if __name__ == "__main__":
    main()
