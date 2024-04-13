import telebot
import pandas as pd
import os
import datetime

# Получаем сегодняшнюю дату
current_date = datetime.date.today()

# Устанавливаем опции для отображения DataFrame
pd.set_option('display.max_colwidth', 50)  # максимальная ширина столбца
pd.set_option('display.max_rows', None)  # вывод всех строк DataFrame
pd.set_option('display.unicode.east_asian_width', True)  # учет ширины символов для адекватного отображения таблицы

# Токен бота
TOKEN = ' '

# Инициализация бота
bot = telebot.TeleBot(TOKEN)

# Загрузка данных из файла сводной таблицы
pivot = pd.read_excel("Pivot.xlsx")

# Преобразование столбца "Кол-во" в числовой формат и заполнение пропущенных значений нулями
pivot["Кол-во"] = pd.to_numeric(pivot["Кол-во"]).fillna(0).astype(int)


def form_formating(form):
    """Функция для форматирования данных из формы."""
    # Преобразование столбца "Кол-во" в числовой формат, замена пропущенных значений нулями и преобразование в целые числа
    form["Кол-во"] = pd.to_numeric(form["Кол-во"], errors='coerce').fillna(0).astype(int)
    # Выбор нужных столбцов и удаление первой строки, сброс индексов
    form = form[['Наименование', 'Кол-во']].iloc[1:109].reset_index(drop=True)
    return form


# Обработчик команды /start
@bot.message_handler(commands=['start'])
def send_welcome(message):
    """Обработчик команды /start."""
    with open("form.xlsx", 'rb') as form:
        # Отправляем приветственное сообщение
        bot.reply_to(message,
                     "Добрый день! Рады приветствовать вас в PontBakery. Для заказа заполните данную форму и отправьте в ответном сообщении. Название файла поменяйте на название вашей точки.")
        # Отправляем файл пользователю
        bot.send_document(message.chat.id, form)

# Обработчик для приема заполненных форм
@bot.message_handler(content_types=['document'])
def handle_document(message):
    """Обработчик для  заполненных форм."""

    # Скачивание файла
    file_info = bot.get_file(message.document.file_id)
    downloaded_file = bot.download_file(file_info.file_path)

    # Сохранение присланного файла
    with open(message.document.file_name, 'wb') as new_file:
        new_file.write(downloaded_file)

    # Чтение и форматирования Excel-файла клиента
    df = pd.read_excel(message.document.file_name, header=2)
    df = form_formating(df)

    # Обновление данных в сводной таблице
    pivot["Кол-во"] = pivot["Кол-во"] + df["Кол-во"]
    pivot.to_excel("pivots/"+str(current_date) + ".xlsx", index=False)

    # Вывод содержимого заказа с ненулевым количеством
    print(df[df["Кол-во"] > 0].to_string(index=False))

    # Форматирование таблицы для отправки в чат
    formatted_table = "Кол-во   Наименование      \n"
    for index, row in df[df["Кол-во"] > 0].iterrows():
        name = row['Наименование']
        quantity = str(row['Кол-во'])
        number_of_digits = int(len(str(quantity)))
        formatted_table += " " * (6 - number_of_digits) + quantity + " " * (16 - number_of_digits) + name + "\n"

    response = "Ваш заказ: \n" + formatted_table + "\n Заказ принят"

    # Отправка ответа пользователю
    bot.reply_to(message, response)

    # Удаление файла после использования
    os.remove(message.document.file_name)


# Запуск бота
bot.polling()
