import telebot
import pandas as pd
from telebot import types
import os

API_TOKEN = ',,,'
bot = telebot.TeleBot(API_TOKEN)

def calculate_percentages(file_path):
    df = pd.read_excel(file_path, header=[0, 1])
    if ('Месяц', 'Проверено') not in df.columns or ('Месяц', 'Выдано') not in df.columns:
        raise ValueError("Столбцы 'Проверено' и/или 'Выдано' не найдены в файле.")

    df[('Месяц', 'Проверено')] = pd.to_numeric(df[('Месяц', 'Проверено')], errors='coerce').fillna(0)
    df[('Месяц', 'Выдано')] = pd.to_numeric(df[('Месяц', 'Выдано')], errors='coerce').fillna(0)

    df['Процент проверенных ДЗ'] = (df[('Месяц', 'Проверено')] / df[('Месяц', 'Выдано')]) * 100
    df['Процент проверенных ДЗ'] = df['Процент проверенных ДЗ'].fillna(0)
    return df[['ФИО преподавателя', 'Процент проверенных ДЗ']]

def calculate_issued_percentages(file_path):
    df = pd.read_excel(file_path, header=[0, 1])
    if ('Месяц', 'Выдано') not in df.columns or ('Неделя', 'Выдано') not in df.columns:
        raise ValueError("Столбцы 'Выдано' не найдены в файле.")

    # Суммируем общее количество выданного ДЗ для каждого преподавателя за месяц и неделю
    df[('Месяц', 'Выдано')] = pd.to_numeric(df[('Месяц', 'Выдано')], errors='coerce').fillna(0)
    df[('Неделя', 'Выдано')] = pd.to_numeric(df[('Неделя', 'Выдано')], errors='coerce').fillna(0)

    total_issued_month = df[('Месяц', 'Выдано')].sum()
    total_issued_week = df[('Неделя', 'Выдано')].sum()

    df['Процент выданного ДЗ (Месяц)'] = (df[('Месяц', 'Выдано')] / total_issued_month) * 100
    df['Процент выданного ДЗ (Неделя)'] = (df[('Неделя', 'Выдано')] / total_issued_week) * 100

    df['Процент выданного ДЗ (Месяц)'] = df['Процент выданного ДЗ (Месяц)'].fillna(0)
    df['Процент выданного ДЗ (Неделя)'] = df['Процент выданного ДЗ (Неделя)'].fillna(0)

    return df[['ФИО преподавателя', 'Процент выданного ДЗ (Месяц)', 'Процент выданного ДЗ (Неделя)']]

@bot.message_handler(commands=['start'])
def handle_start(message):
    chat_id = message.chat.id
    bot.send_message(chat_id, "Привет! Отправьте мне Excel файл для анализа.")

@bot.message_handler(content_types=['document'])
def handle_document(message):
    chat_id = message.chat.id
    file_info = bot.get_file(message.document.file_id)
    downloaded_file = bot.download_file(file_info.file_path)

    # Создаем директорию, если она не существует
    download_dir = 'downloaded_files'
    if not os.path.exists(download_dir):
        os.makedirs(download_dir)

    file_path = os.path.join(download_dir, message.document.file_name)
    with open(file_path, 'wb') as new_file:
        new_file.write(downloaded_file)

    markup = types.ReplyKeyboardMarkup(one_time_keyboard=True)
    markup.add('Подсчитать проценты проверенных ДЗ', 'Подсчитать проценты выданного ДЗ')
    msg = bot.send_message(chat_id, "Что вы хотите сделать?", reply_markup=markup)
    bot.register_next_step_handler(msg, process_choice, file_path)

def process_choice(message, file_path):
    chat_id = message.chat.id
    choice = message.text

    if choice == 'Подсчитать проценты проверенных ДЗ':
        process_calculate_percentages(message, file_path)
    elif choice == 'Подсчитать проценты выданного ДЗ':
        process_issued_percentages(message, file_path)
    else:
        bot.send_message(chat_id, "Неверный выбор. Пожалуйста, попробуйте снова.")

def process_calculate_percentages(message, file_path):
    chat_id = message.chat.id
    try:
        result_df = calculate_percentages(file_path)
        result_df['Процент проверенных ДЗ'] = result_df['Процент проверенных ДЗ'].fillna(0)

        print(result_df)

        message_parts = []
        for index, row in result_df.iterrows():
            teacher_name = row['ФИО преподавателя']
            percentage = row['Процент проверенных ДЗ']
            message = f"ФИО преподавателя: {teacher_name}\nПроцент проверенных ДЗ: {percentage:.2f}%"
            message_parts.append(message)
            print(message)

        for part in message_parts:
            bot.send_message(chat_id, part)

    except Exception as e:
        bot.send_message(chat_id, f"Ошибка при подсчете процентов: {e}")

def process_issued_percentages(message, file_path):
    chat_id = message.chat.id
    try:
        result_df = calculate_issued_percentages(file_path)
        result_df['Процент выданного ДЗ (Месяц)'] = result_df['Процент выданного ДЗ (Месяц)'].fillna(0)
        result_df['Процент выданного ДЗ (Неделя)'] = result_df['Процент выданного ДЗ (Неделя)'].fillna(0)

        print(result_df)

        message_parts = []
        for index, row in result_df.iterrows():
            teacher_name = row['ФИО преподавателя']
            percentage_month = row['Процент выданного ДЗ (Месяц)']
            percentage_week = row['Процент выданного ДЗ (Неделя)']
            message = (f"ФИО преподавателя: {teacher_name}\n"
                       f"Процент выданного ДЗ (Месяц): {percentage_month:.2f}%\n"
                       f"Процент выданного ДЗ (Неделя): {percentage_week:.2f}%")
            message_parts.append(message)
            print(message)

        for part in message_parts:
            bot.send_message(chat_id, part)

    except Exception as e:
        bot.send_message(chat_id, f"Ошибка при подсчете процентов: {e}")

bot.polling()
