import telebot
import pandas as pd
from telebot import types

API_TOKEN = '7680441789:AAEKZP8BdtkUTaWr3p_FBPld-xAkws_Xst4'
bot = telebot.TeleBot(API_TOKEN)

def send_excel_file(chat_id, file_path):
    with open(file_path, 'rb') as file:
        bot.send_document(chat_id, file)

def calculate_percentages(file_path):
    df = pd.read_excel(file_path, header=[0, 1])
    if ('Месяц', 'Проверено') not in df.columns or ('Месяц', 'Выдано') not in df.columns:
        raise ValueError("Столбцы 'Проверено' и/или 'Выдано' не найдены в файле.")

    df[('Месяц', 'Проверено')] = pd.to_numeric(df[('Месяц', 'Проверено')], errors='coerce').fillna(0)
    df[('Месяц', 'Выдано')] = pd.to_numeric(df[('Месяц', 'Выдано')], errors='coerce').fillna(0)

    df['Процент проверенных ДЗ'] = (df[('Месяц', 'Проверено')] / df[('Месяц', 'Выдано')]) * 100
    df['Процент проверенных ДЗ'] = df['Процент проверенных ДЗ'].fillna(0)
    return df[['ФИО преподавателя', 'Процент проверенных ДЗ']]

@bot.message_handler(commands=['send_excel'])
def handle_send_excel(message):
    chat_id = message.chat.id
    msg = bot.send_message(chat_id, "Пожалуйста, отправьте путь к Excel файлу:")
    bot.register_next_step_handler(msg, process_file_path)

def process_file_path(message):
    chat_id = message.chat.id
    file_path = message.text
    try:
        send_excel_file(chat_id, file_path)
        bot.send_message(chat_id, "Файл успешно отправлен!")
    except Exception as e:
        bot.send_message(chat_id, f"Ошибка при отправке файла: {e}")

@bot.message_handler(commands=['calculate_percentages'])
def handle_calculate_percentages(message):
    chat_id = message.chat.id
    msg = bot.send_message(chat_id, "Пожалуйста, отправьте путь к Excel файлу:")
    bot.register_next_step_handler(msg, process_calculate_percentages)

def process_calculate_percentages(message):
    chat_id = message.chat.id
    file_path = message.text
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

@bot.message_handler(commands=['show_percentages'])
def handle_show_percentages(message):
    chat_id = message.chat.id
    msg = bot.send_message(chat_id, "Пожалуйста, отправьте путь к Excel файлу:")
    bot.register_next_step_handler(msg, process_show_percentages)

def process_show_percentages(message):
    chat_id = message.chat.id
    file_path = message.text
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
        bot.send_message(chat_id, f"Ошибка при выводе процентов: {e}")

bot.polling()