import logging
import openpyxl
from telegram import Update, InlineKeyboardButton, InlineKeyboardMarkup
from telegram.request import HTTPXRequest
from telegram.ext import ApplicationBuilder, CommandHandler, CallbackQueryHandler, ContextTypes
import datetime 
from telegram import Bot
from telegram.ext import Application
import os
from dotenv import load_dotenv
import httpx

# Настройка логирования
logging.basicConfig(format='%(asctime)s - %(name)s - %(levelname)s - %(message)s', level=logging.INFO)

# async def start(update: Update, context: ContextTypes.DEFAULT_TYPE) -> None:
#     keyboard = [
#         [InlineKeyboardButton("Отметиться", callback_data='checkin')],
#     ]
#     reply_markup = InlineKeyboardMarkup(keyboard)
#     await update.message.reply_text('Добро пожаловать! Нажмите кнопку ниже, чтобы отметиться.', reply_markup=reply_markup)

# Функция для начала взаимодействия
async def start(update: Update, context: ContextTypes.DEFAULT_TYPE) -> None:
    await update.message.reply_text(
        "Добро пожаловать! Используйте команды:\n"
        "✅ /checkin - Отметить приход\n"
        "❌ /checkout - Отметить уход"
    )

# Функция для отметки прихода
async def checkin(update: Update, context: ContextTypes.DEFAULT_TYPE) -> None:
    user = update.message.from_user
    save_to_excel(user.username, "Check-in")
    await update.message.reply_text(f"✅ {user.username}, вы отметились на вход!")

# Функция для отметки ухода
async def checkout(update: Update, context: ContextTypes.DEFAULT_TYPE) -> None:
    user = update.message.from_user
    save_to_excel(user.username, "Check-out")
    await update.message.reply_text(f"❌ {user.username}, вы отметились на выход!")

# Функция для обработки нажатий кнопок
async def button(update: Update, context: ContextTypes.DEFAULT_TYPE) -> None:
    query = update.callback_query
    await query.answer()

    if query.data == 'checkin':
        user = query.from_user
        save_to_excel(user.username)
        await query.edit_message_text(text=f"Спасибо, {user.username}, вы отметились!")

# Функция для сохранения данных в Excel
def save_to_excel(username, action):
    file_path = 'attendance.xlsx'
    
    try:
        workbook = openpyxl.load_workbook(file_path)
    except FileNotFoundError:
        workbook = openpyxl.Workbook()
        sheet = workbook.active
        sheet.append(["Username", "Action", "Timestamp"])  # Заголовки для столбцов
    else:
        sheet = workbook.active

    current_time = datetime.datetime.now().strftime('%Y-%m-%d %H:%M:%S')
    sheet.append([username, action, current_time])
    workbook.save(file_path)

# Основная функция
def main():
    load_dotenv() 

    # Вставьте токен вашего бота
    request = HTTPXRequest(
        read_timeout=60.0,
        write_timeout=60.0,
        connect_timeout=60.0,
        pool_timeout=60.0,
    )
    application = Application.builder().token(os.getenv("TOKEN")).request(request).build()

    # Добавляем обработчики
    application.add_handler(CommandHandler("start", start))
    application.add_handler(CommandHandler("checkin", checkin))
    application.add_handler(CommandHandler("checkout", checkout))

    # Запуск бота
    print("Bot is running...")
    application.run_polling()

# Запуск программы
if __name__ == '__main__':
    main()