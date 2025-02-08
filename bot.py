import logging
import openpyxl
from telegram import Update
from telegram.request import HTTPXRequest
from telegram.ext import ApplicationBuilder, CommandHandler, ContextTypes
import datetime 
import os
from dotenv import load_dotenv

# Файл для хранения данных
FILE_PATH = 'attendance.xlsx'

# Настройка логирования
logging.basicConfig(format='%(asctime)s - %(name)s - %(levelname)s - %(message)s', level=logging.INFO)


# Функция для начала взаимодействия
async def start(update: Update, context: ContextTypes.DEFAULT_TYPE) -> None:
    await update.message.reply_text(
        "Добро пожаловать! Используйте команды:\n"
        "✅ /checkin - Отметить приход\n"
        "❌ /checkout - Отметить уход"
    )

# Функция для отметки прихода (Check-in)
async def checkin(update: Update, context: ContextTypes.DEFAULT_TYPE) -> None:
    user = update.message.from_user
    response = save_to_excel(user.username, "Check-in")
    await update.message.reply_text(response)

# Функция для отметки ухода (Check-out)
async def checkout(update: Update, context: ContextTypes.DEFAULT_TYPE) -> None:
    user = update.message.from_user
    response = save_to_excel(user.username, "Check-out")
    await update.message.reply_text(response)

# Функция для сохранения данных в Excel
def save_to_excel(username, action):
    """Сохраняет данные о входе/выходе в файл Excel с разделением по месяцам."""
    today = datetime.datetime.now()
    month_sheet = today.strftime('%Y-%m')  # Пример: "2025-02"
    date_str = today.strftime('%Y-%m-%d')
    time_str = today.strftime('%H:%M:%S')

    # Открываем или создаем файл
    if os.path.exists(FILE_PATH):
        workbook = openpyxl.load_workbook(FILE_PATH)
    else:
        workbook = openpyxl.Workbook()

    # Если лист для текущего месяца отсутствует, создаем его
    if month_sheet not in workbook.sheetnames:
        sheet = workbook.create_sheet(title=month_sheet)
        sheet.append(["Date", "Username", "Check-in Time", "Check-out Time", "Work Duration (hh:mm)"])  # Заголовки
    else:
        sheet = workbook[month_sheet]

    # Проверяем, есть ли уже запись о Check-in для этого пользователя в этот день
    found_row = None
    for row in sheet.iter_rows(min_row=2, values_only=True):  # Пропускаем заголовки
        if row[0] == date_str and row[1] == username:
            found_row = row
            break

    if action == "Check-in":
        if found_row:
            return f"⚠️ {username}, вы уже отметились сегодня в {found_row[2]}!"
        else:
            sheet.append([date_str, username, time_str, "", ""])
            workbook.save(FILE_PATH)
            return f"✅ {username}, ваш вход в {time_str} сохранен!"
    
    elif action == "Check-out":
        for row_idx, row in enumerate(sheet.iter_rows(min_row=2, values_only=True), start=2):
            if row[0] == date_str and row[1] == username and row[2]:  # Убеждаемся, что есть Check-in
                sheet.cell(row=row_idx, column=4, value=time_str)  # Записываем Check-out
                checkin_time = datetime.datetime.strptime(row[2], '%H:%M:%S')
                checkout_time = datetime.datetime.strptime(time_str, '%H:%M:%S')
                duration = checkout_time - checkin_time
                sheet.cell(row=row_idx, column=5, value=str(duration))  # Записываем Work Duration
                workbook.save(FILE_PATH)
                return f"❌ {username}, ваш выход в {time_str} сохранен! Рабочее время: {duration}"
        
        return f"⚠️ {username}, нет записи о входе сегодня! Пожалуйста, сначала отметьтесь через /checkin."


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
    application = ApplicationBuilder().token(os.getenv("TOKEN")).request(request).build()

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