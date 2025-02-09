import logging
import openpyxl
from openpyxl.styles import Alignment, PatternFill, Border, Side, Font
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
    """Сохраняет данные о входе/выходе в таблицу с пользователями в строках и днями в столбцах."""
    today = datetime.datetime.now()
    month_sheet = today.strftime("%Y-%m")  # Пример: "2025-02"
    day_col = today.strftime("%d")  # День месяца (01, 02, ..., 31)
    time_str = today.strftime("%H:%M")  # Формат ЧЧ:ММ

    # Открываем или создаем файл
    if os.path.exists(FILE_PATH):
        workbook = openpyxl.load_workbook(FILE_PATH)
    else:
        workbook = openpyxl.Workbook()

    # Если листа для текущего месяца нет, создаем его
    if month_sheet not in workbook.sheetnames:
        sheet = workbook.create_sheet(title=month_sheet)
        # Set headers based on the number of days in the current month
        days_in_month = (datetime.datetime(today.year, today.month + 1, 1) - datetime.timedelta(days=1)).day
        sheet.append(["Username"] + [f"{day:02}" for day in range(1, days_in_month + 1)])

    else:
        sheet = workbook[month_sheet]

    # Поиск строки пользователя (или добавление нового)
    user_row = None
    for row_idx, row in enumerate(sheet.iter_rows(min_row=2, values_only=True), start=2):
        if row[0] == username:
            user_row = row_idx
            break

    if user_row is None:
        user_row = sheet.max_row + 1
        sheet.append([username])  # Добавляем новую строку

    # Определение колонки для текущего дня
    day_col_idx = int(day_col)  # Например, '05' → 5-й столбец
    cell = sheet.cell(row=user_row, column=day_col_idx + 1)  # +1 из-за смещения Username

    if action == "Check-in":
        if cell.value:
            return f"⚠️ {username}, вы уже отметились сегодня!"
        cell.value = f"{time_str} - "
    elif action == "Check-out":
        if not cell.value or "-" not in cell.value:
            return f"⚠️ {username}, отметьтесь через /checkin сначала!"
        
        checkin_time = cell.value.split(" - ")[0]
        checkout_time = time_str
        duration = calculate_duration(checkin_time, checkout_time)
        cell.value = f"{checkin_time} - {checkout_time} ({duration})"

    # Форматируем таблицу
    format_cells(sheet)

    # Сохраняем файл
    workbook.save(FILE_PATH)
    
    return f"✅ {username}, ваш {action.lower()} в {time_str} сохранен!"

# Функция вычисления разницы между check-in и check-out
def calculate_duration(checkin, checkout):
    """Вычисляет длительность работы в формате HH:MM."""
    fmt = "%H:%M"
    checkin_time = datetime.datetime.strptime(checkin, fmt)
    checkout_time = datetime.datetime.strptime(checkout, fmt)
    duration = checkout_time - checkin_time
    hours, remainder = divmod(duration.seconds, 3600)
    minutes = remainder // 60
    return f"{hours:02}:{minutes:02}"

# Функция форматирования таблицы (центрирование, автоширина, стилизация)
def format_cells(sheet):
    """Центрирует значения, подгоняет ширину колонок и добавляет стили."""
    for col in sheet.columns:
        max_length = 0
        col_letter = col[0].column_letter  # Получаем букву колонки
        for cell in col:
            if cell.value:
                max_length = max(max_length, len(str(cell.value)))
            cell.alignment = Alignment(horizontal="center", vertical="center")  # Центрируем
            cell.border = Border(
                left=Side(style='thin'),
                right=Side(style='thin'),
                top=Side(style='thin'),
                bottom=Side(style='thin')
            )

        adjusted_width = max_length + 2  # Добавляем небольшой запас
        sheet.column_dimensions[col_letter].width = adjusted_width
    
    # Apply alternating colors for readability
    for row in sheet.iter_rows(min_row=2, max_row=sheet.max_row):
        for i, cell in enumerate(row):
            if i % 2 == 0:  # Alternate colors for columns
                cell.fill = PatternFill(start_color="E0E0E0", end_color="E0E0E0", fill_type="solid")
            else:
                cell.fill = PatternFill(start_color="FFFFFF", end_color="FFFFFF", fill_type="solid")
    
    # Apply bold and style headers
    for cell in sheet[1]:
        cell.font = Font(bold=True)


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