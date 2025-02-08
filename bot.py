import logging
import openpyxl
from openpyxl.worksheet.table import Table, TableStyleInfo
from openpyxl.styles import Alignment
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

# Функция для сохранения данных в xlsx
def save_to_excel(username, action):
    """Сохраняет данные о входе/выходе в файл xlsx с разделением по месяцам."""
    today = datetime.datetime.now()
    month_sheet = today.strftime('%m-%Y')
    date_str = today.strftime('%d-%m-%Y')
    time_str = today.strftime('%H:%M')

    # Открываем или создаем файл
    if os.path.exists(FILE_PATH):
        workbook = openpyxl.load_workbook(FILE_PATH)
    else:
        workbook = openpyxl.Workbook()

    # Если лист для текущего месяца отсутствует, создаем его
    if month_sheet not in workbook.sheetnames:
        sheet = workbook.create_sheet(title=month_sheet)
        sheet.append(["Дата", "Username", "Время Check-in", "Время Check-out", "Рабочее время"])  # Заголовки
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

            format_as_table(sheet)
            workbook.save(FILE_PATH)

            return f"✅ {username}, ваш вход в {time_str} сохранен!"
    
    elif action == "Check-out":
        for row_idx, row in enumerate(sheet.iter_rows(min_row=2, values_only=True), start=2):
            if row[0] == date_str and row[1] == username and row[2]:  # Убеждаемся, что есть Check-in
                sheet.cell(row=row_idx, column=4, value=time_str)  # Записываем Check-out
                checkin_time = datetime.datetime.strptime(row[2], '%H:%M')
                checkout_time = datetime.datetime.strptime(time_str, '%H:%M')

                duration = checkout_time - checkin_time
                hours, remainder = divmod(duration.seconds, 3600)
                minutes = remainder // 60
                formatted_duration = f"{hours:02}:{minutes:02}"  # Формат ЧЧ:ММ
                sheet.cell(row=row_idx, column=5, value=formatted_duration)  # Записываем Work Duration

                format_as_table(sheet)
                workbook.save(FILE_PATH)

                return f"❌ {username}, ваш выход в {time_str} сохранен! Рабочее время: {duration}"
        
        return f"⚠️ {username}, нет записи о входе сегодня! Пожалуйста, сначала отметьтесь через /checkin."


# Функция для форматирования листа как таблицы
def format_as_table(sheet):
    """Добавляет фильтрацию и делает данные таблицей в Excel."""
    if sheet.max_row < 2:
        return  # Нет данных, нечего форматировать

    table_ref = f"A1:E{sheet.max_row}"  # Определяем диапазон таблицы
    table = Table(displayName="AttendanceTable", ref=table_ref)

    # Стиль таблицы
    style = TableStyleInfo(
        name="TableStyleMedium9", showFirstColumn=False, 
        showLastColumn=False, showRowStripes=True, showColumnStripes=False
    )
    table.tableStyleInfo = style

    # Удаляем старую таблицу (если есть), чтобы избежать дублирования
    sheet._tables.clear()
    
    # Добавляем новую таблицу
    sheet.add_table(table)
    
    # Автоподгонка ширины колонок
    autofit_columns(sheet)
    
def autofit_columns(sheet):
    """Автоматически подгоняет ширину колонок по содержимому."""
    for col in sheet.columns:
        max_length = 0
        col_letter = col[0].column_letter  # Получаем букву колонки
        for cell in col:
            if cell.value:
                max_length = max(max_length, len(str(cell.value)))
                # Применяем выравнивание по центру
                cell.alignment = Alignment(horizontal="center", vertical="center")
        adjusted_width = max_length + 5  # Добавляем небольшой запас
        sheet.column_dimensions[col_letter].width = adjusted_width

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