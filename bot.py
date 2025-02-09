import logging
import openpyxl
from openpyxl.styles import Alignment, Border, Side, Font
from openpyxl.utils import get_column_letter
from telegram import Update
from telegram.request import HTTPXRequest
from telegram.ext import ApplicationBuilder, CommandHandler, ContextTypes
import datetime 
from calendar import monthrange
import os
from dotenv import load_dotenv

# Файл для хранения данных
FILENAME = 'attendance.xlsx'

# Настройка логирования
logging.basicConfig(format='%(asctime)s - %(name)s - %(levelname)s - %(message)s', level=logging.INFO)


# Функция для начала взаимодействия
async def start(update: Update, context: ContextTypes.DEFAULT_TYPE) -> None:
    await update.message.reply_text(
        "Добро пожаловать! Используйте команды:\n"
        "✅ /checkin - Отметить приход\n"
        "❌ /checkout - Отметить уход"
    )

# Ensure the spreadsheet is properly structured
def setup_attendance_sheet():
    now = datetime.datetime.now()
    sheet_name = f"{now.strftime('%B')} {now.year}"
    
    # Load or create workbook
    if os.path.exists(FILENAME):
        workbook = openpyxl.load_workbook(FILENAME)
    else:
        workbook = openpyxl.Workbook()

    # If sheet for current month doesn't exist, create it
    if sheet_name not in workbook.sheetnames:
        workbook.create_sheet(sheet_name)
        sheet = workbook[sheet_name]
        
        # Get number of days in month
        days_in_month = monthrange(now.year, now.month)[1]

        # Styles
        bold_font = Font(bold=True)
        center_align = Alignment(horizontal="center", vertical="center")
        border_style = Border(left=Side(style="thin"), right=Side(style="thin"),
                              top=Side(style="thin"), bottom=Side(style="thin"))

        # Headers
        sheet["A1"] = "Username"
        sheet["A1"].font = bold_font
        sheet["A1"].alignment = center_align
        sheet.column_dimensions["A"].width = 15  # Username column width

        # Create columns for each day (merged headers)
        col = 2
        for day in range(1, days_in_month + 1):
            col_start = get_column_letter(col)
            col_end = get_column_letter(col + 3)
            sheet.merge_cells(f"{col_start}1:{col_end}1")
            sheet[f"{col_start}1"] = str(day)
            sheet[f"{col_start}1"].font = bold_font
            sheet[f"{col_start}1"].alignment = center_align

            # Sub-headers
            sub_headers = ["Check-in", "Check-out", "Duration", "Comment"]
            for i, header in enumerate(sub_headers):
                cell = sheet[f"{get_column_letter(col + i)}2"]
                cell.value = header
                cell.font = bold_font
                cell.alignment = center_align

            col += 4

        # Apply formatting
        for row in sheet.iter_rows():
            for cell in row:
                cell.alignment = center_align
                cell.border = border_style

        workbook.save(FILENAME)

# Save check-in or check-out data
async def record_attendance(username, action):
    setup_attendance_sheet()
    now = datetime.datetime.now()
    sheet_name = f"{now.strftime('%B')} {now.year}"
    workbook = openpyxl.load_workbook(FILENAME)
    sheet = workbook[sheet_name]

    # Find or create user row
    user_row = None
    for row in range(3, sheet.max_row + 1):
        if sheet[f"A{row}"].value == username:
            user_row = row
            break

    if not user_row:
        user_row = sheet.max_row + 1
        sheet[f"A{user_row}"] = username

    # Find the correct column
    day_col_start = 2 + (now.day - 1) * 4
    checkin_cell = sheet[f"{get_column_letter(day_col_start)}{user_row}"]
    checkout_cell = sheet[f"{get_column_letter(day_col_start + 1)}{user_row}"]
    duration_cell = sheet[f"{get_column_letter(day_col_start + 2)}{user_row}"]

    # Store timestamps
    time_now = now.strftime("%HH:%MM")
    if action == "checkin":
        checkin_cell.value = time_now
    elif action == "checkout":
        checkout_cell.value = time_now

        # Calculate duration
        if checkin_cell.value:
            fmt = "%HH:%MM"
            checkin_time = datetime.datetime.strptime(checkin_cell.value, fmt)
            checkout_time = datetime.datetime.strptime(time_now, fmt)
            duration = checkout_time - checkin_time
            duration_cell.value = f"{duration.seconds // 3600}:{(duration.seconds % 3600) // 60}"

    workbook.save(FILENAME)

# Handle /checkin command
async def checkin(update: Update, context: ContextTypes.DEFAULT_TYPE):
    user = update.effective_user
    await record_attendance(user.username, "checkin")
    await update.message.reply_text(f"✅ {user.username}, you have checked in!")

# Handle /checkout command
async def checkout(update: Update, context: ContextTypes.DEFAULT_TYPE):
    user = update.effective_user
    await record_attendance(user.username, "checkout")
    await update.message.reply_text(f"✅ {user.username}, you have checked out!")


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