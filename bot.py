import logging
import openpyxl
from openpyxl.styles import Alignment, Border, Side, Font, PatternFill
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
logging.basicConfig(format='%(asctime)s - %(name)s - %(levelname)s - %(message)s', level=logging.WARN)

# Headers
sub_headers = ["Check-in", "Check-out", "Duration"]

# Columns
first_day_col = 4

# Styles
bold_font = Font(bold=True)
center_align = Alignment(horizontal="center", vertical="center")
border_style = Border(left=Side(style="thin"), right=Side(style="thin"), top=Side(style="thin"), bottom=Side(style="thin"))
color0="E0E0E0"
color1="FFFFFF"

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
    sheet_name = now.strftime("%m-%Y")
    
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

        # Headers
        sheet["A1"] = "Username"
        sheet["B1"] = "Total (1-15)"
        sheet["C1"] = "Total (16-End)"
        for col in ["A", "B", "C"]:
            sheet[col + "1"].font = bold_font
            sheet[col + "1"].alignment = center_align
            sheet.column_dimensions[col].width = 15  # Adjust column widths

        # Create columns for each day (merged headers)
        col = first_day_col
        for day in range(1, days_in_month + 1):
            col_start = get_column_letter(col)
            col_end = get_column_letter(col + len(sub_headers) - 1)
            sheet.merge_cells(f"{col_start}1:{col_end}1")
            sheet[f"{col_start}1"] = str(day)
            sheet[f"{col_start}1"].font = bold_font
            sheet[f"{col_start}1"].alignment = center_align
            color = color0 if day % 2 == 0 else color1
            sheet[f"{col_start}1"].fill = PatternFill(start_color=color, end_color=color, fill_type="solid")


            # Sub-headers
            for i, header in enumerate(sub_headers):
                cell = sheet[f"{get_column_letter(col + i)}2"]
                cell.value = header
                cell.font = bold_font
                cell.alignment = center_align
                cell.fill = PatternFill(start_color=color, end_color=color, fill_type="solid")

            col += len(sub_headers)

        # Apply formatting
        for row in sheet.iter_rows():
            for cell in row:
                cell.alignment = center_align
                cell.border = border_style

        workbook.save(FILENAME)
        
def setup_new_user(sheet, user_row):
    now = datetime.datetime.now()
    days_in_month = monthrange(now.year, now.month)[1]

    for col in ["A", "B", "C"]:
        cell = sheet[f"{col}{user_row}"]
        cell.alignment = center_align
        cell.border = border_style

    col = first_day_col
    for day in range(1, days_in_month + 1):
        color = color0 if day % 2 == 0 else color1
        for i in range(len(sub_headers)):
            cell = sheet[f"{get_column_letter(col + i)}{user_row}"]
            cell.alignment = center_align
            cell.fill = PatternFill(start_color=color, end_color=color, fill_type="solid")
            cell.border = border_style
        col += len(sub_headers)

    update_total_formulas(sheet, user_row, now.month)

# Save check-in or check-out data
async def record_attendance(username, action):
    setup_attendance_sheet()
    now = datetime.datetime.now()
    sheet_name = now.strftime("%m-%Y")
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
        setup_new_user(sheet, user_row)

    # Find the correct column
    day_col_start = first_day_col + (now.day - 1) * len(sub_headers)
    checkin_cell = sheet[f"{get_column_letter(day_col_start)}{user_row}"]
    checkout_cell = sheet[f"{get_column_letter(day_col_start + 1)}{user_row}"]
    duration_cell = sheet[f"{get_column_letter(day_col_start + 2)}{user_row}"]

    # Store timestamps
    if action == "checkin":
        checkin_cell.value = now
        checkin_cell.number_format = "hh:mm"

    elif action == "checkout":
        checkout_cell.value = now
        checkout_cell.number_format = "hh:mm"

        # Calculate duration
        if checkin_cell.value:
            checkin_time = checkin_cell.value
            checkout_time = checkout_cell.value

            duration = checkout_time - checkin_time
            duration_days = duration.total_seconds() / 86400  # Store as fraction of a day
            
            duration_cell.value = duration_days  
            duration_cell.number_format = "[h]:mm"

    workbook.save(FILENAME)

# Update total formulas for each user row
def update_total_formulas(sheet, user_row, month):
    days_in_month = monthrange(datetime.datetime.now().year, month)[1]

    duration_cols_1_15 = []
    duration_cols_16_end = []

    for day in range(1, days_in_month + 1):
        duration_col = get_column_letter(6 + (day - 1) * len(sub_headers))
        if day <= 15:
            duration_cols_1_15.append(f"{duration_col}{user_row}")
        else:
            duration_cols_16_end.append(f"{duration_col}{user_row}")

    sheet[f"B{user_row}"] = f"=ROUNDUP(SUM({','.join(duration_cols_1_15)}), 3)"
    sheet[f"B{user_row}"].number_format = "[h]:mm"

    sheet[f"C{user_row}"] = f"=ROUNDUP(SUM({','.join(duration_cols_16_end)}), 3)"
    sheet[f"C{user_row}"].number_format = "[h]:mm"


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