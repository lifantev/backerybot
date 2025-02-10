import logging
import openpyxl
from openpyxl.styles import Alignment, Border, Side, Font, PatternFill
from openpyxl.utils import get_column_letter
from telegram import Update
from telegram.request import HTTPXRequest
from telegram.ext import ApplicationBuilder, CommandHandler
import datetime
from calendar import monthrange
import os
from dotenv import load_dotenv
import gspread
from gspread import Spreadsheet
from oauth2client.service_account import ServiceAccountCredentials

logging.basicConfig(
    format="%(asctime)s - %(name)s - %(levelname)s - %(message)s", level=logging.WARN
)


# Headers
sub_headers = ["Check-in", "Check-out", "Duration"]

# Columns
days_start_col = 4

# Styles
font_bold = Font(bold=True)
align_center = Alignment(horizontal="center", vertical="center")
border_thin = Border(
    left=Side(style="thin"),
    right=Side(style="thin"),
    top=Side(style="thin"),
    bottom=Side(style="thin"),
)
color_grey = "E0E0E0"
color_white = "FFFFFF"


def apply_styles(
    cell, border=border_thin, alignment=align_center, font=None, color=None
):
    if font:
        cell.font = font
    if alignment:
        cell.alignment = alignment
    if border:
        cell.border = border
    if color:
        cell.fill = PatternFill(start_color=color, end_color=color, fill_type="solid")


def setup_attendance_sheet():
    now = datetime.datetime.now()
    sheet_name = now.strftime("%m-%Y")

    if os.path.exists(FILENAME):
        workbook = openpyxl.load_workbook(FILENAME)
    else:
        workbook = openpyxl.Workbook()

    if sheet_name not in workbook.sheetnames:
        workbook.create_sheet(sheet_name)
        sheet = workbook[sheet_name]

        days_in_month = monthrange(now.year, now.month)[1]

        sheet["A1"] = "Username"
        sheet["B1"] = "Total (1-15)"
        sheet["C1"] = "Total (16-End)"
        for col in ["A", "B", "C"]:
            apply_styles(sheet[f"{col}1"], font=font_bold)
            apply_styles(sheet[f"{col}2"], font=font_bold)
            sheet.column_dimensions[col].width = 15

        col = days_start_col
        for day in range(1, days_in_month + 1):
            col_start = get_column_letter(col)
            col_end = get_column_letter(col + len(sub_headers) - 1)
            sheet.merge_cells(f"{col_start}1:{col_end}1")

            cell = sheet[f"{col_start}1"]
            cell.value = str(day)
            color = color_grey if day % 2 == 0 else color_white
            apply_styles(cell, font=font_bold, color=color)

            for i, header in enumerate(sub_headers):
                cell = sheet[f"{get_column_letter(col + i)}2"]
                cell.value = header
                apply_styles(cell, font=font_bold, color=color)

            col += len(sub_headers)

        workbook.save(FILENAME)


def setup_new_user(sheet, user_row):
    now = datetime.datetime.now()
    days_in_month = monthrange(now.year, now.month)[1]

    for col in ["A", "B", "C"]:
        cell = sheet[f"{col}{user_row}"]
        apply_styles(cell, font=font_bold)

    col = days_start_col
    for day in range(1, days_in_month + 1):
        color = color_grey if day % 2 == 0 else color_white
        for i in range(len(sub_headers)):
            cell = sheet[f"{get_column_letter(col + i)}{user_row}"]
            apply_styles(cell, color=color)
        col += len(sub_headers)

    setup_total_formulas(sheet, user_row, now.month)


async def record_attendance(action, username, filename):
    setup_attendance_sheet()
    now = datetime.datetime.now()
    sheet_name = now.strftime("%m-%Y")
    workbook = openpyxl.load_workbook(filename)
    sheet = workbook[sheet_name]

    user_row = None
    for row in range(3, sheet.max_row + 1):
        if sheet[f"A{row}"].value == username:
            user_row = row
            break
    if not user_row:
        user_row = sheet.max_row + 1
        sheet[f"A{user_row}"] = username
        setup_new_user(sheet, user_row)

    col = days_start_col + (now.day - 1) * len(sub_headers)
    checkin_cell = sheet[f"{get_column_letter(col)}{user_row}"]
    checkout_cell = sheet[f"{get_column_letter(col + 1)}{user_row}"]
    duration_cell = sheet[f"{get_column_letter(col + 2)}{user_row}"]

    if action == "checkin":
        checkin_cell.value = now
        checkin_cell.number_format = "hh:mm"
    elif action == "checkout":
        checkout_cell.value = now
        checkout_cell.number_format = "hh:mm"
        duration_cell.value = f"={get_column_letter(col+1)}{user_row} - {get_column_letter(col)}{user_row}"
        duration_cell.number_format = "hh:mm"

    workbook.save(filename)


def setup_total_formulas(sheet, user_row, month):
    days_in_month = monthrange(datetime.datetime.now().year, month)[1]
    duration_cols_1 = []
    duration_cols_2 = []

    for day in range(1, days_in_month + 1):
        duration_col = get_column_letter(6 + (day - 1) * len(sub_headers))
        if day <= 15:
            duration_cols_1.append(f"{duration_col}{user_row}")
        else:
            duration_cols_2.append(f"{duration_col}{user_row}")

    sheet[f"B{user_row}"] = f"=ROUNDUP(SUM({','.join(duration_cols_1)}), 3)"
    sheet[f"B{user_row}"].number_format = "[h]:mm"

    sheet[f"C{user_row}"] = f"=ROUNDUP(SUM({','.join(duration_cols_2)}), 3)"
    sheet[f"C{user_row}"].number_format = "[h]:mm"


class ActionHandler:

    def __init__(self, spreadsheet: Spreadsheet):
        self.spreadsheet = spreadsheet

    async def start(self, update: Update, context):
        await update.message.reply_text(
            "Добро пожаловать! Используйте команды:\n"
            "✅ /checkin - Отметить приход\n"
            "❌ /checkout - Отметить уход"
        )

    async def checkin(self, update: Update, context):
        user = update.effective_user
        await record_attendance(user.username, "checkin", self.spreadsheet)
        await update.message.reply_text(f"✅ {user.username}, you have checked in!")

    async def checkout(self, update: Update, context):
        user = update.effective_user
        await record_attendance(user.username, "checkout", self.spreadsheet)
        await update.message.reply_text(f"✅ {user.username}, you have checked out!")


def main():
    load_dotenv()

    tg_app = (
        ApplicationBuilder()
        .token(os.getenv("TOKEN"))
        .request(
            HTTPXRequest(
                read_timeout=60.0,
                write_timeout=60.0,
                connect_timeout=60.0,
                pool_timeout=60.0,
            )
        )
        .build()
    )

    scope = [
        "https://spreadsheets.google.com/feeds",
        "https://www.googleapis.com/auth/drive",
    ]
    creds = ServiceAccountCredentials.from_json_keyfile_name(
        os.getenv("SERVICE_ACCOUNT_FILE"), scope
    )
    gspread_client = gspread.authorize(creds)

    spreadsheet = gspread_client.open_by_key(os.getenv("SPREADSHEET_ID"))

    bot = ActionHandler(spreadsheet)

    tg_app.add_handler(CommandHandler("start", bot.start))
    tg_app.add_handler(CommandHandler("checkin", bot.checkin))
    tg_app.add_handler(CommandHandler("checkout", bot.checkout))

    print("Bot is running...")
    tg_app.run_polling()


if __name__ == "__main__":
    main()
