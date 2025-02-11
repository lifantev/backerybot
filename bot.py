import logging
from openpyxl.utils import get_column_letter as col_let
from telegram import Update
from telegram.request import HTTPXRequest
from telegram.ext import ApplicationBuilder, CommandHandler
from datetime import datetime
from calendar import monthrange
import os
from dotenv import load_dotenv
import gspread
from gspread import Spreadsheet, Worksheet
from gspread_formatting import (
    CellFormat,
    Color,
    Borders,
    Border,
    TextFormat,
    format_cell_range,
    set_column_width,
)
from oauth2client.service_account import ServiceAccountCredentials

logging.basicConfig(
    format="%(asctime)s - %(name)s - %(levelname)s - %(message)s", level=logging.WARN
)


# Headers
sub_headers = ["Check-in", "Check-out", "Длительность"]

# Columns
days_start_col = 4

# Styles
border_thin = Borders(
    top=Border(style="SOLID"),
    bottom=Border(style="SOLID"),
    left=Border(style="SOLID"),
    right=Border(style="SOLID"),
)
color_grey = Color(red=0.95, green=0.95, blue=0.95)
color_white = Color(red=1, green=1, blue=1)


def apply_styles(
    sheet: Worksheet,
    cell_range: str,
    bold=False,
    color=None,
):
    fmt = CellFormat(
        textFormat=TextFormat(bold=bold),
        horizontalAlignment="CENTER",
        borders=border_thin,
    )
    if color:
        fmt.backgroundColor = color

    format_cell_range(sheet, cell_range, fmt)


async def record_attendance(
    action: str, username: str | None, spreadsheet: Spreadsheet
):
    now = datetime.now()
    sheet = setup_attendance_sheet(spreadsheet, now)
    user_row = setup_user_row(sheet, username, now)

    now = datetime.now()
    col = days_start_col + (now.day - 1) * len(sub_headers)
    checkin_cell = (user_row, col)
    checkout_cell = (user_row, col + 1)
    duration_cell = (user_row, col + 2)

    if action == "checkin":
        sheet.update_cell(*checkin_cell, now.strftime("%H:%M"))
    elif action == "checkout":
        sheet.update_cell(*checkout_cell, now.strftime("%H:%M"))
        sheet.update_cell(
            *duration_cell,
            f"= {col_let(col+1)}{user_row} - {col_let(col)}{user_row}",
        )


def setup_attendance_sheet(spreadsheet: Spreadsheet, now: datetime) -> Worksheet:
    sheet_name = now.strftime("%m-%Y")

    try:
        sheet = spreadsheet.worksheet(sheet_name)
    except gspread.exceptions.WorksheetNotFound:
        sheet = spreadsheet.add_worksheet(title=sheet_name, rows="100", cols="100")

        days_in_month = monthrange(now.year, now.month)[1]

        row1 = ["Username", "Total (1-15)", "Total (16-End)"] + [
            val for day in range(1, days_in_month + 1) for val in [str(day), "", ""]
        ]
        row2 = ["", "", ""] + sub_headers * days_in_month
        # apply_styles(sheet, "A1:C2", bold=True)
        # set_column_width(sheet, "A:C", 15)

        col = days_start_col
        for _ in range(days_in_month):
            col_start = col_let(col)
            col_end = col_let(col + len(sub_headers) - 1)
            sheet.merge_cells(f"{col_start}1:{col_end}1")

            # cell = f"{col_start}1"
            # color = color_grey if day % 2 == 0 else color_white
            # apply_styles(sheet, cell, bold=True, color=color)
            # apply_styles(
            #     sheet,
            #     f"{col_let(col)}2:{col_let(col + len(sub_headers)- 1)}2",
            #     bold=True,
            #     color=color,
            # )
            col += len(sub_headers)

        sheet.update([row1, row2])

    return sheet


def setup_user_row(sheet: Worksheet, username: str | None, now: datetime) -> int:
    usernames = sheet.col_values(1)
    if username in usernames:
        return usernames.index(username) + 1

    user_row = 3 if len(usernames) == 1 else len(usernames) + 1
    sheet.update_cell(user_row, 1, username)

    # apply_styles(sheet, f"A{user_row}:C{user_row}", bold=True)

    days_in_month = monthrange(now.year, now.month)[1]

    # col = days_start_col
    # for day in range(1, days_in_month + 1):
    #     color = color_grey if day % 2 == 0 else color_white
    #     cells = (
    #         f"{col_let(col)}{user_row}:{col_let(col + len(sub_headers) - 1)}{user_row}"
    #     )
    #     apply_styles(sheet, cells, color=color)
    #     col += len(sub_headers)

    setup_total_formulas(sheet, user_row, days_in_month)

    return user_row


def setup_total_formulas(sheet: Worksheet, user_row: int, days_in_month: int):
    duration_cols_1, duration_cols_2 = [], []
    for day in range(1, days_in_month + 1):
        duration_col = col_let(6 + (day - 1) * len(sub_headers))
        if day <= 15:
            duration_cols_1.append(f"{duration_col}{user_row}")
        else:
            duration_cols_2.append(f"{duration_col}{user_row}")

    sheet.update_acell(f"B{user_row}", f"= {'+'.join(duration_cols_1)}")
    sheet.update_acell(f"C{user_row}", f"= {'+'.join(duration_cols_2)}")


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
        await record_attendance("checkin", user.username, self.spreadsheet)
        await update.message.reply_text(f"✅ {user.username}, you have checked in!")

    async def checkout(self, update: Update, context):
        user = update.effective_user
        await record_attendance("checkout", user.username, self.spreadsheet)
        await update.message.reply_text(f"❌ {user.username}, you have checked out!")


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
