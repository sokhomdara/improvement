"""
GPPC QAQC Telegram Bot
======================
Project: Grand Phnom Penh City
Usage: Team sends /report command with photos in group chat.
       Bot collects data step-by-step and saves to Excel (QAQC template format).

Requirements:
    pip install python-telegram-bot openpyxl

Setup:
    1. Create bot via @BotFather → get BOT_TOKEN
    2. Add bot to your Telegram group as Admin
    3. Set BOT_TOKEN and EXCEL_FILE below
    4. Run: python gppc_qaqc_bot.py
"""

import os
import logging
import json
import threading
from http.server import HTTPServer, BaseHTTPRequestHandler
from datetime import datetime
from telegram import Update, InlineKeyboardButton, InlineKeyboardMarkup
from telegram.ext import (
    ApplicationBuilder, CommandHandler, MessageHandler,
    CallbackQueryHandler, ConversationHandler, ContextTypes, filters
)
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side

# ─────────────────────────────────────────────
# CONFIG — Edit these before running
# ─────────────────────────────────────────────
BOT_TOKEN = os.environ.get("BOT_TOKEN", "8660783157:AAF6Em-gZa0gEz8lynP8p5Z_9u7eCwJAZlc")
EXCEL_FILE = "GPPC_QAQC_Reports.xlsx"  # Auto-created if not exists
PHOTO_DIR  = "qaqc_photos"             # Folder to save photos

# ─────────────────────────────────────────────
# CONVERSATION STATES
# ─────────────────────────────────────────────
(
    ASK_ZONE, ASK_BLOCK, ASK_UNIT, ASK_FLOOR,
    ASK_HTYPE, ASK_ACTION, ASK_WORKTYPE, ASK_VENDOR,
    ASK_SUPERVISOR, ASK_ENGINEER, ASK_RAISED,
    ASK_PHOTO_BEFORE, ASK_PHOTO_AFTER, ASK_COMMENT,
    CONFIRM,
    # Update flow
    UPD_SELECT, UPD_STATUS, UPD_COMMENT, UPD_PHOTO
) = range(19)

logging.basicConfig(level=logging.INFO)
logger = logging.getLogger(__name__)

# ─────────────────────────────────────────────
# EXCEL SETUP
# ─────────────────────────────────────────────
HEADERS = [
    "No", "Date", "Picture Before", "Picture After",
    "Zone", "Block", "# House / Unit", "Floor", "House Type",
    "Action Required", "Defect Raised By", "Supervisor In Charge",
    "Site Engineer in Charge", "Type of Work", "Vendor/ SubCon",
    "Responsible by", "Status", "Type of Comments", "Remark"
]
HEADER_FILL  = PatternFill("solid", start_color="1F3864")
HEADER_FONT  = Font(name="Arial", bold=True, color="FFFFFF", size=10)
YELLOW_FILL  = PatternFill("solid", start_color="FFF2CC")
GREEN_FILL   = PatternFill("solid", start_color="E2EFDA")
RED_FILL     = PatternFill("solid", start_color="FCE4D6")
THIN_BORDER  = Border(
    left=Side(style="thin"), right=Side(style="thin"),
    top=Side(style="thin"), bottom=Side(style="thin")
)

def init_excel():
    """Create Excel file with GPPC template header if not exists."""
    if os.path.exists(EXCEL_FILE):
        return
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "QAQC Reports"
    ws["A1"] = "QUALITY MONITORING AND IMPROVEMENT"
    ws["A1"].font = Font(name="Arial", bold=True, size=13, color="1F3864")
    ws["A2"] = "PROJECT TITLE: GRAND PHNOM PENH CITY"
    ws["A2"].font = Font(name="Arial", bold=True, size=11)
    ws.merge_cells("A1:S1")
    ws.merge_cells("A2:S2")
    ws.row_dimensions[3].height = 20
    for col, h in enumerate(HEADERS, 1):
        cell = ws.cell(row=3, column=col, value=h)
        cell.fill = HEADER_FILL
        cell.font = HEADER_FONT
        cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        cell.border = THIN_BORDER
    # Column widths
    col_widths = [5,12,15,15,8,8,12,10,12,40,15,15,15,12,18,14,12,15,20]
    for i, w in enumerate(col_widths, 1):
        ws.column_dimensions[openpyxl.utils.get_column_letter(i)].width = w
    wb.save(EXCEL_FILE)

def get_next_row_and_no():
    wb = openpyxl.load_workbook(EXCEL_FILE)
    ws = wb.active
    row = ws.max_row + 1
    # Count existing data rows (skip header rows 1-3)
    data_rows = max(0, ws.max_row - 3)
    return row, data_rows + 1

def save_to_excel(data: dict):
    """Append one report row to Excel."""
    init_excel()
    wb = openpyxl.load_workbook(EXCEL_FILE)
    ws = wb.active
    next_row, report_no = get_next_row_and_no()

    status = data.get("status", "Open")
    fill = YELLOW_FILL if status == "Open" else GREEN_FILL if status == "Closed" else RED_FILL

    before_path = data.get("photo_before_path", "")
    after_path  = data.get("photo_after_path", "")
    row_data = [
        report_no,
        data.get("date", ""),
        "",
        "",
        data.get("zone", ""),
        data.get("block", ""),
        data.get("unit", ""),
        data.get("floor", ""),
        data.get("htype", ""),
        data.get("action", ""),
        data.get("raised", ""),
        data.get("supervisor", ""),
        data.get("engineer", ""),
        data.get("worktype", ""),
        data.get("vendor", ""),
        "Site",
        status,
        data.get("comment_type", "1"),
        data.get("comment", ""),
    ]
    for col, val in enumerate(row_data, 1):
        cell = ws.cell(row=next_row, column=col, value=val)
        cell.fill = fill
        cell.font = Font(name="Arial", size=10)
        cell.alignment = Alignment(vertical="center", wrap_text=True)
        cell.border = THIN_BORDER
    row_height = 18
    for ph, col_idx in [(before_path, 3), (after_path, 4)]:
        if ph and os.path.exists(ph):
            try:
                img = openpyxl.drawing.image.Image(ph)
                img.width, img.height = 80, 60
                img.anchor = openpyxl.utils.get_column_letter(col_idx) + str(next_row)
                ws.add_image(img)
                row_height = 50
            except Exception:
                ws.cell(row=next_row, column=col_idx, value=ph)
    ws.row_dimensions[next_row].height = row_height
    wb.save(EXCEL_FILE)
    return report_no

def update_excel_status(report_no: int, new_status: str, remark: str, photo_after: str):
    """Update status and remark for an existing report."""
    wb = openpyxl.load_workbook(EXCEL_FILE)
    ws = wb.active
    for row in ws.iter_rows(min_row=4):
        if row[0].value == report_no:
            row[16].value = new_status   # Status col
            row[18].value = remark        # Remark col
            if photo_after:
                row[3].value = photo_after  # Picture After col
            fill = GREEN_FILL if new_status == "Closed" else YELLOW_FILL
            for cell in row:
                cell.fill = fill
            break
    wb.save(EXCEL_FILE)

def get_all_reports():
    """Return list of (no, zone, block, action, status) for /list."""
    if not os.path.exists(EXCEL_FILE):
        return []
    wb = openpyxl.load_workbook(EXCEL_FILE, data_only=True)
    ws = wb.active
    rows = []
    for row in ws.iter_rows(min_row=4, values_only=True):
        if row[0]:
            rows.append({
                "no": row[0], "date": row[1],
                "zone": row[4], "block": row[5], "unit": row[6],
                "action": row[9], "status": row[16]
            })
    return rows

# ─────────────────────────────────────────────
# HELPERS
# ─────────────────────────────────────────────
def status_emoji(s):
    return {"Open": "🟡", "In Progress": "🔵", "Closed": "🟢"}.get(s, "⚪")

def summary_text(data: dict, report_no=None) -> str:
    no_str = f"#️⃣ Report No: *{report_no}*\n" if report_no else ""
    return (
        f"📋 *GPPC QAQC DEFECT REPORT*\n"
        f"{'─'*30}\n"
        f"{no_str}"
        f"📅 Date: {data.get('date','')}\n"
        f"📍 Zone {data.get('zone','')} / Block {data.get('block','')} {data.get('unit','')}\n"
        f"🏠 Floor: {data.get('floor','')} | Type: {data.get('htype','')}\n"
        f"⚠️ *Defect:* {data.get('action','')}\n"
        f"🔧 Work Type: {data.get('worktype','')}\n"
        f"🏗️ Vendor: {data.get('vendor','')}\n"
        f"👷 Supervisor: {data.get('supervisor','')} | Engineer: {data.get('engineer','')}\n"
        f"🔎 Raised By: {data.get('raised','')}\n"
        f"📌 Status: {status_emoji(data.get('status','Open'))} {data.get('status','Open')}\n"
        f"💬 Comment: {data.get('comment','—')}\n"
        f"{'─'*30}"
    )

async def save_photo(bot, file_id: str, label: str) -> str:
    """Download Telegram photo and save locally. Returns file path."""
    os.makedirs(PHOTO_DIR, exist_ok=True)
    ts = datetime.now().strftime("%Y%m%d_%H%M%S")
    path = os.path.join(PHOTO_DIR, f"{label}_{ts}.jpg")
    photo_file = await bot.get_file(file_id)
    await photo_file.download_to_drive(path)
    return path

# ─────────────────────────────────────────────
# /START & /HELP
# ─────────────────────────────────────────────
async def start(update: Update, ctx: ContextTypes.DEFAULT_TYPE):
    await update.message.reply_text(
        "👋 *GPPC QAQC Bot* is ready!\n\n"
        "Commands:\n"
        "• /report — Log a new defect\n"
        "• /update — Update defect status\n"
        "• /list — View all open defects\n"
        "• /export — Get Excel report file\n"
        "• /cancel — Cancel current action",
        parse_mode="Markdown"
    )

async def help_cmd(update: Update, ctx: ContextTypes.DEFAULT_TYPE):
    await start(update, ctx)

# ─────────────────────────────────────────────
# /REPORT CONVERSATION
# ─────────────────────────────────────────────
async def report_start(update: Update, ctx: ContextTypes.DEFAULT_TYPE):
    ctx.user_data.clear()
    ctx.user_data["date"] = datetime.now().strftime("%Y-%m-%d")
    ctx.user_data["status"] = "Open"
    await update.message.reply_text(
        "📋 *New Defect Report*\nStep 1/13\n\n"
        "Enter *Zone* number:\n_(e.g. 16, 17, 19)_",
        parse_mode="Markdown"
    )
    return ASK_ZONE

async def ask_block(update: Update, ctx: ContextTypes.DEFAULT_TYPE):
    ctx.user_data["zone"] = update.message.text.strip()
    await update.message.reply_text("Step 2/13\n\nEnter *Block* number:\n_(e.g. 1, 2, 8)_", parse_mode="Markdown")
    return ASK_BLOCK

async def ask_unit(update: Update, ctx: ContextTypes.DEFAULT_TYPE):
    ctx.user_data["block"] = update.message.text.strip()
    await update.message.reply_text("Step 3/13\n\nEnter *# House / Unit*:\n_(e.g. #38, #26 — or skip with -)_", parse_mode="Markdown")
    return ASK_UNIT

async def ask_floor(update: Update, ctx: ContextTypes.DEFAULT_TYPE):
    val = update.message.text.strip()
    ctx.user_data["unit"] = "" if val == "-" else val
    kb = [[
        InlineKeyboardButton("GF", callback_data="GF"),
        InlineKeyboardButton("1F", callback_data="1F"),
        InlineKeyboardButton("2F", callback_data="2F"),
        InlineKeyboardButton("3F", callback_data="3F"),
    ],[
        InlineKeyboardButton("RF", callback_data="RF"),
        InlineKeyboardButton("EX", callback_data="EX"),
        InlineKeyboardButton("Exterior", callback_data="Exterior"),
        InlineKeyboardButton("Underground", callback_data="Underground"),
    ]]
    await update.message.reply_text("Step 4/13\n\n🏠 Select *Floor*:", parse_mode="Markdown",
        reply_markup=InlineKeyboardMarkup(kb))
    return ASK_FLOOR

async def floor_chosen(update: Update, ctx: ContextTypes.DEFAULT_TYPE):
    query = update.callback_query
    await query.answer()
    ctx.user_data["floor"] = query.data
    await query.edit_message_text(f"Floor: *{query.data}* ✅\n\nStep 5/13\n\nEnter *House Type*:\n_(e.g. QNBII, TWSE, QNAG, QN)_", parse_mode="Markdown")
    return ASK_HTYPE

async def ask_action(update: Update, ctx: ContextTypes.DEFAULT_TYPE):
    ctx.user_data["htype"] = update.message.text.strip()
    await update.message.reply_text("Step 6/13\n\n⚠️ Describe the *Defect / Action Required*:\n_(Type in Khmer or English)_", parse_mode="Markdown")
    return ASK_ACTION

async def ask_worktype(update: Update, ctx: ContextTypes.DEFAULT_TYPE):
    ctx.user_data["action"] = update.message.text.strip()
    kb = [[
        InlineKeyboardButton("🪟 Finishing", callback_data="Finishing"),
        InlineKeyboardButton("🏗️ Structure", callback_data="Structure"),
        InlineKeyboardButton("⚡ MEP", callback_data="MEP"),
    ]]
    await update.message.reply_text("Step 7/13\n\n🔧 Select *Type of Work*:", parse_mode="Markdown",
        reply_markup=InlineKeyboardMarkup(kb))
    return ASK_WORKTYPE

async def worktype_chosen(update: Update, ctx: ContextTypes.DEFAULT_TYPE):
    query = update.callback_query
    await query.answer()
    ctx.user_data["worktype"] = query.data
    await query.edit_message_text(f"Work Type: *{query.data}* ✅\n\nStep 8/13\n\nEnter *Vendor / SubCon* name:\n_(e.g. ឡាច ពៅ, NIPPON, Dulux)_", parse_mode="Markdown")
    return ASK_VENDOR

async def ask_supervisor(update: Update, ctx: ContextTypes.DEFAULT_TYPE):
    ctx.user_data["vendor"] = update.message.text.strip()
    await update.message.reply_text("Step 9/13\n\nEnter *Supervisor In Charge* code:\n_(e.g. C103, C33, C39, E20)_", parse_mode="Markdown")
    return ASK_SUPERVISOR

async def ask_engineer(update: Update, ctx: ContextTypes.DEFAULT_TYPE):
    ctx.user_data["supervisor"] = update.message.text.strip()
    await update.message.reply_text("Step 10/13\n\nEnter *Site Engineer In Charge* code:\n_(e.g. C64, C118, M21, E83)_", parse_mode="Markdown")
    return ASK_ENGINEER

async def ask_raised(update: Update, ctx: ContextTypes.DEFAULT_TYPE):
    ctx.user_data["engineer"] = update.message.text.strip()
    await update.message.reply_text("Step 11/13\n\nEnter *Defect Raised By* (QAQC code):\n_(e.g. Q21, Q22, Q26)_", parse_mode="Markdown")
    return ASK_RAISED

async def ask_photo_before(update: Update, ctx: ContextTypes.DEFAULT_TYPE):
    ctx.user_data["raised"] = update.message.text.strip()
    await update.message.reply_text(
        "Step 12/13\n\n📷 Send *BEFORE photo* of the defect:\n_(Or type - to skip)_",
        parse_mode="Markdown"
    )
    return ASK_PHOTO_BEFORE

async def receive_photo_before(update: Update, ctx: ContextTypes.DEFAULT_TYPE):
    if update.message.photo:
        file_id = update.message.photo[-1].file_id
        path = await save_photo(update.get_bot(), file_id, "before")
        ctx.user_data["photo_before_path"] = path
        await update.message.reply_text("✅ BEFORE photo saved!")
    else:
        ctx.user_data["photo_before_path"] = ""
    await update.message.reply_text(
        "Step 13/13\n\n📷 Send *AFTER photo* (if repair done):\n_(Or type - to skip)_",
        parse_mode="Markdown"
    )
    return ASK_PHOTO_AFTER

async def receive_photo_after(update: Update, ctx: ContextTypes.DEFAULT_TYPE):
    if update.message.photo:
        file_id = update.message.photo[-1].file_id
        path = await save_photo(update.get_bot(), file_id, "after")
        ctx.user_data["photo_after_path"] = path
        await update.message.reply_text("✅ AFTER photo saved!")
    else:
        ctx.user_data["photo_after_path"] = ""
    await update.message.reply_text(
        "💬 Add a *comment* for the team:\n_(Or type - to skip)_",
        parse_mode="Markdown"
    )
    return ASK_COMMENT

async def ask_confirm(update: Update, ctx: ContextTypes.DEFAULT_TYPE):
    val = update.message.text.strip()
    ctx.user_data["comment"] = "" if val == "-" else val
    ctx.user_data["comment_type"] = "1"

    kb = [[
        InlineKeyboardButton("✅ Confirm & Save", callback_data="confirm"),
        InlineKeyboardButton("❌ Cancel", callback_data="cancel_report"),
    ]]
    await update.message.reply_text(
        summary_text(ctx.user_data) + "\n\nSave this report?",
        parse_mode="Markdown",
        reply_markup=InlineKeyboardMarkup(kb)
    )
    return CONFIRM

async def confirm_report(update: Update, ctx: ContextTypes.DEFAULT_TYPE):
    query = update.callback_query
    await query.answer()
    if query.data == "confirm":
        init_excel()
        report_no = save_to_excel(ctx.user_data)
        await query.edit_message_text(
            f"✅ *Report #{report_no} saved!*\n\n"
            + summary_text(ctx.user_data, report_no)
            + "\n\n📊 Use /export to download the Excel file.",
            parse_mode="Markdown"
        )
    else:
        await query.edit_message_text("❌ Report cancelled.")
    ctx.user_data.clear()
    return ConversationHandler.END

# ─────────────────────────────────────────────
# /UPDATE CONVERSATION
# ─────────────────────────────────────────────
async def update_start(update: Update, ctx: ContextTypes.DEFAULT_TYPE):
    reports = get_all_reports()
    open_reports = [r for r in reports if r["status"] != "Closed"]
    if not open_reports:
        await update.message.reply_text("✅ No open defects found.")
        return ConversationHandler.END
    lines = "\n".join([
        f"{status_emoji(r['status'])} *#{r['no']}* — Zone {r['zone']}/Blk {r['block']} — {str(r['action'])[:40]}..."
        for r in open_reports[-15:]
    ])
    await update.message.reply_text(
        f"📋 *Open Defects:*\n{lines}\n\nType the *Report Number* to update:",
        parse_mode="Markdown"
    )
    return UPD_SELECT

async def upd_select(update: Update, ctx: ContextTypes.DEFAULT_TYPE):
    try:
        ctx.user_data["upd_no"] = int(update.message.text.strip().lstrip("#"))
    except ValueError:
        await update.message.reply_text("❌ Invalid number. Try again or /cancel")
        return UPD_SELECT
    kb = [[
        InlineKeyboardButton("🟡 Open", callback_data="Open"),
        InlineKeyboardButton("🔵 In Progress", callback_data="In Progress"),
        InlineKeyboardButton("🟢 Closed", callback_data="Closed"),
    ]]
    await update.message.reply_text(
        f"Report #{ctx.user_data['upd_no']} — Select new *Status*:",
        parse_mode="Markdown", reply_markup=InlineKeyboardMarkup(kb)
    )
    return UPD_STATUS

async def upd_status_chosen(update: Update, ctx: ContextTypes.DEFAULT_TYPE):
    query = update.callback_query
    await query.answer()
    ctx.user_data["upd_status"] = query.data
    await query.edit_message_text(
        f"Status: *{query.data}* ✅\n\nDescribe what was done to fix the defect:\n_(Or type - to skip)_",
        parse_mode="Markdown"
    )
    return UPD_COMMENT

async def upd_comment(update: Update, ctx: ContextTypes.DEFAULT_TYPE):
    val = update.message.text.strip()
    ctx.user_data["upd_comment"] = "" if val == "-" else val
    await update.message.reply_text(
        "📷 Send *AFTER photo* (completion proof):\n_(Or type - to skip)_",
        parse_mode="Markdown"
    )
    return UPD_PHOTO

async def upd_photo(update: Update, ctx: ContextTypes.DEFAULT_TYPE):
    photo_path = ""
    if update.message.photo:
        file_id = update.message.photo[-1].file_id
        photo_path = await save_photo(update.get_bot(), file_id, f"after_upd{ctx.user_data['upd_no']}")

    update_excel_status(
        ctx.user_data["upd_no"],
        ctx.user_data["upd_status"],
        ctx.user_data["upd_comment"],
        photo_path
    )
    await update.message.reply_text(
        f"✅ *Report #{ctx.user_data['upd_no']} updated!*\n"
        f"Status → {status_emoji(ctx.user_data['upd_status'])} *{ctx.user_data['upd_status']}*\n"
        f"Comment: {ctx.user_data['upd_comment'] or '—'}",
        parse_mode="Markdown"
    )
    ctx.user_data.clear()
    return ConversationHandler.END

# ─────────────────────────────────────────────
# /LIST & /EXPORT
# ─────────────────────────────────────────────
async def list_reports(update: Update, ctx: ContextTypes.DEFAULT_TYPE):
    reports = get_all_reports()
    if not reports:
        await update.message.reply_text("📭 No reports yet. Use /report to log the first defect.")
        return
    lines = []
    for r in reports[-20:]:
        lines.append(
            f"{status_emoji(r['status'])} *#{r['no']}* {r['date']} | "
            f"Z{r['zone']}/B{r['block']} {r['unit']} | {str(r['action'])[:35]}..."
        )
    open_c  = sum(1 for r in reports if r["status"] == "Open")
    prog_c  = sum(1 for r in reports if r["status"] == "In Progress")
    close_c = sum(1 for r in reports if r["status"] == "Closed")
    await update.message.reply_text(
        f"📊 *GPPC QAQC Report Summary*\n"
        f"Total: {len(reports)} | 🟡 Open: {open_c} | 🔵 Progress: {prog_c} | 🟢 Closed: {close_c}\n"
        f"{'─'*30}\n" + "\n".join(lines),
        parse_mode="Markdown"
    )

async def export_excel(update: Update, ctx: ContextTypes.DEFAULT_TYPE):
    if not os.path.exists(EXCEL_FILE):
        await update.message.reply_text("📭 No reports yet.")
        return
    await update.message.reply_document(
        document=open(EXCEL_FILE, "rb"),
        filename=f"GPPC_QAQC_{datetime.now().strftime('%Y%m%d')}.xlsx",
        caption=f"📊 GPPC Quality Monitoring Report\n{datetime.now().strftime('%d %b %Y %H:%M')}"
    )

# ─────────────────────────────────────────────
# CANCEL
# ─────────────────────────────────────────────
async def cancel(update: Update, ctx: ContextTypes.DEFAULT_TYPE):
    ctx.user_data.clear()
    await update.message.reply_text("❌ Cancelled. Use /report to start again.")
    return ConversationHandler.END

# ─────────────────────────────────────────────
# MAIN
# ─────────────────────────────────────────────
# ─────────────────────────────────────────────
# KEEP-ALIVE SERVER (prevents Render from sleeping)
# ─────────────────────────────────────────────
class KeepAliveHandler(BaseHTTPRequestHandler):
    def do_GET(self):
        self.send_response(200)
        self.end_headers()
        self.wfile.write(b"GPPC QAQC Bot is alive!")
    def log_message(self, format, *args):
        pass  # suppress logs

def keep_alive():
    port = int(os.environ.get("PORT", 8080))
    server = HTTPServer(("0.0.0.0", port), KeepAliveHandler)
    thread = threading.Thread(target=server.serve_forever)
    thread.daemon = True
    thread.start()
    print(f"🌐 Keep-alive server running on port {port}")

def main():
    init_excel()
    os.makedirs(PHOTO_DIR, exist_ok=True)
    keep_alive()  # Start web server so Render stays awake

    app = ApplicationBuilder().token(BOT_TOKEN).build()

    # /report conversation
    report_conv = ConversationHandler(
        entry_points=[CommandHandler("report", report_start)],
        states={
            ASK_ZONE:         [MessageHandler(filters.TEXT & ~filters.COMMAND, ask_block)],
            ASK_BLOCK:        [MessageHandler(filters.TEXT & ~filters.COMMAND, ask_unit)],
            ASK_UNIT:         [MessageHandler(filters.TEXT & ~filters.COMMAND, ask_floor)],
            ASK_FLOOR:        [CallbackQueryHandler(floor_chosen)],
            ASK_HTYPE:        [MessageHandler(filters.TEXT & ~filters.COMMAND, ask_action)],
            ASK_ACTION:       [MessageHandler(filters.TEXT & ~filters.COMMAND, ask_worktype)],
            ASK_WORKTYPE:     [CallbackQueryHandler(worktype_chosen)],
            ASK_VENDOR:       [MessageHandler(filters.TEXT & ~filters.COMMAND, ask_supervisor)],
            ASK_SUPERVISOR:   [MessageHandler(filters.TEXT & ~filters.COMMAND, ask_engineer)],
            ASK_ENGINEER:     [MessageHandler(filters.TEXT & ~filters.COMMAND, ask_raised)],
            ASK_RAISED:       [MessageHandler(filters.TEXT & ~filters.COMMAND, ask_photo_before)],
            ASK_PHOTO_BEFORE: [
                MessageHandler(filters.PHOTO, receive_photo_before),
                MessageHandler(filters.TEXT & ~filters.COMMAND, receive_photo_before),
            ],
            ASK_PHOTO_AFTER:  [
                MessageHandler(filters.PHOTO, receive_photo_after),
                MessageHandler(filters.TEXT & ~filters.COMMAND, receive_photo_after),
            ],
            ASK_COMMENT:      [MessageHandler(filters.TEXT & ~filters.COMMAND, ask_confirm)],
            CONFIRM:          [CallbackQueryHandler(confirm_report)],
        },
        fallbacks=[CommandHandler("cancel", cancel)],
    )

    # /update conversation
    update_conv = ConversationHandler(
        entry_points=[CommandHandler("update", update_start)],
        states={
            UPD_SELECT:  [MessageHandler(filters.TEXT & ~filters.COMMAND, upd_select)],
            UPD_STATUS:  [CallbackQueryHandler(upd_status_chosen)],
            UPD_COMMENT: [MessageHandler(filters.TEXT & ~filters.COMMAND, upd_comment)],
            UPD_PHOTO:   [
                MessageHandler(filters.PHOTO, upd_photo),
                MessageHandler(filters.TEXT & ~filters.COMMAND, upd_photo),
            ],
        },
        fallbacks=[CommandHandler("cancel", cancel)],
    )

    app.add_handler(CommandHandler("start", start))
    app.add_handler(CommandHandler("help", help_cmd))
    app.add_handler(CommandHandler("list", list_reports))
    app.add_handler(CommandHandler("export", export_excel))
    app.add_handler(report_conv)
    app.add_handler(update_conv)

    print("🤖 GPPC QAQC Bot is running...")
    print(f"📁 Excel file: {EXCEL_FILE}")
    print(f"📷 Photos saved to: {PHOTO_DIR}/")
    app.run_polling()

if __name__ == "__main__":
    main()
