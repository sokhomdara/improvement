"""
GPPC QAQC Bot - Text Input Version (Most Reliable)
Uses simple step-by-step text questions - no button conflicts
"""
import os, logging, threading
from http.server import HTTPServer, BaseHTTPRequestHandler
from datetime import datetime
from telegram import Update, InlineKeyboardButton, InlineKeyboardMarkup, ReplyKeyboardMarkup, ReplyKeyboardRemove
from telegram.ext import (
    ApplicationBuilder, CommandHandler, MessageHandler,
    CallbackQueryHandler, ConversationHandler, ContextTypes, filters
)
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side

BOT_TOKEN = os.environ.get("BOT_TOKEN", "8660783157:AAEBEb_O-z514MdZspps0rp-IH4yg3cz4NM")
EXCEL_FILE = "GPPC_QAQC_Reports.xlsx"
ADMIN_IDS  = [352178789]
logging.basicConfig(level=logging.WARNING)

# ─── STATES ───
(ZONE, BLOCK, UNIT, FLOOR, HTYPE, WORKTYPE, ACTION,
 VENDOR, SUPERVISOR, ENGINEER, RAISED, PHOTO, COMMENT, CONFIRM,
 UPD_SELECT, UPD_STATUS, UPD_COMMENT, UPD_PHOTO) = range(18)

# ─── DATA ───
FLOORS    = ["GF","1F","2F","3F","RF","EX","Exterior","Underground"]
HTYPES    = ["KG","QN","QNA","QNBII","TWSE","LASII","LBIII","SHC"]
WORKTYPES = ["Finishing","Structure","MEP","Infrastructure"]
VENDORS   = ["ឡាច ពៅ","វ៉ាន់ សាគីន","ថន ផល្លា","Pholla","NIPPON","Dulux"]
SUPS      = ["C9","C16","C32","C45","C63","E6","E20","E22"]
ENGS      = [f"C{i}" for i in range(1,76)] + [f"E{i}" for i in range(1,26)]
RAISED_BY = ["Q2","Q3","Q6","Q7","Q10","Q11","Q12","Q13","Q16","Q17","Q21","Q26"]
DEFECTS   = {
    "Finishing":    ["ឥដ្ឋប៉ោង","ឥដ្ឋមិនពេញរ៉ង","ជញ្ជាំងបូកប៉ោង","ជ្រុងមិនត្រង់","សាច់បូកខ្លុក","ការ៉ូខ្លុក","ការ៉ូបែក","ម៉ាបបែក","ថ្នាំមិនស្អាត"],
    "Structure":    ["បេតុងសសុះចេញដែក","បេតុងប៉ោង","តំទ្បើងដែកខុសប្លង់"],
    "MEP":          ["ទុយោលិច","ឆ្លងភ្លើង"],
    "Infrastructure":["ផ្លូវខូច","លូទឹកស្ទះ","បណ្តាញទឹកលិច","ទ្រនាប់ខូច","ជញ្ជាំងព័រទ្រ"],
}

def make_reply_kb(options, columns=4):
    """Make reply keyboard from list"""
    rows = []
    row = []
    for i, o in enumerate(options, 1):
        row.append(o)
        if i % columns == 0:
            rows.append(row); row = []
    if row: rows.append(row)
    return ReplyKeyboardMarkup(rows, resize_keyboard=True, one_time_keyboard=True)

def remove_kb():
    return ReplyKeyboardRemove()

# ─── EXCEL ───
H_FILL = PatternFill("solid", start_color="1F3864")
Y_FILL = PatternFill("solid", start_color="FFF2CC")
G_FILL = PatternFill("solid", start_color="E2EFDA")
BD = Border(left=Side(style="thin"), right=Side(style="thin"),
            top=Side(style="thin"), bottom=Side(style="thin"))

def init_excel():
    if os.path.exists(EXCEL_FILE): return
    wb = openpyxl.Workbook(); ws = wb.active; ws.title = "QAQC"
    ws["A1"] = "QUALITY MONITORING AND IMPROVEMENT"
    ws["A1"].font = Font(bold=True, size=13, color="1F3864")
    ws["A2"] = "PROJECT TITLE: GRAND PHNOM PENH CITY"
    ws["A2"].font = Font(bold=True, size=11)
    ws.merge_cells("A1:S1"); ws.merge_cells("A2:S2")
    hdrs = ["No","Date","Picture Before","Picture After","Zone","Block","Unit","Floor",
            "House Type","Action Required","Raised By","Supervisor","Engineer",
            "Work Type","Vendor","Responsible","Status","Comments","Remark"]
    for c, h in enumerate(hdrs, 1):
        cell = ws.cell(row=3, column=c, value=h)
        cell.fill = H_FILL
        cell.font = Font(bold=True, color="FFFFFF", size=10)
        cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        cell.border = BD
    for i, w in enumerate([5,12,15,15,8,8,10,8,12,40,12,12,12,14,18,12,10,12,20], 1):
        ws.column_dimensions[openpyxl.utils.get_column_letter(i)].width = w
    wb.save(EXCEL_FILE)

def save_report(d):
    init_excel()
    wb = openpyxl.load_workbook(EXCEL_FILE); ws = wb.active
    row = ws.max_row + 1
    no  = max(0, ws.max_row - 3) + 1
    fill = Y_FILL
    vals = [no, d.get("date",""), "", "", d.get("zone",""), d.get("block",""),
            d.get("unit",""), d.get("floor",""), d.get("htype",""), d.get("action",""),
            d.get("raised",""), d.get("supervisor",""), d.get("engineer",""),
            d.get("worktype",""), d.get("vendor",""), "Site", "Open", "1", d.get("comment","")]
    for c, v in enumerate(vals, 1):
        cell = ws.cell(row=row, column=c, value=v)
        cell.fill = fill; cell.font = Font(size=10)
        cell.alignment = Alignment(vertical="center", wrap_text=True)
        cell.border = BD
    # Embed actual photo thumbnail in cell C
    local_path = d.get("photo_local_path", "")
    if local_path and os.path.exists(local_path):
        try:
            from openpyxl.drawing.image import Image as XLImage
            from PIL import Image as PILImage
            img = PILImage.open(local_path)
            img.thumbnail((150, 120))
            thumb = local_path + "_th.jpg"
            img.convert("RGB").save(thumb, "JPEG", quality=85)
            xl = XLImage(thumb)
            xl.width = 150; xl.height = 120
            xl.anchor = f"C{row}"
            ws.add_image(xl)
            ws.row_dimensions[row].height = 95
            ws.column_dimensions["C"].width = 22
        except:
            # Fallback to link
            cell = ws.cell(row=row, column=3, value="📷 BEFORE")
            if d.get("photo_url"):
                cell.hyperlink = d["photo_url"]
                cell.font = Font(size=10, color="0563C1", underline="single")
            cell.fill = fill; cell.border = BD
            cell.alignment = Alignment(horizontal="center", vertical="center")
    elif d.get("photo_url"):
        cell = ws.cell(row=row, column=3, value="📷 BEFORE")
        cell.hyperlink = d["photo_url"]
        cell.font = Font(size=10, color="0563C1", underline="single")
        cell.fill = fill; cell.border = BD
        cell.alignment = Alignment(horizontal="center", vertical="center")
        ws.row_dimensions[row].height = 18
    else:
        ws.row_dimensions[row].height = 18
    wb.save(EXCEL_FILE)
    return no

def update_excel(no, status, remark, photo_url="", photo_local=""):
    if not os.path.exists(EXCEL_FILE): return
    wb = openpyxl.load_workbook(EXCEL_FILE); ws = wb.active
    for row in ws.iter_rows(min_row=4):
        if row[0].value == no:
            row_idx = row[0].row
            row[16].value = status; row[18].value = remark
            if photo_local and os.path.exists(photo_local):
                try:
                    from openpyxl.drawing.image import Image as XLImage
                    from PIL import Image as PILImage
                    img = PILImage.open(photo_local)
                    img.thumbnail((150, 120))
                    thumb = photo_local + "_th.jpg"
                    img.convert("RGB").save(thumb, "JPEG", quality=85)
                    xl = XLImage(thumb)
                    xl.width = 150; xl.height = 120
                    xl.anchor = f"D{row_idx}"
                    ws.add_image(xl)
                    ws.row_dimensions[row_idx].height = 95
                    ws.column_dimensions["D"].width = 22
                except:
                    if photo_url:
                        c = ws.cell(row=row_idx, column=4, value="✅ AFTER")
                        c.hyperlink = photo_url
                        c.font = Font(size=10, color="0563C1", underline="single")
                        c.alignment = Alignment(horizontal="center", vertical="center")
            elif photo_url:
                c = ws.cell(row=row_idx, column=4, value="✅ AFTER")
                c.hyperlink = photo_url
                c.font = Font(size=10, color="0563C1", underline="single")
                c.alignment = Alignment(horizontal="center", vertical="center")
            fill = G_FILL if status == "Closed" else Y_FILL
            for cell in row: cell.fill = fill
            break
    wb.save(EXCEL_FILE)

def get_reports():
    if not os.path.exists(EXCEL_FILE): return []
    wb = openpyxl.load_workbook(EXCEL_FILE, data_only=True); ws = wb.active
    return [{"no":r[0],"date":r[1],"zone":r[4],"block":r[5],"unit":r[6],
             "action":r[9],"status":r[16]}
            for r in ws.iter_rows(min_row=4, values_only=True) if r[0]]

SE = {"Open":"🟡","In Progress":"🔵","Closed":"🟢"}
def sem(s): return SE.get(s, "⚪")
def pb(n, t=13): return "●"*n + "○"*(t-n) + f" {n}/{t}"

def summary(d, no=None):
    n = f"#️⃣ Report No: *{no}*\n" if no else ""
    return (f"📋 *QAQC DEFECT REPORT*\n{'─'*28}\n{n}"
            f"📅 {d.get('date','')}\n"
            f"📍 Zone {d.get('zone','')} / Block {d.get('block','')} / Unit {d.get('unit','')}\n"
            f"🏠 {d.get('floor','')} — {d.get('htype','')}\n"
            f"⚠️ *{d.get('action','')}* ({d.get('worktype','')})\n"
            f"🔧 Vendor: {d.get('vendor','')}\n"
            f"👷 Sup: {d.get('supervisor','')} | Eng: {d.get('engineer','')}\n"
            f"🔎 Raised: {d.get('raised','')}\n"
            f"📌 🟡 Open\n"
            f"💬 {d.get('comment','—')}")

# ─── /START /PING /RESTART ───
async def start(update, ctx):
    await update.message.reply_text(
        "👋 *Improvement Update Bot is ready!*\n\n"
        "• /report — QAQC log new defect\n"
        "• /update — Site team update fix status\n"
        "• /list — View all open defects\n"
        "• /export — Download Excel report\n"
        "• /ping — Check bot is alive",
        parse_mode="Markdown",
        reply_markup=remove_kb()
    )

async def ping(update, ctx):
    rs = get_reports()
    o = sum(1 for r in rs if r["status"] == "Open")
    c = sum(1 for r in rs if r["status"] == "Closed")
    await update.message.reply_text(
        f"🟢 *Bot is ALIVE!*\n"
        f"🕐 {datetime.now().strftime('%d %b %Y %H:%M:%S')}\n"
        f"📊 {len(rs)} total | 🟡 {o} open | 🟢 {c} closed\n"
        f"✅ All systems running normally!",
        parse_mode="Markdown"
    )

async def restart_cmd(update, ctx):
    if update.effective_user.id not in ADMIN_IDS:
        await update.message.reply_text("❌ Admin only."); return
    await update.message.reply_text("🔄 Restarting bot...")
    await ctx.application.stop()
    await ctx.application.shutdown()
    os._exit(0)

# ─── /REPORT FLOW (Reply Keyboard - most reliable) ───
async def report_start(update, ctx):
    ctx.user_data.clear()
    ctx.user_data["date"] = datetime.now().strftime("%Y-%m-%d")
    # Build zone keyboard 01-50
    zones = [f"{i:02d}" for i in range(1, 51)]
    await update.message.reply_text(
        f"📋 *New Defect Report*\n{pb(1)}\n\n🗺️ Select or type *Zone* (01–50):",
        parse_mode="Markdown",
        reply_markup=make_reply_kb(zones, columns=8)
    )
    return ZONE

async def got_zone(update, ctx):
    ctx.user_data["zone"] = update.message.text.strip()
    blocks = [f"{i:02d}" for i in range(1, 16)]
    await update.message.reply_text(
        f"✅ Zone: *{ctx.user_data['zone']}*\n\n{pb(2)}\n\n🏗️ Select or type *Block* (01–15):",
        parse_mode="Markdown",
        reply_markup=make_reply_kb(blocks, columns=5)
    )
    return BLOCK

async def got_block(update, ctx):
    ctx.user_data["block"] = update.message.text.strip()
    units = [f"{i:02d}" for i in range(1, 100)]
    await update.message.reply_text(
        f"✅ Block: *{ctx.user_data['block']}*\n\n{pb(3)}\n\n🏠 Select or type *Unit/House number*:",
        parse_mode="Markdown",
        reply_markup=make_reply_kb(units, columns=10)
    )
    return UNIT

async def got_unit(update, ctx):
    ctx.user_data["unit"] = update.message.text.strip()
    await update.message.reply_text(
        f"✅ Unit: *{ctx.user_data['unit']}*\n\n{pb(4)}\n\n🏢 Select or type *Floor*:",
        parse_mode="Markdown",
        reply_markup=make_reply_kb(FLOORS, columns=4)
    )
    return FLOOR

async def got_floor(update, ctx):
    ctx.user_data["floor"] = update.message.text.strip()
    await update.message.reply_text(
        f"✅ Floor: *{ctx.user_data['floor']}*\n\n{pb(5)}\n\n🏡 Select or type *House Type*:",
        parse_mode="Markdown",
        reply_markup=make_reply_kb(HTYPES, columns=4)
    )
    return HTYPE

async def got_htype(update, ctx):
    ctx.user_data["htype"] = update.message.text.strip()
    await update.message.reply_text(
        f"✅ House Type: *{ctx.user_data['htype']}*\n\n{pb(6)}\n\n🔧 Select or type *Type of Work*:",
        parse_mode="Markdown",
        reply_markup=make_reply_kb(WORKTYPES, columns=2)
    )
    return WORKTYPE

async def got_worktype(update, ctx):
    wt = update.message.text.strip()
    ctx.user_data["worktype"] = wt
    defects = DEFECTS.get(wt, DEFECTS["Finishing"])
    await update.message.reply_text(
        f"✅ Work Type: *{wt}*\n\n{pb(7)}\n\n⚠️ Select or type *Defect description*:",
        parse_mode="Markdown",
        reply_markup=make_reply_kb(defects, columns=2)
    )
    return ACTION

async def got_action(update, ctx):
    ctx.user_data["action"] = update.message.text.strip()
    await update.message.reply_text(
        f"✅ Defect noted\n\n{pb(8)}\n\n🏢 Select or type *Vendor/SubCon*:",
        parse_mode="Markdown",
        reply_markup=make_reply_kb(VENDORS, columns=2)
    )
    return VENDOR

async def got_vendor(update, ctx):
    ctx.user_data["vendor"] = update.message.text.strip()
    await update.message.reply_text(
        f"✅ Vendor: *{ctx.user_data['vendor']}*\n\n{pb(9)}\n\n👷 Select or type *Supervisor*:",
        parse_mode="Markdown",
        reply_markup=make_reply_kb(SUPS, columns=4)
    )
    return SUPERVISOR

async def got_supervisor(update, ctx):
    ctx.user_data["supervisor"] = update.message.text.strip()
    await update.message.reply_text(
        f"✅ Supervisor: *{ctx.user_data['supervisor']}*\n\n{pb(10)}\n\n🔧 Select or type *Site Engineer*:",
        parse_mode="Markdown",
        reply_markup=make_reply_kb(ENGS, columns=6)
    )
    return ENGINEER

async def got_engineer(update, ctx):
    ctx.user_data["engineer"] = update.message.text.strip()
    await update.message.reply_text(
        f"✅ Engineer: *{ctx.user_data['engineer']}*\n\n{pb(11)}\n\n🔎 Select or type *Raised By* (QAQC code):",
        parse_mode="Markdown",
        reply_markup=make_reply_kb(RAISED_BY, columns=4)
    )
    return RAISED

async def got_raised(update, ctx):
    ctx.user_data["raised"] = update.message.text.strip()
    await update.message.reply_text(
        f"✅ Raised by: *{ctx.user_data['raised']}*\n\n{pb(12)}\n\n📷 Send *BEFORE photo* of the defect\n_(or type - to skip)_",
        parse_mode="Markdown",
        reply_markup=remove_kb()
    )
    return PHOTO

async def got_photo(update, ctx):
    if update.message.photo:
        fid = update.message.photo[-1].file_id
        ctx.user_data["photo_fid"] = fid
        try:
            f = await update.get_bot().get_file(fid)
            ctx.user_data["photo_url"] = f.file_path
            # Download locally for Excel embedding
            os.makedirs("photos", exist_ok=True)
            local_path = f"photos/before_{datetime.now().strftime('%Y%m%d_%H%M%S')}.jpg"
            await f.download_to_drive(local_path)
            ctx.user_data["photo_local_path"] = local_path
        except:
            ctx.user_data["photo_url"] = ""
            ctx.user_data["photo_local_path"] = ""
        await update.message.reply_text(
            f"✅ Photo received!\n\n{pb(13)}\n\n💬 Add *comment* for site team:\n_(or type - to skip)_",
            parse_mode="Markdown",
            reply_markup=remove_kb()
        )
    else:
        # Text - skip photo
        ctx.user_data["photo_fid"] = ""
        ctx.user_data["photo_url"] = ""
        await update.message.reply_text(
            f"⏭️ Photo skipped\n\n{pb(13)}\n\n💬 Add *comment* for site team:\n_(or type - to skip)_",
            parse_mode="Markdown",
            reply_markup=remove_kb()
        )
    return COMMENT

async def got_comment(update, ctx):
    v = update.message.text.strip()
    ctx.user_data["comment"] = "" if v == "-" else v
    # Show summary with YES/NO keyboard
    await update.message.reply_text(
        summary(ctx.user_data) + "\n\n─────────────────────────\nSave this report?",
        parse_mode="Markdown",
        reply_markup=make_reply_kb(["✅ YES - Save", "❌ NO - Cancel"], columns=2)
    )
    return CONFIRM

async def got_confirm(update, ctx):
    v = update.message.text.strip()
    chat_id = update.effective_chat.id
    if "YES" in v.upper() or "✅" in v:
        no = save_report(ctx.user_data)
        await update.message.reply_text(
            f"✅ *Report #{no} saved successfully!*\n\n" +
            summary(ctx.user_data, no) +
            "\n\n📊 Use /export to download Excel",
            parse_mode="Markdown",
            reply_markup=remove_kb()
        )
        # Send BEFORE photo
        fid = ctx.user_data.get("photo_fid", "")
        if fid:
            try:
                await ctx.bot.send_photo(
                    chat_id=chat_id, photo=fid,
                    caption=f"📷 *BEFORE photo — Report #{no}*",
                    parse_mode="Markdown"
                )
            except: pass
        # Alert message
        sup = ctx.user_data.get("supervisor", "—")
        eng = ctx.user_data.get("engineer", "—")
        raised = ctx.user_data.get("raised", "—")
        await ctx.bot.send_message(
            chat_id=chat_id,
            text=(f"🚨 *DEFECT ALERT — Report #{no}*\n{'─'*28}\n"
                  f"📍 Zone {ctx.user_data.get('zone','')} / Block {ctx.user_data.get('block','')} / Unit {ctx.user_data.get('unit','')}\n"
                  f"🏠 {ctx.user_data.get('floor','')} — {ctx.user_data.get('htype','')}\n"
                  f"⚠️ *{ctx.user_data.get('action','')}* ({ctx.user_data.get('worktype','')})\n"
                  f"{'─'*28}\n"
                  f"👷 Supervisor: *{sup}*\n"
                  f"🔧 Engineer: *{eng}*\n"
                  f"🔎 Raised by: *{raised}*\n"
                  f"{'─'*28}\n"
                  f"📌 🟡 Open — please fix and use /update to close!"),
            parse_mode="Markdown"
        )
    else:
        await update.message.reply_text(
            "❌ Report cancelled. Type /report to start again.",
            reply_markup=remove_kb()
        )
    ctx.user_data.clear()
    return ConversationHandler.END

# ─── /UPDATE FLOW ───
async def update_start(update, ctx):
    ctx.user_data.clear()
    rs = [r for r in get_reports() if r["status"] != "Closed"]
    if not rs:
        await update.message.reply_text("✅ No open defects to update.", reply_markup=remove_kb())
        return ConversationHandler.END
    lines = "\n".join([
        f"{sem(r['status'])} *#{r['no']}* Z{r['zone']}/B{r['block']} — {str(r['action'])[:30]}..."
        for r in rs[-15:]
    ])
    nums = [f"#{r['no']}" for r in rs[-15:]]
    await update.message.reply_text(
        f"📋 *Open Defects:*\n{lines}\n\nType the report number to update (e.g. 1):",
        parse_mode="Markdown",
        reply_markup=make_reply_kb(nums, columns=5)
    )
    return UPD_SELECT

async def upd_select(update, ctx):
    v = update.message.text.strip().lstrip("#")
    try:
        ctx.user_data["upd_no"] = int(v)
    except:
        await update.message.reply_text("❌ Please enter a valid number.")
        return UPD_SELECT
    await update.message.reply_text(
        f"Report *#{ctx.user_data['upd_no']}* — Select new *Status*:",
        parse_mode="Markdown",
        reply_markup=make_reply_kb(["🟡 Open","🔵 In Progress","🟢 Closed"], columns=3)
    )
    return UPD_STATUS

async def upd_status(update, ctx):
    v = update.message.text.strip()
    if "Open" in v: ctx.user_data["upd_status"] = "Open"
    elif "Progress" in v: ctx.user_data["upd_status"] = "In Progress"
    elif "Closed" in v: ctx.user_data["upd_status"] = "Closed"
    else: ctx.user_data["upd_status"] = v
    await update.message.reply_text(
        f"✅ Status: *{ctx.user_data['upd_status']}*\n\n💬 Describe what was done to fix it:\n_(or type - to skip)_",
        parse_mode="Markdown",
        reply_markup=remove_kb()
    )
    return UPD_COMMENT

async def upd_comment(update, ctx):
    v = update.message.text.strip()
    ctx.user_data["upd_comment"] = "" if v == "-" else v
    await update.message.reply_text(
        "📷 Send *AFTER photo* (completion proof)\n_(or type - to skip)_",
        parse_mode="Markdown",
        reply_markup=remove_kb()
    )
    return UPD_PHOTO

async def upd_photo(update, ctx):
    no = ctx.user_data["upd_no"]
    status = ctx.user_data["upd_status"]
    comment = ctx.user_data["upd_comment"]
    photo_url = ""
    photo_local = ""
    fid = ""
    if update.message.photo:
        fid = update.message.photo[-1].file_id
        try:
            f = await update.get_bot().get_file(fid)
            photo_url = f.file_path
            os.makedirs("photos", exist_ok=True)
            photo_local = f"photos/after_{no}_{datetime.now().strftime('%Y%m%d_%H%M%S')}.jpg"
            await f.download_to_drive(photo_local)
        except: pass
    update_excel(no, status, comment, photo_url, photo_local)
    await update.message.reply_text(
        f"✅ *Report #{no} updated!*\n"
        f"Status → {sem(status)} *{status}*\n"
        f"💬 {comment or '—'}",
        parse_mode="Markdown",
        reply_markup=remove_kb()
    )
    if fid:
        try:
            await ctx.bot.send_photo(
                chat_id=update.effective_chat.id, photo=fid,
                caption=f"✅ *AFTER photo — Report #{no}*\nStatus: {sem(status)} {status}",
                parse_mode="Markdown"
            )
        except: pass
    ctx.user_data.clear()
    return ConversationHandler.END

# ─── /LIST /EXPORT ───
async def list_cmd(update, ctx):
    rs = get_reports()
    if not rs:
        await update.message.reply_text("📭 No reports yet. Use /report to log first defect.")
        return
    o = sum(1 for r in rs if r["status"] == "Open")
    p = sum(1 for r in rs if r["status"] == "In Progress")
    c = sum(1 for r in rs if r["status"] == "Closed")
    lines = "\n".join([
        f"{sem(r['status'])} *#{r['no']}* {r['date']} | Z{r['zone']}/B{r['block']} | {str(r['action'])[:25]}..."
        for r in rs[-20:]
    ])
    await update.message.reply_text(
        f"📊 *QAQC Report Summary*\n"
        f"Total: {len(rs)} | 🟡 Open: {o} | 🔵 Progress: {p} | 🟢 Closed: {c}\n"
        f"{'─'*28}\n{lines}",
        parse_mode="Markdown"
    )

async def export_cmd(update, ctx):
    if not os.path.exists(EXCEL_FILE):
        await update.message.reply_text("📭 No reports yet.")
        return
    await update.message.reply_document(
        document=open(EXCEL_FILE, "rb"),
        filename=f"GPPC_QAQC_{datetime.now().strftime('%Y%m%d')}.xlsx",
        caption=f"📊 Quality Monitoring Report\n{datetime.now().strftime('%d %b %Y %H:%M')}"
    )

async def cancel(update, ctx):
    ctx.user_data.clear()
    await update.message.reply_text(
        "❌ Cancelled. Type /report to start again.",
        reply_markup=remove_kb()
    )
    return ConversationHandler.END

# ─── KEEP ALIVE ───
class KA(BaseHTTPRequestHandler):
    def do_GET(self):
        self.send_response(200); self.send_header("Content-Type","text/plain"); self.end_headers()
        self.wfile.write(b"OK")
    def do_HEAD(self):
        self.send_response(200); self.end_headers()
    def log_message(self, *a): pass

def keep_alive():
    port = int(os.environ.get("PORT", 8080))
    t = threading.Thread(target=lambda: HTTPServer(("0.0.0.0", port), KA).serve_forever())
    t.daemon = True; t.start()
    print(f"🌐 Keep-alive port {port}")

# ─── MAIN ───
def main():
    keep_alive()
    init_excel()
    app = (ApplicationBuilder().token(BOT_TOKEN)
           .connect_timeout(30).read_timeout(30).write_timeout(30).build())

    TEXT = filters.TEXT & ~filters.COMMAND
    PHOTO_OR_TEXT = filters.PHOTO | (filters.TEXT & ~filters.COMMAND)

    report_conv = ConversationHandler(
        entry_points=[CommandHandler("report", report_start)],
        allow_reentry=True,
        states={
            ZONE:       [MessageHandler(TEXT, got_zone)],
            BLOCK:      [MessageHandler(TEXT, got_block)],
            UNIT:       [MessageHandler(TEXT, got_unit)],
            FLOOR:      [MessageHandler(TEXT, got_floor)],
            HTYPE:      [MessageHandler(TEXT, got_htype)],
            WORKTYPE:   [MessageHandler(TEXT, got_worktype)],
            ACTION:     [MessageHandler(TEXT, got_action)],
            VENDOR:     [MessageHandler(TEXT, got_vendor)],
            SUPERVISOR: [MessageHandler(TEXT, got_supervisor)],
            ENGINEER:   [MessageHandler(TEXT, got_engineer)],
            RAISED:     [MessageHandler(TEXT, got_raised)],
            PHOTO:      [MessageHandler(PHOTO_OR_TEXT, got_photo)],
            COMMENT:    [MessageHandler(TEXT, got_comment)],
            CONFIRM:    [MessageHandler(TEXT, got_confirm)],
        },
        fallbacks=[CommandHandler("cancel", cancel), CommandHandler("report", report_start)],
    )

    update_conv = ConversationHandler(
        entry_points=[CommandHandler("update", update_start)],
        allow_reentry=True,
        states={
            UPD_SELECT:  [MessageHandler(TEXT, upd_select)],
            UPD_STATUS:  [MessageHandler(TEXT, upd_status)],
            UPD_COMMENT: [MessageHandler(TEXT, upd_comment)],
            UPD_PHOTO:   [MessageHandler(PHOTO_OR_TEXT, upd_photo)],
        },
        fallbacks=[CommandHandler("cancel", cancel), CommandHandler("update", update_start)],
    )

    app.add_handler(CommandHandler("start",   start))
    app.add_handler(CommandHandler("help",    start))
    app.add_handler(CommandHandler("ping",    ping))
    app.add_handler(CommandHandler("restart", restart_cmd))
    app.add_handler(CommandHandler("list",    list_cmd))
    app.add_handler(CommandHandler("export",  export_cmd))
    app.add_handler(report_conv)
    app.add_handler(update_conv)

    print("🤖 QAQC Bot running - Text Input Mode")
    app.run_polling(poll_interval=0.5, timeout=20)

if __name__ == "__main__":
    main()
