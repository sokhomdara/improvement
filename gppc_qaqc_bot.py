"""
GPPC QAQC Telegram Bot - Fast & Clean Version
==============================================
Project: Grand Phnom Penh City
"""

import os
import io
import logging
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
# CONFIG
# ─────────────────────────────────────────────
BOT_TOKEN  = os.environ.get("BOT_TOKEN", "8660783157:AAEBEb_O-z514MdZspps0rp-IH4yg3cz4NM")
EXCEL_FILE = "GPPC_QAQC_Reports.xlsx"
ADMIN_IDS  = [352178789]

logging.basicConfig(level=logging.WARNING)
logger = logging.getLogger(__name__)

# ─────────────────────────────────────────────
# CONVERSATION STATES
# ─────────────────────────────────────────────
(
    S_ZONE, S_BLOCK, S_UNIT, S_FLOOR,
    S_HTYPE, S_WORKTYPE, S_ACTION,
    S_VENDOR, S_SUPERVISOR, S_ENGINEER, S_RAISED,
    S_PHOTO, S_COMMENT, S_CONFIRM,
    U_SELECT, U_STATUS, U_COMMENT, U_PHOTO
) = range(18)

# ─────────────────────────────────────────────
# PRESETS
# ─────────────────────────────────────────────
ZONES      = [f"{i:02d}" for i in range(1, 51)]
BLOCKS     = [f"{i:02d}" for i in range(1, 16)]
UNITS      = [f"{i:02d}" for i in range(1, 100)]
FLOORS     = ["GF","1F","2F","3F","RF","EX","Exterior","Underground"]
HTYPES     = ["KG","QN","QNA","QNBII","TWSE","LASII","LBIII","SHC"]
WORKTYPES  = [("🪟 Finishing","Finishing"),("🏗️ Structure","Structure"),("⚡ MEP","MEP"),("🚧 Infrastructure","Infrastructure")]
VENDORS    = ["ឡាច ពៅ","វ៉ាន់ សាគីន","ថន ផល្លា","Pholla","NIPPON","Dulux"]
SUPERVISORS= ["C9","C16","C32","C45","C63","E6","E20","E22"]
ENGINEERS  = [f"C{i}" for i in range(1,76)] + [f"E{i}" for i in range(1,26)]
RAISED_BY  = ["Q2","Q3","Q6","Q7","Q10","Q11","Q12","Q13","Q16","Q17","Q21","Q26"]

DEFECTS = {
    "Finishing":     ["ឥដ្ឋប៉ោង","ឥដ្ឋមិនពេញរ៉ង","ជញ្ជាំងបូកប៉ោង","ជ្រុងមិនត្រង់","សាច់បូកខ្លុក","ការ៉ូខ្លុក","ការ៉ូបែក","ម៉ាបបែក","ថ្នាំមិនស្អាត"],
    "Structure":     ["បេតុងសសុះចេញដែក","បេតុងប៉ោង","តំទ្បើងដែកខុសប្លង់"],
    "MEP":           ["ទុយោលិច","ឆ្លងភ្លើង"],
    "Infrastructure":["ផ្លូវខូច","លូទឹកស្ទះ","បណ្តាញទឹកលិច","ទ្រនាប់ខូច","ជញ្ជាំងព័រទ្រ"],
}

# ─────────────────────────────────────────────
# EXCEL
# ─────────────────────────────────────────────
HEADERS = ["No","Date","Picture Before","Picture After","Zone","Block","# House / Unit",
           "Floor","House Type","Action Required","Defect Raised By","Supervisor In Charge",
           "Site Engineer in Charge","Type of Work","Vendor/ SubCon","Responsible by",
           "Status","Type of Comments","Remark"]
H_FILL  = PatternFill("solid", start_color="1F3864")
H_FONT  = Font(name="Arial", bold=True, color="FFFFFF", size=10)
Y_FILL  = PatternFill("solid", start_color="FFF2CC")
G_FILL  = PatternFill("solid", start_color="E2EFDA")
BORDER  = Border(left=Side(style="thin"),right=Side(style="thin"),top=Side(style="thin"),bottom=Side(style="thin"))

def init_excel():
    if os.path.exists(EXCEL_FILE): return
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "QAQC Reports"
    ws["A1"] = "QUALITY MONITORING AND IMPROVEMENT"
    ws["A1"].font = Font(name="Arial", bold=True, size=13, color="1F3864")
    ws["A2"] = "PROJECT TITLE: GRAND PHNOM PENH CITY"
    ws["A2"].font = Font(name="Arial", bold=True, size=11)
    ws.merge_cells("A1:S1"); ws.merge_cells("A2:S2")
    for col, h in enumerate(HEADERS, 1):
        c = ws.cell(row=3, column=col, value=h)
        c.fill = H_FILL; c.font = H_FONT
        c.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        c.border = BORDER
    for i,w in enumerate([5,12,15,15,8,8,12,10,12,40,15,15,15,12,18,14,12,15,20],1):
        ws.column_dimensions[openpyxl.utils.get_column_letter(i)].width = w
    wb.save(EXCEL_FILE)

def save_report(d):
    init_excel()
    wb = openpyxl.load_workbook(EXCEL_FILE)
    ws = wb.active
    row = ws.max_row + 1
    no  = max(0, ws.max_row - 3) + 1
    fill = Y_FILL if d.get("status") == "Open" else G_FILL
    vals = [no, d.get("date",""), "", "", d.get("zone",""), d.get("block",""),
            d.get("unit",""), d.get("floor",""), d.get("htype",""), d.get("action",""),
            d.get("raised",""), d.get("supervisor",""), d.get("engineer",""),
            d.get("worktype",""), d.get("vendor",""), "Site", d.get("status","Open"), "1", d.get("comment","")]
    for col, val in enumerate(vals, 1):
        c = ws.cell(row=row, column=col, value=val)
        c.fill = fill; c.font = Font(name="Arial", size=10)
        c.alignment = Alignment(vertical="center", wrap_text=True)
        c.border = BORDER
    # Photo link
    for url, col in [(d.get("photo_url",""), 3)]:
        if url:
            c = ws.cell(row=row, column=col, value="📷 BEFORE")
            c.hyperlink = url
            c.font = Font(name="Arial", size=10, color="0563C1", underline="single")
            c.fill = fill; c.border = BORDER
            c.alignment = Alignment(horizontal="center", vertical="center")
    ws.row_dimensions[row].height = 18
    wb.save(EXCEL_FILE)
    return no

def update_report(no, status, remark, photo_url=""):
    if not os.path.exists(EXCEL_FILE): return
    wb = openpyxl.load_workbook(EXCEL_FILE)
    ws = wb.active
    for row in ws.iter_rows(min_row=4):
        if row[0].value == no:
            row[16].value = status
            row[18].value = remark
            if photo_url:
                c = ws.cell(row=row[0].row, column=4, value="✅ AFTER")
                c.hyperlink = photo_url
                c.font = Font(name="Arial", size=10, color="0563C1", underline="single")
                c.alignment = Alignment(horizontal="center", vertical="center")
            fill = G_FILL if status == "Closed" else Y_FILL
            for cell in row: cell.fill = fill
            break
    wb.save(EXCEL_FILE)

def all_reports():
    if not os.path.exists(EXCEL_FILE): return []
    wb = openpyxl.load_workbook(EXCEL_FILE, data_only=True)
    ws = wb.active
    return [{"no":r[0],"date":r[1],"zone":r[4],"block":r[5],"unit":r[6],"action":r[9],"status":r[16]}
            for r in ws.iter_rows(min_row=4, values_only=True) if r[0]]

# ─────────────────────────────────────────────
# HELPERS
# ─────────────────────────────────────────────
SE = {"Open":"🟡","In Progress":"🔵","Closed":"🟢"}
def sem(s): return SE.get(s,"⚪")
def pbar(n,t=13): return "●"*n + "○"*(t-n) + f" {n}/{t}"

def kb(opts, prefix, row=8, custom=True):
    rows, r = [], []
    for i,o in enumerate(opts,1):
        r.append(InlineKeyboardButton(o, callback_data=f"{prefix}:{o}"))
        if i%row==0: rows.append(r); r=[]
    if r: rows.append(r)
    if custom: rows.append([InlineKeyboardButton("✏️ Type custom...", callback_data=f"{prefix}:__c__")])
    return InlineKeyboardMarkup(rows)

def summary(d, no=None):
    return (
        f"📋 *QAQC DEFECT REPORT*\n{'─'*25}\n"
        + (f"#️⃣ Report: *{no}*\n" if no else "")
        + f"📅 {d.get('date','')}\n"
        f"📍 Zone {d.get('zone','')} / Block {d.get('block','')} / Unit {d.get('unit','')}\n"
        f"🏠 {d.get('floor','')} — {d.get('htype','')}\n"
        f"⚠️ *{d.get('action','')}* ({d.get('worktype','')})\n"
        f"🔧 Vendor: {d.get('vendor','')}\n"
        f"👷 Sup: {d.get('supervisor','')} | Eng: {d.get('engineer','')}\n"
        f"🔎 Raised: {d.get('raised','')}\n"
        f"📌 {sem(d.get('status','Open'))} {d.get('status','Open')}\n"
        f"💬 {d.get('comment','—')}"
    )

async def edit(update, ctx, text, markup=None):
    chat_id = update.effective_chat.id
    mid = ctx.user_data.get("mid")
    if update.callback_query:
        try:
            await update.callback_query.edit_message_text(text, parse_mode="Markdown", reply_markup=markup)
            ctx.user_data["mid"] = update.callback_query.message.message_id
            return
        except: pass
    if mid:
        try:
            await ctx.bot.edit_message_text(chat_id=chat_id, message_id=mid, text=text, parse_mode="Markdown", reply_markup=markup)
            return
        except: pass
    m = await ctx.bot.send_message(chat_id=chat_id, text=text, parse_mode="Markdown", reply_markup=markup)
    ctx.user_data["mid"] = m.message_id

async def get_file_url(bot, file_id):
    try:
        f = await bot.get_file(file_id)
        return f.file_path
    except: return ""

# ─────────────────────────────────────────────
# /START /HELP /PING /RESTART
# ─────────────────────────────────────────────
async def start(update, ctx):
    await update.message.reply_text(
        "👋 *Improvement Update Bot is ready!*\n\n"
        "• /report — For QAQC បំពេញឯកសារ Improvement Case\n"
        "• /update — For Site Team បំពេញការជួសជុល\n"
        "• /list — View open defects\n"
        "• /export — Download Excel report\n"
        "• /ping — Check bot status",
        parse_mode="Markdown"
    )

async def ping(update, ctx):
    rs = all_reports()
    o = sum(1 for r in rs if r["status"]=="Open")
    c = sum(1 for r in rs if r["status"]=="Closed")
    await update.message.reply_text(
        f"🟢 *Bot is ALIVE!*\n"
        f"🕐 {datetime.now().strftime('%d %b %Y %H:%M:%S')}\n"
        f"📊 {len(rs)} total | 🟡 {o} open | 🟢 {c} closed\n"
        f"✅ All systems running normally!",
        parse_mode="Markdown"
    )

async def restart(update, ctx):
    if update.effective_user.id not in ADMIN_IDS:
        await update.message.reply_text("❌ Admin only.")
        return
    await update.message.reply_text("🔄 Restarting bot... back in 5 seconds!")
    await ctx.application.stop()
    await ctx.application.shutdown()
    os._exit(0)

# ─────────────────────────────────────────────
# /REPORT FLOW
# ─────────────────────────────────────────────
async def report_start(update, ctx):
    ctx.user_data.clear()
    ctx.user_data["date"] = datetime.now().strftime("%Y-%m-%d")
    ctx.user_data["status"] = "Open"
    try: await update.message.delete()
    except: pass
    m = await update.effective_chat.send_message(
        f"📋 *New Defect Report*\n{pbar(1)}\n\nSelect *Zone* (01–50):",
        parse_mode="Markdown",
        reply_markup=kb(ZONES, "z", row=8, custom=False)
    )
    ctx.user_data["mid"] = m.message_id
    return S_ZONE

async def s_zone(update, ctx):
    q = update.callback_query; await q.answer()
    v = q.data.split(":",1)[1]
    if v=="__c__":
        await edit(update,ctx,f"{pbar(1)}\n\n✏️ Type Zone number:")
        return S_ZONE
    ctx.user_data["zone"] = v
    await edit(update,ctx,f"Zone: *{v}* ✅\n\n{pbar(2)}\n\nSelect *Block* (01–15):",
               kb(BLOCKS,"b",row=5,custom=False))
    return S_BLOCK

async def s_zone_text(update, ctx):
    ctx.user_data["zone"] = update.message.text.strip()
    try: await update.message.delete()
    except: pass
    await edit(update,ctx,f"Zone: *{ctx.user_data['zone']}* ✅\n\n{pbar(2)}\n\nSelect *Block*:",
               kb(BLOCKS,"b",row=5,custom=False))
    return S_BLOCK

async def s_block(update, ctx):
    q = update.callback_query; await q.answer()
    v = q.data.split(":",1)[1]
    if v=="__c__":
        await edit(update,ctx,f"{pbar(2)}\n\n✏️ Type Block number:")
        return S_BLOCK
    ctx.user_data["block"] = v
    await edit(update,ctx,f"Block: *{v}* ✅\n\n{pbar(3)}\n\nSelect *Unit* (01–99):",
               kb(UNITS,"u",row=10,custom=True))
    return S_UNIT

async def s_block_text(update, ctx):
    ctx.user_data["block"] = update.message.text.strip()
    try: await update.message.delete()
    except: pass
    await edit(update,ctx,f"Block: *{ctx.user_data['block']}* ✅\n\n{pbar(3)}\n\nSelect *Unit*:",
               kb(UNITS,"u",row=10,custom=True))
    return S_UNIT

async def s_unit(update, ctx):
    q = update.callback_query; await q.answer()
    v = q.data.split(":",1)[1]
    if v=="__c__":
        await edit(update,ctx,f"{pbar(3)}\n\n✏️ Type Unit/House number:")
        return S_UNIT
    ctx.user_data["unit"] = v
    await edit(update,ctx,f"Unit: *{v}* ✅\n\n{pbar(4)}\n\nSelect *Floor*:",
               kb(FLOORS,"f",row=4,custom=False))
    return S_FLOOR

async def s_unit_text(update, ctx):
    ctx.user_data["unit"] = update.message.text.strip()
    try: await update.message.delete()
    except: pass
    await edit(update,ctx,f"Unit: *{ctx.user_data['unit']}* ✅\n\n{pbar(4)}\n\nSelect *Floor*:",
               kb(FLOORS,"f",row=4,custom=False))
    return S_FLOOR

async def s_floor(update, ctx):
    q = update.callback_query; await q.answer()
    v = q.data.split(":",1)[1]
    ctx.user_data["floor"] = v
    await edit(update,ctx,f"Floor: *{v}* ✅\n\n{pbar(5)}\n\nSelect *House Type*:",
               kb(HTYPES,"ht",row=4,custom=False))
    return S_HTYPE

async def s_htype(update, ctx):
    q = update.callback_query; await q.answer()
    v = q.data.split(":",1)[1]
    if v=="__c__":
        await edit(update,ctx,f"{pbar(5)}\n\n✏️ Type House Type:")
        return S_HTYPE
    ctx.user_data["htype"] = v
    rows=[]
    r=[]
    for i,(lbl,val) in enumerate(WORKTYPES,1):
        r.append(InlineKeyboardButton(lbl,callback_data=f"wt:{val}"))
        if i%2==0: rows.append(r); r=[]
    if r: rows.append(r)
    await edit(update,ctx,f"House Type: *{v}* ✅\n\n{pbar(6)}\n\nSelect *Type of Work*:",
               InlineKeyboardMarkup(rows))
    return S_WORKTYPE

async def s_htype_text(update, ctx):
    ctx.user_data["htype"] = update.message.text.strip()
    try: await update.message.delete()
    except: pass
    rows=[]
    r=[]
    for i,(lbl,val) in enumerate(WORKTYPES,1):
        r.append(InlineKeyboardButton(lbl,callback_data=f"wt:{val}"))
        if i%2==0: rows.append(r); r=[]
    if r: rows.append(r)
    await edit(update,ctx,f"House Type: *{ctx.user_data['htype']}* ✅\n\n{pbar(6)}\n\nSelect *Type of Work*:",
               InlineKeyboardMarkup(rows))
    return S_WORKTYPE

async def s_worktype(update, ctx):
    q = update.callback_query; await q.answer()
    v = q.data.split(":",1)[1]
    ctx.user_data["worktype"] = v
    defects = DEFECTS.get(v, [])
    await edit(update,ctx,f"Work: *{v}* ✅\n\n{pbar(7)}\n\n⚠️ Select *Defect*:",
               kb(defects,"act",row=2,custom=True))
    return S_ACTION

async def s_action(update, ctx):
    q = update.callback_query; await q.answer()
    v = q.data.split(":",1)[1]
    if v=="__c__":
        await edit(update,ctx,f"{pbar(7)}\n\n✏️ Type the defect description:")
        return S_ACTION
    ctx.user_data["action"] = v
    await edit(update,ctx,f"Defect: *{v}* ✅\n\n{pbar(8)}\n\nSelect *Vendor / SubCon*:",
               kb(VENDORS,"vn",row=2,custom=True))
    return S_VENDOR

async def s_action_text(update, ctx):
    ctx.user_data["action"] = update.message.text.strip()
    try: await update.message.delete()
    except: pass
    await edit(update,ctx,f"Defect noted ✅\n\n{pbar(8)}\n\nSelect *Vendor / SubCon*:",
               kb(VENDORS,"vn",row=2,custom=True))
    return S_VENDOR

async def s_vendor(update, ctx):
    q = update.callback_query; await q.answer()
    v = q.data.split(":",1)[1]
    if v=="__c__":
        await edit(update,ctx,f"{pbar(8)}\n\n✏️ Type Vendor name:")
        return S_VENDOR
    ctx.user_data["vendor"] = v
    await edit(update,ctx,f"Vendor: *{v}* ✅\n\n{pbar(9)}\n\nSelect *Supervisor*:",
               kb(SUPERVISORS,"sv",row=4,custom=True))
    return S_SUPERVISOR

async def s_vendor_text(update, ctx):
    ctx.user_data["vendor"] = update.message.text.strip()
    try: await update.message.delete()
    except: pass
    await edit(update,ctx,f"Vendor: *{ctx.user_data['vendor']}* ✅\n\n{pbar(9)}\n\nSelect *Supervisor*:",
               kb(SUPERVISORS,"sv",row=4,custom=True))
    return S_SUPERVISOR

async def s_supervisor(update, ctx):
    q = update.callback_query; await q.answer()
    v = q.data.split(":",1)[1]
    if v=="__c__":
        await edit(update,ctx,f"{pbar(9)}\n\n✏️ Type Supervisor code:")
        return S_SUPERVISOR
    ctx.user_data["supervisor"] = v
    await edit(update,ctx,f"Supervisor: *{v}* ✅\n\n{pbar(10)}\n\nSelect *Site Engineer*:",
               kb(ENGINEERS,"en",row=6,custom=True))
    return S_ENGINEER

async def s_supervisor_text(update, ctx):
    ctx.user_data["supervisor"] = update.message.text.strip()
    try: await update.message.delete()
    except: pass
    await edit(update,ctx,f"Supervisor: *{ctx.user_data['supervisor']}* ✅\n\n{pbar(10)}\n\nSelect *Site Engineer*:",
               kb(ENGINEERS,"en",row=6,custom=True))
    return S_ENGINEER

async def s_engineer(update, ctx):
    q = update.callback_query; await q.answer()
    v = q.data.split(":",1)[1]
    if v=="__c__":
        await edit(update,ctx,f"{pbar(10)}\n\n✏️ Type Engineer code:")
        return S_ENGINEER
    ctx.user_data["engineer"] = v
    await edit(update,ctx,f"Engineer: *{v}* ✅\n\n{pbar(11)}\n\nSelect *Raised By* (QAQC):",
               kb(RAISED_BY,"rb",row=4,custom=True))
    return S_RAISED

async def s_engineer_text(update, ctx):
    ctx.user_data["engineer"] = update.message.text.strip()
    try: await update.message.delete()
    except: pass
    await edit(update,ctx,f"Engineer: *{ctx.user_data['engineer']}* ✅\n\n{pbar(11)}\n\nSelect *Raised By*:",
               kb(RAISED_BY,"rb",row=4,custom=True))
    return S_RAISED

async def s_raised(update, ctx):
    q = update.callback_query; await q.answer()
    v = q.data.split(":",1)[1]
    if v=="__c__":
        await edit(update,ctx,f"{pbar(11)}\n\n✏️ Type QAQC code:")
        return S_RAISED
    ctx.user_data["raised"] = v
    await edit(update,ctx,
               f"Raised by: *{v}* ✅\n\n{pbar(12)}\n\n📷 Send *BEFORE photo*\n_(or tap Skip)_",
               InlineKeyboardMarkup([[InlineKeyboardButton("⏭️ Skip photo",callback_data="ph:skip")]]))
    return S_PHOTO

async def s_raised_text(update, ctx):
    ctx.user_data["raised"] = update.message.text.strip()
    try: await update.message.delete()
    except: pass
    await edit(update,ctx,
               f"Raised by: *{ctx.user_data['raised']}* ✅\n\n{pbar(12)}\n\n📷 Send *BEFORE photo*\n_(or tap Skip)_",
               InlineKeyboardMarkup([[InlineKeyboardButton("⏭️ Skip photo",callback_data="ph:skip")]]))
    return S_PHOTO

async def s_photo_skip(update, ctx):
    q = update.callback_query; await q.answer()
    ctx.user_data["photo_url"] = ""
    ctx.user_data["photo_file_id"] = ""
    await edit(update,ctx,f"{pbar(13)}\n\n💬 Add *comment* for site team:\n_(or type - to skip)_")
    return S_COMMENT

async def s_photo(update, ctx):
    fid = update.message.photo[-1].file_id
    ctx.user_data["photo_file_id"] = fid
    ctx.user_data["photo_url"] = await get_file_url(update.get_bot(), fid)
    try: await update.message.delete()
    except: pass
    await edit(update,ctx,f"📷 Photo received ✅\n\n{pbar(13)}\n\n💬 Add *comment*:\n_(or type - to skip)_")
    return S_COMMENT

async def s_comment(update, ctx):
    v = update.message.text.strip()
    ctx.user_data["comment"] = "" if v=="-" else v
    try: await update.message.delete()
    except: pass
    await edit(update,ctx,
               summary(ctx.user_data)+"\n\nSave this report?",
               InlineKeyboardMarkup([[
                   InlineKeyboardButton("✅ Confirm & Save",callback_data="cf:yes"),
                   InlineKeyboardButton("❌ Cancel",callback_data="cf:no")
               ]]))
    return S_CONFIRM

async def s_confirm(update, ctx):
    q = update.callback_query; await q.answer()
    if q.data=="cf:yes":
        no = save_report(ctx.user_data)
        await q.edit_message_text(
            f"✅ *Report #{no} saved!*\n\n"+summary(ctx.user_data,no)+"\n\n📊 /export to download Excel",
            parse_mode="Markdown"
        )
        # Send BEFORE photo in group
        fid = ctx.user_data.get("photo_file_id","")
        if fid:
            try:
                await ctx.bot.send_photo(
                    chat_id=update.effective_chat.id, photo=fid,
                    caption=f"📷 *BEFORE — Report #{no}*", parse_mode="Markdown"
                )
            except: pass
        # Alert message
        sup = ctx.user_data.get("supervisor","")
        eng = ctx.user_data.get("engineer","")
        alert = (
            f"🚨 *DEFECT ALERT — Report #{no}*\n{'─'*25}\n"
            f"📍 Zone {ctx.user_data.get('zone','')} / Block {ctx.user_data.get('block','')} / Unit {ctx.user_data.get('unit','')}\n"
            f"🏠 {ctx.user_data.get('floor','')} — {ctx.user_data.get('htype','')}\n"
            f"⚠️ *{ctx.user_data.get('action','')}* ({ctx.user_data.get('worktype','')})\n"
            f"{'─'*25}\n"
            f"👷 Supervisor: *{sup}*\n"
            f"🔧 Engineer: *{eng}*\n"
            f"🔎 Raised by: *{ctx.user_data.get('raised','')}*\n"
            f"{'─'*25}\n"
            f"📌 🟡 Open — please fix and /update to close!"
        )
        await ctx.bot.send_message(chat_id=update.effective_chat.id, text=alert, parse_mode="Markdown")
    else:
        await q.edit_message_text("❌ Report cancelled.")
    ctx.user_data.clear()
    return ConversationHandler.END

# ─────────────────────────────────────────────
# /UPDATE FLOW
# ─────────────────────────────────────────────
async def update_start(update, ctx):
    ctx.user_data.clear()
    reports = [r for r in all_reports() if r["status"]!="Closed"]
    if not reports:
        await update.message.reply_text("✅ No open defects.")
        return ConversationHandler.END
    lines = "\n".join([f"{sem(r['status'])} *#{r['no']}* Z{r['zone']}/B{r['block']} — {str(r['action'])[:30]}..." for r in reports[-15:]])
    rows,r=[],[]
    for i,rep in enumerate(reports[-15:],1):
        r.append(InlineKeyboardButton(f"#{rep['no']}",callback_data=f"us:{rep['no']}"))
        if i%5==0: rows.append(r); r=[]
    if r: rows.append(r)
    try: await update.message.delete()
    except: pass
    m = await update.effective_chat.send_message(
        f"📋 *Open Defects:*\n{lines}\n\nTap report to update:",
        parse_mode="Markdown", reply_markup=InlineKeyboardMarkup(rows)
    )
    ctx.user_data["mid"] = m.message_id
    return U_SELECT

async def u_select(update, ctx):
    q = update.callback_query; await q.answer()
    ctx.user_data["upd_no"] = int(q.data.split(":",1)[1])
    await edit(update,ctx,f"Report #{ctx.user_data['upd_no']} — Select new *Status*:",
               InlineKeyboardMarkup([[
                   InlineKeyboardButton("🟡 Open",callback_data="ust:Open"),
                   InlineKeyboardButton("🔵 In Progress",callback_data="ust:In Progress"),
                   InlineKeyboardButton("🟢 Closed",callback_data="ust:Closed"),
               ]]))
    return U_STATUS

async def u_status(update, ctx):
    q = update.callback_query; await q.answer()
    ctx.user_data["upd_status"] = q.data.split(":",1)[1]
    await edit(update,ctx,
               f"Status: *{ctx.user_data['upd_status']}* ✅\n\n💬 Describe what was done:\n_(or type - to skip)_")
    return U_COMMENT

async def u_comment(update, ctx):
    v = update.message.text.strip()
    ctx.user_data["upd_comment"] = "" if v=="-" else v
    try: await update.message.delete()
    except: pass
    await edit(update,ctx,
               "📷 Send *AFTER photo*:\n_(or tap Skip)_",
               InlineKeyboardMarkup([[InlineKeyboardButton("⏭️ Skip photo",callback_data="uph:skip")]]))
    return U_PHOTO

async def u_photo_skip(update, ctx):
    q = update.callback_query; await q.answer()
    await finish_update(update, ctx, "", "")
    return ConversationHandler.END

async def u_photo(update, ctx):
    fid = update.message.photo[-1].file_id
    url = await get_file_url(update.get_bot(), fid)
    try: await update.message.delete()
    except: pass
    await finish_update(update, ctx, fid, url)
    return ConversationHandler.END

async def finish_update(update, ctx, fid, url):
    no     = ctx.user_data["upd_no"]
    status = ctx.user_data["upd_status"]
    comment= ctx.user_data["upd_comment"]
    update_report(no, status, comment, url)
    await edit(update,ctx,
               f"✅ *Report #{no} updated!*\n"
               f"Status → {sem(status)} *{status}*\n"
               f"💬 {comment or '—'}")
    if fid:
        try:
            await ctx.bot.send_photo(
                chat_id=update.effective_chat.id, photo=fid,
                caption=f"✅ *AFTER photo — Report #{no}*\nStatus: {sem(status)} {status}",
                parse_mode="Markdown"
            )
        except: pass
    ctx.user_data.clear()

# ─────────────────────────────────────────────
# /LIST /EXPORT
# ─────────────────────────────────────────────
async def list_cmd(update, ctx):
    rs = all_reports()
    if not rs:
        await update.message.reply_text("📭 No reports yet. Use /report to log the first defect.")
        return
    o=sum(1 for r in rs if r["status"]=="Open")
    p=sum(1 for r in rs if r["status"]=="In Progress")
    c=sum(1 for r in rs if r["status"]=="Closed")
    lines="\n".join([f"{sem(r['status'])} *#{r['no']}* {r['date']} | Z{r['zone']}/B{r['block']} {r['unit']} | {str(r['action'])[:30]}..." for r in rs[-20:]])
    await update.message.reply_text(
        f"📊 *QAQC Report Summary*\nTotal: {len(rs)} | 🟡 Open: {o} | 🔵 Progress: {p} | 🟢 Closed: {c}\n{'─'*25}\n{lines}",
        parse_mode="Markdown"
    )

async def export_cmd(update, ctx):
    if not os.path.exists(EXCEL_FILE):
        await update.message.reply_text("📭 No reports yet.")
        return
    await update.message.reply_document(
        document=open(EXCEL_FILE,"rb"),
        filename=f"GPPC_QAQC_{datetime.now().strftime('%Y%m%d')}.xlsx",
        caption=f"📊 Quality Monitoring Report\n{datetime.now().strftime('%d %b %Y %H:%M')}"
    )

# ─────────────────────────────────────────────
# KEEP ALIVE
# ─────────────────────────────────────────────
class KAH(BaseHTTPRequestHandler):
    def do_GET(self):
        self.send_response(200); self.send_header("Content-Type","text/plain"); self.end_headers()
        self.wfile.write(b"OK")
    def do_HEAD(self):
        self.send_response(200); self.end_headers()
    def log_message(self,*a): pass

def keep_alive():
    port = int(os.environ.get("PORT",8080))
    t = threading.Thread(target=lambda: HTTPServer(("0.0.0.0",port),KAH).serve_forever())
    t.daemon = True; t.start()
    print(f"🌐 Keep-alive on port {port}")

# ─────────────────────────────────────────────
# MAIN
# ─────────────────────────────────────────────
def main():
    keep_alive()
    init_excel()

    app = (
        ApplicationBuilder()
        .token(BOT_TOKEN)
        .connect_timeout(30).read_timeout(30).write_timeout(30)
        .build()
    )

    report_conv = ConversationHandler(
        entry_points=[CommandHandler("report", report_start)],
        allow_reentry=True,
        states={
            S_ZONE:       [CallbackQueryHandler(s_zone,pattern="^z:"),       MessageHandler(filters.TEXT&~filters.COMMAND,s_zone_text)],
            S_BLOCK:      [CallbackQueryHandler(s_block,pattern="^b:"),      MessageHandler(filters.TEXT&~filters.COMMAND,s_block_text)],
            S_UNIT:       [CallbackQueryHandler(s_unit,pattern="^u:"),       MessageHandler(filters.TEXT&~filters.COMMAND,s_unit_text)],
            S_FLOOR:      [CallbackQueryHandler(s_floor,pattern="^f:")],
            S_HTYPE:      [CallbackQueryHandler(s_htype,pattern="^ht:"),     MessageHandler(filters.TEXT&~filters.COMMAND,s_htype_text)],
            S_WORKTYPE:   [CallbackQueryHandler(s_worktype,pattern="^wt:")],
            S_ACTION:     [CallbackQueryHandler(s_action,pattern="^act:"),   MessageHandler(filters.TEXT&~filters.COMMAND,s_action_text)],
            S_VENDOR:     [CallbackQueryHandler(s_vendor,pattern="^vn:"),    MessageHandler(filters.TEXT&~filters.COMMAND,s_vendor_text)],
            S_SUPERVISOR: [CallbackQueryHandler(s_supervisor,pattern="^sv:"),MessageHandler(filters.TEXT&~filters.COMMAND,s_supervisor_text)],
            S_ENGINEER:   [CallbackQueryHandler(s_engineer,pattern="^en:"),  MessageHandler(filters.TEXT&~filters.COMMAND,s_engineer_text)],
            S_RAISED:     [CallbackQueryHandler(s_raised,pattern="^rb:"),    MessageHandler(filters.TEXT&~filters.COMMAND,s_raised_text)],
            S_PHOTO:      [CallbackQueryHandler(s_photo_skip,pattern="^ph:skip"), MessageHandler(filters.PHOTO,s_photo)],
            S_COMMENT:    [MessageHandler(filters.TEXT&~filters.COMMAND,s_comment)],
            S_CONFIRM:    [CallbackQueryHandler(s_confirm,pattern="^cf:")],
        },
        fallbacks=[CommandHandler("report",report_start)],
    )

    update_conv = ConversationHandler(
        entry_points=[CommandHandler("update", update_start)],
        allow_reentry=True,
        states={
            U_SELECT:  [CallbackQueryHandler(u_select,pattern="^us:")],
            U_STATUS:  [CallbackQueryHandler(u_status,pattern="^ust:")],
            U_COMMENT: [MessageHandler(filters.TEXT&~filters.COMMAND,u_comment)],
            U_PHOTO:   [CallbackQueryHandler(u_photo_skip,pattern="^uph:skip"), MessageHandler(filters.PHOTO,u_photo)],
        },
        fallbacks=[CommandHandler("update",update_start)],
    )

    app.add_handler(CommandHandler("start",  start))
    app.add_handler(CommandHandler("help",   start))
    app.add_handler(CommandHandler("ping",   ping))
    app.add_handler(CommandHandler("restart",restart))
    app.add_handler(CommandHandler("list",   list_cmd))
    app.add_handler(CommandHandler("export", export_cmd))
    app.add_handler(report_conv)
    app.add_handler(update_conv)

    print("🤖 QAQC Bot is running...")
    app.run_polling(poll_interval=0.5, timeout=20, drop_pending_updates=True)

if __name__ == "__main__":
    main()
