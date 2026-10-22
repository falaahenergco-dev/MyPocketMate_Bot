import os
import io
import json
import time
import logging
import tempfile
import threading
from datetime import datetime, date, timedelta, timezone, time as dtime
from http.server import HTTPServer, BaseHTTPRequestHandler

import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from PIL import Image

from telegram import Update, InlineKeyboardButton, InlineKeyboardMarkup, ReplyKeyboardMarkup
from telegram.ext import (ApplicationBuilder, CommandHandler, MessageHandler, CallbackQueryHandler,
                          filters, ContextTypes)

from google import genai
from google.genai import types
from google.oauth2.service_account import Credentials
from googleapiclient.discovery import build
from googleapiclient.http import MediaIoBaseDownload, MediaFileUpload

logging.basicConfig(format='%(asctime)s - %(name)s - %(levelname)s - %(message)s', level=logging.INFO)
logging.getLogger("httpx").setLevel(logging.WARNING)  # حتى لا يظهر توكن البوت في السجلات
logger = logging.getLogger(__name__)

# ═══════════════════════ الإعدادات (من Render → Environment) ═══════════════════════
TELEGRAM_TOKEN = os.environ.get("TELEGRAM_TOKEN")
GEMINI_API_KEY = os.environ.get("GEMINI_API_KEY")
SERVICE_ACCOUNT_JSON = os.environ.get("SERVICE_ACCOUNT_JSON")
GOOGLE_FILE_ID = os.environ.get("GOOGLE_FILE_ID", "1jf8dawjd96Cxi2H7Ik6st4R2rzsDH8Gd")
GEMINI_MODEL = os.environ.get("GEMINI_MODEL", "gemini-3.8-flash")
MONTHLY_SALARY = float(os.environ.get("MONTHLY_SALARY", "150000"))
CURRENCY = os.environ.get("CURRENCY", "دينار")
# اختياري: ضع رقم حسابك في تيليغرام حتى لا يستخدم البوت غيرك (يظهر لك عند /start)
ALLOWED_USER_IDS = {x.strip() for x in os.environ.get("ALLOWED_USER_IDS", "").split(",") if x.strip()}
# أسماء المستخدمين لمناداتهم: رقم_الحساب:الاسم مفصولة بفاصلة، مثلاً 1320344513:إيمان,987654321:علي
USER_NAMES = {}
for _pair in os.environ.get("USER_NAMES", "").split(","):
    if ":" in _pair:
        _uid, _name = _pair.split(":", 1)
        if _uid.strip() and _name.strip():
            USER_NAMES[_uid.strip()] = _name.strip()


def name_for(user):
    """اسم المستخدم من USER_NAMES، وإلا اسمه في تيليغرام."""
    if user is None:
        return None
    return USER_NAMES.get(str(user.id)) or (user.first_name or "").strip() or None


def ya(name):
    return f" يا {name}" if name else ""


# التقرير الأسبوعي: يوم الخميس، الساعة (بتوقيت العراق)
REPORT_HOUR = int(os.environ.get("REPORT_HOUR", "21"))
IRAQ_TZ = timezone(timedelta(hours=3))
# الشهر المالي يبدأ يوم نزول الراتب (مثلاً 7 = من 7 الشهر إلى 6 الشهر التالي)
MONTH_START_DAY = min(28, max(1, int(os.environ.get("MONTH_START_DAY", "7"))))
# اختياري: سعر الدولار بالدينار لعرض قيمة تقريبية (مثلاً 1450). اتركه فارغاً إذا لا تريد
USD_RATE = float(os.environ.get("USD_RATE", "0") or 0)


def now_date():
    """تاريخ اليوم بتوقيت العراق (سيرفر Render يعمل بتوقيت UTC)."""
    return datetime.now(IRAQ_TZ).date()


def add_months(d, n):
    m = d.month - 1 + n
    return date(d.year + m // 12, m % 12 + 1, min(d.day, 28))


def period_start(d):
    """بداية الشهر المالي الذي يقع فيه التاريخ d."""
    start = date(d.year, d.month, MONTH_START_DAY)
    return start if d >= start else add_months(start, -1)


def in_period(d, ref):
    if not d:
        return False
    start = period_start(ref)
    return start <= d < add_months(start, 1)


def period_label(ref):
    start = period_start(ref)
    end = add_months(start, 1) - timedelta(days=1)
    return f"{start.strftime('%d/%m')} – {end.strftime('%d/%m/%Y')}"

# ═══════════════════════ هيكل الملف ═══════════════════════
TX_SHEET = "الحركات"
SUMMARY_SHEET = "الملخص الشهري"
BUDGET_SHEET = "الميزانية"
DASH_SHEET = "الوضع المالي"
RATE_CELL = "C5"  # خانة سعر الدولار في ورقة الوضع المالي (تكتبه أنت)
TOTAL_BUDGET = "إجمالي المصاريف"

TX_HEADERS = ["التاريخ", "النوع", "التصنيف", "المبلغ", "الوصف",
              "الرصيد المتاح", "المدخرات", "الذهب", "ديون عليّ", "ديون لي", "الدولار $"]
SUMMARY_HEADERS = ["الشهر", "الدخل", "المصاريف", "الادخار", "السحب من المدخرات",
                   "صافي الادخار", "نسبة الادخار", "المتبقي من الدخل",
                   "رصيد بداية الشهر", "الرصيد المتاح (الآن / نهاية الشهر)"]
SUMMARY_WIDTHS = [12, 14, 14, 14, 18, 14, 14, 18, 18, 24]

# أنواع الحركات وتأثير كل نوع على الأرصدة:
# المفاتيح: avail = الرصيد المتاح، savings = المدخرات، gold = الذهب، owe = ديون عليّ، lent = ديون لي
TYPES = {
    "دخل":                     {"avail": +1},
    "مصروف":                   {"avail": -1},
    "ادخار":                   {"avail": -1, "savings": +1},
    "سحب من المدخرات":         {"avail": +1, "savings": -1},
    "شراء ذهب":                {"avail": -1, "gold": +1},
    "بيع ذهب":                 {"avail": +1, "gold": -1},
    "اقتراض":                  {"avail": +1, "owe": +1},
    "سداد قرض":                {"avail": -1, "owe": -1},
    "إقراض":                   {"avail": -1, "lent": +1},
    "استرداد قرض":             {"avail": +1, "lent": -1},
    "رصيد افتتاحي":            {"avail": +1},
    "رصيد افتتاحي للمدخرات":   {"savings": +1},
    "رصيد افتتاحي للذهب":      {"gold": +1},
    # تسويات: يضيفها البوت تلقائياً عندما تصحح رصيداً (المبلغ قد يكون موجباً أو سالباً)
    "تسوية الرصيد المتاح":     {"avail": +1},
    "تسوية المدخرات":          {"savings": +1},
    "تسوية الذهب":             {"gold": +1},
    "تسوية ديون عليّ":         {"owe": +1},
    "تسوية ديون لي":           {"lent": +1},
    # الدولار (المبلغ في هذه الأنواع بالدولار وليس بالدينار)
    "رصيد افتتاحي للدولار":    {"usd": +1},
    "إيداع دولار":             {"usd": +1},
    "سحب دولار":               {"usd": -1},
    "تسوية الدولار":           {"usd": +1},
    # تصريف العملة (المبلغ بالدينار): لا يُحسب دخلاً ولا مصروفاً
    "تصريف دولار إلى دينار":   {"avail": +1},
    "شراء دولار بالدينار":     {"avail": -1},
}
USD_TYPES = {"رصيد افتتاحي للدولار", "إيداع دولار", "سحب دولار", "تسوية الدولار"}
ADJUST_TYPES = {"avail": "تسوية الرصيد المتاح", "savings": "تسوية المدخرات", "gold": "تسوية الذهب",
                "owe": "تسوية ديون عليّ", "lent": "تسوية ديون لي", "usd": "تسوية الدولار"}
BALANCE_KEYS = ["avail", "savings", "gold", "owe", "lent", "usd"]
BALANCE_COLS = {"avail": "F", "savings": "G", "gold": "H", "owe": "I", "lent": "J", "usd": "K"}
BALANCE_NAMES = {"avail": "💵 الرصيد المتاح", "savings": "🏦 المدخرات بالدينار", "gold": "🪙 الذهب",
                 "owe": "📕 ديون عليّ", "lent": "📗 ديون لي", "usd": "💲 المدخرات بالدولار"}

INCOME_CATEGORIES = ["الراتب", "من العائلة", "دخل إضافي", "هدايا", "أخرى"]
EXPENSE_CATEGORIES = ["الكافيهات والمطاعم", "الجامعة", "الجيم والرياضة", "المواصلات",
                      "الاتصالات والإنترنت", "التسوق والملابس", "الصحة", "الفواتير والمنزل",
                      "هدايا للآخرين", "أخرى"]

HEADER_FILL = PatternFill("solid", fgColor="1F4E78")
HEADER_FONT = Font(bold=True, color="FFFFFF", size=12)
THIN = Side(style="thin", color="D9D9D9")
BORDER = Border(left=THIN, right=THIN, top=THIN, bottom=THIN)
AMOUNT_FMT = '#,##0'


# ═══════════════════════ خادم صغير لـ Render ═══════════════════════
class Health(BaseHTTPRequestHandler):
    def do_GET(self):
        self.send_response(200)
        self.end_headers()
        self.wfile.write(b"Bot is running!")

    def do_HEAD(self):
        self.send_response(200)
        self.end_headers()

    def log_message(self, *a):
        pass


def run_health_server():
    port = int(os.environ.get("PORT", 10000))
    HTTPServer(("0.0.0.0", port), Health).serve_forever()


# ═══════════════════════ Gemini ═══════════════════════
def build_prompt(rows=None):
    today = now_date()
    recent = "\n".join(
        f"#{r['row']} | {r['date']} | {r['type']} | {r['category']} | {int(r['amount'])} | {r['description']}"
        for r in (rows or [])[-25:]) or "(no transactions yet)"
    return f"""You are a precise personal-finance bookkeeper for an Iraqi user. Currency: Iraqi dinar.
Today is {today.isoformat()} (yesterday = {(today - timedelta(days=1)).isoformat()}).
The user writes in Arabic (often Iraqi dialect: شكد = how much, صرفت = I spent, نزل الراتب = salary arrived).

Return ONLY a JSON array. Each element is ONE of:

A) A transaction:
{{"kind":"tx","type":"<TYPE>","category":"<CATEGORY>","amount":NUMBER,"description":"short Arabic text","date":"YYYY-MM-DD"}}

TYPE must be exactly one of:
- "دخل": money received. Category: "الراتب" = his monthly salary; "من العائلة" = money from family
  (أخي / أبي / أمي / أهلي أعطوني); "دخل إضافي" = side work, bonus, selling something; "هدايا" = gifts from others.
  Put the source in the description (e.g. "من أخي").
- "مصروف": money spent
- "ادخار": moving money INTO fixed savings (ادخرت / حطيت بالمدخرات / وفرت)
- "سحب من المدخرات": taking money OUT of savings (سحبت من المدخرات)
- "شراء ذهب" / "بيع ذهب"
- "اقتراض": the user borrowed money (استلفت / اقترضت / تدينت)
- "سداد قرض": the user repaid a debt (سددت / رجعت الدين / دفعت قسط)
- "إقراض": the user lent money to someone (سلفت فلان / أقرضت)
- "استرداد قرض": someone repaid the user (رجعلي فلان فلوسي)
- "رصيد افتتاحي": the user states how much cash he has right now (عندي حالياً / رصيدي الحالي)
- "رصيد افتتاحي للمدخرات": the user states his existing savings total
- "رصيد افتتاحي للذهب": the user states the value of gold he already owns

US DOLLAR types (for these 4 the amount is in DOLLARS, not dinars):
- "رصيد افتتاحي للدولار": the user states how many dollars he already has saved
- "إيداع دولار": dollars added to his dollar savings (received / bought dollars)
- "سحب دولار": dollars taken out of his dollar savings
Exchange (amount in DINARS, not counted as income or expense):
- "تصريف دولار إلى دينار": dinars received when selling dollars -> ALSO add "سحب دولار" with the dollar amount
- "شراء دولار بالدينار": dinars paid to buy dollars -> ALSO add "إيداع دولار" with the dollar amount
Currency rules: "دولار" / "$" / "دولارات" = dollars; otherwise dinars. NEVER put a dollar amount in a dinar type.
If the user paid an expense in dollars, ask (C) how much it is in dinars, unless he said it.

CATEGORY:
- for "دخل": one of {INCOME_CATEGORIES}
- for "مصروف": one of {EXPENSE_CATEGORIES}
- for all other types: use the type name itself.

Rules:
- Amounts: convert words to numbers. ألف = 1000, مليون = 1000000, ربع مليون = 250000, نص مليون = 500000,
  "150" alone in context of salary means 150000. Never invent amounts.
- If the user says the salary arrived without an amount, use {int(MONTHLY_SALARY)}.
- "date": use today unless the user mentions another day.
- One message may contain several transactions -> return several elements.
- A receipt image: return ONE "مصروف" element with the total and a short description.

B) A question:
{{"kind":"query","query":"balance"}}  -> how much do I have / رصيدي / شكد باقي
{{"kind":"query","query":"month"}}    -> this month's summary / شكد صرفت هذا الشهر
{{"kind":"query","query":"category","category":"<EXPENSE CATEGORY>","period":"month|all"}}

E) Correct a balance to an exact value (رصيدي الصحيح / صحح المدخرات / مدخراتي بالضبط):
{{"kind":"set_balance","account":"avail|savings|gold|owe|lent|usd","value":NUMBER}}
   avail = available cash (dinar), savings = fixed savings in dinar, usd = savings in US dollars (value in dollars),
   gold, owe = debts I owe, lent = money others owe me.
   Use this (not "رصيد افتتاحي") when the user says he wants to correct / fix / update a balance,
   or when an opening balance already exists in the transactions below.

F) Delete a previous transaction (احذف / امسح / ألغِ عملية ...):
{{"kind":"delete","row":ROW_NUMBER}}

G) Edit a previous transaction (القهوة كانت 4000 مو 5000 / غير التصنيف ...):
{{"kind":"edit","row":ROW_NUMBER,"amount":NUMBER_or_null,"category":"..."_or_null,"description":"..."_or_null,"date":"YYYY-MM-DD"_or_null}}

For F and G pick ROW_NUMBER from this list of the user's recent transactions (#row | date | type | category | amount | description).
If several rows match and it is unclear which one, use C) to ask (mention the dates/amounts).
{recent}

H) Set a monthly budget limit (ميزانية / حد الصرف / لا أريد أن أصرف أكثر من):
{{"kind":"set_budget","category":"<EXPENSE CATEGORY or إجمالي المصاريف>","amount":NUMBER}}   (amount 0 = remove)

I) Budget status question (شكد باقي من الميزانية / وضع الميزانية): {{"kind":"query","query":"budget"}}
   Weekly report request (تقرير الأسبوع): {{"kind":"query","query":"week"}}

J) If the input is a VOICE message, ALSO add first: {{"kind":"heard","text":"exact Arabic transcription"}}

C) If the amount or the meaning is missing or truly ambiguous, DO NOT guess:
{{"kind":"clarify","question":"short Arabic question to the user"}}

D) Not related to money:
{{"kind":"unknown"}}"""


def model_list():
    models = [m.strip() for m in GEMINI_MODEL.split(",") if m.strip()]
    for m in ["gemini-3.8-flash", "gemini-flash-latest", "gemini-2.5-flash",
              "gemini-flash-lite-latest", "gemini-2.5-flash-lite"]:
        if m not in models:
            models.append(m)
    return models


def call_gemini(contents):
    """يجرب كل موديل؛ إذا كان مزدحماً (503/429) أو غير متاح (404) ينتقل للذي بعده."""
    client = genai.Client(api_key=GEMINI_API_KEY)
    errors = []
    for _ in range(2):
        for model in model_list():
            try:
                resp = client.models.generate_content(
                    model=model,
                    contents=contents,
                    config=types.GenerateContentConfig(
                        response_mime_type="application/json", temperature=0),
                )
                logger.info(f"Gemini model used: {model}")
                raw = resp.text.strip().replace("```json", "").replace("```", "").strip()
                data = json.loads(raw)
                if isinstance(data, dict):
                    data = data.get("items", [data])
                return data
            except Exception as e:
                msg = str(e)
                if any(c in msg for c in ("503", "429", "404", "UNAVAILABLE", "RESOURCE_EXHAUSTED", "NOT_FOUND")):
                    code = "503" if "503" in msg else "429" if "429" in msg else "404"
                    errors.append(f"{model}: {code}")
                    logger.warning(f"Gemini {model} failed ({code}), trying next")
                    continue
                raise
        time.sleep(5)
    raise RuntimeError("كل موديلات Gemini غير متاحة الآن، حاول بعد قليل.\n" + "\n".join(errors[-5:]))


# ═══════════════════════ Google Drive ═══════════════════════
def get_drive_service():
    info = json.loads(SERVICE_ACCOUNT_JSON)
    creds = Credentials.from_service_account_info(info, scopes=["https://www.googleapis.com/auth/drive"])
    return build("drive", "v3", credentials=creds)


def download_workbook():
    req = get_drive_service().files().get_media(fileId=GOOGLE_FILE_ID)
    buf = io.BytesIO()
    dl = MediaIoBaseDownload(buf, req)
    done = False
    while not done:
        _, done = dl.next_chunk()
    buf.seek(0)
    wb = openpyxl.load_workbook(buf)
    ensure_structure(wb)
    return wb


def upload_workbook(wb):
    out = io.BytesIO()
    wb.save(out)
    tmp_path = os.path.join(tempfile.gettempdir(), "pocketmate_upload.xlsx")
    with open(tmp_path, "wb") as f:
        f.write(out.getvalue())
    media = MediaFileUpload(tmp_path, mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
    get_drive_service().files().update(fileId=GOOGLE_FILE_ID, media_body=media).execute()
    try:
        os.unlink(tmp_path)
    except OSError:
        pass


# ═══════════════════════ تجهيز الأوراق ═══════════════════════
def style_header(ws, headers, widths):
    for i, (h, w) in enumerate(zip(headers, widths), start=1):
        c = ws.cell(row=1, column=i, value=h)
        c.fill, c.font, c.border = HEADER_FILL, HEADER_FONT, BORDER
        c.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        ws.column_dimensions[c.column_letter].width = w
    ws.row_dimensions[1].height = 30
    ws.freeze_panes = "A2"
    ws.sheet_view.rightToLeft = True


def ensure_structure(wb):
    """ينشئ ورقتي الحركات والملخص إذا لم تكونا موجودتين (الورقة القديمة تبقى كما هي)."""
    if TX_SHEET not in wb.sheetnames:
        ws = wb.create_sheet(TX_SHEET, 0)
        style_header(ws, TX_HEADERS, [13, 22, 22, 14, 30, 16, 16, 12, 12, 12, 12])
        ws.auto_filter.ref = "A1:K1"
    ws = wb[TX_SHEET]
    if ws.cell(row=1, column=len(TX_HEADERS)).value in (None, ""):
        # ورقة أُنشئت بنسخة سابقة: نضيف عمود الدولار ونعيد كتابة معادلات الأرصدة
        style_header(ws, TX_HEADERS, [13, 22, 22, 14, 30, 16, 16, 12, 12, 12, 12])
        ws.auto_filter.ref = "A1:K1"
        rewrite_balance_formulas(ws, 2)
    if SUMMARY_SHEET not in wb.sheetnames:
        ws = wb.create_sheet(SUMMARY_SHEET, 1)
        style_header(ws, SUMMARY_HEADERS, SUMMARY_WIDTHS)
    ws = wb[SUMMARY_SHEET]
    if ws.cell(row=1, column=len(SUMMARY_HEADERS)).value in (None, ""):
        style_header(ws, SUMMARY_HEADERS, SUMMARY_WIDTHS)
    write_summary_balance_formulas(ws)
    if BUDGET_SHEET not in wb.sheetnames:
        create_budget_sheet(wb)
    write_budget_formulas(wb[BUDGET_SHEET])
    write_dashboard(wb)
    wb.active = wb.sheetnames.index(DASH_SHEET)


def create_budget_sheet(wb):
    ws = wb.create_sheet(BUDGET_SHEET, 2)
    style_header(ws, ["التصنيف", "الحد الشهري", "المصروف هذا الشهر", "المتبقي", "النسبة"], [22, 16, 18, 16, 10])
    for i, cat in enumerate(EXPENSE_CATEGORIES + [TOTAL_BUDGET], start=2):
        for j in range(1, 6):
            c = ws.cell(row=i, column=j)
            c.border = BORDER
            c.alignment = Alignment(horizontal="center")
            c.number_format = "0%" if j == 5 else AMOUNT_FMT
            if cat == TOTAL_BUDGET:
                c.font = Font(bold=True)
        ws.cell(row=i, column=1, value=cat)


def total_formula(key, extra_cond=""):
    """مجموع أثر كل الحركات على رصيد معيّن (avail/savings/...) كمعادلة SUMIFS."""
    t = f"'{TX_SHEET}'"
    part = lambda typ: f'SUMIFS({t}!$D:$D,{t}!$B:$B,"{typ}"{extra_cond})'
    plus = [x for x, eff in TYPES.items() if eff.get(key) == 1]
    minus = [x for x, eff in TYPES.items() if eff.get(key) == -1]
    expr = "+".join(part(x) for x in plus) + "".join("-" + part(x) for x in minus)
    return "=" + (expr or "0")


def write_dashboard(wb):
    """ورقة 'الوضع المالي': كل الأرصدة والشهر الحالي بمعادلات تتحدث تلقائياً."""
    rate = None
    if DASH_SHEET in wb.sheetnames:
        rate = wb[DASH_SHEET][RATE_CELL].value
        del wb[DASH_SHEET]
    if rate in (None, "") and USD_RATE:
        rate = USD_RATE
    ws = wb.create_sheet(DASH_SHEET, 0)
    ws.sheet_view.rightToLeft = True
    ws.sheet_view.showGridLines = False
    ws.column_dimensions["A"].width = 3
    ws.column_dimensions["B"].width = 34
    ws.column_dimensions["C"].width = 22
    ws.column_dimensions["D"].width = 40

    t = f"'{TX_SHEET}'"
    S = MONTH_START_DAY
    start = f"DATE(YEAR(TODAY()),MONTH(TODAY())-IF(DAY(TODAY())<{S},1,0),{S})"
    in_month = f',{t}!$A:$A,">="&{start},{t}!$A:$A,"<"&EDATE({start},1)'
    income = lambda cat=None: (f'=SUMIFS({t}!$D:$D,{t}!$B:$B,"دخل"' +
                               (f',{t}!$C:$C,"{cat}"' if cat else "") + in_month + ")")
    month_type = lambda typ: f'=SUMIFS({t}!$D:$D,{t}!$B:$B,"{typ}"{in_month})'

    title_fill = PatternFill("solid", fgColor="1F4E78")
    section_fill = PatternFill("solid", fgColor="D9E2F3")
    total_fill = PatternFill("solid", fgColor="E2EFDA")

    def title(row, text):
        ws.merge_cells(start_row=row, start_column=2, end_row=row, end_column=4)
        c = ws.cell(row=row, column=2, value=text)
        c.font, c.fill = Font(bold=True, size=16, color="FFFFFF"), title_fill
        c.alignment = Alignment(horizontal="center", vertical="center")
        ws.row_dimensions[row].height = 32

    def section(row, text):
        for col in (2, 3, 4):
            ws.cell(row=row, column=col).fill = section_fill
        c = ws.cell(row=row, column=2, value=text)
        c.font = Font(bold=True, size=13, color="1F4E78")
        ws.row_dimensions[row].height = 24

    def line(row, label, value, note="", fmt_=AMOUNT_FMT, bold=False, fill=None):
        a = ws.cell(row=row, column=2, value=label)
        b = ws.cell(row=row, column=3, value=value)
        n = ws.cell(row=row, column=4, value=note)
        for c in (a, b, n):
            c.border = BORDER
            c.alignment = Alignment(horizontal="right" if c is not b else "center", vertical="center")
            if fill:
                c.fill = fill
        a.font = Font(bold=bold, size=12)
        b.font = Font(bold=True, size=13 if bold else 12, color="1F4E78" if bold else "000000")
        n.font = Font(size=10, color="808080")
        b.number_format = fmt_
        ws.row_dimensions[row].height = 22

    title(2, "📊 الوضع المالي")
    ws.merge_cells("B3:D3")
    c = ws["B3"]
    c.value = '="آخر تحديث: "&TEXT(TODAY(),"yyyy-mm-dd")'
    c.alignment = Alignment(horizontal="center")
    c.font = Font(size=10, color="808080")

    # سعر الدولار (يكتبه المستخدم)
    line(5, "💱 سعر الدولار (اكتبه أنت)", rate, "اكتب سعر الـ 100$ ÷ 100، مثلاً 1450. يُستخدم للتقدير فقط")
    ws[RATE_CELL].fill = PatternFill("solid", fgColor="FFF2CC")

    section(7, "💼 ما تملكه الآن")
    line(8, "💵 الرصيد المتاح للصرف", total_formula("avail"))
    line(9, "🏦 المدخرات الثابتة (دينار)", total_formula("savings"))
    line(10, "💲 المدخرات بالدولار", total_formula("usd"), "", '"$"#,##0')
    line(11, "💲 قيمة الدولار بالدينار (تقريبية)", f'=IF(N({RATE_CELL})>0,C10*{RATE_CELL},"")',
         "تحتاج سعر الدولار في الأعلى")
    line(12, "🪙 الذهب", total_formula("gold"))
    line(13, "📗 ديون لي عند الناس", total_formula("lent"))
    line(14, "📕 ديون عليّ", total_formula("owe"))
    line(15, "💎 إجمالي ما تملك", "=C8+C9+N(C11)+C12+C13-C14",
         "المتاح + المدخرات + الدولار + الذهب + ديون لك − ديون عليك", bold=True, fill=total_fill)

    section(17, "📅 الشهر الحالي (من يوم الراتب)")
    line(18, "يبدأ من", f"={start}", "", "yyyy-mm-dd")
    line(19, "ينتهي في", f"=EDATE({start},1)-1", "", "yyyy-mm-dd")
    line(20, "⏳ الأيام المتبقية", f"=EDATE({start},1)-TODAY()", "", "0")
    line(21, "💰 الراتب", income("الراتب"))
    line(22, "👨‍👩‍👦 من العائلة", income("من العائلة"))
    line(23, "➕ دخل إضافي وهدايا", f"{income('دخل إضافي')}+{income('هدايا')[1:]}+{income('أخرى')[1:]}")
    line(24, "📥 إجمالي الدخل", income(), bold=True)
    line(25, "💸 المصاريف", month_type("مصروف"))
    line(26, "🏦 صافي الادخار", f"{month_type('ادخار')}-{month_type('سحب من المدخرات')[1:]}")
    line(27, "📈 نسبة الادخار من الدخل", '=IF(C24>0,C26/C24,"")', "", "0%")
    line(28, "💵 المتبقي للصرف", "=C8", "الرصيد المتاح الآن", bold=True, fill=total_fill)
    line(29, "📆 تقدر تصرف يومياً", '=IF(C20>0,C28/C20,"")',
         "المتبقي ÷ الأيام المتبقية حتى الراتب القادم", bold=True, fill=total_fill)


def read_usd_rate(wb):
    if DASH_SHEET in wb.sheetnames:
        r = to_number(wb[DASH_SHEET][RATE_CELL].value)
        if r > 0:
            return r
    return USD_RATE


def write_budget_formulas(ws):
    """معادلات المصروف خلال الشهر المالي الحالي (من يوم MONTH_START_DAY)."""
    t = f"'{TX_SHEET}'"
    S = MONTH_START_DAY
    start = f"DATE(YEAR(TODAY()),MONTH(TODAY())-IF(DAY(TODAY())<{S},1,0),{S})"
    month = f'{t}!$A:$A,">="&{start},{t}!$A:$A,"<"&EDATE({start},1)'
    ws.cell(row=1, column=3, value="المصروف هذا الشهر")
    r = 2
    while ws.cell(row=r, column=1).value not in (None, ""):
        cat = str(ws.cell(row=r, column=1).value).strip()
        cat_filter = "" if cat == TOTAL_BUDGET else f',{t}!$C:$C,$A{r}'
        ws.cell(row=r, column=3, value=f'=SUMIFS({t}!$D:$D,{t}!$B:$B,"مصروف"{cat_filter},{month})')
        ws.cell(row=r, column=4, value=f'=IF(N(B{r})>0,B{r}-C{r},"")')
        ws.cell(row=r, column=5, value=f'=IF(N(B{r})>0,C{r}/B{r},"")')
        r += 1


def read_budgets(wb):
    ws = wb[BUDGET_SHEET]
    out = {}
    r = 2
    while ws.cell(row=r, column=1).value not in (None, ""):
        limit = to_number(ws.cell(row=r, column=2).value)
        if limit > 0:
            out[str(ws.cell(row=r, column=1).value).strip()] = limit
        r += 1
    return out


def set_budget(wb, category, amount):
    ws = wb[BUDGET_SHEET]
    r = 2
    while ws.cell(row=r, column=1).value not in (None, ""):
        if str(ws.cell(row=r, column=1).value).strip() == category:
            ws.cell(row=r, column=2, value=amount if amount > 0 else None)
            return True
        r += 1
    return False


def last_tx_row(ws):
    r = 1
    while ws.cell(row=r + 1, column=1).value not in (None, ""):
        r += 1
    return r


def balance_formula(key, r):
    """معادلة الرصيد التراكمي: رصيد الصف السابق + أثر هذه الحركة."""
    col = BALANCE_COLS[key]
    plus = [t for t, eff in TYPES.items() if eff.get(key) == 1]
    minus = [t for t, eff in TYPES.items() if eff.get(key) == -1]
    expr = f"=N({col}{r-1})"
    if plus:
        expr += "+OR(" + ",".join(f'$B{r}="{t}"' for t in plus) + f")*$D{r}"
    if minus:
        expr += "-OR(" + ",".join(f'$B{r}="{t}"' for t in minus) + f")*$D{r}"
    return expr


def ensure_month_row(wb, d):
    ws = wb[SUMMARY_SHEET]
    month_start = period_start(d)
    r = 2
    while ws.cell(row=r, column=1).value not in (None, ""):
        v = ws.cell(row=r, column=1).value
        v = v.date() if isinstance(v, datetime) else v
        if v == month_start:
            return
        r += 1
    t = f"'{TX_SHEET}'"
    rng = lambda col: f"{t}!${col}:${col}"
    period = f'{rng("A")},">="&$A{r},{rng("A")},"<"&EDATE($A{r},1)'
    sumifs = lambda typ: f'SUMIFS({rng("D")},{rng("B")},"{typ}",{period})'
    values = [
        month_start,
        f"={sumifs('دخل')}",
        f"={sumifs('مصروف')}",
        f"={sumifs('ادخار')}",
        f"={sumifs('سحب من المدخرات')}",
        f"=D{r}-E{r}",
        f"=IF(B{r}>0,F{r}/B{r},0)",
        f"=B{r}-C{r}-F{r}",
    ]
    for i, v in enumerate(values, start=1):
        c = ws.cell(row=r, column=i, value=v)
        c.border = BORDER
        c.alignment = Alignment(horizontal="center")
        c.number_format = "yyyy-mm-dd" if i == 1 else ("0%" if i == 7 else AMOUNT_FMT)
    write_summary_balance_formulas(ws)


def avail_until_formula(cond):
    """رصيد الدينار المتاح لكل الحركات التي تاريخها يحقق الشرط cond (مثلاً "<"&A2)."""
    t = f"'{TX_SHEET}'"
    part = lambda typ: f'SUMIFS({t}!$D:$D,{t}!$B:$B,"{typ}",{t}!$A:$A,{cond})'
    plus = [x for x, eff in TYPES.items() if eff.get("avail") == 1]
    minus = [x for x, eff in TYPES.items() if eff.get("avail") == -1]
    return "=" + "+".join(part(x) for x in plus) + "".join("-" + part(x) for x in minus)


def write_summary_balance_formulas(ws):
    r = 2
    while ws.cell(row=r, column=1).value not in (None, ""):
        ws.cell(row=r, column=8, value=f'=IF(B{r}>0,B{r}-C{r}-F{r},"")')  # فارغ إذا لم يُسجَّل دخل
        for col, cond in ((9, f'"<"&$A{r}'), (10, f'"<"&EDATE($A{r},1)')):
            c = ws.cell(row=r, column=col, value=avail_until_formula(cond))
            c.border = BORDER
            c.alignment = Alignment(horizontal="center")
            c.number_format = AMOUNT_FMT
        r += 1


def append_tx(wb, tx):
    ws = wb[TX_SHEET]
    r = last_tx_row(ws) + 1
    values = [tx["date"], tx["type"], tx["category"], tx["amount"], tx.get("description", "")]
    values += [balance_formula(k, r) for k in BALANCE_KEYS]
    for i, v in enumerate(values, start=1):
        c = ws.cell(row=r, column=i, value=v)
        c.border = BORDER
        c.alignment = Alignment(horizontal="center", vertical="center")
        if i == 1:
            c.number_format = "yyyy-mm-dd"
        elif i == 4 or i >= 6:
            c.number_format = AMOUNT_FMT
    ensure_month_row(wb, tx["date"])
    return r


def rewrite_balance_formulas(ws, start_row):
    """بعد حذف صف من الوسط: نعيد كتابة معادلات الأرصدة حتى تبقى متسلسلة وصحيحة."""
    for r in range(max(2, start_row), last_tx_row(ws) + 1):
        for i, k in enumerate(BALANCE_KEYS, start=6):
            ws.cell(row=r, column=i, value=balance_formula(k, r))


def delete_tx_row(wb, row):
    ws = wb[TX_SHEET]
    ws.delete_rows(row)
    rewrite_balance_formulas(ws, row)


# ═══════════════════════ الحسابات (لردود البوت) ═══════════════════════
def to_number(x):
    try:
        return float(str(x).replace(",", "").replace("٬", ""))
    except (TypeError, ValueError):
        return 0.0


def fmt(n, usd=False):
    n = float(n)
    s = f"{int(round(n)):,}" if abs(n - round(n)) < 0.005 else f"{n:,.2f}"
    return f"${s}" if usd else f"{s} {CURRENCY}"


def as_date(v):
    if isinstance(v, datetime):
        return v.date()
    if isinstance(v, date):
        return v
    try:
        return datetime.strptime(str(v)[:10], "%Y-%m-%d").date()
    except ValueError:
        return None


def read_rows(wb):
    ws = wb[TX_SHEET]
    rows = []
    for r in range(2, last_tx_row(ws) + 1):
        rows.append({
            "row": r,
            "date": as_date(ws.cell(row=r, column=1).value),
            "type": str(ws.cell(row=r, column=2).value or "").strip(),
            "category": str(ws.cell(row=r, column=3).value or "").strip(),
            "amount": to_number(ws.cell(row=r, column=4).value),
            "description": str(ws.cell(row=r, column=5).value or ""),
        })
    return rows


def balances(rows):
    b = {k: 0.0 for k in BALANCE_KEYS}
    for row in rows:
        for k, sign in TYPES.get(row["type"], {}).items():
            b[k] += sign * row["amount"]
    return b


def overview_text(wb, rows, name=None):
    """رد زر 💰 رصيدي: الأرصدة + إجمالي ما تملك + كم تقدر تصرف يومياً."""
    b = balances(rows)
    rate = read_usd_rate(wb)
    usd_value = b["usd"] * rate if rate else 0
    total = b["avail"] + b["savings"] + usd_value + b["gold"] + b["lent"] - b["owe"]
    today = now_date()
    next_start = add_months(period_start(today), 1)
    days_left = (next_start - today).days
    m = [r for r in rows if in_period(r["date"], today)]
    income = sum(r["amount"] for r in m if r["type"] == "دخل")
    spent = sum(r["amount"] for r in m if r["type"] == "مصروف")
    lines = [f"💼 وضعك المالي{ya(name)}", f"💬 تبقى لديك {fmt(b['avail'])} للصرف", "", balances_text(b, rate)]
    lines.append(f"\n💎 إجمالي ما تملك: {fmt(total)}" + ("" if rate or not b["usd"] else " (بدون الدولار)"))
    lines.append(f"\n📅 هذا الشهر ({period_label(today)}):")
    lines.append(f"• الدخل: {fmt(income)}")
    lines.append(f"• المصاريف: {fmt(spent)}")
    lines.append(f"• ⏳ باقي {days_left} يوم على الراتب القادم")
    if days_left > 0 and b["avail"] > 0:
        lines.append(f"• 📆 تقدر تصرف يومياً: {fmt(round(b['avail'] / days_left))}")
    if not rate and b["usd"]:
        lines.append("\n💡 لحساب الدولار ضمن الإجمالي، اكتب سعر الصرف في ورقة «الوضع المالي»")
    return "\n".join(lines)


def balances_text(b, rate=None):
    lines = [f"{BALANCE_NAMES['avail']}: {fmt(b['avail'])}",
             f"{BALANCE_NAMES['savings']}: {fmt(b['savings'])}"]
    if abs(b["usd"]) > 0.005:
        line = f"{BALANCE_NAMES['usd']}: {fmt(b['usd'], usd=True)}"
        rate = rate or USD_RATE
        if rate:
            line += f" (≈ {fmt(b['usd'] * rate)})"
        lines.append(line)
    for k in ("gold", "owe", "lent"):
        if abs(b[k]) > 0.005:
            lines.append(f"{BALANCE_NAMES[k]}: {fmt(b[k])}")
    return "\n".join(lines)


def month_text(rows, d=None):
    d = d or now_date()
    m = [r for r in rows if in_period(r["date"], d)]
    s = lambda t: sum(r["amount"] for r in m if r["type"] == t)
    income, spent, saved, withdrawn = s("دخل"), s("مصروف"), s("ادخار"), s("سحب من المدخرات")
    net = saved - withdrawn
    start = period_start(d)
    end = add_months(start, 1)
    avail_until = lambda limit: sum(TYPES.get(r["type"], {}).get("avail", 0) * r["amount"]
                                    for r in rows if r["date"] and r["date"] < limit)
    is_current = in_period(now_date(), d)
    lines = [f"📅 ملخص الشهر ({period_label(d)}):",
             f"• رصيد بداية الشهر: {fmt(avail_until(start))}",
             f"• الدخل: {fmt(income)}",
             f"• المصاريف: {fmt(spent)}",
             f"• صافي الادخار: {fmt(net)}" + (f" ({net / income:.0%} من الدخل)" if income else ""),
             f"• {'الرصيد المتاح الآن' if is_current else 'الرصيد في نهاية الشهر'}: {fmt(avail_until(end))}"]
    by_cat = {}
    for r in m:
        if r["type"] == "مصروف":
            by_cat[r["category"]] = by_cat.get(r["category"], 0) + r["amount"]
    if by_cat:
        lines.append("\n🧾 المصاريف حسب التصنيف:")
        for cat, v in sorted(by_cat.items(), key=lambda x: -x[1]):
            lines.append(f"• {cat}: {fmt(v)}")
    return "\n".join(lines)


def category_text(rows, category, period):
    today = now_date()
    items = [r for r in rows if r["type"] == "مصروف" and r["category"] == category
             and (period == "all" or in_period(r["date"], today))]
    label = "هذا الشهر" if period != "all" else "منذ البداية"
    return f"🧾 مصاريف [{category}] {label}: {fmt(sum(r['amount'] for r in items))} ({len(items)} عملية)"


def month_spent(rows, d, category=None):
    return sum(r["amount"] for r in rows
               if r["type"] == "مصروف" and in_period(r["date"], d)
               and (category is None or r["category"] == category))


def progress_bar(pct, width=10):
    filled = min(width, int(round(pct * width)))
    return "🟥" * filled + "⬜" * (width - filled) if pct >= 1 else \
           ("🟨" if pct >= 0.8 else "🟩") * filled + "⬜" * (width - filled)


def budget_line(label, spent, limit):
    pct = spent / limit if limit else 0
    icon = "🚨" if pct >= 1 else "⚠️" if pct >= 0.8 else "✅"
    left = limit - spent
    tail = f"باقي {fmt(left)}" if left >= 0 else f"تجاوزت بـ {fmt(-left)}"
    return f"{icon} {label}: {fmt(spent)} من {fmt(limit)} ({pct:.0%})\n{progress_bar(pct)} {tail}"


def budget_text(rows, budgets, d=None):
    d = d or now_date()
    if not budgets:
        return ("🎯 لم تحدد ميزانية بعد.\nاكتب مثلاً: ميزانية الكافيهات 30 ألف\n"
                "أو: ميزانيتي الكلية للمصاريف 120 ألف")
    lines = [f"🎯 الميزانية ({period_label(d)}):"]
    for cat, limit in budgets.items():
        spent = month_spent(rows, d, None if cat == TOTAL_BUDGET else cat)
        lines.append(budget_line(cat, spent, limit))
    return "\n\n".join(lines)


def budget_alerts(rows, budgets, categories):
    """تنبيه عند الوصول لـ 80% أو تجاوز الحد، للتصنيفات التي تأثرت بالعملية الحالية."""
    d = now_date()
    out = []
    for cat in list(dict.fromkeys(categories)) + [TOTAL_BUDGET]:
        limit = budgets.get(cat)
        if not limit:
            continue
        spent = month_spent(rows, d, None if cat == TOTAL_BUDGET else cat)
        if spent / limit >= 0.8 or cat != TOTAL_BUDGET:
            out.append(budget_line(cat, spent, limit))
    return out


def weekly_text(rows, budgets, end=None):
    end = end or now_date()
    start = end - timedelta(days=6)
    prev_start, prev_end = start - timedelta(days=7), start - timedelta(days=1)
    week = [r for r in rows if r["date"] and start <= r["date"] <= end]
    prev = [r for r in rows if r["date"] and prev_start <= r["date"] <= prev_end]
    total = lambda rs, t: sum(r["amount"] for r in rs if r["type"] == t)
    spent, prev_spent = total(week, "مصروف"), total(prev, "مصروف")
    income = total(week, "دخل")
    saved = total(week, "ادخار") - total(week, "سحب من المدخرات")

    lines = [f"📊 التقرير الأسبوعي", f"🗓️ من {start.isoformat()} إلى {end.isoformat()}", ""]
    lines.append(f"💸 المصاريف: {fmt(spent)}")
    if prev_spent:
        change = (spent - prev_spent) / prev_spent
        arrow = "🔺" if change > 0 else "🔻"
        lines.append(f"   {arrow} {abs(change):.0%} مقارنة بالأسبوع الماضي ({fmt(prev_spent)})")
    lines.append(f"   📆 المعدل اليومي: {fmt(round(spent / 7))}")
    if income:
        lines.append(f"💰 الدخل: {fmt(income)}")
    if saved:
        lines.append(f"🏦 صافي الادخار: {fmt(saved)}")

    by_cat = {}
    for r in week:
        if r["type"] == "مصروف":
            by_cat[r["category"]] = by_cat.get(r["category"], 0) + r["amount"]
    if by_cat:
        lines.append("\n🧾 أين ذهبت فلوسك:")
        for cat, v in sorted(by_cat.items(), key=lambda x: -x[1])[:6]:
            lines.append(f"• {cat}: {fmt(v)} ({v / spent:.0%})")
        biggest = max((r for r in week if r["type"] == "مصروف"), key=lambda r: r["amount"])
        lines.append(f"\n🔝 أكبر مصروف: {fmt(biggest['amount'])} — {biggest['description'] or biggest['category']}")
    else:
        lines.append("\n✨ لا توجد مصاريف مسجلة هذا الأسبوع")

    if budgets:
        lines.append("\n" + budget_text(rows, budgets, end))
    lines.append("\n" + balances_text(balances(rows)))
    return "\n".join(lines)


# ═══════════════════════ التحقق من صحة ما فهمه Gemini ═══════════════════════
def validate_tx(item):
    """يعيد (الحركة بعد التنظيف, None) أو (None, رسالة خطأ)."""
    t = str(item.get("type", "")).strip()
    if t not in TYPES:
        return None, f"لم أفهم نوع العملية ({t})"
    amount = to_number(item.get("amount"))
    if amount <= 0:
        return None, "لم أجد مبلغاً واضحاً، اكتب المبلغ بالأرقام من فضلك"
    if amount > 1_000_000_000:
        return None, f"المبلغ كبير جداً ({fmt(amount)})، تأكد منه"
    cat = str(item.get("category", "")).strip()
    if t == "دخل":
        cat = cat if cat in INCOME_CATEGORIES else "أخرى"
    elif t == "مصروف":
        cat = cat if cat in EXPENSE_CATEGORIES else "أخرى"
    else:
        cat = t
    d = as_date(item.get("date")) or now_date()
    if d > now_date():
        d = now_date()
    desc = str(item.get("description", "")).strip()[:200]
    if t == "دخل" and cat == "الراتب":
        # راتب وصل قبل يوم بداية الشهر بأيام قليلة = راتب الشهر القادم
        next_start = add_months(period_start(d), 1)
        if 0 < (next_start - d).days <= 5:
            desc = (desc + " • " if desc else "") + f"وصل مبكراً بتاريخ {d.strftime('%d/%m')}"
            d = next_start
    return {"type": t, "category": cat, "amount": amount, "date": d, "description": desc}, None


TYPE_ICONS = {"دخل": "💰", "مصروف": "🧾", "ادخار": "🏦", "سحب من المدخرات": "🔓",
              "شراء ذهب": "🪙", "بيع ذهب": "🪙", "اقتراض": "📕", "سداد قرض": "✅",
              "إقراض": "📗", "استرداد قرض": "💵", "رصيد افتتاحي للدولار": "💲", "إيداع دولار": "💲",
              "سحب دولار": "💲", "تصريف دولار إلى دينار": "💱", "شراء دولار بالدينار": "💱"}


def describe_tx(tx):
    icon = TYPE_ICONS.get(tx["type"], "📌")
    head = tx["type"] if tx["category"] == tx["type"] else f"{tx['type']} • {tx['category']}"
    line = f"{icon} {head}: {fmt(tx['amount'], usd=tx['type'] in USD_TYPES)}"
    if tx["description"]:
        line += f"\n   📝 {tx['description']}"
    if tx["date"] != now_date():
        line += f"\n   📅 {tx['date'].isoformat()}"
    return line


ACCOUNT_KEYS = {"avail", "savings", "gold", "owe", "lent", "usd"}


def process(items, wb, name=None):
    """ينفّذ ما فهمه Gemini. يعيد dict فيه: text, needs_answer, added (العمليات المضافة لزر التراجع)."""
    txs, queries, problems, clarify, heard = [], [], [], None, None
    sets, deletes, edits, budgets_set = [], [], [], []
    for it in items:
        kind = it.get("kind")
        if kind == "tx":
            tx, err = validate_tx(it)
            (txs.append(tx) if tx else problems.append(err))
        elif kind == "query":
            queries.append(it)
        elif kind == "set_balance" and it.get("account") in ACCOUNT_KEYS and it.get("value") is not None:
            sets.append(it)
        elif kind == "delete":
            deletes.append(it)
        elif kind == "edit":
            edits.append(it)
        elif kind == "set_budget":
            budgets_set.append(it)
        elif kind == "heard":
            heard = str(it.get("text", "")).strip()
        elif kind == "clarify":
            clarify = it.get("question") or "ممكن توضح أكثر؟"

    prefix = [f"🎙️ سمعت: «{heard}»"] if heard else []
    result = lambda text, needs=False, added=None: {
        "text": "\n\n".join(prefix + [text]), "needs_answer": needs, "added": added or []}

    changes = txs or sets or deletes or edits or budgets_set
    if clarify and not changes:
        return result(clarify, True)
    if not changes and not queries:
        if problems:
            return result("⚠️ " + "\n".join(problems), True)
        return result("🤔 لم أفهم. أمثلة:\n• قهوة 5000\n• نزل الراتب\n• ادخرت 50 ألف\n"
                      "• سحبت من المدخرات 20 ألف\n• شكد رصيدي؟")

    rows = read_rows(wb)
    by_row = {r["row"]: r for r in rows}
    replies, warnings, added = [], [], []
    ws = wb[TX_SHEET]

    # ── تعديل عمليات سابقة ──
    for e in edits:
        r = by_row.get(int(to_number(e.get("row"))))
        if not r:
            problems.append("لم أجد العملية المطلوب تعديلها")
            continue
        old = describe_tx({**r, "date": r["date"] or now_date()})
        if e.get("amount") not in (None, "") and to_number(e["amount"]) > 0:
            ws.cell(row=r["row"], column=4, value=to_number(e["amount"]))
        if e.get("category") and r["type"] in ("دخل", "مصروف"):
            allowed_cats = INCOME_CATEGORIES if r["type"] == "دخل" else EXPENSE_CATEGORIES
            if e["category"] in allowed_cats:
                ws.cell(row=r["row"], column=3, value=e["category"])
        if e.get("description"):
            ws.cell(row=r["row"], column=5, value=str(e["description"])[:200])
        if e.get("date") and as_date(e["date"]):
            ws.cell(row=r["row"], column=1, value=as_date(e["date"]))
            ensure_month_row(wb, as_date(e["date"]))
        new = next(x for x in read_rows(wb) if x["row"] == r["row"])
        replies.append(f"✏️ تم التعديل:\nقبل: {old}\nبعد: {describe_tx({**new, 'date': new['date'] or now_date()})}")

    # ── حذف عمليات سابقة (من الأسفل للأعلى حتى لا تتغير أرقام الصفوف) ──
    for d in sorted(deletes, key=lambda x: -int(to_number(x.get("row")))):
        r = by_row.get(int(to_number(d.get("row"))))
        if not r:
            problems.append("لم أجد العملية المطلوب حذفها")
            continue
        delete_tx_row(wb, r["row"])
        replies.append("🗑️ تم حذف:\n" + describe_tx({**r, "date": r["date"] or now_date()}))

    # ── عمليات جديدة ──
    running = balances(read_rows(wb))
    for tx in txs:
        if tx["type"] == "سحب من المدخرات" and tx["amount"] > running["savings"] + 0.005:
            warnings.append(f"⚠️ السحب ({fmt(tx['amount'])}) أكبر من المدخرات ({fmt(running['savings'])})")
        if tx["type"] == "سحب دولار" and tx["amount"] > running["usd"] + 0.005:
            warnings.append(f"⚠️ السحب ({fmt(tx['amount'], usd=True)}) أكبر من مدخرات الدولار ({fmt(running['usd'], usd=True)})")
        if tx["type"] == "سداد قرض" and tx["amount"] > running["owe"] + 0.005:
            warnings.append("⚠️ مبلغ السداد أكبر من الديون المسجلة عليك")
        for k, sign in TYPES[tx["type"]].items():
            running[k] += sign * tx["amount"]
        added.append({"row": append_tx(wb, tx), "type": tx["type"], "amount": tx["amount"]})
    if txs:
        replies.append(f"✅ تم تسجيل العملية{ya(name)}:\n" + "\n".join(describe_tx(t) for t in txs))

    # ── تصحيح رصيد إلى قيمة محددة (يُسجَّل كتسوية بالفرق) ──
    for st in sets:
        key, target = st["account"], to_number(st["value"])
        diff = target - running[key]
        if abs(diff) < 0.005:
            replies.append(f"👌 {BALANCE_NAMES[key]} أصلاً {fmt(target, usd=key == 'usd')}")
            continue
        row = append_tx(wb, {"type": ADJUST_TYPES[key], "category": ADJUST_TYPES[key], "amount": diff,
                             "date": now_date(), "description": f"تصحيح إلى {fmt(target, key == 'usd')}"})
        added.append({"row": row, "type": ADJUST_TYPES[key], "amount": diff})
        running[key] = target
        sign = "+" if diff > 0 else "−"
        u = key == "usd"
        replies.append(f"🔧 تم تصحيح {BALANCE_NAMES[key]} إلى {fmt(target, u)} (فرق {sign}{fmt(abs(diff), u)})")

    # ── الميزانية ──
    for b in budgets_set:
        cat = str(b.get("category", "")).strip()
        amount = to_number(b.get("amount"))
        if cat not in EXPENSE_CATEGORIES + [TOTAL_BUDGET]:
            problems.append(f"لا يوجد تصنيف اسمه ({cat})")
        elif set_budget(wb, cat, amount):
            replies.append(f"🎯 ميزانية [{cat}] الشهرية: {fmt(amount)}" if amount > 0
                           else f"🎯 أُلغيت ميزانية [{cat}]")

    if changes:
        upload_workbook(wb)
        rows = read_rows(wb)
        final = balances(rows)
        if final["avail"] < 0:
            warnings.append("⚠️ الرصيد المتاح بالسالب، هل نسيت تسجيل دخل أو رصيد افتتاحي؟")
        if txs or sets or deletes or edits:
            replies.append(f"💬{ya(name)}، تبقى لديك {fmt(final['avail'])} للصرف\n\n" + balances_text(final))
        # تنبيهات الميزانية للتصنيفات التي صُرف عليها الآن
        spent_cats = [t["category"] for t in txs if t["type"] == "مصروف"]
        if spent_cats:
            replies += budget_alerts(rows, read_budgets(wb), spent_cats)

    for q in queries:
        kind = q.get("query")
        if kind == "month":
            replies.append(month_text(rows))
        elif kind == "budget":
            replies.append(budget_text(rows, read_budgets(wb)))
        elif kind == "week":
            replies.append(weekly_text(rows, read_budgets(wb)))
        elif kind == "category" and q.get("category") in EXPENSE_CATEGORIES:
            replies.append(category_text(rows, q["category"], q.get("period", "month")))
        elif not changes:
            replies.append(overview_text(wb, rows, name))

    return result("\n\n".join(replies + warnings + ["⚠️ " + p for p in problems]), False, added)


def undo_added(wb, added):
    """يحذف العمليات التي أضافتها رسالة معينة (زر التراجع) بعد التأكد أنها لم تتغير."""
    rows = {r["row"]: r for r in read_rows(wb)}
    removed = []
    for a in sorted(added, key=lambda x: -x["row"]):
        r = rows.get(a["row"])
        if r and r["type"] == a["type"] and abs(r["amount"] - a["amount"]) < 0.005:
            delete_tx_row(wb, a["row"])
            removed.append(r)
    return removed


# ═══════════════════════ تيليغرام ═══════════════════════
def allowed(update: Update):
    return not ALLOWED_USER_IDS or str(update.effective_user.id) in ALLOWED_USER_IDS


BTN_BALANCE, BTN_MONTH, BTN_WEEK = "💰 رصيدي", "📅 ملخص الشهر", "📊 تقرير الأسبوع"
BTN_BUDGET, BTN_LAST, BTN_HELP = "🎯 الميزانية", "🗂️ آخر العمليات", "❓ مساعدة"
MAIN_KEYBOARD = ReplyKeyboardMarkup(
    [[BTN_BALANCE, BTN_MONTH], [BTN_WEEK, BTN_BUDGET], [BTN_LAST, BTN_HELP]],
    resize_keyboard=True, is_persistent=True,
    input_field_placeholder="اكتب أو سجّل صوتك: قهوة 5000")

HELP_TEXT = (
    "📝 اكتب أو أرسل رسالة صوتية 🎙️ مثل:\n"
    "• نزل الراتب\n"
    "• قهوة 5000\n"
    "• تكسي للجامعة 4 آلاف\n"
    "• ادخرت 50 ألف / سحبت من المدخرات 20 ألف\n"
    "• استلفت من أحمد 100 ألف / سددت لأحمد 50 ألف\n"
    "• سلفت علي 25 ألف / رجعلي علي 25 ألف\n"
    "• عندي حالياً 300 ألف ومدخراتي مليون (أول مرة فقط)\n"
    "• عندي 500 دولار مدخرة / صرّفت 100 دولار بـ 145 ألف\n"
    "• أخي أعطاني 50 ألف / بعت شي بـ 30 ألف\n\n"
    "🎯 الميزانية:\n"
    "• ميزانية الكافيهات 30 ألف\n"
    "• ميزانيتي الكلية للمصاريف 120 ألف\n\n"
    "✏️ التصحيح:\n"
    "• رصيدي الصحيح 350 ألف / صحح المدخرات إلى مليون و200\n"
    "• احذف عملية القهوة 5000\n"
    "• القهوة كانت 4000 مو 5000\n"
    "• أو اضغط ↩️ تراجع تحت أي تسجيل\n\n"
    "📸 أرسل صورة فاتورة\n"
    f"📊 يصلك تقرير أسبوعي كل خميس الساعة {REPORT_HOUR}:00\n"
    f"📅 شهرك المالي يبدأ يوم {MONTH_START_DAY} (يوم الراتب)، ويصلك فيه ملخص الشهر السابق\n\n"
    "⚙️ الأوامر: /balance /month /week /budget /last /undo /help"
)


async def safe_reply(update, coro_factory, name):
    try:
        await coro_factory()
    except Exception as e:
        logger.exception(name)
        await update.effective_message.reply_text(f"❌ خطأ: {type(e).__name__}: {str(e)[:300]}")


async def start(update: Update, context: ContextTypes.DEFAULT_TYPE):
    await update.message.reply_text(
        f"👋 أهلاً{ya(name_for(update.effective_user))}! أنا PocketMate، محاسبك الشخصي.\n"
        f"🆔 رقم حسابك: {update.effective_user.id}\n\n" + HELP_TEXT,
        reply_markup=MAIN_KEYBOARD)


async def help_cmd(update: Update, context: ContextTypes.DEFAULT_TYPE):
    await update.message.reply_text(HELP_TEXT, reply_markup=MAIN_KEYBOARD)


async def balance_cmd(update: Update, context: ContextTypes.DEFAULT_TYPE):
    if allowed(update):
        async def go():
            wb = download_workbook()
            await update.effective_message.reply_text(overview_text(wb, read_rows(wb), name_for(update.effective_user)))
        await safe_reply(update, go, "balance")


async def month_cmd(update: Update, context: ContextTypes.DEFAULT_TYPE):
    if allowed(update):
        async def go():
            wb = download_workbook()
            rows = read_rows(wb)
            text = month_text(rows)
            budgets = read_budgets(wb)
            if budgets:
                text += "\n\n" + budget_text(rows, budgets)
            await update.effective_message.reply_text(text)
        await safe_reply(update, go, "month")


async def week_cmd(update: Update, context: ContextTypes.DEFAULT_TYPE):
    if allowed(update):
        async def go():
            wb = download_workbook()
            await update.effective_message.reply_text(weekly_text(read_rows(wb), read_budgets(wb)))
        await safe_reply(update, go, "week")


async def budget_cmd(update: Update, context: ContextTypes.DEFAULT_TYPE):
    if allowed(update):
        async def go():
            wb = download_workbook()
            await update.effective_message.reply_text(budget_text(read_rows(wb), read_budgets(wb)))
        await safe_reply(update, go, "budget")


async def last_cmd(update: Update, context: ContextTypes.DEFAULT_TYPE):
    if allowed(update):
        async def go():
            rows = read_rows(download_workbook())[-10:]
            if not rows:
                await update.effective_message.reply_text("لا توجد عمليات بعد.")
                return
            lines = [f"{r['date']}  {describe_tx({**r, 'date': now_date()})}" for r in rows]
            await update.effective_message.reply_text(
                "🗂️ آخر العمليات:\n\n" + "\n".join(lines) +
                "\n\nللحذف أو التعديل اكتب مثلاً: احذف عملية القهوة 5000")
        await safe_reply(update, go, "last")


async def undo_cmd(update: Update, context: ContextTypes.DEFAULT_TYPE):
    if allowed(update):
        async def go():
            wb = download_workbook()
            rows = read_rows(wb)
            if not rows:
                await update.effective_message.reply_text("لا توجد عمليات لحذفها.")
                return
            last = rows[-1]
            delete_tx_row(wb, last["row"])
            upload_workbook(wb)
            await update.effective_message.reply_text(
                f"🗑️ حُذفت آخر عملية{ya(name_for(update.effective_user))}:\n" + describe_tx({**last, "date": last["date"] or now_date()}) +
                "\n\n" + balances_text(balances(read_rows(wb))))
        await safe_reply(update, go, "undo")


def undo_markup(context, added):
    """زر ↩️ تراجع تحت الرسالة، مربوط بالعمليات التي أضافتها هذه الرسالة فقط."""
    if not added:
        return None
    token = str(int(time.time() * 1000))[-9:]
    store = context.user_data.setdefault("undo", {})
    store[token] = added
    for old in list(store)[:-10]:  # نحتفظ بآخر 10 فقط
        store.pop(old, None)
    return InlineKeyboardMarkup([[InlineKeyboardButton("↩️ تراجع", callback_data=f"undo:{token}")]])


async def undo_button(update: Update, context: ContextTypes.DEFAULT_TYPE):
    query = update.callback_query
    if not allowed(update):
        await query.answer()
        return
    await query.answer("⏳ جاري التراجع...")
    token = query.data.split(":", 1)[1]
    added = context.user_data.get("undo", {}).pop(token, None)
    if not added:
        await query.edit_message_reply_markup(None)
        await query.message.reply_text("⚠️ لا يمكن التراجع عن هذه العملية (قديمة أو تم التراجع عنها). استخدم /last")
        return
    try:
        wb = download_workbook()
        removed = undo_added(wb, added)
        if removed:
            upload_workbook(wb)
        await query.edit_message_reply_markup(None)
        text = (f"↩️ تم التراجع{ya(name_for(update.effective_user))} عن:\n" + "\n".join(describe_tx({**r, "date": r["date"] or now_date()}) for r in removed)
                if removed else "⚠️ العملية تغيّرت أو حُذفت مسبقاً")
        await query.message.reply_text(text + "\n\n" + balances_text(balances(read_rows(wb))))
    except Exception as e:
        logger.exception("undo button")
        await query.message.reply_text(f"❌ خطأ: {e}")


async def run_ai(update: Update, context: ContextTypes.DEFAULT_TYPE, parts, msg, pending_text=None):
    """parts = محتوى الرسالة (نص/صورة/صوت). نقرأ الملف أولاً حتى يعرف Gemini العمليات السابقة."""
    wb = download_workbook()
    items = call_gemini([build_prompt(read_rows(wb))] + parts)
    res = process(items, wb, name_for(update.effective_user))
    # إذا سأل البوت سؤال توضيح، نحفظ الرسالة الأصلية لنربطها بالجواب
    heard = next((i.get("text") for i in items if i.get("kind") == "heard"), None)
    context.user_data["pending"] = (pending_text or heard) if res["needs_answer"] else None
    await msg.edit_text(res["text"], reply_markup=undo_markup(context, res["added"]))


BUTTON_ACTIONS = {BTN_BALANCE: balance_cmd, BTN_MONTH: month_cmd, BTN_WEEK: week_cmd,
                  BTN_BUDGET: budget_cmd, BTN_LAST: last_cmd, BTN_HELP: help_cmd}


async def handle_message(update: Update, context: ContextTypes.DEFAULT_TYPE):
    if not allowed(update):
        return
    text = update.message.text
    if text in BUTTON_ACTIONS:
        context.user_data["pending"] = None
        await BUTTON_ACTIONS[text](update, context)
        return
    msg = await update.message.reply_text("⏳ جاري المعالجة...")
    try:
        pending = context.user_data.get("pending")
        full = f"{pending}\nUser's answer to your clarification question: {text}" if pending else text
        await run_ai(update, context, [f"User: {full}"], msg, pending_text=full)
    except Exception as e:
        logger.exception("Text error")
        context.user_data["pending"] = None
        await msg.edit_text(f"❌ خطأ: {type(e).__name__}: {str(e)[:300]}")


async def handle_voice(update: Update, context: ContextTypes.DEFAULT_TYPE):
    if not allowed(update):
        return
    msg = await update.message.reply_text("🎙️ جاري الاستماع...")
    try:
        media = update.message.voice or update.message.audio
        if media.duration and media.duration > 120:
            await msg.edit_text("⚠️ الرسالة الصوتية طويلة، أرسل رسالة أقصر من دقيقتين")
            return
        data = bytes(await (await media.get_file()).download_as_bytearray())
        mime = getattr(media, "mime_type", None) or "audio/ogg"
        audio = types.Part.from_bytes(data=data, mime_type=mime)
        pending = context.user_data.get("pending")
        note = "This is a VOICE message from the user (Arabic, maybe Iraqi dialect)."
        if pending:
            note += f" It answers your clarification question about: {pending}"
        await run_ai(update, context, [audio, note], msg)
    except Exception as e:
        logger.exception("Voice error")
        context.user_data["pending"] = None
        await msg.edit_text(f"❌ خطأ: {type(e).__name__}: {str(e)[:300]}")


async def handle_photo(update: Update, context: ContextTypes.DEFAULT_TYPE):
    if not allowed(update):
        return
    msg = await update.message.reply_text("📸 جاري قراءة الفاتورة...")
    try:
        photo_file = await update.message.photo[-1].get_file()
        image = Image.open(io.BytesIO(bytes(await photo_file.download_as_bytearray())))
        caption = update.message.caption or ""
        await run_ai(update, context, [image, f"This is a receipt. User note: {caption}"], msg)
    except Exception as e:
        logger.exception("Photo error")
        await msg.edit_text(f"❌ خطأ: {type(e).__name__}: {str(e)[:300]}")


async def weekly_job(context: ContextTypes.DEFAULT_TYPE):
    """يُرسل التقرير الأسبوعي كل خميس لكل رقم في ALLOWED_USER_IDS."""
    if not ALLOWED_USER_IDS:
        logger.warning("Weekly report skipped: set ALLOWED_USER_IDS in Render")
        return
    try:
        wb = download_workbook()
        text = weekly_text(read_rows(wb), read_budgets(wb))
    except Exception as e:
        logger.exception("weekly report")
        text = f"❌ تعذّر إعداد التقرير الأسبوعي: {e}"
    for uid in ALLOWED_USER_IDS:
        name = USER_NAMES.get(uid)
        greeting = f"🌙 مساء الخير{ya(name)}، هذا تقريرك الأسبوعي\n\n"
        try:
            await context.bot.send_message(chat_id=int(uid), text=greeting + text)
        except Exception:
            logger.exception(f"send weekly to {uid}")


async def salary_day_job(context: ContextTypes.DEFAULT_TYPE):
    """يوم بداية الشهر المالي: ملخص الشهر الذي انتهى + تذكير بتسجيل الراتب."""
    if not ALLOWED_USER_IDS:
        return
    try:
        wb = download_workbook()
        rows = read_rows(wb)
        prev_month_day = now_date() - timedelta(days=1)
        text = ("🗓️ اليوم بداية شهرك المالي الجديد{NAME}!\n\n"
                "📌 ملخص الشهر الذي انتهى:\n" + month_text(rows, prev_month_day) +
                "\n\n" + balances_text(balances(rows)) +
                "\n\n💰 هل نزل الراتب؟ اكتب: نزل الراتب\n🏦 ولا تنسَ تدّخر منه، مثلاً: ادخرت 30 ألف")
    except Exception as e:
        logger.exception("salary day")
        text = "🗓️ اليوم بداية شهرك المالي{NAME}. هل نزل الراتب؟ اكتب: نزل الراتب"
    for uid in ALLOWED_USER_IDS:
        try:
            await context.bot.send_message(chat_id=int(uid), text=text.replace("{NAME}", ya(USER_NAMES.get(uid))))
        except Exception:
            logger.exception(f"send salary-day to {uid}")


def main():
    threading.Thread(target=run_health_server, daemon=True).start()
    app = ApplicationBuilder().token(TELEGRAM_TOKEN).build()
    for name, fn in [("start", start), ("help", help_cmd), ("balance", balance_cmd), ("month", month_cmd),
                     ("week", week_cmd), ("budget", budget_cmd), ("last", last_cmd), ("undo", undo_cmd)]:
        app.add_handler(CommandHandler(name, fn))
    app.add_handler(CallbackQueryHandler(undo_button, pattern=r"^undo:"))
    app.add_handler(MessageHandler(filters.VOICE | filters.AUDIO, handle_voice))
    app.add_handler(MessageHandler(filters.PHOTO, handle_photo))
    app.add_handler(MessageHandler(filters.TEXT & ~filters.COMMAND, handle_message))
    # الخميس: في python-telegram-bot الأيام 0=الأحد ... 4=الخميس
    if app.job_queue:
        app.job_queue.run_daily(weekly_job, time=dtime(hour=REPORT_HOUR, minute=0, tzinfo=IRAQ_TZ),
                                days=(4,), name="weekly_report")
        app.job_queue.run_monthly(salary_day_job, when=dtime(hour=10, minute=0, tzinfo=IRAQ_TZ),
                                  day=MONTH_START_DAY, name="salary_day")
    else:
        logger.warning("JobQueue غير متاح: تأكد أن requirements.txt فيه python-telegram-bot[job-queue]")
    logger.info("🤖 PocketMate يعمل!")
    app.run_polling()


if __name__ == '__main__':
    main()
