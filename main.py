import os
import json
import logging
from io import BytesIO
from datetime import datetime
from PIL import Image

from telegram import Update
from telegram.ext import ApplicationBuilder, CommandHandler, MessageHandler, filters, ContextTypes

from google import genai
from google.genai import types
from google.oauth2.service_account import Credentials
from googleapiclient.discovery import build

logging.basicConfig(format='%(asctime)s - %(name)s - %(levelname)s - %(message)s', level=logging.INFO)

TELEGRAM_TOKEN = os.environ.get("TELEGRAM_TOKEN")
GEMINI_API_KEY = os.environ.get("GEMINI_API_KEY")
SERVICE_ACCOUNT_JSON = os.environ.get("SERVICE_ACCOUNT_JSON")
GOOGLE_FILE_ID = os.environ.get("GOOGLE_FILE_ID", "1jf8dawjd96Cxi2H7Ik6st4R2rzsDH8Gd")

SYSTEM_PROMPT = """You are a personal finance assistant managing a Google Sheet. Analyze Arabic text or receipts.

Categories mapped to columns:
- 'الكافيهات': coffee, cafes, drinks, snacks
- 'الجامعة': university, books, tuition, study
- 'الجيم': gym, fitness, sports, supplements
- 'مقتنيات ذهبية': gold, jewelry, assets
- 'القروض (Loans)': loans, debts, borrowed money
- 'هدايا/دخل إضافي': gifts, bonus, side income
- 'المدخرات الثابتة': fixed savings, deposits

Return JSON array:
[
  {
    "action": "add_expense",
    "data": {
      "category": "exact_category_name",
      "amount": NUMBER,
      "details": "description"
    }
  }
]
Return ONLY a valid JSON array."""

def get_sheets_service():
    info = json.loads(SERVICE_ACCOUNT_JSON)
    creds = Credentials.from_service_account_info(info, scopes=['https://www.googleapis.com/auth/spreadsheets'])
    return build('sheets', '4', credentials=creds)

def append_to_sheet(category, amount, details=""):
    try:
        service = get_sheets_service()
        today = datetime.now().strftime("%Y-%m-%d")
        
        # Mapping categories to sheet columns
        col_map = {
            'الكافيهات': 'B',
            'الجامعة': 'C',
            'الجيم': 'D',
            'مقتنيات ذهبية': 'E',
            'القروض (Loans)': 'F',
            'هدايا/دخل إضافي': 'G',
            'المدخرات الثابتة': 'H'
        }
        
        target_col = col_map.get(category, 'B')
        range_name = f"Sheet1!A3:K9"
        
        # Read existing rows to append properly
        sheet = service.spreadsheets()
        result = sheet.values().get(spreadsheetId=GOOGLE_FILE_ID, range=range_name).execute()
        rows = result.get('values', [])
        
        next_row = len(rows) + 3
        if next_row > 9:
            next_row = 9  # default cap for current structure
            
        value_range_body = {
            'values': [[today]]
        }
        sheet.values().update(
            spreadsheetId=GOOGLE_FILE_ID,
            range=f"Sheet1!A{next_row}",
            valueInputOption="USER_ENTERED",
            body=value_range_body
        ).execute()

        val_body = {'values': [[amount]]}
        sheet.values().update(
            spreadsheetId=GOOGLE_FILE_ID,
            range=f"Sheet1!{target_col}{next_row}",
            valueInputOption="USER_ENTERED",
            body=val_body
        ).execute()
        
        return True
    except Exception as e:
        logging.error(f"Error writing to sheet: {e}")
        return False

async def start(update: Update, context: ContextTypes.DEFAULT_TYPE):
    await update.message.reply_text("أهلاً بك! أنا بوت تتبع المصاريف والمدخرات الشخصية. يمكنك إرسال نص بالمصروف أو صورة فاتورة وسأقوم بتسجيلها مباشرة في شيت المدخرات.")

async def handle_message(update: Update, context: ContextTypes.DEFAULT_TYPE):
    user_text = update.message.text
    client = genai.Client(api_key=GEMINI_API_KEY)
    
    response = client.models.generate_content(
        model='gemini-2.5-flash',
        contents=[SYSTEM_PROMPT, f"Input text: {user_text}"],
        config=types.GenerateContentConfig(response_mime_type="application/json")
    )
    
    try:
        data = json.loads(response.text)
        for item in data:
            if item.get("action") == "add_expense":
                exp = item.get("data", {})
                cat = exp.get("category")
                amt = exp.get("amount")
                det = exp.get("details", "")
                
                success = append_to_sheet(cat, amt, det)
                if success:
                    await update.message.reply_text(f"✅ تم تسجيل المبلغ ({amt}) في عمود [{cat}] بنجاح!")
                else:
                    await update.message.reply_text("❌ حدث خطأ أثناء التحديث في Google Sheets.")
    except Exception as e:
        logging.error(f"Parsing error: {e}")
        await update.message.reply_text("لم أتمكن من فهم المدخلات بشكل دقيق، يرجى إعادة المحاولة.")

async def handle_photo(update: Update, context: ContextTypes.DEFAULT_TYPE):
    photo_file = await update.message.photo[-1].get_file()
    image_bytes = await photo_file.download_as_bytearray()
    image = Image.open(BytesIO(image_bytes))
    
    client = genai.Client(api_key=GEMINI_API_KEY)
    response = client.models.generate_content(
        model='gemini-2.5-flash',
        contents=[SYSTEM_PROMPT, image, "Extract expenses from this receipt."],
        config=types.GenerateContentConfig(response_mime_type="application/json")
    )
    
    try:
        data = json.loads(response.text)
        for item in data:
            if item.get("action") == "add_expense":
                exp = item.get("data", {})
                cat = exp.get("category")
                amt = exp.get("amount")
                
                success = append_to_sheet(cat, amt)
                if success:
                    await update.message.reply_text(f"🧾 تم تحليل الفاتورة وتسجيل ({amt}) في عمود [{cat}] بنجاح!")
                else:
                    await update.message.reply_text("❌ حدث خطأ أثناء التسجيل في Google Sheets.")
    except Exception as e:
        logging.error(f"Photo parsing error: {e}")
        await update.message.reply_text("عذراً، لم أستطع قراءة الفاتورة بوضوح.")

def main():
    app = ApplicationBuilder().token(TELEGRAM_TOKEN).build()
    app.add_handler(CommandHandler("start", start))
    app.add_handler(MessageHandler(filters.TEXT & ~filters.COMMAND, handle_message))
    app.add_handler(MessageHandler(filters.PHOTO, handle_photo))
    app.run_polling()

if __name__ == '__main__':
    main()
