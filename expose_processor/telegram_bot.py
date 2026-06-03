from dotenv import load_dotenv
load_dotenv()

import logging
import subprocess
import os
import glob
import json
from pathlib import Path
from telegram import Update
from telegram.ext import ApplicationBuilder, MessageHandler, filters, ContextTypes

logging.basicConfig(level=logging.INFO)

# Config aus Umgebungsvariablen erstellen
def setup_config():
    config = {
        "anthropic_api_key": os.environ.get("ANTHROPIC_API_KEY", ""),
        "mitarbeiter_nr": "0",
        "mitarbeiter_name": "Timo Klemt",
        "email_absender": os.environ.get("EMAIL_ABSENDER", ""),
        "email_empfaenger": os.environ.get("EMAIL_EMPFAENGER", ""),
        "email_passwort": os.environ.get("EMAIL_PASSWORT", ""),
        "google_drive_hauptordner_id": os.environ.get("GOOGLE_DRIVE_HAUPTORDNER_ID", ""),
        "google_credentials_pfad": "assets/google_service_account.json",
        "template_pfad": "assets/templates/Projektierung_Template.docx",
        "output_pfad": "output/Projektierungen/",
        "log_pfad": "system/logs/stop_log.csv",
        "drive_upload_aktiv": False
    }
    with open("config.json", "w") as f:
        json.dump(config, f, indent=2)

    sa_json = os.environ.get("GOOGLE_SERVICE_ACCOUNT_JSON", "")
    if sa_json:
        Path("assets").mkdir(exist_ok=True)
        with open("assets/google_service_account.json", "w") as f:
            f.write(sa_json)

setup_config()

BOT_TOKEN = os.environ.get("TELEGRAM_BOT_TOKEN", "")
print(f"Token geladen: {BOT_TOKEN[:10]}..." if BOT_TOKEN else "KEIN TOKEN GEFUNDEN!")

async def handle_pdf(update: Update, context: ContextTypes.DEFAULT_TYPE):
    doc = update.message.document
    if not doc.file_name.lower().endswith(".pdf"):
        await update.message.reply_text("❌ Bitte nur PDF-Dateien schicken.")
        return

    await update.message.reply_text(f"📄 Empfangen: {doc.file_name}\n⏳ Verarbeitung läuft...")

    file = await context.bot.get_file(doc.file_id)
    pdf_path = f"/tmp/{doc.file_name}"
    await file.download_to_drive(pdf_path)

    result = subprocess.run(
        ["python3.11", "expose_processor.py", "--pdf", pdf_path],
        capture_output=True, text=True, cwd=os.path.dirname(os.path.abspath(__file__))
    )

    output = result.stdout.strip()
    await update.message.reply_text(f"📋 Ergebnis:\n\n{output}")

    docx_files = glob.glob("output/Projektierungen/**/*.docx", recursive=True)
    if docx_files:
        newest = max(docx_files, key=os.path.getctime)
        await update.message.reply_document(document=open(newest, "rb"))

    os.remove(pdf_path)

async def handle_other(update: Update, context: ContextTypes.DEFAULT_TYPE):
    await update.message.reply_text("👋 Schick mir ein PDF-Exposé und ich verarbeite es!")

app = ApplicationBuilder().token(BOT_TOKEN).build()
app.add_handler(MessageHandler(filters.Document.PDF, handle_pdf))
app.add_handler(MessageHandler(filters.ALL, handle_other))

if __name__ == "__main__":
    print("Bot läuft...")
    app.run_polling()
