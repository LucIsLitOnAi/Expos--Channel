import logging
import subprocess
import os
import glob
import asyncio
from telegram import Update
from telegram.ext import ApplicationBuilder, MessageHandler, filters, ContextTypes

logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s | %(levelname)s | %(message)s"
)

BOT_TOKEN = os.environ.get("TELEGRAM_BOT_TOKEN", "")

def write_config_from_env():
    import json
    config = {
        "anthropic_api_key": os.environ.get("ANTHROPIC_API_KEY", ""),
        "email_absender": os.environ.get("EMAIL_ABSENDER", ""),
        "email_empfaenger": os.environ.get("EMAIL_EMPFAENGER", ""),
        "email_passwort": os.environ.get("EMAIL_PASSWORT", ""),
        "google_drive_hauptordner_id": os.environ.get("GOOGLE_DRIVE_HAUPTORDNER_ID", ""),
        "drive_upload_aktiv": False
    }
    config_path = os.path.join(os.path.dirname(os.path.abspath(__file__)), "config.json")
    with open(config_path, "w") as f:
        json.dump(config, f, indent=2)
    logging.info("config.json geschrieben.")

pdf_queue = asyncio.Queue()

def run_expose_processor(pdf_path: str) -> str:
    script_dir = os.path.dirname(os.path.abspath(__file__))
    result = subprocess.run(
        ["python3.11", "expose_processor.py", "--pdf", pdf_path],
        stdout=subprocess.PIPE,
        stderr=subprocess.STDOUT,
        text=True,
        cwd=script_dir
    )
    return result.stdout.strip()

async def queue_worker():
    logging.info("Queue Worker gestartet.")
    while True:
        update, context, pdf_path, file_name = await pdf_queue.get()
        try:
            logging.info(f"Verarbeitung startet: {file_name}")
            await update.message.reply_text(f"⏳ Verarbeitung läuft: {file_name}")
            output = await asyncio.to_thread(run_expose_processor, pdf_path)
            logging.info(f"Verarbeitung fertig: {file_name}")
            await update.message.reply_text(f"📋 Ergebnis:\n\n{output}")
            script_dir = os.path.dirname(os.path.abspath(__file__))
            docx_files = glob.glob(
                os.path.join(script_dir, "output/Projektierungen/**/*.docx"),
                recursive=True
            )
            if docx_files:
                newest_files = sorted(docx_files, key=os.path.getctime)[-2:]
                for docx in newest_files:
                    logging.info(f"Sende DOCX: {docx}")
                    await update.message.reply_document(document=open(docx, "rb"))
        except Exception as e:
            logging.error(f"Fehler bei {file_name}: {e}")
            await update.message.reply_text(f"❌ Fehler:\n{str(e)}")
        finally:
            if os.path.exists(pdf_path):
                os.remove(pdf_path)
            pdf_queue.task_done()

async def handle_pdf(update: Update, context: ContextTypes.DEFAULT_TYPE):
    doc = update.message.document
    if not doc.file_name.lower().endswith(".pdf"):
        await update.message.reply_text("❌ Bitte nur PDF-Dateien schicken.")
        return
    file = await context.bot.get_file(doc.file_id)
    pdf_path = f"/tmp/{doc.file_name}"
    await file.download_to_drive(pdf_path)
    logging.info(f"PDF heruntergeladen: {pdf_path}")
    queue_pos = pdf_queue.qsize() + 1
    await pdf_queue.put((update, context, pdf_path, doc.file_name))
    if queue_pos == 1:
        await update.message.reply_text(f"📄 Empfangen: {doc.file_name}\n⏳ Verarbeitung startet...")
    else:
        await update.message.reply_text(f"📄 Empfangen: {doc.file_name}\n📬 Warteschlange Position: {queue_pos}")

async def handle_other(update: Update, context: ContextTypes.DEFAULT_TYPE):
    queue_pos = pdf_queue.qsize()
    status = f"\n📬 Aktuell {queue_pos} PDF(s) in der Warteschlange." if queue_pos > 0 else ""
    await update.message.reply_text(f"👋 Schick mir ein PDF-Exposé!{status}")

async def post_init(app):
    asyncio.create_task(queue_worker())
    logging.info("Queue Worker Task erstellt.")

if __name__ == "__main__":
    write_config_from_env()
    logging.info("Bot startet...")
    app = (
        ApplicationBuilder()
        .token(BOT_TOKEN)
        .post_init(post_init)
        .build()
    )
    app.add_handler(MessageHandler(filters.Document.PDF, handle_pdf))
    app.add_handler(MessageHandler(filters.ALL, handle_other))
    app.run_polling()
