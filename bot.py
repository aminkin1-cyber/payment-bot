"""
Payment Tracker Bot
Forwards messages from agent chat → Claude analyzes → morning report + Excel
"""

import os, json, asyncio, logging
from datetime import datetime, time
from pathlib import Path
import httpx
from telegram import Update, Bot
from telegram.ext import Application, MessageHandler, CommandHandler, filters, ContextTypes

# ── Config from environment variables ────────────────────────────────────────
BOT_TOKEN      = os.environ["BOT_TOKEN"]
ANTHROPIC_KEY  = os.environ["ANTHROPIC_KEY"]
MY_CHAT_ID     = int(os.environ["MY_CHAT_ID"])
MORNING_HOUR   = int(os.environ.get("MORNING_HOUR", "9"))
DATA_FILE      = Path("data/messages.json")
EXCEL_FILE     = Path("data/Agent_Model_v2.xlsx")

logging.basicConfig(level=logging.INFO)
log = logging.getLogger(__name__)

# ── Helpers ───────────────────────────────────────────────────────────────────
def load_messages():
    DATA_FILE.parent.mkdir(exist_ok=True)
    if DATA_FILE.exists():
        return json.loads(DATA_FILE.read_text())
    return []

def save_message(msg_dict):
    msgs = load_messages()
    msgs.append(msg_dict)
    DATA_FILE.write_text(json.dumps(msgs, ensure_ascii=False, indent=2))

def clear_messages():
    DATA_FILE.write_text("[]")

async def ask_claude(prompt: str) -> str:
    """Call Claude API and return text response."""
    async with httpx.AsyncClient(timeout=60) as client:
        r = await client.post(
            "https://api.anthropic.com/v1/messages",
            headers={
                "x-api-key": ANTHROPIC_KEY,
                "anthropic-version": "2023-06-01",
                "content-type": "application/json",
            },
            json={
                "model": "claude-opus-4-6",
                "max_tokens": 2000,
                "system": (
                    "You are a financial assistant tracking payments between a company and its financial agent. "
                    "The agent handles payments in AED, CNY, USD, EUR, SGD, RUB. "
                    "AED/USD rate is ~3.6725. Agent charges 0.5% commission on payments, 0.4% on RUB. "
                    "Respond in Russian. Be concise and structured."
                ),
                "messages": [{"role": "user", "content": prompt}],
            },
        )
        data = r.json()
        return data["content"][0]["text"]

# ── Command handlers ──────────────────────────────────────────────────────────
async def cmd_start(update: Update, ctx: ContextTypes.DEFAULT_TYPE):
    await update.message.reply_text(
        "👋 Привет! Я трекер платежей.\n\n"
        "📨 Пересылай мне сообщения от агента — я их запомню.\n\n"
        "Команды:\n"
        "/balance — текущий баланс\n"
        "/pending — что висит\n"
        "/summary — полное саммари\n"
        "/excel — прислать Excel\n"
        "/unknown — неизвестные транзакции\n"
        "/clear — очистить историю сообщений дня"
    )

async def cmd_balance(update: Update, ctx: ContextTypes.DEFAULT_TYPE):
    msgs = load_messages()
    if not msgs:
        await update.message.reply_text("Нет сообщений для анализа. Перешли что-нибудь от агента.")
        return
    prompt = f"Из этих сообщений определи ТОЛЬКО текущий баланс агента в USD. Ответь одной строкой.\n\nСообщения:\n" + _format_msgs(msgs)
    reply = await ask_claude(prompt)
    await update.message.reply_text(f"💰 {reply}")

async def cmd_pending(update: Update, ctx: ContextTypes.DEFAULT_TYPE):
    msgs = load_messages()
    if not msgs:
        await update.message.reply_text("Нет сообщений для анализа.")
        return
    prompt = f"Из этих сообщений выдели только НЕОПЛАЧЕННЫЕ инвойсы и платежи которые ещё не выполнены. Коротко списком.\n\nСообщения:\n" + _format_msgs(msgs)
    reply = await ask_claude(prompt)
    await update.message.reply_text(f"⏳ Ожидают оплаты:\n\n{reply}")

async def cmd_unknown(update: Update, ctx: ContextTypes.DEFAULT_TYPE):
    msgs = load_messages()
    if not msgs:
        await update.message.reply_text("Нет сообщений для анализа.")
        return
    prompt = f"Найди платежи или суммы в этих сообщениях для которых НЕТ инвойса или непонятно кому платили. Коротко.\n\nСообщения:\n" + _format_msgs(msgs)
    reply = await ask_claude(prompt)
    await update.message.reply_text(f"❓ Требуют уточнения:\n\n{reply}")

async def cmd_summary(update: Update, ctx: ContextTypes.DEFAULT_TYPE):
    await _send_morning_report(ctx.bot)

async def cmd_excel(update: Update, ctx: ContextTypes.DEFAULT_TYPE):
    if EXCEL_FILE.exists():
        await ctx.bot.send_document(
            chat_id=MY_CHAT_ID,
            document=EXCEL_FILE.open("rb"),
            filename="Agent_Model.xlsx",
            caption="📎 Актуальный Excel файл"
        )
    else:
        await update.message.reply_text("Excel файл не найден. Положи Agent_Model_v2.xlsx в папку data/")

async def cmd_clear(update: Update, ctx: ContextTypes.DEFAULT_TYPE):
    clear_messages()
    await update.message.reply_text("🗑 История сообщений очищена.")

# ── Message handler (forwarded messages) ─────────────────────────────────────
async def handle_message(update: Update, ctx: ContextTypes.DEFAULT_TYPE):
    msg = update.message
    if msg.chat_id != MY_CHAT_ID:
        return  # ignore messages not from you

    # Extract text
    text = msg.text or msg.caption or ""
    sender = ""
    if msg.forward_from:
        sender = msg.forward_from.full_name or msg.forward_from.username or "Unknown"
    elif msg.forward_sender_name:
        sender = msg.forward_sender_name

    date_str = (msg.forward_date or msg.date).strftime("%d.%m.%Y %H:%M")

    # Document/file attached?
    file_name = ""
    if msg.document:
        file_name = msg.document.file_name or "document"

    entry = {
        "date": date_str,
        "sender": sender,
        "text": text,
        "file": file_name,
    }
    save_message(entry)

    # Quick acknowledgement
    parts = []
    if sender: parts.append(f"от {sender}")
    if file_name: parts.append(f"📎 {file_name}")
    if text: parts.append(f'"{text[:60]}{"…" if len(text)>60 else ""}"')
    await msg.reply_text(f"✅ Сохранено ({date_str}): {' · '.join(parts)}")

def _format_msgs(msgs):
    lines = []
    for m in msgs:
        line = f"[{m['date']}] {m.get('sender','?')}: {m.get('text','')}"
        if m.get('file'):
            line += f" [файл: {m['file']}]"
        lines.append(line)
    return "\n".join(lines)

# ── Morning report ────────────────────────────────────────────────────────────
async def _send_morning_report(bot: Bot):
    msgs = load_messages()
    today = datetime.now().strftime("%d %B %Y")

    if not msgs:
        await bot.send_message(
            chat_id=MY_CHAT_ID,
            text=f"🗓 Утренний отчёт — {today}\n\nНет новых сообщений от агента."
        )
        return

    prompt = f"""Проанализируй эти сообщения от финансового агента и составь утренний отчёт.

Формат отчёта:
🗓 Утренний отчёт — {today}

💰 Баланс агента: [последний известный баланс в USD]

✅ Оплачено/подтверждено:
— [список с суммами]

⏳ Ожидают оплаты:
— [список]

⚠ Требует внимания:
— [проблемы, неизвестные платежи, неподтверждённые переводы]

📊 Транзакции для добавления в Excel:
— [дата | тип | описание | сумма | валюта]

Сообщения от агента:
{_format_msgs(msgs)}"""

    summary = await ask_claude(prompt)

    await bot.send_message(chat_id=MY_CHAT_ID, text=summary)

    if EXCEL_FILE.exists():
        await bot.send_document(
            chat_id=MY_CHAT_ID,
            document=EXCEL_FILE.open("rb"),
            filename=f"Agent_Report_{datetime.now().strftime('%Y%m%d')}.xlsx",
            caption="📎 Excel — добавь новые транзакции вручную по списку выше"
        )

    clear_messages()
    log.info(f"Morning report sent, {len(msgs)} messages processed")

async def morning_job(ctx: ContextTypes.DEFAULT_TYPE):
    await _send_morning_report(ctx.bot)

# ── Main ──────────────────────────────────────────────────────────────────────
def main():
    Path("data").mkdir(exist_ok=True)

    app = Application.builder().token(BOT_TOKEN).build()

    app.add_handler(CommandHandler("start",   cmd_start))
    app.add_handler(CommandHandler("balance", cmd_balance))
    app.add_handler(CommandHandler("pending", cmd_pending))
    app.add_handler(CommandHandler("unknown", cmd_unknown))
    app.add_handler(CommandHandler("summary", cmd_summary))
    app.add_handler(CommandHandler("excel",   cmd_excel))
    app.add_handler(CommandHandler("clear",   cmd_clear))
    app.add_handler(MessageHandler(filters.ALL & ~filters.COMMAND, handle_message))

    # Schedule morning report
    app.job_queue.run_daily(
        morning_job,
        time=time(hour=MORNING_HOUR, minute=0),
    )

    log.info(f"Bot started. Morning report at {MORNING_HOUR}:00")
    app.run_polling(allowed_updates=Update.ALL_TYPES)

if __name__ == "__main__":
    main()
