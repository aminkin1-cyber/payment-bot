"""
Payment Tracker Bot v2 — auto-updates Excel via Claude API
"""
import os, json, logging
from datetime import datetime, time
from pathlib import Path
import httpx
from openpyxl import load_workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from telegram import Update, Bot
from telegram.ext import Application, MessageHandler, CommandHandler, filters, ContextTypes

BOT_TOKEN     = os.environ["BOT_TOKEN"]
ANTHROPIC_KEY = os.environ["ANTHROPIC_KEY"]
MY_CHAT_ID    = int(os.environ["MY_CHAT_ID"])
MORNING_HOUR  = int(os.environ.get("MORNING_HOUR", "9"))
DATA_FILE     = Path("data/messages.json")
EXCEL_FILE    = Path("Agent_Model_v2.xlsx")

logging.basicConfig(level=logging.INFO)
log = logging.getLogger(__name__)

# ── Styles ────────────────────────────────────────────────────────────────────
WHITE  = "FFFFFF"; YELLOW = "FFF2CC"; GREEN  = "E2EFDA"
RED    = "FCE4D6"; ORANGE = "FDEBD0"; LIGHT  = "D6E4F0"; LGRAY  = "F2F2F2"
thin   = Side(style="thin", color="BFBFBF")
def B(): return Border(top=thin, bottom=thin, left=thin, right=thin)

TYPE_BG = {
    "Deposit":    GREEN,
    "Payment":    WHITE,
    "Cash Out":   ORANGE,
    "Cash In":    LIGHT,
    "❓ Unknown": RED,
}
STAT_BG = {
    "✅ Paid":          GREEN,
    "⏳ Pending":       YELLOW,
    "⚠ Partial/Check": ORANGE,
    "❓ Clarify":       RED,
}

def style_cell(cell, bg=WHITE, bold=False, sz=9, fc="000000", num=None,
               align="left", wrap=False):
    cell.font      = Font(name="Arial", bold=bold, size=sz, color=fc)
    cell.fill      = PatternFill("solid", fgColor=bg)
    cell.alignment = Alignment(horizontal=align, vertical="center", wrap_text=wrap)
    cell.border    = B()
    if num: cell.number_format = num

# ── Excel read helpers ────────────────────────────────────────────────────────
def get_balance_from_excel():
    if not EXCEL_FILE.exists():
        return None
    try:
        wb = load_workbook(EXCEL_FILE, data_only=True)
        ws = wb["Transactions"]
        last_bal, last_date = None, None
        for row in ws.iter_rows(min_row=5, max_col=11, values_only=True):
            if row[10] is not None and isinstance(row[10], (int, float)):
                last_bal  = row[10]
                last_date = row[0]
        return (last_bal, last_date) if last_bal else None
    except Exception as e:
        log.error(f"Excel balance error: {e}"); return None

def get_pending_invoices():
    if not EXCEL_FILE.exists(): return []
    try:
        wb = load_workbook(EXCEL_FILE, data_only=True)
        ws = wb["Invoices"]
        out = []
        for row in ws.iter_rows(min_row=5, max_col=10, values_only=True):
            if row[6] and "Pending" in str(row[6]):
                amt = f"{row[4]:,.2f}" if isinstance(row[4], (int,float)) else str(row[4] or "TBC")
                out.append(f"- {row[2] or '?'}: {amt} {row[3] or ''}")
        return out
    except Exception as e:
        log.error(f"Excel pending error: {e}"); return []

def get_unknown_transactions():
    if not EXCEL_FILE.exists(): return []
    try:
        wb = load_workbook(EXCEL_FILE, data_only=True)
        ws = wb["Transactions"]
        out = []
        for row in ws.iter_rows(min_row=5, max_col=12, values_only=True):
            if row[1] and "Unknown" in str(row[1]):
                amt = f"{row[5]:,.2f}" if isinstance(row[5], (int,float)) else str(row[5] or "?")
                out.append(f"- {row[0]}: {row[2] or '?'} | {amt} {row[4] or ''}")
        return out
    except Exception as e:
        log.error(f"Excel unknown error: {e}"); return []

def find_last_tx_row(ws):
    last = 5
    for row in ws.iter_rows(min_row=5, max_col=1):
        if row[0].value is not None:
            last = row[0].row
    return last

def find_last_inv_row(ws):
    last = 5
    for row in ws.iter_rows(min_row=5, max_col=1):
        if row[0].value is not None:
            last = row[0].row
    return last

# ── Excel WRITE ───────────────────────────────────────────────────────────────
def apply_tx_row(ws, r, tx):
    """Write one transaction dict into Transactions sheet row r."""
    tp  = tx.get("type", "Payment")
    bg  = TYPE_BG.get(tp, WHITE)

    vals = [
        tx.get("date",""),
        tp,
        tx.get("description",""),
        tx.get("payee",""),
        tx.get("ccy",""),
        tx.get("amount"),
        tx.get("fx_rate"),   # col G — may be None → formula will fill
        None,                 # col H gross — formula
        tx.get("comm"),       # col I — may be None → formula
        None,                 # col J net — formula
        None,                 # col K balance — formula
        tx.get("notes",""),
    ]
    for col_i, val in enumerate(vals, 1):
        c = ws.cell(r, col_i, val if val is not None else "")
        style_cell(c, bg=bg, wrap=(col_i in (3,12)), sz=9)

    # Formulas
    # G: FX rate — if not provided use Settings lookup
    if not tx.get("fx_rate"):
        ws.cell(r, 7).value = f'=IF(E{r}="","",IFERROR(VLOOKUP(E{r},Settings!$A$7:$B$16,2,FALSE),1))'
    ws.cell(r, 7).number_format = "0.00000"
    style_cell(ws.cell(r,7), bg=YELLOW, fc="0000CC")

    # I: comm — if not provided use Settings lookup
    if not tx.get("comm"):
        ws.cell(r, 9).value = f'=IF(B{r}="","",IFERROR(VLOOKUP(B{r},Settings!$E$19:$F$23,2,FALSE),0))'
    ws.cell(r, 9).number_format = "0.0%"
    style_cell(ws.cell(r,9), bg=YELLOW, fc="0000CC")

    # H: Gross USD
    ws.cell(r, 8).value = f'=IF(OR(F{r}="",G{r}=""),"",F{r}/G{r})'
    ws.cell(r, 8).number_format = '#,##0.00'
    style_cell(ws.cell(r,8), bg=YELLOW)

    # J: Net USD
    ws.cell(r,10).value = (
        f'=IF(H{r}="","",IF(OR(B{r}="Deposit",B{r}="Cash In"),'
        f'H{r},-(H{r}/MAX(1-I{r},0.0001))))'
    )
    ws.cell(r,10).number_format = '#,##0.00'
    style_cell(ws.cell(r,10), bg=YELLOW)

    # K: Running balance
    ws.cell(r,11).value = (
        f'=IF(J{r}="","",IF(ISNUMBER(K{r-1}),K{r-1},Settings!$C$4)+J{r})'
    )
    ws.cell(r,11).number_format = '#,##0.00'
    style_cell(ws.cell(r,11), bg=YELLOW, bold=True)

    ws.row_dimensions[r].height = 28

def apply_inv_update(ws, inv_upd):
    """Update invoice status/date/ref for a matching invoice number."""
    inv_no   = str(inv_upd.get("invoice_no","")).strip().lower()
    status   = inv_upd.get("new_status","✅ Paid")
    date_pd  = inv_upd.get("date_paid","")
    ref      = inv_upd.get("ref","")

    for row in ws.iter_rows(min_row=5, max_col=10):
        cell_inv = str(row[1].value or "").strip().lower()
        if inv_no and inv_no in cell_inv:
            bg = STAT_BG.get(status, YELLOW)
            row[6].value = status
            style_cell(row[6], bg=bg, bold=True, align="center")
            row[7].value = date_pd
            style_cell(row[7], bg=bg)
            if ref:
                row[8].value = ref
                style_cell(row[8], bg=bg, sz=8)
            log.info(f"Updated invoice {cell_inv} → {status}")
            return True
    log.warning(f"Invoice not found: {inv_no}")
    return False

def add_new_invoice(ws, inv, last_row):
    """Add a brand new invoice row."""
    r  = last_row + 1
    st = inv.get("status","⏳ Pending")
    bg = STAT_BG.get(st, YELLOW)
    data = [
        inv.get("date",""), inv.get("invoice_no",""), inv.get("payee",""),
        inv.get("ccy",""), inv.get("amount"), None,  # F=usd formula
        st, inv.get("date_paid",""), inv.get("ref",""), inv.get("notes",""),
    ]
    for col_i, val in enumerate(data, 1):
        c = ws.cell(r, col_i, val if val is not None else "")
        style_cell(c, bg=bg, wrap=(col_i in (3,10)), sz=9)
    # USD equiv formula
    ws.cell(r,6).value = (
        f'=IF(OR(E{r}="",E{r}="TBC"),"TBC",'
        f'IFERROR(E{r}/VLOOKUP(D{r},Settings!$A$7:$B$16,2,FALSE),E{r}))'
    )
    ws.cell(r,6).number_format = '#,##0.00'
    style_cell(ws.cell(r,6), bg=bg)
    ws.row_dimensions[r].height = 26

def write_to_excel(claude_json: dict) -> tuple[int,int,int]:
    """Apply Claude's structured output to Excel. Returns (tx_added, inv_updated, inv_added)."""
    if not EXCEL_FILE.exists():
        return 0,0,0

    wb  = load_workbook(EXCEL_FILE)
    wst = wb["Transactions"]
    wsi = wb["Invoices"]

    tx_added = inv_updated = inv_added = 0

    # 1. New transactions
    for tx in claude_json.get("new_transactions", []):
        last = find_last_tx_row(wst)
        apply_tx_row(wst, last + 1, tx)
        tx_added += 1

    # 2. Invoice status updates
    for upd in claude_json.get("invoice_updates", []):
        if apply_inv_update(wsi, upd):
            inv_updated += 1

    # 3. New invoices
    for inv in claude_json.get("new_invoices", []):
        last = find_last_inv_row(wsi)
        add_new_invoice(wsi, inv, last)
        inv_added += 1

    wb.save(EXCEL_FILE)
    return tx_added, inv_updated, inv_added

# ── Message store ─────────────────────────────────────────────────────────────
def load_messages():
    DATA_FILE.parent.mkdir(exist_ok=True)
    return json.loads(DATA_FILE.read_text()) if DATA_FILE.exists() else []

def save_message(d):
    msgs = load_messages(); msgs.append(d)
    DATA_FILE.write_text(json.dumps(msgs, ensure_ascii=False, indent=2))

def clear_messages():
    DATA_FILE.write_text("[]")

def _fmt(msgs):
    return "\n".join(
        f"[{m['date']}] {m.get('sender','?')}: {m.get('text','')} "
        f"{'[файл: '+m['file']+']' if m.get('file') else ''}"
        for m in msgs
    )

# ── Claude API ────────────────────────────────────────────────────────────────
async def ask_claude(prompt: str, system: str = None) -> str:
    sys = system or (
        "You are a financial assistant. The agent handles payments in AED/CNY/USD/EUR/SGD/RUB. "
        "AED/USD = 3.6725. Commission 0.5% on most payments, 0.4% on RUB, 0% on deposits/cash-in. "
        "Respond in Russian unless asked for JSON."
    )
    async with httpx.AsyncClient(timeout=90) as client:
        r = await client.post(
            "https://api.anthropic.com/v1/messages",
            headers={"x-api-key": ANTHROPIC_KEY,
                     "anthropic-version": "2023-06-01",
                     "content-type": "application/json"},
            json={"model": "claude-opus-4-6", "max_tokens": 3000,
                  "system": sys,
                  "messages": [{"role": "user", "content": prompt}]},
        )
        return r.json()["content"][0]["text"]

async def parse_messages_to_json(msgs_text: str) -> dict:
    """Ask Claude to extract structured data from forwarded messages."""
    prompt = f"""Из этих сообщений от финансового агента извлеки структурированные данные.

Верни ТОЛЬКО валидный JSON без markdown, без пояснений, строго в таком формате:
{{
  "new_transactions": [
    {{
      "date": "DD.MM.YYYY",
      "type": "Payment|Deposit|Cash Out|Cash In|❓ Unknown",
      "description": "краткое описание",
      "payee": "название получателя",
      "ccy": "AED|CNY|USD|EUR|SGD|RUB|INR",
      "amount": 12345.67,
      "fx_rate": null,
      "comm": null,
      "notes": "дополнительно"
    }}
  ],
  "invoice_updates": [
    {{
      "invoice_no": "номер инвойса",
      "new_status": "✅ Paid|⏳ Pending|⚠ Partial/Check|❓ Clarify",
      "date_paid": "DD.MM.YYYY",
      "ref": "референс платежа"
    }}
  ],
  "new_invoices": [
    {{
      "date": "DD.MM.YYYY",
      "invoice_no": "номер",
      "payee": "получатель",
      "ccy": "USD",
      "amount": 12345.67,
      "status": "⏳ Pending",
      "notes": ""
    }}
  ],
  "summary": "2-3 предложения — что нового произошло"
}}

Правила:
- Если сообщение содержит баланс агента — это НЕ транзакция, пропусти
- "ИСПОЛНЕН", "received", "RCVD" = подтверждение оплаты → invoice_updates
- Депозиты от нас агенту = type Deposit
- Кэш который агент нам доставляет = type Cash Out
- Если непонятно — type "❓ Unknown"
- Если нет новых данных — верни пустые массивы

Сообщения:
{msgs_text}"""

    raw = await ask_claude(prompt, system=(
        "You are a JSON extraction assistant. Return ONLY valid JSON, no markdown, no explanation."
    ))
    # Strip markdown if Claude added it anyway
    raw = raw.strip()
    if raw.startswith("```"):
        raw = raw.split("```")[1]
        if raw.startswith("json"):
            raw = raw[4:]
    raw = raw.strip().rstrip("```").strip()
    return json.loads(raw)

# ── Commands ──────────────────────────────────────────────────────────────────
async def cmd_start(update: Update, ctx: ContextTypes.DEFAULT_TYPE):
    await update.message.reply_text(
        "Привет! Я трекер платежей.\n\n"
        "Пересылай мне сообщения от агента — накоплю их.\n"
        "Когда готов обновить Excel — напиши /update\n\n"
        "Команды:\n"
        "/update  — обработать пересланные сообщения и обновить Excel\n"
        "/balance — текущий баланс из Excel\n"
        "/pending — что висит (из Excel)\n"
        "/unknown — неизвестные транзакции\n"
        "/summary — полное саммари\n"
        "/excel   — прислать Excel файл\n"
        "/clear   — очистить накопленные сообщения"
    )

async def cmd_update(update: Update, ctx: ContextTypes.DEFAULT_TYPE):
    """Main command: parse forwarded messages → update Excel → send back."""
    msgs = load_messages()
    if not msgs:
        await update.message.reply_text(
            "Нет накопленных сообщений. Сначала перешли что-нибудь от агента.")
        return

    await update.message.reply_text(
        f"Обрабатываю {len(msgs)} сообщений... Секунду.")

    try:
        data = await parse_messages_to_json(_fmt(msgs))
    except Exception as e:
        await update.message.reply_text(f"Ошибка парсинга: {e}")
        log.error(f"Parse error: {e}")
        return

    tx_n = len(data.get("new_transactions", []))
    upd_n = len(data.get("invoice_updates", []))
    inv_n = len(data.get("new_invoices", []))

    if tx_n + upd_n + inv_n == 0:
        await update.message.reply_text(
            f"Проанализировал сообщения — новых транзакций или инвойсов не найдено.\n\n"
            f"Итог: {data.get('summary','')}")
        clear_messages()
        return

    # Write to Excel
    try:
        tx_added, inv_updated, inv_added = write_to_excel(data)
    except Exception as e:
        await update.message.reply_text(f"Ошибка записи в Excel: {e}")
        log.error(f"Excel write error: {e}")
        return

    summary = data.get("summary","")
    result_text = (
        f"Excel обновлён!\n\n"
        f"Добавлено транзакций: {tx_added}\n"
        f"Обновлено инвойсов: {inv_updated}\n"
        f"Добавлено новых инвойсов: {inv_added}\n\n"
        f"{summary}"
    )
    await update.message.reply_text(result_text)

    # Send updated Excel
    if EXCEL_FILE.exists():
        await ctx.bot.send_document(
            chat_id=MY_CHAT_ID,
            document=EXCEL_FILE.open("rb"),
            filename=f"Agent_{datetime.now().strftime('%Y%m%d_%H%M')}.xlsx",
            caption="Обновлённый Excel"
        )

    clear_messages()

async def cmd_balance(update: Update, ctx: ContextTypes.DEFAULT_TYPE):
    result = get_balance_from_excel()
    if result:
        bal, date = result
        await update.message.reply_text(
            f"БАЛАНС АГЕНТА (из Excel)\n"
            f"${bal:,.2f} USD\n"
            f"Последняя запись: {date}")
    else:
        await update.message.reply_text("Баланс не найден. Попробуй /update или /excel")

async def cmd_pending(update: Update, ctx: ContextTypes.DEFAULT_TYPE):
    pending = get_pending_invoices()
    if pending:
        await update.message.reply_text(
            "ОЖИДАЮТ ОПЛАТЫ:\n\n" + "\n".join(pending))
    else:
        await update.message.reply_text("Нет pending инвойсов в Excel.")

async def cmd_unknown(update: Update, ctx: ContextTypes.DEFAULT_TYPE):
    unknowns = get_unknown_transactions()
    if unknowns:
        await update.message.reply_text(
            "НЕИЗВЕСТНЫЕ ТРАНЗАКЦИИ:\n\n" + "\n".join(unknowns))
    else:
        await update.message.reply_text("Нет неизвестных транзакций — хорошо!")

async def cmd_summary(update: Update, ctx: ContextTypes.DEFAULT_TYPE):
    await _send_report(ctx.bot, triggered_manually=True)

async def cmd_excel(update: Update, ctx: ContextTypes.DEFAULT_TYPE):
    if EXCEL_FILE.exists():
        await ctx.bot.send_document(
            chat_id=MY_CHAT_ID,
            document=EXCEL_FILE.open("rb"),
            filename="Agent_Model.xlsx",
            caption="Актуальный Excel"
        )
    else:
        await update.message.reply_text("Excel файл не найден на сервере.")

async def cmd_clear(update: Update, ctx: ContextTypes.DEFAULT_TYPE):
    clear_messages()
    await update.message.reply_text("Накопленные сообщения очищены.")

# ── Message handler ───────────────────────────────────────────────────────────
async def handle_message(update: Update, ctx: ContextTypes.DEFAULT_TYPE):
    msg = update.message
    if msg.chat_id != MY_CHAT_ID:
        return
    text     = msg.text or msg.caption or ""
    sender   = (msg.forward_from.full_name if msg.forward_from
                else msg.forward_sender_name or "")
    date_str = (msg.forward_date or msg.date).strftime("%d.%m.%Y %H:%M")
    file_n   = msg.document.file_name if msg.document else ""

    save_message({"date": date_str, "sender": sender,
                  "text": text, "file": file_n})

    preview = text[:60] + ("…" if len(text) > 60 else "")
    parts   = [f"от {sender}"] if sender else []
    if file_n:  parts.append(f"файл: {file_n}")
    if preview: parts.append(f'"{preview}"')
    count = len(load_messages())
    await msg.reply_text(
        f"Сохранено ({date_str}): {' | '.join(parts)}\n"
        f"Накоплено сообщений: {count}. Когда готов — пиши /update"
    )

# ── Report ────────────────────────────────────────────────────────────────────
async def _send_report(bot: Bot, triggered_manually=False):
    today   = datetime.now().strftime("%d.%m.%Y")
    msgs    = load_messages()

    # Auto-update Excel if there are pending messages
    updates_text = ""
    if msgs:
        try:
            data = await parse_messages_to_json(_fmt(msgs))
            tx_a, inv_u, inv_a = write_to_excel(data)
            if tx_a + inv_u + inv_a > 0:
                updates_text = (f"\nОбновлено автоматически: "
                                f"+{tx_a} транзакций, {inv_u} инвойсов обновлено, "
                                f"+{inv_a} новых инвойсов")
            clear_messages()
        except Exception as e:
            log.error(f"Auto-update error: {e}")
            updates_text = f"\n(Ошибка автообновления: {e})"

    result  = get_balance_from_excel()
    bal_str = f"${result[0]:,.2f} USD (запись: {result[1]})" if result else "нет данных"
    pending = get_pending_invoices()
    unknown = get_unknown_transactions()

    text = (
        f"ОТЧЁТ — {today}\n"
        f"{updates_text}\n\n"
        f"БАЛАНС АГЕНТА: {bal_str}\n\n"
        f"ОЖИДАЮТ ОПЛАТЫ ({len(pending)}):\n"
        + ("\n".join(pending) if pending else "нет") +
        f"\n\nНЕИЗВЕСТНЫЕ ТРАНЗАКЦИИ ({len(unknown)}):\n"
        + ("\n".join(unknown) if unknown else "нет")
    )

    await bot.send_message(chat_id=MY_CHAT_ID, text=text)
    if EXCEL_FILE.exists():
        await bot.send_document(
            chat_id=MY_CHAT_ID,
            document=EXCEL_FILE.open("rb"),
            filename=f"Agent_{datetime.now().strftime('%Y%m%d')}.xlsx",
            caption="Актуальный Excel"
        )

async def morning_job(ctx: ContextTypes.DEFAULT_TYPE):
    await _send_report(ctx.bot)

# ── Main ──────────────────────────────────────────────────────────────────────
def main():
    Path("data").mkdir(exist_ok=True)
    app = Application.builder().token(BOT_TOKEN).build()
    for cmd, fn in [("start", cmd_start), ("update", cmd_update),
                    ("balance", cmd_balance), ("pending", cmd_pending),
                    ("unknown", cmd_unknown), ("summary", cmd_summary),
                    ("excel", cmd_excel), ("clear", cmd_clear)]:
        app.add_handler(CommandHandler(cmd, fn))
    app.add_handler(MessageHandler(filters.ALL & ~filters.COMMAND, handle_message))
    app.job_queue.run_daily(morning_job, time=time(hour=MORNING_HOUR, minute=0))
    log.info(f"Bot v2 started. Morning report at {MORNING_HOUR}:00")
    app.run_polling(allowed_updates=Update.ALL_TYPES)

if __name__ == "__main__":
    main()
